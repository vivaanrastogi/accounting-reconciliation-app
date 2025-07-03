import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import math

st.set_page_config(page_title="Accounting Reconciliation Portal", layout="centered")

# --- Add your logo here ---
st.image("logo.png", width=150)  # Make sure 'logo.png' is in the same folder
st.title("📄 Accounting Reconciliation Portal")

# Step 1: Inputs
company = st.text_input("Enter company name (e.g., HERCULES):")
month = st.text_input("Enter month in YYYYMM format (e.g., 202504):")
uploaded_tb = st.file_uploader("Upload TB PDF file", type=["pdf"])

# Manual PDF amounts and mapping
pdf_actual_values = {
    "Bank1 amt": 5331520.94,
    "Bank2 amt": None,
    "Bank3 amt": None,
    "Bank4 amt": None,
    "Bank5 amt": None,
    "Bank6 amt": None,
    "Bank7 amt": None,
    "Bank8 amt": None,
    "PND1 amt": 1000.00,
    "PND3 amt": 165.00,
    "PND53 amt": 540.00,
    "PP30 amt": 44145.07,
    "SSO amt": None
}

tb_code_map = {
    "Bank1 amt": "1112-01",
    "Bank2 amt": "1113-01",
    "Bank3 amt": "1114-01",
    "Bank4 amt": "1115-01",
    "Bank5 amt": "1116-01",
    "Bank6 amt": "1117-01",
    "Bank7 amt": "1118-01",
    "Bank8 amt": "1119-01",
    "PND1 amt":  "2132-01",
    "PND3 amt":  "2132-02",
    "PND53 amt": "2132-02",
    "PP30 amt":  "2137-00",
    "SSO amt":   "2131-04"
}

file_map = {
    f"Bank{i} amt": f"bank{i}_{company.lower()}_{month}.pdf" for i in range(1, 9)
}
file_map.update({
    "PND1 amt": f"0.PND1_{month}.pdf",
    "PND3 amt": f"1.PND3_{month}.pdf",
    "PND53 amt": f"2.PND53_{month}.pdf",
    "PP30 amt": f"ภ.พ.30_{month}.pdf",
    "SSO amt": f"สปส1-10_{month}.pdf"
})

if uploaded_tb and company and month:
    st.success("File uploaded. Processing...")

    # Save TB PDF
    tb_filename = f"{company.lower()}_tb_{month}.pdf"
    with open(tb_filename, "wb") as f:
        f.write(uploaded_tb.read())

    # Extract text from TB PDF
    with fitz.open(tb_filename) as doc:
        text = "\n".join(page.get_text() for page in doc)

    pattern = re.compile(
        r"(\d{4}-\d{2})\s+.+?([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})"
    )

    tb_data = []
    for line in text.splitlines():
        match = pattern.match(line)
        if match:
            code = match.group(1)
            balance_debit = float(match.group(6).replace(",", ""))
            balance_credit = float(match.group(7).replace(",", ""))
            amount = balance_debit if balance_debit > 0 else -balance_credit
            tb_data.append({"Code": code, "Amount": amount})

    df_tb = pd.DataFrame(tb_data)

    # Build reconciliation results
    results = []
    staff_name = "tip"  # or customize/get dynamically

    for name, tb_code in tb_code_map.items():
        tb_rows = df_tb[df_tb["Code"] == tb_code]
        tb_amt = tb_rows["Amount"].iloc[0] if not tb_rows.empty else None
        pdf_amt = pdf_actual_values.get(name)

        if tb_amt is not None and pdf_amt is not None:
            result = "Match> Correct" if abs(abs(tb_amt) - abs(pdf_amt)) < 1e-2 else "Mismatch Wrong"
        else:
            result = ""

        # Format amount for display
        tb_amt_fmt = f"{tb_amt:,.2f}" if tb_amt is not None else ""

        results.append({
            "Name": name,
            "source file": file_map[name],
            "TB Code": tb_code,
            "TB code amount column5(+),6(-)": tb_amt_fmt,
            "PDF actual amount": pdf_amt,
            "Results": result,
            "Staff name": staff_name
        })

    df_result = pd.DataFrame(results)

    st.subheader("📊 Reconciliation Summary")
    st.dataframe(df_result)

    # Excel export with formatting and fix for None/NaN
    output_file = f"result_{company.lower()}_{month}.xlsx"
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df_result.to_excel(writer, index=False, sheet_name='Summary')
        workbook = writer.book
        worksheet = writer.sheets['Summary']

        # Header format
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#FFD700',
            'font_color': 'black',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        # Amount format
        amount_format = workbook.add_format({
            'num_format': '#,##0.00',
            'align': 'right',
            'valign': 'vcenter',
            'border': 1
        })

        # Cell format
        cell_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        # Write headers
        for col_num, value in enumerate(df_result.columns):
            worksheet.write(0, col_num, value, header_format)

        # Write data rows with safe type handling
        for row_num in range(1, len(df_result) + 1):
            for col_num in range(len(df_result.columns)):
                val = df_result.iloc[row_num - 1, col_num]
                col_name = df_result.columns[col_num]

                if col_name == "TB code amount column5(+),6(-)":
                    if val not in ("", None):
                        try:
                            val_float = float(str(val).replace(",", ""))
                            worksheet.write(row_num, col_num, val_float, amount_format)
                        except Exception:
                            worksheet.write(row_num, col_num, val, cell_format)  # fallback
                    else:
                        worksheet.write(row_num, col_num, "", cell_format)
                elif val is None or (isinstance(val, float) and (math.isnan(val))):
                    worksheet.write(row_num, col_num, "", cell_format)
                else:
                    worksheet.write(row_num, col_num, val, cell_format)

        # Auto-adjust column widths
        for i, col in enumerate(df_result.columns):
            max_len = max(df_result[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)

    with open(output_file, "rb") as f:
        st.download_button("⬇️ Download Excel", f, file_name=output_file)

else:
    st.info("Please enter company name, month, and upload TB PDF.")
