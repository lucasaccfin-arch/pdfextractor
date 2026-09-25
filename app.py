import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Chakradhara Aerospace - Sales Invoice Extractor", page_icon="✈️", layout="wide")

st.title("✈️ Chakradhara Aerospace - Sales Invoice Extractor")
st.write("Upload sales invoice PDFs to extract structured tax details and export to formatted Excel.")

# Official Company GSTIN List for Chakradhara Aerospace
COMPANY_GSTINS = {
    "30AAHCC1431F1ZF",
    "29AAHCC1431F1ZY",
    "32AAHCC1431F1ZB",
    "34AAHCC1431F1Z7",
    "36AAHCC1431F1Z3",
    "37AAHCC1431F1Z1",
    "33AAHCC1431F1Z9"
}

uploaded_files = st.file_uploader("Upload Sales Invoice PDFs", type=["pdf"], accept_multiple_files=True)

def parse_float(val):
    """Safely extracts clean floating numbers from text."""
    if not val or val == "N/A":
        return 0.0
    cleaned = re.sub(r'[^\d\.]', '', str(val))
    try:
        return float(cleaned) if cleaned else 0.0
    except ValueError:
        return 0.0

def parse_date(date_str):
    """Converts extracted text dates into standard datetime objects."""
    if not date_str or date_str == "N/A":
        return None
    
    date_patterns = [
        r'\d{2}[\/\.-]\d{2}[\/\.-]\d{4}',
        r'\d{2}[\/\.-]\d{2}[\/\.-]\d{2}',
        r'\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}',
        r'\d{1,2}[\/\.-][A-Za-z]{3}[\/\.-]\d{2,4}'
    ]
    
    date_formats = [
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
        "%d/%m/%y", "%d-%m-%y", "%d.%m.%y",
        "%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d-%b-%y"
    ]

    for pattern in date_patterns:
        match = re.search(pattern, date_str)
        if match:
            raw_date = match.group(0).strip()
            for fmt in date_formats:
                try:
                    return datetime.strptime(raw_date, fmt)
                except ValueError:
                    continue
    return None

def extract_sales_invoice_data(pdf_bytes):
    """Extracts required fields from Chakradhara Aerospace invoice formats."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = "".join([page.get_text() for page in doc])
    doc.close()

    # 1. Identify GSTINs
    all_gstins = re.findall(r'\b[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b', text)
    
    my_gstin = "N/A"
    customer_gstin = "N/A"

    for gstin in all_gstins:
        if gstin in COMPANY_GSTINS and my_gstin == "N/A":
            my_gstin = gstin
        elif gstin not in COMPANY_GSTINS and customer_gstin == "N/A":
            customer_gstin = gstin

    # Fallback if both GSTINs are from company or regular pattern lookup
    if my_gstin == "N/A" and len(all_gstins) > 0:
        my_gstin = all_gstins[0]
    if customer_gstin == "N/A" and len(all_gstins) > 1:
        customer_gstin = all_gstins[1]

    # 2. Customer / Recipient Name
    cust_name_match = re.search(r'(?:Billed\s*To|Recipient|Customer\s*Name|Party\s*Name|M/s\.?)\s*[:\-]?\s*([^\n]+)', text, re.IGNORECASE)
    customer_name = cust_name_match.group(1).strip() if cust_name_match else "N/A"

    # 3. Invoice Number
    inv_no_match = re.search(r'(?:Invoice\s*No\.?|Inv\s*No\.?|Invoice\s*#)\s*[:\-]?\s*([A-Za-z0-9\/\-]+)', text, re.IGNORECASE)
    invoice_no = inv_no_match.group(1).strip() if inv_no_match else "N/A"

    # 4. Invoice Date
    inv_date_match = re.search(r'(?:Invoice\s*Date|Dated|Date)\s*[:\-]?\s*(\d{1,2}[\/\.-]\w+[\/\.-]\d{2,4}|\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})', text, re.IGNORECASE)
    raw_date = inv_date_match.group(1).strip() if inv_date_match else "N/A"
    invoice_date = parse_date(raw_date)

    # 5. Place of Supply / State Code
    pos_match = re.search(r'(?:Place\s*of\s*Supply|State\s*Code|State\s*Name|POS)\s*[:\-]?\s*([A-Za-z0-9\s\(\)\-]+)', text, re.IGNORECASE)
    place_of_supply = pos_match.group(1).strip().split('\n')[0] if pos_match else "N/A"

    # 6. HSN / SAC Code
    hsn_match = re.search(r'(?:HSN\s*\/\s*SAC|HSN\s*Code|SAC\s*Code)\s*[:\-]?\s*(\d{4,8})', text, re.IGNORECASE)
    hsn_code = hsn_match.group(1).strip() if hsn_match else "N/A"

    # 7. Amounts & Taxes
    taxable_val_match = re.search(r'(?:Taxable\s*Value|Sub\s*Total|Taxable\s*Amount)\s*[:\-]?\s*(?:₹|Rs\.?)?\s*([\d,]+\.\d{2}|[\d,]+)', text, re.IGNORECASE)
    taxable_value = parse_float(taxable_val_match.group(1)) if taxable_val_match else 0.0

    cgst_match = re.search(r'(?:CGST)\s*(?:@\s*\d+%\s*)?[:\-]?\s*(?:₹|Rs\.?)?\s*([\d,]+\.\d{2}|[\d,]+)', text, re.IGNORECASE)
    cgst = parse_float(cgst_match.group(1)) if cgst_match else 0.0

    sgst_match = re.search(r'(?:SGST|UTGST)\s*(?:@\s*\d+%\s*)?[:\-]?\s*(?:₹|Rs\.?)?\s*([\d,]+\.\d{2}|[\d,]+)', text, re.IGNORECASE)
    sgst = parse_float(sgst_match.group(1)) if sgst_match else 0.0

    igst_match = re.search(r'(?:IGST)\s*(?:@\s*\d+%\s*)?[:\-]?\s*(?:₹|Rs\.?)?\s*([\d,]+\.\d{2}|[\d,]+)', text, re.IGNORECASE)
    igst = parse_float(igst_match.group(1)) if igst_match else 0.0

    total_tax = round(cgst + sgst + igst, 2)

    total_inv_val_match = re.search(r'(?:Grand\s*Total|Invoice\s*Total|Total\s*Amount)\s*[:\-]?\s*(?:₹|Rs\.?)?\s*([\d,]+\.\d{2}|[\d,]+)', text, re.IGNORECASE)
    total_invoice_value = parse_float(total_inv_val_match.group(1)) if total_inv_val_match else (taxable_value + total_tax)

    return {
        "GSTIN": my_gstin,
        "Customer Name": customer_name,
        "Customer GSTIN": customer_gstin,
        "Invoice No": invoice_no,
        "Invoice Date": invoice_date,
        "Place of Supply": place_of_supply,
        "HSN/SAC": hsn_code,
        "Taxable Value": taxable_value,
        "CGST": cgst,
        "SGST": sgst,
        "IGST": igst,
        "Total Tax": total_tax,
        "Total Invoice Value": total_invoice_value
    }

def format_excel_workbook(df):
    """Formats Excel file with correct Date, Number types, Alignment, Auto-filters, and Auto-fit."""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales Invoices"

    headers = list(df.columns)
    ws.append(headers)

    # Style Header Row (Bold, Center Aligned)
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Numeric & Date Column Names
    num_cols = {"Taxable Value", "CGST", "SGST", "IGST", "Total Tax", "Total Invoice Value"}
    
    # Append Data Rows with Proper Excel Types
    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, (col_name, val) in enumerate(zip(headers, row), start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            
            if col_name == "Invoice Date" and pd.notnull(val):
                cell.value = val
                cell.number_format = 'yyyy-mm-dd'
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name in num_cols:
                cell.value = float(val) if val is not None else 0.0
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.value = str(val) if val is not None else ""
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Enable Auto Filter
    ws.auto_filter.ref = ws.dimensions

    # Universal Auto-Fit Column Widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is not None:
                if isinstance(cell.value, datetime):
                    max_len = max(max_len, 10)
                else:
                    max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output)
    return output.getvalue()

if uploaded_files:
    if st.button("Extract Invoice Details"):
        results = []
        for file in uploaded_files:
            pdf_bytes = file.read()
            extracted_info = extract_sales_invoice_data(pdf_bytes)
            extracted_info["Filename"] = file.name
            results.append(extracted_info)

        df = pd.DataFrame(results)
        
        # Order columns cleanly with Filename first
        cols = ["Filename", "GSTIN", "Customer Name", "Customer GSTIN", "Invoice No", "Invoice Date", 
                "Place of Supply", "HSN/SAC", "Taxable Value", "CGST", "SGST", "IGST", 
                "Total Tax", "Total Invoice Value"]
        df = df[cols]

        st.success(f"Processed {len(results)} invoices successfully!")
        st.dataframe(df)

        excel_bytes = format_excel_workbook(df)

        st.download_button(
            label="📥 Download Structured Excel File",
            data=excel_bytes,
            file_name="Chakradhara_Sales_Invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )