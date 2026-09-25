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
st.write("Upload sales invoice PDFs across all formats to extract structured tax details (including Non-Taxable/Exempt items) and export to formatted Excel.")

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

def extract_customer_name(text):
    """Extracts Customer Name directly under Customer header ignoring page numbers."""
    match = re.search(r'Customer\s*[:\-]\s*[\r\n]+\s*([A-Za-z0-9\s&\.\,\-\(\)]+)', text, re.IGNORECASE)
    if match:
        lines = [line.strip() for line in match.group(1).split('\n') if line.strip()]
        for line in lines:
            if not re.search(r'Page|GSTIN|PAN|State|Address|Phone|Invoice|Date|ACK|IRN', line, re.IGNORECASE) and len(line) > 3:
                return line

    patterns = [
        r'Customer\s*[:\-]\s*([^\n]+)',
        r'(?:Client Name|Customer Name)\s*[:\-]?\s*([^\n]+)',
        r'(?:Name\s*&\s*Address\s*of\s*Bill\s*To|Name\s*and\s*address\s*of\s*the\s*receipient)\s*[\n\r]+\s*([^\n]+)'
    ]
    for pattern in patterns:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            name = m.group(1).strip()
            if not re.search(r'Page\s*:|GSTIN|PAN|State|Address|Phone|Invoice', name, re.IGNORECASE) and len(name) > 3:
                return name

    return "N/A"

def extract_place_of_supply(text):
    """Extracts Place of Supply directly matching state names or codes."""
    match = re.search(r'Place\s*of\s*Supply\s*[:\-]?\s*(\[[0-9]+\]\s*[A-Za-z\s]+|[A-Za-z\s]+)', text, re.IGNORECASE)
    if match:
        pos = match.group(1).strip().split('\n')[0]
        if not re.search(r'AAHCC|GSTIN|PAN', pos, re.IGNORECASE):
            return pos
    return "N/A"

def extract_hsn_sac(text):
    """Extracts HSN/SAC code from line item tables while ignoring PIN codes."""
    match_table = re.search(r'(?:SAC\s*\/\s*HSN|HSN\s*\/\s*SAC|SAC|HSN)\s*[\n\r\s]+([0-9]{4,8})', text, re.IGNORECASE)
    if match_table:
        code = match_table.group(1).strip()
        if not code.startswith("6000") and not code.startswith("1100"):
            return code

    match_inline = re.search(r'(?:SAC\s*Code|HSN\s*Code)\s*[:\-]?\s*([0-9]{4,8})', text, re.IGNORECASE)
    if match_inline:
        return match_inline.group(1).strip()

    sac_fallback = re.search(r'\b(99\d{4})\b', text)
    if sac_fallback:
        return sac_fallback.group(1).strip()

    return "N/A"

def extract_line_item_breakdown(text):
    """
    Parses line item tables to extract Non-Taxable / Exempt Values, Taxable Values,
    CGST, SGST, and IGST line by line.
    """
    non_taxable_exempt = 0.0
    taxable_value = 0.0
    cgst = 0.0
    sgst = 0.0
    igst = 0.0

    # Pattern for Reimbursement / Non GST Exempt Line Items (e.g., CUSTOMS DUTY -REIM - SI)
    reim_items = re.findall(r'(?:REIM|REIMBURSEMENT|DUTY|NON-GST|EXEMPT)[\s\S]*?([\d,]+\.\d{2})', text, re.IGNORECASE)
    
    # Check table structure for "Non GST Exempt Value (INR)" column
    exempt_match = re.search(r'Non\s*GST\s*Exempt\s*Value\s*(?:\(INR\))?[\s\S]*?([\d,]+\.\d{2})', text, re.IGNORECASE)
    if exempt_match:
        non_taxable_exempt = parse_float(exempt_match.group(1))

    # Check Taxable Value column
    taxable_match = re.search(r'Taxable\s*Value\s*(?:\(INR\))?\s*[:\-]?\s*(?:₹|Rs\.?|INR)?\s*([\d,]+\.\d{2})', text, re.IGNORECASE)
    if taxable_match:
        taxable_value = parse_float(taxable_match.group(1))

    # Check IGST column
    igst_match = re.search(r'IGST[\s\S]{1,30}?Tax[\s\S]{1,30}?([\d,]+\.\d{2})', text, re.IGNORECASE)
    if igst_match:
        igst = parse_float(igst_match.group(1))

    # Check CGST / SGST columns
    cgst_match = re.search(r'CGST[\s\S]{1,30}?Tax[\s\S]{1,30}?([\d,]+\.\d{2})', text, re.IGNORECASE)
    if cgst_match:
        cgst = parse_float(cgst_match.group(1))

    sgst_match = re.search(r'SGST[\s\S]{1,30}?Tax[\s\S]{1,30}?([\d,]+\.\d{2})', text, re.IGNORECASE)
    if sgst_match:
        sgst = parse_float(sgst_match.group(1))

    return non_taxable_exempt, taxable_value, cgst, sgst, igst

def extract_sales_invoice_data(pdf_bytes):
    """Extracts required fields from Chakradhara Aerospace invoice formats."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = "".join([page.get_text() for page in doc])
    doc.close()

    # 1. GSTIN Identification
    all_gstins = re.findall(r'\b[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b', text)
    
    my_gstin = "N/A"
    customer_gstin = "N/A"

    for gstin in all_gstins:
        if gstin in COMPANY_GSTINS and my_gstin == "N/A":
            my_gstin = gstin
        elif gstin not in COMPANY_GSTINS and customer_gstin == "N/A":
            customer_gstin = gstin

    if my_gstin == "N/A" and len(all_gstins) > 0:
        my_gstin = all_gstins[0]
    if customer_gstin == "N/A" and len(all_gstins) > 1:
        customer_gstin = all_gstins[1]

    # 2. Customer Name & Place of Supply
    customer_name = extract_customer_name(text)
    place_of_supply = extract_place_of_supply(text)

    # 3. Invoice Number
    inv_no_match = re.search(r'(?:Invoice\s*No\.?|Inv\s*No\.?|INVOICE\040NO|Invoice\s*#)\s*[:\-]?\s*([A-Za-z0-9\/\-]+)', text, re.IGNORECASE)
    invoice_no = inv_no_match.group(1).strip() if inv_no_match else "N/A"

    # 4. Invoice Date
    inv_date_match = re.search(r'(?:Invoice\s*Date|Inv\.\s*Date|DATE|Dated|Date)\s*[:\-]?\s*(\d{1,2}[\/\.-]\w+[\/\.-]\d{2,4}|\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})', text, re.IGNORECASE)
    raw_date = inv_date_match.group(1).strip() if inv_date_match else "N/A"
    invoice_date = parse_date(raw_date)

    # 5. HSN / SAC Code
    hsn_code = extract_hsn_sac(text)

    # 6. Extract Line Item Tax Breakdown
    non_taxable_exempt, taxable_value, cgst, sgst, igst = extract_line_item_breakdown(text)

    total_tax = round(cgst + sgst + igst, 2)

    # 7. Total Invoice Value
    total_match = re.search(r'Total\s*(?:\(INR\)|Value|Amount)?\s*[:\-]?\s*(?:₹|Rs\.?|INR)?\s*([\d,]+\.\d{2})', text, re.IGNORECASE)
    if total_match:
        total_invoice_value = parse_float(total_match.group(1))
    else:
        total_invoice_value = round(taxable_value + non_taxable_exempt + total_tax, 2)

    return {
        "GSTIN": my_gstin,
        "Customer Name": customer_name,
        "Customer GSTIN": customer_gstin,
        "Invoice No": invoice_no,
        "Invoice Date": invoice_date,
        "Place of Supply": place_of_supply,
        "HSN/SAC": hsn_code,
        "Non-Taxable / Exempt Value": non_taxable_exempt,
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

    # Style Header Row
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    num_cols = {"Non-Taxable / Exempt Value", "Taxable Value", "CGST", "SGST", "IGST", "Total Tax", "Total Invoice Value"}
    
    # Append Data Rows
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
        
        cols = ["Filename", "GSTIN", "Customer Name", "Customer GSTIN", "Invoice No", "Invoice Date", 
                "Place of Supply", "HSN/SAC", "Non-Taxable / Exempt Value", "Taxable Value", 
                "CGST", "SGST", "IGST", "Total Tax", "Total Invoice Value"]
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