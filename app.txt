import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Sales Invoice Extractor", page_icon="🧾", layout="wide")

st.title("🧾 Sales Invoice Data Extractor to Excel")
st.write("Upload sales invoice PDFs to extract structured tax fields directly into a styled Excel file.")

uploaded_files = st.file_uploader("Upload Sales Invoice PDFs", type=["pdf"], accept_multiple_files=True)

def extract_sales_invoice_data(pdf_bytes):
    """Extracts sales invoice tax fields from PDF text using optimized regex patterns."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = "".join([page.get_text() for page in doc])
    doc.close()

    # Regex patterns for Indian Sales Invoices
    gstin_pattern = r'\b[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b'
    gstin_matches = re.findall(gstin_pattern, text)
    
    # Identify My GSTIN vs Customer GSTIN based on occurrences
    my_gstin = gstin_matches[0] if len(gstin_matches) > 0 else "N/A"
    customer_gstin = gstin_matches[1] if len(gstin_matches) > 1 else (gstin_matches[0] if len(gstin_matches) == 1 else "N/A")

    # Extract Invoice Number
    inv_no_match = re.search(r'(?:Invoice\s*No\.?|Inv\s*No\.?|Invoice\s*#)\s*[:\-]?\s*([A-Za-z0-9\/\-]+)', text, re.IGNORECASE)
    invoice_no = inv_no_match.group(1).strip() if inv_no_match else "N/A"

    # Extract Invoice Date
    inv_date_match = re.search(r'(?:Invoice\s*Date|Dated|Date)\s*[:\-]?\s*(\d{2}[\/\.-]\d{2}[\/\.-]\d{2,4}|\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})', text, re.IGNORECASE)
    invoice_date = inv_date_match.group(1).strip() if inv_date_match else "N/A"

    # Extract Place of Supply
    pos_match = re.search(r'(?:Place\s*of\s*Supply|State\s*Name|POS)\s*[:\-]?\s*([A-Za-z\s]+(?:\(\d+\))?)', text, re.IGNORECASE)
    place_of_supply = pos_match.group(1).strip().split('\n')[0] if pos_match else "N/A"

    # Extract Tax Amounts & Percentages
    taxable_val_match = re.search(r'(?:Taxable\s*Value|Sub\s*Total|Taxable\s*Amount)\s*[:\-]?\s*(?:₹|Rs\.?)?\s*([\d,]+\.\d{2}|\d+)', text, re.IGNORECASE)
    taxable_value = taxable_val_match.group(1).replace(',', '') if taxable_val_match else "N/A"

    cgst_match = re.search(r'(?:CGST)\s*[:\-]?\s*(?:₹|Rs\.?)?\s*([\d,]+\.\d{2}|\d+)', text, re.IGNORECASE)
    cgst = cgst_match.group(1).replace(',', '') if cgst_match else "0.00"

    sgst_match = re.search(r'(?:SGST|UTGST)\s*[:\-]?\s*(?:₹|Rs\.?)?\s*([\d,]+\.\d{2}|\d+)', text, re.IGNORECASE)
    sgst = sgst_match.group(1).replace(',', '') if sgst_match else "0.00"

    igst_match = re.search(r'(?:IGST)\s*[:\-]?\s*(?:₹|Rs\.?)?\s*([\d,]+\.\d{2}|\d+)', text, re.IGNORECASE)
    igst = igst_match.group(1).replace(',', '') if igst_match else "0.00"

    tax_rate_match = re.search(r'(\d{1,2}(?:\.\d+)?)\s*%', text)
    tax_percentage = f"{tax_rate_match.group(1)}%" if tax_rate_match else "N/A"

    total_tax_match = re.search(r'(?:Total\s*Tax|Tax\s*Amount)\s*[:\-]?\s*(?:₹|Rs\.?)?\s*([\d,]+\.\d{2}|\d+)', text, re.IGNORECASE)
    total_tax = total_tax_match.group(1).replace(',', '') if total_tax_match else "N/A"

    total_inv_val_match = re.search(r'(?:Grand\s*Total|Invoice\s*Total|Total\s*Amount|Total)\s*[:\-]?\s*(?:₹|Rs\.?)?\s*([\d,]+\.\d{2}|\d+)', text, re.IGNORECASE)
    total_invoice_value = total_inv_val_match.group(1).replace(',', '') if total_inv_val_match else "N/A"

    return {
        "My Company GSTIN": my_gstin,
        "Customer GSTIN": customer_gstin,
        "Invoice No": invoice_no,
        "Invoice Date": invoice_date,
        "Place of Supply": place_of_supply,
        "Taxable Value": taxable_value,
        "CGST": cgst,
        "SGST": sgst,
        "IGST": igst,
        "Tax Percentage": tax_percentage,
        "Total Tax": total_tax,
        "Total Invoice Value": total_invoice_value,
    }

def format_excel_workbook(df):
    """Generates an Excel file with specific header styling, alignments, auto-filters, and auto-fitted columns."""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales Invoices"

    # Write Headers
    headers = list(df.columns)
    ws.append(headers)

    # Format Header Row (Bold, Center Aligned)
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Write Data Rows (Left Aligned)
    for row in df.itertuples(index=False):
        ws.append(list(row))

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal="left", vertical="center")

    # Enable Auto Filter across all columns
    ws.auto_filter.ref = ws.dimensions

    # Auto-fit Column Widths universally for all columns
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max(max_length + 4, 12)

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
        
        # Reorder columns with Filename first
        cols = ["Filename"] + [c for c in df.columns if c != "Filename"]
        df = df[cols]

        st.success(f"Processed {len(results)} sales invoice files successfully!")
        st.dataframe(df)

        # Generate styled Excel file
        excel_bytes = format_excel_workbook(df)

        st.download_button(
            label="📥 Download Structured Excel File",
            data=excel_bytes,
            file_name="Sales_Invoices_Extracted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )