import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# OCR fallback for scanned / image-only PDFs (no embedded text layer).
# Optional: if pytesseract or the tesseract binary aren't installed, the app
# still runs fine -- it just can't rescue image-only pages, and will flag
# them in the "Extraction Notes" column instead of silently returning N/A.
try:
    import pytesseract
    from PIL import Image
    _OCR_AVAILABLE = True
except ImportError:
    _OCR_AVAILABLE = False

st.set_page_config(page_title="Chakradhara Aerospace - Sales Invoice Extractor", page_icon="✈️", layout="wide")

st.title("✈️ Chakradhara Aerospace - Sales Invoice Extractor")
st.write("Upload sales invoice PDFs across all formats to extract structured details and export to formatted Excel.")

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

GSTIN_RE = re.compile(r'\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b')

# Standard GST state codes (first two digits of any GSTIN) -> state/UT name.
# Used as a last-resort fallback to derive Place of Supply when the invoice
# text has no explicit "Place of Supply" / "State" label with a value.
GST_STATE_CODES = {
    "01": "Jammu and Kashmir", "02": "Himachal Pradesh", "03": "Punjab",
    "04": "Chandigarh", "05": "Uttarakhand", "06": "Haryana", "07": "Delhi",
    "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim",
    "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur",
    "15": "Mizoram", "16": "Tripura", "17": "Meghalaya", "18": "Assam",
    "19": "West Bengal", "20": "Jharkhand", "21": "Odisha",
    "22": "Chhattisgarh", "23": "Madhya Pradesh", "24": "Gujarat",
    "25": "Daman and Diu", "26": "Dadra and Nagar Haveli",
    "27": "Maharashtra", "28": "Andhra Pradesh (Old)", "29": "Karnataka",
    "30": "Goa", "31": "Lakshadweep", "32": "Kerala", "33": "Tamil Nadu",
    "34": "Puducherry", "35": "Andaman and Nicobar Islands",
    "36": "Telangana", "37": "Andhra Pradesh", "38": "Ladakh",
}

NAME_CONTINUATION_RE = re.compile(
    r'^(PRIVATE\s+LIMITED|LIMITED|LTD\.?|PVT\.?\s*LTD\.?|\(P\)\s*LTD\.?|CO\.?\s*LTD\.?|PRIVATE\s+LTD\.?)$',
    re.IGNORECASE,
)

# ---------------------------------------------------------------------------
# Basic helpers
# ---------------------------------------------------------------------------

def parse_float(val):
    """Safely extracts a clean float from text like '1,23,456.78', '₹ 1,234.00' or '-'."""
    if val is None:
        return 0.0
    s = str(val).strip()
    if not s or s.upper() == "N/A" or s == "-":
        return 0.0
    cleaned = re.sub(r'[^\d\.\-]', '', s)
    if cleaned in ("", "-", ".", "-."):
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def parse_date(date_str):
    """Converts extracted text dates into datetime objects, trying multiple formats."""
    if not date_str:
        return None

    date_patterns = [
        r'\d{1,2}[\/\.-]\d{1,2}[\/\.-]\d{4}',
        r'\d{1,2}[\/\.-]\d{1,2}[\/\.-]\d{2}',
        r'\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}',
        r'\d{1,2}[\/\.-][A-Za-z]{3}[\/\.-]\d{2,4}',
    ]
    date_formats = [
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
        "%d/%m/%y", "%d-%m-%y", "%d.%m.%y",
        "%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d-%b-%y",
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


def clean_text(text):
    """
    Strips digital-signature appearance-stream garbage that PyMuPDF sometimes
    leaks into extracted text as one giant unbroken token (this never occurs
    in normal invoice text, which is short words/numbers separated by
    whitespace).
    """
    return re.sub(r'\S{60,}', ' ', text)


def _ocr_extract(doc):
    """Rasterizes each page at high DPI and runs Tesseract OCR over it.
    Used only when the PDF has little or no embedded text (i.e. it's a
    scanned image / photo of the invoice rather than a digitally
    generated one)."""
    if not _OCR_AVAILABLE:
        return ""
    chunks = []
    for page in doc:
        try:
            pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))  # ~300 DPI
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            chunks.append(pytesseract.image_to_string(img))
        except Exception:
            continue
    return "\n".join(chunks)


def get_pdf_text(pdf_bytes):
    """
    Extracts text from the PDF, trying multiple strategies and picking
    whichever yields the most usable result:

    1. PyMuPDF with sort=True, which reorders text into natural reading
       order. Without this, PyMuPDF returns text in internal PDF-object
       order, which for table-heavy invoices comes out scrambled (numbers
       from unrelated rows/columns interleave with each other).
    2. PyMuPDF with sort=False (default order) -- for a minority of
       templates the raw object order actually keeps label/value pairs
       together better than the sort heuristic does.
    3. OCR (Tesseract), used only as a last resort when neither of the
       above produced meaningful text. This is what rescues invoices that
       are scanned/photographed rather than digitally generated -- those
       have no embedded text layer at all, so PyMuPDF returns next to
       nothing and every field would otherwise come back "N/A".

    Returns a tuple: (text, extraction_method) so the caller can record
    how the text was obtained (useful for the "Extraction Notes" column).
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    text_sorted = clean_text("\n".join(page.get_text(sort=True) for page in doc))
    text_plain = clean_text("\n".join(page.get_text(sort=False) for page in doc))

    if len(text_sorted.strip()) >= len(text_plain.strip()):
        best_text, method = text_sorted, "text (sorted)"
    else:
        best_text, method = text_plain, "text (raw order)"

    # Fewer than ~40 non-whitespace characters across the whole document
    # means there's essentially no embedded text -- almost certainly a
    # scanned/photographed page rather than a digitally generated PDF.
    if len(re.sub(r'\s', '', best_text)) < 40:
        ocr_text = clean_text(_ocr_extract(doc))
        if len(re.sub(r'\s', '', ocr_text)) > len(re.sub(r'\s', '', best_text)):
            best_text, method = ocr_text, ("OCR" if ocr_text.strip() else "none (no text found)")
        elif not best_text.strip():
            method = "none (no text found)"

    doc.close()
    return best_text, method


# ---------------------------------------------------------------------------
# Header field extraction
# ---------------------------------------------------------------------------

def extract_gstins(text):
    """Finds all GSTIN-like strings and splits into company GSTIN vs customer GSTIN."""
    all_gstins = GSTIN_RE.findall(text)
    my_gstin, customer_gstin = "N/A", "N/A"
    for g in all_gstins:
        if g in COMPANY_GSTINS and my_gstin == "N/A":
            my_gstin = g
        elif g not in COMPANY_GSTINS and customer_gstin == "N/A":
            customer_gstin = g
    if my_gstin == "N/A" and all_gstins:
        my_gstin = all_gstins[0]
    if customer_gstin == "N/A":
        remaining = [g for g in all_gstins if g != my_gstin]
        if remaining:
            customer_gstin = remaining[0]
    return my_gstin, customer_gstin


def _first_segment(s):
    """Takes the text before the first run of 2+ spaces (i.e. before the next
    visual 'column' on the same printed line), trimming a trailing comma."""
    s = s.strip()
    if not s:
        return ''
    return re.split(r'\s{2,}', s)[0].strip().rstrip(',')


def _looks_like_label_junk(s):
    """True if a candidate name is actually a neighbouring field label or
    table-column header that got pulled onto the same line/position by the
    PDF's column layout (very common in freight/GTA consignment-note
    templates, where headers like "Billing Remarks", "LR No", "Vehicle No"
    etc. sit right next to the "Name and Address of Recipient" block once
    the text is reordered)."""
    return bool(re.search(
        r':|Invoice\s*No|Invoice\s*Date|\bGSTIN\b|\bPAN\b|Ref\.\s*No|Division\s*:|\bAck\b|'
        r'State\s*:|Address\s*:|^Date\b|Delivery|Ship\s*To|Address\s*Of|'
        r'Billing\s*Remarks|^Remarks$|LR\s*No|L\.?R\.?\s*No|Vehicle\s*No|Driver|'
        r'Consignment|Freight|E-?way|Bilty|Description|Particulars|'
        r'Weight|Quantity|^Rate\b|^Amount\b|\bHSN\b|\bSAC\b|Bank|IFSC|'
        r'Terms\s*of|Bill\s*Type|Reverse\s*Charge',
        s, re.IGNORECASE,
    ))


def extract_customer_name(text):
    lines = text.split('\n')
    label_res = [
        re.compile(r'^\s*Customer\s*-?\s*:\s*(.*)$', re.IGNORECASE),
        re.compile(r'^\s*Client\s*Name\s*:\s*(.*)$', re.IGNORECASE),
    ]
    for i, raw_line in enumerate(lines):
        for lr in label_res:
            m = lr.match(raw_line)
            if not m:
                continue
            candidate = _first_segment(m.group(1))
            if candidate and _looks_like_label_junk(candidate):
                candidate = ''
            if not candidate:
                # Value wasn't on this line (e.g. "Customer :" alone) -- the
                # actual name usually appears on the next non-blank line.
                for j in range(i + 1, min(i + 3, len(lines))):
                    nxt = lines[j].strip()
                    if not nxt:
                        continue
                    seg = _first_segment(nxt)
                    if seg and len(seg) > 3 and not _looks_like_label_junk(seg):
                        candidate = seg
                    break
            if candidate and len(candidate) > 3:
                # Some names wrap onto a second line (e.g. "...COMPANY" /
                # "PRIVATE LIMITED") -- reattach a short recognizable suffix.
                for j in range(i + 1, min(i + 3, len(lines))):
                    cont_seg = _first_segment(lines[j])
                    if cont_seg and NAME_CONTINUATION_RE.match(cont_seg):
                        candidate = f"{candidate} {cont_seg}"
                    break
                return candidate

    # Fallback for formats with no "Customer:" label at all.
    m = re.search(r'Name\s*(?:&|and)?\s*Address\s*of\s*the\s*rec[ei]{1,2}pient', text, re.IGNORECASE)
    if m:
        for line in text[m.end():].split('\n'):
            line = line.strip()
            if not line:
                continue
            seg = _first_segment(line)
            if seg and len(seg) > 3 and not _looks_like_label_junk(seg):
                return seg
    return "N/A"


def extract_invoice_no(text):
    """Finds the invoice number after an "Invoice No" / "Inv No" label.

    Some templates have another short field (e.g. "Bill Type: B") sitting
    right next to "Invoice No" once the PDF's columns get flattened into a
    single text stream, so the very next token after the label isn't
    always the real invoice number -- it can be a stray single letter.
    To guard against that, every token on the rest of that line is
    considered and the first one that actually looks like an invoice
    number (at least 3 characters, contains a digit) is preferred; only if
    nothing on the line qualifies do we fall back to the first token."""
    label_re = re.compile(
        r'(?:Invoice\s*No\.?(?!umber)|Inv\.?\s*No\.?|INVOICE\s*NO)\s*[:\-]?\s*',
        re.IGNORECASE,
    )
    fallback = None
    for m in label_re.finditer(text):
        # Look at the rest of this line, and -- in case the label and its
        # value ended up on separate lines -- the following couple of
        # lines too.
        following_lines = text[m.end():].split('\n')[:3]
        rest_of_line = following_lines[0]
        tokens = re.findall(r'[A-Za-z0-9][\w\/\-]*', rest_of_line)
        if not tokens:
            continue
        if fallback is None:
            fallback = tokens[0]
        good = next((t for t in tokens if len(t) >= 3 and re.search(r'\d', t)), None)
        if good:
            return good
        for extra_line in following_lines[1:]:
            extra_tokens = re.findall(r'[A-Za-z0-9][\w\/\-]*', extra_line.strip())
            good = next((t for t in extra_tokens if len(t) >= 3 and re.search(r'\d', t)), None)
            if good:
                return good
    return fallback if fallback else "N/A"


def extract_invoice_date(text):
    # Prefer an explicit "Invoice Date" label.
    m = re.search(
        r'Invoice\s*Date\s*[:\-]?\s*(\d{1,2}[\/\.\- ][A-Za-z0-9]{1,9}[\/\.\- ]\d{2,4})',
        text, re.IGNORECASE,
    )
    if m:
        return parse_date(m.group(1))

    # Otherwise fall back to a generic "Date" label, skipping "Ack Date",
    # "Due Date", "Doc. Date", "LR Date" etc. which are different fields.
    for m in re.finditer(
        r'\bDate\b\s*[:\-]?\s*(\d{1,2}[\/\.\- ][A-Za-z0-9]{1,9}[\/\.\- ]\d{2,4})',
        text, re.IGNORECASE,
    ):
        preceding = text[max(0, m.start() - 5):m.start()]
        if re.search(r'Ack\s*$|Due\s*$|Doc\.\s*$|LR\s*$|Ref\s*$', preceding, re.IGNORECASE):
            continue
        return parse_date(m.group(1))
    return None


def extract_place_of_supply(text, customer_gstin):
    m = re.search(
        r'Place\s*of\s*Supply\s*[:\-]?\s*(\[?\d{0,2}\]?\s*[A-Za-z][A-Za-z\s]+)',
        text, re.IGNORECASE,
    )
    if m:
        val = m.group(1).strip().split('\n')[0].strip()
        if not re.search(r'AAHCC|GSTIN|PAN', val, re.IGNORECASE):
            return val

    # Fallback: look for a "State" mention near the customer's GSTIN.
    if customer_gstin and customer_gstin != "N/A":
        idx = text.find(customer_gstin)
        if idx != -1:
            window = text[max(0, idx - 150):idx + 150]
            m2 = re.search(
                r'State\s*(?:Code)?(?:\s*&\s*Name)?\s*:?\s*(?:\d{1,2}\s*-\s*)?([A-Za-z][A-Za-z\s]{2,30})',
                window, re.IGNORECASE,
            )
            if m2:
                val = _first_segment(m2.group(1))
                if val and val.lower() not in ("code", "name", "address", "gstin") \
                        and not re.search(r'GSTIN|Invoice|Name', val, re.IGNORECASE):
                    return val

    # Final fallback: derive the state purely from the customer's GSTIN
    # state-code prefix (first 2 digits) -- reliable even when the invoice
    # template prints a blank "State Code & Name:" field.
    if customer_gstin and customer_gstin != "N/A" and len(customer_gstin) >= 2:
        return GST_STATE_CODES.get(customer_gstin[:2], "N/A")
    return "N/A"


def extract_hsn(text):
    """Tries, in order: an explicit HSN/SAC label; a bare service SAC code
    (these commonly start with '99'); a comma/decimal-mangled code some
    invoice templates print (e.g. '9,96,791.00' meaning HSN 996791); finally
    a bare 8-digit goods HSN code."""
    m = re.search(r'(?:HSN|SAC)\s*(?:Code)?\s*:?\s*(\d{4,8})\b', text, re.IGNORECASE)
    if m:
        return m.group(1)

    candidates = re.findall(r'\b(99\d{2,6})\b', text)
    if candidates:
        six_digit = [c for c in candidates if len(c) == 6]
        return six_digit[0] if six_digit else candidates[0]

    m = re.search(r'\b(\d{1,2},\d{2},\d{3})\.00\b', text)
    if m:
        return m.group(1).replace(',', '')

    m = re.search(r'\b(\d{8})\b', text)
    if m:
        return m.group(1)

    return "N/A"


# ---------------------------------------------------------------------------
# Tax figure extraction
# ---------------------------------------------------------------------------

def _find_vertical_tax_amount(text, keyword):
    """
    Fallback for the 'vertical rate-of-tax' layout used by several GTA/freight
    templates, where CGST/SGST/IGST each get their own row with a % rate then
    an amount (amount may be '-' meaning nil, common under reverse charge).
    """
    m = re.search(
        rf'\b{keyword}\b\s*[:\-]?((?:\s*(?:[\d,]+\.?\d*%?|-)){{1,4}})',
        text, re.IGNORECASE,
    )
    if not m:
        return 0.0
    tokens = re.findall(r'[\d,]+\.?\d*%?|-', m.group(1))
    if not tokens:
        return 0.0
    last = tokens[-1]
    if last == '-':
        return 0.0
    return parse_float(last.rstrip('%'))


def extract_tax_details(text):
    result = {"taxable": 0.0, "cgst": 0.0, "sgst": 0.0, "igst": 0.0}

    # Layout family 1: a "Sub Total" row that lists Taxable Value followed by
    # CGST/SGST (or IGST) followed by a running total, all on one line.
    sub_total_m = re.search(
        r'Sub\s*Total\s*(?:\[INR\])?\s*((?:[\d,]+\.\d{2}\s*){1,4})',
        text, re.IGNORECASE,
    )
    if sub_total_m:
        nums = [parse_float(n) for n in re.findall(r'[\d,]+\.\d{2}', sub_total_m.group(1))]
        if len(nums) >= 4:
            result["taxable"], result["cgst"], result["sgst"] = nums[0], nums[1], nums[2]
        elif len(nums) == 3:
            result["taxable"], result["igst"] = nums[0], nums[1]
        elif len(nums) >= 1:
            result["taxable"] = nums[0]
    else:
        # Layout family 2: a single "Total Taxable Amount" figure with
        # CGST/SGST/IGST given separately (often as a vertical rate/amount
        # table, frequently zero under GTA reverse-charge).
        taxable_m = re.search(
            r'Total\s*Taxable\s*Amount\s*[:\-]?\s*([\d,]+\.?\d{0,2})',
            text, re.IGNORECASE,
        )
        if taxable_m:
            result["taxable"] = parse_float(taxable_m.group(1))
        else:
            # Layout family 3: a single-row SAC/Taxable-Value table, e.g.
            # "996791   58700.00   ...   0 0 0 0"
            m = re.search(r'\b\d{6}\b\s+([\d,]+\.\d{2})', text)
            if m:
                result["taxable"] = parse_float(m.group(1))

        result["cgst"] = _find_vertical_tax_amount(text, "CGST")
        result["sgst"] = _find_vertical_tax_amount(text, "SGST")
        result["igst"] = _find_vertical_tax_amount(text, "IGST")

    total_tax = round(result["cgst"] + result["sgst"] + result["igst"], 2)

    # ---- Grand total, tried in priority order across the templates -------
    grand_total = None
    m = re.search(
        r'Total\s*Invoice\s*Value\s*(?:\(?INR\)?)?\s*[:\-]?\s*(?:INR)?\s*([\d,]+\.\d{1,2})',
        text, re.IGNORECASE,
    )
    if m:
        grand_total = parse_float(m.group(1))
    if grand_total is None:
        m = re.search(r'Gross\s*Total\s*[:\-]?\s*([\d,]+\.?\d{0,2})', text, re.IGNORECASE)
        if m:
            grand_total = parse_float(m.group(1))
    if grand_total is None:
        matches = re.findall(r'Total\s*\[INR\]\s*([\d,]+\.\d{2})', text, re.IGNORECASE)
        if matches:
            grand_total = parse_float(matches[-1])
    if grand_total is None:
        m_all = list(re.finditer(r'Total\s*Amount\s*\[INR\]\s*((?:[\d,]+\.\d{2}\s*)+)', text, re.IGNORECASE))
        if m_all:
            nums = re.findall(r'[\d,]+\.\d{2}', m_all[-1].group(1))
            if nums:
                grand_total = parse_float(nums[-1])

    # Non-Taxable / Exempt Value is picked up ONLY when the invoice
    # explicitly labels an amount as Non-Taxable, Exempt(ed), Nil-Rated or
    # Non-GST -- never inferred just from the document being titled "Bill
    # of Supply", "Reimbursement Invoice", etc. Those document types can
    # still carry GST payable by the recipient under reverse charge (as in
    # a GTA Bill of Supply), so the line-item value itself is a normal
    # taxable amount, not an exempt one, unless the invoice says so.
    non_taxable_m = re.search(
        r'(?:Non[-\s]?Taxable|Nil[-\s]?Rated|Exempt(?:ed)?|Non[-\s]?GST(?:\s*Supply)?)'
        r'\s*(?:Value|Amount)?\s*(?:\[INR\])?\s*[:\-]?\s*(?:INR)?\s*([\d,]+\.\d{1,2})',
        text, re.IGNORECASE,
    )
    non_taxable = parse_float(non_taxable_m.group(1)) if non_taxable_m else 0.0

    taxable = result["taxable"]

    # Layout family 4: some templates (e.g. Credit Memos, Bills of Supply)
    # print only a "Line Total [INR]" / "Total [INR]" figure with no
    # separate taxable-value or tax breakdown at all. If nothing else
    # matched and there's no tax and no explicit non-taxable/exempt wording,
    # the grand total itself IS the taxable value by default.
    if taxable == 0.0 and non_taxable == 0.0 and total_tax == 0.0 and grand_total:
        taxable = grand_total

    if grand_total is None:
        grand_total = round(taxable + non_taxable + total_tax, 2)

    return {
        "non_taxable": round(non_taxable, 2),
        "taxable": round(taxable, 2),
        "cgst": round(result["cgst"], 2),
        "sgst": round(result["sgst"], 2),
        "igst": round(result["igst"], 2),
        "total_tax": total_tax,
        "grand_total": round(grand_total, 2),
    }


# ---------------------------------------------------------------------------
# Invoice type / document type extraction
# ---------------------------------------------------------------------------

# Checked in priority order (most specific document types first) so that,
# e.g., a "Bill of Supply" that happens to mention "invoice" elsewhere in
# its body doesn't get misclassified as a generic "Tax Invoice".
_INVOICE_TYPE_PATTERNS = [
    (re.compile(r'\bBill\s*Of\s*Supply\b', re.IGNORECASE), "Bill of Supply"),
    (re.compile(r'\bReimbursement\s*Invoice\b', re.IGNORECASE), "Reimbursement Invoice"),
    (re.compile(r'\bCredit\s*Memo\b', re.IGNORECASE), "Credit Memo"),
    (re.compile(r'\bCredit\s*Note\b', re.IGNORECASE), "Credit Note"),
    (re.compile(r'\bDebit\s*Memo\b', re.IGNORECASE), "Debit Memo"),
    (re.compile(r'\bDebit\s*Note\b', re.IGNORECASE), "Debit Note"),
    (re.compile(r'\bProforma\s*Invoice\b', re.IGNORECASE), "Proforma Invoice"),
    (re.compile(r'\bDelivery\s*Challan\b', re.IGNORECASE), "Delivery Challan"),
    (re.compile(r'\bTax\s*Invoice\b', re.IGNORECASE), "Tax Invoice"),
]


def extract_invoice_type(text):
    """Identifies the document type (Tax Invoice / Bill of Supply / Credit
    Memo / etc.) from the heading printed at the top of the document.
    Checks only the first ~800 characters first, since that's where the
    document title is always printed; falls back to searching the whole
    text in case the heading landed further down after text reordering."""
    head = text[:800]
    for pattern, label in _INVOICE_TYPE_PATTERNS:
        if pattern.search(head):
            return label
    for pattern, label in _INVOICE_TYPE_PATTERNS:
        if pattern.search(text):
            return label
    if re.search(r'\bINVOICE\b', text, re.IGNORECASE):
        return "Invoice"
    return "N/A"


# ---------------------------------------------------------------------------
# Main per-invoice extraction
# ---------------------------------------------------------------------------

def extract_sales_invoice_data(pdf_bytes):
    """Extracts required fields from a Chakradhara Aerospace sales invoice PDF."""
    text, extraction_method = get_pdf_text(pdf_bytes)

    my_gstin, customer_gstin = extract_gstins(text)
    customer_name = extract_customer_name(text)
    invoice_type = extract_invoice_type(text)
    invoice_no = extract_invoice_no(text)
    invoice_date = extract_invoice_date(text)
    place_of_supply = extract_place_of_supply(text, customer_gstin)
    hsn_code = extract_hsn(text)
    tax = extract_tax_details(text)

    fields = [my_gstin, customer_name, customer_gstin, invoice_type, invoice_no, place_of_supply, hsn_code]
    na_count = sum(1 for f in fields if f == "N/A") + (1 if invoice_date is None else 0)

    if extraction_method == "none (no text found)":
        note = "No text extracted -- likely a scanned/photographed PDF; install OCR support or re-export the invoice as a text PDF."
    elif extraction_method == "OCR":
        note = "Recovered via OCR (scanned PDF) -- please double-check the values."
    elif na_count >= 5:
        note = "Most fields missing -- this invoice's layout may not match any known template. Check manually."
    elif na_count > 0:
        note = "Some fields missing -- check manually."
    else:
        note = "OK"

    return {
        "GSTIN": my_gstin,
        "Customer Name": customer_name,
        "Customer GSTIN": customer_gstin,
        "Invoice Type": invoice_type,
        "Invoice No": invoice_no,
        "Invoice Date": invoice_date,
        "Place of Supply": place_of_supply,
        "HSN/SAC": hsn_code,
        "Non-Taxable / Exempt Value": tax["non_taxable"],
        "Taxable Value": tax["taxable"],
        "CGST": tax["cgst"],
        "SGST": tax["sgst"],
        "IGST": tax["igst"],
        "Total Tax": tax["total_tax"],
        "Total Invoice Value": tax["grand_total"],
        "Extraction Notes": note,
        "_raw_text": text,
    }


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def format_excel_workbook(df):
    """Formats Excel file: bold + center-aligned header with auto-filter,
    left-aligned data rows, and dates shown as dd/mm/yyyy."""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales Invoices"

    headers = list(df.columns)
    ws.append(headers)

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    num_cols = {"Non-Taxable / Exempt Value", "Taxable Value", "CGST", "SGST", "IGST", "Total Tax", "Total Invoice Value"}

    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, (col_name, val) in enumerate(zip(headers, row), start=1):
            cell = ws.cell(row=row_idx, column=col_idx)

            if col_name == "Invoice Date" and pd.notnull(val):
                cell.value = val
                cell.number_format = 'dd/mm/yyyy'
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_name in num_cols:
                cell.value = float(val) if val is not None else 0.0
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.value = str(val) if val is not None else ""
                cell.alignment = Alignment(horizontal="left", vertical="center")

    ws.auto_filter.ref = ws.dimensions

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


# ---------------------------------------------------------------------------
# Streamlit app
# ---------------------------------------------------------------------------

if "extracted_df" not in st.session_state:
    st.session_state.extracted_df = None
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "raw_text_by_file" not in st.session_state:
    st.session_state.raw_text_by_file = {}

if not _OCR_AVAILABLE:
    st.info(
        "OCR support isn't installed, so scanned/photographed invoices with no "
        "embedded text layer can't be recovered automatically (they'll show "
        "'No text extracted' in Extraction Notes). To enable OCR, add "
        "`pytesseract` and `pillow` to requirements.txt and `tesseract-ocr` "
        "to packages.txt (if deploying on Streamlit Community Cloud).",
        icon="ℹ️",
    )


def clear_all():
    """Resets the app back to its initial state: clears results and the
    file uploader (bumping its widget key forces Streamlit to recreate it
    empty)."""
    st.session_state.extracted_df = None
    st.session_state.excel_bytes = None
    st.session_state.raw_text_by_file = {}
    st.session_state.uploader_key += 1


uploaded_files = st.file_uploader(
    "Upload Sales Invoice PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_key}",
)

if uploaded_files:
    col1, col2 = st.columns([1, 1])
    with col1:
        extract_clicked = st.button("Extract Invoice Details", type="primary")
    with col2:
        st.button("🔄 Clear", on_click=clear_all)

    if extract_clicked:
        results = []
        for file in uploaded_files:
            pdf_bytes = file.read()
            extracted_info = extract_sales_invoice_data(pdf_bytes)
            extracted_info["Filename"] = file.name
            results.append(extracted_info)

        df = pd.DataFrame(results)

        # Keep the raw extracted text alongside the results (for the
        # in-app debug viewer below) but never write it into the Excel
        # export -- it's only there to help diagnose problem invoices.
        raw_text_by_file = dict(zip(df["Filename"], df["_raw_text"]))
        st.session_state.raw_text_by_file = raw_text_by_file

        cols = ["Filename", "GSTIN", "Customer Name", "Customer GSTIN", "Invoice Type", "Invoice No", "Invoice Date",
                "Place of Supply", "HSN/SAC", "Non-Taxable / Exempt Value", "Taxable Value",
                "CGST", "SGST", "IGST", "Total Tax", "Total Invoice Value", "Extraction Notes"]
        df = df[cols]

        st.session_state.extracted_df = df
        st.session_state.excel_bytes = format_excel_workbook(df)
elif st.session_state.extracted_df is not None:
    # Files were cleared from the uploader but results are still held from a
    # previous run -- offer Clear so the person can fully reset the page.
    st.button("🔄 Clear", on_click=clear_all)

if st.session_state.extracted_df is not None:
    df_display = st.session_state.extracted_df
    flagged = df_display[df_display["Extraction Notes"] != "OK"]
    if len(flagged) > 0:
        st.warning(
            f"{len(flagged)} of {len(df_display)} invoice(s) have missing or "
            "uncertain fields -- see the 'Extraction Notes' column below, and "
            "use 'Inspect raw extracted text' to see exactly what was pulled "
            "from the PDF for a given file."
        )
    else:
        st.success(f"Processed {len(df_display)} invoices successfully!")
    st.dataframe(df_display)

    with st.expander("🔍 Inspect raw extracted text (for troubleshooting)"):
        file_choice = st.selectbox(
            "Choose a file to inspect", list(st.session_state.raw_text_by_file.keys())
        )
        if file_choice:
            raw = st.session_state.raw_text_by_file.get(file_choice, "")
            st.caption(f"{len(raw)} characters extracted")
            st.text_area("Extracted text", raw or "(nothing was extracted from this PDF)", height=300)

    st.download_button(
        label="📥 Download Structured Excel File",
        data=st.session_state.excel_bytes,
        file_name="Chakradhara_Sales_Invoices.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
