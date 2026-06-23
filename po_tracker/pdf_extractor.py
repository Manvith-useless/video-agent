"""
Semi-automatic PDF extraction for POs and Invoices.
Returns best-effort extracted fields; user must verify before saving.
"""
import re
import pdfplumber


def extract_text(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text


def _find(pattern, text, group=1, flags=re.IGNORECASE):
    m = re.search(pattern, text, flags)
    return m.group(group).strip() if m else ""


def _parse_quantity(val):
    if not val:
        return 0
    cleaned = re.sub(r"[^\d.]", "", val)
    try:
        return float(cleaned)
    except ValueError:
        return 0


def extract_po_data(pdf_path):
    """Extract PO fields from PDF text. Returns dict of best-guess values."""
    text = extract_text(pdf_path)

    po_number = (
        _find(r"P\.?O\.?\s*(?:No|Number|#)[:\s]*([A-Z0-9\-/]+)", text)
        or _find(r"Purchase\s+Order\s*[:\s]*([A-Z0-9\-/]+)", text)
        or _find(r"PO[:\s#]*([A-Z0-9\-/]+)", text)
    )

    po_date = (
        _find(r"(?:PO\s+Date|Order\s+Date|Date)[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})", text)
        or _find(r"(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})", text)
    )

    due_date = (
        _find(r"(?:Due\s+Date|Delivery\s+Date|Required\s+By|Delivery\s+By)[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})", text)
        or _find(r"Due[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})", text)
    )

    items = _extract_table_items_po(pdf_path, text)

    return {
        "po_number": po_number,
        "po_date": _normalise_date(po_date),
        "due_date": _normalise_date(due_date),
        "items": items,
        "raw_text": text[:3000],
    }


def extract_invoice_data(pdf_path):
    """Extract Invoice fields from PDF text."""
    text = extract_text(pdf_path)

    invoice_number = (
        _find(r"Invoice\s*(?:No|Number|#)[:\s]*([A-Z0-9\-/]+)", text)
        or _find(r"Inv\.?\s*(?:No|#)[:\s]*([A-Z0-9\-/]+)", text)
    )

    invoice_date = (
        _find(r"(?:Invoice\s+Date|Date)[:\s]*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})", text)
        or _find(r"(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})", text)
    )

    po_number = (
        _find(r"(?:Against\s+PO|PO\s*(?:No|Number|#)|Purchase\s+Order)[:\s]*([A-Z0-9\-/]+)", text)
        or _find(r"P\.?O\.?\s*[:\s#]*([A-Z0-9\-/]+)", text)
    )

    items = _extract_table_items_invoice(pdf_path, text)

    return {
        "invoice_number": invoice_number,
        "invoice_date": _normalise_date(invoice_date),
        "po_number": po_number,
        "items": items,
        "raw_text": text[:3000],
    }


def _extract_table_items_po(pdf_path, text):
    items = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table:
                        continue
                    headers = [str(c).lower().strip() if c else "" for c in table[0]]
                    code_col = _find_col(headers, ["code", "item code", "product code", "part no", "sku"])
                    name_col = _find_col(headers, ["description", "product", "name", "item", "particulars"])
                    qty_col = _find_col(headers, ["qty", "quantity", "ordered", "order qty"])

                    if qty_col is None:
                        continue

                    for row in table[1:]:
                        if not row or all(c is None or str(c).strip() == "" for c in row):
                            continue
                        code = str(row[code_col]).strip() if code_col is not None and row[code_col] else ""
                        name = str(row[name_col]).strip() if name_col is not None and row[name_col] else ""
                        qty_raw = str(row[qty_col]).strip() if row[qty_col] else "0"
                        qty = _parse_quantity(qty_raw)
                        if qty > 0 or code:
                            items.append({"product_code": code, "product_name": name, "ordered_qty": qty})
    except Exception:
        pass

    if not items:
        # Fallback: regex scan lines for qty patterns
        for line in text.splitlines():
            m = re.match(r"([A-Z0-9\-]+)\s+(.+?)\s+(\d+(?:\.\d+)?)\s*$", line.strip())
            if m:
                items.append({
                    "product_code": m.group(1),
                    "product_name": m.group(2).strip(),
                    "ordered_qty": float(m.group(3)),
                })

    return items


def _extract_table_items_invoice(pdf_path, text):
    items = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table:
                        continue
                    headers = [str(c).lower().strip() if c else "" for c in table[0]]
                    code_col = _find_col(headers, ["code", "item code", "product code", "part no", "sku"])
                    name_col = _find_col(headers, ["description", "product", "name", "item", "particulars"])
                    qty_col = _find_col(headers, ["qty", "quantity", "dispatched", "shipped"])

                    if qty_col is None:
                        continue

                    for row in table[1:]:
                        if not row or all(c is None or str(c).strip() == "" for c in row):
                            continue
                        code = str(row[code_col]).strip() if code_col is not None and row[code_col] else ""
                        name = str(row[name_col]).strip() if name_col is not None and row[name_col] else ""
                        qty_raw = str(row[qty_col]).strip() if row[qty_col] else "0"
                        qty = _parse_quantity(qty_raw)
                        if qty > 0 or code:
                            items.append({"product_code": code, "product_name": name, "dispatched_qty": qty})
    except Exception:
        pass

    if not items:
        for line in text.splitlines():
            m = re.match(r"([A-Z0-9\-]+)\s+(.+?)\s+(\d+(?:\.\d+)?)\s*$", line.strip())
            if m:
                items.append({
                    "product_code": m.group(1),
                    "product_name": m.group(2).strip(),
                    "dispatched_qty": float(m.group(3)),
                })

    return items


def _find_col(headers, keywords):
    for kw in keywords:
        for i, h in enumerate(headers):
            if kw in h:
                return i
    return None


def _normalise_date(d):
    """Convert DD/MM/YYYY or MM/DD/YYYY to YYYY-MM-DD best-effort."""
    if not d:
        return ""
    d = d.replace("/", "-")
    parts = d.split("-")
    if len(parts) == 3:
        if len(parts[2]) == 4:
            # DD-MM-YYYY
            try:
                return f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}"
            except Exception:
                return d
        elif len(parts[0]) == 4:
            # Already YYYY-MM-DD
            return d
    return d
