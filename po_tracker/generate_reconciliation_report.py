"""
Item-wise PO vs Invoice reconciliation report.
Re-parses PO and Invoice PDFs for unit price/value, matches against the
existing PURCHASE_ORDER_ITEMS / INVOICE_ITEMS records by ProductCode.
Run: python3 generate_reconciliation_report.py
Output: data/reports/PO_Invoice_Reconciliation.xlsx
"""
import os
import re
import sqlite3
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(__file__)
DB_PATH = os.path.join(BASE, "data", "po_tracker.db")
OUT_DIR = os.path.join(BASE, "data", "reports")
OUT_PATH = os.path.join(OUT_DIR, "PO_Invoice_Reconciliation.xlsx")

HEADER_FILL = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(bold=True, size=14)
THIN = Side(style="thin", color="D1D5DB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

STATUS_FILL = {
    "Matched": PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"),
    "Short Supplied": PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid"),
    "Excess Supplied": PatternFill(start_color="BFDBFE", end_color="BFDBFE", fill_type="solid"),
    "Not Invoiced": PatternFill(start_color="E5E7EB", end_color="E5E7EB", fill_type="solid"),
    "Extra Item": PatternFill(start_color="FECACA", end_color="FECACA", fill_type="solid"),
}

LINE_RE = re.compile(
    r"^(\d+)\s*(.+?)\s+(\d{4,8})\s+([\d,]+(?:\.\d+)?)\s+(NOS|MTS|MT)\s+([\d,]+(?:\.\d+)?)\s+(?:NOS|MTS|MT)\s+([\d,]+(?:\.\d+)?)\s*$",
    re.IGNORECASE,
)


def _num(s):
    return float(s.replace(",", ""))


def product_code(desc):
    d = desc.upper().strip()
    m_mm = re.findall(r"(\d+)MM", d)
    dims = re.findall(r"X\s*/?\s*(\d+)", d)
    thickness = m_mm[-1] if m_mm else (dims[-1] if dims else "")
    if "MORTAR" in d and "MGT" in d:
        return "MGT-MORTAR", "MGT Mortar"
    if "MORTAR" in d and "MCH" in d:
        return "MCH-MORTAR", "MCH Mortar"
    if "MORTAR" in d:
        return "AL70-MORTAR", "70% AL Mortar"
    if "MCH" in d:
        if thickness == "30":
            return "MCH-30", "MCH 230x115x30 Fire Bricks"
        if thickness == "40":
            return "MCH-40", "MCH 230x115x40 Fire Bricks"
        if thickness in ("64", "65"):
            return "MCH-65", "MCH 230x115x65 Fire Bricks"
    if "MGT" in d:
        if thickness in ("65", "110"):
            return "MGT-65", "MGT 230x115x65 Fire Bricks"
        if thickness in ("75", "76"):
            return "MGT-76", "MGT 230x115x76 Fire Bricks"
    if "70%" in d or ("AL" in d and "70" in d and "ALUM" not in d):
        if thickness in ("64", "65"):
            return "AL70-65", "AL70% 230x115x65 Fire Bricks"
        if thickness in ("75", "76"):
            return "AL70-76", "AL70% 230x115x76 Fire Bricks"
        return "AL70-32", "AL70% 230x115x32 Fire Bricks"
    if "60%" in d:
        if thickness not in ("32", "33"):
            return "AL60-FB", "AL60% Fire Bricks"
        return "AL60-32", "AL60% 230x115x32 Fire Bricks"
    if "50%" in d:
        return "AL60-FB", "AL50% Fire Bricks"
    if thickness in ("64", "65"):
        return "AL70-65", "AL70% 230x115x65 Fire Bricks"
    return "UNKNOWN", desc[:40]


def parse_invoice_pdf(path):
    """Returns dict: invoice_number -> list of {desc, qty, rate, amount, product_code}"""
    invoices = {}
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if "Tax Invoice" not in text[:30]:
                continue
            m = re.search(r"\b(20\d\d-\d\d/\d+)\b", text)
            if not m:
                continue
            inv_no = m.group(1)
            lines = text.splitlines()
            items = []
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                m2 = LINE_RE.match(line)
                if m2:
                    sl, cat, hsn, qty, unit, rate, amount = m2.groups()
                    desc = cat
                    if i + 1 < len(lines):
                        nxt = lines[i + 1].strip()
                        if nxt and not re.match(r"^[\d,]+\.\d\d", nxt) and "Total" not in nxt and "IGST" not in nxt:
                            desc = nxt
                    code, name = product_code(desc)
                    items.append({
                        "desc": desc, "qty": _num(qty), "rate": _num(rate),
                        "amount": _num(amount), "product_code": code, "product_name": name,
                    })
                i += 1
            if items:
                invoices[inv_no] = items
    return invoices


# ---------- PO item data with unit price (re-keyed from PDF text already reviewed) ----------
PO_PRICED_ITEMS = {
    "AI/25-26/550": [
        ("MCH-30", "MCH 230x115x30 Fire Bricks", 6275, 55.00),
        ("MCH-40", "MCH 230x115x40 Fire Bricks", 3000, 70.00),
        ("MGT-76", "MGT 230x115x76 Fire Bricks", 2000, 170.00),
        ("MCH-65", "MCH 230x115x65 Fire Bricks", 6250, 115.00),
        ("AL60-32", "AL60% 230x115x32 Fire Bricks", 10000, 30.00),
        ("MGT-65", "MGT 230x115x65 Fire Bricks", 10000, 125.00),
    ],
    "AI/400 TON/25-26/641": [
        ("AL70-32", "AL70% 230x115x32 Fire Bricks", 10000, 34.00),
        ("AL60-32", "AL60% 230x115x32 Fire Bricks", 15000, 30.00),
        ("MGT-76", "MGT 230x115x76 Fire Bricks", 5000, 170.00),
        ("MGT-65", "MGT 230x115x65 Fire Bricks", 10000, 125.00),
    ],
    "AI/25-26/712": [
        ("AL70-MORTAR", "70% AL Mortar", 35, 16500.00),
        ("MGT-76", "MGT 230x115x76 Fire Bricks", 10000, 170.00),
        ("MGT-65", "MGT 230x115x65 Fire Bricks", 10000, 125.00),
    ],
    "AI/25-26/780": [
        ("AL60-32", "AL60% 230x115x32 Fire Bricks", 12500, 30.00),
        ("AL70-32", "AL70% 230x115x32 Fire Bricks", 12500, 34.00),
        ("MGT-65", "MGT 230x115x65 Fire Bricks", 25000, 125.00),
    ],
    "AI/25-26/817": [
        ("MCH-30", "MCH 230x115x30 Fire Bricks", 15000, 55.00),
        ("MCH-65", "MCH 230x115x65 Fire Bricks", 15000, 115.00),
    ],
    "AI/25-26/1031": [
        ("MCH-30", "MCH 230x115x30 Fire Bricks", 45000, 55.00),
    ],
    "AI/25-26/1175": [
        ("MGT-76", "MGT 230x115x76 Fire Bricks", 35000, 165.00),
        ("AL70-32", "AL70% 230x115x32 Fire Bricks", 20000, 32.00),
        ("MGT-65", "MGT 230x115x65 Fire Bricks", 53000, 122.00),
        ("MCH-40", "MCH 230x115x40 Fire Bricks", 4600, 70.00),
    ],
    "AI/26-27/110": [
        ("AL60-FB", "AL50% 230x114x75/65 Fire Bricks", 5000, 53.00),
        ("MGT-76", "MGT 230x115x76 Fire Bricks", 2000, 180.00),
    ],
    "MI/26-27/149": [
        ("AL70-76", "AL70% 230x114x76 Fire Bricks", 2512, 110.00),
        ("MCH-65", "MCH 230x114x64 Fire Bricks", 1312, 132.00),
        ("AL70-MORTAR", "70% AL Mortar", 0.6, 19200.00),
        ("MCH-MORTAR", "MCH Mortar", 0.6, 24900.00),
    ],
    "AI/26-27/190": [
        ("MGT-76", "MGT 230x115x76 Fire Bricks", 4800, 183.00),
        ("AL70-76", "AL70% 230x115x76 Fire Bricks", 1330, 110.00),
        ("AL70-MORTAR", "70% AL Mortar", 1.5, 19200.00),
    ],
    "AI/26-27/191": [
        ("AL70-65", "AL70% 230x115x65 Fire Bricks", 1175, 85.00),
        ("AL70-32", "AL70% 230x115x35 Fire Bricks", 4000, 34.00),
        ("MGT-65", "MGT 230x115x65 Fire Bricks", 3250, 120.00),
        ("MGT-MORTAR", "MGT Mortar", 1.25, 21000.00),
        ("AL70-MORTAR", "70% AL Mortar", 1.25, 16200.00),
    ],
    "AI/BRICKS/26-27/215": [
        ("AL70-32", "AL70% 230x115x32 Fire Bricks", 20000, 34.00),
    ],
    "AI/26-27/227": [
        ("MCH-30", "MCH 230x114x30 Fire Bricks", 20000, 55.00),
    ],
    "SC/26-27/228": [
        ("MGT-76", "MGT 230x115x76 Fire Bricks", 6000, 180.00),
    ],
}


def _style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def _autosize(ws):
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(length + 3, 12), 42)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # Parse all invoice PDFs for priced line items, keyed by invoice number
    invoice_dir = os.path.join(BASE, "data", "invoice_pdfs")
    priced_invoices = {}
    for fname in os.listdir(invoice_dir):
        if fname.lower().endswith(".pdf"):
            priced_invoices.update(parse_invoice_pdf(os.path.join(invoice_dir, fname)))

    # Aggregate invoice items by (PONumber, ProductCode) using DB for invoice->PO mapping
    invoices_db = cur.execute("SELECT InvoiceID, InvoiceNumber, PONumber FROM INVOICES").fetchall()
    po_invoice_agg = {}  # (PONumber, ProductCode) -> {qty, amount}
    unmatched_extra = {}  # (PONumber, ProductCode) -> {qty, amount, name} for items not in PO

    po_numbers_set = set(PO_PRICED_ITEMS.keys())

    for inv in invoices_db:
        inv_no = inv["InvoiceNumber"]
        po_no = inv["PONumber"]
        items = priced_invoices.get(inv_no, [])
        for it in items:
            key = (po_no, it["product_code"])
            agg = po_invoice_agg.setdefault(key, {"qty": 0.0, "amount": 0.0, "name": it["product_name"]})
            agg["qty"] += it["qty"]
            agg["amount"] += it["amount"]

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Reconciliation"
    ws1["A1"] = "Item-wise PO vs Invoice Reconciliation"
    ws1["A1"].font = TITLE_FONT
    ws1.append([])
    headers = ["PO Number", "Item Description", "PO Qty Ordered", "Invoice Qty",
               "Qty Difference", "PO Unit Price", "Invoice Unit Price",
               "PO Value", "Invoice Value", "Status"]
    ws1.append(headers)
    _style_header(ws1, 3, len(headers))

    total_po_items = 0
    total_inv_items = 0
    counts = {"Matched": 0, "Short Supplied": 0, "Excess Supplied": 0, "Not Invoiced": 0, "Extra Item": 0}

    seen_keys = set()

    for po_no, items in PO_PRICED_ITEMS.items():
        for code, name, ordered_qty, unit_price in items:
            total_po_items += 1
            key = (po_no, code)
            seen_keys.add(key)
            agg = po_invoice_agg.get(key, {"qty": 0.0, "amount": 0.0})
            inv_qty = agg["qty"]
            inv_amount = agg["amount"]
            inv_unit_price = (inv_amount / inv_qty) if inv_qty > 0 else 0
            po_value = ordered_qty * unit_price
            qty_diff = inv_qty - ordered_qty

            if inv_qty == 0:
                status = "Not Invoiced"
            elif abs(qty_diff) < 0.01:
                status = "Matched"
            elif inv_qty < ordered_qty:
                status = "Short Supplied"
            else:
                status = "Excess Supplied"

            counts[status] += 1
            ws1.append([po_no, name, ordered_qty, inv_qty, qty_diff, unit_price,
                        round(inv_unit_price, 2), po_value, round(inv_amount, 2), status])
            r = ws1.max_row
            for c in range(1, len(headers) + 1):
                ws1.cell(row=r, column=c).border = BORDER
            ws1.cell(row=r, column=10).fill = STATUS_FILL[status]

    # Extra items: invoiced under a PO/product not present in that PO's item list
    for (po_no, code), agg in po_invoice_agg.items():
        if po_no not in po_numbers_set:
            continue
        if (po_no, code) in seen_keys:
            continue
        total_inv_items += 1
        counts["Extra Item"] += 1
        ws1.append([po_no, agg["name"], 0, agg["qty"], agg["qty"], 0, round(agg["amount"] / agg["qty"], 2) if agg["qty"] else 0,
                    0, round(agg["amount"], 2), "Extra Item"])
        r = ws1.max_row
        for c in range(1, len(headers) + 1):
            ws1.cell(row=r, column=c).border = BORDER
        ws1.cell(row=r, column=10).fill = STATUS_FILL["Extra Item"]

    total_inv_items += sum(1 for k in seen_keys if po_invoice_agg.get(k, {"qty": 0})["qty"] > 0)

    _autosize(ws1)
    ws1.freeze_panes = "A4"

    # ---------- Summary sheet ----------
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Reconciliation Summary"
    ws2["A1"].font = TITLE_FONT
    ws2.append([])
    rows = [
        ("Total PO Items", total_po_items),
        ("Total Invoice Items (matched to PO lines)", total_inv_items),
        ("Total Matched Items", counts["Matched"]),
        ("Total Short-Supplied Items", counts["Short Supplied"]),
        ("Total Excess-Supplied Items", counts["Excess Supplied"]),
        ("Total Not-Invoiced (Missing) Items", counts["Not Invoiced"]),
        ("Total Extra Items (not in PO)", counts["Extra Item"]),
    ]
    for label, val in rows:
        ws2.append([label, val])
        r = ws2.max_row
        ws2.cell(row=r, column=1).font = Font(bold=True)
        ws2.cell(row=r, column=1).border = BORDER
        ws2.cell(row=r, column=2).border = BORDER

    ws2.append([])
    ws2.append(["Action Required — Discrepancies Needing Review"])
    ws2.cell(row=ws2.max_row, column=1).font = Font(bold=True, size=12)
    for po_no, items in PO_PRICED_ITEMS.items():
        for code, name, ordered_qty, unit_price in items:
            key = (po_no, code)
            agg = po_invoice_agg.get(key, {"qty": 0.0, "amount": 0.0})
            inv_qty = agg["qty"]
            if inv_qty == 0:
                ws2.append([f"{po_no} — {name}: NOT INVOICED YET ({ordered_qty} pending)"])
            elif inv_qty < ordered_qty - 0.01:
                ws2.append([f"{po_no} — {name}: SHORT by {ordered_qty - inv_qty}"])
            elif inv_qty > ordered_qty + 0.01:
                ws2.append([f"{po_no} — {name}: EXCESS by {inv_qty - ordered_qty}"])

    _autosize(ws2)

    wb.save(OUT_PATH)
    conn.close()
    print(f"Saved: {OUT_PATH}")
    print(f"PO items: {total_po_items} | Matched: {counts['Matched']} | Short: {counts['Short Supplied']} | "
          f"Excess: {counts['Excess Supplied']} | Not Invoiced: {counts['Not Invoiced']} | Extra: {counts['Extra Item']}")


if __name__ == "__main__":
    main()
