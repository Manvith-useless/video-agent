import os
import json
import shutil
from flask import Flask, request, jsonify, render_template, redirect, url_for, send_file
from werkzeug.utils import secure_filename
from database import init_db, db
from tracking_engine import (
    get_po_tracking,
    get_all_pos_tracking,
    get_dashboard_counts,
    get_pending_dispatch_report,
    get_invoice_dispatch_history,
    recalculate_po_status,
)
from pdf_extractor import extract_po_data, extract_invoice_data

BASE_DIR = os.path.dirname(__file__)
PO_PDF_DIR = os.path.join(BASE_DIR, "data", "po_pdfs")
INV_PDF_DIR = os.path.join(BASE_DIR, "data", "invoice_pdfs")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


# ── Startup ──────────────────────────────────────────────────────────────────

init_db()


# ── Helpers ──────────────────────────────────────────────────────────────────

def _save_pdf(file, directory):
    filename = secure_filename(file.filename)
    dest = os.path.join(directory, filename)
    # Avoid collisions
    base, ext = os.path.splitext(filename)
    counter = 1
    while os.path.exists(dest):
        dest = os.path.join(directory, f"{base}_{counter}{ext}")
        counter += 1
    file.save(dest)
    return dest


def _upsert_product(conn, code, name):
    if not code:
        return
    existing = conn.execute(
        "SELECT ProductCode FROM PRODUCT_MASTER WHERE ProductCode = ?", (code,)
    ).fetchone()
    if existing:
        if name:
            conn.execute(
                "UPDATE PRODUCT_MASTER SET ProductName = ? WHERE ProductCode = ?",
                (name, code),
            )
    else:
        conn.execute(
            "INSERT INTO PRODUCT_MASTER (ProductCode, ProductName) VALUES (?, ?)",
            (code, name),
        )


# ── Dashboard ────────────────────────────────────────────────────────────────

@app.route("/")
def dashboard():
    counts = get_dashboard_counts()
    return render_template("dashboard.html", counts=counts)


# ── PO Management ────────────────────────────────────────────────────────────

@app.route("/po/upload", methods=["GET"])
def po_upload_page():
    return render_template("po_upload.html")


@app.route("/po/extract", methods=["POST"])
def po_extract():
    if "pdf" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["pdf"]
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400

    path = _save_pdf(file, PO_PDF_DIR)
    data = extract_po_data(path)
    data["pdf_path"] = path
    return jsonify(data)


@app.route("/po/save", methods=["POST"])
def po_save():
    data = request.get_json()
    po_number = data.get("po_number", "").strip()
    po_date = data.get("po_date", "").strip()
    due_date = data.get("due_date", "").strip()
    pdf_path = data.get("pdf_path", "").strip()
    items = data.get("items", [])

    if not po_number:
        return jsonify({"error": "PO Number is required"}), 400
    if not items:
        return jsonify({"error": "At least one product item is required"}), 400

    with db() as conn:
        existing = conn.execute(
            "SELECT POID FROM PURCHASE_ORDERS WHERE PONumber = ?", (po_number,)
        ).fetchone()
        if existing:
            return jsonify({"error": f"PO Number '{po_number}' already exists"}), 409

        cur = conn.execute(
            "INSERT INTO PURCHASE_ORDERS (PONumber, PODate, DueDate, Status, PDFPath) VALUES (?,?,?,?,?)",
            (po_number, po_date, due_date, "NOT STARTED", pdf_path),
        )
        po_id = cur.lastrowid

        for item in items:
            code = str(item.get("product_code", "")).strip()
            name = str(item.get("product_name", "")).strip()
            qty = float(item.get("ordered_qty", 0) or 0)
            conn.execute(
                "INSERT INTO PURCHASE_ORDER_ITEMS (POID, ProductCode, ProductName, OrderedQty) VALUES (?,?,?,?)",
                (po_id, code, name, qty),
            )
            _upsert_product(conn, code, name)

        recalculate_po_status(po_id, conn)

    return jsonify({"success": True, "po_id": po_id})


@app.route("/po/list")
def po_list():
    tracking = get_all_pos_tracking()
    return render_template("po_list.html", pos=tracking)


@app.route("/po/<int:po_id>")
def po_detail(po_id):
    tracking = get_po_tracking(po_id)
    if not tracking:
        return "PO not found", 404
    history = get_invoice_dispatch_history(tracking["PONumber"])
    return render_template("po_detail.html", po=tracking, history=history)


@app.route("/po/<int:po_id>/edit", methods=["GET"])
def po_edit_page(po_id):
    with db() as conn:
        po = conn.execute("SELECT * FROM PURCHASE_ORDERS WHERE POID = ?", (po_id,)).fetchone()
        items = conn.execute("SELECT * FROM PURCHASE_ORDER_ITEMS WHERE POID = ?", (po_id,)).fetchall()
    if not po:
        return "PO not found", 404
    return render_template("po_edit.html", po=dict(po), items=[dict(i) for i in items])


@app.route("/po/<int:po_id>/update", methods=["POST"])
def po_update(po_id):
    data = request.get_json()
    po_date = data.get("po_date", "").strip()
    due_date = data.get("due_date", "").strip()
    items = data.get("items", [])

    with db() as conn:
        conn.execute(
            "UPDATE PURCHASE_ORDERS SET PODate = ?, DueDate = ? WHERE POID = ?",
            (po_date, due_date, po_id),
        )
        conn.execute("DELETE FROM PURCHASE_ORDER_ITEMS WHERE POID = ?", (po_id,))
        for item in items:
            code = str(item.get("product_code", "")).strip()
            name = str(item.get("product_name", "")).strip()
            qty = float(item.get("ordered_qty", 0) or 0)
            conn.execute(
                "INSERT INTO PURCHASE_ORDER_ITEMS (POID, ProductCode, ProductName, OrderedQty) VALUES (?,?,?,?)",
                (po_id, code, name, qty),
            )
            _upsert_product(conn, code, name)
        recalculate_po_status(po_id, conn)

    return jsonify({"success": True})


# ── Invoice Management ────────────────────────────────────────────────────────

@app.route("/invoice/upload", methods=["GET"])
def invoice_upload_page():
    with db() as conn:
        pos = conn.execute(
            "SELECT PONumber FROM PURCHASE_ORDERS WHERE Status != 'COMPLETED' ORDER BY PONumber"
        ).fetchall()
    return render_template("invoice_upload.html", open_pos=[r["PONumber"] for r in pos])


@app.route("/invoice/extract", methods=["POST"])
def invoice_extract():
    if "pdf" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["pdf"]
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400

    path = _save_pdf(file, INV_PDF_DIR)
    data = extract_invoice_data(path)
    data["pdf_path"] = path
    return jsonify(data)


@app.route("/invoice/save", methods=["POST"])
def invoice_save():
    data = request.get_json()
    invoice_number = data.get("invoice_number", "").strip()
    invoice_date = data.get("invoice_date", "").strip()
    po_number = data.get("po_number", "").strip()
    pdf_path = data.get("pdf_path", "").strip()
    items = data.get("items", [])

    if not invoice_number:
        return jsonify({"error": "Invoice Number is required"}), 400
    if not po_number:
        return jsonify({"error": "PO Number is required"}), 400
    if not items:
        return jsonify({"error": "At least one dispatch item is required"}), 400

    with db() as conn:
        po = conn.execute(
            "SELECT POID FROM PURCHASE_ORDERS WHERE PONumber = ?", (po_number,)
        ).fetchone()
        if not po:
            return jsonify({"error": f"PO Number '{po_number}' not found"}), 404

        existing = conn.execute(
            "SELECT InvoiceID FROM INVOICES WHERE InvoiceNumber = ?", (invoice_number,)
        ).fetchone()
        if existing:
            return jsonify({"error": f"Invoice '{invoice_number}' already exists"}), 409

        cur = conn.execute(
            "INSERT INTO INVOICES (InvoiceNumber, InvoiceDate, PONumber, PDFPath) VALUES (?,?,?,?)",
            (invoice_number, invoice_date, po_number, pdf_path),
        )
        inv_id = cur.lastrowid

        for item in items:
            code = str(item.get("product_code", "")).strip()
            name = str(item.get("product_name", "")).strip()
            qty = float(item.get("dispatched_qty", 0) or 0)
            conn.execute(
                "INSERT INTO INVOICE_ITEMS (InvoiceID, ProductCode, ProductName, DispatchedQty) VALUES (?,?,?,?)",
                (inv_id, code, name, qty),
            )
            _upsert_product(conn, code, name)

        recalculate_po_status(po["POID"], conn)

    return jsonify({"success": True, "invoice_id": inv_id})


@app.route("/invoice/list")
def invoice_list():
    with db() as conn:
        invoices = conn.execute(
            "SELECT * FROM INVOICES ORDER BY CreatedAt DESC"
        ).fetchall()
    return render_template("invoice_list.html", invoices=[dict(i) for i in invoices])


# ── Reports ───────────────────────────────────────────────────────────────────

@app.route("/reports/pending")
def report_pending():
    rows = get_pending_dispatch_report()
    return render_template("report_pending.html", rows=rows)


@app.route("/reports/completed")
def report_completed():
    with db() as conn:
        pos = conn.execute(
            "SELECT POID FROM PURCHASE_ORDERS WHERE Status = 'COMPLETED' ORDER BY CreatedAt DESC"
        ).fetchall()
    data = [get_po_tracking(p["POID"]) for p in pos]
    return render_template("report_completed.html", pos=data)


@app.route("/reports/overdue")
def report_overdue():
    with db() as conn:
        pos = conn.execute(
            "SELECT POID FROM PURCHASE_ORDERS WHERE Status = 'OVERDUE' ORDER BY DueDate"
        ).fetchall()
    data = [get_po_tracking(p["POID"]) for p in pos]
    return render_template("report_overdue.html", pos=data)


# ── API Endpoints (JSON) ──────────────────────────────────────────────────────

@app.route("/api/po/<int:po_id>")
def api_po(po_id):
    data = get_po_tracking(po_id)
    if not data:
        return jsonify({"error": "Not found"}), 404
    return jsonify(data)


@app.route("/api/po/<int:po_id>/products")
def api_po_products(po_id):
    """Return product codes for a PO — used by invoice form autocomplete."""
    with db() as conn:
        items = conn.execute(
            "SELECT ProductCode, ProductName FROM PURCHASE_ORDER_ITEMS WHERE POID = ?",
            (po_id,),
        ).fetchall()
    return jsonify([dict(i) for i in items])


@app.route("/api/po_by_number/<po_number>")
def api_po_by_number(po_number):
    with db() as conn:
        po = conn.execute(
            "SELECT POID FROM PURCHASE_ORDERS WHERE PONumber = ?", (po_number,)
        ).fetchone()
    if not po:
        return jsonify({"error": "Not found"}), 404
    return api_po(po["POID"])


@app.route("/api/products")
def api_products():
    with db() as conn:
        rows = conn.execute("SELECT * FROM PRODUCT_MASTER ORDER BY ProductCode").fetchall()
    return jsonify([dict(r) for r in rows])


if __name__ == "__main__":
    app.run(debug=True, port=5000)
