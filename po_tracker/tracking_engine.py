"""
Core tracking engine — single source of truth for all quantity calculations.
All status and quantity logic lives here. Never calculated elsewhere.
"""
from database import db
from datetime import date


def _determine_status(ordered_qty, dispatched_qty, due_date_str):
    if ordered_qty <= 0:
        return "NOT STARTED"

    pct = dispatched_qty / ordered_qty * 100

    if pct >= 100:
        return "COMPLETED"

    if due_date_str:
        try:
            due = date.fromisoformat(due_date_str)
            if due < date.today() and dispatched_qty < ordered_qty:
                return "OVERDUE"
        except ValueError:
            pass

    if pct == 0:
        return "NOT STARTED"

    return "IN PROGRESS"


def get_po_tracking(po_id):
    """Return full tracking data for a PO including per-product breakdown."""
    with db() as conn:
        po = conn.execute(
            "SELECT * FROM PURCHASE_ORDERS WHERE POID = ?", (po_id,)
        ).fetchone()
        if not po:
            return None

        items = conn.execute(
            "SELECT * FROM PURCHASE_ORDER_ITEMS WHERE POID = ?", (po_id,)
        ).fetchall()

        result_items = []
        total_ordered = 0
        total_dispatched = 0

        for item in items:
            dispatched = conn.execute(
                """
                SELECT COALESCE(SUM(ii.DispatchedQty), 0) as total
                FROM INVOICE_ITEMS ii
                JOIN INVOICES inv ON ii.InvoiceID = inv.InvoiceID
                WHERE ii.ProductCode = ? AND inv.PONumber = ?
                """,
                (item["ProductCode"], po["PONumber"]),
            ).fetchone()["total"]

            pending = max(0, item["OrderedQty"] - dispatched)
            status = _determine_status(item["OrderedQty"], dispatched, po["DueDate"])

            result_items.append({
                "POItemID": item["POItemID"],
                "ProductCode": item["ProductCode"],
                "ProductName": item["ProductName"],
                "OrderedQty": item["OrderedQty"],
                "DispatchedQty": dispatched,
                "PendingQty": pending,
                "Status": status,
            })

            total_ordered += item["OrderedQty"]
            total_dispatched += dispatched

        overall_status = _determine_status(total_ordered, total_dispatched, po["DueDate"])

        return {
            "POID": po["POID"],
            "PONumber": po["PONumber"],
            "PODate": po["PODate"],
            "DueDate": po["DueDate"],
            "PDFPath": po["PDFPath"],
            "CreatedAt": po["CreatedAt"],
            "TotalOrdered": total_ordered,
            "TotalDispatched": total_dispatched,
            "TotalPending": max(0, total_ordered - total_dispatched),
            "Status": overall_status,
            "Items": result_items,
        }


def recalculate_po_status(po_id, conn):
    """Recalculate and persist PO status — call inside an open db() context."""
    po = conn.execute(
        "SELECT PONumber, DueDate FROM PURCHASE_ORDERS WHERE POID = ?", (po_id,)
    ).fetchone()
    if not po:
        return

    items = conn.execute(
        "SELECT OrderedQty, ProductCode FROM PURCHASE_ORDER_ITEMS WHERE POID = ?",
        (po_id,),
    ).fetchall()

    total_ordered = 0
    total_dispatched = 0

    for item in items:
        dispatched = conn.execute(
            """
            SELECT COALESCE(SUM(ii.DispatchedQty), 0) as total
            FROM INVOICE_ITEMS ii
            JOIN INVOICES inv ON ii.InvoiceID = inv.InvoiceID
            WHERE ii.ProductCode = ? AND inv.PONumber = ?
            """,
            (item["ProductCode"], po["PONumber"]),
        ).fetchone()["total"]
        total_ordered += item["OrderedQty"]
        total_dispatched += dispatched

    status = _determine_status(total_ordered, total_dispatched, po["DueDate"])
    conn.execute(
        "UPDATE PURCHASE_ORDERS SET Status = ? WHERE POID = ?", (status, po_id)
    )


def get_all_pos_tracking():
    """Return summary tracking for all POs."""
    with db() as conn:
        pos = conn.execute(
            "SELECT POID FROM PURCHASE_ORDERS ORDER BY CreatedAt DESC"
        ).fetchall()

    return [get_po_tracking(po["POID"]) for po in pos]


def get_dashboard_counts():
    with db() as conn:
        today = date.today().isoformat()

        open_count = conn.execute(
            "SELECT COUNT(*) FROM PURCHASE_ORDERS WHERE Status NOT IN ('COMPLETED')"
        ).fetchone()[0]

        completed_count = conn.execute(
            "SELECT COUNT(*) FROM PURCHASE_ORDERS WHERE Status = 'COMPLETED'"
        ).fetchone()[0]

        overdue_count = conn.execute(
            "SELECT COUNT(*) FROM PURCHASE_ORDERS WHERE Status = 'OVERDUE'"
        ).fetchone()[0]

        week_end = _week_end()
        due_this_week = conn.execute(
            """
            SELECT COUNT(*) FROM PURCHASE_ORDERS
            WHERE Status NOT IN ('COMPLETED')
              AND DueDate >= ? AND DueDate <= ?
            """,
            (today, week_end),
        ).fetchone()[0]

        recent_completed = conn.execute(
            """
            SELECT PONumber, DueDate, CreatedAt FROM PURCHASE_ORDERS
            WHERE Status = 'COMPLETED'
            ORDER BY CreatedAt DESC LIMIT 5
            """
        ).fetchall()

        return {
            "open": open_count,
            "completed": completed_count,
            "overdue": overdue_count,
            "due_this_week": due_this_week,
            "recent_completed": [dict(r) for r in recent_completed],
        }


def get_pending_dispatch_report():
    """All pending products across open POs."""
    rows = []
    with db() as conn:
        pos = conn.execute(
            "SELECT POID, PONumber, DueDate FROM PURCHASE_ORDERS WHERE Status != 'COMPLETED'"
        ).fetchall()

        for po in pos:
            items = conn.execute(
                "SELECT * FROM PURCHASE_ORDER_ITEMS WHERE POID = ?", (po["POID"],)
            ).fetchall()

            for item in items:
                dispatched = conn.execute(
                    """
                    SELECT COALESCE(SUM(ii.DispatchedQty), 0) as total
                    FROM INVOICE_ITEMS ii
                    JOIN INVOICES inv ON ii.InvoiceID = inv.InvoiceID
                    WHERE ii.ProductCode = ? AND inv.PONumber = ?
                    """,
                    (item["ProductCode"], po["PONumber"]),
                ).fetchone()["total"]

                pending = max(0, item["OrderedQty"] - dispatched)
                if pending > 0:
                    status = _determine_status(
                        item["OrderedQty"], dispatched, po["DueDate"]
                    )
                    rows.append({
                        "PONumber": po["PONumber"],
                        "DueDate": po["DueDate"],
                        "ProductCode": item["ProductCode"],
                        "ProductName": item["ProductName"],
                        "OrderedQty": item["OrderedQty"],
                        "DispatchedQty": dispatched,
                        "PendingQty": pending,
                        "Status": status,
                    })

    return rows


def get_invoice_dispatch_history(po_number):
    """All invoices and items linked to a PO number."""
    with db() as conn:
        invoices = conn.execute(
            "SELECT * FROM INVOICES WHERE PONumber = ? ORDER BY InvoiceDate",
            (po_number,),
        ).fetchall()

        result = []
        for inv in invoices:
            items = conn.execute(
                "SELECT * FROM INVOICE_ITEMS WHERE InvoiceID = ?", (inv["InvoiceID"],)
            ).fetchall()
            result.append({
                "InvoiceID": inv["InvoiceID"],
                "InvoiceNumber": inv["InvoiceNumber"],
                "InvoiceDate": inv["InvoiceDate"],
                "PDFPath": inv["PDFPath"],
                "Items": [dict(i) for i in items],
            })

        return result


def _week_end():
    from datetime import timedelta
    today = date.today()
    days_until_sunday = 6 - today.weekday()
    return (today + timedelta(days=days_until_sunday)).isoformat()
