import sqlite3
import os
from contextlib import contextmanager

DB_PATH = os.path.join(os.path.dirname(__file__), "data", "po_tracker.db")


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@contextmanager
def db():
    conn = get_connection()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    with db() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS PURCHASE_ORDERS (
                POID        INTEGER PRIMARY KEY AUTOINCREMENT,
                PONumber    TEXT NOT NULL UNIQUE,
                PODate      TEXT,
                DueDate     TEXT,
                Status      TEXT NOT NULL DEFAULT 'NOT STARTED',
                PDFPath     TEXT,
                CreatedAt   TEXT NOT NULL DEFAULT (datetime('now'))
            );

            CREATE TABLE IF NOT EXISTS PURCHASE_ORDER_ITEMS (
                POItemID    INTEGER PRIMARY KEY AUTOINCREMENT,
                POID        INTEGER NOT NULL,
                ProductCode TEXT NOT NULL,
                ProductName TEXT,
                OrderedQty  REAL NOT NULL DEFAULT 0,
                FOREIGN KEY (POID) REFERENCES PURCHASE_ORDERS(POID)
            );

            CREATE TABLE IF NOT EXISTS INVOICES (
                InvoiceID     INTEGER PRIMARY KEY AUTOINCREMENT,
                InvoiceNumber TEXT NOT NULL,
                InvoiceDate   TEXT,
                PONumber      TEXT,
                PDFPath       TEXT,
                CreatedAt     TEXT NOT NULL DEFAULT (datetime('now'))
            );

            CREATE TABLE IF NOT EXISTS INVOICE_ITEMS (
                InvoiceItemID  INTEGER PRIMARY KEY AUTOINCREMENT,
                InvoiceID      INTEGER NOT NULL,
                ProductCode    TEXT NOT NULL,
                ProductName    TEXT,
                DispatchedQty  REAL NOT NULL DEFAULT 0,
                FOREIGN KEY (InvoiceID) REFERENCES INVOICES(InvoiceID)
            );

            CREATE TABLE IF NOT EXISTS PRODUCT_MASTER (
                ProductCode TEXT PRIMARY KEY,
                ProductName TEXT
            );
        """)
