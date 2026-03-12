import sqlite3
import os
import sys


def get_base_path():
    # If running as EXE
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    # If running normally (python app.py)
    return os.path.dirname(os.path.abspath(__file__))


DB_NAME = os.path.join(get_base_path(), "accounts.db")


def get_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    cur = conn.cursor()

    # Customers — fixed missing commas between columns
    cur.execute("""
    CREATE TABLE IF NOT EXISTS customers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT,
        address TEXT,
        gst_no TEXT
    )
    """)

    # Sellers — fixed missing commas between columns
    cur.execute("""
    CREATE TABLE IF NOT EXISTS sellers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT,
        address TEXT,
        gst_no TEXT
    )
    """)

    # Sales
    cur.execute("""
    CREATE TABLE IF NOT EXISTS sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_date TEXT NOT NULL,
        customer_id INTEGER NOT NULL,
        bill_no TEXT NOT NULL,
        amount REAL NOT NULL DEFAULT 0,
        paid_amount REAL NOT NULL DEFAULT 0,
        notes TEXT,
        FOREIGN KEY(customer_id) REFERENCES customers(id)
    )
    """)

    # Purchases
    cur.execute("""
    CREATE TABLE IF NOT EXISTS purchases (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        purchase_date TEXT NOT NULL,
        seller_id INTEGER NOT NULL,
        payment_method TEXT,
        payment_ref TEXT,
        bill_no TEXT,
        product_description TEXT,
        amount REAL NOT NULL DEFAULT 0,
        notes TEXT,
        FOREIGN KEY(seller_id) REFERENCES sellers(id)
    )
    """)

    # Payment history log
    cur.execute("""
    CREATE TABLE IF NOT EXISTS payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_id INTEGER NOT NULL,
        payment_date TEXT NOT NULL,
        amount REAL NOT NULL,
        notes TEXT,
        FOREIGN KEY(sale_id) REFERENCES sales(id)
    )
    """)

    # --- Migration: add missing columns safely ---

    # customers table
    try:
        cur.execute("ALTER TABLE customers ADD COLUMN address TEXT")
    except Exception:
        pass

    try:
        cur.execute("ALTER TABLE customers ADD COLUMN gst_no TEXT")
    except Exception:
        pass

    # sellers table
    try:
        cur.execute("ALTER TABLE sellers ADD COLUMN address TEXT")
    except Exception:
        pass

    try:
        cur.execute("ALTER TABLE sellers ADD COLUMN gst_no TEXT")
    except Exception:
        pass

    conn.commit()
    conn.close()
