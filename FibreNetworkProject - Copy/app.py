from flask import Flask, render_template, request, redirect, url_for, session, send_file
import sqlite3
import csv
import io
import re
from datetime import datetime
import json
import pandas as pd
import os
import uuid
import glob
import secrets
import openpyxl
from dotenv import load_dotenv
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

load_dotenv()

# Optional: enable PDF statement parsing if installed
try:
    import pdfplumber  # type: ignore
except Exception:
    pdfplumber = None

# Optional OCR fallback for scanned PDFs (image-only)
try:
    from pdf2image import convert_from_path  # type: ignore
except Exception:
    convert_from_path = None

try:
    import pytesseract  # type: ignore
except Exception:
    pytesseract = None


def detect_tesseract_cmd():
    # Common Windows install paths (winget installer)
    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def detect_poppler_bin():
    # pdf2image needs the folder that contains pdftoppm.exe
    # 1) Respect explicit env override
    env_path = os.environ.get("POPPLER_BIN") or os.environ.get("POPPLER_PATH")
    if env_path and os.path.exists(env_path):
        return env_path

    # 2) Typical winget extraction location
    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        pattern = os.path.join(
            local,
            "Microsoft",
            "WinGet",
            "Packages",
            "oschwartz10612.Poppler_*",
            "poppler-*",
            "Library",
            "bin",
        )
        matches = sorted(glob.glob(pattern), reverse=True)
        for m in matches:
            if os.path.exists(os.path.join(m, "pdftoppm.exe")):
                return m

    # 3) Common manual install path
    candidates = [
        r"C:\poppler\Library\bin",
        r"C:\Program Files\poppler\Library\bin",
    ]
    for path in candidates:
        if os.path.exists(os.path.join(path, "pdftoppm.exe")):
            return path

    return None

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE = os.path.abspath(os.environ.get(
    "DATABASE_PATH",
    os.path.join(BASE_DIR, "database", "fiber.db"),
))
CURRENT_PAID_MAP = {}
PENDING_DIR = os.path.abspath(os.environ.get(
    "PENDING_STATEMENTS_DIR",
    os.path.join(BASE_DIR, "database", "pending_statements"),
))
os.makedirs(PENDING_DIR, exist_ok=True)
os.makedirs(os.path.dirname(DATABASE), exist_ok=True)

APP_ENV = os.environ.get("APP_ENV", "development").lower()
if APP_ENV == "production" and not all(
    os.environ.get(name) for name in ("SECRET_KEY", "ADMIN_USERNAME", "ADMIN_PASSWORD")
):
    raise RuntimeError(
        "Production requires SECRET_KEY, ADMIN_USERNAME, and ADMIN_PASSWORD."
    )

app.secret_key = os.environ.get("SECRET_KEY") or "nila-cable-fibre-networks-secret-2025"
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "change-me")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "change-me")


def table_exists(cur, name):
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (name,))
    return cur.fetchone() is not None


def table_columns(cur, name):
    cur.execute(f"PRAGMA table_info({name})")
    return [row[1] for row in cur.fetchall()]


def create_customer_table(cur):
    cur.execute('''
        CREATE TABLE IF NOT EXISTS customer (
            customer_id INTEGER PRIMARY KEY AUTOINCREMENT,
            box_id TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            phone TEXT NOT NULL,
            phone_number TEXT NOT NULL,
            subscription REAL NOT NULL,
            status TEXT NOT NULL,
            paid_status TEXT NOT NULL DEFAULT 'unpaid',
            connection_id TEXT,
            plan_amount REAL,
            upi_id TEXT
        )
    ''')


def create_payment_table(cur):
    cur.execute('''
        CREATE TABLE IF NOT EXISTS payment (
            payment_id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER,
            phone TEXT,
            amount REAL NOT NULL,
            payment_date TEXT NOT NULL,
            month TEXT NOT NULL,
            upi_id TEXT,
            transaction_id TEXT UNIQUE,
            status TEXT NOT NULL,
            FOREIGN KEY(customer_id) REFERENCES customer(customer_id)
        )
    ''')


def migrate_customer_table(cur, old_columns):
    cur.execute("ALTER TABLE customer RENAME TO customer_old")
    create_customer_table(cur)
    cur.execute(f'''
        INSERT INTO customer (
            customer_id, box_id, name, phone, phone_number, subscription, status, paid_status,
            connection_id, plan_amount, upi_id
        )
        SELECT
            id,
            box_id,
            name,
            phone,
            phone,
            subscription,
            status,
            paid_status,
            box_id,
            subscription,
            NULL
        FROM customer_old
    ''')
    cur.execute("DROP TABLE customer_old")


def migrate_payment_table(cur, old_columns):
    has_txn_id = 'txn_id' in old_columns
    cur.execute("ALTER TABLE payment RENAME TO payment_old")
    create_payment_table(cur)
    txn_expr = 'txn_id' if has_txn_id else 'NULL'
    cur.execute(f'''
        INSERT INTO payment (
            payment_id, customer_id, phone, amount, payment_date, month, upi_id, transaction_id, status
        )
        SELECT
            id,
            NULL,
            phone,
            amount,
            date,
            strftime('%Y-%m', date),
            NULL,
            {txn_expr},
            'Paid'
        FROM payment_old
    ''')
    cur.execute("DROP TABLE payment_old")


def ensure_schema():
    conn = sqlite3.connect(DATABASE, timeout=30.0)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA busy_timeout = 30000;")
    cur = conn.cursor()
    if not table_exists(cur, 'customer'):
        create_customer_table(cur)
    else:
        cols = table_columns(cur, 'customer')
        if 'customer_id' not in cols:
            migrate_customer_table(cur, cols)
    if not table_exists(cur, 'payment'):
        create_payment_table(cur)
    else:
        cols = table_columns(cur, 'payment')
        if 'payment_id' not in cols or 'payment_date' not in cols or 'month' not in cols:
            migrate_payment_table(cur, cols)
    cur.execute('''
        CREATE TABLE IF NOT EXISTS unmatched_transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            raw_description TEXT,
            amount REAL,
            payment_date TEXT,
            reason TEXT,
            logged_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()


ensure_schema()


def get_db_connection():
    conn = sqlite3.connect(DATABASE, timeout=30.0)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA busy_timeout = 30000;")
    conn.row_factory = sqlite3.Row
    return conn


def reconcile_payment_amounts():
    known_amounts = get_known_plan_amounts()
    max_expected = max(known_amounts) if known_amounts else 5000
    conn = get_db_connection()
    conn.execute('''
        UPDATE payment
        SET status = 'Needs Verification'
        WHERE payment_id IN (
            SELECT p.payment_id
            FROM payment p
            JOIN customer c ON c.customer_id = p.customer_id
            WHERE p.status = 'Paid'
              AND ABS(p.amount - COALESCE(c.plan_amount, c.subscription)) > 2
        )
    ''')
    conn.execute('''
        UPDATE payment
        SET status = 'Invalid Statement Row'
        WHERE status IN ('Unidentified Payment', 'Needs Verification')
          AND (
              amount > ?
              OR LENGTH(REPLACE(CAST(CAST(amount AS INTEGER) AS TEXT), '.', '')) > 6
          )
    ''', (max_expected * 5,))
    conn.commit()
    conn.close()


def normalize_customer_row(row):
    if not row:
        return None
    data = dict(row)
    data['customer_id'] = data.get('customer_id') or data.get('id')
    data['id'] = data['customer_id']
    data['phone'] = data.get('phone_number') or data.get('phone')
    data['plan_amount'] = data.get('plan_amount') or data.get('subscription')
    return data


def normalize_payment_row(row):
    data = dict(row)
    data['txn_id'] = data.get('transaction_id')
    data['date'] = data.get('payment_date')
    data['name'] = data.get('customer_name') or 'Unknown'
    return data


def get_all_customers():
    conn = get_db_connection()
    rows = conn.execute('SELECT * FROM customer ORDER BY name').fetchall()
    conn.close()
    return [normalize_customer_row(r) for r in rows]


def get_customer_by_id(customer_id):
    conn = get_db_connection()
    row = conn.execute('SELECT * FROM customer WHERE customer_id = ?', (customer_id,)).fetchone()
    conn.close()
    return normalize_customer_row(row)


def add_customer(box_id, name, phone, subscription, status, connection_id=None, plan_amount=None, upi_id=None):
    conn = get_db_connection()
    try:
        conn.execute('''
            INSERT INTO customer (
                box_id, name, phone, phone_number, subscription, status, paid_status,
                connection_id, plan_amount, upi_id
            ) VALUES (?, ?, ?, ?, ?, ?, 'unpaid', ?, ?, ?)
        ''', (
            box_id,
            name,
            phone,
            phone,
            subscription,
            status,
            connection_id or box_id,
            plan_amount or subscription,
            upi_id
        ))
        conn.commit()
        conn.close()
        return True, "Customer added successfully!"
    except sqlite3.IntegrityError as e:
        conn.close()
        return False, f"Error: {str(e)}"


def update_customer(customer_id, box_id, name, phone, subscription, status, connection_id=None, plan_amount=None, upi_id=None):
    conn = get_db_connection()
    try:
        conn.execute('''
            UPDATE customer 
            SET box_id=?, name=?, phone=?, phone_number=?, subscription=?, status=?, connection_id=?, plan_amount=?, upi_id=?
            WHERE customer_id=?
        ''', (
            box_id, name,
            phone, phone,
            subscription, status,
            connection_id, plan_amount, upi_id,
            customer_id
        ))
        conn.commit()
        conn.close()
        return True, "Customer updated successfully!"
    except sqlite3.IntegrityError as e:
        conn.close()
        return False, f"Error: {str(e)}"


def delete_customer(customer_id):
    conn = get_db_connection()
    try:
        conn.execute('DELETE FROM customer WHERE customer_id = ?', (customer_id,))
        conn.commit()
        conn.close()
        return True, "Customer deleted successfully!"
    except Exception as e:
        conn.close()
        return False, f"Error: {str(e)}"


def get_payment_by_phone(phone, month=None):
    conn = get_db_connection()
    if month is None:
        query = 'SELECT * FROM payment WHERE phone = ? ORDER BY payment_date DESC LIMIT 1'
        payment = conn.execute(query, (phone,)).fetchone()
    else:
        query = 'SELECT * FROM payment WHERE phone = ? AND month = ? LIMIT 1'
        payment = conn.execute(query, (phone, month)).fetchone()
    conn.close()
    return payment


def get_paid_payment_for_customer(customer_id, month):
    conn = get_db_connection()
    payment = conn.execute(
        '''
        SELECT *
        FROM payment
        WHERE customer_id = ? AND month = ? AND status = ?
        ORDER BY payment_date DESC, payment_id DESC
        LIMIT 1
        ''',
        (customer_id, month, 'Paid')
    ).fetchone()
    conn.close()
    return payment


def mark_customer_paid(customer_id, month=None, upi=None):
    if not customer_id:
        return
    conn = get_db_connection()
    if upi:
        conn.execute('UPDATE customer SET paid_status=?, upi_id=? WHERE customer_id=?', ('paid', upi, customer_id))
    else:
        conn.execute('UPDATE customer SET paid_status=? WHERE customer_id=?', ('paid', customer_id))
    conn.commit()
    conn.close()
    if month:
        CURRENT_PAID_MAP[(customer_id, month)] = True


def get_customer_by_phone(phone):
    if not phone:
        return None
    conn = get_db_connection()
    row = conn.execute('SELECT * FROM customer WHERE phone = ? OR phone_number = ? LIMIT 1', (phone, phone)).fetchone()
    conn.close()
    return normalize_customer_row(row)


def calculate_payment_status():
    reconcile_payment_amounts()
    conn = get_db_connection()
    customers = conn.execute('SELECT customer_id FROM customer').fetchall()
    current_month = datetime.now().strftime('%Y-%m')
    paid_rows = conn.execute(
        'SELECT customer_id, amount FROM payment WHERE month = ? AND status = ? AND customer_id IS NOT NULL',
        (current_month, 'Paid')
    ).fetchall()
    conn.close()
    paid_ids = {row['customer_id'] for row in paid_rows if row['customer_id']}
    paid_count = len(paid_ids)
    unpaid_count = max(len(customers) - paid_count, 0)
    total_collection = sum(row['amount'] for row in paid_rows)
    return {
        'total': len(customers),
        'paid': paid_count,
        'unpaid': unpaid_count,
        'collection': total_collection
    }


def calculate_payment_status_for_month(month):
    reconcile_payment_amounts()
    conn = get_db_connection()
    customers = conn.execute('SELECT customer_id FROM customer').fetchall()
    # Only count payments that are linked to a customer (customer_id IS NOT NULL)
    # Unlinked payments (customer_id NULL) cannot be attributed to a specific customer
    paid_rows = conn.execute(
        'SELECT customer_id, amount FROM payment WHERE month = ? AND status = ? AND customer_id IS NOT NULL',
        (month, 'Paid')
    ).fetchall()
    conn.close()
    paid_ids = {row['customer_id'] for row in paid_rows if row['customer_id']}
    paid_count = len(paid_ids)
    unpaid_count = max(len(customers) - paid_count, 0)
    total_collection = sum(row['amount'] for row in paid_rows)
    return {
        'total': len(customers),
        'paid': paid_count,
        'unpaid': unpaid_count,
        'collection': total_collection
    }


def get_latest_paid_month():
    conn = get_db_connection()
    row = conn.execute("SELECT MAX(month) as m FROM payment WHERE status = ?", ('Paid',)).fetchone()
    conn.close()
    return (row['m'] if row and row['m'] else None)


def get_default_billing_month():
    return get_latest_paid_month() or datetime.now().strftime('%Y-%m')


def get_available_months():
    conn = get_db_connection()
    rows = conn.execute("SELECT DISTINCT month FROM payment WHERE month IS NOT NULL AND month != '' ORDER BY month DESC").fetchall()
    conn.close()
    months = [r['month'] for r in rows if r['month']]
    current = datetime.now().strftime('%Y-%m')
    if current not in months:
        months.append(current)
    # Deduplicate and sort descending
    return sorted(list(set(months)), reverse=True)


def load_paid_map():
    global CURRENT_PAID_MAP
    conn = get_db_connection()
    rows = conn.execute('SELECT customer_id, month FROM payment WHERE status = ?', ('Paid',)).fetchall()
    conn.close()
    CURRENT_PAID_MAP = {(row['customer_id'], row['month']): True for row in rows if row['customer_id']}


def get_payment_history(limit=100):
    reconcile_payment_amounts()
    conn = get_db_connection()
    rows = conn.execute('''
        SELECT p.*, c.name as customer_name
        FROM payment p
        LEFT JOIN customer c ON c.customer_id = p.customer_id
        WHERE p.status != 'Invalid Statement Row'
        ORDER BY payment_date DESC, payment_id DESC
        LIMIT ?
    ''', (limit,)).fetchall()
    conn.close()
    return [normalize_payment_row(r) for r in rows]


def insert_payment_record(customer_id, phone, amount, payment_date, month, upi_id, transaction_id, status):
    conn = get_db_connection()
    conn.execute('''
        INSERT INTO payment (customer_id, phone, amount, payment_date, month, upi_id, transaction_id, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (customer_id, phone, amount, payment_date, month, upi_id, transaction_id, status))
    conn.commit()
    conn.close()


def log_unmatched_transaction(description, amount, payment_date, reason):
    conn = get_db_connection()
    conn.execute('''
        INSERT INTO unmatched_transactions (raw_description, amount, payment_date, reason)
        VALUES (?, ?, ?, ?)
    ''', (description, amount, payment_date, reason))
    conn.commit()
    conn.close()


def parse_amount(value):
    if not value:
        return None
    normalized = re.sub(r'[^\d\.]', '', str(value))
    try:
        return float(normalized)
    except ValueError:
        return None


def amounts_close(a, b, tolerance=2):
    if a is None or b is None:
        return False
    try:
        return abs(float(a) - float(b)) <= tolerance
    except (TypeError, ValueError):
        return False


def get_known_plan_amounts():
    conn = get_db_connection()
    rows = conn.execute('''
        SELECT COALESCE(plan_amount, subscription) as amount
        FROM customer
        WHERE COALESCE(plan_amount, subscription) IS NOT NULL
    ''').fetchall()
    conn.close()
    amounts = []
    for row in rows:
        try:
            amount = float(row['amount'])
        except (TypeError, ValueError):
            continue
        if amount > 0:
            amounts.append(amount)
    return amounts


def is_valid_customer_amount(amount, known_amounts=None):
    if amount is None:
        return False
    known_amounts = known_amounts if known_amounts is not None else get_known_plan_amounts()
    if not known_amounts:
        return 1 <= float(amount) <= 5000
    return any(amounts_close(amount, plan) for plan in known_amounts)


def parse_date(value):
    if not value:
        return None
    text = str(value).strip()
    formats = ['%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%d.%m.%Y', '%d %b %Y', '%d %B %Y']
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    try:
        parsed = datetime.fromisoformat(text)
        return parsed.strftime('%Y-%m-%d')
    except ValueError:
        return None


def extract_upi(description):
    if not description:
        return ''
    match = re.search(r'([\*\w\.]+@[\w\.]+)', description)
    return match.group(1) if match else ''


def get_upi_suffix(upi):
    if not upi:
        return ''
    parts = upi.lower().split('@')
    if len(parts) != 2:
        return upi.lower()
    user, domain = parts
    user = user.lstrip('*')
    suffix = user[-4:] if len(user) > 4 else user
    return f"{suffix}@{domain}"


def generate_transaction_id(transaction):
    base = (transaction.get('name') or 'txn').replace(' ', '')[:5]
    date_part = transaction['payment_date'].replace('-', '')
    amount_part = str(int(transaction['amount'])) if transaction['amount'] else '0'
    return f"UPI_{base}_{amount_part}_{date_part}"


def find_field(row, candidates):
    if not row:
        return ''
    lookup = {key.lower(): val for key, val in row.items()}
    lower_candidates = [c.lower() for c in candidates]
    for key in lookup:
        if key in lower_candidates:
            return lookup[key]
    for candidate in lower_candidates:
        if candidate in lookup:
            return lookup[candidate]
    return ''


def extract_transaction_from_row(row):
    name = find_field(row, ['customer name', 'beneficiary', 'name']) or 'Unknown'
    amount = parse_amount(find_field(row, ['amount', 'amt', 'credit', 'debit']))
    if amount is None:
        return None
    payment_date = parse_date(find_field(row, ['date', 'payment date', 'txn date']))
    if not payment_date:
        payment_date = datetime.now().strftime('%Y-%m-%d')
    upi_id = find_field(row, ['upi id', 'upi', 'vpa', 'receiver upi'])
    description = find_field(row, ['description', 'narration', 'remarks', 'details'])
    if not upi_id:
        upi_id = extract_upi(description)
    txn_id = find_field(row, ['transaction_id', 'txn id', 'ref', 'reference']) or ''
    if not txn_id:
        txn_id = generate_transaction_id({'name': name, 'amount': amount, 'payment_date': payment_date})
    return {
        'name': name.strip(),
        'amount': amount,
        'payment_date': payment_date,
        'month': payment_date[:7],
        'upi_id': upi_id.strip(),
        'description': description,
        'transaction_id': txn_id.strip(),
        'phone': find_field(row, ['phone', 'mobile'])
    }


def parse_statement_text(text):
    """
    Best-effort parser for statement text (TXT/PDF extracted text).
    Many bank PDFs don't contain customer name in a usable way, so name may be 'Unknown'.
    Matching will still work via UPI suffix + amount/month.

    Strategy:
    - Find lines containing a UPI VPA (something@bank)
    - Look in a small window around that line for a date and amount
    - Prefer credits (CR/credit/received) when indicated
    """
    lines = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        # normalize whitespace
        line = re.sub(r'\s+', ' ', line)
        lines.append(line)

    if not lines:
        return []

    upi_re = re.compile(r'([\*\w\.\-]+@[\w\.]+)')
    date_re = re.compile(r'(\d{4}-\d{2}-\d{2}|\d{2}[/-]\d{2}[/-]\d{4}|\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})')
    # Amount patterns: with or without currency, and optional CR/DR markers.
    amt_re = re.compile(r'(?:₹|INR|Rs\.?)\s*([0-9][0-9,]*\.?[0-9]{0,2})', re.IGNORECASE)
    plain_amt_re = re.compile(r'\b([0-9][0-9,]*\.?[0-9]{0,2})\b')
    known_plan_amounts = get_known_plan_amounts()
    bad_amount_words = (
        'balance', 'bal ', 'closing', 'opening', 'available', 'ledger',
        'ref', 'reference', 'utr', 'rrn', 'account', 'a/c', 'ac no',
        'mobile', 'phone', 'transaction id', 'txn id'
    )

    def find_date(window_text):
        dm = date_re.search(window_text)
        if not dm:
            return None
        return parse_date(dm.group(1))

    def amount_context(text, start, end, width=28):
        return text[max(0, start - width):min(len(text), end + width)].lower()

    def likely_transaction_amount(text, match):
        context = amount_context(text, match.start(), match.end())
        if any(word in context for word in bad_amount_words):
            return False
        if date_re.search(context):
            return False
        return True

    def find_amount(window_text, prefer_credit=True):
        txt = window_text.lower()
        credit_words = (' cr', 'credit', 'received', 'deposit', 'credited')
        debit_words = (' dr', 'debit', 'withdrawn', 'paid to')
        if prefer_credit and any(word in txt for word in debit_words) and not any(word in txt for word in credit_words):
            return None

        candidates = []
        for m in amt_re.finditer(window_text):
            val = parse_amount(m.group(1))
            if val is None or val < 1 or not is_valid_customer_amount(val, known_plan_amounts):
                continue
            if likely_transaction_amount(window_text, m):
                candidates.append(val)

        if not candidates:
            for m in plain_amt_re.finditer(window_text):
                val = parse_amount(m.group(1))
                if val is None or val < 1 or not is_valid_customer_amount(val, known_plan_amounts):
                    continue
                if not likely_transaction_amount(window_text, m):
                    continue
                raw = m.group(1).replace(',', '').replace('.', '')
                if len(raw) > 6:
                    continue
                candidates.append(val)

        if not candidates:
            return None

        plan_like = [v for v in candidates if is_valid_customer_amount(v, known_plan_amounts)]
        if plan_like:
            return min(plan_like)
        return None

    transactions = []
    seen = set()

    for idx, line in enumerate(lines):
        um = upi_re.search(line)
        if not um:
            continue

        # Build a local window around the UPI line to find date/amount.
        start = max(0, idx - 2)
        end = min(len(lines), idx + 4)
        window = " | ".join(lines[start:end])

        upi_id = um.group(1)
        payment_date = find_date(window) or datetime.now().strftime('%Y-%m-%d')
        amount = find_amount(window, prefer_credit=True)
        if amount is None:
            continue

        tx_id = generate_transaction_id({'name': get_upi_suffix(upi_id), 'amount': amount, 'payment_date': payment_date})
        if tx_id in seen:
            continue
        seen.add(tx_id)

        transactions.append({
            'name': 'Unknown',
            'amount': amount,
            'payment_date': payment_date,
            'month': payment_date[:7],
            'upi_id': upi_id,
            'description': window,
            'transaction_id': tx_id,
            'phone': ''
        })

    return transactions


def parse_pdf_statement(file_storage, password='', use_ocr=False):
    if not pdfplumber:
        raise ValueError("PDF parsing not available. Install pdfplumber: pip install pdfplumber")

    # pdfplumber can accept a file-like object
    try:
        with pdfplumber.open(file_storage, password=password or None) as pdf:
            full_text = ""
            for page in pdf.pages:
                full_text += (page.extract_text() or "") + "\n"
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypted" in msg:
            raise ValueError("This PDF looks encrypted/password protected. Please enter the PDF password, or download a CSV statement.") from e
        raise

    transactions = parse_statement_text(full_text)
    if transactions:
        return transactions

    # Common when the PDF is scanned (image-only) or text extraction is blocked.
    if use_ocr and convert_from_path and pytesseract:
        return ocr_pdf_statement(file_storage, password=password)

    raise ValueError(
        "No transactions found after unlocking the PDF. This usually means the PDF has no extractable text (scanned image PDF). "
        "If you cannot export CSV, you must enable OCR, or use a different statement format."
    )


def ocr_pdf_statement(file_storage, password=''):
    """
    OCR fallback for scanned PDFs. Requires:
    - pip install pytesseract pdf2image
    - Windows: install Poppler and Tesseract OCR, and add them to PATH
    """
    if not convert_from_path or not pytesseract:
        raise ValueError("OCR support not installed (need pytesseract + pdf2image).")

    # Make OCR work even if PATH was not refreshed yet.
    tesseract_cmd = detect_tesseract_cmd()
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    poppler_bin = detect_poppler_bin()
    if not poppler_bin:
        raise ValueError("Poppler not found. Install Poppler or set POPPLER_BIN to the folder containing pdftoppm.exe.")

    # Save upload to a temp file because pdf2image expects a file path
    tmp_name = f"ocr_{uuid.uuid4().hex}.pdf"
    tmp_path = os.path.join(PENDING_DIR, tmp_name)
    file_storage.stream.seek(0)
    with open(tmp_path, "wb") as f:
        f.write(file_storage.read())

    try:
        images = convert_from_path(tmp_path, poppler_path=poppler_bin)
        full_text = ""
        for img in images:
            full_text += pytesseract.image_to_string(img) + "\n"
        transactions = parse_statement_text(full_text)
        if not transactions:
            raise ValueError("OCR ran, but still no UPI+amount lines were detected in the PDF.")
        return transactions
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


def parse_statement_file(file):
    filename = (file.filename or "").lower()

    # PDF upload (common confusion): don't try to decode as UTF-8
    if filename.endswith('.pdf'):
        raise ValueError("PDF uploaded. Please use CSV/TXT, or use the PDF upload with password support.")

    content = file.read().decode('utf-8', errors='ignore')
    transactions = []

    if filename.endswith('.csv'):
        reader = csv.DictReader(io.StringIO(content))
        for row in reader:
            tx = extract_transaction_from_row(row)
            if tx:
                transactions.append(tx)
    else:
        # TXT: either comma-separated or tab-separated lines
        lines = [line for line in content.splitlines() if line.strip()]
        for line in lines:
            tokens = re.split(r'[,\t]+', line)
            row = {str(idx): token for idx, token in enumerate(tokens)}
            tx = extract_transaction_from_row(row)
            if tx:
                transactions.append(tx)

        # If structured parse failed, fall back to text parser (UPI+amount+date in same line)
        if not transactions:
            transactions = parse_statement_text(content)

    if not transactions:
        raise ValueError("No transactions parsed from statement. If this is a PDF, export CSV from your bank app, or upload PDF with password.")

    return transactions


def match_payment(transaction, customers):
    global CURRENT_PAID_MAP
    month = transaction['month']
    name = (transaction.get('name') or '').strip().lower()
    amount = transaction['amount']
    upi_suffix = get_upi_suffix(transaction.get('upi_id'))
    level1 = []
    level1_amount_mismatch = []
    level2 = []
    for cust in customers:
        if not cust or cust.get('customer_id') is None:
            continue
        paid_key = (cust['customer_id'], month)
        if CURRENT_PAID_MAP.get(paid_key):
            continue
        cust_upi_suffix = get_upi_suffix(cust.get('upi_id'))
        cust_amount = cust.get('plan_amount') or cust.get('subscription') or 0
        if upi_suffix and cust_upi_suffix and upi_suffix == cust_upi_suffix:
            if amounts_close(cust_amount, amount):
                level1.append(cust)
            else:
                mismatch = dict(cust)
                mismatch['expected_amount'] = cust_amount
                level1_amount_mismatch.append(mismatch)
            continue
        cust_name = cust.get('name', '').lower()
        if cust_name == name and amounts_close(cust_amount, amount):
            level2.append(cust)
    if level1:
        if len(level1) == 1:
            return {'status': 'Paid', 'customer': level1[0], 'reason': 'UPI + amount match'}
        return {'status': 'Needs Verification', 'candidates': level1, 'reason': 'Multiple UPI matches'}
    if level1_amount_mismatch:
        return {
            'status': 'Needs Verification',
            'candidates': level1_amount_mismatch,
            'reason': 'UPI matched but amount differs from plan'
        }
    if level2:
        if len(level2) == 1:
            return {'status': 'Paid', 'customer': level2[0], 'reason': 'Name + amount match'}
        return {'status': 'Needs Verification', 'candidates': level2, 'reason': 'Multiple name/amount matches'}
    return {'status': 'Unidentified Payment', 'reason': 'No customer matched'}


def process_transactions(transactions):
    load_paid_map()
    customers = get_all_customers()
    known_amounts = get_known_plan_amounts()
    results = []
    conn = get_db_connection()
    for transaction in transactions:
        txn_id = transaction.get('transaction_id') or generate_transaction_id(transaction)
        transaction['transaction_id'] = txn_id
        if not is_valid_customer_amount(transaction.get('amount'), known_amounts):
            log_unmatched_transaction(
                transaction.get('description'),
                transaction.get('amount'),
                transaction.get('payment_date'),
                'Ignored: extracted amount does not match any customer plan'
            )
            results.append({
                'transaction_id': txn_id,
                'status': 'Invalid Statement Row',
                'reason': 'Extracted amount does not match any customer plan',
                'amount': transaction.get('amount'),
                'month': transaction.get('month'),
                'name': transaction.get('name', 'Unknown')
            })
            continue
        exists = conn.execute('SELECT 1 FROM payment WHERE transaction_id = ?', (txn_id,)).fetchone()
        if exists:
            results.append({
                'transaction_id': txn_id,
                'status': 'Duplicate',
                'reason': 'Already recorded',
                'amount': transaction['amount'],
                'month': transaction['month'],
                'name': transaction.get('name', 'Unknown')
            })
            continue
        match = match_payment(transaction, customers)
        phone = transaction.get('phone')
        if match['status'] == 'Paid':
            cust = match['customer']
            customer_upi = cust.get('upi_id') or transaction.get('upi_id')
            insert_payment_record(
                cust['customer_id'],
                phone or cust.get('phone'),
                transaction['amount'],
                transaction['payment_date'],
                transaction['month'],
                customer_upi,
                txn_id,
                'Paid'
            )
            mark_customer_paid(cust['customer_id'], transaction['month'], customer_upi)
            results.append({
                'transaction_id': txn_id,
                'status': 'Paid',
                'reason': match['reason'],
                'amount': transaction['amount'],
                'month': transaction['month'],
                'name': cust.get('name', 'Unknown')
            })
        elif match['status'] == 'Needs Verification':
            insert_payment_record(
                None,
                phone,
                transaction['amount'],
                transaction['payment_date'],
                transaction['month'],
                transaction.get('upi_id'),
                txn_id,
                'Needs Verification'
            )
            log_unmatched_transaction(transaction.get('description'), transaction['amount'], transaction['payment_date'], match['reason'])
            results.append({
                'transaction_id': txn_id,
                'status': 'Needs Verification',
                'reason': match['reason'],
                'amount': transaction['amount'],
                'month': transaction['month'],
                'name': ', '.join(c.get('name', '') for c in match.get('candidates', []))
            })
        else:
            insert_payment_record(
                None,
                phone,
                transaction['amount'],
                transaction['payment_date'],
                transaction['month'],
                transaction.get('upi_id'),
                txn_id,
                'Unidentified Payment'
            )
            log_unmatched_transaction(transaction.get('description'), transaction['amount'], transaction['payment_date'], match['reason'])
            results.append({
                'transaction_id': txn_id,
                'status': 'Unidentified Payment',
                'reason': match['reason'],
                'amount': transaction['amount'],
                'month': transaction['month'],
                'name': transaction.get('name', 'Unknown')
            })
    conn.close()
    return results


def save_pending_transactions(transactions):
    # Flask's default session is a signed cookie; keep only a token there.
    token = uuid.uuid4().hex
    path = os.path.join(PENDING_DIR, f"{token}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(transactions, f, ensure_ascii=True)
    return token


def load_pending_transactions(token, delete_after=True):
    if not token:
        return None
    path = os.path.join(PENDING_DIR, f"{token}.json")
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data
    finally:
        if delete_after:
            try:
                os.remove(path)
            except OSError:
                pass


def pop_pending_transactions():
    token = session.pop("pending_token", None)
    return load_pending_transactions(token, delete_after=True)


def get_sidebar_metrics():
    conn = get_db_connection()
    pending_row = conn.execute(
        "SELECT COUNT(*) as cnt FROM payment WHERE status IN ('Needs Verification', 'Unidentified Payment')"
    ).fetchone()
    pending_review_count = pending_row['cnt'] if pending_row else 0
    active_row = conn.execute(
        "SELECT COUNT(*) as cnt FROM customer WHERE LOWER(status) = 'active'"
    ).fetchone()
    active_customers_count = active_row['cnt'] if active_row else 0
    conn.close()
    return active_customers_count, pending_review_count


def reconcile_database_payments():
    """Runs a full DB auto-reconciliation pass:
       1. Runs reconcile_payment_amounts() to flag amount mismatches.
       2. Finds unlinked payments and attempts to match them to customers by phone or UPI.
       3. Recalculates customer paid_status for corresponding months.
    """
    reconcile_payment_amounts()
    conn = get_db_connection()
    customers = conn.execute("SELECT customer_id, name, phone, phone_number, plan_amount, subscription, upi_id FROM customer").fetchall()
    unlinked_payments = conn.execute("SELECT payment_id, phone, amount, payment_date, month, upi_id, status FROM payment WHERE customer_id IS NULL AND status != 'Invalid Statement Row'").fetchall()
    
    matched_count = 0
    for p in unlinked_payments:
        pid = p['payment_id']
        p_phone = (p['phone'] or '').strip()
        p_upi = get_upi_suffix(p['upi_id'])
        p_amount = p['amount']
        
        matched_cust = None
        for c in customers:
            c_phone = (c['phone'] or c['phone_number'] or '').strip()
            c_upi = get_upi_suffix(c['upi_id'])
            c_amount = c['plan_amount'] or c['subscription'] or 0
            
            if p_phone and c_phone and p_phone == c_phone and amounts_close(p_amount, c_amount):
                matched_cust = c
                break
            if p_upi and c_upi and p_upi == c_upi and amounts_close(p_amount, c_amount):
                matched_cust = c
                break
                
        if matched_cust:
            conn.execute(
                "UPDATE payment SET customer_id = ?, status = 'Paid' WHERE payment_id = ?",
                (matched_cust['customer_id'], pid)
            )
            conn.execute(
                "UPDATE customer SET paid_status = 'paid', upi_id = COALESCE(?, upi_id) WHERE customer_id = ?",
                (p['upi_id'] or matched_cust['upi_id'], matched_cust['customer_id'])
            )
            matched_count += 1
            
    conn.commit()
    
    paid_cnt = conn.execute("SELECT COUNT(*) FROM payment WHERE status = 'Paid'").fetchone()[0]
    verify_cnt = conn.execute("SELECT COUNT(*) FROM payment WHERE status = 'Needs Verification'").fetchone()[0]
    unidentified_cnt = conn.execute("SELECT COUNT(*) FROM payment WHERE status = 'Unidentified Payment'").fetchone()[0]
    conn.close()
    
    return {
        'reconciled_unlinked': matched_count,
        'paid_count': paid_cnt,
        'verify_count': verify_cnt,
        'unidentified_count': unidentified_cnt
    }


def generate_payment_report_excel(month=None):
    reconcile_payment_amounts()
    conn = get_db_connection()

    if month:
        payments_query = '''
            SELECT p.payment_id, p.transaction_id, p.payment_date, p.month,
                   COALESCE(c.name, 'Unidentified') as customer_name,
                   COALESCE(c.box_id, '—') as box_id,
                   COALESCE(c.connection_id, '—') as connection_id,
                   COALESCE(p.phone, c.phone, '—') as phone,
                   COALESCE(c.plan_amount, c.subscription, 0) as plan_amount,
                   p.amount as paid_amount,
                   p.upi_id,
                   p.status
            FROM payment p
            LEFT JOIN customer c ON c.customer_id = p.customer_id
            WHERE p.month = ? AND p.status != 'Invalid Statement Row'
            ORDER BY p.payment_date DESC, p.payment_id DESC
        '''
        payments = conn.execute(payments_query, (month,)).fetchall()
    else:
        payments_query = '''
            SELECT p.payment_id, p.transaction_id, p.payment_date, p.month,
                   COALESCE(c.name, 'Unidentified') as customer_name,
                   COALESCE(c.box_id, '—') as box_id,
                   COALESCE(c.connection_id, '—') as connection_id,
                   COALESCE(p.phone, c.phone, '—') as phone,
                   COALESCE(c.plan_amount, c.subscription, 0) as plan_amount,
                   p.amount as paid_amount,
                   p.upi_id,
                   p.status
            FROM payment p
            LEFT JOIN customer c ON c.customer_id = p.customer_id
            WHERE p.status != 'Invalid Statement Row'
            ORDER BY p.payment_date DESC, p.payment_id DESC
        '''
        payments = conn.execute(payments_query).fetchall()

    target_month = month or datetime.now().strftime('%Y-%m')
    unpaid_query = '''
        SELECT c.customer_id, c.box_id, c.connection_id, c.name, c.phone,
               COALESCE(c.plan_amount, c.subscription, 0) as plan_amount,
               c.status as connection_status
        FROM customer c
        WHERE LOWER(c.status) = 'active'
          AND c.customer_id NOT IN (
              SELECT customer_id FROM payment 
              WHERE month = ? AND status = 'Paid' AND customer_id IS NOT NULL
          )
        ORDER BY c.name ASC
    '''
    unpaid_customers = conn.execute(unpaid_query, (target_month,)).fetchall()

    stats = calculate_payment_status_for_month(target_month)
    active_row = conn.execute(
        "SELECT COUNT(*) as cnt, SUM(COALESCE(plan_amount, subscription)) as total_due FROM customer WHERE LOWER(status) = 'active'"
    ).fetchone()
    total_potential = round(active_row['total_due'] or 0.0, 2)
    active_customers_count = active_row['cnt'] or 0
    collection = stats['collection']
    realization_rate = round((collection / total_potential * 100), 1) if total_potential > 0 else 0
    outstanding_dues = max(0.0, round(total_potential - collection, 2))

    conn.close()

    wb = openpyxl.Workbook()

    title_font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11, color="1F2937")
    bold_font = Font(name="Calibri", size=11, bold=True, color="1F2937")

    brand_fill = PatternFill(start_color="151A28", end_color="151A28", fill_type="solid")
    header_fill = PatternFill(start_color="252D3D", end_color="252D3D", fill_type="solid")
    alt_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    paid_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    paid_font = Font(name="Calibri", size=11, bold=True, color="15803D")

    verify_fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    verify_font = Font(name="Calibri", size=11, bold=True, color="A16207")

    unident_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    unident_font = Font(name="Calibri", size=11, bold=True, color="B91C1C")

    thin_border_side = Side(border_style="thin", color="D1D5DB")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # 1. Payment Transactions Sheet
    ws1 = wb.active
    ws1.title = "Payment Transactions"
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:L1")
    title_cell = ws1["A1"]
    title_cell.value = f"NILA CABLE FIBRE NETWORKS — PAYMENT REPORT {'(' + month + ')' if month else '(ALL PERIODS)'}"
    title_cell.font = title_font
    title_cell.fill = brand_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 36

    ws1.merge_cells("A2:L2")
    sub_cell = ws1["A2"]
    sub_cell.value = f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Total Transactions: {len(payments)}"
    sub_cell.font = Font(name="Calibri", size=10, italic=True, color="6B7280")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 20

    headers1 = [
        "Payment ID", "Transaction ID", "Payment Date", "Billing Month",
        "Customer Name", "Box ID", "Connection ID", "Phone",
        "Plan Amount (₹)", "Amount Paid (₹)", "UPI / Ref", "Status"
    ]
    ws1.append([])
    ws1.row_dimensions[3].height = 8
    ws1.append(headers1)
    ws1.row_dimensions[4].height = 26

    for col_num in range(1, len(headers1) + 1):
        c = ws1.cell(row=4, column=col_num)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = cell_border

    current_row = 5
    total_amount_sum = 0
    for p in payments:
        amount = float(p['paid_amount'] or 0)
        total_amount_sum += amount
        plan = float(p['plan_amount'] or 0)
        row_data = [
            p['payment_id'],
            p['transaction_id'] or '—',
            p['payment_date'],
            p['month'],
            p['customer_name'],
            p['box_id'],
            p['connection_id'],
            p['phone'],
            plan,
            amount,
            p['upi_id'] or '—',
            p['status']
        ]
        ws1.append(row_data)
        ws1.row_dimensions[current_row].height = 20

        for col_idx in range(1, len(row_data) + 1):
            cell = ws1.cell(row=current_row, column=col_idx)
            cell.font = data_font
            cell.border = cell_border
            if current_row % 2 == 0:
                cell.fill = alt_fill

            if col_idx in (1, 3, 4, 6, 7, 8):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in (9, 10):
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '"₹"#,##0.00'
            elif col_idx == 12:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                st = str(p['status']).strip()
                if st == 'Paid':
                    cell.fill = paid_fill
                    cell.font = paid_font
                elif st == 'Needs Verification':
                    cell.fill = verify_fill
                    cell.font = verify_font
                else:
                    cell.fill = unident_fill
                    cell.font = unident_font
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

        current_row += 1

    ws1.append(["TOTAL", "", "", "", "", "", "", "", "", total_amount_sum, "", ""])
    ws1.row_dimensions[current_row].height = 24
    ws1.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=9)
    total_label = ws1.cell(row=current_row, column=1)
    total_label.font = bold_font
    total_label.alignment = Alignment(horizontal="right", vertical="center")
    total_label.border = cell_border

    tot_val = ws1.cell(row=current_row, column=10)
    tot_val.font = bold_font
    tot_val.number_format = '"₹"#,##0.00'
    tot_val.alignment = Alignment(horizontal="right", vertical="center")
    tot_val.border = cell_border

    for c_idx in (11, 12):
        ws1.cell(row=current_row, column=c_idx).border = cell_border

    for col in ws1.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row in (1, 2):
                continue
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws1.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # 2. Unpaid Subscribers Sheet
    ws2 = wb.create_sheet(title="Unpaid Subscribers")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:G1")
    t2 = ws2["A1"]
    t2.value = f"UNPAID SUBSCRIBERS — PERIOD: {target_month}"
    t2.font = title_font
    t2.fill = brand_fill
    t2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 36

    headers2 = ["Customer ID", "Box ID", "Connection ID", "Customer Name", "Phone", "Monthly Plan (₹)", "Status"]
    ws2.append([])
    ws2.row_dimensions[2].height = 8
    ws2.append(headers2)
    ws2.row_dimensions[3].height = 26

    for col_num in range(1, len(headers2) + 1):
        c = ws2.cell(row=3, column=col_num)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = cell_border

    r2_idx = 4
    total_unpaid_due = 0
    for u in unpaid_customers:
        plan = float(u['plan_amount'] or 0)
        total_unpaid_due += plan
        r_data = [
            u['customer_id'],
            u['box_id'],
            u['connection_id'] or '—',
            u['name'],
            u['phone'],
            plan,
            u['connection_status'].capitalize()
        ]
        ws2.append(r_data)
        ws2.row_dimensions[r2_idx].height = 20
        for col_idx in range(1, len(r_data) + 1):
            cell = ws2.cell(row=r2_idx, column=col_idx)
            cell.font = data_font
            cell.border = cell_border
            if r2_idx % 2 == 0:
                cell.fill = alt_fill
            if col_idx in (1, 2, 3, 5, 7):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 6:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '"₹"#,##0.00'
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        r2_idx += 1

    ws2.append(["TOTAL PENDING DUES", "", "", "", "", total_unpaid_due, ""])
    ws2.merge_cells(start_row=r2_idx, start_column=1, end_row=r2_idx, end_column=5)
    ws2.cell(row=r2_idx, column=1).font = bold_font
    ws2.cell(row=r2_idx, column=1).alignment = Alignment(horizontal="right", vertical="center")
    ws2.cell(row=r2_idx, column=1).border = cell_border
    due_c = ws2.cell(row=r2_idx, column=6)
    due_c.font = bold_font
    due_c.number_format = '"₹"#,##0.00'
    due_c.alignment = Alignment(horizontal="right", vertical="center")
    due_c.border = cell_border
    ws2.cell(row=r2_idx, column=7).border = cell_border

    for col in ws2.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row == 1:
                continue
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws2.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # 3. Collection Summary Sheet
    ws3 = wb.create_sheet(title="Collection Summary")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:B1")
    t3 = ws3["A1"]
    t3.value = f"REVENUE & BILLING SUMMARY — {target_month}"
    t3.font = title_font
    t3.fill = brand_fill
    t3.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 36

    summary_rows = [
        ("Billing Period", target_month),
        ("Active Subscriptions", active_customers_count),
        ("Total Potential Revenue", total_potential),
        ("Total Realized Collection", collection),
        ("Outstanding Pending Dues", outstanding_dues),
        ("Collection Realization Rate", f"{realization_rate}%"),
        ("Paid Customers Count", stats['paid']),
        ("Unpaid Customers Count", stats['unpaid']),
        ("Total Payment Transactions Recorded", len(payments)),
    ]

    ws3.append([])
    ws3.append(["Metric", "Value"])
    ws3.row_dimensions[3].height = 24
    ws3.cell(row=3, column=1).font = header_font
    ws3.cell(row=3, column=1).fill = header_fill
    ws3.cell(row=3, column=1).border = cell_border
    ws3.cell(row=3, column=2).font = header_font
    ws3.cell(row=3, column=2).fill = header_fill
    ws3.cell(row=3, column=2).border = cell_border

    r3_idx = 4
    for label, val in summary_rows:
        ws3.append([label, val])
        ws3.row_dimensions[r3_idx].height = 22
        c1 = ws3.cell(row=r3_idx, column=1)
        c2 = ws3.cell(row=r3_idx, column=2)
        c1.font = bold_font
        c1.border = cell_border
        c2.font = data_font
        c2.border = cell_border
        if r3_idx % 2 == 0:
            c1.fill = alt_fill
            c2.fill = alt_fill
        if isinstance(val, (int, float)) and ("Revenue" in label or "Collection" in label or "Dues" in label):
            c2.number_format = '"₹"#,##0.00'
            c2.alignment = Alignment(horizontal="right", vertical="center")
        else:
            c2.alignment = Alignment(horizontal="left", vertical="center")
        r3_idx += 1

    ws3.column_dimensions["A"].width = 36
    ws3.column_dimensions["B"].width = 24

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream


def generate_payment_report_csv(month=None):
    reconcile_payment_amounts()
    conn = get_db_connection()
    if month:
        query = '''
            SELECT p.payment_id, p.transaction_id, p.payment_date, p.month,
                   COALESCE(c.name, 'Unidentified') as customer_name,
                   COALESCE(c.box_id, '') as box_id,
                   COALESCE(c.connection_id, '') as connection_id,
                   COALESCE(p.phone, c.phone, '') as phone,
                   COALESCE(c.plan_amount, c.subscription, 0) as plan_amount,
                   p.amount, p.upi_id, p.status
            FROM payment p
            LEFT JOIN customer c ON c.customer_id = p.customer_id
            WHERE p.month = ? AND p.status != 'Invalid Statement Row'
            ORDER BY p.payment_date DESC, p.payment_id DESC
        '''
        rows = conn.execute(query, (month,)).fetchall()
    else:
        query = '''
            SELECT p.payment_id, p.transaction_id, p.payment_date, p.month,
                   COALESCE(c.name, 'Unidentified') as customer_name,
                   COALESCE(c.box_id, '') as box_id,
                   COALESCE(c.connection_id, '') as connection_id,
                   COALESCE(p.phone, c.phone, '') as phone,
                   COALESCE(c.plan_amount, c.subscription, 0) as plan_amount,
                   p.amount, p.upi_id, p.status
            FROM payment p
            LEFT JOIN customer c ON c.customer_id = p.customer_id
            WHERE p.status != 'Invalid Statement Row'
            ORDER BY p.payment_date DESC, p.payment_id DESC
        '''
        rows = conn.execute(query).fetchall()
    conn.close()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "Payment ID", "Transaction ID", "Date", "Month",
        "Customer Name", "Box ID", "Connection ID", "Phone",
        "Plan Amount", "Amount Paid", "UPI ID", "Status"
    ])
    for r in rows:
        writer.writerow([
            r['payment_id'], r['transaction_id'], r['payment_date'], r['month'],
            r['customer_name'], r['box_id'], r['connection_id'], r['phone'],
            r['plan_amount'], r['amount'], r['upi_id'], r['status']
        ])
    output.seek(0)
    return io.BytesIO(output.getvalue().encode('utf-8'))


@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            session["logged_in"] = True
            return redirect(url_for("dashboard"))
        else:
            return "Invalid Username or Password"
    return render_template("login.html")


@app.route("/dashboard")
def dashboard():
    available_months = get_available_months()
    requested_month = request.args.get("month", "").strip()
    month = requested_month or get_default_billing_month()

    stats = calculate_payment_status_for_month(month)
    total_customers = stats['total']
    paid_customers = stats['paid']
    unpaid_customers = stats['unpaid']
    collection = stats['collection']
    chart_labels = json.dumps(["Paid", "Unpaid"])
    chart_values = json.dumps([paid_customers, unpaid_customers])

    conn = get_db_connection()
    week_query = '''
        SELECT payment_date, SUM(amount) as total
        FROM payment
        WHERE month = ? AND status = 'Paid'
        GROUP BY payment_date
        ORDER BY payment_date
    '''
    daily_rows = conn.execute(week_query, (month,)).fetchall()
    week_labels = json.dumps([r['payment_date'] for r in daily_rows])
    week_values = json.dumps([round(r['total'] or 0, 2) for r in daily_rows])

    br_query = '''SELECT status, COUNT(*) as cnt FROM customer GROUP BY status'''
    br_rows = conn.execute(br_query).fetchall()
    breakdown = {r['status']: r['cnt'] for r in br_rows}
    breakdown_labels = json.dumps(list(breakdown.keys()))
    breakdown_values = json.dumps(list(breakdown.values()))

    trend_query = '''
        SELECT payment_date, COUNT(*) as cnt
        FROM payment
        WHERE month = ? AND status != 'Invalid Statement Row'
        GROUP BY payment_date
        ORDER BY payment_date
    '''
    trend_rows = conn.execute(trend_query, (month,)).fetchall()
    trend_labels = [r['payment_date'] for r in trend_rows]
    trend_values = [r['cnt'] for r in trend_rows]
    trend_labels_js = json.dumps(trend_labels)
    trend_values_js = json.dumps(trend_values)

    # Calculate collection potential & realization metrics
    active_row = conn.execute(
        "SELECT COUNT(*) as cnt, SUM(COALESCE(plan_amount, subscription)) as total_due FROM customer WHERE LOWER(status) = 'active'"
    ).fetchone()
    total_potential = round(active_row['total_due'] or 0.0, 2)
    active_customers_count = active_row['cnt'] or 0
    realization_rate = round((collection / total_potential * 100), 1) if total_potential > 0 else 0
    outstanding_dues = max(0.0, round(total_potential - collection, 2))

    # Alert badge: pending verification payments
    pending_row = conn.execute(
        "SELECT COUNT(*) as cnt FROM payment WHERE status IN ('Needs Verification', 'Unidentified Payment')"
    ).fetchone()
    pending_review_count = pending_row['cnt'] if pending_row else 0

    # 6-Month Monthly Collection History
    monthly_history_rows = conn.execute('''
        SELECT month, SUM(amount) as total
        FROM payment
        WHERE status = 'Paid' AND month IS NOT NULL AND month != ''
        GROUP BY month
        ORDER BY month DESC
        LIMIT 12
    ''').fetchall()
    monthly_history_rows = list(reversed(monthly_history_rows))
    monthly_history_labels = [r['month'] for r in monthly_history_rows]
    monthly_history_values = [round(r['total'] or 0, 2) for r in monthly_history_rows]

    # Plan / Tariff tier breakdown
    plan_rows = conn.execute('''
        SELECT CAST(COALESCE(plan_amount, subscription) AS TEXT) as plan_tier, COUNT(*) as cnt
        FROM customer
        GROUP BY plan_tier
        ORDER BY COALESCE(plan_amount, subscription) ASC
    ''').fetchall()
    plan_labels = [f"₹{r['plan_tier']}" for r in plan_rows if r['plan_tier']]
    plan_values = [r['cnt'] for r in plan_rows if r['plan_tier']]

    conn.close()

    return render_template(
        "dashboard.html",
        dashboard_month=month,
        available_months=available_months,
        total_customers=total_customers,
        paid_customers=paid_customers,
        unpaid_customers=unpaid_customers,
        collection=collection,
        total_potential=total_potential,
        realization_rate=realization_rate,
        outstanding_dues=outstanding_dues,
        active_customers_count=active_customers_count,
        pending_review_count=pending_review_count,
        chart_labels=chart_labels,
        chart_values=chart_values,
        week_labels=week_labels,
        week_values=week_values,
        breakdown_labels=breakdown_labels,
        breakdown_values=breakdown_values,
        trend_labels=trend_labels_js,
        trend_values=trend_values_js,
        monthly_history_labels=json.dumps(monthly_history_labels),
        monthly_history_values=json.dumps(monthly_history_values),
        plan_labels=json.dumps(plan_labels),
        plan_values=json.dumps(plan_values),
    )


@app.route("/customers")
def customers():
    available_months = get_available_months()
    customers_list = get_all_customers()
    enhanced = []
    requested_month = request.args.get("month", "").strip()
    month = requested_month or get_default_billing_month()
    for c in customers_list:
        cust = dict(c)
        cust['txn_id_from_payment'] = ''
        cust['payment_amount'] = ''
        cust['payment_month'] = month
        cust['payment_date'] = ''
        payment = get_paid_payment_for_customer(cust['customer_id'], month)
        if payment:
            cust['txn_id_from_payment'] = payment['transaction_id']
            cust['payment_amount'] = payment['amount']
            cust['payment_date'] = payment['payment_date']
            if cust.get('paid_status') != 'paid':
                conn = get_db_connection()
                conn.execute("UPDATE customer SET paid_status='paid' WHERE customer_id=?", (cust['customer_id'],))
                conn.commit()
                conn.close()
            cust['paid_status'] = 'paid'
        else:
            cust['paid_status'] = 'unpaid'
        enhanced.append(cust)

    conn = get_db_connection()
    pending_row = conn.execute("SELECT COUNT(*) as cnt FROM payment WHERE status IN ('Needs Verification', 'Unidentified Payment')").fetchone()
    pending_review_count = pending_row['cnt'] if pending_row else 0
    active_row = conn.execute("SELECT COUNT(*) as cnt FROM customer WHERE LOWER(status) = 'active'").fetchone()
    active_customers_count = active_row['cnt'] if active_row else 0
    conn.close()

    return render_template(
        "view_customers.html",
        customers=enhanced,
        selected_month=month,
        available_months=available_months,
        pending_review_count=pending_review_count,
        active_customers_count=active_customers_count
    )


@app.route("/upload-customers", methods=["GET", "POST"])
def upload_customers():
    active_customers_count, pending_review_count = get_sidebar_metrics()
    if request.method == "POST":
        if 'file' not in request.files:
            return render_template("upload_customers.html", error="No file part in request",
                                   active_customers_count=active_customers_count, pending_review_count=pending_review_count)
        file = request.files['file']
        if file.filename == "":
            return render_template("upload_customers.html", error="No file selected",
                                   active_customers_count=active_customers_count, pending_review_count=pending_review_count)
        try:
            df = pd.read_excel(file)
            required = ['box_id', 'name', 'phone', 'subscription', 'status']
            for col in required:
                if col not in df.columns.str.lower() and col not in df.columns:
                    return render_template("upload_customers.html", error=f"Missing column: {col}",
                                           active_customers_count=active_customers_count, pending_review_count=pending_review_count)
            added = 0
            errors = []
            for idx, row in df.iterrows():
                try:
                    box_id = row.get('box_id') or row.get('Box ID')
                    name = row.get('name') or row.get('Name')
                    phone = row.get('phone') or row.get('Phone')
                    subscription = row.get('subscription') or row.get('Subscription')
                    status = row.get('status') or row.get('Status')
                    connection = row.get('connection_id') or row.get('Connection ID') or box_id
                    plan_amount = row.get('plan_amount') or row.get('Plan Amount') or subscription
                    upi_id = row.get('upi_id') or row.get('UPI ID') or ''
                    if pd.isna(box_id) or pd.isna(name) or pd.isna(phone) or pd.isna(subscription) or pd.isna(status):
                        raise ValueError("One or more required values missing")
                    success, msg = add_customer(
                        str(box_id).strip(),
                        str(name).strip(),
                        str(phone).strip(),
                        float(subscription),
                        str(status).strip().lower(),
                        connection_id=str(connection).strip(),
                        plan_amount=float(plan_amount) if plan_amount else None,
                        upi_id=str(upi_id).strip()
                    )
                    if success:
                        added += 1
                    else:
                        errors.append(msg)
                except Exception as e:
                    errors.append(f"Row {idx+2}: {e}")
            message = f"Imported {added} customers."
            if errors:
                message += " Some rows failed: " + "; ".join(errors[:5])
            active_customers_count, pending_review_count = get_sidebar_metrics()
            return render_template("upload_customers.html", success=message,
                                   active_customers_count=active_customers_count, pending_review_count=pending_review_count)
        except Exception as e:
            return render_template("upload_customers.html", error=f"Error processing file: {e}",
                                   active_customers_count=active_customers_count, pending_review_count=pending_review_count)
    return render_template("upload_customers.html",
                           active_customers_count=active_customers_count, pending_review_count=pending_review_count)


@app.route("/add-customer", methods=["GET", "POST"])
def add_customer_page():
    active_customers_count, pending_review_count = get_sidebar_metrics()
    if request.method == "POST":
        box_id = request.form.get("box_id")
        name = request.form.get("name")
        phone = request.form.get("phone")
        subscription = request.form.get("subscription")
        status = request.form.get("status")
        connection = request.form.get("connection_id") or box_id
        plan_amount = request.form.get("plan_amount") or subscription
        upi_id = request.form.get("upi_id")
        success, message = add_customer(
            box_id,
            name,
            phone,
            float(subscription),
            status,
            connection_id=connection,
            plan_amount=float(plan_amount) if plan_amount else float(subscription),
            upi_id=upi_id
        )
        if success:
            return redirect(url_for("customers"))
        else:
            return render_template("add_customer.html", error=message,
                                   active_customers_count=active_customers_count, pending_review_count=pending_review_count)
    return render_template("add_customer.html",
                           active_customers_count=active_customers_count, pending_review_count=pending_review_count)


@app.route("/edit-customer/<int:customer_id>", methods=["GET", "POST"])
def edit_customer_page(customer_id):
    active_customers_count, pending_review_count = get_sidebar_metrics()
    customer = get_customer_by_id(customer_id)
    if not customer:
        return "Customer not found", 404
    if request.method == "POST":
        box_id = request.form.get("box_id")
        name = request.form.get("name")
        phone = request.form.get("phone")
        subscription = request.form.get("subscription")
        status = request.form.get("status")
        connection = request.form.get("connection_id") or box_id
        plan_amount = request.form.get("plan_amount") or subscription
        upi_id = request.form.get("upi_id")
        success, message = update_customer(
            customer_id,
            box_id,
            name,
            phone,
            float(subscription),
            status,
            connection_id=connection,
            plan_amount=float(plan_amount) if plan_amount else float(subscription),
            upi_id=upi_id
        )
        if success:
            return redirect(url_for("customers"))
        else:
            return render_template("edit_customer.html", customer=customer, error=message,
                                   active_customers_count=active_customers_count, pending_review_count=pending_review_count)
    return render_template("edit_customer.html", customer=customer,
                           active_customers_count=active_customers_count, pending_review_count=pending_review_count)


@app.route("/delete-customer/<int:customer_id>", methods=["GET", "POST"])
def delete_customer_page(customer_id):
    delete_customer(customer_id)
    return redirect(url_for("customers"))


@app.route("/payments")
def payments():
    payments_list = get_payment_history(limit=500)
    available_months = get_available_months()
    active_customers_count, pending_review_count = get_sidebar_metrics()
    return render_template(
        "payments.html",
        payments=payments_list,
        available_months=available_months,
        pending_review_count=pending_review_count,
        active_customers_count=active_customers_count
    )


@app.route("/add-payment", methods=["GET", "POST"])
def add_payment_page():
    active_customers_count, pending_review_count = get_sidebar_metrics()
    if request.method == "POST":
        txn_id = request.form.get("txn_id")
        phone = request.form.get("phone")
        amount = request.form.get("amount")
        payment_date = request.form.get("date")
        upi_id = request.form.get("upi_id")
        month = datetime.strptime(payment_date, '%Y-%m-%d').strftime('%Y-%m')
        try:
            # Resolve customer by phone so the payment is linked (enables dashboard stats)
            customer = get_customer_by_phone(phone)
            customer_id = customer['customer_id'] if customer else None
            insert_payment_record(
                customer_id,
                phone,
                float(amount),
                payment_date,
                month,
                upi_id or (customer.get('upi_id') if customer else None),
                txn_id,
                'Paid'
            )
            if customer:
                mark_customer_paid(customer_id, month, upi_id or customer.get('upi_id'))
            return redirect(url_for("payments"))
        except sqlite3.IntegrityError as e:
            error = f"Database error: {e}"
            return render_template("add_payment.html", error=error, today=datetime.utcnow().strftime('%Y-%m-%d'),
                                   active_customers_count=active_customers_count, pending_review_count=pending_review_count)
    return render_template("add_payment.html", today=datetime.utcnow().strftime('%Y-%m-%d'),
                           active_customers_count=active_customers_count, pending_review_count=pending_review_count)


@app.route("/delete-payment/<int:payment_id>", methods=["POST", "GET"])
def delete_payment_page(payment_id):
    conn = get_db_connection()
    conn.execute('DELETE FROM payment WHERE payment_id = ?', (payment_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("payments"))


@app.route("/verify-payment/<int:payment_id>", methods=["POST"])
def verify_payment(payment_id):
    # Mark a payment (usually one flagged as 'Needs Verification') as verified/paid.
    conn = get_db_connection()
    row = conn.execute('SELECT * FROM payment WHERE payment_id = ?', (payment_id,)).fetchone()
    if not row:
        conn.close()
        return "Payment not found", 404

    # Update payment status to Paid
    conn.execute('UPDATE payment SET status = ? WHERE payment_id = ?', ('Paid', payment_id))

    # If payment isn't linked to a customer, try to attach via phone
    customer_id = row['customer_id']
    phone = row['phone']
    month = row['month']
    upi_id = row['upi_id']

    conn.commit()
    conn.close()

    # Associate and mark customer paid if possible
    if not customer_id and phone:
        cust = get_customer_by_phone(phone)
        if cust:
            conn2 = get_db_connection()
            conn2.execute('UPDATE payment SET customer_id = ? WHERE payment_id = ?', (cust['customer_id'], payment_id))
            conn2.commit()
            conn2.close()
            mark_customer_paid(cust['customer_id'], month, upi_id or cust.get('upi_id'))
    else:
        if customer_id:
            mark_customer_paid(customer_id, month, upi_id)

    return redirect(url_for('payments'))


@app.route("/reports")
def reports():
    stats = calculate_payment_status()
    conn = get_db_connection()
    monthly_rows = conn.execute('''
        SELECT month, SUM(amount) as total
        FROM payment
        WHERE status = ?
        GROUP BY month
        ORDER BY month DESC
        LIMIT 6
    ''', ('Paid',)).fetchall()
    recent_rows = conn.execute('''
        SELECT p.*, c.name as customer_name
        FROM payment p
        LEFT JOIN customer c ON c.customer_id = p.customer_id
        ORDER BY payment_date DESC
        LIMIT 6
    ''').fetchall()
    unmatched = conn.execute('SELECT COUNT(*) as cnt FROM unmatched_transactions').fetchone()['cnt']
    conn.close()
    return render_template(
        "reports.html",
        stats=stats,
        monthly=monthly_rows,
        recent=recent_rows,
        unmatched=unmatched
    )


@app.route("/delete-payments-by-month", methods=["POST"])
def delete_payments_by_month():
    month = request.form.get("month", "").strip()
    if not month:
        return redirect(url_for("reports"))
    conn = get_db_connection()
    affected = conn.execute(
        "SELECT DISTINCT customer_id FROM payment WHERE month = ? AND customer_id IS NOT NULL",
        (month,)
    ).fetchall()
    affected_ids = [r["customer_id"] for r in affected]
    conn.execute("DELETE FROM payment WHERE month = ?", (month,))
    conn.commit()
    for cid in affected_ids:
        still_paid = conn.execute(
            "SELECT COUNT(*) as cnt FROM payment WHERE customer_id = ? AND status = 'Paid'",
            (cid,)
        ).fetchone()["cnt"]
        if still_paid == 0:
            conn.execute(
                "UPDATE customer SET paid_status = 'unpaid' WHERE customer_id = ?",
                (cid,)
            )
    conn.commit()
    conn.close()
    return redirect(url_for("reports"))


@app.route("/delete-all-payments", methods=["POST"])
def delete_all_payments():
    conn = get_db_connection()
    conn.execute("DELETE FROM payment")
    conn.execute("DELETE FROM unmatched_transactions")
    conn.execute("UPDATE customer SET paid_status = 'unpaid'")
    conn.commit()
    conn.close()
    # redirect back to where the user came from
    referrer = request.referrer or url_for("reports")
    return redirect(referrer)


@app.route("/settings")
def settings():
    active_customers_count, pending_review_count = get_sidebar_metrics()
    return render_template("settings.html", active_customers_count=active_customers_count, pending_review_count=pending_review_count)


@app.route("/upload_statement", methods=["GET", "POST"])
def upload_statement():
    active_customers_count, pending_review_count = get_sidebar_metrics()
    if request.method == "POST":
        statement = request.files.get("statement")
        pdf_password = request.form.get("pdf_password", "").strip()
        use_ocr = request.form.get("use_ocr") == "on"
        if not statement or statement.filename == "":
            return render_template("upload_statement.html", error="Please select a file",
                                   active_customers_count=active_customers_count, pending_review_count=pending_review_count)
        try:
            filename = (statement.filename or "").lower()
            if filename.endswith(".pdf"):
                transactions = parse_pdf_statement(statement, password=pdf_password, use_ocr=use_ocr)
            else:
                transactions = parse_statement_file(statement)
            token = save_pending_transactions(transactions)
            # Keep in session (normal flow) and also in query param (works even if cookies are blocked)
            session["pending_token"] = token
            return redirect(url_for("process_payments", token=token))
        except Exception as e:
            return render_template("upload_statement.html", error=f"Parsing failed: {e}",
                                   active_customers_count=active_customers_count, pending_review_count=pending_review_count)
    return render_template("upload_statement.html", active_customers_count=active_customers_count,
                           pending_review_count=pending_review_count)


@app.route("/process_payments")
def process_payments():
    token = request.args.get("token", "").strip()
    transactions = load_pending_transactions(token, delete_after=True) if token else pop_pending_transactions()
    active_customers_count, pending_review_count = get_sidebar_metrics()

    conn = get_db_connection()
    unidentified_count = conn.execute("SELECT COUNT(*) FROM payment WHERE status = 'Unidentified Payment'").fetchone()[0]
    total_paid_count = conn.execute("SELECT COUNT(*) FROM payment WHERE status = 'Paid'").fetchone()[0]
    conn.close()

    recon_stats = {
        'paid_count': total_paid_count,
        'verify_count': pending_review_count,
        'unidentified_count': unidentified_count
    }

    if not transactions:
        return render_template(
            "process_payments.html",
            recon_stats=recon_stats,
            active_customers_count=active_customers_count,
            pending_review_count=pending_review_count
        )
    results = process_transactions(transactions)
    active_customers_count, pending_review_count = get_sidebar_metrics()
    return render_template(
        "process_payments.html",
        results=results,
        recon_stats=recon_stats,
        active_customers_count=active_customers_count,
        pending_review_count=pending_review_count
    )


@app.route("/auto-reconcile-db", methods=["GET", "POST"])
def auto_reconcile_db():
    stats = reconcile_database_payments()
    message = (
        f"Database auto-reconciliation complete! {stats['reconciled_unlinked']} unlinked payment(s) successfully matched to customers. "
        f"Database status: {stats['paid_count']} Paid, {stats['verify_count']} Needs Verification, {stats['unidentified_count']} Unidentified."
    )
    active_customers_count, pending_review_count = get_sidebar_metrics()
    return render_template(
        "process_payments.html",
        message=message,
        recon_stats=stats,
        active_customers_count=active_customers_count,
        pending_review_count=pending_review_count
    )


@app.route("/payment_history")
def payment_history():
    history = get_payment_history(limit=500)
    active_customers_count, pending_review_count = get_sidebar_metrics()
    return render_template(
        "payment_history.html",
        payments=history,
        active_customers_count=active_customers_count,
        pending_review_count=pending_review_count
    )


@app.route("/download-report")
@app.route("/download-payment-report")
def download_report():
    month = request.args.get("month", "").strip() or None
    fmt = request.args.get("format", "xlsx").strip().lower()
    file_prefix = f"Payment_Report_{month}" if month else "Payment_Report_All"

    if fmt == "csv":
        stream = generate_payment_report_csv(month)
        return send_file(
            stream,
            as_attachment=True,
            download_name=f"{file_prefix}.csv",
            mimetype="text/csv"
        )
    else:
        stream = generate_payment_report_excel(month)
        return send_file(
            stream,
            as_attachment=True,
            download_name=f"{file_prefix}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    app.run(debug=True)
