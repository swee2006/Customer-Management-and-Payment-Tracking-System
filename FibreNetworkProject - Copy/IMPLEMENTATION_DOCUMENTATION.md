# Application Implementation Documentation

This document describes the requested implementation areas in the Flask fibre-network customer and payment application. The code is in `app.py`; database initialization is also available in `create_db.py`. The excerpts below follow the current project behavior.

## 1. Flask application initialization and configuration

The project uses Flask with SQLite directly; it does not use Flask-SQLAlchemy or `db.init_app()`.

```python
app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE = os.path.join(BASE_DIR, 'database', 'fiber.db')
CURRENT_PAID_MAP = {}
PENDING_DIR = os.path.join(BASE_DIR, "database", "pending_statements")
os.makedirs(PENDING_DIR, exist_ok=True)
os.makedirs(os.path.dirname(DATABASE), exist_ok=True)

app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "change-me")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "change-me")


def get_db_connection():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn
```

`load_dotenv()` is called when the application starts, so local environment variables can be loaded from a `.env` file. The database file is `database/fiber.db`.

## 2. Database tables

The application creates these tables in `ensure_schema()` (`app.py`) and in the database setup script (`create_db.py`).

```sql
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
);

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
);

CREATE TABLE IF NOT EXISTS unmatched_transactions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    raw_description TEXT,
    amount REAL,
    payment_date TEXT,
    reason TEXT,
    logged_at TEXT DEFAULT CURRENT_TIMESTAMP
);
```

`ensure_schema()` runs when `app.py` is imported. It creates missing tables and migrates certain older customer/payment schemas. `create_db.py` drops and recreates all three tables, so running that script resets the existing database and its records.

## 3. Admin login route

The current login handler is registered at `/` (not `/login`). It compares the submitted username and password against the configured admin values.

```python
@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            return redirect(url_for("dashboard"))
        return "Invalid Username or Password"
    return render_template("login.html")
```

The current handler does not set a logged-in session value, and other routes do not check for one. The `change-me` defaults should be overridden with `ADMIN_USERNAME`, `ADMIN_PASSWORD`, and `SECRET_KEY` environment variables before deployment.

## 4. Customer add, edit, and delete

Database operations are parameterized. `add_customer()` inserts a new customer, `update_customer()` changes an existing row, and `delete_customer()` removes one.

```python
def add_customer(box_id, name, phone, subscription, status,
                 connection_id=None, plan_amount=None, upi_id=None):
    conn = get_db_connection()
    try:
        conn.execute('''
            INSERT INTO customer (
                box_id, name, phone, phone_number, subscription, status, paid_status,
                connection_id, plan_amount, upi_id
            ) VALUES (?, ?, ?, ?, ?, ?, 'unpaid', ?, ?, ?)
        ''', (box_id, name, phone, phone, subscription, status,
              connection_id or box_id, plan_amount or subscription, upi_id))
        conn.commit()
        conn.close()
        return True, "Customer added successfully!"
    except sqlite3.IntegrityError as e:
        conn.close()
        return False, f"Error: {str(e)}"


def update_customer(customer_id, box_id, name, phone, subscription, status,
                    connection_id=None, plan_amount=None, upi_id=None):
    conn = get_db_connection()
    try:
        conn.execute('''
            UPDATE customer
            SET box_id=?, name=?, phone=?, phone_number=?, subscription=?, status=?,
                connection_id=?, plan_amount=?, upi_id=?
            WHERE customer_id=?
        ''', (box_id, name, phone, phone, subscription, status,
              connection_id, plan_amount, upi_id, customer_id))
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
```

The corresponding page routes are `/add-customer`, `/edit-customer/<customer_id>`, and `/delete-customer/<customer_id>`. Add and edit accept GET and POST; delete currently accepts GET and POST.

## 5. Excel customer import

The `/upload-customers` route reads an uploaded spreadsheet with Pandas. The expected columns are `box_id`, `name`, `phone`, `subscription`, and `status`; optional fields include `connection_id`, `plan_amount`, and `upi_id`.

```python
@app.route("/upload-customers", methods=["GET", "POST"])
def upload_customers():
    active_customers_count, pending_review_count = get_sidebar_metrics()
    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            return render_template("upload_customers.html", error="No file selected",
                                   active_customers_count=active_customers_count,
                                   pending_review_count=pending_review_count)
        try:
            df = pd.read_excel(file)
            required = ['box_id', 'name', 'phone', 'subscription', 'status']
            for col in required:
                if col not in df.columns.str.lower() and col not in df.columns:
                    return render_template("upload_customers.html", error=f"Missing column: {col}",
                                           active_customers_count=active_customers_count,
                                           pending_review_count=pending_review_count)

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
                    if any(pd.isna(value) for value in
                           (box_id, name, phone, subscription, status)):
                        raise ValueError("One or more required values missing")
                    success, message = add_customer(
                        str(box_id).strip(), str(name).strip(), str(phone).strip(),
                        float(subscription), str(status).strip().lower(),
                        connection_id=str(connection).strip(),
                        plan_amount=float(plan_amount) if plan_amount else None,
                        upi_id=str(upi_id).strip())
                    if success:
                        added += 1
                    else:
                        errors.append(message)
                except Exception as e:
                    errors.append(f"Row {idx + 2}: {e}")

            message = f"Imported {added} customers."
            if errors:
                message += " Some rows failed: " + "; ".join(errors[:5])
            return render_template("upload_customers.html", success=message,
                                   active_customers_count=active_customers_count,
                                   pending_review_count=pending_review_count)
        except Exception as e:
            return render_template("upload_customers.html", error=f"Error processing file: {e}",
                                   active_customers_count=active_customers_count,
                                   pending_review_count=pending_review_count)
    return render_template("upload_customers.html",
                           active_customers_count=active_customers_count,
                           pending_review_count=pending_review_count)
```

Pandas uses OpenPyXL to read `.xlsx` files in this project. The app also imports `openpyxl` directly to create Excel payment reports.

## 6. Bank statement upload and parsing

The upload route accepts a statement file and optional PDF password/OCR flag. PDF files are routed to the PDF parser; other files are sent to the CSV/TXT parser. Parsed transactions are staged, then the user is redirected to payment processing.

```python
@app.route("/upload_statement", methods=["GET", "POST"])
def upload_statement():
    active_customers_count, pending_review_count = get_sidebar_metrics()
    if request.method == "POST":
        statement = request.files.get("statement")
        pdf_password = request.form.get("pdf_password", "").strip()
        use_ocr = request.form.get("use_ocr") == "on"
        if not statement or statement.filename == "":
            return render_template("upload_statement.html", error="Please select a file",
                                   active_customers_count=active_customers_count,
                                   pending_review_count=pending_review_count)
        try:
            filename = (statement.filename or "").lower()
            if filename.endswith(".pdf"):
                transactions = parse_pdf_statement(
                    statement, password=pdf_password, use_ocr=use_ocr)
            else:
                transactions = parse_statement_file(statement)
            token = save_pending_transactions(transactions)
            session["pending_token"] = token
            return redirect(url_for("process_payments", token=token))
        except Exception as e:
            return render_template("upload_statement.html", error=f"Parsing failed: {e}",
                                   active_customers_count=active_customers_count,
                                   pending_review_count=pending_review_count)
    return render_template("upload_statement.html",
                           active_customers_count=active_customers_count,
                           pending_review_count=pending_review_count)
```

`parse_statement_file()` reads CSV rows with `csv.DictReader`. TXT input is split on comma or tab delimiters, with a free-text fallback that looks for UPI IDs, dates, and plausible payment amounts. PDF text extraction uses `pdfplumber`; encrypted files can use the supplied password. Statement formats vary, so parsing is best-effort and uses customer plan amounts to reject implausible extracted amounts.

## 7. OCR fallback for scanned PDFs

OCR is optional and requires `pdf2image`, `pytesseract`, Poppler, and the Tesseract executable. `pdf2image` converts PDF pages into image objects (Pillow-backed); the app sends each image to Tesseract. There is no direct `PIL` import in `app.py`.

```python
def ocr_pdf_statement(file_storage, password=''):
    if not convert_from_path or not pytesseract:
        raise ValueError("OCR support not installed (need pytesseract + pdf2image).")

    tesseract_cmd = detect_tesseract_cmd()
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    poppler_bin = detect_poppler_bin()
    if not poppler_bin:
        raise ValueError(
            "Poppler not found. Install Poppler or set POPPLER_BIN to the folder containing pdftoppm.exe.")

    tmp_path = os.path.join(PENDING_DIR, f"ocr_{uuid.uuid4().hex}.pdf")
    file_storage.stream.seek(0)
    with open(tmp_path, "wb") as f:
        f.write(file_storage.read())

    try:
        images = convert_from_path(tmp_path, poppler_path=poppler_bin)
        full_text = "".join(pytesseract.image_to_string(img) + "\n" for img in images)
        transactions = parse_statement_text(full_text)
        if not transactions:
            raise ValueError("OCR ran, but still no UPI+amount lines were detected in the PDF.")
        return transactions
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
```

`parse_pdf_statement()` first extracts text using `pdfplumber`. It invokes this OCR function only when normal text extraction produces no transactions and the upload form has OCR enabled.

## 8. Automatic payment reconciliation

UPI IDs are normalized to the last four characters before `@` plus the bank/domain portion. Matching prioritizes UPI suffix plus amount, then exact normalized customer name plus amount. Amounts are considered close when the difference is at most 2 currency units.

```python
def amounts_close(a, b, tolerance=2):
    if a is None or b is None:
        return False
    try:
        return abs(float(a) - float(b)) <= tolerance
    except (TypeError, ValueError):
        return False


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


def match_payment(transaction, customers):
    month = transaction['month']
    name = (transaction.get('name') or '').strip().lower()
    amount = transaction['amount']
    upi_suffix = get_upi_suffix(transaction.get('upi_id'))
    upi_matches = []
    upi_amount_mismatches = []
    name_matches = []

    for customer in customers:
        if not customer or customer.get('customer_id') is None:
            continue
        if CURRENT_PAID_MAP.get((customer['customer_id'], month)):
            continue
        customer_upi = get_upi_suffix(customer.get('upi_id'))
        plan_amount = customer.get('plan_amount') or customer.get('subscription') or 0

        if upi_suffix and customer_upi and upi_suffix == customer_upi:
            if amounts_close(plan_amount, amount):
                upi_matches.append(customer)
            else:
                mismatch = dict(customer)
                mismatch['expected_amount'] = plan_amount
                upi_amount_mismatches.append(mismatch)
            continue

        if customer.get('name', '').lower() == name and amounts_close(plan_amount, amount):
            name_matches.append(customer)

    if len(upi_matches) == 1:
        return {'status': 'Paid', 'customer': upi_matches[0], 'reason': 'UPI + amount match'}
    if len(upi_matches) > 1:
        return {'status': 'Needs Verification', 'candidates': upi_matches,
                'reason': 'Multiple UPI matches'}
    if upi_amount_mismatches:
        return {'status': 'Needs Verification', 'candidates': upi_amount_mismatches,
                'reason': 'UPI matched but amount differs from plan'}
    if len(name_matches) == 1:
        return {'status': 'Paid', 'customer': name_matches[0], 'reason': 'Name + amount match'}
    if len(name_matches) > 1:
        return {'status': 'Needs Verification', 'candidates': name_matches,
                'reason': 'Multiple name/amount matches'}
    return {'status': 'Unidentified Payment', 'reason': 'No customer matched'}
```

The separate `/auto-reconcile-db` action links existing unassigned payment rows to customers by phone plus amount first, then UPI suffix plus amount. Both paths avoid customers already recorded as paid for the same month where applicable.

## 9. Dashboard totals, collection, and realization rate

`calculate_payment_status_for_month(month)` returns the number of customers, distinct customers with a `Paid` payment in that month, unpaid customer count, and the sum of paid payment amounts. The dashboard uses those values and calculates collection potential from active customer plans.

```python
stats = calculate_payment_status_for_month(month)
total_customers = stats['total']
paid_customers = stats['paid']
unpaid_customers = stats['unpaid']
collection = stats['collection']

active_row = conn.execute(
    "SELECT COUNT(*) as cnt, "
    "SUM(COALESCE(plan_amount, subscription)) as total_due "
    "FROM customer WHERE LOWER(status) = 'active'"
).fetchone()
total_potential = round(active_row['total_due'] or 0.0, 2)
active_customers_count = active_row['cnt'] or 0
realization_rate = round(collection / total_potential * 100, 1) if total_potential > 0 else 0
outstanding_dues = max(0.0, round(total_potential - collection, 2))
```

The dashboard route is `/dashboard`; it accepts an optional `month` query parameter and passes the calculated values to `dashboard.html`. Realization rate is the selected month's collection divided by total active-customer plan potential.

## 10. Reports and CSV/Excel downloads

`/reports` displays aggregate payment statistics, the six latest monthly paid totals, recent payment rows, and the count of unmatched transactions. CSV and Excel files are served by `/download-report` and `/download-payment-report`; the `format` query parameter selects CSV (`format=csv`) or Excel (default `.xlsx`), and an optional `month` filters the report.

```python
@app.route("/download-report")
@app.route("/download-payment-report")
def download_report():
    month = request.args.get("month", "").strip() or None
    fmt = request.args.get("format", "xlsx").strip().lower()
    file_prefix = f"Payment_Report_{month}" if month else "Payment_Report_All"

    if fmt == "csv":
        stream = generate_payment_report_csv(month)
        return send_file(stream, as_attachment=True,
                         download_name=f"{file_prefix}.csv",
                         mimetype="text/csv")

    stream = generate_payment_report_excel(month)
    return send_file(
        stream, as_attachment=True,
        download_name=f"{file_prefix}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
```

`generate_payment_report_csv()` builds a payment/customer join and writes columns for payment ID, transaction ID, date, month, customer, box/connection IDs, phone, plan amount, amount paid, UPI ID, and status. `generate_payment_report_excel()` builds an OpenPyXL workbook with payment details, unpaid customers, and summary metrics, then returns the workbook as an in-memory download.
