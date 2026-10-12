"""School Management System - FastAPI + SQLite database.

The live database is a single SQLite file (school.db) in this folder — still
just one file, no server/setup needed, exactly like the old Data.xlsx was.
SQLite is used instead of Excel because it handles thousands of rows and
multiple concurrent users far better (indexed lookups instead of re-reading
the whole spreadsheet on every request, no need to rewrite the entire file on
every save).

Data.xlsx is kept as an *export/backup* format only:
- If you're upgrading from an older version of this app and already have a
  Data.xlsx with real data in it, the very first time this new version runs
  it automatically imports everything from Data.xlsx into school.db, once.
- At any time, use "Export to Excel" (Settings page, or GET /api/export/excel)
  to get a fresh, up-to-date Data.xlsx snapshot for viewing/printing/backup
  in Excel. That exported file is not read back in automatically.
"""
import os, datetime, threading, secrets, hashlib, base64, re, sqlite3, glob, zipfile, json, shutil, tempfile
from typing import Optional, List
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import database as pg_database

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = "school.db"
DB_PATH = os.path.join(BASE_DIR, DB_NAME)
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
LEGACY_XLSX_PATH = os.path.join(BASE_DIR, "Data.xlsx")
PHOTO_DIR = os.path.join(BASE_DIR, "photos")
DUE_SOON_DAYS = 7
LATE_COUNT_START_DAYS = 31
DEFAULT_ADMIN_PIN = "123"
DEFAULT_ADMIN_USERNAME = "Administrator"
DEFAULT_ADMIN_PASSWORD = "123"
DEFAULT_ADMIN_DISPLAY_NAME = "Administrator"
LOCK = threading.RLock()
SESSIONS = {}
GROUP_MANAGEMENT_ROLES = {"developer administrator", "admin", "developer_admin", "developer-admin", "bm"}
DEVELOPER_ADMIN_ROLES = {"admin", "developer administrator", "developer_admin", "developer-admin"}
OPERATIONAL_ROLES = {"teacher", "bm", "receptionist"}
ROLE_OPTIONS = ["Teacher", "BM", "Receptionist"]
DEPARTMENT_OPTIONS = [
    "ចំណេះដឹងទូទៅ ភាសាខ្មែរ",
    "ភាសាអង់គ្លេស ពេញម៉ោង",
    "ភាសាអង់គ្លេស ក្រៅម៉ោង",
    "ភាសាចិន ក្រៅម៉ោង",
]

# Column definitions per table, in the exact order used for Excel import/export.
TABLE_COLUMNS = {
    "Settings": ["Key", "Value"],
    "Classes": ["ClassID", "ClassName", "MonthlyFee", "Department"],
    "Students": ["StudentID", "StudentCode", "Surname", "GivenName", "LatinName", "FullName", "Gender", "DOB", "ParentName", "ParentOccupation", "GuardianPhone", "Address", "ClassID", "EnrollmentDate", "Status", "PhotoPath", "Subject", "RegisterNo", "Room", "Schedule", "Department"],
    "Payments": ["PaymentID", "InvoiceNo", "StudentID", "DurationMonths", "AmountPaid", "PaymentDate", "NextDueDate", "ReceivedBy", "PayerAccount", "TotalAmount", "Discount", "Deposit", "Owed", "PaymentMethod", "CheckNumber"],
    "Attendance": ["AttendanceID", "StudentID", "AttDate", "Status"],
    "Scores": ["ScoreID", "StudentID", "ExamName", "Subject", "Score", "MaxScore", "ExamDate"],
    "Users": ["UserID", "Username", "Password", "PasswordHash", "DisplayName", "Role", "Active"],
    "WarehouseBooks": ["BookID", "BookName", "ClassName", "Stock", "Price", "CreatedAt"],
    "WarehouseUniforms": ["UniformID", "UniformType", "Size", "Stock", "Price", "CreatedAt"],
    "Sales": ["SaleID", "SaleType", "StudentID", "ItemID", "ItemName", "Quantity", "UnitPrice", "TotalAmount", "PaymentDate", "ReceivedBy", "PaymentMethod"],
}
SEED_CLASSES = [(1, "ថ្នាក់ទី១", 25.0), (2, "ថ្នាក់ទី២", 30.0), (3, "ថ្នាក់ទី៣", 35.0)]

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS Settings (
    "Key" TEXT PRIMARY KEY,
    "Value" TEXT
);
CREATE TABLE IF NOT EXISTS Classes (
    ClassID INTEGER PRIMARY KEY,
    ClassName TEXT,
    MonthlyFee REAL,
    Department TEXT
);
CREATE TABLE IF NOT EXISTS Students (
    StudentID INTEGER PRIMARY KEY,
    StudentCode TEXT,
    Surname TEXT,
    GivenName TEXT,
    LatinName TEXT,
    FullName TEXT,
    Gender TEXT,
    DOB TEXT,
    ParentName TEXT,
    ParentOccupation TEXT,
    GuardianPhone TEXT,
    Address TEXT,
    ClassID INTEGER,
    EnrollmentDate TEXT,
    Status TEXT,
    PhotoPath TEXT,
    Subject TEXT,
    RegisterNo TEXT,
    Room TEXT,
    Schedule TEXT,
    Department TEXT,
    FOREIGN KEY (ClassID) REFERENCES Classes(ClassID) ON DELETE SET NULL
);
CREATE TABLE IF NOT EXISTS Payments (
    PaymentID INTEGER PRIMARY KEY,
    InvoiceNo TEXT,
    StudentID INTEGER,
    DurationMonths INTEGER,
    AmountPaid REAL,
    PaymentDate TEXT,
    NextDueDate TEXT,
    ReceivedBy TEXT,
    PayerAccount TEXT,
    TotalAmount REAL,
    Discount REAL,
    Deposit REAL,
    Owed REAL,
    PaymentMethod TEXT,
    CheckNumber TEXT,
    FOREIGN KEY (StudentID) REFERENCES Students(StudentID) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS Attendance (
    AttendanceID INTEGER PRIMARY KEY,
    StudentID INTEGER,
    AttDate TEXT,
    Status TEXT,
    FOREIGN KEY (StudentID) REFERENCES Students(StudentID) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS Scores (
    ScoreID INTEGER PRIMARY KEY,
    StudentID INTEGER NOT NULL,
    ExamName TEXT NOT NULL,
    Subject TEXT NOT NULL,
    Score REAL DEFAULT 0,
    MaxScore REAL DEFAULT 100,
    ExamDate TEXT,
    FOREIGN KEY (StudentID) REFERENCES Students(StudentID) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS Users (
    UserID INTEGER PRIMARY KEY,
    Username TEXT,
    Password TEXT,
    PasswordHash TEXT,
    DisplayName TEXT,
    Role TEXT,
    Active TEXT
);
CREATE INDEX IF NOT EXISTS idx_students_classid ON Students(ClassID);
CREATE INDEX IF NOT EXISTS idx_students_code ON Students(StudentCode);
CREATE INDEX IF NOT EXISTS idx_students_status ON Students(Status);
CREATE INDEX IF NOT EXISTS idx_payments_studentid ON Payments(StudentID);
CREATE INDEX IF NOT EXISTS idx_payments_nextdue ON Payments(NextDueDate);
CREATE INDEX IF NOT EXISTS idx_payments_date ON Payments(PaymentDate);
CREATE INDEX IF NOT EXISTS idx_attendance_studentid ON Attendance(StudentID);
CREATE INDEX IF NOT EXISTS idx_attendance_date ON Attendance(AttDate);
CREATE INDEX IF NOT EXISTS idx_scores_student ON Scores(StudentID);
CREATE INDEX IF NOT EXISTS idx_scores_exam ON Scores(ExamName);
CREATE INDEX IF NOT EXISTS idx_users_username ON Users(Username);
CREATE TABLE IF NOT EXISTS WarehouseBooks (
    BookID INTEGER PRIMARY KEY AUTOINCREMENT,
    BookName TEXT NOT NULL,
    ClassName TEXT DEFAULT '',
    Stock INTEGER NOT NULL DEFAULT 0,
    Price REAL NOT NULL DEFAULT 0,
    CreatedAt TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS WarehouseUniforms (
    UniformID INTEGER PRIMARY KEY AUTOINCREMENT,
    UniformType TEXT NOT NULL,
    Size TEXT DEFAULT '',
    Stock INTEGER NOT NULL DEFAULT 0,
    Price REAL NOT NULL DEFAULT 0,
    CreatedAt TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_books_name ON WarehouseBooks(BookName);
CREATE INDEX IF NOT EXISTS idx_uniform_type ON WarehouseUniforms(UniformType);
CREATE TABLE IF NOT EXISTS Rooms (
    RoomID INTEGER PRIMARY KEY AUTOINCREMENT,
    RoomName TEXT NOT NULL UNIQUE,
    Capacity INTEGER NOT NULL DEFAULT 0,
    CreatedAt TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_rooms_name ON Rooms(RoomName);
CREATE TABLE IF NOT EXISTS Sales (
    SaleID INTEGER PRIMARY KEY AUTOINCREMENT,
    SaleType TEXT NOT NULL,
    StudentID INTEGER,
    ItemID INTEGER,
    ItemName TEXT NOT NULL,
    Quantity INTEGER NOT NULL DEFAULT 1,
    UnitPrice REAL NOT NULL DEFAULT 0,
    TotalAmount REAL NOT NULL DEFAULT 0,
    PaymentDate TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    ReceivedBy TEXT DEFAULT '',
    PaymentMethod TEXT DEFAULT 'Cash',
    FOREIGN KEY (StudentID) REFERENCES Students(StudentID) ON DELETE SET NULL
);
CREATE INDEX IF NOT EXISTS idx_sales_date ON Sales(PaymentDate);
"""

app = FastAPI(title="School Management System API", version="4.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])
os.makedirs(PHOTO_DIR, exist_ok=True)
app.mount("/photos", StaticFiles(directory=PHOTO_DIR), name="photos")


@app.get("/", include_in_schema=False)
def serve_frontend():
    """Serve index.html so other computers on the LAN can just open
    http://<this-computer-ip>:8000/ in a browser instead of copying files."""
    return FileResponse(os.path.join(BASE_DIR, "index.html"))


def _hash_password(password):
    return hashlib.sha256(str(password).encode("utf-8")).hexdigest()


def _connect():
    if DATABASE_URL:
        return pg_database.connect(DATABASE_URL)
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.execute("PRAGMA foreign_keys=ON")
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=30000")
    return conn


def _rows(cursor_result):
    return [dict(r) for r in cursor_result]


def _next_id(conn, table, id_col):
    row = conn.execute(f"SELECT COALESCE(MAX({id_col}),0)+1 FROM {table}").fetchone()
    return int(row[0])


def _iso(v):
    if isinstance(v, datetime.datetime):
        return v.date().isoformat()
    if isinstance(v, datetime.date):
        return v.isoformat()
    return v


def _date(s, default=None):
    if not s:
        return default or datetime.date.today()
    try:
        return datetime.datetime.strptime(str(s), "%Y-%m-%d").date()
    except ValueError:
        raise HTTPException(422, f"ទម្រង់កាលបរិច្ឆេទមិនត្រឹមត្រូវ: {s}")


def add_months(d, months):
    mi = d.month - 1 + months
    y = d.year + mi // 12
    m = mi % 12 + 1
    days = [31, 29 if y % 4 == 0 and (y % 100 != 0 or y % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    return datetime.date(y, m, min(d.day, days[m - 1]))


def _settings_dict(conn):
    return {r["Key"]: r["Value"] for r in conn.execute('SELECT "Key", "Value" FROM Settings') if r["Key"]}


def _set_setting(conn, key, value):
    conn.execute(
        'INSERT INTO Settings ("Key","Value") VALUES (?,?) ON CONFLICT("Key") DO UPDATE SET "Value"=excluded."Value"',
        (key, value),
    )


def _admin_ok(conn, pin):
    row = conn.execute('SELECT "Value" FROM Settings WHERE "Key"=?', ("AdminPIN",)).fetchone()
    return str(pin or "") == str(row["Value"] if row else DEFAULT_ADMIN_PIN)


DEFAULT_SETTINGS = {
    "AdminPIN": DEFAULT_ADMIN_PIN,
    "SchoolName": "សាលារៀន",
    "SchoolAddress": "",
    "SchoolPhone": "",
    "SchoolLogo": "",
    "InvoiceTitle": "PAYMENT RECEIPT",
    "InvoiceNoLabel": "Receipt",
    "InvoiceStudentLabel": "Full Name",
    "InvoiceCodeLabel": "Code",
    "InvoiceDurationLabel": "Course Period",
    "InvoiceDateLabel": "Date",
    "InvoiceDueLabel": "Next due",
    "InvoiceReceiverLabel": "Cashier",
    "InvoiceAccountLabel": "Payer Account",
    "InvoiceAmountLabel": "Total Paid",
    "InvoiceFooter": "សូមអរគុណ!",
}


def _import_legacy_excel(conn, path):
    """One-time import of an older Data.xlsx (openpyxl-based version of this
    app) into the new SQLite database. Only runs when school.db is brand new."""
    try:
        wb = load_workbook(path, data_only=True)
    except Exception:
        return
    for table, headers in TABLE_COLUMNS.items():
        if table not in wb.sheetnames:
            continue
        ws = wb[table]
        existing_headers = [c.value for c in ws[1]] if ws.max_row >= 1 else []
        if not existing_headers:
            continue
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if all(v is None for v in row):
                continue
            rows.append(dict(zip(existing_headers, row)))
        if not rows:
            continue
        cols_present = [h for h in headers if h in existing_headers]
        placeholders = ",".join("?" for _ in cols_present)
        col_list = ",".join(f'"{c}"' for c in cols_present)
        for r in rows:
            values = []
            for c in cols_present:
                v = r.get(c)
                if isinstance(v, (datetime.date, datetime.datetime)):
                    v = _iso(v)
                values.append(v)
            try:
                conn.execute(f'INSERT OR IGNORE INTO {table} ({col_list}) VALUES ({placeholders})', values)
            except sqlite3.Error:
                pass
    wb.close()
    conn.commit()



def seed_test_database(conn, count=50):
    """Create a linked demo dataset for QA/testing. Idempotent via TEST-* codes."""
    existing = conn.execute("SELECT COUNT(*) FROM Students WHERE StudentCode LIKE 'TEST-%'").fetchone()[0]
    if existing >= count:
        return
    # Remove an incomplete previous test set so all relationships remain consistent.
    old_ids = [r[0] for r in conn.execute("SELECT StudentID FROM Students WHERE StudentCode LIKE 'TEST-%'").fetchall()]
    if old_ids:
        marks = ','.join('?' for _ in old_ids)
        for table in ('Attendance','Scores','Payments','Sales'):
            conn.execute(f"DELETE FROM {table} WHERE StudentID IN ({marks})", old_ids)
        conn.execute(f"DELETE FROM Students WHERE StudentID IN ({marks})", old_ids)
    class_ids = [r[0] for r in conn.execute("SELECT ClassID FROM Classes ORDER BY ClassID").fetchall()]
    if not class_ids:
        conn.execute("INSERT INTO Classes(ClassID,ClassName,MonthlyFee,Department) VALUES (1,'Test Class A',20,'Test')")
        class_ids = [1]
    next_student = _next_id(conn, 'Students', 'StudentID')
    today = datetime.date.today()
    students = []
    for i in range(1, count + 1):
        sid = next_student + i - 1
        code = f"TEST-{i:04d}"
        gender = 'Male' if i % 2 else 'Female'
        cls = class_ids[(i - 1) % len(class_ids)]
        students.append((sid, code, 'Test', f'Student {i:02d}', f'Test Student {i:02d}', f'Test Student {i:02d}', gender,
                         f"201{(i % 8) + 2}-{((i % 12) + 1):02d}-{((i * 3) % 27) + 1:02d}",
                         f'Parent {i:02d}', 'Business', f'010000{i:04d}', f'Test Address {i:02d}', cls,
                         (today - datetime.timedelta(days=i)).isoformat(), 'Active', '', 'General', f'REG-{i:04d}',
                         f'R{((i - 1) % 5) + 1}', 'ព្រឹក' if i % 2 else 'ល្ងាច', 'Test'))
    conn.executemany("""INSERT INTO Students
        (StudentID,StudentCode,Surname,GivenName,LatinName,FullName,Gender,DOB,ParentName,ParentOccupation,GuardianPhone,Address,ClassID,EnrollmentDate,Status,PhotoPath,Subject,RegisterNo,Room,Schedule,Department)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", students)

    # 50 books + 50 uniforms for warehouse relationship tests.
    for i in range(1, 51):
        conn.execute("INSERT INTO WarehouseBooks(BookName,ClassName,Stock,Price) VALUES (?,?,?,?)",
                     (f'Test Book {i:02d}', f'Test Class {((i - 1) % len(class_ids)) + 1}', 100 + i, 5 + i))
        conn.execute("INSERT INTO WarehouseUniforms(UniformType,Size,Stock,Price) VALUES (?,?,?,?)",
                     (f'Test Uniform {i:02d}', ['S','M','L','XL'][i % 4], 100 + i, 8 + i))

    book_ids = [r[0] for r in conn.execute("SELECT BookID FROM WarehouseBooks ORDER BY BookID DESC LIMIT 50").fetchall()][::-1]
    uniform_ids = [r[0] for r in conn.execute("SELECT UniformID FROM WarehouseUniforms ORDER BY UniformID DESC LIMIT 50").fetchall()][::-1]
    admin = conn.execute("SELECT DisplayName FROM Users ORDER BY UserID LIMIT 1").fetchone()
    receiver = admin[0] if admin else 'Test Admin'

    for i, st in enumerate(students, 1):
        sid = st[0]
        pay_date = today - datetime.timedelta(days=(i % 30))
        months = [1, 2, 3, 6][(i - 1) % 4]
        monthly = 20 + ((i - 1) % 3) * 5
        gross = monthly * months
        discount = 10 if i % 10 == 0 else 0
        total = round(gross * (1 - discount / 100), 2)
        paid = total
        next_due = add_months(pay_date, months)
        conn.execute("""INSERT INTO Payments(PaymentID,InvoiceNo,StudentID,DurationMonths,AmountPaid,PaymentDate,NextDueDate,ReceivedBy,PayerAccount,TotalAmount,Discount,Deposit,Owed,PaymentMethod,CheckNumber)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                     (100000 + i, f'TEST-INV-{i:04d}', sid, months, paid, pay_date.isoformat(), next_due.isoformat(), receiver,
                      f'TEST-PAYER-{i:04d}', total, discount, 0, 0, 'Cash' if i % 2 else 'Transfer', ''))
        att_status = 'Present' if i % 5 else 'Absent'
        conn.execute("INSERT INTO Attendance(AttendanceID,StudentID,AttDate,Status) VALUES (?,?,?,?)",
                     (200000 + i, sid, today.isoformat(), att_status))
        for j, subject in enumerate(['Khmer','English','Math'], 1):
            conn.execute("INSERT INTO Scores(ScoreID,StudentID,ExamName,Subject,Score,MaxScore,ExamDate) VALUES (?,?,?,?,?,?,?)",
                         (300000 + (i - 1) * 3 + j, sid, 'Test Midterm', subject, 60 + ((i * j) % 41), 100,
                          (today - datetime.timedelta(days=7)).isoformat()))
        # One linked sale per student, alternating books/uniforms.
        if i % 2:
            item_id = book_ids[i - 1]
            row = conn.execute("SELECT BookName,Price FROM WarehouseBooks WHERE BookID=?", (item_id,)).fetchone()
            sale_type, item_name, unit = 'book', row[0], row[1]
        else:
            item_id = uniform_ids[i - 1]
            row = conn.execute("SELECT UniformType,Price FROM WarehouseUniforms WHERE UniformID=?", (item_id,)).fetchone()
            sale_type, item_name, unit = 'uniform', row[0], row[1]
        conn.execute("INSERT INTO Sales(SaleType,StudentID,ItemID,ItemName,Quantity,UnitPrice,TotalAmount,PaymentDate,ReceivedBy,PaymentMethod) VALUES (?,?,?,?,?,?,?,?,?,?)",
                     (sale_type, sid, item_id, item_name, 1, unit, unit, pay_date.isoformat(), receiver, 'Cash'))

    conn.commit()


def repair_database_relationships(conn):
    """Remove orphaned child rows and expose relationship integrity through SQLite FK enforcement."""
    conn.execute('PRAGMA foreign_keys=OFF')
    conn.execute("DELETE FROM Payments WHERE StudentID IS NOT NULL AND StudentID NOT IN (SELECT StudentID FROM Students)")
    conn.execute("DELETE FROM Attendance WHERE StudentID IS NOT NULL AND StudentID NOT IN (SELECT StudentID FROM Students)")
    conn.execute("DELETE FROM Scores WHERE StudentID NOT IN (SELECT StudentID FROM Students)")
    conn.execute("DELETE FROM Sales WHERE StudentID IS NOT NULL AND StudentID NOT IN (SELECT StudentID FROM Students)")
    conn.execute('PRAGMA foreign_keys=ON')
    conn.commit()

def ensure_db():
    with LOCK:
        if DATABASE_URL:
            conn = pg_database.connect(DATABASE_URL)
            pg_schema = SCHEMA_SQL.replace("INTEGER PRIMARY KEY AUTOINCREMENT", "INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY")
            conn.executescript(pg_schema)
            if conn.execute("SELECT COUNT(*) FROM Classes").fetchone()[0] == 0:
                conn.executemany("INSERT INTO Classes (ClassID,ClassName,MonthlyFee) VALUES (?,?,?)", SEED_CLASSES)
            existing_keys = {row[0] for row in conn.execute('SELECT "Key" FROM Settings')}
            for key, value in DEFAULT_SETTINGS.items():
                if key not in existing_keys:
                    conn.execute('INSERT INTO Settings ("Key","Value") VALUES (?,?)', (key, value))
            admin = conn.execute("SELECT UserID FROM Users WHERE lower(Username)=lower(?) LIMIT 1", (DEFAULT_ADMIN_USERNAME,)).fetchone()
            if not admin:
                conn.execute(
                    "INSERT INTO Users (UserID,Username,Password,PasswordHash,DisplayName,Role,Active) VALUES (1,?,?,?,?,?,?)",
                    (DEFAULT_ADMIN_USERNAME, DEFAULT_ADMIN_PASSWORD, _hash_password(DEFAULT_ADMIN_PASSWORD), DEFAULT_ADMIN_DISPLAY_NAME, "Developer Administrator", "Yes"),
                )
            _set_setting(conn, "AdminPIN", DEFAULT_ADMIN_PIN)
            conn.commit()
            repair_database_relationships(conn)
            seed_test_database(conn, 50)
            room_names = [r[0] for r in conn.execute("SELECT DISTINCT trim(Room) FROM Students WHERE trim(COALESCE(Room,''))<>''").fetchall()]
            for room_name in room_names:
                conn.execute("INSERT OR IGNORE INTO Rooms(RoomName,Capacity) VALUES (?,0)", (room_name,))
            conn.commit()
            conn.close()
            return
        is_new = not os.path.exists(DB_PATH)
        conn = sqlite3.connect(DB_PATH)
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute("PRAGMA journal_mode=WAL")
        conn.executescript(SCHEMA_SQL)
        # Migrate: add any columns that a newer version of this app introduced
        # to a database created by an older version (SQLite supports adding
        # columns to existing tables directly).
        for table, headers in TABLE_COLUMNS.items():
            existing = {row[1] for row in conn.execute(f"PRAGMA table_info({table})")}
            for h in headers:
                if h not in existing:
                    conn.execute(f'ALTER TABLE {table} ADD COLUMN "{h}" TEXT')
        if conn.execute("SELECT COUNT(*) FROM Classes").fetchone()[0] == 0:
            conn.executemany("INSERT INTO Classes (ClassID,ClassName,MonthlyFee) VALUES (?,?,?)", SEED_CLASSES)
        existing_keys = {row[0] for row in conn.execute('SELECT "Key" FROM Settings')}
        for key, value in DEFAULT_SETTINGS.items():
            if key not in existing_keys:
                conn.execute('INSERT INTO Settings ("Key","Value") VALUES (?,?)', (key, value))
        # Always guarantee the built-in full-access account.
        # This repairs old databases copied from another PC or older releases.
        admin_row = conn.execute("SELECT UserID FROM Users WHERE lower(Username)=lower(?) LIMIT 1", (DEFAULT_ADMIN_USERNAME,)).fetchone()
        legacy_row = conn.execute("SELECT UserID FROM Users WHERE lower(Username)=lower(?) LIMIT 1", ("admin",)).fetchone()
        row = admin_row or legacy_row
        if row:
            conn.execute(
                "UPDATE Users SET Username=?, Password=?, PasswordHash=?, DisplayName=?, Role='Developer Administrator', Active='Yes' WHERE UserID=?",
                (DEFAULT_ADMIN_USERNAME, DEFAULT_ADMIN_PASSWORD, _hash_password(DEFAULT_ADMIN_PASSWORD), DEFAULT_ADMIN_DISPLAY_NAME, row[0]),
            )
        else:
            conn.execute(
                "INSERT INTO Users (UserID,Username,Password,PasswordHash,DisplayName,Role,Active) VALUES (1,?,?,?,?,?,?)",
                (DEFAULT_ADMIN_USERNAME, DEFAULT_ADMIN_PASSWORD, _hash_password(DEFAULT_ADMIN_PASSWORD), DEFAULT_ADMIN_DISPLAY_NAME, "Developer Administrator", "Yes"),
            )
        _set_setting(conn, "AdminPIN", DEFAULT_ADMIN_PIN)
        conn.commit()
        if is_new and os.path.exists(LEGACY_XLSX_PATH):
            _import_legacy_excel(conn, LEGACY_XLSX_PATH)
        repair_database_relationships(conn)
        seed_test_database(conn, 50)
        # Keep the Room dropdown backed by the Rooms table even for legacy
        # databases that stored room names only in Students.Room.
        room_names = [r[0] for r in conn.execute("SELECT DISTINCT trim(Room) FROM Students WHERE trim(COALESCE(Room,''))<>''").fetchall()]
        for room_name in room_names:
            conn.execute("INSERT OR IGNORE INTO Rooms(RoomName,Capacity) VALUES (?,0)", (room_name,))
        conn.commit()
        conn.close()


def _login_user(username, password):
    with LOCK:
        conn = _connect()
        login_name = str(username or "").strip()
        # Accept the legacy admin name as an alias for the built-in Administrator account.
        lookup_name = DEFAULT_ADMIN_USERNAME if login_name.lower() == "admin" else login_name
        row = conn.execute(
            "SELECT * FROM Users WHERE lower(Username)=lower(?)", (lookup_name,)
        ).fetchone()
        conn.close()
    if not row:
        return None
    user = dict(row)
    if str(user.get("Active") or "Yes").lower() not in {"yes", "true", "1", "active"}:
        return None
    plain = str(user.get("Password") or "")
    hashed = str(user.get("PasswordHash") or "")
    valid = (plain != "" and plain == str(password or "")) or (hashed != "" and hashed == _hash_password(password))
    if not valid:
        return None
    token = secrets.token_urlsafe(32)
    role = user.get("Role") or "Staff"
    if str(role).strip().lower() == "admin":
        role = "Developer Administrator"
    SESSIONS[token] = {"UserID": user.get("UserID"), "Username": user.get("Username"), "DisplayName": user.get("DisplayName") or user.get("Username"), "Role": role}
    return token, SESSIONS[token]


def _current_user(request):
    token = request.headers.get("X-Session-Token", "")
    user = SESSIONS.get(token)
    if not user:
        raise HTTPException(401, "សូម Login មុនពេលប្រើប្រាស់ប្រព័ន្ធ។")
    return user


def _require_role(request, allowed):
    user = _current_user(request)
    role = str(user.get("Role") or "").strip().lower()
    allowed_norm = {str(x).strip().lower() for x in allowed}
    if role not in allowed_norm:
        raise HTTPException(403, "គណនីនេះមិនមានសិទ្ធិសម្រាប់មុខងារនេះទេ។")
    return user


def _require_developer_admin(request):
    user = _current_user(request)
    if str(user.get("Role") or "").strip().lower() not in DEVELOPER_ADMIN_ROLES:
        raise HTTPException(403, "មានតែ Developer Administrator ប៉ុណ្ណោះដែលអាចកែការកំណត់ប្រព័ន្ធនេះបាន។")
    return user


class StudentIn(BaseModel):
    StudentCode: Optional[str] = ""
    Surname: Optional[str] = ""
    GivenName: Optional[str] = ""
    LatinName: Optional[str] = ""
    FullName: Optional[str] = ""
    Gender: str
    DOB: Optional[str] = None
    ParentName: Optional[str] = ""
    ParentOccupation: Optional[str] = ""
    GuardianPhone: Optional[str] = ""
    Address: Optional[str] = ""
    ClassID: int
    EnrollmentDate: Optional[str] = None
    Status: str = "Active"
    PhotoData: Optional[str] = None
    Subject: Optional[str] = ""
    RegisterNo: Optional[str] = ""
    Room: Optional[str] = ""
    Schedule: Optional[str] = ""
    Department: Optional[str] = ""


class PaymentAddOnIn(BaseModel):
    SaleType: str
    ItemID: int
    Quantity: int = Field(default=1, ge=1)

class PaymentIn(BaseModel):
    StudentID: int
    DurationMonths: int
    AmountPaid: float
    TotalAmount: Optional[float] = None
    Discount: Optional[float] = 0
    Deposit: Optional[float] = 0
    PaymentMethod: Optional[str] = "Cash"
    CheckNumber: Optional[str] = ""
    AdditionalItems: List[PaymentAddOnIn] = []


class AttendanceEntry(BaseModel):
    StudentID: int
    Status: str


class AttendanceBatchIn(BaseModel):
    AttDate: Optional[str] = None
    Entries: List[AttendanceEntry]

class ScoreEntry(BaseModel):
    StudentID: int
    ExamName: str
    Subject: str
    Score: float = 0
    MaxScore: float = 100
    ExamDate: Optional[str] = None

class ScoreBatchIn(BaseModel):
    Entries: List[ScoreEntry]


class AdminClassPriceIn(BaseModel):
    AdminPIN: str
    ClassID: int
    MonthlyFee: float


class AdminPinIn(BaseModel):
    AdminPIN: str
    NewPIN: str = Field(..., min_length=4, max_length=20)


class NewUserIn(BaseModel):
    Username: str = Field(..., min_length=2, max_length=50)
    Password: str = Field(..., min_length=3, max_length=100)
    DisplayName: str = Field(..., min_length=1, max_length=100)
    Role: str = "Receptionist"


class UserUpdateIn(BaseModel):
    Password: Optional[str] = Field(None, min_length=3, max_length=100)
    DisplayName: Optional[str] = Field(None, min_length=1, max_length=100)
    Role: Optional[str] = None
    Active: Optional[bool] = None


class ClassCreateIn(BaseModel):
    ClassName: str = Field(..., min_length=1, max_length=100)
    MonthlyFee: float = Field(..., gt=0)
    Department: Optional[str] = ""

class RoomCreateIn(BaseModel):
    RoomName: str = Field(..., min_length=1, max_length=100)
    Capacity: int = Field(default=0, ge=0, le=1000)

class ClassUpdateIn(BaseModel):
    ClassName: Optional[str] = Field(None, min_length=1, max_length=100)
    MonthlyFee: Optional[float] = Field(None, gt=0)
    Department: Optional[str] = None

class StudentUpdateIn(BaseModel):
    Surname: Optional[str] = None
    GivenName: Optional[str] = None
    LatinName: Optional[str] = None
    FullName: Optional[str] = None
    Gender: Optional[str] = None
    DOB: Optional[str] = None
    ParentName: Optional[str] = None
    ParentOccupation: Optional[str] = None
    GuardianPhone: Optional[str] = None
    Address: Optional[str] = None
    ClassID: Optional[int] = None
    EnrollmentDate: Optional[str] = None
    Status: Optional[str] = None
    Subject: Optional[str] = None
    RegisterNo: Optional[str] = None
    Room: Optional[str] = None
    Schedule: Optional[str] = None
    Department: Optional[str] = None
    PhotoData: Optional[str] = None


class SystemSettingsIn(BaseModel):
    SchoolName: Optional[str] = None
    SchoolAddress: Optional[str] = None
    SchoolPhone: Optional[str] = None
    InvoiceTitle: Optional[str] = None
    InvoiceNoLabel: Optional[str] = None
    InvoiceStudentLabel: Optional[str] = None
    InvoiceCodeLabel: Optional[str] = None
    InvoiceDurationLabel: Optional[str] = None
    InvoiceDateLabel: Optional[str] = None
    InvoiceDueLabel: Optional[str] = None
    InvoiceReceiverLabel: Optional[str] = None
    InvoiceAccountLabel: Optional[str] = None
    InvoiceAmountLabel: Optional[str] = None
    InvoiceFooter: Optional[str] = None
    LogoData: Optional[str] = None


VALID_DURATIONS = {1, 2, 3, 6}
VALID_STATUS = {"Present", "Absent", "Permission"}


@app.middleware("http")
async def auth_middleware(request: Request, call_next):
    if request.method == "OPTIONS" or not request.url.path.startswith("/api/") or request.url.path in {"/api/health", "/api/login", "/api/settings"}:
        return await call_next(request)
    token = request.headers.get("X-Session-Token", "")
    if token not in SESSIONS:
        from starlette.responses import JSONResponse
        return JSONResponse({"detail": "សូម Login មុនពេលប្រើប្រាស់ប្រព័ន្ធ។"}, status_code=401)
    return await call_next(request)


@app.on_event("startup")
def startup():
    ensure_db()


@app.get("/api/health")
def health():
    return {"status": "ok", "dbFound": os.path.exists(DB_PATH), "dbPath": DB_PATH, "database": "SQLite"}


class LoginIn(BaseModel):
    Username: str
    Password: str


@app.post("/api/login")
def login(payload: LoginIn):
    result = _login_user(payload.Username, payload.Password)
    if not result:
        raise HTTPException(401, "Username ឬ Password មិនត្រឹមត្រូវ។")
    token, user = result
    return {"token": token, "user": user}


@app.get("/api/me")
def me(request: Request):
    return _current_user(request)


@app.post("/api/logout")
def logout(request: Request):
    token = request.headers.get("X-Session-Token", "")
    SESSIONS.pop(token, None)
    return {"message": "បាន Logout ដោយជោគជ័យ"}


@app.get("/api/dashboard")
def dashboard():
    today = datetime.date.today()
    conn = _connect()
    active = conn.execute("SELECT COUNT(*) FROM Students WHERE Status='Active'").fetchone()[0]
    absent = conn.execute("SELECT COUNT(*) FROM Attendance WHERE AttDate=? AND Status='Absent'", (today.isoformat(),)).fetchone()[0]
    payments = _rows(conn.execute("SELECT StudentID, NextDueDate FROM Payments").fetchall())
    conn.close()
    latest = {}
    for p in payments:
        d = _iso(p.get("NextDueDate"))
        sid = p.get("StudentID")
        if d and (sid not in latest or d > latest[sid]):
            latest[sid] = d
    limit = (today + datetime.timedelta(days=DUE_SOON_DAYS)).isoformat()
    due = sum(1 for d in latest.values() if d <= limit)
    return {"totalStudents": active, "absentToday": absent, "duePayments": due}


@app.get("/api/departments")
def list_departments():
    return [{"DepartmentID": i + 1, "DepartmentName": name} for i, name in enumerate(DEPARTMENT_OPTIONS)]


@app.get("/api/classes")
def list_classes():
    conn = _connect()
    rows = _rows(conn.execute("SELECT * FROM Classes").fetchall())
    conn.close()
    return sorted(rows, key=lambda x: str(x.get("ClassName") or ""))


@app.post("/api/admin/classes", status_code=201)
def create_class(payload: ClassCreateIn, request: Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES)
    name = payload.ClassName.strip()
    with LOCK:
        conn = _connect()
        if conn.execute("SELECT 1 FROM Classes WHERE lower(ClassName)=lower(?)", (name,)).fetchone():
            conn.close()
            raise HTTPException(409, "ថ្នាក់/កំរិតនេះមានរួចហើយ។")
        cid = _next_id(conn, "Classes", "ClassID")
        conn.execute("INSERT INTO Classes (ClassID,ClassName,MonthlyFee,Department) VALUES (?,?,?,?)",
                     (cid, name, payload.MonthlyFee, (payload.Department or "").strip()))
        conn.commit()
        conn.close()
    return {"message":"បានបង្កើតថ្នាក់/កំរិតដោយជោគជ័យ","ClassID":cid}


@app.patch("/api/admin/classes/{class_id}")
def update_class(class_id: int, payload: ClassUpdateIn, request: Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES)
    updates = payload.model_dump(exclude_unset=True)
    if not updates:
        raise HTTPException(400, "គ្មានទិន្នន័យត្រូវកែ។")
    with LOCK:
        conn = _connect()
        row = conn.execute("SELECT * FROM Classes WHERE ClassID=?", (class_id,)).fetchone()
        if not row:
            conn.close()
            raise HTTPException(404, "រកមិនឃើញថ្នាក់/កំរិតនេះទេ។")
        if "ClassName" in updates:
            updates["ClassName"] = updates["ClassName"].strip()
            dup = conn.execute("SELECT 1 FROM Classes WHERE lower(ClassName)=lower(?) AND ClassID<>?",
                               (updates["ClassName"], class_id)).fetchone()
            if dup:
                conn.close()
                raise HTTPException(409, "ឈ្មោះថ្នាក់/កំរិតនេះមានរួចហើយ។")
        sets=[]; params=[]
        for k,v in updates.items():
            sets.append(f'"{k}"=?'); params.append(v)
        params.append(class_id)
        conn.execute(f'UPDATE Classes SET {", ".join(sets)} WHERE ClassID=?', params)
        conn.commit(); conn.close()
    return {"message":"បានកែថ្នាក់/កំរិតដោយជោគជ័យ"}


@app.delete("/api/admin/classes/{class_id}")
def delete_class(class_id: int, request: Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES)
    with LOCK:
        conn=_connect()
        row=conn.execute("SELECT ClassName FROM Classes WHERE ClassID=?",(class_id,)).fetchone()
        if not row:
            conn.close(); raise HTTPException(404,"រកមិនឃើញថ្នាក់/កំរិតនេះទេ។")
        count=conn.execute("SELECT COUNT(*) FROM Students WHERE ClassID=?",(class_id,)).fetchone()[0]
        if count:
            conn.close()
            raise HTTPException(409,f"មិនអាចលុបបានទេ។ មានសិស្ស {count} នាក់នៅក្នុងថ្នាក់នេះ។ សូមផ្លាស់ថ្នាក់សិស្សជាមុន។")
        conn.execute("DELETE FROM Classes WHERE ClassID=?",(class_id,))
        conn.commit(); conn.close()
    return {"message":"បានលុបថ្នាក់/កំរិតដោយជោគជ័យ"}


@app.post("/api/admin/rooms", status_code=201)
def create_room(payload: RoomCreateIn, request: Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES)
    name = payload.RoomName.strip()
    with LOCK:
        conn = _connect()
        if conn.execute("SELECT 1 FROM Rooms WHERE lower(RoomName)=lower(?)", (name,)).fetchone():
            conn.close(); raise HTTPException(409, "បន្ទប់នេះមានរួចហើយ។")
        conn.execute("INSERT INTO Rooms(RoomName,Capacity) VALUES (?,?)", (name, payload.Capacity))
        conn.commit(); conn.close()
    return {"message":"បានបង្កើតបន្ទប់ដោយជោគជ័យ","RoomName":name}

@app.get("/api/admin/rooms/all")
def list_all_rooms(request: Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES)
    conn=_connect(); rows=_rows(conn.execute("SELECT * FROM Rooms ORDER BY RoomName").fetchall()); conn.close()
    return rows

@app.get("/api/rooms/options")
def list_room_options():
    """Return rooms from the Rooms table for student registration/edit dropdowns."""
    conn = _connect()
    rows = _rows(conn.execute("SELECT RoomID, RoomName, Capacity FROM Rooms ORDER BY RoomName").fetchall())
    conn.close()
    return rows


@app.get("/api/admin/rooms")
def list_rooms(request: Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES)
    conn=_connect()
    rows=conn.execute("SELECT Room, COUNT(*) AS StudentCount FROM Students WHERE trim(COALESCE(Room,''))<>'' GROUP BY Room ORDER BY Room").fetchall()
    conn.close()
    return _rows(rows)


@app.patch("/api/admin/rooms/{room_name:path}")
def update_room(room_name: str, payload: dict, request: Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES)
    new_name=str(payload.get("Room") or "").strip()
    old_name=str(room_name or "").strip()
    if not old_name or not new_name:
        raise HTTPException(422,"សូមបញ្ចូលឈ្មោះបន្ទប់។")
    with LOCK:
        conn=_connect()
        conn.execute("UPDATE Students SET Room=? WHERE Room=?",(new_name,old_name))
        conn.commit(); conn.close()
    return {"message":"បានកែបន្ទប់ដោយជោគជ័យ","Room":new_name}


@app.delete("/api/admin/rooms/{room_name:path}")
def delete_room(room_name: str, request: Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES)
    name=str(room_name or "").strip()
    if not name: raise HTTPException(422,"ឈ្មោះបន្ទប់មិនត្រឹមត្រូវ។")
    with LOCK:
        conn=_connect()
        conn.execute("UPDATE Students SET Room='' WHERE Room=?",(name,))
        conn.commit(); conn.close()
    return {"message":"បានលុបបន្ទប់ដោយជោគជ័យ"}

def payment_status_for_student(student_id, payments_by_student, today=None):
    today = today or datetime.date.today()
    mine = payments_by_student.get(student_id) or []
    if not mine:
        return {"LastPaymentDate": None, "NextDueDate": None, "LateDays": 0, "PaymentStatus": "មិនទាន់បង់"}
    latest = max(mine, key=lambda p: _iso(p.get("PaymentDate")) or "")
    due = _date(_iso(latest.get("NextDueDate")))
    late = max(0, (today - due).days)
    return {"LastPaymentDate": _iso(latest.get("PaymentDate")), "NextDueDate": due.isoformat(), "LateDays": late, "PaymentStatus": "ហួសកំណត់" if late > 0 else "បានបង់ / មិនទាន់ហួស"}


@app.get("/api/students")
def list_students(status: Optional[str] = None):
    conn = _connect()
    students = _rows(conn.execute("SELECT * FROM Students").fetchall())
    classes = {r["ClassID"]: r for r in _rows(conn.execute("SELECT * FROM Classes").fetchall())}
    payments = _rows(conn.execute("SELECT StudentID, PaymentDate, NextDueDate FROM Payments WHERE NextDueDate IS NOT NULL").fetchall())
    conn.close()
    payments_by_student = {}
    for p in payments:
        payments_by_student.setdefault(p.get("StudentID"), []).append(p)
    out = []
    for s in students:
        if status and s.get("Status") != status:
            continue
        c = classes.get(s.get("ClassID"), {})
        s["ClassName"] = c.get("ClassName")
        s["MonthlyFee"] = c.get("MonthlyFee")
        for k in ("DOB", "EnrollmentDate"):
            s[k] = _iso(s.get(k))
        s.update(payment_status_for_student(s.get("StudentID"), payments_by_student))
        out.append(s)
    return sorted(out, key=lambda x: str(x.get("FullName") or ""))


@app.get("/api/students/next-code")
def next_student_code():
    conn = _connect()
    nid = _next_id(conn, "Students", "StudentID")
    conn.close()
    return {"StudentID": nid, "StudentCode": f"STU-{nid:04d}"}


@app.post("/api/students", status_code=201)
def create_student(payload: StudentIn):
    dob = _date(payload.DOB, None) if payload.DOB else None
    enrollment = _date(payload.EnrollmentDate)
    surname = (payload.Surname or "").strip()
    given = (payload.GivenName or "").strip()
    latin = (payload.LatinName or "").strip()
    full = (payload.FullName or "").strip() or " ".join(x for x in [surname, given] if x).strip()
    if not full:
        raise HTTPException(422, "សូមបញ្ចូល នាមត្រកូល និង នាមខ្លួន។")
    code = (payload.StudentCode or "").strip()
    with LOCK:
        conn = _connect()
        if not conn.execute("SELECT 1 FROM Classes WHERE ClassID=?", (payload.ClassID,)).fetchone():
            conn.close()
            raise HTTPException(400, "ថ្នាក់រៀនដែលបានជ្រើសរើសមិនមានទេ។")
        if (payload.Room or "").strip() and not conn.execute("SELECT 1 FROM Rooms WHERE RoomName=?", ((payload.Room or "").strip(),)).fetchone():
            conn.close()
            raise HTTPException(400, "បន្ទប់ដែលបានជ្រើសរើសមិនមានក្នុង Database ទេ។")
        new_id = _next_id(conn, "Students", "StudentID")
        if not code:
            code = f"STU-{new_id:04d}"
        if conn.execute("SELECT 1 FROM Students WHERE lower(StudentCode)=lower(?)", (code,)).fetchone():
            conn.close()
            raise HTTPException(409, "លេខសម្គាល់សិស្សនេះមានរួចហើយ។")
        photo_path = ""
        if payload.PhotoData:
            m = re.match(r"^data:(image/(?:jpeg|jpg|png|webp));base64,(.+)$", payload.PhotoData, re.I | re.S)
            if not m:
                conn.close()
                raise HTTPException(422, "រូបថតត្រូវជា JPG, PNG ឬ WEBP។")
            raw = base64.b64decode(m.group(2), validate=True)
            if len(raw) > 3 * 1024 * 1024:
                conn.close()
                raise HTTPException(413, "រូបថតធំពេក (អតិបរមា 3MB)។")
            os.makedirs(PHOTO_DIR, exist_ok=True)
            ext = {"image/jpeg": "jpg", "image/jpg": "jpg", "image/png": "png", "image/webp": "webp"}[m.group(1).lower()]
            filename = f"{code}.{ext}"
            with open(os.path.join(PHOTO_DIR, filename), "wb") as f:
                f.write(raw)
            photo_path = os.path.join("photos", filename).replace("\\", "/")
        conn.execute(
            """INSERT INTO Students (StudentID,StudentCode,Surname,GivenName,LatinName,FullName,Gender,DOB,ParentName,ParentOccupation,GuardianPhone,Address,ClassID,EnrollmentDate,Status,PhotoPath,Subject,RegisterNo,Room,Schedule,Department)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (new_id, code, surname, given, latin, full, payload.Gender, dob.isoformat() if dob else None,
             payload.ParentName or "", payload.ParentOccupation or "", payload.GuardianPhone or "", payload.Address or "",
             payload.ClassID, enrollment.isoformat(), payload.Status, photo_path,
             payload.Subject or "", payload.RegisterNo or "", payload.Room or "", payload.Schedule or "", payload.Department or ""),
        )
        conn.commit()
        conn.close()
    return {"message": "បានចុះឈ្មោះសិស្សដោយជោគជ័យ", "StudentID": new_id, "StudentCode": code, "FullName": full, "PhotoPath": photo_path}


@app.patch("/api/students/{student_id}")
def update_student(student_id: int, payload: StudentUpdateIn, request: Request):
    _current_user(request)
    with LOCK:
        conn = _connect()
        row = conn.execute("SELECT * FROM Students WHERE StudentID=?", (student_id,)).fetchone()
        if not row:
            conn.close()
            raise HTTPException(404, "រកមិនឃើញសិស្សនេះទេ។")
        if payload.ClassID is not None:
            if not conn.execute("SELECT 1 FROM Classes WHERE ClassID=?", (payload.ClassID,)).fetchone():
                conn.close()
                raise HTTPException(400, "ថ្នាក់រៀនដែលបានជ្រើសរើសមិនមានទេ។")
        if payload.Room is not None and str(payload.Room).strip():
            if not conn.execute("SELECT 1 FROM Rooms WHERE RoomName=?", (str(payload.Room).strip(),)).fetchone():
                conn.close()
                raise HTTPException(400, "បន្ទប់ដែលបានជ្រើសរើសមិនមានក្នុង Database ទេ។")
        updates = payload.model_dump(exclude_unset=True)
        # PhotoData is transport-only; store the image as a normal student photo file.
        photo_data = updates.pop("PhotoData", None)
        set_clauses = []
        params = []
        for key, value in updates.items():
            if key in {"DOB", "EnrollmentDate"} and value:
                value = _date(value).isoformat()
            set_clauses.append(f'"{key}"=?')
            params.append(value)
        if photo_data:
            m = re.match(r"^data:(image/(?:jpeg|jpg|png|webp));base64,(.+)$", photo_data, re.I | re.S)
            if not m:
                conn.close()
                raise HTTPException(422, "រូបថតត្រូវជា JPG, PNG ឬ WEBP។")
            try:
                raw = base64.b64decode(m.group(2), validate=True)
            except Exception:
                conn.close()
                raise HTTPException(422, "ទិន្នន័យរូបថតមិនត្រឹមត្រូវ។")
            if len(raw) > 3 * 1024 * 1024:
                conn.close()
                raise HTTPException(413, "រូបថតធំពេក (អតិបរមា 3MB)។")
            os.makedirs(PHOTO_DIR, exist_ok=True)
            ext = {"image/jpeg": "jpg", "image/jpg": "jpg", "image/png": "png", "image/webp": "webp"}[m.group(1).lower()]
            code = row["StudentCode"]
            filename = f"{code}.{ext}"
            with open(os.path.join(PHOTO_DIR, filename), "wb") as f:
                f.write(raw)
            set_clauses.append('"PhotoPath"=?')
            params.append(os.path.join("photos", filename).replace("\\", "/"))

        if payload.Surname is not None or payload.GivenName is not None:
            surname = payload.Surname if payload.Surname is not None else row["Surname"]
            given = payload.GivenName if payload.GivenName is not None else row["GivenName"]
            full = " ".join(x for x in [str(surname or "").strip(), str(given or "").strip()] if x).strip()
            set_clauses.append('"FullName"=?')
            params.append(full)
        if set_clauses:
            params.append(student_id)
            conn.execute(f"UPDATE Students SET {', '.join(set_clauses)} WHERE StudentID=?", params)
            conn.commit()
        conn.close()
    return {"message": "បានកែព័ត៌មានសិស្សដោយជោគជ័យ"}


@app.delete("/api/students/{student_id}")
def delete_student(student_id: int, request: Request):
    _require_developer_admin(request)
    with LOCK:
        conn=_connect()
        row=conn.execute("SELECT PhotoPath FROM Students WHERE StudentID=?",(student_id,)).fetchone()
        if not row:
            conn.close(); raise HTTPException(404,"រកមិនឃើញសិស្សនេះទេ។")
        # Keep every student relationship consistent when a student is deleted.
        conn.execute("DELETE FROM Attendance WHERE StudentID=?",(student_id,))
        conn.execute("DELETE FROM Scores WHERE StudentID=?",(student_id,))
        conn.execute("DELETE FROM Payments WHERE StudentID=?",(student_id,))
        conn.execute("DELETE FROM Sales WHERE StudentID=?",(student_id,))
        conn.execute("DELETE FROM Students WHERE StudentID=?",(student_id,))
        conn.commit(); conn.close()
    photo=row["PhotoPath"] if row else ""
    if photo:
        try:
            path=os.path.join(BASE_DIR, photo.replace("/", os.sep))
            if os.path.isfile(path): os.remove(path)
        except OSError:
            pass
    return {"message":"បានលុបសិស្សដោយជោគជ័យ"}


@app.post("/api/payments", status_code=201)
def create_payment(payload: PaymentIn, request: Request):
    user = _current_user(request)
    if payload.DurationMonths not in VALID_DURATIONS:
        raise HTTPException(422, "រយៈពេលបង់ប្រាក់ត្រូវតែជា ១, ២, ៣ ឬ ៦ ខែ។")
    if payload.AmountPaid <= 0:
        raise HTTPException(422, "ចំនួនទឹកប្រាក់មិនត្រឹមត្រូវ។")
    today = datetime.date.today()
    next_due = today + datetime.timedelta(days=LATE_COUNT_START_DAYS * payload.DurationMonths)
    invoice = f"INV-{datetime.datetime.now().strftime('%Y%m%d%H%M%S%f')}"
    with LOCK:
        conn = _connect()
        try:
            student = conn.execute("SELECT * FROM Students WHERE StudentID=?", (payload.StudentID,)).fetchone()
            if not student:
                raise HTTPException(404, "រកមិនឃើញសិស្សនេះទេ។")
            student = dict(student)
            class_row = conn.execute("SELECT * FROM Classes WHERE ClassID=?", (student.get("ClassID"),)).fetchone()
            class_row = dict(class_row) if class_row else {}
            discount = float(payload.Discount or 0)
            deposit = float(payload.Deposit or 0)
            tuition_total = float(payload.TotalAmount) if payload.TotalAmount is not None else payload.AmountPaid
            if discount < 0 or deposit < 0:
                raise HTTPException(422, "ចំនួន Discount/Deposit មិនត្រឹមត្រូវ។")

            # Validate and price optional books/uniforms before changing anything.
            add_ons = []
            add_on_total = 0.0
            for item in payload.AdditionalItems or []:
                sale_type = item.SaleType.strip().lower()
                if sale_type not in {"book", "uniform"}:
                    raise HTTPException(400, "ប្រភេទទំនិញមិនត្រឹមត្រូវ។")
                table = "WarehouseBooks" if sale_type == "book" else "WarehouseUniforms"
                id_col = "BookID" if sale_type == "book" else "UniformID"
                name_col = "BookName" if sale_type == "book" else "UniformType"
                row = conn.execute(f"SELECT * FROM {table} WHERE {id_col}=?", (item.ItemID,)).fetchone()
                if not row:
                    raise HTTPException(404, "រកមិនឃើញទំនិញដែលបានជ្រើសរើស។")
                if int(row["Stock"] or 0) < int(item.Quantity):
                    raise HTTPException(400, f"ស្តុក {row[name_col]} មិនគ្រប់ចំនួន។")
                unit = float(row["Price"] or 0)
                line_total = round(unit * int(item.Quantity), 2)
                add_on_total += line_total
                add_ons.append({
                    "SaleType": sale_type, "ItemID": item.ItemID, "ItemName": str(row[name_col]),
                    "Size": str(row["Size"] or "") if sale_type == "uniform" else "",
                    "Quantity": int(item.Quantity), "UnitPrice": unit, "TotalAmount": line_total
                })

            # Tuition charge and optional merchandise are one payment transaction.
            tuition_charge = max(0.0, tuition_total - discount - deposit)
            grand_total = round(tuition_charge + add_on_total, 2)
            owed = round(max(0.0, grand_total - payload.AmountPaid), 2)
            payment_method = (payload.PaymentMethod or "Cash").strip() or "Cash"
            check_number = (payload.CheckNumber or "").strip()
            pid = _next_id(conn, "Payments", "PaymentID")
            conn.execute(
                """INSERT INTO Payments (PaymentID,InvoiceNo,StudentID,DurationMonths,AmountPaid,PaymentDate,NextDueDate,ReceivedBy,PayerAccount,TotalAmount,Discount,Deposit,Owed,PaymentMethod,CheckNumber)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (pid, invoice, payload.StudentID, payload.DurationMonths, payload.AmountPaid, today.isoformat(), next_due.isoformat(),
                 user["DisplayName"], user["Username"], tuition_total, discount, deposit, owed, payment_method, check_number),
            )
            for item in add_ons:
                table = "WarehouseBooks" if item["SaleType"] == "book" else "WarehouseUniforms"
                id_col = "BookID" if item["SaleType"] == "book" else "UniformID"
                conn.execute(f"UPDATE {table} SET Stock=Stock-? WHERE {id_col}=?", (item["Quantity"], item["ItemID"]))
                conn.execute(
                    "INSERT INTO Sales(SaleType,StudentID,ItemID,ItemName,Quantity,UnitPrice,TotalAmount,ReceivedBy,PaymentMethod) VALUES(?,?,?,?,?,?,?,?,?)",
                    (item["SaleType"], payload.StudentID, item["ItemID"], item["ItemName"], item["Quantity"], item["UnitPrice"], item["TotalAmount"], user["DisplayName"], payment_method)
                )
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
    return {
        "message": "ការបង់ប្រាក់បានជោគជ័យ", "InvoiceNo": invoice, "PaymentID": pid, "StudentID": payload.StudentID,
        "StudentName": student.get("FullName"), "StudentCode": student.get("StudentCode"), "Gender": student.get("Gender"),
        "Subject": student.get("Subject"), "RegisterNo": student.get("RegisterNo"), "Room": student.get("Room"),
        "Schedule": student.get("Schedule"), "ClassName": class_row.get("ClassName"), "DurationMonths": payload.DurationMonths,
        "AmountPaid": payload.AmountPaid, "TotalAmount": tuition_total, "Discount": discount, "Deposit": deposit, "Owed": owed,
        "PaymentMethod": payment_method, "CheckNumber": check_number, "PaymentDate": today.isoformat(), "NextDueDate": next_due.isoformat(),
        "ReceivedBy": user["DisplayName"], "PayerAccount": user["Username"], "AddOnTotal": round(add_on_total, 2),
        "GrandTotal": grand_total, "AdditionalItems": add_ons,
    }


@app.get("/api/payments")
def list_payments(student_id: Optional[int] = None, student_code: Optional[str] = None):
    conn = _connect()
    payments = _rows(conn.execute("SELECT * FROM Payments").fetchall())
    students = _rows(conn.execute("SELECT * FROM Students").fetchall())
    conn.close()
    smap = {s.get("StudentID"): s for s in students}
    code_map = {str(s.get("StudentCode") or "").strip().lower(): s for s in students}
    if student_code:
        st = code_map.get(student_code.strip().lower())
        student_id = st.get("StudentID") if st else -1
    today = datetime.date.today()
    out = []
    for p in payments:
        if student_id is not None and p.get("StudentID") != student_id:
            continue
        s = smap.get(p.get("StudentID"), {})
        due = _iso(p.get("NextDueDate"))
        late = max(0, (today - _date(due)).days) if due else 0
        out.append({**p, "PaymentDate": _iso(p.get("PaymentDate")), "NextDueDate": due, "StudentCode": s.get("StudentCode"), "FullName": s.get("FullName"), "LateDays": late})
    return sorted(out, key=lambda x: x.get("PaymentDate") or "", reverse=True)


@app.post("/api/attendance", status_code=201)
def attendance(payload: AttendanceBatchIn):
    day = _date(payload.AttDate)
    if any(e.Status not in VALID_STATUS for e in payload.Entries):
        raise HTTPException(422, "ស្ថានភាពវត្តមានមិនត្រឹមត្រូវ។")
    with LOCK:
        conn = _connect()
        for e in payload.Entries:
            conn.execute("DELETE FROM Attendance WHERE StudentID=? AND AttDate=?", (e.StudentID, day.isoformat()))
            new_id = _next_id(conn, "Attendance", "AttendanceID")
            conn.execute("INSERT INTO Attendance (AttendanceID,StudentID,AttDate,Status) VALUES (?,?,?,?)", (new_id, e.StudentID, day.isoformat(), e.Status))
        conn.commit()
        conn.close()
    return {"message": f"បានកត់ត្រាវត្តមានសម្រាប់សិស្សចំនួន {len(payload.Entries)} នាក់", "AttDate": day.isoformat()}


@app.get("/api/attendance")
def get_attendance(att_date: Optional[str] = None, date_from: Optional[str] = None, date_to: Optional[str] = None, class_id: Optional[int] = None, status: Optional[str] = None, search: Optional[str] = None):
    day = _date(att_date) if att_date else None
    start = _date(date_from) if date_from else day
    end = _date(date_to) if date_to else day
    conn = _connect()
    if start or end:
        # Filter by date range in SQL directly — avoids pulling the entire
        # (potentially very large, ever-growing) attendance history into
        # memory when a date filter is available, which is the normal case.
        sql = "SELECT * FROM Attendance WHERE 1=1"
        params = []
        if start:
            sql += " AND AttDate>=?"
            params.append(start.isoformat())
        if end:
            sql += " AND AttDate<=?"
            params.append(end.isoformat())
        rows = _rows(conn.execute(sql, params).fetchall())
    else:
        # No date filter at all: cap to a recent window instead of loading
        # the entire history, so this endpoint stays fast as records grow
        # into the tens/hundreds of thousands over time.
        cutoff = (datetime.date.today() - datetime.timedelta(days=90)).isoformat()
        rows = _rows(conn.execute("SELECT * FROM Attendance WHERE AttDate>=?", (cutoff,)).fetchall())
    students = _rows(conn.execute("SELECT * FROM Students").fetchall())
    classes = _rows(conn.execute("SELECT * FROM Classes").fetchall())
    conn.close()
    smap = {s["StudentID"]: s for s in students}
    cmap = {c["ClassID"]: c.get("ClassName") for c in classes}
    q = (search or "").strip().lower()
    out = []
    for a in rows:
        ad = _iso(a.get("AttDate"))
        s = smap.get(a.get("StudentID"), {})
        if class_id is not None and s.get("ClassID") != class_id:
            continue
        if status and a.get("Status") != status:
            continue
        if q and q not in str(s.get("StudentID", "")).lower() and q not in str(s.get("StudentCode", "")).lower() and q not in str(s.get("FullName", "")).lower():
            continue
        out.append({**a, "FullName": s.get("FullName"), "StudentCode": s.get("StudentCode"), "ClassName": cmap.get(s.get("ClassID")), "AttDate": ad})
    return sorted(out, key=lambda x: (x.get("AttDate") or "", x.get("FullName") or ""), reverse=True)


@app.get("/api/scores")
def list_scores(exam_name: str = ""):
    conn = _connect()
    rows = _rows(conn.execute("SELECT * FROM Scores ORDER BY StudentID, Subject" if not exam_name else "SELECT * FROM Scores WHERE ExamName=? ORDER BY StudentID, Subject", (() if not exam_name else (exam_name,))).fetchall())
    conn.close()
    return rows

@app.post("/api/scores", status_code=201)
def save_scores(payload: ScoreBatchIn, request: Request):
    _current_user(request)
    if not payload.Entries:
        return {"saved": 0}
    conn = _connect()
    try:
        saved = 0
        for e in payload.Entries:
            exam = str(e.ExamName or "").strip()
            subject = str(e.Subject or "").strip()
            if not exam or not subject:
                continue
            if e.MaxScore <= 0 or e.Score < 0 or e.Score > e.MaxScore:
                raise HTTPException(400, f"ពិន្ទុ {subject} មិនត្រឹមត្រូវ។")
            row = conn.execute("SELECT ScoreID FROM Scores WHERE StudentID=? AND ExamName=? AND Subject=?", (e.StudentID, exam, subject)).fetchone()
            if row:
                conn.execute("UPDATE Scores SET Score=?,MaxScore=?,ExamDate=? WHERE ScoreID=?", (e.Score,e.MaxScore,e.ExamDate,row[0]))
            else:
                sid = _next_id(conn, "Scores", "ScoreID")
                conn.execute("INSERT INTO Scores (ScoreID,StudentID,ExamName,Subject,Score,MaxScore,ExamDate) VALUES (?,?,?,?,?,?,?)", (sid,e.StudentID,exam,subject,e.Score,e.MaxScore,e.ExamDate))
            saved += 1
        conn.commit()
        return {"saved": saved}
    finally:
        conn.close()

@app.delete("/api/scores/{score_id}")
def delete_score(score_id: int, request: Request):
    _current_user(request)
    conn = _connect(); conn.execute("DELETE FROM Scores WHERE ScoreID=?", (score_id,)); conn.commit(); conn.close()
    return {"message":"បានលុបពិន្ទុ"}

@app.get("/api/reports/due-payments")
def due_report(class_id: Optional[int] = None, min_late: Optional[int] = None, search: Optional[str] = None):
    today = datetime.date.today()
    limit = today + datetime.timedelta(days=DUE_SOON_DAYS)
    conn = _connect()
    payments = _rows(conn.execute("SELECT StudentID, NextDueDate FROM Payments").fetchall())
    students = _rows(conn.execute("SELECT * FROM Students").fetchall())
    classes = _rows(conn.execute("SELECT * FROM Classes").fetchall())
    conn.close()
    cmap = {c["ClassID"]: c.get("ClassName") for c in classes}
    smap = {s["StudentID"]: s for s in students}
    latest = {}
    for p in payments:
        d = _iso(p.get("NextDueDate"))
        sid = p.get("StudentID")
        if d and (sid not in latest or d > latest[sid]):
            latest[sid] = d
    q = (search or "").strip().lower()
    out = []
    for sid, d in latest.items():
        if d > limit.isoformat():
            continue
        s = smap.get(sid, {})
        if class_id is not None and s.get("ClassID") != class_id:
            continue
        late = max(0, (today - _date(d)).days)
        if min_late is not None and late < min_late:
            continue
        if q and q not in str(s.get("StudentID", "")).lower() and q not in str(s.get("StudentCode", "")).lower() and q not in str(s.get("FullName", "")).lower():
            continue
        out.append({"StudentID": sid, "StudentCode": s.get("StudentCode"), "FullName": s.get("FullName"), "GuardianPhone": s.get("GuardianPhone"), "ClassName": cmap.get(s.get("ClassID")), "LastDue": d, "LateDays": late, "IsOverdue": late > 0})
    return sorted(out, key=lambda x: (-x["LateDays"], x["LastDue"]))


@app.get("/api/reports/late-payments")
def late_payment_report(class_id: Optional[int] = None, min_late: int = 1, search: Optional[str] = None):
    return due_report(class_id, min_late, search)


@app.get("/api/reports/attendance")
def attendance_report(date_from: Optional[str] = None, date_to: Optional[str] = None, class_id: Optional[int] = None, status: Optional[str] = None, search: Optional[str] = None):
    return get_attendance(date_from=date_from, date_to=date_to, class_id=class_id, status=status, search=search)



@app.get("/api/reports/export/excel")
def export_report_excel(
    report_type: str = "late",
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    class_id: Optional[int] = None,
    status: Optional[str] = None,
    min_late: int = 1,
    search: Optional[str] = None,
    request: Request = None,
):
    _current_user(request)
    report_type = (report_type or "late").strip().lower()

    if report_type == "late":
        rows = due_report(class_id=class_id, min_late=min_late, search=search)
        headers = ["លេខសិស្ស", "ឈ្មោះ", "ថ្នាក់", "ថ្ងៃកំណត់", "យឺត (ថ្ងៃ)"]
        values = [[r.get("StudentCode") or "", r.get("FullName") or "", r.get("ClassName") or "",
                   r.get("LastDue") or "", r.get("LateDays", 0)] for r in rows]
        filename = f"Late_Payment_Report_{datetime.date.today().isoformat()}.xlsx"
        title = "របាយការណ៍សិស្សបង់ប្រាក់យឺត"
    elif report_type == "attendance":
        rows = attendance_report(date_from=date_from, date_to=date_to, class_id=class_id,
                                 status=status, search=search)
        headers = ["កាលបរិច្ឆេទ", "លេខសិស្ស", "ឈ្មោះ", "ថ្នាក់", "ស្ថានភាព"]
        status_map = {"Absent": "អវត្តមាន", "Permission": "សុំច្បាប់", "Present": "មកសិក្សា"}
        values = [[r.get("AttDate") or "", r.get("StudentCode") or "", r.get("FullName") or "",
                   r.get("ClassName") or "", status_map.get(r.get("Status"), r.get("Status") or "")]
                  for r in rows]
        filename = f"Attendance_Report_{datetime.date.today().isoformat()}.xlsx"
        title = "របាយការណ៍វត្តមាន / អវត្តមាន"
    else:
        raise HTTPException(400, "ប្រភេទរបាយការណ៍មិនត្រឹមត្រូវ")

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ws.append([title])
    ws.append(["Exported", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws.append([])
    ws.append(headers)
    for row in values:
        ws.append(row)
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(headers))}{max(4, ws.max_row)}"
    for cell in ws[4]:
        cell.font = cell.font.copy(bold=True)
    for col in range(1, len(headers) + 1):
        max_len = max([len(str(ws.cell(r, col).value or "")) for r in range(1, ws.max_row + 1)] + [10])
        ws.column_dimensions[get_column_letter(col)].width = min(max_len + 2, 35)

    path = os.path.join(tempfile.gettempdir(), filename)
    wb.save(path)
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )

@app.get("/api/admin/users")
def list_users(request: Request):
    _require_developer_admin(request)
    conn = _connect()
    users = _rows(conn.execute("SELECT * FROM Users").fetchall())
    conn.close()
    return [{"UserID": u.get("UserID"), "Username": u.get("Username"), "DisplayName": u.get("DisplayName"), "Role": u.get("Role"), "Active": str(u.get("Active") or "Yes").lower() in {"yes", "true", "1", "active"}} for u in users]


@app.post("/api/admin/users", status_code=201)
def create_user(payload: NewUserIn, request: Request):
    _require_developer_admin(request)
    username = payload.Username.strip()
    display_name = payload.DisplayName.strip()
    role = payload.Role.strip() or "Receptionist"
    if not username or len(username) < 2:
        raise HTTPException(422, "សូមបញ្ចូល Username យ៉ាងតិច 2 តួអក្សរ។")
    if not display_name:
        raise HTTPException(422, "សូមបញ្ចូលឈ្មោះអ្នកប្រើ។")
    if role.lower() not in {r.lower() for r in ROLE_OPTIONS}:
        raise HTTPException(422, "Role ត្រូវជា Teacher, BM ឬ Receptionist។")
    with LOCK:
        conn = _connect()
        if conn.execute("SELECT 1 FROM Users WHERE lower(Username)=lower(?)", (username,)).fetchone():
            conn.close()
            raise HTTPException(409, "Username នេះមានរួចហើយ។ សូមជ្រើស Username ថ្មី។")
        uid = _next_id(conn, "Users", "UserID")
        try:
            conn.execute(
                "INSERT INTO Users (UserID,Username,Password,PasswordHash,DisplayName,Role,Active) VALUES (?,?,?,?,?,?,?)",
                (uid, username, payload.Password, _hash_password(payload.Password), display_name, role, "Yes"),
            )
            conn.commit()
        except Exception as exc:
            conn.close()
            raise HTTPException(500, f"មិនអាចបង្កើតគណនីបានទេ។ ({exc})")
        conn.close()
    return {"message": f"បានបង្កើតគណនី {role} ដោយជោគជ័យ", "UserID": uid, "Username": username, "DisplayName": display_name, "Role": role}


@app.patch("/api/admin/users/{user_id}")
def update_user(user_id: int, payload: UserUpdateIn, request: Request):
    _require_developer_admin(request)
    updates = payload.model_dump(exclude_unset=True)
    if "Role" in updates and updates["Role"] and updates["Role"].lower() not in {r.lower() for r in ROLE_OPTIONS} and updates["Role"].lower() not in DEVELOPER_ADMIN_ROLES:
        raise HTTPException(422, "Role មិនត្រឹមត្រូវ។")
    with LOCK:
        conn = _connect()
        row = conn.execute("SELECT * FROM Users WHERE UserID=?", (user_id,)).fetchone()
        if not row:
            conn.close()
            raise HTTPException(404, "រកមិនឃើញគណនីនេះទេ។")
        current_username = str(row["Username"] or "").strip().lower()
        if "Active" in updates and not updates["Active"] and current_username == "admin":
            conn.close()
            raise HTTPException(400, "មិនអាចបិទគណនី Developer Administrator admin បានទេ។")
        set_clauses = []
        params = []
        if "DisplayName" in updates:
            set_clauses.append("DisplayName=?")
            params.append(updates["DisplayName"].strip())
        if "Role" in updates and updates["Role"]:
            role = updates["Role"].strip()
            if role.lower() == "admin":
                role = "Developer Administrator"
            set_clauses.append("Role=?")
            params.append(role)
        if "Active" in updates:
            set_clauses.append("Active=?")
            params.append("Yes" if updates["Active"] else "No")
        if "Password" in updates and updates["Password"]:
            set_clauses.append("Password=?")
            params.append(updates["Password"])
            set_clauses.append("PasswordHash=?")
            params.append(_hash_password(updates["Password"]))
        if set_clauses:
            params.append(user_id)
            conn.execute(f"UPDATE Users SET {', '.join(set_clauses)} WHERE UserID=?", params)
            conn.commit()
        conn.close()
    return {"message": "បានកែគណនីដោយជោគជ័យ"}


@app.delete("/api/admin/users/{user_id}")
def delete_user(user_id: int, request: Request):
    current=_require_developer_admin(request)
    if int(current.get("UserID") or 0)==user_id:
        raise HTTPException(400,"មិនអាចលុបគណនីដែលកំពុង Login បានទេ។")
    with LOCK:
        conn=_connect()
        row=conn.execute("SELECT Username FROM Users WHERE UserID=?",(user_id,)).fetchone()
        if not row:
            conn.close(); raise HTTPException(404,"រកមិនឃើញគណនីនេះទេ។")
        if str(row["Username"]).lower()=="admin":
            conn.close(); raise HTTPException(400,"មិនអាចលុប Developer Administrator បានទេ។")
        conn.execute("DELETE FROM Users WHERE UserID=?",(user_id,))
        conn.commit(); conn.close()
    return {"message":"បានលុបគណនី/Teacher ដោយជោគជ័យ"}


class WarehouseBookIn(BaseModel):
    BookName: str = Field(min_length=1, max_length=200)
    ClassName: str = ""
    Stock: int = Field(default=0, ge=0)
    Price: float = Field(default=0, ge=0)

class WarehouseUniformIn(BaseModel):
    UniformType: str = Field(min_length=1, max_length=200)
    Size: str = ""
    Stock: int = Field(default=0, ge=0)
    Price: float = Field(default=0, ge=0)

class SaleIn(BaseModel):
    SaleType: str
    StudentID: Optional[int] = None
    ItemID: int
    Quantity: int = Field(default=1, ge=1)
    PaymentMethod: str = "Cash"

@app.post("/api/warehouse/sales", status_code=201)
def create_warehouse_sale(payload: SaleIn, request: Request):
    user = _require_role(request, GROUP_MANAGEMENT_ROLES | OPERATIONAL_ROLES)
    sale_type = payload.SaleType.strip().lower()
    if sale_type not in {"book", "uniform"}:
        raise HTTPException(400, "ប្រភេទទិញមិនត្រឹមត្រូវ")
    table = "WarehouseBooks" if sale_type == "book" else "WarehouseUniforms"
    id_col = "BookID" if sale_type == "book" else "UniformID"
    name_col = "BookName" if sale_type == "book" else "UniformType"
    with LOCK:
        conn=_connect()
        row=conn.execute(f"SELECT * FROM {table} WHERE {id_col}=?",(payload.ItemID,)).fetchone()
        if not row:
            conn.close(); raise HTTPException(404,"រកមិនឃើញទំនិញ")
        if int(row["Stock"] or 0) < int(payload.Quantity):
            conn.close(); raise HTTPException(400,"ស្តុកមិនគ្រប់ចំនួន")
        unit=float(row["Price"] or 0); total=unit*int(payload.Quantity)
        conn.execute(f"UPDATE {table} SET Stock=Stock-? WHERE {id_col}=?",(int(payload.Quantity),payload.ItemID))
        cur=conn.execute("INSERT INTO Sales(SaleType,StudentID,ItemID,ItemName,Quantity,UnitPrice,TotalAmount,ReceivedBy,PaymentMethod) VALUES(?,?,?,?,?,?,?,?,?)",(sale_type,payload.StudentID,payload.ItemID,str(row[name_col]),int(payload.Quantity),unit,total,user.get("DisplayName") or user.get("Username") or "",payload.PaymentMethod))
        conn.commit(); out=conn.execute("SELECT * FROM Sales WHERE SaleID=?",(cur.lastrowid,)).fetchone(); conn.close()
    return dict(out)

@app.get("/api/warehouse/sales")
def list_warehouse_sales(request: Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES | OPERATIONAL_ROLES)
    conn=_connect(); rows=_rows(conn.execute("SELECT * FROM Sales ORDER BY SaleID DESC").fetchall()); conn.close(); return rows

@app.get("/api/warehouse/books")
def warehouse_books(search: Optional[str] = None):
    conn=_connect()
    q=(search or "").strip().lower()
    rows=_rows(conn.execute("SELECT * FROM WarehouseBooks ORDER BY BookID DESC").fetchall())
    conn.close()
    if q:
        rows=[r for r in rows if q in str(r.get("BookName") or "").lower() or q in str(r.get("ClassName") or "").lower()]
    return rows

@app.post("/api/warehouse/books", status_code=201)
def create_warehouse_book(payload: WarehouseBookIn, request: Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES | OPERATIONAL_ROLES)
    with LOCK:
        conn=_connect()
        cur=conn.execute("INSERT INTO WarehouseBooks(BookName,ClassName,Stock,Price) VALUES(?,?,?,?)",(payload.BookName.strip(),payload.ClassName.strip(),int(payload.Stock),float(payload.Price)))
        conn.commit(); row=conn.execute("SELECT * FROM WarehouseBooks WHERE BookID=?",(cur.lastrowid,)).fetchone(); conn.close()
    return dict(row)

@app.patch("/api/warehouse/books/{book_id}")
def update_warehouse_book(book_id:int,payload:WarehouseBookIn,request:Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES | OPERATIONAL_ROLES)
    with LOCK:
        conn=_connect(); cur=conn.execute("UPDATE WarehouseBooks SET BookName=?,ClassName=?,Stock=?,Price=? WHERE BookID=?",(payload.BookName.strip(),payload.ClassName.strip(),int(payload.Stock),float(payload.Price),book_id)); conn.commit()
        if cur.rowcount==0: conn.close(); raise HTTPException(404,"រកមិនឃើញសៀវភៅ")
        row=conn.execute("SELECT * FROM WarehouseBooks WHERE BookID=?",(book_id,)).fetchone(); conn.close()
    return dict(row)

@app.delete("/api/warehouse/books/{book_id}")
def delete_warehouse_book(book_id:int,request:Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES | OPERATIONAL_ROLES)
    with LOCK:
        conn=_connect(); cur=conn.execute("DELETE FROM WarehouseBooks WHERE BookID=?",(book_id,)); conn.commit(); conn.close()
    if cur.rowcount==0: raise HTTPException(404,"រកមិនឃើញសៀវភៅ")
    return {"message":"បានលុបសៀវភៅដោយជោគជ័យ"}

@app.get("/api/warehouse/uniforms")
def warehouse_uniforms(search: Optional[str] = None):
    conn=_connect(); q=(search or "").strip().lower(); rows=_rows(conn.execute("SELECT * FROM WarehouseUniforms ORDER BY UniformID DESC").fetchall()); conn.close()
    if q: rows=[r for r in rows if q in str(r.get("UniformType") or "").lower() or q in str(r.get("Size") or "").lower()]
    return rows

@app.post("/api/warehouse/uniforms", status_code=201)
def create_warehouse_uniform(payload: WarehouseUniformIn, request: Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES | OPERATIONAL_ROLES)
    uniform_type = (payload.UniformType or "").strip()
    size = (payload.Size or "").strip()
    if not uniform_type:
        raise HTTPException(400, "សូមបញ្ចូលប្រភេទឯកសណ្ឋាន")
    if not size:
        raise HTTPException(400, "សូមបញ្ចូលទំហំឯកសណ្ឋាន")
    with LOCK:
        conn=_connect(); cur=conn.execute("INSERT INTO WarehouseUniforms(UniformType,Size,Stock,Price) VALUES(?,?,?,?)",(uniform_type,size,int(payload.Stock),float(payload.Price))); conn.commit(); row=conn.execute("SELECT * FROM WarehouseUniforms WHERE UniformID=?",(cur.lastrowid,)).fetchone(); conn.close()
    return dict(row)

@app.patch("/api/warehouse/uniforms/{uniform_id}")
def update_warehouse_uniform(uniform_id:int,payload:WarehouseUniformIn,request:Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES | OPERATIONAL_ROLES)
    with LOCK:
        conn=_connect(); cur=conn.execute("UPDATE WarehouseUniforms SET UniformType=?,Size=?,Stock=?,Price=? WHERE UniformID=?",(payload.UniformType.strip(),payload.Size.strip(),int(payload.Stock),float(payload.Price),uniform_id)); conn.commit()
        if cur.rowcount==0: conn.close(); raise HTTPException(404,"រកមិនឃើញឯកសណ្ឋាន")
        row=conn.execute("SELECT * FROM WarehouseUniforms WHERE UniformID=?",(uniform_id,)).fetchone(); conn.close()
    return dict(row)

@app.delete("/api/warehouse/uniforms/{uniform_id}")
def delete_warehouse_uniform(uniform_id:int,request:Request):
    _require_role(request, GROUP_MANAGEMENT_ROLES | OPERATIONAL_ROLES)
    with LOCK:
        conn=_connect(); cur=conn.execute("DELETE FROM WarehouseUniforms WHERE UniformID=?",(uniform_id,)); conn.commit(); conn.close()
    if cur.rowcount==0: raise HTTPException(404,"រកមិនឃើញឯកសណ្ឋាន")
    return {"message":"បានលុបឯកសណ្ឋានដោយជោគជ័យ"}

@app.get("/api/accounting/prices")
def get_accounting_prices():
    conn = _connect()
    settings = _settings_dict(conn)
    conn.close()
    return {
        "UniformPrice": float(settings.get("AccountingUniformPrice") or 0),
        "BookPrice": float(settings.get("AccountingBookPrice") or 0),
        "StudyPrice": float(settings.get("AccountingStudyPrice") or 0),
        "TransportPrice": float(settings.get("AccountingTransportPrice") or 0),
    }


class AccountingPricesIn(BaseModel):
    UniformPrice: float = 0
    BookPrice: float = 0
    StudyPrice: float = 0
    TransportPrice: float = 0


@app.post("/api/accounting/prices")
def save_accounting_prices(payload: AccountingPricesIn, request: Request):
    _require_developer_admin(request)
    vals = {
        "AccountingUniformPrice": max(0, float(payload.UniformPrice or 0)),
        "AccountingBookPrice": max(0, float(payload.BookPrice or 0)),
        "AccountingStudyPrice": max(0, float(payload.StudyPrice or 0)),
        "AccountingTransportPrice": max(0, float(payload.TransportPrice or 0)),
    }
    with LOCK:
        conn = _connect()
        for k, v in vals.items():
            conn.execute('INSERT INTO Settings("Key","Value") VALUES(?,?) ON CONFLICT("Key") DO UPDATE SET "Value"=excluded."Value"', (k, str(v)))
        conn.commit(); conn.close()
    return {"message": "បានរក្សាទុកតម្លៃដោយជោគជ័យ", **{k.replace("Accounting", ""): v for k,v in vals.items()}}


@app.get("/api/accounting/daily-close")
def accounting_daily_close(date_from: Optional[str] = None, date_to: Optional[str] = None, class_id: Optional[int] = None, search: Optional[str] = None):
    today = datetime.date.today().isoformat()
    date_from = date_from or today
    date_to = date_to or date_from
    conn = _connect()
    rows = _rows(conn.execute("""
        SELECT p.*, s.StudentCode, s.FullName, s.ClassID, c.ClassName
        FROM Payments p
        LEFT JOIN Students s ON s.StudentID=p.StudentID
        LEFT JOIN Classes c ON c.ClassID=s.ClassID
        WHERE date(p.PaymentDate) BETWEEN date(?) AND date(?)
        ORDER BY p.PaymentDate DESC, p.PaymentID DESC
    """, (date_from, date_to)).fetchall())
    new_students = conn.execute("SELECT COUNT(*) FROM Students WHERE date(EnrollmentDate) BETWEEN date(?) AND date(?)", (date_from, date_to)).fetchone()[0]
    conn.close()
    q = (search or '').strip().lower()
    if class_id is not None:
        rows = [r for r in rows if r.get('ClassID') == class_id]
    if q:
        rows = [r for r in rows if q in str(r.get('StudentCode') or '').lower() or q in str(r.get('FullName') or '').lower() or q in str(r.get('InvoiceNo') or '').lower()]
    paid = sum(float(r.get('AmountPaid') or 0) for r in rows)
    owed = sum(float(r.get('Owed') or 0) for r in rows)
    return {"date_from": date_from, "date_to": date_to, "new_students": int(new_students), "payment_students": len({r.get('StudentID') for r in rows}), "payment_count": len(rows), "total_paid": paid, "total_owed": owed, "rows": rows}


@app.get("/api/settings")
def get_settings():
    conn = _connect()
    data = _settings_dict(conn)
    conn.close()
    data.pop("AdminPIN", None)
    return data


@app.post("/api/admin/settings")
def save_settings(payload: SystemSettingsIn, request: Request):
    _require_developer_admin(request)
    updates = payload.model_dump(exclude_unset=True)
    logo_data = updates.pop("LogoData", None)
    with LOCK:
        conn = _connect()
        if logo_data:
            m = re.match(r"^data:(image/(?:jpeg|jpg|png|webp));base64,(.+)$", logo_data, re.I | re.S)
            if not m:
                conn.close()
                raise HTTPException(422, "Logo ត្រូវជា JPG, PNG ឬ WEBP។")
            try:
                raw = base64.b64decode(m.group(2), validate=True)
            except Exception:
                conn.close()
                raise HTTPException(422, "Logo មិនត្រឹមត្រូវ។")
            if len(raw) > 2 * 1024 * 1024:
                conn.close()
                raise HTTPException(413, "Logo ធំពេក (អតិបរមា 2MB)។")
            ext = {"image/jpeg": "jpg", "image/jpg": "jpg", "image/png": "png", "image/webp": "webp"}[m.group(1).lower()]
            os.makedirs(PHOTO_DIR, exist_ok=True)
            filename = f"school_logo.{ext}"
            with open(os.path.join(PHOTO_DIR, filename), "wb") as f:
                f.write(raw)
            _set_setting(conn, "SchoolLogo", f"photos/{filename}")
        for key, value in updates.items():
            if value is not None:
                _set_setting(conn, key, value)
        conn.commit()
        conn.close()
    return {"message": "បានរក្សាទុកការកំណត់សាលា និង Invoice Template ដោយជោគជ័យ"}


@app.post("/api/admin/class-price")
def admin_class_price(payload: AdminClassPriceIn, request: Request):
    _require_developer_admin(request)
    if payload.MonthlyFee <= 0:
        raise HTTPException(422, "តម្លៃត្រូវតែធំជាង 0")
    with LOCK:
        conn = _connect()
        if not _admin_ok(conn, payload.AdminPIN):
            conn.close()
            raise HTTPException(403, "Admin PIN មិនត្រឹមត្រូវ")
        cur = conn.execute("UPDATE Classes SET MonthlyFee=? WHERE ClassID=?", (payload.MonthlyFee, payload.ClassID))
        if cur.rowcount == 0:
            conn.close()
            raise HTTPException(404, "រកមិនឃើញថ្នាក់")
        conn.commit()
        conn.close()
    return {"message": "បានកែតម្លៃថ្នាក់ដោយជោគជ័យ", "ClassID": payload.ClassID, "MonthlyFee": payload.MonthlyFee}


@app.post("/api/admin/change-pin")
def admin_change_pin(payload: AdminPinIn, request: Request):
    _require_developer_admin(request)
    with LOCK:
        conn = _connect()
        if not _admin_ok(conn, payload.AdminPIN):
            conn.close()
            raise HTTPException(403, "Admin PIN មិនត្រឹមត្រូវ")
        _set_setting(conn, "AdminPIN", payload.NewPIN)
        conn.commit()
        conn.close()
    return {"message": "បានប្តូរ Admin PIN ដោយជោគជ័យ"}



@app.get("/api/admin/backup")
def backup_database(request: Request):
    """Create a complete release backup containing a consistent SQLite snapshot,
    photos, and a small manifest. Only Developer Administrator may use it."""
    _require_developer_admin(request)
    if DATABASE_URL:
        raise HTTPException(501, "Database ZIP backup is available for SQLite only. Use Export to Excel for PostgreSQL data.")
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = os.path.join(BASE_DIR, "backups")
    os.makedirs(backup_dir, exist_ok=True)
    filename = f"SchoolBackup_{timestamp}.zip"
    zip_path = os.path.join(backup_dir, filename)
    temp_db = os.path.join(backup_dir, f".backup_{timestamp}.db")
    try:
        # Use SQLite's online backup API rather than copying school.db directly;
        # this also captures data safely when WAL mode is enabled.
        with LOCK:
            src_conn = _connect()
            try:
                src_conn.execute("PRAGMA wal_checkpoint(FULL)")
            except sqlite3.Error:
                pass
            dst_conn = sqlite3.connect(temp_db)
            try:
                src_conn.backup(dst_conn)
            finally:
                dst_conn.close()
                src_conn.close()

        manifest = {
            "format": "School Management System Backup",
            "version": "SQLite Database v6 + Backup Restore",
            "created_at": datetime.datetime.now().isoformat(timespec="seconds"),
            "database": "school.db",
            "includes": ["school.db", "photos/"],
        }
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("backup_manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
            z.write(temp_db, "school.db")
            if os.path.isdir(PHOTO_DIR):
                for root, dirs, files in os.walk(PHOTO_DIR):
                    for name in files:
                        full = os.path.join(root, name)
                        arc = os.path.relpath(full, BASE_DIR).replace("\\", "/")
                        z.write(full, arc)
        return FileResponse(zip_path, filename=filename, media_type="application/zip")
    except Exception as exc:
        try:
            if os.path.exists(zip_path):
                os.remove(zip_path)
        except OSError:
            pass
        raise HTTPException(500, f"Backup មិនជោគជ័យទេ: {exc}")
    finally:
        try:
            if os.path.exists(temp_db):
                os.remove(temp_db)
        except OSError:
            pass


@app.post("/api/admin/restore")
async def restore_database(request: Request):
    """Restore a SchoolBackup ZIP sent as application/zip.
    Uses raw request bytes so python-multipart is NOT required."""
    _require_developer_admin(request)
    if DATABASE_URL:
        raise HTTPException(501, "SQLite ZIP restore is not available when PostgreSQL is configured.")

    # Browsers/Windows may send ZIP as application/zip, application/x-zip-compressed,
    # application/octet-stream, or another MIME type. Do not reject a valid ZIP
    # merely because its MIME type differs.
    content_type = (request.headers.get("content-type") or "").lower()

    temp_root = tempfile.mkdtemp(prefix="school_restore_", dir=BASE_DIR)
    upload_zip = os.path.join(temp_root, "upload.zip")
    extracted = os.path.join(temp_root, "extracted")
    os.makedirs(extracted, exist_ok=True)

    try:
        raw = await request.body()
        if not raw:
            raise HTTPException(400, "Backup file ទទេ។")
        if len(raw) > 200 * 1024 * 1024:
            raise HTTPException(413, "Backup file ធំពេក (អតិបរមា 200MB)។")

        with open(upload_zip, "wb") as f:
            f.write(raw)

        try:
            with zipfile.ZipFile(upload_zip, "r") as z:
                names = z.namelist()
                db_name = next((n for n in names if n.replace("\\", "/") == "school.db"), None)
                if not db_name:
                    raise HTTPException(400, "Backup ZIP មិនមាន school.db ទេ។")

                root_abs = os.path.abspath(extracted)
                for n in names:
                    target = os.path.abspath(os.path.join(extracted, n))
                    if not (target == root_abs or target.startswith(root_abs + os.sep)):
                        raise HTTPException(400, "Backup ZIP មិនមានសុវត្ថិភាពទេ។")
                z.extractall(extracted)
        except zipfile.BadZipFile:
            raise HTTPException(400, "Backup file មិនមែនជា ZIP ត្រឹមត្រូវទេ។")

        restore_db = os.path.join(extracted, "school.db")
        if not os.path.isfile(restore_db):
            raise HTTPException(400, "រកមិនឃើញ school.db ក្នុង Backup។")

        # Validate SQLite before touching the live database.
        check = sqlite3.connect(restore_db)
        try:
            integrity = check.execute("PRAGMA integrity_check").fetchone()[0]
            if str(integrity).lower() != "ok":
                raise HTTPException(400, "school.db ក្នុង Backup ខូច ឬមិនអាចប្រើបាន។")
            tables = {r[0] for r in check.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            )}
            required = {"Students", "Payments", "Attendance", "Classes", "Users", "Settings"}
            missing = sorted(required - tables)
            if missing:
                raise HTTPException(400, "Backup មិនមែនជា Database របស់ប្រព័ន្ធនេះទេ។")
        finally:
            check.close()

        with LOCK:
            SESSIONS.clear()

            if os.path.exists(DB_PATH):
                old_conn = sqlite3.connect(DB_PATH, timeout=30)
                try:
                    old_conn.execute("PRAGMA wal_checkpoint(FULL)")
                except sqlite3.Error:
                    pass
                finally:
                    old_conn.close()

            live_dir = os.path.dirname(DB_PATH)
            restore_copy = os.path.join(live_dir, ".school_restore_ready.db")
            shutil.copy2(restore_db, restore_copy)

            for suffix in ("-wal", "-shm"):
                try:
                    os.remove(DB_PATH + suffix)
                except FileNotFoundError:
                    pass

            os.replace(restore_copy, DB_PATH)

            restored_photos = os.path.join(extracted, "photos")
            if os.path.isdir(restored_photos):
                old_photos = os.path.join(BASE_DIR, "photos_before_restore")
                try:
                    if os.path.isdir(old_photos):
                        shutil.rmtree(old_photos)
                    if os.path.isdir(PHOTO_DIR):
                        os.replace(PHOTO_DIR, old_photos)
                    shutil.copytree(restored_photos, PHOTO_DIR)
                    if os.path.isdir(old_photos):
                        shutil.rmtree(old_photos)
                except Exception:
                    if os.path.isdir(PHOTO_DIR):
                        shutil.rmtree(PHOTO_DIR, ignore_errors=True)
                    if os.path.isdir(old_photos):
                        os.replace(old_photos, PHOTO_DIR)
                    raise

        ensure_db()
        return {
            "message": "បាន Restore ទិន្នន័យដោយជោគជ័យ។ សូម Login ម្តងទៀត។",
            "restartLoginRequired": True
        }
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(500, f"Restore មិនជោគជ័យទេ: {exc}")
    finally:
        shutil.rmtree(temp_root, ignore_errors=True)


@app.get("/api/export/excel")
def export_excel(request: Request):
    """Export a fresh snapshot of the live SQLite database into an .xlsx file
    for viewing, printing, or backing up in Excel. This file is NOT read back
    in automatically — it's for human use only. The workbook and each sheet
    are protected with password 123 (matching the old version's behavior);
    unprotect a sheet in Excel with that password to edit it."""
    _current_user(request)
    conn = _connect()
    wb = Workbook()
    wb.remove(wb.active)
    for table, headers in TABLE_COLUMNS.items():
        ws = wb.create_sheet(table)
        ws.append(headers)
        ws.freeze_panes = "A2"
        for i, h in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(i)].width = max(14, len(h) + 4)
        col_list = ",".join(f'"{h}"' for h in headers)
        for row in conn.execute(f"SELECT {col_list} FROM {table}"):
            ws.append(list(row))
        if table != "Settings":
            ws.protection.set_password("123")
            ws.protection.sheet = True
    conn.close()
    wb.security.lockStructure = True
    wb.security.workbookPassword = "123"
    export_dir = os.path.join(BASE_DIR, "exports")
    os.makedirs(export_dir, exist_ok=True)
    for old in glob.glob(os.path.join(export_dir, "Export_*.xlsx")):
        try:
            os.remove(old)
        except OSError:
            pass
    filename = f"Export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    export_path = os.path.join(export_dir, filename)
    wb.save(export_path)
    return FileResponse(export_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
