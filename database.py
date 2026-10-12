"""Small adapter that lets the existing SQLite-style queries run on PostgreSQL."""
import re
from urllib.parse import urlparse, urlunparse

import psycopg
from psycopg.rows import tuple_row


_AUTO_IDS = {
    "WarehouseBooks": "BookID",
    "WarehouseUniforms": "UniformID",
    "Rooms": "RoomID",
    "Sales": "SaleID",
}
_CANONICAL = {name.lower(): name for name in """
Key Value ClassID ClassName MonthlyFee Department StudentID StudentCode Surname GivenName LatinName FullName Gender DOB ParentName ParentOccupation GuardianPhone Address EnrollmentDate Status PhotoPath Subject RegisterNo Room Schedule PaymentID InvoiceNo DurationMonths AmountPaid PaymentDate NextDueDate ReceivedBy PayerAccount TotalAmount Discount Deposit Owed PaymentMethod CheckNumber AttendanceID AttDate ScoreID ExamName Score MaxScore ExamDate UserID Username Password PasswordHash DisplayName Role Active BookID BookName Stock Price CreatedAt UniformID UniformType Size RoomID RoomName Capacity SaleID SaleType ItemID ItemName Quantity UnitPrice AdminPIN
""".split()}


class Row:
    def __init__(self, values, columns):
        self._values = tuple(values)
        self._columns = tuple(_CANONICAL.get(name.lower(), name) for name in columns)
        self._mapping = dict(zip(self._columns, self._values))

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._values[key]
        if key in self._mapping:
            return self._mapping[key]
        return self._mapping[_CANONICAL.get(key.lower(), key)]

    def keys(self):
        return self._columns

    def __iter__(self):
        return iter(self._values)

    def __len__(self):
        return len(self._values)


class Cursor:
    def __init__(self, cursor, lastrowid=None):
        self._cursor = cursor
        self.lastrowid = lastrowid

    def _row(self, values):
        return None if values is None else Row(values, [c.name for c in self._cursor.description])

    def fetchone(self):
        return self._row(self._cursor.fetchone())

    def fetchall(self):
        return [self._row(row) for row in self._cursor.fetchall()]

    def __iter__(self):
        for row in self._cursor:
            yield self._row(row)


class Connection:
    def __init__(self, url):
        self._conn = psycopg.connect(url, row_factory=tuple_row, connect_timeout=10)

    @staticmethod
    def _sql(sql):
        ignored = bool(re.search(r"INSERT\s+OR\s+IGNORE\s+INTO", sql, re.I))
        sql = re.sub(r"INSERT\s+OR\s+IGNORE\s+INTO", "INSERT INTO", sql, flags=re.I)
        sql = re.sub(r"\bdate\(([^()]*)\)", r"CAST(\1 AS DATE)", sql, flags=re.I)
        sql = sql.replace("?", "%s")
        if ignored:
            sql = sql.rstrip().rstrip(";") + " ON CONFLICT DO NOTHING"
        return sql

    def execute(self, sql, params=()):
        if re.match(r"\s*PRAGMA\s+", sql, re.I):
            return Cursor(self._conn.cursor())
        sql = self._sql(sql)
        table = re.search(r"INSERT\s+INTO\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(([^)]*)\)", sql, re.I)
        cursor = self._conn.cursor()
        if table and table.group(1) in _AUTO_IDS:
            id_col = _AUTO_IDS[table.group(1)]
            if id_col.lower() not in {part.strip().strip('"').lower() for part in table.group(2).split(",")}:
                cursor.execute(sql + f' RETURNING "{id_col}"', params)
                value = cursor.fetchone()
                return Cursor(cursor, value[0] if value else None)
        cursor.execute(sql, params)
        return Cursor(cursor)

    def executemany(self, sql, values):
        self._conn.cursor().executemany(self._sql(sql), values)

    def executescript(self, script):
        with self._conn.cursor() as cursor:
            for statement in script.split(";"):
                if statement.strip():
                    cursor.execute(self._sql(statement))

    def commit(self):
        self._conn.commit()

    def close(self):
        self._conn.close()


def connect(url):
    parsed = urlparse(url)
    if parsed.scheme == "postgres":
        url = urlunparse(parsed._replace(scheme="postgresql"))
    return Connection(url)
