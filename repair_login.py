import sqlite3, hashlib, os
BASE_DIR=os.path.dirname(os.path.abspath(__file__))
DB=os.path.join(BASE_DIR,"school.db")

def h(p): return hashlib.sha256(p.encode('utf-8')).hexdigest()
conn=sqlite3.connect(DB)
conn.execute("CREATE TABLE IF NOT EXISTS Users (UserID INTEGER PRIMARY KEY, Username TEXT, Password TEXT, PasswordHash TEXT, DisplayName TEXT, Role TEXT, Active TEXT)")
# Prefer existing Administrator; otherwise rename legacy admin; otherwise create.
row=conn.execute("SELECT UserID FROM Users WHERE lower(trim(Username))='administrator' ORDER BY UserID LIMIT 1").fetchone()
if not row:
    row=conn.execute("SELECT UserID FROM Users WHERE lower(trim(Username))='admin' ORDER BY UserID LIMIT 1").fetchone()
if row:
    uid=row[0]
    conn.execute("UPDATE Users SET Username='Administrator',Password='123',PasswordHash=?,DisplayName='Administrator',Role='Developer Administrator',Active='Yes' WHERE UserID=?",(h('123'),uid))
else:
    conn.execute("INSERT INTO Users (UserID,Username,Password,PasswordHash,DisplayName,Role,Active) VALUES (1,'Administrator','123',?,'Administrator','Developer Administrator','Yes')",(h('123'),))
# Keep only one built-in Administrator account if duplicates exist.
rows=conn.execute("SELECT UserID FROM Users WHERE lower(trim(Username))='administrator' ORDER BY UserID").fetchall()
for r in rows[1:]: conn.execute("DELETE FROM Users WHERE UserID=?",(r[0],))
conn.commit(); conn.close()
print('LOGIN OK: Administrator / 123 / Developer Administrator / Full Access')
