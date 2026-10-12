# School Management System – SQLite Database v6

## Database
- **Live database: `school.db`** (SQLite — a single file, no server/setup needed, same idea as the old `Data.xlsx` but built to handle thousands of records and multiple people using it at the same time without slowing down).
- User accounts are stored in `school.db` → `Users` table.
- The login page does not display any default password.

### ⚠️ Upgrading from an older (Excel-based) version — read this first
If you already have a working copy of this app with a real `Data.xlsx` full of your students/payments:
1. **Do not delete your existing `Data.xlsx`.** Just copy in the new `main.py`, `index.html`, `README.md`, `run.bat`, `install.bat`, `requirements.txt`, and `schema_setup.py` from this update, keeping your existing `Data.xlsx` and `photos` folder as they are.
2. The first time you run the new `run.bat`, the app automatically creates `school.db` and **imports everything from your existing `Data.xlsx` into it, once.** Your students, payments, attendance, settings and user accounts all carry over.
3. After that first run, `school.db` is the live database. Editing `Data.xlsx` by hand afterwards will **not** be picked up — see "Export to Excel" below if you want to view or back up your data in Excel format going forward.
4. If you're starting completely fresh (no existing `Data.xlsx`), just run `install.bat` then `run.bat` as normal — a new `school.db` is created automatically with the default classes and `admin` account.

### Export to Excel
Settings → **Database → Export to Excel** downloads a fresh `.xlsx` snapshot of everything currently in `school.db`, for viewing, printing, or backing up in Excel. Protected the same way as before (workbook/sheet password `123`). This exported file is for human use only — it is **not** read back into the app automatically.

## Account roles
- **Developer Administrator** (the existing `admin` account is treated as this role): full system administration.
- **Teacher**: operational data access; can edit student data and change a student's class. No Developer Administrator settings.
- **BM**: operational data access; can edit student data and change a student's class. No Developer Administrator settings.
- **Receptionist**: operational data access; can edit student data and change a student's class. No Developer Administrator settings.

Only the Developer Administrator can:
- Open the Settings button and manage all system settings.
- Create and edit Teacher / BM / Receptionist accounts from Settings.
- Change class monthly fees and Admin PIN.
- Upload the school logo.
- Change the school name, address and phone.
- Edit all invoice template text fields.

## School branding and invoice template
The Developer Administrator page can edit:
- School name, address, phone, logo
- Invoice title and every field label (Receipt, Full Name, Code, Course Period, Date, Next due, Cashier, Payer Account, Total Paid)
- Invoice footer

These settings are stored in `school.db` → `Settings` table. The logo file is stored in the `photos` folder.

## Payment / invoice
Every payment automatically records the currently logged-in user's `DisplayName` (Received By / Cashier) and `Username` (Payer Account), plus the new receipt fields: Subject, Time, Room, Register, Total Amount, Discount, Deposit, Owed, and Payment Method (Cash/Check). The invoice prints as a single A5-landscape page styled like a real payment receipt.

## Student class editing
Teacher, BM and Receptionist accounts can change a student's class from the Student List. The server also provides a protected student-update endpoint for operational data editing (Subject, Room, Time, Register No. included).

## Run
1. Run `install.bat` (on the computer that will act as the server — the one keeping `school.db`).
2. Run `run.bat`. Keep this window open; closing it stops the server for everyone.
3. Open the app in Firefox or Chrome using one of the addresses printed by `run.bat`.

## Using it from multiple computers with one shared database
This app is client/server: `main.py` + `school.db` is the "server", and `index.html` is just the screen. To share one database across several computers on the same WiFi/LAN:

1. Pick **one** computer to be the server. Put the whole `sms_package` folder there and run `install.bat` then `run.bat` on it only. Do **not** run `run.bat` on the other computers.
2. `run.bat` prints one or more addresses like `http://192.168.1.23:8000/`. Note the server computer's address.
3. On every other computer, just open a normal web browser (Chrome/Firefox/Edge) and type that address into the address bar — no files need to be copied to those computers. Everyone will read and write the same `school.db` on the server computer.
4. If other computers can't connect:
   - Make sure all computers are on the same WiFi/router (not different networks or mobile hotspots).
   - Windows Firewall may block incoming connections the first time — choose "Allow access" if prompted, or add an inbound rule for TCP port 8000.
   - The server computer's LAN IP can change if the router reassigns it; giving the server computer a fixed/static local IP in your router avoids the address changing later.
5. Keep the server computer turned on and `run.bat` running while others are using the system. SQLite safely handles multiple people reading and writing at the same time, much better than the old Excel file did at scale (thousands of students, years of attendance records).
6. For use over the internet (not just local WiFi) rather than a single LAN, you'd need port-forwarding/a hosting setup and additional security — ask if you need help with that.

## Default existing account
- Username: `admin`
- Password: `1234`

For security, create/edit accounts from Settings after first setup (don't hand-edit `school.db` directly — use the app or, for a human-readable copy, Export to Excel).

## Settings button
All Developer Administrator controls are consolidated under **Settings / ការកំណត់ប្រព័ន្ធ**: school branding, logo, invoice template, class fees, Admin PIN, database export, and user accounts. The user manager validates input, reports duplicate usernames clearly, and provides an Edit action for existing accounts.

## Backup Data / Restore Data
- Settings → Database → **Backup Data** creates a ZIP containing a consistent copy of `school.db`, the `photos` folder, and a backup manifest.
- **Restore Data** accepts a ZIP created by this application and restores the database and photos.
- Restore is restricted to the Developer Administrator and requires confirmation.
- After a successful restore, all active sessions are cleared and users must log in again.
- Keep at least one backup outside the application folder (USB drive/cloud storage) for disaster recovery.

## One-click start
- Double-click **START_SCHOOL_SYSTEM.bat** (recommended).
- It starts the local FastAPI server and opens the system in the browser automatically.
- `index.html` also redirects to `http://127.0.0.1:8000/` if opened directly, but a normal HTML file cannot start a Python server by itself because Firefox blocks local file pages from launching programs.

## v5 Restore fix
Restore now uploads the ZIP as raw `application/zip`, so the system does not require `python-multipart` and avoids the FastAPI form-parsing error.

## v6 Restore compatibility fix
Restore no longer rejects a valid ZIP based on browser MIME type. The browser sends the selected backup as raw bytes and the server validates the ZIP/database contents directly.

## v7 CRUD Update
- Student: Edit / Delete
- Teacher/User: Edit / Delete
- Grade / ថ្នាក់ / កំរិត: Add / Edit / Delete
- Room / បន្ទប់: Edit / Delete
- Student Room field can be updated from Student Edit
- Class deletion is protected when students are still assigned to that class.
- Student deletion also removes that student's Payment and Attendance records.

## v8 CRUD enhancement
Student Edit covers DOB, parent occupation, address, room, subject, register number, schedule and status. Teacher/User Edit includes Active/Disabled. Existing Grade/ថ្នាក់/កំរិត and Room Edit/Delete actions are retained.

## v9 School Departments
The school now has 3 department/program choices:
1. ចំណេះដឹងទូទៅ ភាសាខ្មែរ
2. ភាសាអង់គ្លេស ពេញម៉ោង — ព្រឹក និងថ្ងៃ
3. ភាសាអង់គ្លេស ក្រៅម៉ោង / ភាសាចិន ក្រៅម៉ោង — ល្ងាច
Department is saved with each student, shown in Student List, filterable, included in Excel data, and editable.

## v10 Department-dependent Register Form
- Department dropdown has 4 separate programs.
- General Khmer and Full-time English: Time = ព្រឹក / ល្ងាច.
- Part-time English and Part-time Chinese: Time = 4:00-5:00 PM / 5:00-6:00 PM / 6:00-7:00 PM.
- Grade dropdown shows only grades assigned to the selected Department.
- Grade management now assigns each grade to a Department, so administrators can configure the dependent Grade list.

## v12 Group menu placement / role
- Group / ក្រុម menu moved to the bottom of the sidebar, directly below របាយការណ៍, matching the requested location.
- Group menu and Group CRUD are available only to Admin / Developer Administrator and BM.
- Teacher and Receptionist do not see Group and backend endpoints return 403 for unauthorized roles.

## Database Relationship + Test Data Fix (v4.0.5)
- SQLite foreign-key enforcement is enabled for runtime connections.
- Students -> Classes: `ON DELETE SET NULL`.
- Payments / Attendance / Scores -> Students: `ON DELETE CASCADE`.
- Sales -> Students: `ON DELETE SET NULL`.
- Student deletion now removes related payments, attendance, scores, and sales.
- Startup repairs orphan child records without inventing parent records.
- Includes an idempotent linked QA dataset: 50 `TEST-0001` ... `TEST-0050` students, 50 payments, 50 attendance records, 150 score records, 50 books, 50 uniforms, and 50 sales.
