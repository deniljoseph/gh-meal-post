============================================================
  WORKFORCE & MEAL MANAGEMENT SYSTEM - v1.0
  Complete Local Web Application
============================================================

QUICK START
-----------
1. Install Python 3.10+ from https://python.org
   (Check "Add Python to PATH" during installation)

2. Double-click START.bat
   The browser will open automatically.

3. Login with:
   Username: admin
   Password: admin123

URLS
----
Main Application:   http://localhost:8000
Meal Lookup Portal: http://localhost:8000/lookup  (no login needed)

DEFAULT ACCOUNTS
----------------
admin    / admin123  -> Full access
hr       / hr123     -> HR management
supervisor/ sup123   -> Attendance & food updates
viewer   / view123   -> Read-only

FEATURES
--------
✅ Dashboard with live charts
✅ Employee management (Add/Edit/Delete/Search)
✅ Bulk import from Excel
✅ Export employees to Excel
✅ Daily food count calculator
✅ Attendance marking (mark absent/present)
✅ Suspension management (auto-exclude from food count)
✅ Fasting management (auto-adjust food count)
✅ Location & Shift management
✅ Food preference management
✅ Reports (Excel export)
✅ Public meal lookup portal (no login required)
✅ Kiosk mode for cafeteria screens
✅ Dark mode support
✅ Audit log
✅ Role-based access control

FOOD COUNT LOGIC
----------------
Daily food count automatically:
- Excludes suspended employees
- Excludes absent employees (marked that day)
- Excludes fasting employees
- Updates instantly when any status changes

DATABASE
--------
File: workforce.db (SQLite, auto-created on first run)
Location: Same folder as app.py
Backup: Simply copy workforce.db file

IMPORTING FROM EXCEL
--------------------
Your existing Excel file can be imported via:
Employees -> Import Excel button

Supported columns (any order, auto-detected):
- Emp ID / Employee ID
- Full Name / Name
- Department / Dept
- Designation / Position
- Location / Camp / Site
- Shift / Shift Name
- Food Preference / Food Pref
- Nationality
- Mobile / Phone
- Joining Date

TROUBLESHOOTING
---------------
Q: "python is not recognized"
A: Re-install Python and check "Add Python to PATH"

Q: Port 8000 already in use
A: Edit app.py last line, change port=8000 to port=8001
   Then access http://localhost:8001

Q: Missing module errors
A: Run: pip install -r requirements.txt

Q: How to upgrade to PostgreSQL?
A: Change DB_PATH and connection string in app.py
   Replace sqlite3.connect with psycopg2/asyncpg

FOLDER STRUCTURE
----------------
workforce_app/
├── app.py              <- Main application (backend)
├── requirements.txt    <- Python packages needed
├── START.bat           <- One-click launcher (Windows)
├── README.txt          <- This file
├── workforce.db        <- Database (auto-created)
└── static/
    └── index.html      <- Frontend (all-in-one)

============================================================
