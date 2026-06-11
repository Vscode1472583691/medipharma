# MediPharma - Medicine Billing System

Professional medicine and pharmacy billing software with inventory management, payment tracking, and real-time alerts.

## Quick Start

### Option 1: Run with Python (Development)

```bash
# Create virtual environment
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py
```

Visit: http://127.0.0.1:5000

### Option 2: Run Standalone EXE (No Python Required!)

#### Using Command Prompt:
```bash
build.bat
```

#### Using PowerShell:
```powershell
.\build.ps1
```

The EXE will be created at: `dist\MediPharma.exe`

Simply double-click to run the application!

---

## Default Login Credentials

- **Username**: `admin`
- **Password**: `admin1472`

---

## Features

✅ **Inventory Management**
- Track stock levels
- Monitor expiry dates
- Low stock alerts
- Automatic notifications

✅ **Billing & Sales**
- Create professional invoices
- GST calculation
- Discount handling
- Print-friendly bills
- Payment modes: Full / Partial / None

✅ **Purchase Management**
- Record supplier purchases
- Track payment terms
- Schedule payment reminders
- Supplier contact management

✅ **Payment Tracking**
- Customer payment history
- Supplier payment records
- Scheduled payment reminders
- Email notifications for due dates

✅ **Customers & Suppliers**
- Manage customer database
- Track supplier details
- Contact information storage

✅ **Reports & Analytics**
- Profit & Loss statements
- Sales reports
- Purchase history
- Date-range filtering

✅ **Professional UI**
- Modern dashboard
- Real-time notifications
- Responsive design
- Role-based access

---

## Environment Configuration

### Optional Environment Variables

Create a `.env` file in the project root:

```env
# Security
SECRET_KEY=your-secret-key-here
DEFAULT_ADMIN_PASS=your-password

# Database
DATABASE_URL=sqlite:///billing.db

# Email Notifications (for reminders and alerts)
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=your-email@gmail.com
SMTP_PASS=your-app-password
ALERT_EMAIL_TO=owner@example.com
```

---

## Running Tests

```bash
pytest -q
```

---

## Building EXE (Advanced)

The build scripts use PyInstaller to create standalone executables.

**Prerequisites:**
- Python 3.9+
- Windows OS

**Build Steps:**
1. Run `build.bat` or `build.ps1`
2. Wait for the build to complete
3. Find `dist\MediPharma.exe`

The EXE includes:
- All dependencies bundled
- Templates and static files
- Database (SQLite)
- No Python installation required!

Persistent storage when bundled
- When the application is packaged as an EXE the SQLite database is placed in a persistent folder in the user's home directory: `%USERPROFILE%\\MediPharma\\billing.db` (Windows). This ensures your data remains after updates or across runs until you delete it manually.

Building 32-bit and 64-bit EXEs
- To produce installers for both 32-bit and 64-bit Windows you need to build on matching Python builds (PyInstaller does not reliably cross-compile between architectures).

Steps:
1. Install 64-bit Python and create a virtualenv, then run `build.bat` to produce a 64-bit EXE.
2. Install 32-bit Python (the installer labeled "Windows x86") and create a separate virtualenv using that interpreter, then run `build.bat` from that environment to produce a 32-bit EXE.

Example (PowerShell):
```powershell
# 64-bit build (example)
& 'C:\\Python3x64\\python.exe' -m venv venv64
venv64\\Scripts\\Activate.ps1
pip install -r requirements.txt
```

## Database Migration Repair

Older SQLite databases may be missing legacy `sale` columns like `gst_percent`, `discount`, and `seller_name`.
Run the repair script against the database file you want to fix:

```bash
python migrate_add_sale_columns.py
```

If you use a custom database path, set `DATABASE_URL` first, for example:

```bash
set DATABASE_URL=sqlite:///billing.db
python migrate_add_sale_columns.py
```
.\build.bat

# 32-bit build (example)
& 'C:\\Python3x86\\python.exe' -m venv venv32
venv32\\Scripts\\Activate.ps1
pip install -r requirements.txt
.\build.bat
```

Note: The produced EXE runs a bundled Flask server and UI in a desktop window; you do not need to run it from a browser. If you prefer a true native installer (MSI/NSIS) wrap the produced EXE using an installer builder.

Creating a native Windows installer (NSIS)
- A simple NSIS script `installer.nsi` is included in the project root. It packages the `dist` output into `MediPharma-Setup-x64.exe` and creates Start Menu and Desktop shortcuts.

Build steps (on Windows):
1. Install NSIS from https://nsis.sourceforge.io/ and ensure `makensis` is in your PATH.
2. Run `build_installer.bat` from the project root — it will invoke `makensis installer.nsi` and produce `MediPharma-Setup-x64.exe`.

If you don't have NSIS available, you can copy the `dist\MediPharma.exe` folder to users manually or use third-party installer builders (InnoSetup, WiX, etc.).

Automating multi-architecture builds and portable packages
- `build_all.bat [path-to-32bit-python]` — builds the 64-bit EXE (using current Python) and, if you provide the path to a 32-bit Python executable, builds a 32-bit EXE into `dist32`.
- `package_portable.bat` — creates ZIP packages from `dist` and `dist32` into `packages\MediPharma-x64.zip` and `packages\MediPharma-x86.zip` using PowerShell's `Compress-Archive`.

Examples:

```powershell
# 64-bit only (already built by default)
.\build.bat

# 64-bit + 32-bit (supply 32-bit Python path)
.\build_all.bat "C:\\Python3x86\\python.exe"

# Create portable ZIPs
.\package_portable.bat
```

---

## Technical Stack

- **Framework**: Flask 2.2.5
- **Database**: SQLAlchemy (SQLite)
- **Scheduling**: APScheduler
- **Security**: Werkzeug password hashing
- **Frontend**: Bootstrap 5 + Font Awesome
- **Packaging**: PyInstaller

---

## File Structure

```
├── app.py                 # Main Flask application
├── launcher.py           # EXE launcher
├── templates/            # HTML templates
│   ├── base.html
│   ├── dashboard.html
│   ├── billing.html
│   ├── inventory.html
│   ├── purchase.html
│   ├── suppliers.html
│   ├── customers.html
│   ├── payments.html
│   ├── reminders.html
│   ├── reports.html
│   ├── notifications.html
│   ├── print_bill.html
│   └── change_password.html
├── build.bat             # Windows batch build script
├── build.ps1             # PowerShell build script
├── MediPharma.spec       # PyInstaller spec file
├── requirements.txt      # Python dependencies
├── billing.db            # SQLite database (auto-created)
└── README.md             # This file
```

---

## Notes

- First run will create `billing.db` and initialize with default admin
- Database is stored locally (no cloud sync by default)
- All calculations in Indian Rupees (₹)
- 24-hour time format
- Responsive design works on desktop, tablet, mobile

---

## Troubleshooting

**EXE won't start:**
- Ensure Windows Firewall allows the app
- Try running as Administrator
- Check that port 5000 is available

**Browser doesn't open automatically:**
- Manually visit http://127.0.0.1:5000
- Check if port 5000 is in use: `netstat -ano | findstr 5000`

**Database issues:**
- Delete `billing.db` to reset
- Database recreates automatically on first run

---

## Features Coming Soon

- SMS notifications
- Advanced reporting (PDF export)
- Barcode scanning
- Multi-user support with role management
- Cloud backup integration
- API for external integrations

---

## License

Proprietary - SRI RAM POULTRY FARM BILLING SYSTEM

---

## Support

For issues or feature requests, contact the development team.
