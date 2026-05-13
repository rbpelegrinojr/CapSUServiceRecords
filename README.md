# CapSU Service Records System

A standalone desktop-style service records management system for **Capiz State University (CapSU)**. When the admin launches the application, it starts a local Flask server and automatically opens the system in the default web browser.

---

## Features

- 📋 **Employee Management** — Add, edit, view, and delete employee records
- 📝 **Service Records** — Manage individual service record entries per employee
- 🖨️ **DOCX Generation** — Generate official service record documents using a template
- 📥 **Excel Import** — Bulk import employees and service records from `.xlsx` files
- 📤 **Excel Export** — Export any employee's records to Excel
- ⚙️ **Settings** — Configure certifier name, title, and university info
- 🗄️ **SQLite Database** — Lightweight file-based database, no server required

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3, Flask, Flask-SQLAlchemy |
| Database | SQLite |
| Templating | Jinja2 |
| DOCX Generation | python-docx |
| Excel Import/Export | pandas, openpyxl |
| Launcher | Python webbrowser + threading |
| Styling | Bootstrap 5 (CDN) |

---

## Project Structure

```
CapSUServiceRecords/
├── app.py                        # Flask entry point + auto-browser launcher
├── config.py                     # App configuration
├── database.py                   # DB init helper
├── create_template.py            # Generates DOCX template programmatically
├── create_sample_data.py         # Generates sample Excel file
├── requirements.txt
├── README.md
│
├── models/
│   ├── employee.py
│   ├── service_record.py
│   └── setting.py
│
├── routes/
│   ├── employees.py
│   ├── records.py
│   ├── import_export.py
│   ├── print_doc.py
│   └── settings.py
│
├── templates/                    # Jinja2 HTML templates
├── static/                       # CSS and JS
├── docx_templates/               # DOCX template (auto-created on first run)
├── sample_data/                  # Sample Excel file (auto-created on first run)
└── data/                         # SQLite database (auto-created on first run)
```

---

## Setup Instructions

### 1. Prerequisites

- Python 3.8 or higher
- pip (Python package installer)

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the Application

```bash
python app.py
```

The app will:
1. Auto-create the SQLite database at `data/records.db`
2. Seed default settings (certifier name, university info)
3. Generate the DOCX template at `docx_templates/service_record_template.docx`
4. Generate the sample Excel file at `sample_data/sample_service_records.xlsx`
5. Open your default browser at `http://127.0.0.1:5000`

---

## How to Use

### Adding an Employee

1. Click **Employees** in the navbar
2. Click **Add Employee**
3. Fill in the form and click **Save Employee**

### Adding Service Records

1. Open an employee's profile
2. Click **Add Service Record**
3. Fill in the service record details and save

### Importing from Excel

1. Click **Import** in the navbar
2. Download the sample template to see the expected format
3. Prepare your `.xlsx` file with the required columns
4. Upload the file and click **Import Records**

**Required columns:** `Surname`, `Given Name`

**Optional columns:** `Middle Name`, `Birth Date`, `Birth Place`, `Date From`, `Date To`, `Designation`, `Status`, `Monthly Salary`, `Annual Salary`, `Station/Place of`, `Branch`, `LV ABS WO Pay`, `Separation Date`, `Separation Cause`

Each row = one service record. Rows with the same Surname + Given Name are linked to the same employee.

### Printing a Service Record PDF

1. Open an employee's profile
2. Click **Print Service Record**
3. A PDF preview opens in a new tab and triggers the print dialog

### Exporting to Excel

1. Open an employee's profile
2. Click **Export to Excel**

### Configuring Settings

1. Click **Settings** in the navbar
2. Update certifier name, title, and university info
3. Click **Save Settings**

---

## Generating the DOCX Template Manually

If the template was not auto-generated, run:

```bash
python create_template.py
```

This creates `docx_templates/service_record_template.docx`.

---

## Generating Sample Data Manually

```bash
python create_sample_data.py
```

This creates `sample_data/sample_service_records.xlsx` with 5 sample employees and multiple service record rows.

---

## Database

The SQLite database is stored at `data/records.db`. To reset the database, simply delete this file and restart the app.

---

## Packaging as a Standalone `.exe` (Windows)

### Install PyInstaller

```bash
pip install pyinstaller
```

### Build

```bash
pyinstaller --onefile --noconsole \
    --add-data "templates;templates" \
    --add-data "static;static" \
    --add-data "docx_templates;docx_templates" \
    --add-data "sample_data;sample_data" \
    app.py
```

The executable will be in the `dist/` folder.

> **Note:** On macOS/Linux, use `:` instead of `;` as the path separator:
> ```bash
> pyinstaller --onefile --noconsole \
>     --add-data "templates:templates" \
>     --add-data "static:static" \
>     --add-data "docx_templates:docx_templates" \
>     --add-data "sample_data:sample_data" \
>     app.py
> ```

---

## Notes

- The app runs on `http://127.0.0.1:5000` by default
- The browser opens automatically 1.5 seconds after the server starts
- All data is stored locally — no internet connection required
- Salary values support decimal precision (e.g., `18,549.00`)
- Use `-do-` for repeated values in service records (as per CapSU format)
