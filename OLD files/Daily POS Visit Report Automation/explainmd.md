# 📋 POS Daily Report Automation — Full Script Explanation

> **File:** `pos_daily_report_automation.py`  
> **Purpose:** Automates the daily collection, formatting, compression, and email distribution of Point-of-Sale (POS) visit reports for Grameenphone (GP) field teams.

---

## 📦 1. Imports & Dependencies

```python
import os, shutil, smtplib, pandas as pd, psycopg2, zipfile
from datetime import datetime, timedelta
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Side, Border, Alignment
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.utils import get_column_letter, column_index_from_string
from sqlalchemy import create_engine, text
import matplotlib.pyplot as plt
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.utils import make_msgid
from email import encoders
import xlwings as xw
import re, time
from PIL import ImageGrab
import warnings
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
```

| Library | Purpose |
|---|---|
| `os`, `shutil` | File and folder operations (create, rename, copy) |
| `pandas` | Data manipulation and Excel writing |
| `psycopg2` | Low-level PostgreSQL connection (used indirectly) |
| `sqlalchemy` | High-level database engine and SQL execution |
| `openpyxl` | Excel workbook creation and styling |
| `zipfile` | Compressing report folders into `.zip` archives |
| `smtplib` | Sending emails via SMTP (Gmail) |
| `email.*` | Building MIME email messages with HTML body |
| `xlwings` | Excel automation via COM (imported but not actively called in main flow) |
| `matplotlib` | Chart generation (imported for potential chart use) |
| `PIL.ImageGrab` | Screenshot capture (imported for potential use) |
| `re` | Regular expressions for file name pattern matching |
| `pydrive` / `pydrive2` | Google Drive file upload and link generation |
| `warnings` | Suppressing non-critical xlsxwriter warnings |

---

## 🔐 2. Database Configuration

```python
DB_PARAMS = {
    "dbname": "gp_dev",
    "user": "report_user",
    "password": "",   # ⚠️ Credential placeholder
    "host": "",       # ⚠️ Credential placeholder
    "port": 5432
}

DATABASE_URL = f"postgresql://{DB_PARAMS['user']}:{DB_PARAMS['password']}@{DB_PARAMS['host']}:{DB_PARAMS['port']}/{DB_PARAMS['dbname']}"
engine = create_engine(DATABASE_URL)
```

- Connects to a **PostgreSQL** database (originally an AWS RDS instance on `ap-southeast-1`).
- Uses **SQLAlchemy** engine for executing parameterized queries.
- Credentials are **intentionally blanked out** in the committed version for security.

---

## 🗄️ 3. SQL Query — `pos_daily_rpt(custom_date)`

This is the core data extraction function. It builds a complex SQL query using **4 Common Table Expressions (CTEs)** and returns a `text()` SQLAlchemy object.

### CTE 1: `mv_data` — Market Visit Records

```sql
WITH mv_data AS (
    SELECT ... RANK() OVER (PARTITION BY o.outlet_code, mvc.user_id ORDER BY mvc.id DESC) rank
    FROM gp.market_visit_contacts mvc, gp.outlets o, gp.routes r
    WHERE mvc.visit_date = '{custom_date}'
)
WHERE rank = 1 AND x.route NOT LIKE '%V2TECH%'
```

**What it does:**
- Pulls visit records for a specific date from `gp.market_visit_contacts`
- Joins with `gp.outlets` and `gp.routes` to get outlet code, route, and geo info
- Uses `RANK()` window function — picks **only the latest visit per outlet per user** (rank = 1)
- Excludes internal V2TECH test routes

**Key columns extracted:**
- `visit_id`, `visit_date`, `outlet_code`, `route`, `is_skitto`
- `visit_lat_long` — concatenated GPS coordinates
- `operational_status`, `scheduled_visit`

---

### CTE 2: `user_data` — ME (Market Executive) Details

```sql
user_data AS (WITH loc_data AS (
    SELECT Circle, Region, Cluster, Territory, Distribution House, DH Code, DH Lat-Long
    FROM gp.locations c, r, cl, tr, h, gp.house dh
    WHERE [hierarchy conditions]
))
SELECT ld.*, u.id, ui.full_name, ui.official_contact, u.email, b.bundle_code
FROM gp.users u, gp.user_infos ui, gp.users_bundle_maps ubm, gp.bundles b, gp.bundles_house_maps bhm, loc_data ld
WHERE [mapping conditions]
AND '{custom_date}' BETWEEN ubm.from_date AND COALESCE(ubm.to_date, '{custom_date}')
```

**What it does:**
- Builds the full **geographic hierarchy**: Circle → Region → Cluster → Territory → Distribution House
- Joins ME (field agent) user profile with their assigned DH and bundle/route
- Date-range filtering ensures only MEs **active on that exact date** are included
- Excludes test bundles and test DHs

**Key columns extracted:**
- `Circle`, `Region`, `Cluster`, `Territory`, `Distribution House`, `DH Code`
- `ME Name`, `ME Contact No.`, `ME Email`, `ME Code`

---

### CTE 3: `survey_data` — Visit Survey Answers (Pivot)

```sql
survey_data AS (
    SELECT visit_id,
        MAX(CASE WHEN question = 'pos_status' THEN answer END) AS pos_status,
        MAX(CASE WHEN question = 'posm_exists' THEN answer END) AS posm_exists,
        MAX(CASE WHEN question = 'pre_execution_photo' THEN answer END) AS pre_execution_photo,
        MAX(CASE WHEN question = 'gp_fascia' THEN answer END) AS gp_fascia,
        ...
    FROM gp.market_visit_contact_flow_maps mvcf
    WHERE mvcf.visit_date = '{custom_date}'
    GROUP BY visit_id
)
```

**What it does:**
- The survey answers are stored as **key-value rows** (`question`, `answer`) in `market_visit_contact_flow_maps`
- Uses **conditional aggregation (PIVOT)** with `MAX(CASE WHEN ...)` to convert rows into columns
- Photo fields are converted to **full S3 URLs** for access:
  ```
  https://gp-web-uploads.s3.ap-southeast-1.amazonaws.com/Production/gp/Images/Retail/MarketVisit/YYYYMM/<filename>
  ```

**Key columns extracted (30+ survey fields):**
- `pos_status`, `pos_structure`, `pos_location`, `pos_business_type`
- `posm_exists`, `old_posm_list`, `do_posm_remove`, `removed_posm_counts`
- `gp_fascia`, `gp_fascia_type`, `other_fascia`
- `sim_sell`, `gp_sim_sell`, `other_sim_sell`
- `finger_print_scanner`, `other_company_scanner`
- `pre_execution_photo`, `post_execution_photo`, `execution_photo_left/center/right`
- `temporarily_&_permanently_closed_photo`, `not_found_&_moved_photo`

---

### CTE 4: `posm_data` — POSM Material Counts (Pivot)

```sql
posm_data AS (
    SELECT mvp.visit_id,
        SUM(COALESCE(CASE WHEN posm."name" = 'FST_Sim Bikroy_GA_NOV_24' THEN mvp.amount ELSE 0 END, 0)) "FST_Sim Bikroy_GA_NOV_24",
        SUM(COALESCE(CASE WHEN posm."name" = 'SS_Limitless_DATA_NOV_24' THEN mvp.amount ELSE 0 END, 0)) "SS_Limitless_DATA_NOV_24",
        ...  -- 60+ POSM items
    FROM gp.market_visit_posm_counts mvp
    JOIN gp.materials posm ON mvp.posm_id = posm.id
    WHERE mvp.visit_date = '{custom_date}'
    GROUP BY mvp.visit_id
)
```

**What it does:**
- Pulls Point-of-Sale Material (POSM) placement counts per visit
- POSM items include: Festoons (FST), Shop Screens (SS), Cover Stickers (COVS), Posters (PSTR), Poster Display Boards (PDB), Banners (BNT), Stickers (ST), etc.
- Each POSM has a naming convention: `TYPE_Description_Category_Month_Year`
- Items from **Nov 2024 to July 2026** are tracked
- Commented-out items (`--`) represent discontinued/inactive POSM campaigns
- Uses `SUM + COALESCE` to safely handle NULL amounts as 0

**POSM Naming Convention:**
| Prefix | Type |
|---|---|
| `PSTR` | Poster (Regular, Small, Large, Medium) |
| `FST` | Festoon |
| `SS` | Shop Screen |
| `COVS` | Cover Sticker |
| `PDB` | Poster Display Board |
| `BNT` | Banner |
| `ST` | Sticker |
| `DLR` | Dealer |

---

### Final SELECT — Joining All 4 CTEs

```sql
SELECT mv.*, ud.*, sd.*, pd.*
FROM mv_data mv
LEFT JOIN user_data ud ON mv.me = ud."ME Code"
LEFT JOIN survey_data sd ON mv.visit_id = sd.visit_id
LEFT JOIN posm_data pd ON mv.visit_id = pd.visit_id
```

**What it produces:**
- One row per unique POS visit
- Full columns: Visit Info + Geographic Hierarchy + ME Details + Survey Answers + POSM Counts
- Photo URLs built inline using `CONCAT` with S3 base path + date + filename
- Left joins ensure visits without user/survey/posm data are still included

---

## 🔄 4. Helper Functions

### `fetch_raw_data(custom_date, query_function)`
```python
def fetch_raw_data(custom_date, query_function):
    query = query_function(custom_date)
    df = pd.read_sql_query(query, con=engine)
    return df
```
- Calls the query function with the date
- Executes it against the database using `pandas.read_sql_query()`
- Returns a DataFrame; returns empty DataFrame on error

---

### `get_day_with_suffix(day)` & `format_date_for_file_name()`
```python
def get_day_with_suffix(day):
    # Returns "1st", "2nd", "3rd", "4th"... etc.

def format_date_for_file_name(start_date_str, end_date_str):
    # Returns "25th June to 9th July"
```
- Converts raw dates to **human-readable file naming format**
- Used in folder names and Excel file names

---

### `format_sheet_name(custom_date)`
```python
def format_sheet_name(custom_date):
    # Returns "9th July -2026"
```
- Formats the Excel **sheet tab name** for each day's data

---

### `sheet_exists(file_path, sheet_name)`
```python
def sheet_exists(file_path, sheet_name):
    wb = load_workbook(file_path)
    return sheet_name in wb.sheetnames
```
- Checks if a sheet already exists in the Excel file to **avoid overwriting existing data**

---

### `format_excel_file(excel_file_path, sheet_name)`
```python
def format_excel_file(excel_file_path, sheet_name):
    # Applies:
    # - Bold white text on dark blue (#366092) header row
    # - Auto-adjusted column widths (capped at 50 chars)
    # - Thin borders on all cells
    # - Center-aligned headers
```
- Applies professional styling to the Excel report:
  - **Header row**: Bold white font + dark blue fill + center alignment
  - **Column widths**: Auto-sized up to max 50 characters
  - **All cells**: Thin border on all 4 sides

---

### `get_or_rename_folder(base_path, start_date, end_date)`
```python
def get_or_rename_folder(base_path, start_date, end_date):
    # Looks for any folder starting with "POS Daily Report"
    # Renames it if date range has changed
    # Creates it if it doesn't exist
    # Returns the folder path
```
- Manages a **single rolling folder** for the entire reporting period
- As new days are added, the folder name is updated to reflect the expanded date range
  - E.g., `POS Daily Report (25th June to 8th July)` → `POS Daily Report (25th June to 9th July)`

---

### `rename_existing_file_if_needed(folder_path, circle, to_date, custom_date)`
```python
def rename_existing_file_if_needed(folder_path, circle, to_date, custom_date):
    # Uses regex to find: CIRCLE_POS Daily Report (Xth Month to Yth Month).xlsx
    # Renames it to the updated date range if needed
    # Returns the file path (new or existing)
```
- Each **circle** gets its own Excel file (e.g., `Dhaka_POS Daily Report (25th June to 9th July).xlsx`)
- On each run, the file is renamed to include the new date
- Uses regex: `^\{circle\}_POS Daily Report \(\d{1,2}(?:st|nd|rd|th)? \w+ to \d{1,2}(?:st|nd|rd|th)? \w+\)\.xlsx$`

---

## 💾 5. `save_the_data()` — Main Data Save Function

```python
def save_the_data():
    df = fetch_raw_data(custom_date, pos_daily_rpt)
    
    folder_path = get_or_rename_folder(base_folder, to_date, custom_date)

    for circle in df['Circle'].unique():
        circle_df = df[df['Circle'] == circle]
        file_path = rename_existing_file_if_needed(folder_path, circle, to_date, custom_date)
        sheet_name = format_sheet_name(custom_date)

        if sheet_exists(file_path, sheet_name):
            # Skip — already processed
            continue

        if not os.path.exists(file_path):
            wb = Workbook(); wb.save(file_path)  # Create empty workbook

        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
            circle_df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        format_excel_file(file_path, sheet_name)
```

**Flow:**
1. Fetch all data from DB for the date
2. Get/create/rename the report folder
3. Loop through each unique **Circle** in the data
4. Get/rename the circle's Excel file
5. Check if today's sheet already exists → skip if yes (idempotent)
6. Create new workbook if file doesn't exist
7. Append the new sheet with today's data
8. Apply Excel formatting

---

## 🗜️ 6. `zip_files(folder_path, zip_output_path)`

```python
def zip_files(folder_path, zip_output_path):
    with zipfile.ZipFile(zip_output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                zipf.write(file_path, arcname=os.path.relpath(file_path, folder_path))
```

- Walks through the entire report folder
- Compresses all Excel files into a single `.zip` archive using **DEFLATE compression**
- The zip is saved alongside the folder with the same name

---

## 📧 7. `send_email_with_attachments()` — Email Delivery

```python
def send_email_with_attachments(subject, body, to_emails, cc_emails, attachments):
```

This function has **two stages**:

### Stage 1: Google Drive Upload
```python
gauth = GoogleAuth()
gauth.LoadClientConfigFile("client_secrets.json")
gauth.LoadCredentialsFile("mycreds.txt")
# ... token refresh logic ...
drive = GoogleDrive(gauth)

for attachment_path in attachments:
    file_drive = drive.CreateFile({'title': os.path.basename(attachment_path)})
    file_drive.SetContentFile(attachment_path)
    file_drive.Upload()
    file_drive.InsertPermission({'type': 'anyone', 'value': 'anyone', 'role': 'reader'})
    drive_links.append(file_drive['alternateLink'])
```
- Authenticates with **Google Drive using OAuth2** (`pydrive2`)
- Uploads the `.zip` attachment to Google Drive
- Makes it **publicly readable** (anyone with link)
- Collects the shareable Drive link

> **Why Drive?** Gmail has a 25MB attachment limit. Uploading to Drive and sharing a link avoids size restrictions.

### Stage 2: HTML Email via SMTP
```python
body = f"""
<p>Dear Concern,</p>
<p>Please find the <b>POS Daily Reports circle wise ({date_range})</b> herewith:</p>
<p><a href="{drive_link}">{drive_link}</a></p>
<p><i>[Automated email. Do not reply.]</i><br>Regards,<br><b>Cockpit GLM System</b></p>
<hr><small>Disclaimer: ...</small>
"""

msg = MIMEMultipart()
msg['From'] = group_mail
msg['Subject'] = subject
msg['To'] = ", ".join(to_emails)
msg['Cc'] = ", ".join(cc_emails)
msg.attach(MIMEText(body, 'html'))

with smtplib.SMTP('smtp.gmail.com', 587) as server:
    server.starttls()
    server.login(from_email, email_password)
    server.sendmail(group_mail, all_recipients, msg.as_string())
```
- Builds a **professional HTML email** with:
  - Drive links as clickable hyperlinks
  - Legal disclaimer in small text
  - Automated system signature
- Sends via **Gmail SMTP** on port 587 with TLS
- All credentials are blanked out in the committed version

---

## ▶️ 8. `__main__` — Entry Point

```python
if __name__ == "__main__":
    custom_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
    to_date = '2026-06-25'
    formatted_range = format_date_for_file_name(to_date, custom_date)
    custom_date_filename = datetime.strptime(custom_date, '%Y-%m-%d').strftime('%d %B %Y')

    save_the_data()

    base_folder = r'E:/GP/Report/POS Daily Report/Daily POS Visit Reports'
    folder_name = f"POS Daily Report ({formatted_range})"
    folder_path = os.path.join(base_folder, folder_name)
    zip_path = folder_path + ".zip"

    zip_files(folder_path, zip_path)

    subject = f"POS Daily Report ({formatted_range})"
    send_email_with_attachments(subject, None, to_emails, cc_emails, [zip_path])
```

**Execution Order:**
1. Set `custom_date` = **yesterday's date** (auto-calculated)
2. Set `to_date` = the **start date** of the reporting period (manually configured)
3. Call `save_the_data()` → fetch → process → save Excel files
4. Zip the entire report folder
5. Upload zip to Google Drive + send email with Drive link

**Key Hardcoded Config:**
| Variable | Value | Purpose |
|---|---|---|
| `custom_date` | Yesterday's date (auto) | Date to pull data for |
| `to_date` | `'2026-06-25'` (manual) | Start of current reporting cycle |
| `base_folder` | `E:/GP/Report/...` | Output directory on Windows machine |

---

## 🏗️ Overall Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│                    SCRIPT EXECUTION FLOW                        │
│                                                                 │
│  1. CONFIGURE DATE                                              │
│     custom_date = yesterday | to_date = cycle start            │
│                                                                 │
│  2. FETCH DATA FROM PostgreSQL                                  │
│     pos_daily_rpt(custom_date) ──► SQLAlchemy ──► DataFrame    │
│     ┌─────────────────────────────────────────────────────┐    │
│     │  mv_data   → Visit records (latest per outlet/ME)   │    │
│     │  user_data → ME details + DH + geographic hierarchy │    │
│     │  survey_data → Survey Q&A (pivoted to columns)      │    │
│     │  posm_data  → POSM material counts (pivoted)        │    │
│     └─────────────────────────────────────────────────────┘    │
│                                                                 │
│  3. SPLIT BY CIRCLE & SAVE TO EXCEL                            │
│     For each Circle in data:                                    │
│       ├── Get/Rename report folder                              │
│       ├── Get/Rename circle Excel file                          │
│       ├── Append new sheet (skips if exists)                   │
│       └── Format: Headers, Borders, Column Widths              │
│                                                                 │
│  4. ZIP THE FOLDER                                              │
│     POS Daily Report (25th June to 9th July).zip               │
│                                                                 │
│  5. UPLOAD TO GOOGLE DRIVE                                      │
│     OAuth2 → Upload zip → Set public permission → Get link     │
│                                                                 │
│  6. SEND EMAIL                                                  │
│     Gmail SMTP → HTML email with Drive link → Recipients       │
└─────────────────────────────────────────────────────────────────┘
```

---

## ⚠️ Important Notes & Security

| Item | Status | Note |
|---|---|---|
| DB Password | `""` blanked | Must be set before running |
| DB Host | `""` blanked | AWS RDS endpoint needed |
| Gmail `from_email` | `""` blanked | Sender email needed |
| Gmail `email_password` | `""` blanked | App password (not account password) |
| `group_mail` | `""` blanked | Group/team mailbox address |
| `to_emails` | `[]` empty list | Recipient list must be populated |
| `cc_emails` | `[]` empty list | CC list must be populated |
| `client_secrets.json` | External file | Google OAuth client credentials |
| `mycreds.txt` | External file | Saved Google Drive tokens |
| `to_date` | Hardcoded | Update at start of each new reporting cycle |
| `base_folder` | Windows path | Must exist on the machine running the script |

---

## 📂 Output File Structure

```
E:/GP/Report/POS Daily Report/Daily POS Visit Reports/
└── POS Daily Report (25th June to 9th July)/
    ├── Dhaka_POS Daily Report (25th June to 9th July).xlsx
    │   ├── Sheet: "25th June -2026"
    │   ├── Sheet: "26th June -2026"
    │   └── ... (one sheet per day)
    ├── Chittagong_POS Daily Report (25th June to 9th July).xlsx
    ├── Sylhet_POS Daily Report (25th June to 9th July).xlsx
    └── [Other circles...]
POS Daily Report (25th June to 9th July).zip   ← Emailed via Drive link
```

---

## 🔁 How to Run Daily

This script is designed to run **once per day**, typically scheduled via:
- **Windows Task Scheduler** (since `base_folder` is a Windows path `E:/GP/...`)
- Or manually executed each morning

**Each daily run:**
1. Picks up yesterday's data automatically
2. Updates existing Excel files with a new sheet
3. Updates folder and file names to reflect the new date range
4. Zips and emails the updated reports

---

*Generated explanation for `pos_daily_report_automation.py` — Grameenphone POS Visit Report Automation System*
