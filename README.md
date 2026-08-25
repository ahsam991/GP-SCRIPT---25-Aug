# GP SCRIPT — POS Daily Report Automation (GUI)

Automates the daily POS visit report for all circles: **generates circle-wise Excel
reports**, **zips them**, and **sends one email per circle** with that circle's own report,
using a recipient mapping file (Excel or CSV).

The data comes from the **exact same database query** as the original reference script, so
the report content is unchanged and accurate. Only the delivery is improved: each circle now
gets its own report instead of one combined ZIP email.

---

## 📁 Folder Structure

```
GP SCRIPT/
├── run.py                                  → LAUNCHER: checks/installs libraries, then opens the GUI
├── pos_daily_report_automation_gui.py      → main automation (GUI application)
├── requirements.txt                        → Python libraries required by the automation
├── POS_Daily_Report_Email_list_Testing.xlsx→ sample recipient mapping file (sheet per circle)
├── client_secrets.json                     → Google Drive OAuth client (required for sending)
├── mycreds.txt                             → saved Google Drive login (created on first send)
├── Report/                                 → sample generated output (per-circle files + ZIP)
└── OLD files/                              → reference/backup only — NOT needed to run
    ├── pos_daily_report_automation.py                  (original single-ZIP script)
    ├── pos_daily_report_automation_split_circlewise.py (GUI reference implementation)
    ├── table.tsv                                       (requirements mapping table)
    ├── pos_daily_report_gui_plan.md                    (design plan)
    └── Daily POS Visit Report Automation/              (older production copy)
```

> Everything inside `OLD files/` is kept for reference and rollback. Do not modify it;
> the automation never reads from there.

---

## 🚀 Quick Start

### 1. Install (automatic)

Just run the launcher — it checks which libraries are missing and installs the latest
versions automatically:

```bash
python run.py
```

Or install manually:

```bash
python -m pip install -r requirements.txt
```

Required libraries: `pandas`, `openpyxl`, `SQLAlchemy`, `psycopg2-binary`, `pydrive2`, `xlrd`.

### 2. Run the GUI

```bash
python run.py
```

or directly:

```bash
python pos_daily_report_automation_gui.py
```

> On Windows, double-click `run.py` (or run `python run.py`) — the console stays open so you
> can see progress and errors.

---

## 🖥️ Using the GUI

| Field | What it does |
|---|---|
| **Start Date** | First day of the report range, used for the file/folder name (e.g. `2026-07-25`). |
| **End Date** | The report date — **this is the day whose data is fetched** (e.g. `2026-07-28`). |
| **Recipient mapping file** | The Excel/CSV file containing circle-wise email lists (Browse…). |
| **Destination folder** | Where the generated report files are saved (Choose Folder…). |

### Buttons

- **Download Only** — fetches data, writes one formatted `.xlsx` per circle, then creates the
  ZIP of the folder. Sends nothing.
- **Download & Send** — does the same download, then loads the recipient mapping and asks for
  confirmation before emailing each circle.

A progress bar shows the percentage and the current step. Every action is logged with a timestamp.

---

## 📧 Recipient Mapping File (Excel or CSV)

### Option A — Excel (`.xlsx` / `.xls`)
- **One worksheet per Circle.** The sheet name must match the circle name (case-insensitive).
- Row 1 must contain columns **`To`** and **`Cc`**.
- Every non-empty email in those columns (any number of rows) is collected.

Example — workbook with sheets `DHAKA`, `CHITTAGONG`, … each looking like:

| To | Cc |
|----|----|
| a@example.com | cc1@example.com |
| b@example.com | cc2@example.com |

### Option B — CSV
- Columns: **`Circle`**, **`To`**, **`Cc`** (header row required).
- One row per recipient; a cell can hold several emails separated by comma/semicolon/newline.

Example `recipients.csv`:

```csv
Circle,To,Cc
Dhaka,a@example.com; b@example.com,cc1@example.com
Chattogram,c@example.com,
Sylhet,d@example.com,
```

> Circle matching is case-insensitive (e.g. `dhaka` matches sheet `DHAKA`). Circles that have
> **no To recipients** are skipped with a warning — their report is still downloaded.

---

## ✉️ What Each Email Contains

- Subject: `{Circle} POS Daily Report ({date range})`
- Body: the circle's report name + the Google Drive share link
- Attachment: that circle's own `.xlsx` file (attached directly, up to 24 MB)
  - If the file is larger than 24 MB, only the Drive link is sent.
- One email is sent per circle to that circle's `To` (+ `Cc`) recipients only.

### First-time Google Drive login
Sending uploads the file to Google Drive using `client_secrets.json`. On the first send the
browser opens for Google login; afterwards `mycreds.txt` is reused automatically.

---

## 🧪 Testing Without Sending Real Emails

Open `pos_daily_report_automation_gui.py` and set near the top:

```python
DRY_RUN = True
```

With `DRY_RUN = True`, the Send step only prints what would be emailed
(recipients + file) and does **not** upload or send anything. Set it back to `False` for real
sending.

---

## 🗂️ Generated Output

```
<Destination Folder>/
└── POS Daily Report (25th July to 28th July)/
    ├── DHAKA_POS Daily Report (25th July to 28th July).xlsx
    ├── CHITTAGONG_POS Daily Report (25th July to 28th July).xlsx
    ├── ... (one file per circle)
    └── (then zipped) → POS Daily Report (25th July to 28th July).zip
```

Each circle file contains one sheet named after the report date (e.g. `28th July -2026`),
formatted with headers, borders and auto-width. Every written file is **verified** (sheet
exists + row count matches) before being marked done.

---

## 🛠️ Safety & Conflict Rules

- If a report file is **already downloaded** for that date, it is reused — no duplicate sheets.
- If a file is **open in Excel**, that circle is skipped and you are told to close it; other
  circles continue.
- No blank default sheets — files start directly with the report sheet.
- The ZIP always contains only the current date-range folder.

---

## ❓ Troubleshooting

| Problem | Solution |
|---|---|
| `ModuleNotFoundError: No module named 'openpyxl'` | Run `python run.py` (auto-installs), or check `which python` and make sure it is the same interpreter that has the packages. |
| Sending fails / no Drive login | Make sure `client_secrets.json` is in the same folder as `pos_daily_report_automation_gui.py` (the `GP SCRIPT` folder) and run from there. |
| Circle skipped when sending | That circle has no `To` emails in the mapping file, or its name doesn't match any sheet/row. |
| File skipped during download | The file is open in Excel (or another program). Close it and re-run. |
| `Bad CRC-32` / corrupted file | A previous write was interrupted (often Excel had the file open). Delete the file and re-run. |
| Emails go to wrong circle | Check the sheet names in the mapping file match the circle names exactly (case doesn't matter). |

---

## 📌 Notes

- **Data accuracy:** the SQL query is byte-for-byte identical to the original reference script
  (`OLD files/pos_daily_report_automation.py`) — same single report date, same deduplication.
  No data limit or pagination is applied.
- **Passwords/credentials** are intentionally blanked in the sanitized copy inside
  `OLD files/Daily POS Visit Report Automation/`. Do not commit secrets.
