# POS Daily Report — GUI Automation (Download Only / Download & Send)

## Goal
Build a new Tkinter GUI script that reuses the **old script's accurate data + file generation**
(`pos_daily_report_automation.py`) and the **new script's circle-wise email sending + recipient
mapping** (`pos_daily_report_automation_split_circlewise.py`), exposing exactly two buttons:
**Download Only** and **Download & Send**, with a progress percentage, file-conflict safety,
and a visible log.

## Context from `table.tsv` (What I Need column)
| Topic | Decision |
|---|---|
| Date Handling | Keep old process. GUI date fields (Start = `2026-07-25`, End/Report = `2026-07-28`) replace manual in-code edits. Old SQL (`visit_date = custom_date`) stays — only End Date drives the data. |
| Report Download | Download only = generate circle files to selected folder. |
| Password Protection | No password-protected files are generated. Out of scope (neither script implements it). |
| Report Save Location | User picks destination folder in the GUI (Browse). |
| Email Recipient Selection | User picks the recipient mapping `.xlsx` in the GUI (Browse). |
| File Generation | Keep old behavior: per-circle `.xlsx` files **+ one ZIP** of the folder. |
| Email Sending | Implement new-script behavior in the new GUI script: one email per circle, matched by circle name → sheet name. |
| Recipient Mapping | Add matching logic: each sheet in the recipient `.xlsx` = one circle; columns `To`/`Cc`; all non-empty cells across rows collected. |
| Mail Distribution | Each circle receives its own report via a separate email. Combined ZIP email to hardcoded recipients is **dropped**. |
| Visit Data Accuracy | Base on the old script's query (accurate, complete visit data). |
| User Interface | New GUI **supersedes** the old "no interface" note — this is the requested deliverable. |

## Architecture
- **Create new file** `pos_daily_report_automation_gui.py` next to the existing scripts.
- Do **not** modify `pos_daily_report_automation.py` or `pos_daily_report_automation_split_circlewise.py`.
- All heavy work runs in a **background thread**; GUI updates via `root.after()` (pattern already proven in the new script).

## Code to keep (ported/unchanged from old script)
- `pos_daily_rpt(custom_date)` — exact SQL, `fetch_raw_data(custom_date, ...)`.
- `get_day_with_suffix`, `format_date_for_file_name(to_date, custom_date)`, `format_sheet_name(custom_date)` (old signatures — single report-date sheet per file).
- `sheet_exists`, `format_excel_file`, `get_or_rename_folder`, `rename_existing_file_if_needed`, `zip_files`.
- DB credentials/engine, imports (keep top-level `pydrive` imports exactly as the old script — they load in the target env).

## Code to port from new script (adapted)
- `is_file_locked(file_path)` — pre-write lock check (prevents Bad CRC-32/corruption).
- `load_email_recipients(xlsx_path, log)` — sheet-per-circle → `{"CIRCLE_UPPER": {"to": [...], "cc": [...]}}`; header columns `to`/`cc` (case-insensitive); dedupe emails.
- `send_circle_email(circle, file_path, to_emails, cc_emails, formatted_range, log)` — subject `"{circle} POS Daily Report ({range})"`, uploads the single circle file to Google Drive (`pydrive2`, `client_secrets.json` + `mycreds.txt`), inserts "anyone/reader" permission, **attaches the `.xlsx` directly** (≤ 24 MB, else Drive-link-only body), sends via `smtplib` (Gmail 587/TLS).
  - **Fix required:** use the working password constant from the new script (`email_password = "ujmt gpgi ctbx dohe"`); the old script's `email_password` was commented out and would raise `NameError`.
  - `from_email = "shadman.sayeid@v2.ltd"`, `group_mail = "cockpit.glm@v2.ltd"` (From header).
- Threading + button-disable/enable + log-window patterns.

## GUI Spec (Tkinter, ~640x560)
- **Start Date** field (default `2026-07-25`), **End Date / Report Date** field (default `2026-07-28`) — validated against `^\d{4}-\d{2}-\d{2}$`.
- **Recipient mapping Excel file**: read-only entry + "Browse…" (`filedialog.askopenfilename`, `*.xlsx *.xls`). Only required for Download & Send.
- **Destination folder**: read-only entry + "Choose Folder…" (`filedialog.askdirectory`); auto-create if missing.
- **Progress bar** (`ttk.Progressbar`, 0–100) + percentage label ("Downloading circle files… 42%").
- **Buttons**: `Download Only` (primary blue), `Download & Send` (green). Disabled while a job runs.
- **Log**: `scrolledtext.ScrolledText`, read-only, auto-scroll; per-line timestamps.
- **Status label** at top (idle / running / done / error).

## Data Flow
### Download Only
1. Validate dates + destination folder (error dialog on failure).
2. Background thread:
   - Fetch data (10%).
   - For each circle: resolve file path (rename stale range file if needed), lock-check, skip if target sheet already exists, else write + format sheet (10%→80%, equal steps per circle).
   - Zip the date-range folder (90%).
   - Post-write verify each written file (reopen; confirm sheet exists and row count == circle rows) (100%).
3. Re-enable buttons; status "Download complete — N files + ZIP in <folder>".

### Download & Send
1. Validate dates, destination folder, **and** recipient file (error dialog on failure).
2. Download phase — same as above, but **reuse** already-generated files for the date range if present (no re-fetch needed when sheets already exist).
3. Load recipient mapping (`load_email_recipients`).
4. **Confirmation dialog** (`messagebox.askyesno`): "Send emails for N circle(s): <names>?" — circles with no recipients are listed and skipped.
5. Background thread: for each circle with recipients → upload to Drive + attach + SMTP send; per-circle progress to 100%; failures logged and counted; other circles continue.
6. Summary in log: "Emails sent: X, skipped: Y".

## Progress Reporting
- Weighted phases: Fetch 10%, circle files 10–80% (even split), ZIP 90%, Send loop 90–100% (even split per circle).
- Worker thread never touches widgets directly; it sends progress via `root.after(0, ...)` (or a thread-safe callback that schedules the GUI update).
- Percentage label updates with phase text so the user sees remaining work.

## Conflict & Correctness Rules (explicit requirements)
1. **Never write to a locked file**: `is_file_locked()` → log "open in Excel, skipping", continue others.
2. **No duplicate/conflicting filenames**: reuse `rename_existing_file_if_needed`; if a same-prefix file with an older range name exists, rename it to the current range instead of creating a new one.
3. **No partial/duplicate sheets**: if the report-date sheet already exists in the circle file, skip rewrite and treat the circle as downloaded.
4. **No append-mode corruption**: write a fresh sheet only into an existing workbook whose other sheets are unrelated; wrap in try/except catching `BadZipFile`, `PermissionError`, generic `Exception` (message: close file in Excel and re-run).
5. **ZIP only the current date-range folder** (old `zip_files` walks that folder only).
6. **Emails**: dedupe To/Cc; skip empty To; loose email-format check (contains `@`); only send when mapping has To recipients for the circle.
7. **GUI stays responsive**: all network/DB/file work in a daemon thread; buttons disabled during jobs.

## Failure Modes & Handling
- No data fetched → stop, error dialog, no files.
- Recipient file missing/unreadable → error for Download & Send only.
- Circle with no matching sheet / no To → skip email, log warning (file still downloaded).
- Attachment > 24 MB → Drive link only body.
- SMTP/Drive auth failure → log error for that circle, continue; final summary lists failures.
- Date format invalid → dialog, no work starts.

## Validation Plan
- `python -m py_compile pos_daily_report_automation_gui.py`.
- Add a module-level `DRY_RUN = False` constant; when `True`, the send phase only logs what would be emailed (no Drive/SMTP) — used for testing.
- Test 1 (no DB/network needed): launch GUI, verify fields/buttons render, date validation works, buttons disable during a simulated run (DRY_RUN + short job).
- Test 2 (real): click **Download Only** with a small date range → verify per-circle `.xlsx` files (correct sheet, formatted header, row count matches DB) + ZIP created, no duplicate files, locked-file skip works (open one file in Excel first).
- Test 3 (real, single circle): point mapping at a test file containing one circle sheet with one test email; **Download & Send** → confirm dialog lists 1 circle; verify one email with the circle's `.xlsx` attached + Drive link.
- Test 4: repeat Download & Send for the same range → files reused, no re-send prompt duplication, no sheet duplication.

## Rollout / Migration
- No DB, config, or schema changes. Requires `client_secrets.json` + `mycreds.txt` next to the script (same as existing scripts) for Drive upload.
- Run from the project folder on the Windows machine (`E:` drive paths already in use).
- Existing scripts remain untouched and usable.

## Out of Scope
- Password-protected report files (table row 3 behavior).
- Any new DB queries, dashboards, or web modules.
- Modifying the existing two scripts.
