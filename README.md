# Loan Shoppee — Customer Tracker

A single-user web app for managing loan-customer records. Data lives in **real
`.xlsx` files** on the server — every change is written straight back to the
workbook, so you always have a plain Excel file you can open, download, or back
up. No database required.

- **Frontend:** vanilla JavaScript SPA (no build step)
- **Backend:** Flask + openpyxl + pandas + `formulas`
- **Data:** `.xlsx` workbooks on disk

---

## Table of contents

1. [Quick start](#quick-start)
2. [Login & security](#login--security)
3. [Workbooks & sheets](#workbooks--sheets)
4. [Viewing records](#viewing-records)
5. [Adding, editing & deleting records](#adding-editing--deleting-records)
6. [Flags](#flags)
7. [Duplicate detection](#duplicate-detection)
8. [Column types](#column-types)
9. [Amount columns & INR formatting](#amount-columns--inr-formatting)
10. [Formulas](#formulas)
11. [Column totals](#column-totals)
12. [Managing columns](#managing-columns)
13. [Changes & backups](#changes--backups)
14. [Download & upload](#download--upload)
15. [Where data is stored](#where-data-is-stored)
16. [Configuration (environment variables)](#configuration-environment-variables)
17. [Deploying to PythonAnywhere](#deploying-to-pythonanywhere)
18. [Security notes](#security-notes)

---

## Quick start

```bash
cd webapp
.venv/bin/python scripts/create_password_hash.py   # prints a PASSWORD_HASH value
export PASSWORD_HASH='scrypt:32768:8:1...'          # paste the printed value (single quotes!)
export SECRET_KEY='some-random-long-string'          # optional; auto-generated if omitted
.venv/bin/python app.py
```

Then open **http://127.0.0.1:5000** and sign in with the password you hashed.

> The hash contains `$` characters — always wrap it in single quotes in a shell.

### First-run steps

1. **Upload a workbook** (`Upload` in the top bar) — or add sheets from scratch.
2. Pick the workbook and sheet from the dropdowns.
3. Add your first record with **+ Add Record**.

---

## Login & security

- Single password login (`PASSWORD_HASH`, generated with `scripts/create_password_hash.py`).
- Sessions expire after **60 minutes** idle or **30 days** absolute.
- **CSRF protection** on every write; session cookie is `HttpOnly` + `SameSite=Strict`.
- **Login rate limiting**: 10 tries per 15 minutes, then 20 per hour.
- General API rate limit: 60 requests/minute by default.
- Security headers (`X-Frame-Options: DENY`, `nosniff`, `no-store`, CSP) on all responses.

---

## Workbooks & sheets

A **workbook** is one `.xlsx` file. A **sheet** is one tab inside it.

- **Switch workbook/sheet** — dropdowns in the toolbar.
- **Add a sheet** — the `+` button next to the Sheet dropdown. You can define
  the new sheet's columns and their types up front (optional — defaults to
  `DATE` / `CUSTOMER NAME`).
- The app remembers the **last workbook and sheet** you had open and returns to it next time.

### Layout of a sheet

The app understands three parts of your Excel sheet:

| Row | Purpose |
| --- | --- |
| Row 1 | Column headers (e.g. `DATE`, `CUSTOMER NAME`, `AMOUNT`) |
| Row 2 | Optional **template row** — formulas/defaults copied into every new record |
| Row 3+ | Your records |

Sheets with a title row above the header are detected automatically.

---

## Viewing records

- **Search** — type in the search box; the table filters as you type (300 ms debounce).
- **Sort** — click a column header to sort ascending, click again for descending.
- **Flagged only** — check the box to show only flagged rows.
- **Totals footer** — a sticky row at the bottom (see [Column totals](#column-totals)).
- Date columns show `Mon DD`; **Amount** columns show Indian-style commas (see below).
- A banner appears when any formula values could not be calculated.

---

## Adding, editing & deleting records

### Add a record
- Click **+ Add Record**, fill the form, click **Save**.
- **New records go to** — choose *Bottom* (default) or *Top*. This is remembered per workbook.
- Blank numeric columns default to `0` on save.

### Edit a record
- Click any cell (except formula cells) and type; press **Enter** or click away to save, **Esc** to cancel.
- Date columns open a date picker. If the edit triggers a duplicate warning, you can confirm it anyway.
- Formula columns are calculated — you can't edit them directly.

### Delete records
- Tick the checkboxes on the left of one or more rows, then click **Delete**.
- Deleting can't be undone from the table, but a **backup snapshot is saved first**
  (see [Changes & backups](#changes--backups)).

---

## Flags

Flags mark records for attention (e.g. follow-ups, pending documents).

- Click the **flag icon** in a row to toggle it.
- Select multiple rows and click **Flag** to flag them all at once.
- Use **Flagged only** to focus on them.
- The flag is stored in a `FLAGGED` column in the workbook (added automatically
  the first time you use it) and survives reloads.

---

## Duplicate detection

When you add or edit a record, the app checks existing rows for near-identical values
and warns you before saving.

- **Automatic** (default): scans columns whose names look like identifiers —
  `CUSTOMER NAME`, `CONTACT NUMBER`, `CONTACT`, `PHONE`, `CAR`, `VEHICLE`, `MODEL`,
  `EMAIL`, `ACCOUNT`.
- **Manual**: in **Format → Advanced: duplicate detection columns**, check specific
  columns to only flag duplicates on those.
- On a match you get a dialog listing the similar rows. Choose **Continue** to
  save anyway (or **Cancel** to go back and fix it).

---

## Column types

Each column has a type that drives how it's entered, aligned, and formatted:

| Type | Behavior |
| --- | --- |
| **Text** | Plain text (default). |
| **Number** | Right-aligned, numeric input, counted/summed in totals. Contact or bank numbers are usually fine as **Number** — they stay unformatted. |
| **Date** | Date picker when entering/editing; displayed as `Mon DD`. |
| **Amount** | Treated like a number, but **displayed with Indian-style comma grouping** (e.g. `12,34,567.89`). No other numbers get commas. |

Column types can be set when **creating a sheet**, when **adding a column**, and changed
any time in **Format → Manage columns** (see below).

---

## Amount columns & INR formatting

Amount columns make money columns readable with Indian number-grouping
(lakhs/crores style):

| Stored value | Displayed |
| --- | --- |
| `1234` | `1,234` |
| `123456` | `1,23,456` |
| `1234567.89` | `12,34,567.89` |
| `100000000` | `10,00,00,000` |

- Commas are **display-only** — the value stored in Excel stays a plain number.
- Decimals are shown exactly as stored (no rounding).
- Only **Amount** columns get commas. Normal numbers (contact numbers, bank
  numbers) stay unformatted.
- Amount totals in the footer use the same grouping.

**To make a column an Amount column:** open **Format → Manage columns** and change
that column's type dropdown to **Amount**.

---

## Formulas

Formula columns compute values automatically for every new record — handy for
tax, interest, or running balances.

- Open **Format**; the **Formulas** section lists every column with its Excel
  letter next to it.
- Write an Excel-style formula in a column, e.g. `=D2*1.18` to add 18% GST to
  whatever is in column D of that row.
- Column letters are absolute references to that row's cells. Formula cells
  can't be edited directly and are recalculated whenever an edit affects them.

### Supported functions

```
SUM   SUMIF   COUNT   COUNTA   COUNTBLANK   COUNTIF
AVERAGE   AVERAGEA   MEDIAN   MAX   MIN
```

- `SUMIF` / `COUNTIF` accept a criteria (e.g. `=SUMIF(D2:D100, ">1000")`).
- Ranges and criteria use the same Excel letter/number references.
- If a value can't be calculated for some rows, a warning banner lists them.

A formula column can also be set to **Amount** type so its result shows INR commas.

---

## Column totals

A **sticky footer row** (pinned above the scrollbar) shows a total per column.

- Open **Format → Column totals** and pick a mode per column:
  - **Off** — no total
  - **All rows** — every record in the sheet
  - **Visible** — records matching the current search + flagged filter
  - **Flagged** — only flagged records
  - **Visible + Flagged** — both filters applied
- Numeric/Amount columns show a **sum**; other columns show a **count**.
- The footer updates live with every add, edit, delete, flag, search, and filter change.

---

## Managing columns

**Format → Manage columns** lets you:

- **Add** a column (name + type).
- **Rename** a column — data is preserved and every reference updated.
- **Delete** a column — removes it and its data from every row (a backup is created first).
- **Change the type** of an existing column via its dropdown — e.g. switch a
  number column to **Amount**.

> Note: a column full of numbers will keep behaving as a number even if you set
> its type to Text — type changes to Amount/Date/Number stick, but inference can
> override an attempt to make numeric data "Text".

---

## Changes & backups

Every change (add/edit/delete/flag/formula/column action) writes a new snapshot.

- Click **Changes** in the top bar.
- **Recent save points** — each snapshot is labeled from the audit log
  (e.g. `ADD_ROW — Sheet1 — row 4`). Click **Revert to this version** to restore it.
  Reverting saves the current version as a new snapshot first, so it's undoable.
- **Raw backup files** — the underlying `.bak` files, one per save point; download them directly.
- Snapshots are kept per workbook; the **last 5** are retained.

---

## Download & upload

- **Download** — saves the current workbook (`.xlsx`) to your computer, exactly as it exists on the server.
- **Upload** — replace a workbook by uploading an `.xlsx` file (max **5 MB**).
  - Uploading over an existing workbook creates a backup of the old one first.
  - Invalid or non-`.xlsx` files are rejected.

---

## Where data is stored

Inside `webapp/data/` (or wherever `DATA_DIR` points):

| Path | Contents |
| --- | --- |
| `*.xlsx` | Your workbooks |
| `settings.json` | App state: append direction, duplicate-check columns, column types, totals, last-opened pointer |
| `.bak/` | Automatic `.bak` snapshots (last 5 per workbook) |
| `audit.log` | Timestamped change log (rotated at 5 MB); bank/account values are redacted |

Back up `data/` and your workbooks are safe.

---

## Configuration (environment variables)

| Variable | Default | Purpose |
| --- | --- | --- |
| `PASSWORD_HASH` | *(none — login disabled)* | Password hash from `scripts/create_password_hash.py` |
| `SECRET_KEY` | random per boot | Session signing key; set it so sessions survive restarts |
| `DATA_DIR` | `webapp/data` | Where workbooks/settings live |
| `PORT` | `5000` | HTTP port |
| `COOKIE_SECURE` | `0` | Set `1` when served over HTTPS |
| `TIMEZONE` | `Asia/Kolkata` | Audit timestamps |
| `MAX_ROWS` | `20000` | Row cap |
| `LOGIN_RATE_LIMIT_BURST` | `10 per 15 minutes` | Login burst limit |
| `LOGIN_RATE_LIMIT_SUSTAINED` | `20 per hour` | Login sustained limit |
| `DEFAULT_RATE_LIMIT` | `60 per minute` | General API limit |

---

## Deploying to PythonAnywhere

1. Upload the `webapp/` contents to your account (or clone the repo).
2. Create a virtualenv and install dependencies:

   ```bash
   pip install --no-deps formulas==1.3.4
   pip install -r requirements.txt
   ```

   `formulas` declares `scipy` as a dependency, but it isn't actually needed —
   install it *without deps* as above (or uninstall scipy afterwards). The
   `pip check` warning about scipy is harmless.

3. Point your **WSGI** configuration at the app:

   ```python
   from app import app as application
   ```

4. Set `PASSWORD_HASH`, `SECRET_KEY`, and `COOKIE_SECURE=1` (HTTPS) as
   PythonAnywhere environment variables, and configure a static file mapping for
   `/static/` → `webapp/static/`.

5. Restart the web app. The app stores data under `webapp/data/` (or whatever
   `DATA_DIR` points to).

---

## Security notes

- Workbook and sheet names are validated to prevent path traversal.
- Only `.xlsx` files up to 5 MB are accepted; uploads are validated as real
  zip/xlsx archives.
- Bank and account values are **redacted** from the audit log.
- The `FLAGGED` column is reserved and can't be renamed or deleted.
- CSRF tokens protect every mutating request; responses carry `no-store` caching
  headers so sensitive data is never cached.