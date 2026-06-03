# CoreCast Invoice Automation — Claude Code Handoff

## What this project does

Automates weekly invoice validation for CoreCast (Perth precast concrete company).
Subcontractors email invoices each week. This system:
1. Reads those invoices from Outlook (Microsoft Graph API)
2. Reads the weekly timesheet export from Connecteam (manually exported XLSX)
3. Matches each invoice sender's claimed hours against Connecteam hours
4. Emails Tom a report flagging any mismatches

It also files email attachments into SharePoint by job keyword — see "Email
attachment filing" below.

## Why we switched from Connecteam API to XLSX export

Connecteam's time-activities API requires their Expert plan ($99/month). CoreCast is on
the Small Business plan. The API authenticates fine and returns the time clock list, but
returns empty shift data. The workaround: Tom exports the weekly timesheet XLSX manually
from the Connecteam UI each week and drops it in a watch folder.

---

## Project structure

```
corecast-invoice-automation/
├── src/
│   ├── config.js                      # All env var config
│   ├── index.js                       # Invoice ingestion from Outlook (entry point)
│   ├── email/
│   │   ├── outlook.js                 # Microsoft Graph API client
│   │   ├── parser.js                  # Extract hours/amount/date from invoice emails
│   │   └── filter.js                  # Heuristics to identify invoice emails
│   ├── connecteam/
│   │   ├── timesheetParser.js         # ✅ WORKING — parse Connecteam XLSX export
│   │   └── client.js                  # ⚠️  Connecteam REST API client (limited by plan — not used in main flow)
│   └── storage/
│       └── invoices.js                # JSON file store for ingested invoices
├── scripts/
│   ├── test-timesheet-parser.js       # Test the XLSX parser against a real export
│   └── test-connecteam.js             # (Legacy) Connecteam API tests — ignore
├── data/
│   └── invoices.json                  # Persisted invoice records
├── .env                               # Local secrets (not committed)
├── .env.example                       # Template
└── package.json
```

---

## What needs to be built (the main task)

The core weekly matching flow does not exist yet. Here is exactly what needs to be built:

### New file: `src/matcher/matchWeek.js`

This is the main weekly job. It should:

1. **Find the timesheet XLSX** — scan a configurable watch folder (e.g. `./timesheets/`)
   for the most recent Connecteam export file matching the pattern:
   `timeclock-timesheet_overview_YYYY-MM-DD_YYYY-MM-DD.xlsx`

2. **Parse the XLSX** — call `parseTimesheetExport(filePath)` from `timesheetParser.js`.
   Returns: `{ weekStart, weekEnd, employees[], byName: { "aaron norris": 19.23 } }`

3. **Load invoices for that week** — from `InvoiceStore` in `storage/invoices.js`,
   filter invoices where `weekEnding` falls within the timesheet's week range.
   (Invoices are ingested separately by `src/index.js`)

4. **Match each invoice to timesheet hours** — for each invoice, call
   `lookupHours(invoice.sender, byName)` from `timesheetParser.js`.
   Compare `invoice.hoursClaimed` vs timesheet hours.

5. **Classify each match result** as one of:
   - ✅ `match` — hours within tolerance (±0.5h — allow for rounding)
   - ⚠️  `mismatch` — hours differ by more than tolerance
   - ❓ `not_found` — sender name not found in timesheet at all
   - ⏸️  `no_hours_in_timesheet` — found in timesheet but totalHours = 0 (not clocked in)
   - ❓ `invoice_hours_missing` — invoice was parsed but hours weren't extracted

6. **Send a summary email** to `tom@corecastconcrete.com.au` via Microsoft Graph API
   (reuse the `OutlookClient` / Graph client already in `outlook.js`).

7. **Update invoice statuses** in the JSON store — set matched invoices to `validated`,
   mismatches to `mismatch`, etc.

### New file: `src/matcher/sendReport.js`

Sends the weekly validation email. Should produce a clear plain-text + HTML email with:
- Week dates at the top
- A section for each subcontractor:
  - Invoice claimed: X hours / $Y
  - Connecteam shows: Z hours
  - Status: ✅ Match / ⚠️ MISMATCH (difference: N hours)
- Summary counts at the bottom
- Subject line: `[CoreCast] Timesheet Validation — Week of {weekStart}`
- Send FROM `accounts@corecastconcrete.com.au` TO `tom@corecastconcrete.com.au`

### New script: `scripts/run-weekly-match.js`

CLI entry point for the weekly job:
```
node scripts/run-weekly-match.js
node scripts/run-weekly-match.js --timesheet ./timesheets/timeclock-timesheet_overview_2026-04-27_2026-05-03.xlsx
```

### Update `src/config.js`

Add:
```js
timesheets: {
  watchFolder: process.env.TIMESHEETS_FOLDER || './timesheets',
},
alerts: {
  reportRecipient: process.env.REPORT_RECIPIENT || 'tom@corecastconcrete.com.au',
  reportSender: process.env.REPORT_SENDER || 'accounts@corecastconcrete.com.au',
},
matcher: {
  hoursTolerance: Number(process.env.HOURS_TOLERANCE || 0.5),
},
```

### Update `.env.example`

Add:
```
# Folder where Connecteam XLSX exports are dropped each week
TIMESHEETS_FOLDER=./timesheets

# Email report settings
REPORT_RECIPIENT=tom@corecastconcrete.com.au
REPORT_SENDER=accounts@corecastconcrete.com.au

# Tolerance in hours before flagging a mismatch (default 0.5)
HOURS_TOLERANCE=0.5
```

---

## Timesheet XLSX format (confirmed from real export)

File: `timeclock-timesheet_overview_2026-04-27_2026-05-03.xlsx`
Sheet: `All Employees` (also has one sheet per employee — use All Employees)

Columns (row 1 is header):
| Col | Name | Notes |
|-----|------|-------|
| A | First name | Only on first row per employee |
| B | Last name | Only on first row per employee |
| C | Type | "Shift", "Lunch break", etc |
| E | Start Date | datetime object |
| F | In | Clock-in time string "06:43" |
| I | Out | Clock-out time string "15:57" |
| M | Shift hours | Per-shift duration "09:14" |
| N | Daily total hours | Daily total "09:14" |
| O | **Weekly total hours** | ⭐ THIS IS THE ONE WE WANT — only on first row per employee |
| P | Total work hours | Same as weekly for most cases |

Hours format: `"HH:MM"` strings e.g. `"19:14"` = 19 hours 14 minutes.
Convert to decimal: `19 + 14/60 = 19.23h`

Employees with no shifts this week: appear as a row with name but no time values.

Real data from week 2026-04-27 to 2026-05-03:
| Employee | Raw | Decimal hours |
|----------|-----|---------------|
| Aaron Norris | 19:14 | 19.23h |
| Barsbaatar Binderiya | — | 0h (no shifts) |
| Jack Henderson | 33:57 | 33.95h |
| James Elliott | 37:41 | 37.68h |
| Josh Brennan | — | 0h (no shifts) |
| Orgil Khurelchuluun | 07:51 | 7.85h |
| Tom Bache | — | 0h (admin) |
| Zolbayar Munkhbaatar | 07:50 | 7.83h |

---

## Existing working code to reuse

### `src/connecteam/timesheetParser.js` — ✅ Complete and tested

```js
const { parseTimesheetExport, lookupHours } = require('./src/connecteam/timesheetParser');

// Parse the XLSX
const { weekStart, weekEnd, employees, byName } = parseTimesheetExport('/path/to/file.xlsx');
// byName = { "aaron norris": 19.23, "jack henderson": 33.95, ... }

// Look up a sender name (tolerant matching)
const hours = lookupHours("Aaron Norris", byName); // returns 19.23
const hours2 = lookupHours("Norris", byName);       // also returns 19.23 (last-name match)
const hours3 = lookupHours("Unknown", byName);      // returns null
```

### `src/email/outlook.js` — ✅ Complete

Microsoft Graph client. To send an email, add a `sendMail` method using:
```
POST /users/{sender}/sendMail
```
The client already handles auth via `acquireTokenByClientCredential`.
The app needs `Mail.Send` application permission in Azure AD (may already be set).

### `src/storage/invoices.js` — Invoice store

Stores invoices as JSON. Each invoice has:
```js
{
  id: "uuid",
  sender: "Aaron Norris",           // from email From: header
  amount: 1234.56,                  // extracted from invoice
  hoursClaimed: 19.25,              // extracted from invoice
  weekEnding: "2026-05-01",         // extracted from invoice (usually a Friday)
  date: "2026-05-03",               // date invoice was received
  status: "pending",                // pending | validated | mismatch | review_needed
  notes: "awaiting Connecteam validation",
}
```

---

## Environment variables (`.env`)

```
# Microsoft Graph (Azure AD app — needs Mail.Read + Mail.Send application permissions)
MS_TENANT_ID=
MS_CLIENT_ID=
MS_CLIENT_SECRET=

# Mailboxes to scan for invoices
MAILBOXES=tom@corecastconcrete.com.au,accounts@corecastconcrete.com.au

# Connecteam (API key kept for potential future use — not used in main flow)
CONNECTEAM_API_KEY=
CONNECTEAM_BASE_URL=https://api.connecteam.com
CONNECTEAM_TIME_CLOCK_ID=16246267

# Folder where Tom drops the Connecteam XLSX export each week
TIMESHEETS_FOLDER=./timesheets

# Email reporting
REPORT_RECIPIENT=tom@corecastconcrete.com.au
REPORT_SENDER=accounts@corecastconcrete.com.au
HOURS_TOLERANCE=0.5

# Storage
INVOICES_FILE=./data/invoices.json
```

---

## Weekly workflow (manual steps + automation)

1. **Monday morning**: Tom exports timesheet from Connecteam UI
   - Go to Time Clock → Reports → Timesheet Overview
   - Select previous week (Mon–Sun)
   - Export as XLSX
   - Drop file into `./timesheets/` folder

2. **Automatically**: `node scripts/run-weekly-match.js`
   - Finds the latest XLSX in `./timesheets/`
   - Parses hours for each employee
   - Loads invoices received for that week from `data/invoices.json`
   - Matches each invoice sender to their timesheet hours
   - Sends validation report email to tom@corecastconcrete.com.au

3. **Tom reviews** the email:
   - ✅ Matches: approve and pay
   - ⚠️ Mismatches: follow up with subcontractor before paying

---

## Key decisions made

- **Not using Connecteam API** — plan limitation (Small Business plan returns empty shift data)
- **Tolerance of ±0.5h** — invoices are often rounded to nearest 15min, Connecteam tracks to the minute
- **Match by name** — `lookupHours()` does tolerant matching (case-insensitive, substring, last-name fallback)
- **XLSX sheet**: always use "All Employees" sheet, not individual employee sheets
- **Weekly total**: use column O ("Weekly total hours") — only populated on first row per employee
- **No userId API filter** — Connecteam's `userId` param on time-activities silently returns empty shifts

---

## npm packages

Already installed: `dotenv`, `@azure/msal-node`, `@microsoft/microsoft-graph-client`, `node-fetch`, `pdf-parse`
Added: `xlsx` (^0.18.5) — run `npm install` before using timesheetParser.js

---

## Email attachment filing (`src/filing/`, `src/sharepoint/`)

A separate flow from invoice ingestion. Polls the same mailboxes, extracts each
attachment, and uploads it to SharePoint based on a keyword match against the
email subject line.

### Layout

```
src/
├── sharepoint/
│   └── client.js                       # SharePointClient — uploads + folder ops via Graph
├── filing/
│   ├── keywords.js                     # Load + match keyword config
│   ├── state.js                        # FilingStore — JSON record of what's been filed
│   └── fileAttachments.js              # Orchestrator
config/
└── job-keywords.json                   # Keyword → {library, folder} map (user-editable)
data/
└── filed-attachments.json              # Per-attachment filing state (created on first run)
scripts/
└── file-attachments.js                 # CLI entry point (run from Task Scheduler)
```

### Run

```
npm run file-attachments                                  # last FILING_LOOKBACK_HOURS hours
node scripts/file-attachments.js --dry-run
node scripts/file-attachments.js --lookback 12            # override window for one run
node scripts/file-attachments.js --since 2026-05-01T00:00:00Z   # hard date overrides lookback
node scripts/file-attachments.js --lookback 0             # disable window (full inbox top N)
node scripts/file-attachments.js --mailbox tom@corecastconcrete.com.au --limit 100
```

The default rolling window (`FILING_LOOKBACK_HOURS`, 24 by default) plus the
dedup state file means re-runs are safe — overlap is intentional. Only emails
received inside the window are scanned, and any already-filed attachment is
skipped via `data/filed-attachments.json`.

### Behaviour — three-tier routing

For each attachment (after dedup against `data/filed-attachments.json`):

1. **Keyword tier** — match subject against `config/job-keywords.json`
   (case-insensitive, word-boundary; longest keyword wins on conflict).
   → Upload to the matched job folder. No notification.
2. **Domain tier** — if no keyword matched, look up the sender's domain in
   `config/supplier-domains.json`.
   → Upload to Accounts Payable (or per-domain override). No notification.
3. **To Sort fallback** — neither matched.
   → Upload to `To Sort` AND send a notification email to
     `FILING_NOTIFY_RECIPIENT` listing the sender and subject.

The run summary surfaces unique unmatched sender domains so you can grow the
supplier list iteratively. Independent of the invoice ingest flow — the same
email may be filed AND parsed into `data/invoices.json`.

### Adding a new job

Edit `config/job-keywords.json`:

```json
"Newkeyword": { "library": "ACTIVE JOBS", "folder": "Newkeyword - Site Address" }
```

### Adding a new supplier (routes to Accounts Payable)

Edit `config/supplier-domains.json`:

```json
"supplier.com.au": {}
```

Use `{}` for the default Accounts Payable destination, or override per supplier:

```json
"supplier.com.au": { "library": "ACTIVE JOBS", "folder": "Specific Folder" }
```

Both files are read on each run — no restart required.

### Required Azure permissions

Beyond the existing `Mail.Read` and `Mail.Send`:

- `Sites.ReadWrite.All` (application) — admin consent required.

After granting, verify with `node scripts/probe-sharepoint.js`. The first
filing run will surface a 403 if the permission isn't yet propagated.

### Scheduling

Add a Windows Task Scheduler entry similar to `CoreCast Dashboard Refresh`.
Suggested: every 30 minutes during business hours, command:

```
node C:\Users\tomba\corecast-invoice-automation\scripts\file-attachments.js
```
