# BIC Work Orders — Project Summary

This document contains the full context of the BIC Work Orders system for use in future AI-assisted development sessions.

---

## What This Is

A mobile-first web application for Bell Intercoolers LLC (BIC Machine Shop) that captures Purchase Orders by photo, extracts data using AI, stores it in Google Sheets, and displays a live dashboard for tracking job status and due dates.

Built by: Chris (machining@bellintercoolers.com)
GitHub: BICmachining
Status: Live and operational as of April 2026

---

## Live URLs

| What | URL |
|---|---|
| Capture app | https://BICmachining.github.io/bic-work-orders/index.html |
| Dashboard | https://BICmachining.github.io/bic-work-orders/dashboard.html |
| GitHub repo | https://github.com/BICmachining/bic-work-orders |
| Google Sheet | https://docs.google.com/spreadsheets/d/16-aMTlmhRaWvdZzBifXMJEVo5IA60TWcHC8rhHcLGQw |

---

## Architecture

```
Phone camera
    ↓
index.html (GitHub Pages)
    ↓ photo(s) as base64
Gemini Vision API (gemini-2.5-flash-lite)
    ↓ extracted JSON
Review screen (user edits inline)
    ↓ confirmed data
Google Apps Script Web App (GET request)
    ↓
Google Sheets (two tabs: POs + Line Items)
    ↑
dashboard.html reads via Apps Script GET
```

**Key technical decisions:**
- All API calls use GET (not POST) to avoid CORS issues with Apps Script redirects
- Credentials (Gemini API key, Apps Script URL) stored in browser localStorage — not in code
- No backend server — GitHub Pages (static) + Apps Script as the only server-side component
- Gemini free tier (gemini-2.5-flash-lite) handles image extraction at no cost

---

## Files

### index.html
The capture app. Screens: Config → Capture → Processing → Review → Success.

- **Config screen**: First-run setup for Gemini API key and Apps Script URL. Saved to localStorage.
- **Capture screen**: Multi-page photo queue. User adds pages one at a time (portrait mode). Each page shows as a thumbnail with remove button. "Read PO with AI" sends all pages to Gemini in one request.
- **Processing screen**: Spinner while Gemini processes. Shows page count.
- **Review screen**: Fully editable table of extracted data (landscape mode). PO header fields + line items table. Inline editing on all cells. Per-part material calculation is auto-computed from materialUsed ÷ qty.
- **Success screen**: Confirmation with line item count. Links to dashboard or new PO.

### dashboard.html
The live dashboard. Reads from Google Sheets via Apps Script GET.

- Stats bar: Open POs, Total Lines, Overdue, In Process, Waiting on Material, Complete
- Filter buttons by status + search by PO # or part number
- PO cards sorted by urgency (overdue first, then due-soon, then on-track)
- Each card expandable to show line items with inline status dropdowns (saves immediately on change)
- Color coding: red = overdue, yellow = due soon, green = complete
- Handwritten notes shown in amber callout
- Trends section at bottom: parts completed by month (appears once data exists)

### apps-script.js
Google Apps Script — the only server-side code. Paste into Extensions → Apps Script in the Google Sheet. Deploy as Web App, Execute as Me, Who has access: Anyone.

Handles three GET actions:
- `?action=getDashboardData` — returns all POs and line items as JSON
- `?action=submitPO&po=...&lineItems=...` — writes a new PO and its line items
- `?action=updateStatus&poNumber=...&line=...&status=...` — updates a single line item status

Includes `formatDate()` helper to convert Google Sheets date objects to M/D/YY strings.

---

## Google Sheets Structure

**Sheet ID:** `16-aMTlmhRaWvdZzBifXMJEVo5IA60TWcHC8rhHcLGQw`

### Tab: POs
| A | B | C | D | E | F |
|---|---|---|---|---|---|
| PO Number | PO Date | Ordered By | Order Description | Handwritten Notes | Date Added |

### Tab: Line Items
| A | B | C | D | E | F | G | H | I | J | K | L |
|---|---|---|---|---|---|---|---|---|---|---|---|
| PO Number | Line | Part Number | Qty | Description | Material Part # | Material Description | Material Used (in) | Material Per Part (in) | Due Date | Status | Notes |

**PO Number is the join key between tabs.**

---

## PO Document Structure

Bell Intercoolers POs are printed documents with this structure:

**Header fields (repeat on every page):**
- PO Number, PO Date, Ordered By (initials: HC, MS, etc.), Order Desc

**Line item table columns:**
- Line (001, 002, 003...), Order Qty, Part Number / Description, Unit Cost, UM, Amount, Dt Req'd

**M1 sub-lines:**
Each numbered line item has an M1 sub-line directly below it showing the raw material to pull from stock. Example:
```
001   2.00   MAC103-001LEE2 — Flange for CA300057150LEE2...    EA   04/27/26
M1   -9.00   M00105-0064-ALU — FLAT,Alu,1.00"x10.00"          IN   04/27/26
```
The M1 line is NOT a separate work task — it's the bill of material for the line above it. The negative quantity (-9.00) means 9 inches of that stock is consumed total. Per-part = 9.00 ÷ 2 = 4.50 inches.

**Handwritten notes** sometimes appear anywhere on the document (e.g. "NOT A Priority", "already done").

---

## Gemini Extraction Prompt Logic

The prompt instructs Gemini to:
1. Take header data from page 1 only (header repeats on all pages)
2. Collect ALL numbered line items across all pages sequentially
3. Fold each M1 sub-line into its parent line item
4. Return materialUsed as a positive number
5. Return only valid JSON — no markdown fences, no explanation

Model: `gemini-2.5-flash-lite` (free tier, no billing required)
Max output tokens: 4096 (handles large multi-page POs)
Temperature: 0.1 (low for consistent structured output)

---

## Status Values

Exactly four statuses used throughout (case-sensitive, must match exactly):
- `Not Started`
- `Waiting on Material`
- `In Process`
- `Complete`

---

## Credentials & Config

Credentials are stored in browser localStorage under key `bic_config`:
```json
{
  "gemini": "AIza...",
  "script": "https://script.google.com/macros/s/.../exec"
}
```

Stored per-device per-browser. Must be entered separately on desktop and phone.

**Current Apps Script deployment URL:**
`https://script.google.com/macros/s/AKfycbzhIw7tYtt2H6v3OYC4wKmyG-Nbh-aFYvjwH5LMp0HwQCACd4QEgZvkc4jDJCAQduJcCw/exec`

Note: If the Apps Script is redeployed, this URL changes and must be updated in Settings on all devices.

---

## Known Issues & Solutions Encountered

**Gemini model changes (April 2026):**
- `gemini-2.0-flash` — deprecated, requires billing even on free tier
- `gemini-1.5-flash-latest` — no longer available
- `gemini-2.5-flash` — works but high demand causes intermittent errors
- `gemini-2.5-flash-lite` — current working model, less congested

**CORS / Apps Script redirect issue:**
Apps Script POST requests fail from GitHub Pages due to CORS. Solution: all requests use GET with URL parameters. Data is JSON-encoded and URI-encoded in the query string.

**Date format issue:**
Google Sheets auto-converts date strings to Date objects. Apps Script `formatDate()` helper converts them back to M/D/YY before returning JSON.

**Apps Script caching:**
When redeploying Apps Script, always use "New deployment" — editing an existing deployment can cache the old version.

**Mobile localStorage:**
Credentials entered on desktop are NOT shared with mobile. Must be entered separately on each device/browser.

---

## Future Development Ideas

- **Material inventory tracking** — use the Material Used field to track stock consumption over time
- **Trend analytics** — parts completed per month/year already captured in dashboard trends section
- **Email notifications** — Apps Script can send email on new PO submission (Gmail integration built into Apps Script)
- **Photo archive** — save original PO photos to Google Drive alongside Sheet entries
- **Part number lookup** — cross-reference part numbers against a master parts list tab in Sheets
- **Overdue alerts** — Apps Script time-based trigger to email when jobs are past due date
- **Multi-user** — currently single-user (one machinist). Could be extended with an "Assigned To" field
- **Print view** — dashboard print stylesheet for shop floor paper reference

---

## Deployment Process

To update any file:
1. Go to github.com/BICmachining/bic-work-orders
2. Click the file → pencil icon (Edit)
3. Select all, delete, paste new content
4. Commit changes
5. Wait ~60 seconds for GitHub Pages to rebuild

To update Apps Script:
1. Google Sheet → Extensions → Apps Script
2. Replace all code, save (Ctrl+S)
3. Deploy → **New deployment** (not manage existing)
4. Web app, Execute as Me, Who has access: Anyone
5. Copy new URL → update Settings on all devices
