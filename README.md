# Frido — Low RTD Level Report · Email Automation

Automated daily email report that identifies all SKUs with a **Low RTD (Ready-To-Dispatch) Level** across Frido's fulfilment centres, generates a formatted Excel workbook, and delivers it to the supply-chain team twice every day via Microsoft 365 — all running serverlessly on **Google Cloud Run + Cloud Scheduler**.

---

## Table of Contents

1. [What is RTD Level?](#1-what-is-rtd-level)
2. [What This System Does](#2-what-this-system-does)
3. [Complete Flow Diagram](#3-complete-flow-diagram)
4. [Repository Structure](#4-repository-structure)
5. [Architecture Overview](#5-architecture-overview)
6. [Data Source](#6-data-source)
7. [Code Walkthrough — generate\_low\_rtd\_report.py](#7-code-walkthrough--generate_low_rtd_reportpy)
8. [Code Walkthrough — send\_low\_rtd\_email.py](#8-code-walkthrough--send_low_rtd_emailpy)
9. [Excel Report — Sheet by Sheet](#9-excel-report--sheet-by-sheet)
10. [DOI & Stock Status Logic](#10-doi--stock-status-logic)
11. [Email Configuration](#11-email-configuration)
12. [Docker Container](#12-docker-container)
13. [Google Cloud Infrastructure](#13-google-cloud-infrastructure)
14. [Schedule — When Emails Go Out](#14-schedule--when-emails-go-out)
15. [How to Re-deploy (Step by Step)](#15-how-to-re-deploy-step-by-step)
16. [How to Run Locally](#16-how-to-run-locally)
17. [How to Trigger Manually on Cloud](#17-how-to-trigger-manually-on-cloud)
18. [Monitoring & Logs](#18-monitoring--logs)
19. [Changing Recipients](#19-changing-recipients)
20. [Changing the Schedule](#20-changing-the-schedule)
21. [Dependencies](#21-dependencies)
22. [Credentials Reference](#22-credentials-reference)
23. [Troubleshooting](#23-troubleshooting)

---

## 1. What is RTD Level?

**RTD = Ready-To-Dispatch** inventory — stock that is fully processed, QC-passed, and physically available to ship to a customer the moment an order is placed.

The RTD Level for a SKU at a location is calculated using **Days of Inventory (DOI)**:

```
DOI = floor(RTD Inventory / ceil(Avg Daily Sale))
```

| DOI Range | RTD Level  | Colour in Report |
|-----------|------------|-----------------|
| 0         | Out of Stock | — |
| Avg Sale = 0 | No Demand | — |
| 1 – 2     | **Critical** | 🔴 Red |
| 3 – 15    | **Low**      | 🟠 Orange |
| 16 – 45   | Sufficient   | 🟢 Green |
| > 45      | Excess       | 🟡 Yellow |

This report focuses exclusively on SKUs whose **overall RTD Level = Low** (DOI 3–15 days across all FCs combined).

---

## 2. What This System Does

Every day at **7:15 AM IST** and **4:15 PM IST**:

1. Fetches live inventory JSON from the Frido Navigator dashboard API
2. Filters all SKUs where `rtdLevel == "Low"`
3. Builds per-location detail rows (only Regular facility type)
4. Writes a multi-sheet, colour-coded Excel workbook
5. Sends the workbook as an email attachment via Microsoft 365 Graph API
6. From: `reports@myfrido.com`
7. To: Supply-chain team leads
8. CC: Analytics, ops, and management stakeholders

---

## 3. Complete Flow Diagram

```
╔══════════════════════════════════════════════════════════════════════════════════╗
║                        DAILY TRIGGER (twice a day)                              ║
║                                                                                  ║
║   7:15 AM IST ──┐                                                                ║
║                 ├──► Google Cloud Scheduler ──► HTTP POST                        ║
║   4:15 PM IST ──┘         (asia-south1)                                          ║
╚══════════════════════════════════════════════════════════════════════════════════╝
                                      │
                                      ▼
╔══════════════════════════════════════════════════════════════════════════════════╗
║                    GOOGLE CLOUD RUN JOB  (low-rtd-email-job)                    ║
║                         Region: asia-south1  │  Python 3.11                     ║
║                                                                                  ║
║   ┌─────────────────────────────────────────────────────────────────────────┐   ║
║   │  STEP 1 — generate_low_rtd_report.py                                    │   ║
║   │                                                                         │   ║
║   │  fetch_inv_json()                                                        │   ║
║   │       │                                                                  │   ║
║   │       ▼                                                                  │   ║
║   │  GET https://frido-navigator.vercel.app/inv-data-7d.json                │   ║
║   │       │                                                                  │   ║
║   │       ▼  (JSON with all SKUs + location breakdown)                      │   ║
║   │                                                                         │   ║
║   │  build_detail_rows(data)                                                 │   ║
║   │       │                                                                  │   ║
║   │       ├── For each SKU: filter rtdLevel == "Low"                        │   ║
║   │       ├── For each location: aggregate only facilityType=Regular        │   ║
║   │       ├── Skip rows where totalInvt==0 AND avgSale==0                   │   ║
║   │       └── Returns list of flat dicts (one row per SKU × location)       │   ║
║   │                                                                         │   ║
║   │  write_excel(detail_rows, today, output_path)                           │   ║
║   │       │                                                                  │   ║
║   │       ├── Sheet: inventory_detail  (all SKUs × locations, sorted)       │   ║
║   │       ├── Sheet: Sheet1            (pivot: SKU rows × location cols)    │   ║
║   │       ├── Sheet: PNQ               (Pune FC — if has data)              │   ║
║   │       ├── Sheet: GGN               (Gurgaon FC — if has data)           │   ║
║   │       ├── Sheet: BLR               (Bangalore FC — if has data)         │   ║
║   │       ├── Sheet: MUM               (Mumbai FC — if has data)            │   ║
║   │       ├── Sheet: KOL               (Kolkata FC — if has data)           │   ║
║   │       ├── Sheet: HYD               (Hyderabad FC — if has data)         │   ║
║   │       └── Sheet: CHN               (Chennai FC — if has data)           │   ║
║   │                                                                         │   ║
║   │  Returns: (file_path, stats_dict)                                       │   ║
║   └─────────────────────────────────────────────────────────────────────────┘   ║
║                                      │                                           ║
║                                      ▼                                           ║
║   ┌─────────────────────────────────────────────────────────────────────────┐   ║
║   │  STEP 2 — send_low_rtd_email.py                                         │   ║
║   │                                                                         │   ║
║   │  get_access_token()                                                      │   ║
║   │       │                                                                  │   ║
║   │       ▼                                                                  │   ║
║   │  POST https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token   │   ║
║   │       │  (client_credentials grant — app-only, no user login needed)    │   ║
║   │       ▼                                                                  │   ║
║   │  Bearer Token (JWT, valid 1 hour)                                        │   ║
║   │                                                                         │   ║
║   │  send_email(report_path, stats)                                          │   ║
║   │       │                                                                  │   ║
║   │       ├── Read .xlsx → base64 encode                                    │   ║
║   │       ├── Build JSON payload (subject, HTML body, toRecipients,         │   ║
║   │       │   ccRecipients, fileAttachment)                                  │   ║
║   │       └── POST https://graph.microsoft.com/v1.0/                        │   ║
║   │               users/reports@myfrido.com/sendMail                        │   ║
║   │                                                                         │   ║
║   │  HTTP 202 Accepted → Email queued in M365 for delivery                  │   ║
║   └─────────────────────────────────────────────────────────────────────────┘   ║
╚══════════════════════════════════════════════════════════════════════════════════╝
                                      │
                                      ▼
╔══════════════════════════════════════════════════════════════════════════════════╗
║                       MICROSOFT 365 MAIL DELIVERY                               ║
║                                                                                  ║
║   From : reports@myfrido.com                                                     ║
║                                                                                  ║
║   To   : kunal.Kumbhar@myfrido.com                                               ║
║          navnath.k@myfrido.com                                                   ║
║          gaurav.k@myfrido.com                                                    ║
║          saurabh.m@myfrido.com                                                   ║
║          sagar.k@myfrido.com                                                     ║
║                                                                                  ║
║   CC   : hemant.k@myfrido.com      nikhil.c@myfrido.com                          ║
║          Ipshit.k@myfrido.com      pravin.k@myfrido.com                          ║
║          amit.c@myfrido.com        suraj.s@myfrido.com                           ║
║          yash.s@myfrido.com        yash.p@myfrido.com                            ║
║          abhishek.v@myfrido.com    tushar.g@myfrido.com                          ║
║          pranjal.t@myfrido.com                                                   ║
║                                                                                  ║
║   Attachment: Low_RTD_Level_DD_Mon_YYYY.xlsx                                     ║
╚══════════════════════════════════════════════════════════════════════════════════╝
```

---

## 4. Repository Structure

```
cloud_run_Low_RTD_Level_Report/
│
├── generate_low_rtd_report.py   # Core report generator — fetches data, builds Excel
├── send_low_rtd_email.py        # Email sender — calls generator, sends via Graph API
│
├── requirements.txt             # Python package dependencies (pinned versions)
├── Dockerfile                   # Container definition for Cloud Run Job
├── .dockerignore                # Files excluded from Docker build context
│
├── deploy.ps1                   # One-shot PowerShell deployment script (run once)
└── README.md                    # This file
```

---

## 5. Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                        Google Cloud (frido-429506)              │
│                         Region: asia-south1                      │
│                                                                 │
│   ┌───────────────────────┐      ┌──────────────────────────┐  │
│   │   Cloud Scheduler     │      │   Container Registry     │  │
│   │                       │      │   (gcr.io)               │  │
│   │  low-rtd-email-morning│      │                          │  │
│   │  ┌─────────────────┐  │      │  gcr.io/frido-429506/    │  │
│   │  │ 15 7 * * *      │  │      │  low-rtd-email:latest    │  │
│   │  │ Asia/Kolkata    │  │      │                          │  │
│   │  └────────┬────────┘  │      └──────────┬───────────────┘  │
│   │           │            │                 │                  │
│   │  low-rtd-email-evening│                 │ image            │
│   │  ┌─────────────────┐  │                 │                  │
│   │  │ 15 16 * * *     │  │      ┌──────────▼───────────────┐  │
│   │  │ Asia/Kolkata    │  │      │   Cloud Run Job          │  │
│   │  └────────┬────────┘  │      │   low-rtd-email-job      │  │
│   └───────────┼────────────┘      │                          │  │
│               │  HTTP POST        │   Memory : 512 Mi        │  │
│               └──────────────────►│   CPU    : 1             │  │
│                  trigger job run  │   Timeout: 600s          │  │
│                                   │   Retries: 1             │  │
│                                   └──────────────────────────┘  │
│                                                                 │
│   Service Account: for-quickcomm-automation@frido-429506        │
│                    .iam.gserviceaccount.com                     │
└─────────────────────────────────────────────────────────────────┘
         │                                    │
         │ fetches JSON                       │ sends email
         ▼                                    ▼
  frido-navigator                    Microsoft 365
  .vercel.app                        Graph API
  /inv-data-7d.json                  graph.microsoft.com
```

---

## 6. Data Source

**URL:** `https://frido-navigator.vercel.app/inv-data-7d.json`

This is the same pre-computed JSON that powers the Frido Navigator inventory dashboard. Using the same source guarantees that numbers in the email report **exactly match** what the team sees on the dashboard — no discrepancy.

**JSON structure (simplified):**

```json
{
  "asOf": "2026-10-12T05:30:00Z",
  "lastSnapshotUpdated": "2026-10-12T04:00:00Z",
  "avgSaleWindow": 7,
  "skus": [
    {
      "sku": "FR-WC-001",
      "category": "Wheelchairs",
      "subCategory": "Manual Wheelchair Standard",
      "rtdLevel": "Low",
      "websiteStatus": "Live",
      "totalInvt": 45,
      "rtdInvt": 30,
      "rawInvt": 10,
      "rawBlockedInvt": 5,
      "avgSale": 3.2,
      "locations": [
        {
          "location": "PNQ",
          "totalInvt": 20,
          "rtdInvt": 15,
          "rawInvt": 4,
          "rawBlockedInvt": 1,
          "avgSale": 1.5,
          "doi": 10,
          "stockStatus": "Low"
        }
        ...
      ],
      "facilities": [
        {
          "location": "PNQ",
          "facilityType": "Regular",
          "totalInvt": 18,
          "rtdInvt": 14,
          "rawInvt": 3,
          "rawBlockedInvt": 1
        },
        {
          "location": "PNQ",
          "facilityType": "B2B",
          "totalInvt": 2,
          ...
        }
      ]
    }
  ]
}
```

**Key fields:**

| Field | Meaning |
|---|---|
| `rtdLevel` | Overall RTD health across all FCs (`Low`, `Sufficient`, `Excess`, etc.) |
| `websiteStatus` | Whether SKU is live on website (`Live` / `Inactive`) |
| `totalInvt` | Total inventory (RTD + RAW + Blocked) |
| `rtdInvt` | Ready-to-dispatch units |
| `rawInvt` | Raw / unprocessed units |
| `rawBlockedInvt` | Blocked units (damaged, QC hold, etc.) |
| `avgSale` | Average daily B2C sales over the last 7 days |
| `doi` | Days of Inventory = floor(totalInvt / ceil(avgSale)) |
| `facilities[]` | Breakdown by facility type within each location |

---

## 7. Code Walkthrough — `generate_low_rtd_report.py`

### 7.1 `fetch_inv_json()`

```python
def fetch_inv_json():
    with urllib.request.urlopen(INV_JSON_URL, timeout=60) as r:
        data = json.load(r)
    return data
```

- Opens an HTTP connection to the Vercel-hosted JSON endpoint
- 60-second timeout prevents the job from hanging indefinitely
- Returns the full parsed JSON dictionary

---

### 7.2 `regular_locations(s)` — The Critical Filter

This function mirrors the `regularLocations()` JavaScript function in `InventoryHealthPage.jsx` of the dashboard frontend. It is the most important logic for data accuracy.

**Why it exists:** Each physical warehouse location can have multiple facility types — `Regular` (B2C stock), `B2B` (bulk/dealer stock), `Transit`, etc. The dashboard only shows Regular facility stock. If we included B2B stock, the numbers would not match the dashboard.

```
SKU at PNQ location
├── Facility: Regular  → totalInvt: 18  ← INCLUDE
├── Facility: B2B      → totalInvt: 2   ← EXCLUDE
└── Facility: Transit  → totalInvt: 0   ← EXCLUDE
```

**Step-by-step logic:**

```
If SKU has no facilities[] array:
    → Fall back to s.locations[] directly (older data format)

Else:
    1. Loop through facilities[]
       - Skip any facility where facilityType != "Regular"
       - Skip any facility at a location not in LOCATION_ORDER
       - Aggregate totalInvt, rtdInvt, rawInvt, rawBlockedInvt by location

    2. Merge avgSale, doi, stockStatus from s.locations[]
       (These are computed at the location level, not facility level)

    3. Recompute DOI and stockStatus based on aggregated Regular stock:
       DOI = floor(totalInvt / ceil(avgSale))
       stockStatus = based on DOI thresholds (see Section 10)

    4. Return list of per-location aggregated dicts
```

---

### 7.3 `build_detail_rows(data)`

Converts the raw JSON into a flat list of rows, one row per **SKU × Location** combination.

```
For each SKU in data["skus"]:
    If SKU's rtdLevel != "Low" → skip entirely

    Capture SKU-level all-FC totals:
        all_fc_total   = SKU.totalInvt   (sum across ALL locations, ALL facilities)
        all_fc_rtd     = SKU.rtdInvt
        all_fc_raw     = SKU.rawInvt
        all_fc_blocked = SKU.rawBlockedInvt
        all_fc_avg     = SKU.avgSale

    For each location returned by regular_locations(SKU):
        If totalInvt == 0 AND avgSale == 0 → skip (no meaningful data)

        Append row dict:
        {
            category, sub_category, product_id,
            location, rtd_invt, raw_invt, raw_blocked,
            total_invt, avg_sale, doi, stock_status,
            website_status,
            all_fc_total, all_fc_rtd, all_fc_raw,
            all_fc_blocked, all_fc_avg_sale, rtd_level
        }
```

**Output example (one row):**

```python
{
    "category":        "Wheelchairs",
    "sub_category":    "Manual Wheelchair Standard",
    "product_id":      "FR-WC-001",
    "location":        "PNQ",
    "rtd_invt":        14,
    "raw_invt":        3,
    "raw_blocked":     1,
    "total_invt":      18,
    "avg_sale":        1.5,
    "doi":             10,
    "stock_status":    "Low",
    "website_status":  "Live",
    "all_fc_total":    45,
    "all_fc_rtd":      30,
    "all_fc_raw":      10,
    "all_fc_blocked":  5,
    "all_fc_avg_sale": 3.2,
    "rtd_level":       "Low",
}
```

---

### 7.4 `write_excel(detail_rows, report_date, output_path)`

Builds the multi-sheet Excel workbook using `xlsxwriter`. See [Section 9](#9-excel-report--sheet-by-sheet) for a full breakdown of each sheet.

**Colour palette used:**

| Variable | Hex | Used For |
|---|---|---|
| `NAV` | `#1A3A5C` | Header rows, tab colour |
| `GRY` | `#F5F5F5` | Alternating row background |
| `WHT` | `#FFFFFF` | Default cell background |
| `RED` | `#F4CCCC` | Critical stock / DOI ≤ 2 |
| `ORA` | `#FCE5CD` | Low stock / DOI 3–15 |
| `YEL` | `#FFF2CC` | Excess stock / DOI > 45 |
| `GRN` | `#D9EAD3` | Sufficient stock / DOI 16–45 |
| `BLU` | `#CFE2F3` | Total / grand total rows |

---

### 7.5 `main()`

Entry point for the report generator:

```
1. Get today's date (UTC)
2. fetch_inv_json() → raw data
3. build_detail_rows(data) → flat list of dicts
4. If no rows → return (None, {}) — email will not be sent
5. write_excel(...) → saves Low_RTD_Level_DD_Mon_YYYY.xlsx
6. Return (file_path, stats_dict)
```

---

## 8. Code Walkthrough — `send_low_rtd_email.py`

### 8.1 `get_access_token()`

Uses **OAuth 2.0 Client Credentials flow** (app-only authentication — no user password needed):

```
App Registration in Azure AD (Entra ID)
    CLIENT_ID     = Azure App ID
    CLIENT_SECRET = App secret
    TENANT_ID     = Frido's M365 tenant

POST https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token
    grant_type=client_credentials
    scope=https://graph.microsoft.com/.default

← Returns: { "access_token": "eyJ0...", "expires_in": 3600 }
```

The token is valid for 1 hour — well within our job's 10-minute runtime.

---

### 8.2 `build_email_body(report_date_str, stats)`

Returns a minimal HTML email body:

```html
<html>
<body style="font-family: Arial, sans-serif; ...">
  <p>Hi All,</p>
  <p>Please find attached the <strong>Low RTD Level Report for 12 Oct 2026</strong>.</p>
  <p>
    Thanks and Regards<br>
    <strong>Data Analytics Team</strong><br>
    Frido
  </p>
</body>
</html>
```

---

### 8.3 `send_email(report_path, stats)`

```
1. Read .xlsx file from disk → base64 encode

2. Build Microsoft Graph API payload:
   {
     "message": {
       "subject": "Frido Low RTD Level Report - 12 Oct 2026",
       "body": { "contentType": "HTML", "content": "<html>..." },
       "toRecipients": [ {emailAddress: {address: "..."}}, ... ],
       "ccRecipients":  [ {emailAddress: {address: "..."}}, ... ],
       "attachments": [
         {
           "@odata.type": "#microsoft.graph.fileAttachment",
           "name": "Low_RTD_Level_12_Oct_2026.xlsx",
           "contentType": "application/vnd.openxmlformats-...",
           "contentBytes": "<base64 string>"
         }
       ]
     },
     "saveToSentItems": true
   }

3. POST https://graph.microsoft.com/v1.0/users/reports@myfrido.com/sendMail
   Authorization: Bearer <token>
   Content-Type: application/json

4. HTTP 202 Accepted → success
   Any other status → raise RuntimeError with details
```

---

### 8.4 `main()`

```
1. Call generate_report() → (report_path, stats)
2. If report_path is None → print message, exit (no email sent)
3. Call send_email(report_path, stats)
```

---

## 9. Excel Report — Sheet by Sheet

### Sheet 1: `inventory_detail`

**Purpose:** Complete raw data for all Low RTD SKUs across all locations.

**Filters applied:** `rtdLevel == "Low"` only. All locations included (no website status filter here).

**Sort order:** Category → Sub-category → Product ID → Location

**Columns:**

| Column | Description |
|---|---|
| Category | Product category (e.g., Wheelchairs) |
| Sub-category | Product name / sub-category |
| Product ID | SKU code |
| Location | FC code (PNQ / GGN / BLR / MUM / KOL / HYD / CHN) |
| RTD Inventory | Ready-to-dispatch units at this FC |
| RAW Inventory | Unprocessed units at this FC |
| RAW Blocked Inventory | Blocked/QC-hold units at this FC |
| **Total Inventory** | RTD + RAW + Blocked (bold) |
| Avg Sale (B2C) | 7-day average daily B2C sales at this FC |
| DOI | Days of Inventory = floor(Total / ceil(Avg Sale)) |
| Stock Status | Critical / Low / Sufficient / Excess / Out of Stock |

**Colour coding in Stock Status column:**

| Value | Background |
|---|---|
| Critical | 🔴 Red `#F4CCCC` |
| Low | 🟠 Orange `#FCE5CD` |
| Sufficient | 🟢 Green `#D9EAD3` |
| Excess | 🟡 Yellow `#FFF2CC` |

**Total row** at bottom in blue (`#CFE2F3`), summing all numeric columns.

---

### Sheet 2: `Sheet1` (Pivot View)

**Purpose:** Cross-tab matrix — one row per SKU, columns grouped by FC location.

**Filters applied (stricter than inventory_detail):**
- Total Inventory > 0
- Avg Sale (B2C) > 0
- Website Status = **Live** only
- DOI > 1 (excludes DOI = 0 or 1, i.e., near-stockout)

**Layout:**

```
Rows 0–3: Filter labels (shown at top of sheet for transparency)
Row 5:    Location group headers  [PNQ] [GGN] [BLR] [MUM] [KOL] [HYD] [CHN] [Totals]
Row 6:    Value sub-headers       RTD Invt | RAW Invt | RAW Blocked | Avg Sale (repeated per FC)
Row 7+:   One row per Product ID (sorted alphabetically)
Last row: Grand Total
```

**Value columns per location:**
1. RTD Invt
2. RAW Invt
3. RAW Blocked Invt
4. Regional Avg Sale

**Total columns (rightmost):**
1. Total RTD Invt (sum across all FCs)
2. Total RAW Invt
3. Total RAW Blocked Invt
4. Total Regional Avg Sale

Cells with zero value display `"-"` instead of `0` for cleaner reading.

---

### Sheets 3–9: Per-Location Sheets (`PNQ`, `GGN`, `BLR`, `MUM`, `KOL`, `HYD`, `CHN`)

**Purpose:** Warehouse-specific view showing both local (This FC) and network-wide (All FCs) data side by side.

**One sheet per FC** — a sheet is only created if that FC has at least one qualifying row.

**Filters:** Same as Sheet1 (Live, Total > 0, Avg Sale > 0, DOI > 1)

**Sort order:** Category → Sub-category → Product ID

**Column groups:**

```
┌──────────────────────────┬──────────────────────────────────────┬─────────────────────────────┬──────────┐
│  (Identity — dark navy)  │    This FC  (steel blue #1F6B8E)     │  All FCs  (amber #B45309)   │ Status   │
├────┬──────────────┬───── ┼────┬────┬──────┬──────────┬──────────┼───────┬──────┬────────┬──── ┼──────────┤
│SKU │ Product Name │ Cat  │RTD │RAW │Blkd  │ Reg Avg  │ All Tot  │All RTD│AllAvg│All DOI │     │RTD Level │
│    │              │      │Invt│Invt│ Invt │  Sale    │  Invt    │       │      │        │     │          │
└────┴──────────────┴──────┴────┴────┴──────┴──────────┴──────────┴───────┴──────┴────────┴─────┴──────────┘
```

**All FCs DOI** is computed dynamically:
```python
All FCs DOI = floor(all_fc_rtd / all_fc_avg_sale)  if all_fc_avg_sale > 0
            = None                                    otherwise
```

**DOI cell colour coding:**

| DOI Value | Background |
|---|---|
| None or 0 | 🔴 Red |
| 1 – 2 | 🔴 Red |
| 3 – 15 | 🟠 Orange |
| 16 – 45 | 🟡 Yellow |
| > 45 | 🟢 Green |

**Overall RTD Level column:**

| Value | Background |
|---|---|
| Low | 🔴 Red |
| Anything else | 🟢 Green |

---

## 10. DOI & Stock Status Logic

This exactly mirrors the JavaScript logic in the Frido Navigator dashboard:

```python
denom = math.ceil(avg_sale)        # round up avg sale to nearest integer

if total_invt == 0:
    doi = 0
    stock_status = "Out of Stock"

elif denom == 0:                    # SKU has no demand / no recent sales
    doi = None
    stock_status = "No Demand"

else:
    doi = math.floor(total_invt / denom)

    if   doi <= 2:   stock_status = "Critical"
    elif doi <= 15:  stock_status = "Low"
    elif doi <= 45:  stock_status = "Sufficient"
    else:            stock_status = "Excess"
```

**Why `ceil` for avgSale and `floor` for DOI?**
- `ceil(avgSale)` = conservative demand estimate (avoids overestimating cover)
- `floor(DOI)` = conservative days remaining (avoids false comfort)

---

## 11. Email Configuration

| Property | Value |
|---|---|
| **From** | `reports@myfrido.com` |
| **Subject** | `Frido Low RTD Level Report - DD Mon YYYY` |
| **Body** | HTML (see Section 8.2) |
| **Attachment** | `Low_RTD_Level_DD_Mon_YYYY.xlsx` |
| **API** | Microsoft Graph API v1.0 |
| **Auth** | OAuth 2.0 Client Credentials (app-only) |
| **Save to Sent** | Yes (`saveToSentItems: true`) |

### To Recipients

| Name | Email |
|---|---|
| Kunal Kumbhar | kunal.Kumbhar@myfrido.com |
| Navnath Kamble | navnath.k@myfrido.com |
| Gaurav Kumar | gaurav.k@myfrido.com |
| Saurabh Moon | saurabh.m@myfrido.com |
| Sagar Kokkul | sagar.k@myfrido.com |

### CC Recipients

| Name | Email |
|---|---|
| Hemant Kale | hemant.k@myfrido.com |
| Ipshit Kumar | Ipshit.k@myfrido.com |
| Nikhil Chandel | nikhil.c@myfrido.com |
| Pravin Kore | pravin.k@myfrido.com |
| Amit Chavan | amit.c@myfrido.com |
| Suraj Singh | suraj.s@myfrido.com |
| Yash Sakure | yash.s@myfrido.com |
| Yash Pawar | yash.p@myfrido.com |
| Abhishek Varma | abhishek.v@myfrido.com |
| Tushar Gupta | tushar.g@myfrido.com |
| Pranjal Tripati | pranjal.t@myfrido.com |

---

## 12. Docker Container

### `Dockerfile`

```dockerfile
FROM python:3.11-slim        # Minimal Python 3.11 base image (~150 MB)

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt   # Install deps (no cache = smaller image)

COPY generate_low_rtd_report.py .
COPY send_low_rtd_email.py .

CMD ["python", "send_low_rtd_email.py"]              # Entry point
```

### `.dockerignore`

Excluded from Docker build context (keeps image lean):
- `__pycache__/` and `*.pyc` — Python bytecode
- `*.xlsx` — any local report files
- `Untitled` — scratch file
- `deploy.ps1` — deployment script (not needed at runtime)

### Image location

```
gcr.io/frido-429506/low-rtd-email:latest
```

---

## 13. Google Cloud Infrastructure

| Resource | Name / Value |
|---|---|
| **GCP Project** | `frido-429506` |
| **Region** | `asia-south1` (Mumbai) |
| **Cloud Run Job** | `low-rtd-email-job` |
| **Container Image** | `gcr.io/frido-429506/low-rtd-email` |
| **Memory** | 512 Mi |
| **CPU** | 1 vCPU |
| **Task Timeout** | 600 seconds (10 minutes) |
| **Max Retries** | 1 |
| **Service Account** | `for-quickcomm-automation@frido-429506.iam.gserviceaccount.com` |
| **Scheduler Job 1** | `low-rtd-email-morning` |
| **Scheduler Job 2** | `low-rtd-email-evening` |

---

## 14. Schedule — When Emails Go Out

| Job Name | Cron | Timezone | IST Time | UTC Time |
|---|---|---|---|---|
| `low-rtd-email-morning` | `15 7 * * *` | Asia/Kolkata | **7:15 AM** | 1:45 AM |
| `low-rtd-email-evening` | `15 16 * * *` | Asia/Kolkata | **4:15 PM** | 10:45 AM |

Runs every day including weekends. The morning email gives the team data at the start of the day; the evening email captures any intra-day inventory changes.

---

## 15. How to Re-deploy (Step by Step)

Use this if you change the Python code and need to push a new version to Cloud.

**Prerequisites:**
- [Google Cloud SDK](https://cloud.google.com/sdk/docs/install) installed
- Logged in: `gcloud auth login`
- Project set: `gcloud config set project frido-429506`

**Steps:**

```powershell
# 1. Navigate to the project folder
cd "path\to\Low_RTD_Email_Automation"

# 2. Build new Docker image and deploy Cloud Run Job
#    (this replaces the existing job with the new image)
gcloud run jobs deploy low-rtd-email-job `
  --source . `
  --region asia-south1 `
  --project frido-429506 `
  --service-account for-quickcomm-automation@frido-429506.iam.gserviceaccount.com `
  --task-timeout 600 `
  --max-retries 1 `
  --memory 512Mi `
  --cpu 1
```

The Cloud Scheduler jobs (`low-rtd-email-morning` and `low-rtd-email-evening`) do **not** need to be re-created — they already point to the job name and will automatically use the new image on the next trigger.

---

## 16. How to Run Locally

```bash
# Install dependencies
pip install -r requirements.txt

# Run the full pipeline (generate report + send email)
python send_low_rtd_email.py

# Or just generate the Excel (without sending email)
python generate_low_rtd_report.py
```

The Excel file will be saved in the same directory as the script, named:
`Low_RTD_Level_DD_Mon_YYYY.xlsx`

---

## 17. How to Trigger Manually on Cloud

To run the job immediately (outside the scheduled time):

```bash
gcloud run jobs execute low-rtd-email-job \
  --region asia-south1 \
  --project frido-429506 \
  --wait
```

`--wait` keeps the terminal open until the job completes, showing success or failure.

Or trigger via Cloud Scheduler (simulates what the cron would do):

```bash
gcloud scheduler jobs run low-rtd-email-morning \
  --location asia-south1 \
  --project frido-429506
```

---

## 18. Monitoring & Logs

### View recent job executions

```bash
gcloud run jobs executions list \
  --job low-rtd-email-job \
  --region asia-south1 \
  --project frido-429506
```

### View logs for the most recent execution

```bash
gcloud logging read \
  "resource.type=cloud_run_job AND resource.labels.job_name=low-rtd-email-job" \
  --project frido-429506 \
  --limit 50 \
  --format "value(textPayload)"
```

### What healthy logs look like

```
Fetching dashboard JSON: https://frido-navigator.vercel.app/inv-data-7d.json
  asOf          : 2026-10-12T05:30:00Z
  snapshotUpdated: 2026-10-12T04:00:00Z
  avgSaleWindow : 7
  Total SKUs    : 1842
Building detail rows (RTD Level = Low)...
  267 Low RTD SKUs  (504 location-rows)
Writing Excel: Low_RTD_Level_12_Oct_2026.xlsx
  Writing inventory_detail sheet...
  Writing Sheet1 (pivot)...
  Writing sheet: PNQ (38 SKUs)...
  Writing sheet: GGN (54 SKUs)...
  Writing sheet: BLR (4 SKUs)...
  Writing sheet: MUM (4 SKUs)...
  Writing sheet: KOL (6 SKUs)...
  Writing sheet: HYD (6 SKUs)...
  Excel written: /app/Low_RTD_Level_12_Oct_2026.xlsx
Acquiring Microsoft 365 token...
Sending email to: kunal.Kumbhar@myfrido.com, navnath.k@myfrido.com, ...
CC: hemant.k@myfrido.com, ...
Email sent successfully!
```

---

## 19. Changing Recipients

Open [send_low_rtd_email.py](send_low_rtd_email.py) and edit the `TO_EMAILS` or `CC_EMAILS` lists:

```python
TO_EMAILS = [
    "kunal.Kumbhar@myfrido.com",
    # add new person here
]

CC_EMAILS = [
    "hemant.k@myfrido.com",
    # add new person here
]
```

Then re-deploy (see [Section 15](#15-how-to-re-deploy-step-by-step)).

---

## 20. Changing the Schedule

To update the schedule (e.g., add a 12:00 PM run):

```bash
# Update existing job schedule
gcloud scheduler jobs update http low-rtd-email-morning \
  --location asia-south1 \
  --project frido-429506 \
  --schedule "15 7 * * *" \
  --time-zone "Asia/Kolkata"

# Or create a new one
gcloud scheduler jobs create http "low-rtd-email-noon" \
  --location asia-south1 \
  --project frido-429506 \
  --schedule "0 12 * * *" \
  --time-zone "Asia/Kolkata" \
  --uri "https://asia-south1-run.googleapis.com/apis/run.googleapis.com/v1/namespaces/frido-429506/jobs/low-rtd-email-job:run" \
  --http-method POST \
  --oauth-service-account-email for-quickcomm-automation@frido-429506.iam.gserviceaccount.com \
  --attempt-deadline 600s
```

**Cron quick reference (Asia/Kolkata):**

| Time | Cron Expression |
|---|---|
| 7:15 AM | `15 7 * * *` |
| 9:00 AM | `0 9 * * *` |
| 12:00 PM | `0 12 * * *` |
| 4:15 PM | `15 16 * * *` |
| 6:00 PM | `0 18 * * *` |

---

## 21. Dependencies

| Package | Version | Purpose |
|---|---|---|
| `msal` | 1.31.0 | Microsoft Authentication Library — gets M365 OAuth token |
| `requests` | 2.32.3 | HTTP client — calls Microsoft Graph API |
| `pandas` | 2.2.3 | DataFrame manipulation for sorting and aggregating rows |
| `xlsxwriter` | 3.2.0 | Writes `.xlsx` with full formatting, colours, merged cells |
| `openpyxl` | 3.1.5 | Excel compatibility (used by pandas as engine) |

---

## 22. Credentials Reference

### Microsoft 365 (Azure AD App Registration)

| Property | Value |
|---|---|
| Azure App (Client) ID | `0ea585d7-9550-481b-8c4a-5cd2f5237975` |
| Tenant ID | `7e5ccb3a-f189-4e3a-95b2-5cfcc0e30f90` |
| Sender Mailbox | `reports@myfrido.com` |
| Permission Required | `Mail.Send` (application permission on Graph API) |

### Google Cloud

| Property | Value |
|---|---|
| Project ID | `frido-429506` |
| Service Account | `for-quickcomm-automation@frido-429506.iam.gserviceaccount.com` |
| SA Key File | `../sa_key.json` (outside this repo, never commit to git) |
| Roles needed | Cloud Run Developer, Cloud Scheduler Admin (for setup only) |

> ⚠️ **Never commit `sa_key.json` or any credentials to this repository.**

---

## 23. Troubleshooting

### Email not received

1. Check Cloud Run execution logs (Section 18)
2. Confirm the job ran: `gcloud run jobs executions list ...`
3. Check if Graph API returned non-202 — look for `Graph API error` in logs
4. Verify M365 token acquisition succeeded — look for `Failed to get token` in logs

### `No report generated — email not sent`

The inventory JSON returned zero SKUs with `rtdLevel == "Low"`. This is a valid outcome — it means all SKUs currently have sufficient or excess stock. No action needed.

### Job failed / timed out

- Default timeout is 600s (10 min). If the JSON fetch is slow, it may rarely hit this.
- Check network connectivity from Cloud Run to `frido-navigator.vercel.app`
- Retry manually: `gcloud run jobs execute low-rtd-email-job --region asia-south1 --wait`

### Cloud Scheduler not triggering

```bash
# List scheduler jobs and check STATE = ENABLED
gcloud scheduler jobs list --location asia-south1 --project frido-429506

# Force trigger immediately
gcloud scheduler jobs run low-rtd-email-morning --location asia-south1 --project frido-429506
```

### Re-deploying after code change doesn't take effect

Make sure `gcloud run jobs deploy` completed successfully. The image is rebuilt from source every time you run `deploy`. Confirm with:

```bash
gcloud run jobs describe low-rtd-email-job --region asia-south1 --project frido-429506
```

---

*Built and maintained by the Frido Data Analytics Team · reports@myfrido.com*
