# FieldServe Ticket Board — Architecture & Two-Way Sync Engine

> **Submission for FieldServe Take-Home Assignment**  
> **Author**: Dushyant Joshi  
> **Stack**: Python 3.9+, `openpyxl`, `requests`  
> **Interface**: Fully compliant with `submission.json` automated grader contract  

---

## 1. Executive Summary & Problem Context

FieldServe installs, repairs, and maintains electronic security infrastructure (CCTV surveillance, intrusion alarms, access control) across retail branches of **OmniMart**.

- **The Operational Challenge**: The helpdesk system records tickets via a REST API, but the field operations team works primarily out of spreadsheets. Crucially, raw tickets do not answer the real operational questions:
  1. *What specific step is this ticket blocked on?*
  2. *Whose move is it right now (Action Owner)?*
  3. *Has paperwork been skipped or falsified with junk text?*
- **Our Solution**: A robust, bidirectional synchronization engine (`sync.py`) bridging the FieldServe REST API and an operational Excel spreadsheet (`tickets.xlsx`), powered by a 6-step sequential workflow state machine, 3-way merge conflict resolution, automated data cleaning heuristics, and an automated CLI (`sheet.py`) matching the grader contract.

---

## 2. System Architecture & Two-Way Sync

```
                     ┌─────────────────────────────────────────────────┐
                     │          FieldServe Helpdesk REST API           │
                     │          http://13.233.55.184:8377              │
                     │   GET /agents │ GET /tickets │ PUT /tickets/{id}│
                     └────────────────────────┬────────────────────────┘
                                              │ ▲
                   1. Windowed GET (500/batch)│ │ 5. Targeted PUT (Payload)
                      Rate-limit 429 backoff  │ │    Mapped status & engineer
                                              ▼ │
                     ┌─────────────────────────────────────────────────┐
                     │          Two-Way Sync Engine (sync.py)          │
                     │                                                 │
                     │   [Fetcher]  ──> [Parser & Noise Filter]        │
                     │       │                      │                  │
                     │       ▼                      ▼                  │
                     │   [3-Way Diff] <── [Workflow State Machine]     │
                     │       │                      │                  │
                     │       ▼                      ▼                  │
                     │   [Conflict Resolver] ──> [Integrity Auditor]   │
                     └────────────────────────┬────────────────────────┘
                                              │ ▲
                         6. Formatted Rewrite │ │ 4. Read User Edits
                                              ▼ │
┌──────────────────────────────────────────────────────────────────────────────────────┐
│                              Spreadsheet (tickets.xlsx)                              │
│                                                                                      │
│  Tab 1: "Tickets" (1,007 rows with Stage, Action Owner, Store Code, Flags)           │
│  Tab 2: "Summary" (Aggregated workload count per person)                             │
│                                                                                      │
│  Controlled via Grader Interface:                                                    │
│    python3 sheet.py export /path/to/out.csv                                          │
│    python3 sheet.py apply  /path/to/edits.csv                                        │
└──────────────────────────────────────────────────────────────────────────────────────┘
```

### How the Sync Works Cycle-by-Cycle
1. **Load State Snapshot**: Reads `local_state.json` containing the base snapshot from the previous sync cycle (`last_sync` timestamp, agent roster, and raw field values).
2. **Fetch Roster**: Retrieves active agents from `GET /api/v2/agents`. Categorizes Coordinators, Managers, Engineers, and Finance.
3. **Fetch Remote Tickets (Windowed)**:
   - If initializing or sheet is missing: performs a full fetch from epoch (`1970-01-01T00:00:00Z`).
   - If incremental: fetches tickets with `updated_since = last_sync - 60s` (buffer for server clock skew).
4. **Load Local Sheet**: Reads current cell values from `tickets.xlsx` via `openpyxl`.
5. **3-Way Diff & Push**: Compares $\text{Base}$ vs $\text{Local (Sheet)}$ vs $\text{Server (API)}$ on all writable fields. Local modifications are converted to API payloads and pushed via `PUT /api/v2/tickets/{id}`.
6. **Sheet Regeneration**: Writes all current tickets to `tickets.xlsx` with recomputed Stages, Action Owners, Store Codes, and Integrity Flags, and builds the `Summary` workload tab.
7. **Persist Snapshot**: Writes the updated state to `local_state.json`.

---

## 3. Conflict Resolution Policy (§6)

When both the API and the spreadsheet change concurrently, data must never be silently clobbered. We track a **3-Way Merge** across all writable fields (`Status`, `Engineer`, `Manager`, `Purchase Order`, `Delivery Note`, `Work Report`, `Invoice`):

- **$\text{Base}(T)$**: Snapshot of the field value at the conclusion of the last sync.
- **$\text{Local}(T)$**: Current field value in `tickets.xlsx`.
- **$\text{Server}(T)$**: Current field value returned by the API.

### Decision Matrix

| Local Modified? ($\Delta_{\text{local}}$) | Server Modified? ($\Delta_{\text{server}}$) | Values Identical? | Resolution / Action |
|:---:|:---:|:---:|---|
| **No** | **No** | Yes | **No-op**: No data changed. Zero API requests. |
| **Yes** | **No** | — | **Push to Server**: Human edited sheet. Send `PUT /api/v2/tickets/{id}` with transformed payload. |
| **No** | **Yes** | — | **Pull to Sheet**: Server changed (e.g., engineer updated on mobile). Update sheet cell. |
| **Yes** | **Yes** | Yes | **Clean Merge**: Both sides made the same edit. Update base snapshot. |
| **Yes** | **Yes** | No | **Conflict — Server Wins**: Both sides changed differently. Server value takes precedence. |

### Why "Server Wins"?
- Field engineers and customer store managers frequently submit updates live on-site from mobile devices or email. Overwriting live field reports from an asynchronous desktop spreadsheet would clobber fresh field operational data.
- All conflicts are logged to `stdout` and saved under `local_state.json["conflicts"]` for auditability.
- **Critical Guarantee**: An unchanged sheet cell ($\Delta_{\text{local}} = \text{False}$) **never** overwrites a fresh server-side change.

### PUT Payload Mapping
- **Status**: Maps display strings (`Open`, `Pending`, `Resolved`, `Closed`) $\rightarrow$ numeric codes (`2`, `3`, `4`, `5`).
- **Engineer**: Maps exact name string (e.g., `"Barnelbel Panmo"`) $\rightarrow$ `engineer_id` (`9964`) via `/agents`. Blank string maps to `null`.
- Only changed fields are included in the JSON payload (sparse updates).

---

## 4. API Limitations & Handling Strategies (§5)

| Limitation | Impact | Implementation Strategy |
|---|---|---|
| **Page Cap (Max 10 Pages)** | `page > 10` returns `HTTP 400`. With `per_page=50`, only 500 tickets can be returned in a single listing. | **Pagination Windowing**: We query pages 1–10. When 500 tickets are reached, we extract the `updated_at` of the 500th ticket and advance `updated_since = last_ticket.updated_at`. If all 500 tickets share the same second, we advance by `+1s`. This successfully enumerates all 1,007+ tickets in ~3–4 seconds. |
| **Strict Timestamp Format** | Timestamps with microseconds (`.001000Z`) return `HTTP 400: updated_since must look like 2026-01-31T00:00:00Z`. | We strictly format all timestamps using `strftime("%Y-%m-%dT%H:%M:%SZ")`. Microseconds are stripped before submission. |
| **Rate Limiting (60 req/min)** | Rapid calls return `HTTP 429 Too Many Requests`. | The client inspects the `Retry-After` header, sleeps `retry + 1` seconds, and retries with backoff up to 5 attempts. Both GET and PUT operations are protected. |
| **No Search / Filtering** | Cannot filter by status or store code server-side. | Client-side classification and filtering in Python. |

---

## 5. Workflow State Machine & Integrity Rules (§2 & §3)

A ticket owes 6 strictly ordered steps. The **first unmet step** defines the ticket's **Stage**, and its owner is the **Action Owner**:

```
Step 1: Assign Manager
   │  Met when: manager is non-empty string
   │  Action Owner: Coordinator by ticket_id % 2 (sorted alphabetically by name)
   ▼
Step 2: Assign Engineer
   │  Met when: engineer_id is set (> 0)
   │  Action Owner: Ticket's Manager
   ▼
Step 3: Purchase Order (PO)
   │  Met when: purchase_order is a valid URL (Waivable: NO)
   │  Action Owner: Ticket's Manager
   ▼
Step 4: Delivery Note (DN)
   │  Met when: delivery_note is a valid URL or exact text 'NA' (Waivable: YES)
   │  Action Owner: Finance (Rogelnor Nelkin)
   ▼
Step 5: Work Report (WR)
   │  Met when: work_report is a valid URL (Waivable: NO)
   │  Action Owner: Assigned Field Engineer
   ▼
Step 6: Invoice
   │  Met when: invoice is a valid URL or exact text 'NA' (Waivable: YES)
   │  Action Owner: Finance (Rogelnor Nelkin)
   ▼
Step 7: Done (Owner is blank)
```

### Document Validation & Junk Text Handling
- A document field is valid **only** if it is a well-formed HTTP/HTTPS URL (`http://` or `https://` with a valid domain and netloc) or the exact string `'NA'` where allowed.
- **Junk Text**: Real-world data contains entries like `"done"`, `"ok"`, `"asked Ravi"`, `"yes"`, `"uploading tonight"`, `"sent on mail"`. These are rejected by `is_doc_met()` and the stage is blocked.
- **Waiver Rules**:
  - `purchase_order`: Non-waivable. (`'NA'` is invalid).
  - `delivery_note`: Waivable. Exact text `'NA'` satisfies the step.
  - `work_report`: Non-waivable under any circumstances.
  - `invoice`: Waivable. Exact text `'NA'` satisfies the step.

### Integrity Rules
Surfaced in the `Integrity Flag` column:
1. **`Closed but paperwork missing`**: Ticket has status `5` (Closed) but one or more of the 6 steps remain unmet.
2. **`Resolved without Work Report`**: Ticket has status `4` (Resolved) but lacks a verified Work Report URL.
3. **`Manager not in roster`**: Ticket manager name is not recognized in `/agents`.

---

## 6. Messy Data Extraction (§4)

### Store Code Extraction
Store codes are 6-character alphanumeric identifiers (e.g., `K7Q2XR`, `M9GDCU`, `HHV8WY`, `DKCKJT`). We extract them through a tiered heuristic:
1. **Requester Email**: Regex `store.([a-zA-Z0-9]{6})@omnimart.example` (Highest confidence).
2. **Requester Name**: Parenthetical token, e.g., `OmniMart Rarsi (M9GDCU)`.
3. **Subject Parentheses**: e.g., `PM AT OMNIMART TARVENO (K7Q2XR)`.
4. **Alphanumeric Tokens**: 6-character tokens near `OMNIMART` or `STORE`, filtered against common 6-letter English words (`CAMERA`, `SYSTEM`, `REPORT`, etc.).

### Work Type Classification
- **Maintenance**: Subject contains `PM` (Preventive Maintenance) or `AMC`.
- **Installation**: Subject contains `INSTALL` or `INSTALLATION`.
- **Service Call**: Subject indicates break-fix (`NOT WORKING`, `FAULT`, `DOWN`, `COMPLAINT`, `CAMERA`, `FOOTAGE`, `OFFLINE`).
- **Other**: Administrative, payment, vendor billing, or generic tickets.

### Noise Detection
- Tickets containing spam keywords (`offer`, `listing`, `business with`, `sell smart`, `earn more`, `licenses`, `lottery`, `promo`) or automated email replies (`automatic reply:`, `out of office`) are classified as `Is Noise: YES`.
- **Retention Decision**: Noise tickets are retained on the operational sheet to prevent silent data loss and remain subject to the exact same workflow rules (§7). They can be filtered out easily in Excel.

---

## 7. Automated Grader Contract (§7)

### `submission.json`
Located at repo root:
```json
{
  "sync_cmd": "python3 sync.py",
  "sheet_cmd": "python3 sheet.py"
}
```

### Grader CLI Interface (`sheet.py`)
- **Export**: `python3 sheet.py export /path/to/out.csv`
  - Reads the `Tickets` tab from `tickets.xlsx`.
  - Resolves all formulas (`values_only=True`).
  - Exports all 11 required columns with exact names:
    `Ticket ID`, `Subject`, `Status`, `Engineer`, `Manager`, `Purchase Order`, `Delivery Note`, `Work Report`, `Invoice`, `Stage`, `Action Owner`.
- **Apply**: `python3 sheet.py apply /path/to/edits.csv`
  - Accepts CSV with columns: `Ticket ID,Column,Value`.
  - Modifies the targeted cells in `tickets.xlsx` as if typed by a human.
  - **Does NOT trigger sync as a side effect**.
- **Sync**: `python3 sync.py`
  - Performs one complete two-way sync cycle and exits with code `0`.

---

## 8. Summary Workload View

In addition to the main `Tickets` tab, `tickets.xlsx` contains a dedicated **`Summary`** tab detailing the current queue per individual:

```text
Action Owner         : Tickets Waiting
---------------------------------------
Nefo Zagi            : 401
Pahifan Kinga        : 202
Rogelnor Nelkin      : 159  (Finance)
Farho Kelpar         : 50   (Coordinator)
Nikintel Zarfi       : 45   (Coordinator)
Kisanmor Huvor       : 9
Felboso Hamar        : 6
Melsar Marfu         : 6
Mihorzan Dinme       : 5
Borga Bosor          : 5
Parrar Vapo          : 5
Ninmarpi Gelban      : 3
Delhinma Belbor      : 2
Pekan Parmel         : 2
Ninfindor Gelgor     : 2
Pudovor Ninin        : 1
```

---

## 9. Assumptions Made

1. **Coordinator Parity**: Coordinators are sorted alphabetically by name. Ticket ID parity `ticket_id % 2` determines assignment (0 $\rightarrow$ `Farho Kelpar`, 1 $\rightarrow$ `Nikintel Zarfi`).
2. **Finance Team Mailbox**: Since Finance operates from a shared mailbox, the single Finance agent from `/agents` (`Rogelnor Nelkin`) is assigned as the Action Owner for `Delivery Note` and `Invoice` stages.
3. **Manager Matching**: `manager` is stored as a string on tickets. During push, if an engineer name does not match any agent in the roster, the update is skipped with an error log to avoid `HTTP 422`.
4. **Authoritative Server**: In cases of conflicting concurrent edits, server-side data is preferred to protect on-site field entries from desktop overwrites.
5. **No Client-side Ticket Creation**: Tickets originate exclusively from the Helpdesk API; the sheet does not create new ticket IDs.

---

## 10. Verification & Test Suite

Run the complete test suite:
```bash
python3 test_all.py
```

### Test Coverage (100% Passing):
- `test_doc_validation`: Valid URLs, waivable/non-waivable `'NA'`, junk strings (`"done"`, `"asked Ravi"`, `"ok"`).
- `test_stage_logic`: All 6 steps in sequence, coordinator parity by ticket ID modulo 2.
- `test_junk_handling_in_stages`: Verifies junk text blocks stage progression.
- `test_integrity_rules`: Detects Closed tickets with missing paperwork and Resolved tickets missing Work Reports.
- `test_store_code_and_work_type`: Store code extraction from emails, names, subjects, and work type classification.
- `test_three_way_merge_logic`: Verifies 3-way merge truth table and conflict policy.
- `test_sheet_contract`: End-to-end test of `sheet.py export` and `sheet.py apply`.
- `test_submission_contract`: Validates `submission.json` format and executability.

---

## 11. Quickstart Commands

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Set environment variables (optional, defaults provided)
export FIELDOPS_BASE_URL="http://13.233.55.184:8377"
export FIELDOPS_API_KEY="fk_90cf43b35e8eeaa60a0b6cf1"

# 3. Run test suite
python3 test_all.py

# 4. Run two-way sync
python3 sync.py

# 5. Export sheet to CSV
python3 sheet.py export /tmp/out.csv

# 6. Apply sample edit and re-sync
python3 sheet.py apply /tmp/edits.csv
python3 sync.py
```

---

## 12. What I Would Build Next With More Time

1. **Google Sheets Real-time Webhooks**: Connect via Google Sheets API v4 with Apps Script webhooks to trigger immediate delta-syncs upon cell modification rather than periodic batch polling.
2. **Exponential Jitter & Circuit Breaker**: Introduce randomized jitter to HTTP retry mechanisms to mitigate thundering herd behavior during rate limit resets.
3. **Persistent SQLite Store**: Transition `local_state.json` to an embedded SQLite database for ACID compliance and transactional history at larger scale.
4. **Operational Alerting**: Webhook integration with Slack/PagerDuty for tickets flagged with integrity violations (e.g., Closed without invoices).
