# Project: FieldServe Ticket Board (Two-Way Sync Engine)

## What this product does
A robust, bi-directional bridge between FieldServe's Helpdesk REST API and an operational Excel spreadsheet (`tickets.xlsx`). It answers the core operational question: *"What is this ticket waiting on, and whose move is it right now?"* by applying a 6-step sequential workflow state machine, 3-way merge conflict resolution, automated junk paperwork detection, and store code heuristics.

## Stack
- **Language**: Python 3.9+
- **HTTP Client**: `requests` (with HTTP Basic Auth & exponential 429 retry backoff)
- **Spreadsheet Engine**: `openpyxl` (data-only evaluation, multi-sheet support)
- **State Persistence**: `local_state.json` (3-way merge snapshot storage)
- **Test Framework**: Native Python test suite (`test_all.py`)

## Rules
- **Never commit secrets**: All keys configured via `FIELDOPS_API_KEY` and `FIELDOPS_BASE_URL` (documented in `.env.example`).
- **Server is Source of Truth**: On concurrent conflicting edits, server-side data wins to protect live field updates.
- **Unchanged cells never overwrite**: An unedited spreadsheet cell must never overwrite a fresh server change.
- **Strict ISO-8601 formatting**: API requires `YYYY-MM-DDTHH:MM:SSZ` without microseconds.
- **Grader Contract Compliance**:
  - `sheet.py export <path>` must resolve formulas and produce exact required column names.
  - `sheet.py apply <path>` must write cell values without triggering sync as a side-effect.
  - `sync.py` must run one full sync cycle and exit with code 0.
- **Verify before declaring done**: Run full test suite (`python3 test_all.py`), export CSV, and verify integrity flags.

## Commands
- **Test Suite**: `python3 test_all.py`
- **Two-Way Sync**: `python3 sync.py`
- **Export CSV**: `python3 sheet.py export /tmp/out.csv`
- **Apply Edits**: `python3 sheet.py apply /tmp/edits.csv`

## Architecture Decisions (ADRs)
1. **ADR-001: 3-Way Merge with Server Wins**: Resolves concurrent edits between desktop spreadsheet users and mobile field engineers without data loss.
2. **ADR-002: Windowed Timestamp Pagination**: Overcomes the 10-page / 500-ticket API limit by advancing `updated_since` windows.
3. **ADR-003: Pure-Python Heuristics for Messy Data**: Multi-tiered regex extraction for 6-character store codes and work types without heavyweight external ML dependencies.
