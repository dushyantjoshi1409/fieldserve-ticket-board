# ADR-001: 3-Way Merge with "Server Wins" Conflict Policy

## Context
FieldServe's operations team edits ticket fields in spreadsheets (assigning managers/engineers, pasting document URLs, changing statuses), while field engineers and customer store managers update tickets simultaneously via the Helpdesk API (mobile app, automated emails). 

A naive last-write-wins or full-row overwrite would inadvertently clobber concurrent updates made on either side. Specifically, an unchanged cell in a user's spreadsheet could wipe out a fresh status update submitted from an engineer's phone.

## Decision
We implemented a field-level **3-Way Merge** engine backed by a persistent snapshot state (`local_state.json`):
1. **Base ($B$)**: State of each ticket's writable fields at the last sync.
2. **Local ($L$)**: Current state read from `tickets.xlsx`.
3. **Server ($S$)**: Current state fetched from `GET /api/v2/tickets`.

### Conflict Rule:
- If only $L$ changed ($L \neq B, S = B$): **Push to Server** via `PUT /api/v2/tickets/{id}`.
- If only $S$ changed ($S \neq B, L = B$): **Pull to Sheet** (sheet adopts server value).
- If both changed to the same value ($L = S$): Clean merge.
- If both changed to different values ($L \neq S$): **Conflict — Server Wins**. Server value is adopted and the conflict is logged to `local_state.json["conflicts"]`.
- An unchanged sheet cell ($L = B$) **never** overwrites a fresh server change ($S \neq B$).

## Alternatives Considered
- **Client Wins (Sheet Wins)**: Rejected because field engineers on site have ground-truth knowledge; overwriting mobile updates from a delayed desktop spreadsheet causes field coordination failures.
- **Timestamp / Last-Write-Wins**: Rejected because cell edits in local Excel files lack per-field timestamps, making reliable sub-field timestamp comparison impossible.
- **Fail / Abort on Conflict**: Rejected because it halts the entire batch synchronization for 1,000+ tickets due to an issue on a single ticket.

## Consequences
- Zero data loss for mobile/field updates.
- Transparent conflict auditing via console logs and `local_state.json`.
- Strict compliance with assignment section 6 and section 7.
