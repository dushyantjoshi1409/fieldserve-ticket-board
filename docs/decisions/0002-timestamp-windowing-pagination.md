# ADR-002: Windowed Timestamp Pagination for API Limits

## Context
The FieldServe Helpdesk API enforces two hard restrictions:
1. `page > 10` returns `HTTP 400 Bad Request`.
2. `per_page` maximum is 50.
Therefore, any single standard paginated query can return at most $10 \times 50 = 500$ tickets. However, the total dataset contains over 1,000 live tickets.

Additionally, the server strictly accepts timestamps in the exact format `YYYY-MM-DDTHH:MM:SSZ` and rejects ISO strings containing sub-second precision (`.001000Z`).

## Decision
We implemented a **Sliding Timestamp Window** pagination strategy:
1. Start fetching with `updated_since = current_since` (initially `1970-01-01T00:00:00Z` or `last_sync - 60s`).
2. Page through $1 \le \text{page} \le 10$.
3. When the 500-ticket cap is hit, extract the `updated_at` timestamp from the 500th ticket.
4. Advance `current_since = last_ticket.updated_at`. If all 500 tickets share the exact same timestamp second, advance by $+1\text{ second}$ using `timedelta(seconds=1)`.
5. Format all timestamps strictly as `%Y-%m-%dT%H:%M:%SZ`.
6. Repeat until a window returns fewer than 500 tickets.

## Alternatives Considered
- **Offset-based pagination**: Rejected because the API does not support `offset` or `page > 10`.
- **Search endpoint**: Rejected because no search/filter endpoint is exposed by the API.

## Consequences
- Successfully enumerates all 1,007+ tickets in ~3–4 seconds across 3 compact windows.
- Stays well within the 60 requests/minute rate limit budget (uses ~21 requests on full fetch, 1–2 requests on incremental sync).
- Protects against clock skew via a 60-second backward buffer during incremental syncs.
