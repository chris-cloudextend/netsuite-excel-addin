# Speed Reference: NetSuite Excel Add-in Optimization

This document explains the architecture and techniques used to achieve fast performance when pulling NetSuite data into Excel.

---

## The Challenge

NetSuite has several limitations that make real-time Excel integration challenging:

| Limitation | Impact |
|------------|--------|
| **Concurrency Limit** | Max 5 concurrent API requests |
| **1000-Row Limit** | SuiteQL queries return max 1000 rows |
| **Query Latency** | Each query takes 2-15 seconds |
| **Rate Limiting** | Too many requests = 429 errors |

Without optimization, refreshing 100 accounts × 12 months = **1,200 individual API calls** = hours of waiting.

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────────┐
│                           EXCEL ADD-IN                                  │
│  ┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐     │
│  │  Custom         │    │   Build Mode    │    │    Caching      │     │
│  │  Functions      │───▶│   Detection     │───▶│    Layer        │     │
│  │  (NS.GLABAL)    │    │   (Batching)    │    │  (3-tier)       │     │
│  └─────────────────┘    └─────────────────┘    └─────────────────┘     │
│                                   │                     │               │
│                                   ▼                     ▼               │
│                         ┌─────────────────────────────────┐             │
│                         │      Cloudflare Worker          │             │
│                         │      (CORS Proxy)               │             │
│                         └─────────────────────────────────┘             │
│                                        │                                │
│                                        ▼                                │
│                         ┌─────────────────────────────────┐             │
│                         │      Cloudflare Tunnel          │             │
│                         │   (localhost:5002 exposed)      │             │
│                         └─────────────────────────────────┘             │
│                                        │                                │
│                                        ▼                                │
│                         ┌─────────────────────────────────┐             │
│                         │      Python Backend             │             │
│                         │   - SuiteQL Query Builder       │             │
│                         │   - Pagination Handler          │             │
│                         │   - Backend Cache               │             │
│                         └─────────────────────────────────┘             │
│                                        │                                │
│                                        ▼                                │
│                         ┌─────────────────────────────────┐             │
│                         │         NetSuite                │             │
│                         │      (SuiteQL API)              │             │
│                         └─────────────────────────────────┘             │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## Speed Optimization Techniques

### 1. Intelligent Batching

**Problem:** Each `=NS.GLABAL()` formula triggers an API call.

**Solution:** Formulas don't call the API immediately. Instead, they queue requests and batch them together.

```javascript
// Without batching: 100 formulas = 100 API calls
=NS.GLABAL("4220", "Jan 2024", "Jan 2024")  → API call
=NS.GLABAL("4230", "Jan 2024", "Jan 2024")  → API call
=NS.GLABAL("4240", "Jan 2024", "Jan 2024")  → API call
// ... 97 more API calls

// With batching: 100 formulas = 1 API call
// All requests queued for 150ms, then sent as single batch:
POST /batch/balance
{
  "accounts": ["4220", "4230", "4240", ... 97 more],
  "periods": ["Jan 2024"]
}
```

**Result:** 100x fewer API calls

---

### 2. Build Mode (Drag Detection)

**Problem:** When user drags a formula across 12 columns, Excel creates 12 formulas nearly simultaneously. The 150ms batch timer fires before all formulas are created.

**Solution:** "Build Mode" detects rapid formula creation and defers ALL processing until the user stops dragging.

```javascript
// Detection criteria:
// - 3+ formulas created within 500ms triggers Build Mode
// - Build Mode shows #BUSY placeholder immediately
// - After 800ms of inactivity, processes all queued formulas at once

User drags formula across 12 months:
  → Formula 1: enters Build Mode, shows #BUSY
  → Formula 2-12: queued in Build Mode
  → User stops dragging
  → 800ms passes
  → Single optimized batch request for all 12 months
  → All cells update simultaneously
```

**Result:** Dragging feels instant, no partial updates

---

### 3. Full Year Refresh (Fast Path)

**Problem:** When dragging across 6+ months, even batched requests are slow because the backend runs separate queries per account type.

**Solution:** When Build Mode detects 6+ months for the same year, it uses `/batch/full_year_refresh` which:
1. Runs a single optimized query for ALL P&L accounts for the entire year
2. Returns ~200 accounts × 12 months in one response
3. Caches everything for instant subsequent lookups

```javascript
// Regular path: 
// - Separate query for Income accounts
// - Separate query for Expense accounts  
// - Separate Balance Sheet queries per period
// Total: 5-10 queries, 30-60 seconds

// Fast path (full_year_refresh):
// - Single query: all P&L accounts × all periods
// - Returns in 15-20 seconds
// - Caches ALL accounts, not just requested ones
```

**Result:** First row drag: 15-20 sec. Subsequent rows: INSTANT (from cache)

---

### 4. Three-Tier Caching

```
┌─────────────────────────────────────────────────────────────┐
│  TIER 1: In-Memory Cache (functions.js)                     │
│  - Fastest: microseconds                                     │
│  - Lost on page refresh                                      │
│  - Map: cache.balance.set(key, value)                       │
├─────────────────────────────────────────────────────────────┤
│  TIER 2: localStorage Cache                                  │
│  - Fast: milliseconds                                        │
│  - Persists across taskpane refreshes                        │
│  - TTL: 5 minutes                                            │
│  - Shared between taskpane and custom functions              │
├─────────────────────────────────────────────────────────────┤
│  TIER 3: Backend Cache (server.py)                          │
│  - Medium: avoids NetSuite query                             │
│  - TTL: 5 minutes                                            │
│  - balance_cache dict in Python                              │
└─────────────────────────────────────────────────────────────┘
```

**Cache Key Format:**
```
{account}:{period}:{subsidiary}:{department}:{location}:{class}
Example: "4220:Jan 2024:1:::"
```

**Result:** Second request for same data = instant (no API call)

---

### 5. Explicit Zero Caching

**Problem:** Accounts with $0 balance aren't returned by NetSuite (no rows = no transactions). This causes cache misses, which trigger new queries.

**Solution:** After `full_year_refresh`, explicitly cache `$0` for any requested period NOT in the response.

```javascript
// NetSuite returns:
{ "4220": { "Jan 2024": 50000, "Feb 2024": 45000 } }
// Note: Mar 2024 missing = $0 balance

// We explicitly cache:
cache.set("4220:Mar 2024", 0);  // Now cached as $0, not a miss
```

**Result:** Zero-balance accounts resolve instantly instead of triggering re-queries

---

### 6. Backend Pagination

**Problem:** NetSuite SuiteQL returns max 1000 rows per query. Full year data often exceeds this.

**Solution:** Backend uses API-level pagination with `limit` and `offset` parameters:

```python
def run_paginated_suiteql(query, page_size=1000):
    all_rows = []
    offset = 0
    while True:
        url = f"{base_url}?limit={page_size}&offset={offset}"
        response = query_netsuite(url)
        all_rows.extend(response)
        if len(response) < page_size:
            break  # Last page
        offset += page_size
    return all_rows
```

**Result:** Can retrieve 3000+ rows across multiple pages

---

### 7. Period Chunking

**Problem:** Requesting too many periods at once causes Cloudflare 524 timeout (>100 seconds).

**Solution:** Split periods into chunks of 3:

```javascript
// Instead of: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
// Chunk into: [["Jan", "Feb", "Mar"], ["Apr", "May", "Jun"], ["Jul", "Aug", "Sep"], ["Oct", "Nov", "Dec"]]
```

**Result:** Each request completes within timeout limits

---

### 8. Account Title Preloading

**Problem:** When formulas recalculate, each `=NS.GLATITLE()` call hits NetSuite, causing 429 concurrency errors.

**Solution:** On taskpane load, fetch ALL account titles in one batch:

```python
# Backend endpoint: /account/preload_titles
# Returns: {"4220": "Revenue - Products", "4230": "Revenue - Services", ...}
# Frontend caches all 400+ titles on startup
```

**Result:** Account titles never hit NetSuite after initial load

---

## Performance Summary

| Scenario | Before Optimization | After Optimization |
|----------|--------------------|--------------------|
| Single formula | 2-5 sec | 2-5 sec (no change) |
| 20 formulas (batch) | 40-100 sec | 5-10 sec |
| Drag 12 months | 60-180 sec + timeouts | 15-20 sec first row, instant after |
| Full sheet (200 accounts × 12 months) | Hours + many errors | 30-60 sec |
| Second refresh (cached) | Same as first | Instant |

---

## Key Files

| File | Purpose |
|------|---------|
| `docs/functions.js` | Custom functions, batching, Build Mode, caching |
| `docs/taskpane.html` | UI, refresh controls, cache management |
| `backend/server.py` | SuiteQL queries, pagination, backend cache |
| `CLOUDFLARE-WORKER-CODE.js` | CORS proxy configuration |

---

## Configuration Constants

```javascript
// functions.js
const BATCH_TIMER_MS = 150;           // Wait before processing batch
const BUILD_MODE_THRESHOLD = 3;        // Formulas to trigger Build Mode
const BUILD_MODE_WINDOW_MS = 500;      // Time window for detection
const BUILD_MODE_SETTLE_MS = 800;      // Wait after last formula
const MAX_ACCOUNTS_PER_BATCH = 50;     // Chunk size for accounts
const MAX_PERIODS_PER_BATCH = 3;       // Chunk size for periods
const CACHE_TTL_MS = 5 * 60 * 1000;   // 5 minute cache TTL
```

---

## Future Improvements

1. **WebSocket Connection** - Real-time updates without polling
2. **Service Worker** - Offline caching of static lookups
3. **Incremental Refresh** - Only fetch changed periods
4. **Query Deduplication** - Detect identical pending requests
5. **Predictive Prefetch** - Anticipate next period based on patterns

---

*Document created: December 2024*
*Add-in Version: 3.0.5.77*

