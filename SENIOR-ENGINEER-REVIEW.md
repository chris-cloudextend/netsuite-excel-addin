# Excel Add-in Performance Analysis - Senior Engineer Review Request

## Executive Summary

We successfully completed **Phase 1** of the Excel add-in architecture migration (converting from streaming to non-streaming async functions), which eliminated the auto-recalculation (`@` symbol) issue. However, we're now facing a **NetSuite SuiteQL performance bottleneck** that causes a 6-8 minute refresh time for 100 accounts × 12 periods.

**Request:** Review our SuiteQL queries and batching strategy to identify optimization opportunities.

---

## Phase 1 Accomplishments ✅

1. **Converted custom functions from `@streaming` to standard `async` functions**
   - `NS.GLABAL(account, fromPeriod, toPeriod, subsidiary, dept, location, class)` - Get GL account balance
   - `NS.GLABUD(account, fromPeriod, toPeriod, subsidiary, dept, location, class)` - Get GL budget amount
   - `NS.GLATITLE(account)` - Get account name
   - `NS.GLACCTTYPE(account)` - Get account type (Income, Expense, etc.)
   - `NS.GLAPARENT(account)` - Get parent account

2. **Implemented non-volatile functions (`volatile: false`)**
   - Formulas no longer auto-recalculate on sheet open ✅
   - Users explicitly trigger refresh via task pane buttons ✅

3. **Implemented client-side caching**
   - Results cached in JavaScript `Map` objects
   - Cache cleared before each refresh to ensure fresh data
   - Instant loading on subsequent sheet opens ✅

4. **Implemented batching mechanism**
   - Queue requests during Excel's calculation phase
   - Process in batches to reduce API calls
   - Multiple batching strategies tested (see below)

---

## Current Performance Issue ⚠️

### Observed Behavior:
- **Sheet:** ~100 accounts × 12 months = 1,200 formulas
- **Refresh Time:** 6-8 minutes
- **Backend Response Time:** 
  - 6 accounts × 12 periods = 226 seconds (4 minutes!)
  - 10 accounts × 1 period = 19 seconds
  - **~2-3 seconds per account-period combination**

### User Experience:
- First daily refresh: 6-8 minutes ☕
- Subsequent opens: Instant (cached) ✅
- No `@` symbols on open ✅
- Predictable, progressive updates (month by month) ✅

---

## SuiteQL Queries (Backend Implementation)

### 1. Batch Balance Query (`/batch/balance` endpoint)

**Request Format:**
```json
{
  "accounts": ["4220", "60010", "60040", ...],  // Array of account numbers
  "periods": ["Jan 2024", "Feb 2024", ...],     // Array of period names
  "subsidiary": "",                              // Optional filter
  "department": "",                              // Optional filter
  "location": "",                                // Optional filter
  "class": ""                                    // Optional filter
}
```

**Response Format:**
```json
{
  "balances": {
    "4220": {
      "Jan 2024": 401569.18,
      "Feb 2024": 301881.19,
      ...
    },
    "60010": {
      "Jan 2024": 4277813.58,
      ...
    }
  }
}
```

### 2. P&L Account Query (Activity for Period Range)

```sql
WITH sub_info AS (
    SELECT 
        COALESCE({target_subsidiary}, parent) AS target_sub,
        (SELECT COUNT(*) FROM Subsidiary WHERE isinactive = 'F') AS sub_count
    FROM Subsidiary
    WHERE id = COALESCE({target_subsidiary}, parent)
    AND parent IS NULL
    LIMIT 1
)
SELECT 
    a.acctnumber,
    a.accttype,
    ap.periodname,
    SUM(
        CASE 
            WHEN sub_info.sub_count > 1 THEN 
                TO_NUMBER(
                    BUILTIN.CONSOLIDATE(
                        tal.amount,
                        'LEDGER',
                        'DEFAULT',
                        'DEFAULT',
                        sub_info.target_sub,
                        t.postingperiod,
                        'DEFAULT'
                    )
                )
            ELSE 
                tal.amount
        END
        * 
        CASE 
            WHEN a.accttype IN ('Income', 'OthIncome') THEN -1 
            ELSE 1 
        END
    ) AS total_amount
FROM TransactionAccountingLine tal
CROSS JOIN sub_info
INNER JOIN Transaction t ON t.id = tal.transaction
INNER JOIN Account a ON a.id = tal.account
INNER JOIN AccountingPeriod ap ON ap.id = t.postingperiod
WHERE t.posting = 'T'
  AND tal.posting = 'T'
  AND tal.accountingbook = 1
  AND ap.periodname IN ('Jan 2024', 'Feb 2024', 'Mar 2024', ...)
  AND a.acctnumber IN ('4220', '60010', '60040', ...)
  AND a.accttype IN ('Income', 'COGS', 'Cost of Goods Sold', 'Expense', 'OthIncome', 'OthExpense')
  AND COALESCE(a.eliminate, 'F') = 'F'
  -- Optional filters:
  AND (t.subsidiary = {subsidiary} OR {subsidiary} IS NULL)
  AND (tal.department = {department} OR {department} IS NULL)
  AND (tal.location = {location} OR {location} IS NULL)
  AND (tal.class = {class} OR {class} IS NULL)
GROUP BY a.acctnumber, a.accttype, ap.periodname
ORDER BY a.acctnumber, ap.periodname
```

**Key Points:**
- Uses `BUILTIN.CONSOLIDATE` when multiple subsidiaries exist
- Consolidates to top-level parent subsidiary by default
- Handles multi-currency conversion automatically
- Signs adjusted for Income accounts (negative → positive)

### 3. Balance Sheet Account Query (Cumulative Balance to Period End)

```sql
WITH sub_info AS (
    -- Same as P&L query
),
period_dates AS (
    SELECT id, periodname, enddate
    FROM AccountingPeriod
    WHERE periodname IN ('Jan 2024', 'Feb 2024', ...)
      AND isquarter = 'F'
      AND isyear = 'F'
)
SELECT 
    a.acctnumber,
    a.accttype,
    pd.periodname,
    SUM(
        CASE 
            WHEN sub_info.sub_count > 1 THEN 
                TO_NUMBER(
                    BUILTIN.CONSOLIDATE(
                        tal.amount,
                        'LEDGER',
                        'DEFAULT',
                        'DEFAULT',
                        sub_info.target_sub,
                        pd.id,
                        'DEFAULT'
                    )
                )
            ELSE 
                tal.amount
        END
    ) AS total_amount
FROM TransactionAccountingLine tal
CROSS JOIN sub_info
CROSS JOIN period_dates pd
INNER JOIN Transaction t ON t.id = tal.transaction
INNER JOIN Account a ON a.id = tal.account
WHERE t.posting = 'T'
  AND tal.posting = 'T'
  AND tal.accountingbook = 1
  AND t.trandate <= pd.enddate  -- Cumulative up to period end
  AND a.acctnumber IN ('15000-1', '10100', ...)
  AND a.accttype NOT IN ('Income', 'COGS', 'Cost of Goods Sold', 'Expense', 'OthIncome', 'OthExpense')
  -- Optional filters (same as P&L)
GROUP BY a.acctnumber, a.accttype, pd.periodname
ORDER BY a.acctnumber, pd.periodname
```

**Key Differences from P&L:**
- Uses `t.trandate <= pd.enddate` instead of `ap.periodname IN (...)`
- Calculates cumulative balance from inception to period end
- No sign adjustment (Balance Sheet accounts don't invert)

---

## Why BUILTIN.CONSOLIDATE is Slow

### Technical Analysis:

1. **Multi-Currency Conversion:**
   - `BUILTIN.CONSOLIDATE` performs real-time currency conversion for each transaction line
   - Fetches exchange rates for each period
   - Applies conversion at the transaction level

2. **Multi-Subsidiary Aggregation:**
   - When `sub_count > 1`, it must aggregate across subsidiary hierarchy
   - Resolves intercompany eliminations
   - Applies consolidation rules per period

3. **Per-Transaction Processing:**
   - Function called for **every row** in the result set
   - If 1 account has 1,000 transactions in 1 period, it's called 1,000 times
   - Then aggregated with `SUM()`

4. **Cross Join with Periods:**
   - In Balance Sheet query, we `CROSS JOIN period_dates`
   - If we request 12 periods, the `BUILTIN.CONSOLIDATE` function is called 12× more

### Performance Test Results:

| Test | Accounts | Periods | Time | Per Account-Period |
|------|----------|---------|------|-------------------|
| Test 1 | 6 | 12 | 226s | ~3.14s |
| Test 2 | 10 | 1 | 19s | ~1.90s |
| Test 3 | 50 | 1 | ~19s* | ~0.38s |

*Estimated based on Test 2

**Extrapolation for Production:**
- 100 accounts × 12 periods = 1,200 combinations
- At 2-3 seconds each = **40-60 minutes** without batching
- Current batching: **6-8 minutes** (~8× improvement)

---

## Batching Strategies Tested

### Strategy 1: Group by Filters + Period (v1.0.0.93 - Current)

**Frontend Logic:**
```javascript
// Group requests by filters AND period
const filterKey = JSON.stringify({
    subsidiary: params.subsidiary || '',
    department: params.department || '',
    location: params.location || '',
    class: params.classId || '',
    fromPeriod: params.fromPeriod || '',  // ← Period included
    toPeriod: params.toPeriod || ''        // ← Period included
});
```

**Result:**
- 12 separate batches (one per month)
- Each batch: 100 accounts × 1 period
- Split into 2 chunks of 50 accounts each
- API calls: 2 per month × 12 months = **24 API calls**
- Time: ~38 seconds per month × 12 = **~7-8 minutes**

**Pros:**
- Predictable, progressive updates (column by column)
- User sees visible progress
- Avoids massive multi-period timeouts
- Each query is relatively fast (~19s)

**Cons:**
- Many API calls (24)
- Total time still high (7-8 minutes)

---

### Strategy 2: Group by Filters Only (v1.0.0.92 - Failed)

**Frontend Logic:**
```javascript
// Group requests by filters ONLY (NOT period)
const filterKey = JSON.stringify({
    subsidiary: params.subsidiary || '',
    department: params.department || '',
    location: params.location || '',
    class: params.classId || ''
    // ← NO period grouping
});

// Collect ALL unique periods across all requests
const allPeriods = [...new Set(groupRequests.flatMap(r => [r.fromPeriod, r.toPeriod]))];
// Backend receives: ["Jan 2024", "Feb 2024", ..., "Dec 2024"]
```

**Result:**
- 1 batch for all filters
- Backend receives: 50 accounts × 12 periods
- API calls: 2 chunks = **2 API calls**
- Time: 226 seconds (4 minutes) for just 6 accounts!

**Why it Failed:**
- `CROSS JOIN period_dates` in Balance Sheet query multiplied processing
- `BUILTIN.CONSOLIDATE` called 12× more per account
- NetSuite query optimizer struggled with large multi-period requests
- Many cells returned `$0` (timeouts/failures)

---

## Questions for Senior Engineer Review

1. **SuiteQL Optimization:**
   - Can we avoid `BUILTIN.CONSOLIDATE` while maintaining multi-subsidiary support?
   - Should we pre-calculate exchange rates and apply manually?
   - Can we use a different NetSuite API (RESTlet, SuiteTalk) for better performance?

2. **Query Structure:**
   - Is our `CROSS JOIN period_dates` approach in Balance Sheet queries optimal?
   - Should we denormalize and query one period at a time?
   - Can we use `UNION ALL` to combine P&L and Balance Sheet queries?

3. **Batching Strategy:**
   - Is period-by-period batching the best we can do?
   - Should we implement adaptive chunk sizing based on response times?
   - Can we parallelize backend queries (e.g., multiple Flask workers)?

4. **Caching:**
   - Should we implement backend caching (Redis/Memcached)?
   - Can we cache `BUILTIN.CONSOLIDATE` results at the transaction level?
   - What's the right cache invalidation strategy?

5. **Alternative Approaches:**
   - Should we build a NetSuite Saved Search instead?
   - Can we use NetSuite's native Excel connector (CloudExtend) for comparison?
   - Should we pre-aggregate data in a custom NetSuite record?

---

## Current Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                        Excel (Client)                            │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │  Formulas: NS.GLABAL(4220, "Jan 2024", "Jan 2024")        │ │
│  │  (100 accounts × 12 months = 1,200 formulas)              │ │
│  └────────────────────────────────────────────────────────────┘ │
│                              ↓                                   │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │  functions.js (Batching Engine)                            │ │
│  │  • Queue requests (pendingRequests.balance)                │ │
│  │  • Group by filters+period                                 │ │
│  │  • Split into chunks of 50 accounts                        │ │
│  │  • Debounce with 150ms batch delay                         │ │
│  └────────────────────────────────────────────────────────────┘ │
│                              ↓                                   │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │  Client-Side Cache (Map objects)                           │ │
│  │  • cache.balance: Map<cacheKey, value>                     │ │
│  │  • cache.title: Map<account, name>                         │ │
│  │  • Cleared before each refresh                             │ │
│  └────────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────┘
                              ↓ HTTPS
┌─────────────────────────────────────────────────────────────────┐
│              Cloudflare Worker (Proxy)                           │
│  • CORS headers                                                  │
│  • Forwards to Cloudflare Tunnel                                │
└─────────────────────────────────────────────────────────────────┘
                              ↓ HTTPS
┌─────────────────────────────────────────────────────────────────┐
│              Cloudflare Tunnel (Local Proxy)                     │
│  • Exposes localhost:5002 securely                              │
└─────────────────────────────────────────────────────────────────┘
                              ↓ HTTP
┌─────────────────────────────────────────────────────────────────┐
│              Flask Backend (server.py)                           │
│  • /batch/balance endpoint                                       │
│  • OAuth1 authentication to NetSuite                            │
│  • Builds SuiteQL queries                                        │
│  • Separates P&L vs Balance Sheet logic                         │
│  • Returns JSON: { balances: { account: { period: value } } }   │
└─────────────────────────────────────────────────────────────────┘
                              ↓ HTTPS (OAuth1)
┌─────────────────────────────────────────────────────────────────┐
│              NetSuite (SuiteQL REST API)                         │
│  • TransactionAccountingLine table (millions of rows)           │
│  • BUILTIN.CONSOLIDATE function (slow!)                         │
│  • Query execution time: 2-3 seconds per account-period         │
└─────────────────────────────────────────────────────────────────┘
```

---

## Performance Metrics Summary

### Current State (v1.0.0.93):
- **First Daily Refresh:** 6-8 minutes for 100 accounts × 12 months
- **Subsequent Opens:** Instant (cached)
- **API Calls:** 24 (2 per month)
- **User Experience:** Acceptable, but could be better

### Ideal State:
- **First Daily Refresh:** 1-2 minutes (target)
- **Subsequent Opens:** Instant (cached) ✅ Already achieved
- **API Calls:** Minimize while maintaining accuracy
- **User Experience:** Fast enough for production use

---

## Files for Reference

### Frontend:
- `docs/functions.js` - Custom function logic, batching, caching
- `docs/functions.json` - Function metadata
- `docs/taskpane.html` - Task pane UI with refresh buttons

### Backend:
- `backend/server.py` - Flask API, SuiteQL queries
- `backend/netsuite_config.json` - NetSuite OAuth credentials (not in repo)

### Manifest:
- `excel-addin/manifest-claude.xml` - Excel add-in configuration

---

## Specific Code to Review

### Backend Query Builder (Python):
```python
def build_pl_query(accounts, periods, filters):
    """Build SuiteQL for P&L accounts (activity for period)"""
    account_list = ','.join([f"'{acc}'" for acc in accounts])
    period_list = ','.join([f"'{p}'" for p in periods])
    
    query = f"""
    WITH sub_info AS (
        SELECT 
            COALESCE({target_subsidiary}, parent) AS target_sub,
            (SELECT COUNT(*) FROM Subsidiary WHERE isinactive = 'F') AS sub_count
        FROM Subsidiary
        WHERE id = COALESCE({target_subsidiary}, parent)
        AND parent IS NULL
        LIMIT 1
    )
    SELECT 
        a.acctnumber,
        a.accttype,
        ap.periodname,
        SUM(
            CASE 
                WHEN sub_info.sub_count > 1 THEN 
                    TO_NUMBER(
                        BUILTIN.CONSOLIDATE(
                            tal.amount,
                            'LEDGER',
                            'DEFAULT',
                            'DEFAULT',
                            sub_info.target_sub,
                            t.postingperiod,
                            'DEFAULT'
                        )
                    )
                ELSE 
                    tal.amount
            END
            * 
            CASE 
                WHEN a.accttype IN ('Income', 'OthIncome') THEN -1 
                ELSE 1 
            END
        ) AS total_amount
    FROM TransactionAccountingLine tal
    CROSS JOIN sub_info
    INNER JOIN Transaction t ON t.id = tal.transaction
    INNER JOIN Account a ON a.id = tal.account
    INNER JOIN AccountingPeriod ap ON ap.id = t.postingperiod
    WHERE t.posting = 'T'
      AND tal.posting = 'T'
      AND tal.accountingbook = 1
      AND ap.periodname IN ({period_list})
      AND a.acctnumber IN ({account_list})
      AND a.accttype IN ('Income', 'COGS', 'Cost of Goods Sold', 'Expense', 'OthIncome', 'OthExpense')
      AND COALESCE(a.eliminate, 'F') = 'F'
      {filter_conditions}
    GROUP BY a.acctnumber, a.accttype, ap.periodname
    ORDER BY a.acctnumber, ap.periodname
    """
    return query

def build_bs_query(accounts, periods, filters):
    """Build SuiteQL for Balance Sheet accounts (cumulative to period end)"""
    account_list = ','.join([f"'{acc}'" for acc in accounts])
    period_list = ','.join([f"'{p}'" for p in periods])
    
    query = f"""
    WITH sub_info AS (
        -- Same as P&L
    ),
    period_dates AS (
        SELECT id, periodname, enddate
        FROM AccountingPeriod
        WHERE periodname IN ({period_list})
          AND isquarter = 'F'
          AND isyear = 'F'
    )
    SELECT 
        a.acctnumber,
        a.accttype,
        pd.periodname,
        SUM(
            CASE 
                WHEN sub_info.sub_count > 1 THEN 
                    TO_NUMBER(
                        BUILTIN.CONSOLIDATE(
                            tal.amount,
                            'LEDGER',
                            'DEFAULT',
                            'DEFAULT',
                            sub_info.target_sub,
                            pd.id,
                            'DEFAULT'
                        )
                    )
                ELSE 
                    tal.amount
            END
        ) AS total_amount
    FROM TransactionAccountingLine tal
    CROSS JOIN sub_info
    CROSS JOIN period_dates pd
    INNER JOIN Transaction t ON t.id = tal.transaction
    INNER JOIN Account a ON a.id = tal.account
    WHERE t.posting = 'T'
      AND tal.posting = 'T'
      AND tal.accountingbook = 1
      AND t.trandate <= pd.enddate
      AND a.acctnumber IN ({account_list})
      AND a.accttype NOT IN ('Income', 'COGS', 'Cost of Goods Sold', 'Expense', 'OthIncome', 'OthExpense')
      {filter_conditions}
    GROUP BY a.acctnumber, a.accttype, pd.periodname
    ORDER BY a.acctnumber, pd.periodname
    """
    return query
```

---

## Request for Feedback

**Primary Question:** How can we reduce the 6-8 minute refresh time to 1-2 minutes while maintaining:
1. Multi-subsidiary support with proper currency consolidation
2. Accurate P&L (period activity) vs Balance Sheet (cumulative) logic
3. Real-time data (no stale cached data from NetSuite)

**Specific Areas of Concern:**
1. Is `BUILTIN.CONSOLIDATE` the bottleneck, or is it the query structure?
2. Should we abandon SuiteQL and use a different NetSuite API?
3. Can we optimize the queries to reduce `BUILTIN.CONSOLIDATE` calls?
4. Is backend caching a viable solution (with what invalidation strategy)?
5. Are there NetSuite-specific optimizations we're missing?

Thank you for your review!

