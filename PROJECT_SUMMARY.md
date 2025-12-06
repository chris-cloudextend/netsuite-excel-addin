# NetSuite Excel Add-in - Technical Summary

## Project Overview

This is an Excel Add-in that provides custom functions to retrieve financial data from NetSuite via SuiteQL queries. Users can build financial reports in Excel using formulas like `=XAVI.BALANCE("4010", "Jan 2025", "Jan 2025")` to fetch account balances, and the add-in handles consolidation, caching, and performance optimization.

### Key Features
- **Custom Excel Functions**: `XAVI.BALANCE` (balance), `XAVI.NAME` (account name), `XAVI.TYPE` (account type)
- **Multi-currency Consolidation**: Uses NetSuite's `BUILTIN.CONSOLIDATE` function
- **Smart Caching**: In-memory and localStorage caching with intelligent invalidation
- **Build Mode Detection**: Detects rapid formula entry (drag-drop) and batches requests
- **Optimized Queries**: Different strategies for P&L vs Balance Sheet accounts

---

## Architecture

```
┌─────────────────────┐     ┌─────────────────────┐     ┌─────────────────┐
│   Excel Add-in      │────▶│   Flask Backend     │────▶│    NetSuite     │
│   (functions.js)    │     │   (server.py)       │     │    SuiteQL      │
│                     │◀────│                     │◀────│                 │
│   - Custom funcs    │     │   - Query building  │     │   - Account     │
│   - Caching         │     │   - Consolidation   │     │   - Transaction │
│   - Build mode      │     │   - Caching         │     │   - Period      │
└─────────────────────┘     └─────────────────────┘     └─────────────────┘
         │
         ▼
┌─────────────────────┐
│   Taskpane UI       │
│   (taskpane.html)   │
│                     │
│   - Prep Data       │
│   - Refresh All     │
│   - Refresh Selected│
└─────────────────────┘
```

---

## Balance Sheet vs Income Statement Approaches

### The Core Challenge

**P&L (Income Statement) accounts** show activity for a specific period:
- Revenue for January 2025 = Sum of all revenue transactions in January 2025

**Balance Sheet accounts** show cumulative balances:
- Cash as of January 2025 = Sum of all cash transactions from inception through January 2025

Additionally, for **multi-currency** environments, Balance Sheet accounts must be revalued at the reporting period's exchange rate, not the transaction's historical rate.

---

### P&L Query Approach (Optimized Pivoted Query)

**Strategy**: Single query fetches ALL P&L accounts for an entire fiscal year in ~6-15 seconds.

**Key Optimization**: Pivot all 12 months into columns using `SUM(CASE WHEN...)`:

```sql
SELECT
  a.acctnumber AS account_number,
  a.accttype AS account_type,
  SUM(CASE WHEN TO_CHAR(ap.startdate,'YYYY-MM')='2025-01' THEN cons_amt ELSE 0 END) AS jan,
  SUM(CASE WHEN TO_CHAR(ap.startdate,'YYYY-MM')='2025-02' THEN cons_amt ELSE 0 END) AS feb,
  SUM(CASE WHEN TO_CHAR(ap.startdate,'YYYY-MM')='2025-03' THEN cons_amt ELSE 0 END) AS mar,
  -- ... all 12 months
FROM (
  SELECT
    tal.account,
    t.postingperiod,
    CASE
      WHEN subs_count > 1 THEN
        TO_NUMBER(
          BUILTIN.CONSOLIDATE(
            tal.amount,
            'LEDGER',
            'DEFAULT',
            'DEFAULT',
            1,  -- target subsidiary
            t.postingperiod,
            'DEFAULT'
          )
        )
      ELSE tal.amount
    END
    * CASE WHEN a.accttype IN ('Income','OthIncome') THEN -1 ELSE 1 END AS cons_amt
  FROM transactionaccountingline tal
    JOIN transaction t ON t.id = tal.transaction
    JOIN account a ON a.id = tal.account
    JOIN accountingperiod apf ON apf.id = t.postingperiod
    CROSS JOIN (
      SELECT COUNT(*) AS subs_count
      FROM subsidiary
      WHERE isinactive = 'F'
    ) subs_cte
  WHERE t.posting = 'T'
    AND tal.posting = 'T'
    AND tal.accountingbook = 1
    AND apf.isyear = 'F' 
    AND apf.isquarter = 'F'
    AND TO_CHAR(apf.startdate,'YYYY') = '2025'
    AND a.accttype IN ('Income', 'COGS', 'Cost of Goods Sold', 'Expense', 'OthIncome', 'OthExpense')
) x
JOIN accountingperiod ap ON ap.id = x.postingperiod
JOIN account a ON a.id = x.account
GROUP BY a.acctnumber, a.accttype
ORDER BY a.acctnumber
```

**Why This Works**:
- Returns ~100-300 rows (one per account) instead of ~1000+ rows (one per account/month)
- Single query, no pagination needed
- CROSS JOIN subquery instead of CTE (better SuiteQL compatibility)

---

### Balance Sheet Query Approach (Multi-Period with Fixed Consolidation)

**Strategy**: Single query fetches ALL Balance Sheet accounts for requested periods in ~60-70 seconds.

**Critical Insight**: For multi-currency consolidation, `BUILTIN.CONSOLIDATE` must use the **target period's ID** (not the transaction's posting period) to get correct exchange rates.

```sql
SELECT 
  a.acctnumber AS account_number,
  a.accttype AS account_type,
  -- January 2025 cumulative balance (uses Jan 2025 exchange rate)
  SUM(CASE WHEN ap.startdate <= p_2025_01.enddate
    THEN TO_NUMBER(BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT', 1, p_2025_01.id, 'DEFAULT'))
         * CASE WHEN a.accttype IN ('AcctPay', 'CredCard', 'OthCurrLiab', 'LongTermLiab', 'DeferRevenue', 'Equity', 'RetainedEarnings') 
                THEN -1 ELSE 1 END
    ELSE 0 END) AS bal_2025_01,
  -- February 2025 cumulative balance (uses Feb 2025 exchange rate)
  SUM(CASE WHEN ap.startdate <= p_2025_02.enddate
    THEN TO_NUMBER(BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT', 1, p_2025_02.id, 'DEFAULT'))
         * CASE WHEN a.accttype IN ('AcctPay', 'CredCard', 'OthCurrLiab', 'LongTermLiab', 'DeferRevenue', 'Equity', 'RetainedEarnings') 
                THEN -1 ELSE 1 END
    ELSE 0 END) AS bal_2025_02,
  -- December 2024 cumulative balance (uses Dec 2024 exchange rate)
  SUM(CASE WHEN ap.startdate <= p_2024_12.enddate
    THEN TO_NUMBER(BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT', 1, p_2024_12.id, 'DEFAULT'))
         * CASE WHEN a.accttype IN ('AcctPay', 'CredCard', 'OthCurrLiab', 'LongTermLiab', 'DeferRevenue', 'Equity', 'RetainedEarnings') 
                THEN -1 ELSE 1 END
    ELSE 0 END) AS bal_2024_12

FROM transactionaccountingline tal
  INNER JOIN transaction t ON t.id = tal.transaction
  INNER JOIN account a ON a.id = tal.account
  INNER JOIN accountingperiod ap ON ap.id = t.postingperiod
  -- INNER JOIN for each target period (gets period ID for CONSOLIDATE)
  INNER JOIN accountingperiod p_2025_01 
    ON TO_CHAR(p_2025_01.startdate, 'YYYY-MM') = '2025-01' 
    AND p_2025_01.isquarter = 'F' AND p_2025_01.isyear = 'F'
  INNER JOIN accountingperiod p_2025_02 
    ON TO_CHAR(p_2025_02.startdate, 'YYYY-MM') = '2025-02' 
    AND p_2025_02.isquarter = 'F' AND p_2025_02.isyear = 'F'
  INNER JOIN accountingperiod p_2024_12 
    ON TO_CHAR(p_2024_12.startdate, 'YYYY-MM') = '2024-12' 
    AND p_2024_12.isquarter = 'F' AND p_2024_12.isyear = 'F'

WHERE 
  t.posting = 'T'
  AND tal.posting = 'T'
  AND tal.accountingbook = 1
  AND a.accttype NOT IN ('Income', 'COGS', 'Cost of Goods Sold', 'Expense', 'OthIncome', 'OthExpense')
  AND ap.startdate <= p_2025_02.enddate  -- Use chronologically latest period
  AND ap.isyear = 'F'
  AND ap.isquarter = 'F'

GROUP BY a.acctnumber, a.accttype
ORDER BY a.acctnumber
```

**Key Points**:
1. **Fixed Period ID for CONSOLIDATE**: Each `SUM(CASE WHEN...)` uses its own period's ID (e.g., `p_2025_01.id`)
2. **Sign Convention**: Liabilities and Equity are multiplied by -1 to show positive on reports
3. **Cumulative Filter**: `ap.startdate <= p_YYYY_MM.enddate` includes all transactions from inception
4. **WHERE Clause**: Uses the **chronologically latest** period (not last in list)

---

## Account Type Sign Conventions

```
BALANCE SHEET ACCOUNTS:
-----------------------
ASSETS (Debit balance - stored positive, NO sign flip):
  - Bank              Bank/Cash accounts
  - AcctRec           Accounts Receivable
  - OthCurrAsset      Other Current Asset
  - FixedAsset        Fixed Asset
  - OthAsset          Other Asset
  - DeferExpense      Deferred Expense
  - UnbilledRec       Unbilled Receivable

LIABILITIES (Credit balance - stored negative, FLIP × -1):
  - AcctPay           Accounts Payable
  - CredCard          Credit Card
  - OthCurrLiab       Other Current Liability
  - LongTermLiab      Long Term Liability
  - DeferRevenue      Deferred Revenue

EQUITY (Credit balance - stored negative, FLIP × -1):
  - Equity            Equity accounts
  - RetainedEarnings  Retained Earnings

PROFIT & LOSS ACCOUNTS:
-----------------------
INCOME (Credit balance - stored negative, FLIP × -1):
  - Income            Revenue/Sales
  - OthIncome         Other Income

EXPENSES (Debit balance - stored positive, NO sign flip):
  - COGS              Cost of Goods Sold
  - Expense           Operating Expense
  - OthExpense        Other Expense
```

---

## Frontend Behavior Modes

### 1. Single Cell Resolution

When a user enters a single formula like `=XAVI.BALANCE("4010", "Jan 2025", "Jan 2025")`:

```javascript
// In BALANCE function
async function BALANCE(account, fromPeriod, toPeriod, subsidiary, department, location, classId) {
    // 1. Check in-memory cache
    const cacheKey = getCacheKey('balance', { account, fromPeriod, toPeriod, ... });
    if (cache.balance.has(cacheKey)) {
        return cache.balance.get(cacheKey);
    }
    
    // 2. Check localStorage cache
    const lsValue = checkLocalStorageCache(account, fromPeriod, toPeriod);
    if (lsValue !== null) {
        cache.balance.set(cacheKey, lsValue);
        return lsValue;
    }
    
    // 3. Queue for batch processing
    return new Promise((resolve, reject) => {
        pendingQueue.push({ params: {...}, resolve, reject, cacheKey });
        startBatchTimer(); // 500ms delay to collect more requests
    });
}
```

### 2. Drag-Drop / Build Mode

When user drags a formula down multiple rows, Excel rapidly fires many formula calls:

```javascript
// Build mode detection
let buildModeFormulas = 0;
let buildModeStartTime = null;

function detectBuildMode() {
    const now = Date.now();
    if (!buildModeStartTime || (now - buildModeStartTime) > BUILD_MODE_WINDOW_MS) {
        buildModeStartTime = now;
        buildModeFormulas = 0;
    }
    buildModeFormulas++;
    
    if (buildModeFormulas >= BUILD_MODE_THRESHOLD) {
        enterBuildMode();
    }
}

async function runBuildModeBatch() {
    // 1. Group formulas by filter combination (subsidiary, dept, loc, class)
    const filterGroups = new Map();
    for (const item of buildModePending) {
        const filterKey = getFilterKey(item.params);
        if (!filterGroups.has(filterKey)) {
            filterGroups.set(filterKey, []);
        }
        filterGroups.get(filterKey).push(item);
    }
    
    // 2. For each filter group:
    for (const [filterKey, groupItems] of filterGroups) {
        const filters = parseFilterKey(filterKey);
        
        // 3. Detect account types
        const accountTypes = await batchGetAccountTypes(accountsArray);
        
        // 4. Split into P&L and BS accounts
        const plAccounts = [], bsAccounts = [];
        for (const acct of accountsArray) {
            if (isBalanceSheetType(accountTypes[acct])) {
                bsAccounts.push(acct);
            } else {
                plAccounts.push(acct);
            }
        }
        
        // 5. Fetch BS accounts using efficient multi-period query
        if (bsAccounts.length > 0) {
            const response = await fetch(`${SERVER_URL}/batch/bs_periods`, {
                method: 'POST',
                body: JSON.stringify({
                    periods: periodsArray,
                    subsidiary: filters.subsidiary,
                    ...
                })
            });
            // Cache results...
        }
        
        // 6. Fetch P&L accounts using full_year_refresh
        if (plAccounts.length >= 5) {
            const response = await fetch(`${SERVER_URL}/batch/full_year_refresh`, {
                method: 'POST',
                body: JSON.stringify({
                    year: year,
                    skip_bs: true,  // Already handled above
                    ...
                })
            });
            // Cache results...
        }
        
        // 7. Resolve all pending promises
        for (const item of groupItems) {
            const value = allBalances[item.params.account][item.params.fromPeriod];
            item.resolve(value);
        }
    }
}
```

### 3. Refresh All (Taskpane)

Scans the active sheet for all `XAVI.*` formulas, detects accounts and periods, then fetches fresh data:

```javascript
async function refreshCurrentSheet() {
    // 1. Scan all formulas on sheet
    const formulas = await getFormulasByType('NS.GLA');
    
    // 2. Extract unique accounts and periods
    const accounts = new Set();
    const periods = new Set();
    const years = new Set();
    for (const formula of formulas) {
        accounts.add(parseAccount(formula));
        periods.add(parsePeriod(formula));
        years.add(extractYear(parsePeriod(formula)));
    }
    
    // 3. Get account types
    const accountTypes = await fetch('/batch/account_types', {
        method: 'POST',
        body: JSON.stringify({ accounts: Array.from(accounts) })
    });
    
    // 4. Fetch P&L for each year
    for (const year of years) {
        await fetch('/batch/full_year_refresh', {
            method: 'POST',
            body: JSON.stringify({ year, skip_bs: true })
        });
    }
    
    // 5. Fetch BS for requested periods
    if (bsAccounts.length > 0) {
        await fetch('/batch/bs_periods', {
            method: 'POST',
            body: JSON.stringify({ periods: Array.from(periods) })
        });
    }
    
    // 6. Re-evaluate formulas
    await recalculateSheet();
}
```

### 4. Refresh Selected

Only refreshes the selected cells:

```javascript
async function refreshSelected() {
    // 1. Get selected range
    const range = context.workbook.getSelectedRange();
    
    // 2. Parse formulas to extract account/period pairs
    const items = [];
    for (const cell of range) {
        const formula = cell.formula;
        if (formula.includes('XAVI.BALANCE')) {
            const account = parseAccountFromFormula(formula);
            const period = parsePeriodFromFormula(formula);
            items.push({ account, period });
        }
    }
    
    // 3. Clear cache for these specific items
    await Excel.run(async (context) => {
        const tempRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
        const itemsList = items.map(i => `${i.account}:${i.period}`).join(',');
        tempRange.formulas = [[`=XAVI.BALANCE("__CLEARCACHE__","${itemsList}","")`]];
        await context.sync();
    });
    
    // 4. Trigger recalculation
    range.calculate();
}
```

---

## Backend Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/batch/full_year_refresh` | POST | Fetch all P&L accounts for a fiscal year |
| `/batch/bs_periods` | POST | Fetch all BS accounts for specific periods |
| `/batch/balance` | POST | Fetch specific accounts for specific periods |
| `/batch/account_types` | POST | Get account types for a list of accounts |
| `/account/<number>/name` | GET | Get account name |
| `/account/<number>/type` | GET | Get account type |

---

## Caching Strategy

### Frontend Caches
1. **In-memory cache** (`cache.balance`): Fastest, lost on page reload
2. **localStorage cache**: Persists across sessions, 5-minute TTL
3. **fullYearCache**: Stores complete year data for instant lookup

### Backend Cache
- `balance_cache`: Server-side cache for fast repeat queries
- `account_title_cache`: Account names preloaded at startup

### Cache Key Format
```javascript
const cacheKey = JSON.stringify({
    type: 'balance',
    account: '4010',
    fromPeriod: 'Jan 2025',
    toPeriod: 'Jan 2025',
    subsidiary: '1',
    department: '',
    location: '',
    classId: ''
});
```

---

## Key Learnings & Gotchas

1. **SuiteQL doesn't support CTEs well**: Use CROSS JOIN subqueries instead
2. **BUILTIN.CONSOLIDATE needs fixed period ID**: For BS accounts, use target period's ID, not transaction's period
3. **Sign conventions matter**: Liabilities/Equity stored as negative, must flip for display
4. **Account type spelling**: NetSuite uses `CredCard` not `CreditCard`
5. **WHERE clause ordering**: Must use chronologically latest period, not last in input list
6. **Rate limiting**: NetSuite has concurrent request limits, batch operations help
7. **$0 accounts not returned**: Must explicitly cache zeros for accounts not in query results

---

## File Structure

```
├── backend/
│   └── server.py              # Flask API server
├── docs/
│   ├── functions.js           # Custom Excel functions + caching
│   ├── functions.json         # Function metadata for Excel
│   └── taskpane.html          # Taskpane UI + JavaScript
├── excel-addin/
│   └── manifest-claude.xml    # Excel add-in manifest
└── CLOUDFLARE-WORKER-CODE.js  # Proxy for production deployment
```

---

## Security Architecture

### Current State & Recommendations

#### 1. Credential Storage ✅ Partially Addressed

**Current Implementation:**
```
backend/netsuite_config.json (NOT in git - see .gitignore)
{
  "account_id": "589861",
  "consumer_key": "...",
  "consumer_secret": "...",
  "token_id": "...",
  "token_secret": "..."
}
```

**Status:** 
- ✅ File is in `.gitignore` (lines 28, 37)
- ✅ Pattern `**/netsuite_config*.json` blocks all variants
- ⚠️ Plain text file on disk (not encrypted)
- ⚠️ No secret rotation mechanism

**Recommendations for Production:**
1. **Environment Variables**: Move to `os.environ.get('NETSUITE_CONSUMER_KEY')` etc.
2. **AWS Secrets Manager / HashiCorp Vault**: For enterprise deployments
3. **Encrypted config**: Use `python-dotenv` with encrypted `.env` files
4. **Example:** Add `netsuite_config.example.json` with placeholder values

#### 2. Transport Security ✅ HTTPS Throughout

**Current Implementation:**
```javascript
// docs/functions.js line 12
const SERVER_URL = 'https://netsuite-proxy.chris-corcoran.workers.dev';
```

**Full Chain:**
```
Excel Add-in (HTTPS) → Cloudflare Worker (HTTPS) → Cloudflare Tunnel (HTTPS) → Flask (localhost:5002)
     ↑                        ↑                           ↑                          ↑
   TLS 1.3              Edge TLS               Encrypted Tunnel              Local only
```

**Status:**
- ✅ All external traffic over HTTPS
- ✅ Cloudflare provides TLS termination
- ✅ Tunnel encrypts local traffic
- ✅ Flask only binds to localhost (not exposed)

#### 3. Cloudflare Worker Role ✅ Secure Design

**What it does:** CORS proxy only (no credential access)

```javascript
// CLOUDFLARE-WORKER-CODE.js
export default {
  async fetch(request) {
    const TUNNEL_URL = 'https://xxx.trycloudflare.com';
    // Simply forwards requests - no secrets stored here
    const response = await fetch(TUNNEL_URL + url.pathname, { ... });
    // Adds CORS headers and returns
  }
}
```

**Status:**
- ✅ No credentials in worker code
- ✅ Worker only sees encrypted traffic (can't decrypt TLS from tunnel)
- ✅ All authentication happens in Flask backend (via OAuth1)

**Why needed:** Excel add-ins require CORS headers. NetSuite doesn't provide them.

#### 4. Client-Side Exposure ⚠️ Risk on Shared Machines

**Current Implementation:**
```javascript
// localStorage caching in functions.js
localStorage.setItem('netsuite_balance_cache', JSON.stringify(balanceData));
localStorage.setItem('netsuite_account_titles', JSON.stringify(titles));
```

**What's Cached:**
- Account balances (financial amounts)
- Account names
- Account types
- Subsidiary/department/location/class IDs

**Risks:**
- 🔴 Data persists after logout/session end
- 🔴 Other browser tabs/extensions could potentially read
- 🔴 On shared machines, next user could see data

**Mitigations Already in Place:**
- ✅ 5-minute TTL (data expires quickly)
- ✅ No PII (just account numbers and amounts)
- ✅ No authentication tokens in localStorage

**Recommendations:**
1. **Use sessionStorage instead**: Cleared when browser tab closes
2. **Encrypt cached data**: Use Web Crypto API with session key
3. **Clear on deactivate**: `Office.addin.onVisibilityModeChanged` to clear cache
4. **Add clear cache button**: Already implemented in taskpane ✅

**Example Enhancement:**
```javascript
// Replace localStorage with sessionStorage for sensitive data
const CACHE_STORAGE = sessionStorage; // Instead of localStorage

// Or encrypt before storing
async function encryptAndStore(key, data) {
    const cryptoKey = await getSessionKey();
    const encrypted = await crypto.subtle.encrypt(
        { name: 'AES-GCM', iv: window.crypto.getRandomValues(new Uint8Array(12)) },
        cryptoKey, 
        new TextEncoder().encode(JSON.stringify(data))
    );
    localStorage.setItem(key, btoa(String.fromCharCode(...new Uint8Array(encrypted))));
}
```

### Security Checklist

| Area | Status | Notes |
|------|--------|-------|
| Credentials in git | ✅ | `.gitignore` blocks all config files |
| HTTPS for all traffic | ✅ | Cloudflare provides TLS |
| OAuth1 authentication | ✅ | HMAC-SHA256 signature |
| CORS properly configured | ✅ | Worker adds headers, not wildcard on sensitive endpoints |
| Client cache exposure | ⚠️ | Consider sessionStorage or encryption |
| Secret rotation | ❌ | Manual process, consider automating |
| Audit logging | ❌ | No request logging (add for compliance) |
| Rate limiting | ⚠️ | NetSuite enforces limits, no client-side throttling |

---

## Detailed Technical Review - Q&A

### Finance Questions

#### Q1: Why both 'COGS' AND 'Cost of Goods Sold'?

**Answer**: NetSuite historically used different values. Testing against live data shows:
- Modern NetSuite accounts use `COGS`
- Legacy imported accounts may have `Cost of Goods Sold`
- Some NetSuite documentation references both

**Verification Query**:
```sql
SELECT DISTINCT accttype FROM Account WHERE accttype LIKE '%COGS%' OR accttype LIKE '%Cost%'
```

**Resolution**: Both are included defensively. This is documented as a known NetSuite quirk.

---

#### Q2: NonPosting/Stat accounts - intentionally excluded?

**Answer**: Yes, intentionally excluded.

- `NonPosting` accounts cannot have transaction amounts (used for grouping/headers)
- `Stat` accounts are statistical KPIs (unit counts, not currency amounts)

Neither would return meaningful balance data from `transactionaccountingline`.

**Documentation added** to account types reference:
```
OTHER ACCOUNT TYPES (Excluded from queries):
  - NonPosting        Statistical/Non-posting (no transactions)
  - Stat              Statistical accounts (non-financial KPIs)
```

---

#### Q3: Zero Balance vs. No Data vs. Account Doesn't Exist

**Current Behavior**:

| Scenario | What Happens | User Sees |
|----------|--------------|-----------|
| Account exists, $0 balance | Query returns no row (NetSuite optimization) | `0` (cached explicitly) |
| Account in query but filtered out | Not returned | `0` (potentially incorrect!) |
| Account doesn't exist | Query returns nothing | `0` (should error) |
| Network/API error | Catch block | Error message or `""` |

**The Risk**: Silently returning $0 for non-existent accounts is indeed a material misstatement risk.

**Recommended Fix**:
```javascript
// Instead of returning 0 for missing accounts:
if (!accountExists(account)) {
    return "#INVALID_ACCT";  // Excel error value
}
```

**Implementation needed**: Pre-validate account numbers against a cached account list.

---

### Engineering Questions

#### Q4: SQL Injection Risk

**Current Mitigation**: All inputs go through `escape_sql()`:
```python
def escape_sql(text):
    """Escape single quotes in SQL strings"""
    if text is None:
        return ""
    return str(text).replace("'", "''")

# Usage:
accounts_in = ','.join([f"'{escape_sql(acc)}'" for acc in accounts])
```

**Example attack**: `1; DROP TABLE account--` becomes `'1; DROP TABLE account--'` (treated as literal string).

**Why SuiteQL is safer**:
- Read-only API (no DDL/DML)
- NetSuite validates queries server-side
- No direct database access

**Recommendation**: Add input validation regex:
```python
def validate_account_number(acct):
    """Only allow alphanumeric and hyphens"""
    if not re.match(r'^[a-zA-Z0-9\-]+$', str(acct)):
        raise ValueError(f"Invalid account number: {acct}")
```

---

#### Q5: Error Handling Strategy

**Current Implementation** (172 error handling blocks in server.py):

| Error Type | Backend Response | Frontend Display |
|------------|------------------|------------------|
| NetSuite 429 (rate limit) | 429 + retry-after | "Rate limited, retrying..." |
| NetSuite timeout | 504 Gateway Timeout | "Query timeout, try smaller selection" |
| Auth failure | 401 Unauthorized | "Authentication error" |
| Network error | 502 Bad Gateway | "Server unavailable" |
| Invalid account | 200 + empty result | `0` (see Q3 issue above) |

**Frontend retry logic**:
```javascript
const MAX_RETRIES = 3;
for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    try {
        const response = await fetch(...);
        if (response.status === 429) {
            await sleep(2000 * (attempt + 1));  // Exponential backoff
            continue;
        }
        // ... handle other statuses
    } catch (e) {
        if (attempt === MAX_RETRIES) throw e;
    }
}
```

---

#### Q6: 60-70 Second Query Time UX

**Current mitigations**:
1. Status bar shows "Fetching Balance Sheet accounts..." with spinner
2. `localStorage` broadcasts progress updates
3. Build mode batches requests (users see "busy" indicator in cells)

**Recommendations not yet implemented**:
- [ ] Progressive loading (show P&L results while BS fetches)
- [ ] Stale-while-revalidate (show cached, refresh in background)
- [ ] Query splitting (fetch 20 accounts at a time)

**Why it's slow**: Balance Sheet queries use `BUILTIN.CONSOLIDATE` which revalues every transaction at period-end exchange rate. This is CPU-intensive in NetSuite.

---

#### Q7: Caching Race Conditions

**Identified Issue**: Two identical requests could both miss cache and both query.

**Current mitigation** (partial):
```javascript
// Build mode queues all requests, processes once
if (buildModeActive) {
    buildModePending.push(request);  // Deduplicated by key
    return;  // Don't process individually
}
```

**Recommended enhancement**:
```javascript
const inflightRequests = new Map();  // Track pending fetches

async function fetchWithDedup(key, fetchFn) {
    if (inflightRequests.has(key)) {
        return inflightRequests.get(key);  // Return same promise
    }
    const promise = fetchFn();
    inflightRequests.set(key, promise);
    try {
        return await promise;
    } finally {
        inflightRequests.delete(key);
    }
}
```

---

#### Q8: Cache Invalidation

| Action | Cache Behavior |
|--------|----------------|
| Formula evaluation | Check cache first (5-min TTL) |
| "Refresh Selected" | Clears specific entries, then re-fetches |
| "Refresh All" | Clears all entries, full re-fetch |
| "Clear Cache" button | Clears all caches (in-memory + localStorage) |
| User posts JE in NetSuite | Must manually refresh in Excel |

**Real-time sync not implemented** - would require NetSuite webhooks or polling.

---

#### Q9: Backend Scalability

**Current state**: In-memory dict caches
```python
balance_cache = {}  # Lost on restart, unbounded growth
account_title_cache = {}
```

**Recommendations for production**:
1. Redis for distributed caching
2. Cache size limits (LRU eviction)
3. Cache warming on startup

---

#### Q10: API Design - GET vs POST

**Issue**: `GET /account/<number>/name` exposes account numbers in URL logs.

**Current endpoints**:
| Method | Endpoint | Risk |
|--------|----------|------|
| GET | `/account/<num>/name` | Account in URL |
| GET | `/account/<num>/type` | Account in URL |
| POST | `/batch/balance` | Secure (body) |

**Recommendation**: Migrate to POST for all account lookups.

---

#### Q11: Magic Strings → Constants

**Recommendation** (not yet implemented):
```python
# backend/constants.py
class AccountTypes:
    # Assets
    BANK = 'Bank'
    ACCT_REC = 'AcctRec'
    # ... etc
    
    BALANCE_SHEET = {BANK, ACCT_REC, ...}
    PL_TYPES = {'Income', 'OthIncome', 'COGS', 'Expense', 'OthExpense'}
```

---

#### Q12: Function Naming (Historical Note)

**History**: Original function was `NS.GLABAL` (GL Account BALance).

**Resolution**: Renamed to `XAVI.BALANCE` in v1.5.0.0 with cleaner naming:
- `XAVI.BALANCE` - Account balance
- `XAVI.BUDGET` - Budget amount
- `XAVI.NAME` - Account name
- `XAVI.TYPE` - Account type
- `XAVI.PARENT` - Parent account

---

#### Q13: Hardcoded Subsidiary

**Clarification**: Subsidiary IS passed dynamically:
```python
def build_bs_multi_period_query(..., target_sub, ...):
    # target_sub comes from request, defaults to parent subsidiary
    select_columns.append(f"""
        BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT', 
            {target_sub},  -- Dynamic!
            {alias}.id, 'DEFAULT')
    """)
```

The `1` in comments is example value, not hardcoded.

---

#### Q14-16: Testing, Monitoring, Documentation

**Current State**:
| Area | Status | Notes |
|------|--------|-------|
| Unit tests | ❌ | Not implemented |
| Integration tests | ❌ | Manual QA only |
| API docs | ⚠️ | Inline comments only |
| Logging | ⚠️ | Print statements, no structured logging |
| Metrics | ❌ | No Prometheus/DataDog |
| Alerting | ❌ | None |

**Recommendations**:
1. Add pytest for sign convention and query building
2. Use Flask-RESTx for auto-generated Swagger docs
3. Add `structlog` for JSON logging
4. Implement health check endpoint

---

## Action Items Summary

| Priority | Issue | Status |
|----------|-------|--------|
| 🔴 Critical | SQL injection | ✅ `escape_sql()` implemented |
| 🔴 Critical | Zero vs error ambiguity | ⚠️ Needs account validation |
| 🔴 Critical | 60-70s UX | ⚠️ Status bar added, progressive loading pending |
| 🟠 High | Race conditions | ⚠️ Build mode helps, dedup recommended |
| 🟠 High | Error messages | ✅ Implemented (could improve) |
| 🟡 Medium | Magic strings | ❌ Refactor needed |
| 🟡 Medium | GET → POST | ❌ API change needed |
| 🟡 Medium | Testing | ❌ Not implemented |
| 🟢 Low | Function naming | ✅ Renamed to XAVI.* |
| 🟢 Low | Monitoring | ❌ Production concern |

---

## Version History (Recent)

- **1.4.89.0**: Complete account types documentation
- **1.4.88.0**: Fixed CredCard type mismatch
- **1.4.87.0**: Balance Sheet sign convention fix
- **1.4.86.0**: Fixed BS query WHERE clause (chronological order)
- **1.4.85.0**: Optimized P&L query (pivoted columns)
- **1.4.84.0**: Cache $0 values for BS accounts
- **1.4.83.0**: Group formulas by filters for correct caching

