# NetSuite Excel Add-in - Technical Summary

## Project Overview

This is an Excel Add-in that provides custom functions to retrieve financial data from NetSuite via SuiteQL queries. Users can build financial reports in Excel using formulas like `=NS.GLABAL("4010", "Jan 2025", "Jan 2025")` to fetch account balances, and the add-in handles consolidation, caching, and performance optimization.

### Key Features
- **Custom Excel Functions**: `NS.GLABAL` (balance), `NS.GLATITLE` (account name), `NS.GLACCTTYPE` (account type)
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

When a user enters a single formula like `=NS.GLABAL("4010", "Jan 2025", "Jan 2025")`:

```javascript
// In GLABAL function
async function GLABAL(account, fromPeriod, toPeriod, subsidiary, department, location, classId) {
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

Scans the active sheet for all `NS.GLA*` formulas, detects accounts and periods, then fetches fresh data:

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
        if (formula.includes('NS.GLABAL')) {
            const account = parseAccountFromFormula(formula);
            const period = parsePeriodFromFormula(formula);
            items.push({ account, period });
        }
    }
    
    // 3. Clear cache for these specific items
    await Excel.run(async (context) => {
        const tempRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
        const itemsList = items.map(i => `${i.account}:${i.period}`).join(',');
        tempRange.formulas = [[`=NS.GLABAL("__CLEARCACHE__","${itemsList}","")`]];
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

## Version History (Recent)

- **1.4.89.0**: Complete account types documentation
- **1.4.88.0**: Fixed CredCard type mismatch
- **1.4.87.0**: Balance Sheet sign convention fix
- **1.4.86.0**: Fixed BS query WHERE clause (chronological order)
- **1.4.85.0**: Optimized P&L query (pivoted columns)
- **1.4.84.0**: Cache $0 values for BS accounts
- **1.4.83.0**: Group formulas by filters for correct caching

