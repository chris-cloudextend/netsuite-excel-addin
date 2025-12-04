# NetSuite Excel Add-in - SuiteQL Queries Summary

## Purpose
This document summarizes all SuiteQL queries used in the NetSuite Excel Add-in backend (`server.py`). These queries power the custom Excel functions (NS.GLABAL, NS.GLATITLE, NS.GLACCTTYPE, etc.) and the Guide Me wizard.

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────┐
│                         EXCEL ADD-IN                                │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐                 │
│  │ NS.GLABAL   │  │ NS.GLATITLE │  │NS.GLACCTTYPE│                 │
│  │ (balances)  │  │ (names)     │  │ (types)     │                 │
│  └──────┬──────┘  └──────┬──────┘  └──────┬──────┘                 │
└─────────┼────────────────┼────────────────┼────────────────────────┘
          │                │                │
          ▼                ▼                ▼
┌─────────────────────────────────────────────────────────────────────┐
│                      PYTHON BACKEND (server.py)                     │
│                                                                     │
│  ┌──────────────────┐  ┌──────────────────┐  ┌──────────────────┐  │
│  │ /batch/balance   │  │ /account/{#}/name│  │ /account/{#}/type│  │
│  │ /full_year_refresh│ │                  │  │                  │  │
│  └────────┬─────────┘  └────────┬─────────┘  └────────┬─────────┘  │
│           │                     │                     │            │
│           ▼                     ▼                     ▼            │
│  ┌────────────────────────────────────────────────────────────┐    │
│  │                    SuiteQL Queries                          │    │
│  └────────────────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────────────────┘
          │
          ▼
┌─────────────────────────────────────────────────────────────────────┐
│                         NETSUITE API                                │
│                    (SuiteQL REST endpoint)                          │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Query Categories

### 1. LOOKUP QUERIES (Cached at Startup)

These run once when the server starts to populate lookup caches.

#### 1.1 Get Subsidiaries with Hierarchy
**Endpoint:** Startup initialization  
**Purpose:** Load all subsidiaries for dropdown and consolidation logic

```sql
SELECT id, name
FROM Subsidiary
WHERE parent IS NULL
```

Then for each parent:
```sql
SELECT id, name
FROM Subsidiary
WHERE parent = {parent_id}
```

#### 1.2 Get Departments
**Endpoint:** Startup initialization  
**Purpose:** Load departments for filtering

```sql
SELECT DISTINCT tl.department as id
FROM TransactionLine tl
WHERE tl.department IS NOT NULL AND tl.department != 0
```

Then get names:
```sql
SELECT id, name
FROM Department
WHERE id IN ({ids})
```

#### 1.3 Get Classes
**Endpoint:** Startup initialization  
**Purpose:** Load classes for filtering

```sql
SELECT DISTINCT c.id, c.name
FROM Classification c
WHERE c.id IN (
    SELECT DISTINCT tl.class
    FROM TransactionLine tl
    WHERE tl.class IS NOT NULL AND tl.class != 0
)
```

#### 1.4 Get Locations
**Endpoint:** Startup initialization  
**Purpose:** Load locations for filtering

```sql
SELECT DISTINCT l.id, l.name
FROM Location l
WHERE l.id IN (
    SELECT DISTINCT tl.location
    FROM TransactionLine tl
    WHERE tl.location IS NOT NULL AND tl.location != 0
)
```

---

### 2. PERIOD LOOKUP QUERY

#### 2.1 Get Period Dates from Name
**Endpoint:** Called by balance queries  
**Purpose:** Convert "Jan 2025" to period dates and ID

```sql
SELECT startdate, enddate, id
FROM AccountingPeriod
WHERE periodname = '{period_name}'
  AND isyear = 'F' 
  AND isquarter = 'F'
```

---

### 3. ACCOUNT METADATA QUERIES

#### 3.1 Get Account Name (Title)
**Endpoint:** `GET /account/{account_number}/name`  
**Excel Function:** `NS.GLATITLE(account)`  
**Purpose:** Get display name for an account

```sql
SELECT accountsearchdisplaynamecopy AS account_name
FROM Account
WHERE acctnumber = '{account_number}'
```

#### 3.2 Get Account Type
**Endpoint:** `GET /account/{account_number}/type`  
**Excel Function:** `NS.GLACCTTYPE(account)`  
**Purpose:** Get account type (Income, Expense, COGS, etc.)

```sql
SELECT accttype AS account_type
FROM Account
WHERE acctnumber = '{account_number}'
```

#### 3.3 Get Account Parent
**Endpoint:** `GET /account/{account_number}/parent`  
**Excel Function:** `NS.GLAPARENT(account)`  
**Purpose:** Get parent account number

```sql
SELECT p.acctnumber AS parent_number
FROM Account a
LEFT JOIN Account p ON a.parent = p.id
WHERE a.acctnumber = '{account_number}'
```

---

### 4. P&L BALANCE QUERIES

#### 4.1 Full Year P&L Query (Optimized CTE Pattern)
**Endpoint:** `POST /batch/full_year_refresh`  
**Purpose:** Get ALL P&L accounts for entire fiscal year in ONE query  
**Performance:** ~15-30 seconds for entire year

```sql
WITH sub_cte AS (
  SELECT COUNT(*) AS subs_count
  FROM Subsidiary
  WHERE isinactive = 'F'
),
base AS (
  SELECT
    tal.account AS account_id,
    t.postingperiod AS period_id,
    CASE
      WHEN (SELECT subs_count FROM sub_cte) > 1 THEN
        TO_NUMBER(
          BUILTIN.CONSOLIDATE(
            tal.amount,
            'LEDGER',
            'DEFAULT',
            'DEFAULT',
            {target_subsidiary},
            t.postingperiod,
            'DEFAULT'
          )
        )
      ELSE tal.amount
    END
    * CASE WHEN a.accttype IN ('Income','OthIncome') THEN -1 ELSE 1 END
    AS cons_amt
  FROM TransactionAccountingLine tal
  JOIN Transaction t ON t.id = tal.transaction
  JOIN Account a ON a.id = tal.account
  JOIN AccountingPeriod ap ON ap.id = t.postingperiod
  CROSS JOIN sub_cte
  WHERE t.posting = 'T'
    AND tal.posting = 'T'
    AND tal.accountingbook = 1
    AND ap.isyear = 'F'
    AND ap.isquarter = 'F'
    AND EXTRACT(YEAR FROM ap.startdate) = {fiscal_year}
    AND COALESCE(a.eliminate, 'F') = 'F'
    AND a.accttype IN ('Income','COGS','Cost of Goods Sold','Expense','OthIncome','OthExpense')
    {optional_filters}
)
SELECT
  a.acctnumber AS account_number,
  a.accttype AS account_type,
  TO_CHAR(ap.startdate,'YYYY-MM') AS month,
  SUM(b.cons_amt) AS amount
FROM base b
JOIN AccountingPeriod ap ON ap.id = b.period_id
JOIN Account a ON a.id = b.account_id
GROUP BY a.acctnumber, a.accttype, ap.startdate
ORDER BY a.acctnumber, ap.startdate
```

**Key Features:**
- Uses CTE to consolidate FIRST, then group (much faster)
- BUILTIN.CONSOLIDATE handles multi-subsidiary consolidation
- Sign flip for Income/OthIncome (NetSuite stores as negative)
- Filters by fiscal year using EXTRACT(YEAR FROM ap.startdate)

#### 4.2 Batch P&L Query (Specific Accounts/Periods)
**Endpoint:** `POST /batch/balance`  
**Purpose:** Get balances for specific accounts and periods

```sql
SELECT 
    a.acctnumber,
    ap.periodname,
    SUM(cons_amt) AS balance
FROM (
    SELECT
        tal.account,
        t.postingperiod,
        CASE
            WHEN subs_count > 1 THEN
                TO_NUMBER(
                    BUILTIN.CONSOLIDATE(
                        tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT',
                        {target_sub}, t.postingperiod, 'DEFAULT'
                    )
                )
            ELSE tal.amount
        END
        * CASE WHEN a.accttype IN ('Income','OthIncome') THEN -1 ELSE 1 END AS cons_amt
    FROM TransactionAccountingLine tal
        JOIN Transaction t ON t.id = tal.transaction
        JOIN Account a ON a.id = tal.account
        JOIN AccountingPeriod apf ON apf.id = t.postingperiod
        CROSS JOIN (
            SELECT COUNT(*) AS subs_count
            FROM Subsidiary
            WHERE isinactive = 'F'
        ) subs_cte
    WHERE t.posting = 'T'
        AND tal.posting = 'T'
        AND a.acctnumber IN ({account_list})
        AND apf.periodname IN ({period_list})
        AND a.accttype IN ('Income', 'OthIncome', 'COGS', 'Expense', 'OthExpense')
        AND COALESCE(a.eliminate, 'F') = 'F'
) x
JOIN Account a ON a.id = x.account
JOIN AccountingPeriod ap ON ap.id = x.postingperiod
GROUP BY a.acctnumber, ap.periodname
ORDER BY a.acctnumber, ap.periodname
```

---

### 5. BALANCE SHEET QUERIES

Balance Sheet accounts require CUMULATIVE balances (inception through period end), not period-specific activity.

#### 5.1 Full Year Balance Sheet Query (Wide Format)
**Endpoint:** `POST /batch/full_year_refresh` (second part)  
**Purpose:** Get ALL BS accounts with cumulative balances for each month

```sql
SELECT
  a.acctnumber AS account_number,
  a.accttype AS account_type,
  SUM(
    CASE 
      WHEN t.trandate <= p_jan.enddate 
      THEN TO_NUMBER(BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT', {target_sub}, p_jan.id, 'DEFAULT'))
      ELSE 0
    END
  ) AS Jan_{year},
  SUM(
    CASE 
      WHEN t.trandate <= p_feb.enddate 
      THEN TO_NUMBER(BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT', {target_sub}, p_feb.id, 'DEFAULT'))
      ELSE 0
    END
  ) AS Feb_{year},
  -- ... (repeat for all 12 months)
FROM TransactionAccountingLine tal
  INNER JOIN Transaction t ON t.id = tal.transaction
  INNER JOIN Account a ON a.id = tal.account
  CROSS JOIN (
    SELECT id, enddate FROM AccountingPeriod 
    WHERE TO_CHAR(startdate, 'YYYY-MM') = '{year}-01' 
      AND isquarter = 'F' AND isyear = 'F'
    FETCH FIRST 1 ROWS ONLY
  ) p_jan
  CROSS JOIN (
    SELECT id, enddate FROM AccountingPeriod 
    WHERE TO_CHAR(startdate, 'YYYY-MM') = '{year}-02' 
      AND isquarter = 'F' AND isyear = 'F'
    FETCH FIRST 1 ROWS ONLY
  ) p_feb
  -- ... (repeat for all 12 months)
WHERE 
  t.posting = 'T'
  AND tal.posting = 'T'
  AND tal.accountingbook = 1
  AND COALESCE(a.eliminate, 'F') = 'F'
  AND a.accttype NOT IN ('Income', 'OthIncome', 'COGS', 'Cost of Goods Sold', 'Expense', 'OthExpense')
  AND t.trandate <= p_dec.enddate
GROUP BY 
  a.acctnumber, 
  a.accttype
ORDER BY 
  a.acctnumber
```

**Key Features:**
- Uses CASE WHEN with 12 CROSS JOINs (one per period)
- Each period returns exactly ONE row (via FETCH FIRST 1 ROWS ONLY)
- Cumulative balance: `t.trandate <= period.enddate`
- Excludes P&L account types

#### 5.2 Single Period Balance Sheet Query
**Endpoint:** `POST /batch/balance`  
**Purpose:** Get BS balances for specific accounts/period

```sql
SELECT 
    a.acctnumber,
    SUM(cons_amt) AS balance
FROM TransactionAccountingLine tal
    JOIN Transaction t ON t.id = tal.transaction
    JOIN Account a ON a.id = tal.account
    CROSS JOIN (
        SELECT COUNT(*) AS subs_count
        FROM Subsidiary
        WHERE isinactive = 'F'
    ) subs_cte
WHERE t.posting = 'T'
    AND tal.posting = 'T'
    AND t.trandate <= '{period_end_date}'
    AND a.acctnumber IN ({account_list})
    AND a.accttype NOT IN ('Income', 'OthIncome', 'COGS', 'Expense', 'OthExpense')
    AND COALESCE(a.eliminate, 'F') = 'F'
GROUP BY a.acctnumber
```

---

### 6. TRANSACTION DRILL-DOWN QUERY

**Endpoint:** `GET /transactions`  
**Purpose:** Get transaction details for a specific account/period

```sql
SELECT 
    t.tranid,
    t.trandate,
    t.type,
    t.status,
    t.memo,
    tal.amount,
    tal.debit,
    tal.credit
FROM TransactionAccountingLine tal
JOIN Transaction t ON t.id = tal.transaction
JOIN Account a ON a.id = tal.account
JOIN AccountingPeriod ap ON ap.id = t.postingperiod
WHERE t.posting = 'T'
    AND tal.posting = 'T'
    AND a.acctnumber = '{account_number}'
    AND ap.periodname = '{period_name}'
ORDER BY t.trandate DESC
```

---

### 7. GUIDE ME ACCOUNT LIST QUERY

**Endpoint:** `GET /lookups/accounts`  
**Purpose:** Get accounts for Guide Me wizard

```sql
SELECT 
    acctnumber AS number,
    accountsearchdisplaynamecopy AS name,
    accttype AS type
FROM Account
WHERE isinactive = 'F'
  AND accttype = 'Income'
ORDER BY acctnumber
```

---

### 8. BULK ACCOUNT NAMES QUERY

**Endpoint:** `POST /batch/full_year_refresh` (third part)  
**Purpose:** Get all account names in ONE query (prevents 429 concurrency errors)

```sql
SELECT acctnumber AS number, accountsearchdisplaynamecopy AS name
FROM Account
WHERE acctnumber IN ('{acct1}', '{acct2}', '{acct3}', ...)
```

---

## Performance Considerations

### Query Execution Times (Typical)
| Query Type | Typical Time | Notes |
|------------|--------------|-------|
| Account name lookup | <1 sec | Single row |
| Account type lookup | <1 sec | Single row |
| Full Year P&L | 15-30 sec | All accounts × 12 months |
| Full Year BS | 60-90 sec | Complex CONSOLIDATE calls |
| Batch balance (10 accts × 12 periods) | 5-10 sec | |
| Transaction drill-down | 2-5 sec | Depends on volume |

### Pagination
NetSuite limits SuiteQL results to 1000 rows. We use API-level pagination:
```
POST /query/v1/suiteql?limit=1000&offset=0
POST /query/v1/suiteql?limit=1000&offset=1000
...
```

### Caching Strategy
1. **Backend cache:** 5-minute TTL for balance data
2. **Frontend localStorage:** Balance, type, name caches
3. **In-memory cache:** Populated via Shared Runtime

---

## Questions for Review

1. **P&L Query:** Is the CTE pattern optimal for consolidation?
2. **BS Query:** Can the 12 CROSS JOINs be optimized?
3. **Sign Flip:** Is `* -1` for Income/OthIncome correct for all scenarios?
4. **Filters:** Are department/class/location filters applied correctly?
5. **BUILTIN.CONSOLIDATE:** Are the parameters optimal for multi-subsidiary environments?

---

*Generated: December 4, 2025*
*Add-in Version: 1.4.33.0*

