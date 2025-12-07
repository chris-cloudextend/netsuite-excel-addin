# NetSuite Excel Add-in - Technical Summary

## Project Overview

**XAVI for NetSuite** is an Excel Add-in that provides custom functions to retrieve financial data from NetSuite via SuiteQL queries. Users can build financial reports in Excel using formulas like `=XAVI.BALANCE("4010", "Jan 2025", "Jan 2025")` to fetch account balances. The add-in handles multi-currency consolidation, multi-book accounting, intelligent caching, and performance optimization.

### Key Features
- **Custom Excel Functions**: `XAVI.BALANCE`, `XAVI.BUDGET`, `XAVI.NAME`, `XAVI.TYPE`, `XAVI.PARENT`
- **Multi-Book Accounting**: Support for GAAP, IFRS, Tax, and other accounting books
- **Multi-currency Consolidation**: Uses NetSuite's `BUILTIN.CONSOLIDATE` function
- **Smart Caching**: In-memory and localStorage caching with intelligent invalidation
- **Build Mode Detection**: Detects rapid formula entry (drag-drop) and batches requests
- **Optimized Queries**: Different strategies for P&L vs Balance Sheet accounts
- **Smart Period Expansion**: Automatically pre-caches adjacent months during drag operations

---

## Custom Functions Reference

### XAVI.BALANCE
Get GL account balance with full filter support.

```
=XAVI.BALANCE(account, fromPeriod, toPeriod, [subsidiary], [department], [location], [classId], [accountingBook])
```

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `account` | Text/Number | ✅ | Account number (e.g., "4010" or 4010) |
| `fromPeriod` | Text/Date | ✅ | Starting period (e.g., "Jan 2025" or 1/1/2025) |
| `toPeriod` | Text/Date | ✅ | Ending period (same as fromPeriod for single month) |
| `subsidiary` | Text/Number | ❌ | Subsidiary name or ID (e.g., "Celigo Inc." or 1) |
| `department` | Text/Number | ❌ | Department name or ID |
| `location` | Text/Number | ❌ | Location name or ID |
| `classId` | Text/Number | ❌ | Class name or ID |
| `accountingBook` | Number | ❌ | Accounting Book ID (default: 1 = Primary Book) |

**Examples:**
```
=XAVI.BALANCE("4010", "Jan 2025", "Jan 2025")
=XAVI.BALANCE("4010", "Jan 2025", "Dec 2025", "Celigo Inc.")
=XAVI.BALANCE("4010", A1, A1, $B$1, , , , 2)  ← Using Secondary Book (ID 2)
```

---

### XAVI.BUDGET
Get budget amount with filters.

```
=XAVI.BUDGET(account, fromPeriod, toPeriod, [subsidiary], [department], [location], [classId], [accountingBook])
```

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `account` | Text/Number | ✅ | Account number |
| `fromPeriod` | Text/Date | ✅ | Starting period |
| `toPeriod` | Text/Date | ✅ | Ending period |
| `subsidiary` | Text/Number | ❌ | Subsidiary filter |
| `department` | Text/Number | ❌ | Department filter |
| `location` | Text/Number | ❌ | Location filter |
| `classId` | Text/Number | ❌ | Class filter |
| `accountingBook` | Number | ❌ | Accounting Book ID (default: Primary Book) |

**Example:**
```
=XAVI.BUDGET("4010", "Jan 2025", "Dec 2025", "Celigo Inc.")
```

---

### XAVI.NAME
Get account name from account number.

```
=XAVI.NAME(accountNumber)
```

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `accountNumber` | Text/Number | ✅ | Account number or ID |

**Example:**
```
=XAVI.NAME("4010")  → "Product Revenue"
```

---

### XAVI.TYPE
Get account type (Income, Expense, Bank, etc.).

```
=XAVI.TYPE(accountNumber)
```

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `accountNumber` | Text/Number | ✅ | Account number or ID |

**Example:**
```
=XAVI.TYPE("4010")  → "Income"
=XAVI.TYPE("1010")  → "Bank"
```

---

### XAVI.PARENT
Get parent account number (for sub-accounts).

```
=XAVI.PARENT(accountNumber)
```

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `accountNumber` | Text/Number | ✅ | Account number or ID |

**Example:**
```
=XAVI.PARENT("4010-1")  → "4010"
```

---

## Multi-Book Accounting Support

NetSuite's Multi-Book Accounting feature allows organizations to maintain multiple sets of books for different accounting standards or purposes.

### Common Accounting Books
| ID | Book Type | Use Case |
|----|-----------|----------|
| 1 | Primary Book | Main GAAP/local statutory reporting |
| 2+ | Secondary Books | IFRS, Tax, Management, etc. |

### How It Works

1. **Backend**: All SuiteQL queries include `accountingbook` filter:
   ```sql
   AND tal.accountingbook = {accountingbook}
   ```

2. **Frontend**: BALANCE and BUDGET functions accept optional `accountingBook` parameter

3. **Taskpane UI**: "Accounting Book" dropdown in Filters section

### Formula Example with Multi-Book
```
-- Primary Book (default)
=XAVI.BALANCE("4010", "Jan 2025", "Jan 2025")

-- Secondary Book (IFRS)
=XAVI.BALANCE("4010", "Jan 2025", "Jan 2025", "", "", "", "", 2)
```

---

## Architecture

```
┌─────────────────────┐     ┌─────────────────────┐     ┌─────────────────┐
│   Excel Add-in      │────▶│   Cloudflare        │────▶│   Flask Backend │
│   (functions.js)    │     │   Worker + Tunnel   │     │   (server.py)   │
│                     │◀────│                     │◀────│                 │
│   - Custom funcs    │     │   - CORS proxy      │     │   - SuiteQL     │
│   - Caching         │     │   - TLS termination │     │   - OAuth1      │
│   - Build mode      │     │                     │     │   - Caching     │
└─────────────────────┘     └─────────────────────┘     └─────────────────┘
         │                                                       │
         ▼                                                       ▼
┌─────────────────────┐                              ┌─────────────────┐
│   Taskpane UI       │                              │    NetSuite     │
│   (taskpane.html)   │                              │    SuiteQL API  │
│                     │                              │                 │
│   - Prep Data       │                              │   - Account     │
│   - Refresh All     │                              │   - Transaction │
│   - Refresh Selected│                              │   - Period      │
│   - Filter Lookups  │                              │   - Budget      │
└─────────────────────┘                              └─────────────────┘
```

---

## Security Architecture

### 1. Credential Storage ✅

**Implementation:**
```
backend/netsuite_config.json (NOT in git - .gitignore protected)
{
  "account_id": "589861",
  "consumer_key": "...",
  "consumer_secret": "...",
  "token_id": "...",
  "token_secret": "..."
}
```

**Security measures:**
- ✅ File in `.gitignore` (pattern `**/netsuite_config*.json`)
- ✅ OAuth1 with HMAC-SHA256 signatures
- ✅ Token-Based Authentication (TBA)

### 2. Transport Security ✅

**Full chain is HTTPS:**
```
Excel Add-in (HTTPS) → Cloudflare Worker (HTTPS) → Cloudflare Tunnel (Encrypted) → Flask (localhost:5002)
     ↑                        ↑                           ↑                              ↑
   TLS 1.3              Edge TLS               Encrypted Tunnel                    Local only
```

- ✅ All external traffic over HTTPS
- ✅ Cloudflare provides TLS termination and DDoS protection
- ✅ Tunnel encrypts local traffic
- ✅ Flask only binds to localhost (not exposed to internet)

### 3. Cloudflare Worker Role ✅

The worker is a **CORS proxy only** - no credentials are stored there:

```javascript
export default {
  async fetch(request) {
    const TUNNEL_URL = 'https://xxx.trycloudflare.com';
    // Simply forwards requests - no secrets stored
    const response = await fetch(TUNNEL_URL + url.pathname, {...});
    // Adds CORS headers and returns
  }
}
```

### 4. SQL Injection Prevention ✅

All user inputs are sanitized:

```python
def escape_sql(text):
    """Escape single quotes in SQL strings"""
    if text is None:
        return ""
    return str(text).replace("'", "''")

# Usage:
accounts_in = ','.join([f"'{escape_sql(acc)}'" for acc in accounts])
```

### 5. Client-Side Caching ⚠️

**What's cached in localStorage:**
- Account balances (5-minute TTL)
- Account names
- Account types

**Mitigations:**
- ✅ 5-minute TTL (data expires quickly)
- ✅ No PII or authentication tokens cached
- ✅ "Clear Cache" button in taskpane
- ⚠️ Consider sessionStorage for shared machines

### Security Checklist

| Area | Status | Notes |
|------|--------|-------|
| Credentials in git | ✅ | `.gitignore` blocks all config files |
| HTTPS for all traffic | ✅ | Cloudflare provides TLS |
| OAuth1 authentication | ✅ | HMAC-SHA256 signature |
| SQL injection | ✅ | `escape_sql()` function |
| CORS configured | ✅ | Worker adds headers |
| Client cache exposure | ⚠️ | 5-min TTL mitigates risk |

---

## Balance Sheet vs Income Statement Approaches

### The Core Challenge

**P&L (Income Statement) accounts** show activity for a specific period:
- Revenue for January 2025 = Sum of all revenue transactions in January 2025

**Balance Sheet accounts** show cumulative balances:
- Cash as of January 2025 = Sum of all cash transactions from inception through January 2025

Additionally, for **multi-currency** environments, Balance Sheet accounts must be revalued at the reporting period's exchange rate.

---

### P&L Query Approach (Optimized Pivoted Query)

**Strategy**: Single query fetches ALL P&L accounts for an entire fiscal year in ~6-15 seconds.

```sql
SELECT
  a.acctnumber AS account_number,
  a.accttype AS account_type,
  SUM(CASE WHEN TO_CHAR(ap.startdate,'YYYY-MM')='2025-01' THEN cons_amt ELSE 0 END) AS jan,
  SUM(CASE WHEN TO_CHAR(ap.startdate,'YYYY-MM')='2025-02' THEN cons_amt ELSE 0 END) AS feb,
  -- ... all 12 months
FROM (
  SELECT
    tal.account,
    t.postingperiod,
    CASE
      WHEN subs_count > 1 THEN
        TO_NUMBER(BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT',
            {target_sub}, t.postingperiod, 'DEFAULT'))
      ELSE tal.amount
    END
    * CASE WHEN a.accttype IN ('Income','OthIncome') THEN -1 ELSE 1 END AS cons_amt
  FROM transactionaccountingline tal
    JOIN transaction t ON t.id = tal.transaction
    JOIN account a ON a.id = tal.account
    JOIN accountingperiod apf ON apf.id = t.postingperiod
    CROSS JOIN (SELECT COUNT(*) AS subs_count FROM subsidiary WHERE isinactive = 'F') subs_cte
  WHERE t.posting = 'T'
    AND tal.posting = 'T'
    AND tal.accountingbook = {accountingbook}
    AND apf.isyear = 'F' AND apf.isquarter = 'F'
    AND TO_CHAR(apf.startdate,'YYYY') = '2025'
    AND a.accttype IN ('Income', 'COGS', 'Cost of Goods Sold', 'Expense', 'OthIncome', 'OthExpense')
) x
JOIN accountingperiod ap ON ap.id = x.postingperiod
JOIN account a ON a.id = x.account
GROUP BY a.acctnumber, a.accttype
```

---

### Balance Sheet Query Approach (Multi-Period with Fixed Consolidation)

**Strategy**: Single query fetches ALL Balance Sheet accounts for requested periods in ~60-70 seconds.

**Critical**: `BUILTIN.CONSOLIDATE` uses the **target period's ID** for correct exchange rates.

```sql
SELECT 
  a.acctnumber AS account_number,
  a.accttype AS account_type,
  SUM(CASE WHEN ap.startdate <= p_2025_01.enddate
    THEN TO_NUMBER(BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT', 
         {target_sub}, p_2025_01.id, 'DEFAULT'))
         * CASE WHEN a.accttype IN ('AcctPay', 'CredCard', 'OthCurrLiab', 'LongTermLiab', 
                'DeferRevenue', 'Equity', 'RetainedEarnings', 'Income', 'OthIncome') 
                THEN -1 ELSE 1 END
    ELSE 0 END) AS bal_2025_01
FROM transactionaccountingline tal
  INNER JOIN transaction t ON t.id = tal.transaction
  INNER JOIN account a ON a.id = tal.account
  INNER JOIN accountingperiod ap ON ap.id = t.postingperiod
  INNER JOIN accountingperiod p_2025_01 
    ON TO_CHAR(p_2025_01.startdate, 'YYYY-MM') = '2025-01' 
    AND p_2025_01.isquarter = 'F' AND p_2025_01.isyear = 'F'
WHERE 
  t.posting = 'T'
  AND tal.posting = 'T'
  AND tal.accountingbook = {accountingbook}
  AND a.accttype NOT IN ('Income', 'COGS', 'Cost of Goods Sold', 'Expense', 'OthIncome', 'OthExpense')
  AND ap.startdate <= p_2025_01.enddate
  AND ap.isyear = 'F' AND ap.isquarter = 'F'
GROUP BY a.acctnumber, a.accttype
```

---

## Account Type Reference

### Balance Sheet Accounts

```
ASSETS (Debit balance - stored positive, NO sign flip):
  Bank              Bank/Cash accounts
  AcctRec           Accounts Receivable
  OthCurrAsset      Other Current Asset
  FixedAsset        Fixed Asset
  OthAsset          Other Asset
  DeferExpense      Deferred Expense
  UnbilledRec       Unbilled Receivable

LIABILITIES (Credit balance - stored negative, FLIP × -1):
  AcctPay           Accounts Payable
  CredCard          Credit Card (NOT "CreditCard")
  OthCurrLiab       Other Current Liability
  LongTermLiab      Long Term Liability
  DeferRevenue      Deferred Revenue

EQUITY (Credit balance - stored negative, FLIP × -1):
  Equity            Equity accounts
  RetainedEarnings  Retained Earnings
```

### Profit & Loss Accounts

```
INCOME (Credit balance - stored negative, FLIP × -1):
  Income            Revenue/Sales
  OthIncome         Other Income

EXPENSES (Debit balance - stored positive, NO sign flip):
  COGS              Cost of Goods Sold
  Expense           Operating Expense
  OthExpense        Other Expense
```

### Excluded Account Types
```
  NonPosting        Statistical/Non-posting (no transactions)
  Stat              Statistical accounts (non-financial KPIs)
```

---

## Frontend Behavior Modes

### 1. Single Cell Resolution
User enters a formula → check cache → if miss, queue for batch processing with 500ms delay.

### 2. Build Mode (Drag-Drop)
Detects rapid formula entry, batches all requests, groups by filter combination, then:
- Fetches BS accounts using `batch/bs_periods`
- Fetches P&L accounts using `batch/full_year_refresh`

### 3. Smart Period Expansion
When dragging formulas, automatically pre-caches one month before and after the requested range:
```
User drags: Jan 2025, Feb 2025
System fetches: Dec 2024, Jan 2025, Feb 2025, Mar 2025
```

### 4. Refresh All (Taskpane)
Scans sheet for all XAVI formulas, extracts accounts/periods, fetches fresh data.

### 5. Refresh Selected
Only refreshes the selected cells using `__CLEARCACHE__` command.

---

## Backend Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/batch/full_year_refresh` | POST | Fetch all P&L accounts for fiscal year |
| `/batch/bs_periods` | POST | Fetch all BS accounts for specific periods |
| `/batch/balance` | POST | Fetch specific accounts for specific periods |
| `/batch/account_types` | POST | Get account types for a list of accounts |
| `/lookups/all` | GET | Get all filter lookups (subsidiaries, departments, classes, locations, accounting books) |
| `/lookups/accountingbooks` | GET | Get accounting books list |
| `/account/name` | POST | Get account name |
| `/account/type` | POST | Get account type |
| `/account/parent` | POST | Get parent account |
| `/budget` | GET | Get budget amount |

---

## Caching Strategy

### Frontend Caches
1. **In-memory cache** (`cache.balance`): Fastest, lost on page reload
2. **localStorage cache**: Persists across sessions, 5-minute TTL
3. **fullYearCache**: Stores complete year data for instant lookup

### Backend Cache
- `balance_cache`: Server-side cache for fast repeat queries
- `account_title_cache`: Account names preloaded at startup
- `lookup_cache`: Subsidiaries, departments, etc.

---

## File Structure

```
├── backend/
│   ├── server.py              # Flask API server
│   ├── constants.py           # Account type constants
│   └── netsuite_config.json   # Credentials (NOT in git)
├── docs/
│   ├── functions.js           # Custom Excel functions + caching
│   ├── functions.json         # Function metadata for Excel
│   ├── functions.html         # Functions runtime page
│   └── taskpane.html          # Taskpane UI + JavaScript
├── excel-addin/
│   └── manifest-claude.xml    # Excel add-in manifest
├── CLOUDFLARE-WORKER-CODE.js  # Proxy for production deployment
└── PROJECT_SUMMARY.md         # This file
```

---

## Key Technical Decisions

1. **SuiteQL doesn't support CTEs well**: Use CROSS JOIN subqueries instead
2. **BUILTIN.CONSOLIDATE needs fixed period ID**: For BS accounts, use target period's ID
3. **Sign conventions matter**: Liabilities/Equity stored as negative, must flip for display
4. **Account type spelling**: NetSuite uses `CredCard` not `CreditCard`
5. **Multi-Book Accounting**: All queries filter by `accountingbook` (default: 1)
6. **Smart period expansion**: Pre-cache adjacent months for better UX

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.5.3.0 | Dec 2025 | Multi-Book Accounting support, accountingBook parameter |
| 1.5.2.0 | Dec 2025 | Smart period expansion for caching |
| 1.5.1.0 | Dec 2025 | Fixed XAVI namespace in manifest |
| 1.5.0.0 | Dec 2025 | Renamed NS.GLA* to XAVI.* functions |
| 1.4.90.0 | Dec 2025 | GET to POST endpoint migration, constants refactor |
| 1.4.89.0 | Dec 2025 | Complete account types documentation |
| 1.4.88.0 | Dec 2025 | Fixed CredCard type mismatch |
| 1.4.87.0 | Dec 2025 | Balance Sheet sign convention fix |

---

## Production Deployment Checklist

- [ ] Update Cloudflare Worker with new tunnel URL
- [ ] Verify GitHub Pages deployment (~1 minute after push)
- [ ] Bump manifest version for cache-busting
- [ ] Test all functions with different accounting books
- [ ] Verify filter dropdowns populate correctly
- [ ] Test drag-drop caching behavior
- [ ] Confirm multi-currency consolidation accuracy
