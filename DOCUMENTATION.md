# XAVI for NetSuite - Complete Documentation

## Table of Contents

1. [Overview](#overview)
2. [For Finance Users (CPA Perspective)](#for-finance-users-cpa-perspective)
3. [For Engineers (Technical Reference)](#for-engineers-technical-reference)
4. [SuiteQL Deep Dive](#suiteql-deep-dive)
5. [BUILTIN.CONSOLIDATE Explained](#builtinconsolidate-explained)
6. [Account Types & Sign Conventions](#account-types--sign-conventions)
7. [Troubleshooting](#troubleshooting)

---

# Overview

**XAVI for NetSuite** is an Excel Add-in that provides custom formulas to retrieve financial data directly from NetSuite. Finance teams can build dynamic reports in Excel that pull live data from their ERP.

### Why XAVI?

| Traditional Approach | With XAVI |
|---------------------|-----------|
| Export CSV from NetSuite | Live formulas pull data on demand |
| Manual copy/paste | Auto-refresh with one click |
| Stale data within hours | Real-time accuracy |
| Breaking links when structure changes | Dynamic account references |

### Available Functions

| Function | Purpose | Example |
|----------|---------|---------|
| `XAVI.BALANCE` | Get GL account balance | `=XAVI.BALANCE("4010", "Jan 2025", "Jan 2025")` |
| `XAVI.BUDGET` | Get budget amount | `=XAVI.BUDGET("4010", "Jan 2025", "Dec 2025")` |
| `XAVI.NAME` | Get account name | `=XAVI.NAME("4010")` → "Product Revenue" |
| `XAVI.TYPE` | Get account type | `=XAVI.TYPE("4010")` → "Income" |
| `XAVI.PARENT` | Get parent account | `=XAVI.PARENT("4010-1")` → "4010" |
| `XAVI.RETAINEDEARNINGS` | Calculate Retained Earnings | `=XAVI.RETAINEDEARNINGS("Dec 2024")` |
| `XAVI.NETINCOME` | Calculate Net Income YTD | `=XAVI.NETINCOME("Mar 2025")` |
| `XAVI.CTA` | Calculate Cumulative Translation Adjustment | `=XAVI.CTA("Dec 2024")` |

---

# For Finance Users (CPA Perspective)

## Understanding the Formulas

### XAVI.BALANCE - The Foundation

This is your primary formula for building financial statements. It retrieves the balance for any GL account for any period.

**Syntax:**
```
=XAVI.BALANCE(account, fromPeriod, toPeriod, [subsidiary], [department], [location], [class], [accountingBook])
```

**For Balance Sheet accounts** (Assets, Liabilities, Equity):
- The formula returns the **cumulative balance** as of period end
- This matches how NetSuite displays Balance Sheet balances
- Example: Cash as of Jan 2025 = all cash transactions from inception through Jan 31, 2025

**For Income Statement accounts** (Revenue, Expenses):
- The formula returns **activity for the period**
- This matches NetSuite's P&L presentation
- Example: Revenue for Jan 2025 = only January revenue transactions

### Why Are RE, NI, and CTA Separate?

NetSuite doesn't store Retained Earnings, Net Income, or CTA as actual account balances. Instead, it **calculates them dynamically** when you run reports. XAVI replicates these calculations:

#### Retained Earnings
```
RE = All P&L from company inception through prior fiscal year end
   + Any journal entries posted directly to Retained Earnings accounts
```

**When to use:** Balance Sheet reports showing the equity section.

#### Net Income
```
NI = All P&L from fiscal year start through the report period
```

**When to use:** Balance Sheet reports (completes equity section) or to verify P&L totals.

#### CTA (Cumulative Translation Adjustment)

**What is CTA?**
In multi-currency companies, when you translate foreign subsidiary balances to your reporting currency:
- Balance Sheet accounts translate at **period-end exchange rate**
- Income Statement accounts translate at **average or transaction rate**
- The difference creates an imbalance → CTA is the "plug" that balances the Balance Sheet

**Why the "plug method"?**
NetSuite calculates additional translation adjustments at runtime that are never posted to any account. The only way to get 100% accuracy is:
```
CTA = (Total Assets - Total Liabilities) - Posted Equity - Retained Earnings - Net Income
```

This guarantees Assets = Liabilities + Equity, matching NetSuite exactly.

### Multi-Book Accounting

If your organization maintains multiple sets of books (GAAP, IFRS, Tax, etc.), use the `accountingBook` parameter:

```
=XAVI.BALANCE("4010", "Jan 2025", "Jan 2025", "", "", "", "", 1)  ← Primary Book
=XAVI.BALANCE("4010", "Jan 2025", "Jan 2025", "", "", "", "", 2)  ← Secondary Book (IFRS)
```

The accounting book ID can be found in NetSuite under Setup → Accounting → Accounting Books.

### Consolidation

When running consolidated reports across subsidiaries:
1. Use the parent subsidiary name or ID in the subsidiary parameter
2. XAVI will automatically consolidate all child subsidiaries
3. Foreign currency amounts are translated at the appropriate exchange rates

```
=XAVI.BALANCE("4010", "Dec 2024", "Dec 2024", "Parent Company")
```

### Best Practices

1. **Use Refresh Accounts** before presenting reports - ensures all data is fresh
2. **Recalculate Retained Earnings** separately - these calculations take 30-60 seconds each
3. **Reference periods from cells** - makes it easy to change the report date
4. **Use the subsidiary hierarchy** - parent subsidiaries automatically include children

---

# For Engineers (Technical Reference)

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
└─────────────────────┘                              └─────────────────┘
```

## Key Design Decisions

### 1. Why SuiteQL over Saved Searches?
- **Flexibility:** SQL-like queries can be constructed dynamically
- **Performance:** Single query can return multiple accounts/periods
- **Consolidation:** `BUILTIN.CONSOLIDATE` function handles currency translation

### 2. Why Cloudflare Tunnel?
- Excel Add-ins require HTTPS
- Quick tunnels provide free HTTPS endpoint
- Worker provides stable URL (tunnel URL changes on restart)

### 3. Why Separate RE/NI/CTA?
NetSuite doesn't expose these as queryable account balances. They're calculated at report runtime. We replicate the calculation using the same logic.

### 4. Batching Strategy
When users drag formulas, we detect rapid formula creation and batch them:
- **Build Mode:** 3+ formulas in 500ms triggers batching
- **Batch Delay:** 150-500ms to collect requests
- **Chunking:** Max 50 accounts per request to avoid timeouts

## File Structure

```
├── backend/
│   ├── server.py              # Flask API server (SuiteQL queries)
│   ├── constants.py           # Account type constants
│   ├── requirements.txt       # Python dependencies
│   └── netsuite_config.json   # Credentials (gitignored)
├── docs/
│   ├── functions.js           # Custom Excel functions + caching
│   ├── functions.json         # Function metadata for Excel
│   ├── functions.html         # Functions runtime page
│   ├── taskpane.html          # Taskpane UI + refresh logic
│   ├── commands.html          # Ribbon commands
│   └── commands.js            # Command implementations
├── excel-addin/
│   └── manifest-claude.xml    # Excel add-in manifest
├── CLOUDFLARE-WORKER-CODE.js  # Proxy worker code
└── DOCUMENTATION.md           # This file
```

## Backend Endpoints

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/batch/full_year_refresh` | POST | Fetch all P&L accounts for fiscal year |
| `/batch/bs_periods` | POST | Fetch all BS accounts for specific periods |
| `/batch/balance` | POST | Fetch specific accounts for specific periods |
| `/batch/account_types` | POST | Get account types for classification |
| `/retained-earnings` | POST | Calculate Retained Earnings |
| `/net-income` | POST | Calculate Net Income |
| `/cta` | POST | Calculate CTA |
| `/account/name` | POST | Get account name |
| `/account/type` | POST | Get account type |
| `/lookups/all` | GET | Get filter lookups |

## Caching Strategy

### Three-Tier Cache

```
┌─────────────────────────────────────────────────────────┐
│  TIER 1: In-Memory Cache (functions.js)                 │
│  - Fastest: microseconds                                │
│  - Lost on page refresh                                 │
│  - Map: cache.balance.set(key, value)                   │
├─────────────────────────────────────────────────────────┤
│  TIER 2: localStorage Cache                             │
│  - Fast: milliseconds                                   │
│  - Persists across taskpane refreshes                   │
│  - TTL: 5 minutes                                       │
├─────────────────────────────────────────────────────────┤
│  TIER 3: Backend Cache (server.py)                      │
│  - Medium: avoids NetSuite query                        │
│  - TTL: 5 minutes                                       │
└─────────────────────────────────────────────────────────┘
```

### Cache Key Format
```
{account}:{period}:{subsidiary}:{department}:{location}:{class}:{book}
Example: "4220:Jan 2024:1::::1"
```

---

# SuiteQL Deep Dive

## Core Tables

| Table | Purpose | Key Fields |
|-------|---------|------------|
| `TransactionAccountingLine` | Actual posted amounts | `account`, `amount`, `posting`, `accountingbook` |
| `Transaction` | Transaction header | `id`, `trandate`, `postingperiod`, `posting` |
| `Account` | Chart of Accounts | `acctnumber`, `accttype`, `fullname` |
| `AccountingPeriod` | Fiscal periods | `id`, `periodname`, `startdate`, `enddate` |
| `Subsidiary` | Legal entities | `id`, `name`, `parent`, `iselimination` |

## Standard Balance Query Structure

```sql
SELECT 
    a.acctnumber,
    SUM(cons_amount) AS balance
FROM TransactionAccountingLine tal
    JOIN Transaction t ON t.id = tal.transaction
    JOIN Account a ON a.id = tal.account
    JOIN AccountingPeriod ap ON ap.id = t.postingperiod
WHERE t.posting = 'T'                    -- Only posted transactions
  AND tal.posting = 'T'                  -- Only posting lines
  AND tal.accountingbook = {book_id}     -- Specific accounting book
  AND a.acctnumber IN ({accounts})       -- Filter accounts
  AND ap.periodname IN ({periods})       -- Filter periods
GROUP BY a.acctnumber
```

## P&L vs Balance Sheet Queries

**The fundamental difference:**

| Type | Date Filter | What it Returns |
|------|-------------|-----------------|
| P&L (Income Statement) | `ap.periodname IN ('Jan 2025')` | Activity for the period |
| Balance Sheet | `ap.enddate <= '2025-01-31'` | Cumulative balance through period end |

**P&L Query:**
```sql
WHERE ...
  AND ap.periodname IN ('Jan 2025', 'Feb 2025')  -- Specific periods only
  AND a.accttype IN ('Income', 'Expense', 'COGS', ...)
```

**Balance Sheet Query:**
```sql
WHERE ...
  AND ap.enddate <= TO_DATE('2025-01-31', 'YYYY-MM-DD')  -- All time through period
  AND a.accttype IN ('Bank', 'AcctRec', 'AcctPay', ...)
```

---

# BUILTIN.CONSOLIDATE Explained

## Why It's Critical

In multi-currency, multi-subsidiary environments, `BUILTIN.CONSOLIDATE` is the **only way** to get correct consolidated amounts. It handles:

1. **Currency Translation:** Converts foreign currency to reporting currency
2. **Intercompany Elimination:** Removes intercompany transactions
3. **Subsidiary Rollup:** Aggregates child subsidiaries to parent

## Syntax

```sql
BUILTIN.CONSOLIDATE(
    tal.amount,           -- Source amount (transaction currency)
    'LEDGER',             -- Amount type
    'DEFAULT',            -- Exchange rate type
    'DEFAULT',            -- Consolidation type
    {target_sub},         -- Target subsidiary ID
    {target_period_id},   -- Period ID for exchange rates
    'DEFAULT'             -- Elimination handling
)
```

## The Critical Period Parameter

**This is the #1 source of bugs!**

```sql
-- WRONG: Uses each transaction's posting period for exchange rate
BUILTIN.CONSOLIDATE(tal.amount, ..., t.postingperiod, ...)

-- CORRECT: Uses report period for all translations
BUILTIN.CONSOLIDATE(tal.amount, ..., {target_period_id}, ...)
```

**Why it matters:**

A January transaction in EUR at 1.10 USD/EUR:
- **Wrong way:** Translates at January rate (1.10) = $110
- **Correct way:** Translates at December rate (1.15) = $115

The Balance Sheet must show ALL amounts at the **same period-end rate** to balance correctly.

## When NOT to Use BUILTIN.CONSOLIDATE

- Single-currency environments (no translation needed)
- Single subsidiary (no consolidation needed)
- Non-OneWorld NetSuite accounts

The backend detects these cases:
```python
if target_sub:
    cons_amount = f"BUILTIN.CONSOLIDATE(tal.amount, ...)"
else:
    cons_amount = "tal.amount"  # Use raw amount
```

---

# Account Types & Sign Conventions

## NetSuite's Internal Storage

| Account Type | Natural Balance | Stored As | Display Multiply |
|--------------|----------------|-----------|------------------|
| **Assets** (Bank, AcctRec, etc.) | Debit | Positive | × 1 |
| **Liabilities** (AcctPay, etc.) | Credit | Negative | × -1 |
| **Equity** | Credit | Negative | × -1 |
| **Income** | Credit | Negative | × -1 |
| **Expenses** (COGS, Expense) | Debit | Positive | × 1 |

## Account Type Constants

**CRITICAL: Exact spelling required!**

```python
# CORRECT (from constants.py)
DEFERRED_EXPENSE = 'DeferExpense'    # NOT 'DeferExpens'
DEFERRED_REVENUE = 'DeferRevenue'    # NOT 'DeferRevenu'
CRED_CARD = 'CredCard'               # NOT 'CreditCard'
```

These typos caused a $60M+ CTA discrepancy. The queries silently exclude accounts with misspelled types.

## Complete Type Reference

### Balance Sheet - Assets
```
Bank              Bank/Cash accounts
AcctRec           Accounts Receivable
OthCurrAsset      Other Current Asset
FixedAsset        Fixed Asset
OthAsset          Other Asset
DeferExpense      Deferred Expense (prepaid)
UnbilledRec       Unbilled Receivable
```

### Balance Sheet - Liabilities
```
AcctPay           Accounts Payable
CredCard          Credit Card
OthCurrLiab       Other Current Liability
LongTermLiab      Long Term Liability
DeferRevenue      Deferred Revenue (unearned)
```

### Balance Sheet - Equity
```
Equity            Common stock, APIC, etc.
RetainedEarnings  Retained Earnings
```

### Income Statement
```
Income            Revenue
OthIncome         Other Income
COGS              Cost of Goods Sold (modern)
Cost of Goods Sold  COGS (legacy - include BOTH!)
Expense           Operating Expense
OthExpense        Other Expense
```

---

# Troubleshooting

## Common Issues

### #N/A in Cells

**Cause:** Network error, timeout, or invalid parameters

**Solution:**
1. Check connection status in taskpane
2. Verify account number exists
3. Verify period format ("Jan 2025")
4. Try "Refresh Selected" on the cell

### #TIMEOUT# in Special Formulas

**Cause:** Backend query took >5 minutes

**Solution:**
1. Ensure tunnel is running
2. Check server logs for errors
3. Try during off-peak hours
4. Contact NetSuite if persistent

### Values Don't Match NetSuite

**Check:**
1. **Subsidiary:** Are you using the correct consolidation level?
2. **Period:** Is it the exact same period end date?
3. **Accounting Book:** Are you querying the same book?
4. **Account Types:** Are any accounts excluded due to type mismatches?

### Slow Performance

**Optimize by:**
1. Use "Refresh Accounts" instead of individual cell refreshes
2. Reduce the number of unique filter combinations
3. Consider using fewer periods per sheet
4. Check tunnel latency (should be <500ms)

## Logs and Debugging

### Browser Console (F12)
Shows client-side logs from functions.js:
```
⚡ CACHE HIT [balance]: 4010:Jan 2025
📥 CACHE MISS [balance]: 4020:Jan 2025
```

### Server Logs
Backend prints detailed query information:
```
📊 Calculating CTA (PLUG METHOD) for Dec 2024
   📜 total_assets SQL: SELECT SUM(...)
   ✓ total_assets: 53,322,353.28
   ✓ total_liabilities: 59,987,254.08
   = CTA (plug): -239,639.06
```

---

# Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.5.36.0 | Dec 2025 | Remove duplicate tooltips |
| 1.5.35.0 | Dec 2025 | TYPE formula batching for drag operations |
| 1.5.34.0 | Dec 2025 | Improved Recalculate Retained Earnings UX |
| 1.5.33.0 | Dec 2025 | Separate Refresh Accounts from RE/NI/CTA |
| 1.5.32.0 | Dec 2025 | Fix account type spellings (DeferExpense, DeferRevenue) |

---

*Document Version: 1.0*
*Last Updated: December 2025*

