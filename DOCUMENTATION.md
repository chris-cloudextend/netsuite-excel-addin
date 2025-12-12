# XAVI for NetSuite - Complete Documentation

## Table of Contents

1. [Overview](#overview)
2. [For Finance Users (CPA Perspective)](#for-finance-users-cpa-perspective)
3. [For Engineers (Technical Reference)](#for-engineers-technical-reference)
4. [Why SuiteQL Over ODBC](#why-suiteql-over-odbc)
5. [Pre-Caching & Drag-Drop Optimization](#pre-caching--drag-drop-optimization)
6. [SuiteQL Deep Dive](#suiteql-deep-dive)
7. [BUILTIN.CONSOLIDATE Explained](#builtinconsolidate-explained)
8. [Account Types & Sign Conventions](#account-types--sign-conventions)
   - [Special Account Sign Handling (sspecacct)](#special-account-sign-handling-sspecacct)
9. [AWS Migration Roadmap](#aws-migration-roadmap)
10. [CEFI Integration (CloudExtend Federated Integration)](#cefi-integration-cloudextend-federated-integration)
11. [Troubleshooting](#troubleshooting)

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

## Architecture (Current)

```
┌─────────────────────┐     ┌─────────────────────┐     ┌─────────────────┐
│   Excel Add-in      │────▶│   Cloudflare        │────▶│   Flask Backend │
│   (functions.js)    │     │   Worker + Tunnel   │     │   (server.py)   │
│                     │◀────│                     │◀────│   localhost:5002│
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

---

# Why SuiteQL (REST API) Over ODBC

## Executive Summary

We chose SuiteQL via REST API over NetSuite's ODBC driver (SuiteAnalytics Connect) for three primary reasons:

| Factor | ODBC (SuiteAnalytics Connect) | SuiteQL (REST API) |
|--------|-------------------------------|-------------------|
| **Annual Cost** | $5,000 - $20,000+ | $0 (included) |
| **Client Setup** | Driver installation required | No installation |
| **Firewall** | Database port required | HTTPS only (443) |

> **Important Clarification:** Both ODBC and REST API can execute SuiteQL queries. ODBC via NetSuite2.com data source supports SuiteQL syntax including `BUILTIN.CONSOLIDATE`. The key differences are licensing cost and deployment simplicity, not query capabilities.

## Cost Analysis

### ODBC Driver Costs
NetSuite's ODBC driver (SuiteAnalytics Connect) requires additional licensing purchased separately from your core NetSuite platform:

| Cost Component | Annual Cost |
|----------------|-------------|
| SuiteAnalytics Connect License | **$5,000 - $20,000/year** |
| Additional user seats (if required) | Variable |
| **Total** | **$5,000 - $20,000+/year** |

> *Pricing varies significantly by negotiation. Community reports range from $5K to $20K annually. Some customers have negotiated inclusion in their base contract.*

### SuiteQL REST API Costs
- **License Cost:** $0 - Included with all NetSuite subscriptions
- **API Calls:** Included in standard governance limits
- **Infrastructure:** Only backend server costs (~$50/month for AWS hosting)

### ROI Calculation

For an organization:
```
ODBC Approach:
  SuiteAnalytics Connect License: $5,000-20,000/year
  Driver deployment/maintenance: Time cost

XAVI with SuiteQL REST:
  License: $0
  AWS hosting: ~$50/month = $600/year
  Total: $600/year

Annual Savings: $4,400 - $19,400 (88-97% reduction)
```

## Technical Comparison

### What's the SAME (Both Use SuiteQL)

Both ODBC and REST API can run SuiteQL queries with:
- ✅ `BUILTIN.CONSOLIDATE` for multi-currency consolidation
- ✅ Complex JOINs, GROUP BY, aggregations
- ✅ SQL-92 syntax support
- ✅ Oracle syntax support (REST API has fewer limitations)

**Key limitation for ODBC:** Cannot use WITH clauses (CTEs) via ODBC. REST API supports full SuiteQL syntax.

### What's DIFFERENT

| Aspect | ODBC | REST API |
|--------|------|----------|
| **Licensing** | Additional purchase required | Included |
| **Client Setup** | ODBC driver installation | None |
| **Firewall** | Database port access | HTTPS (443) only |
| **Authentication** | User/Pass, OAuth 2.0, or TBA | OAuth 1.0 (TBA) |
| **Power BI** | Import only (no DirectQuery) | N/A for Excel |
| **WITH Clauses** | Not supported | Supported |

### Why We Chose REST API

1. **Zero Licensing Cost:** No SuiteAnalytics Connect purchase required
2. **No Driver Installation:** Users don't need ODBC drivers on their machines
3. **Simpler Firewall:** Only HTTPS (port 443) needed, no database ports
4. **Full SuiteQL Support:** Including WITH clauses and modern Oracle features
5. **Web-Native:** Works with Excel Add-ins hosted via GitHub Pages

### ODBC Advantages We Traded Away

- **Direct BI Tool Integration:** Power BI, Tableau can connect directly via ODBC
- **Familiar SQL Tools:** Works with any ODBC-compatible application
- **Bulk Data Export:** May be faster for very large one-time exports

For our use case (real-time Excel formulas), REST API's zero-cost and no-installation benefits outweigh ODBC's BI tool compatibility.

## Authentication Clarification

**Both approaches support modern authentication:**

| Method | ODBC | REST API |
|--------|------|----------|
| Username/Password | ✅ | ❌ |
| OAuth 2.0 | ✅ | ❌ |
| Token-Based Auth (TBA) | ✅ | ✅ |
| OAuth 1.0 | ❌ | ✅ |

We use OAuth 1.0 with Token-Based Authentication (TBA) for the REST API.

## References

- NetSuite SuiteQL Documentation: [docs.oracle.com/netsuite](https://docs.oracle.com/en/cloud/saas/netsuite/ns-online-help/chapter_157108952762.html)
- SuiteAnalytics Connect: [NetSuite Help Center](https://docs.oracle.com/en/cloud/saas/netsuite/ns-online-help/section_3aborqgzaqc.html)
- Community Pricing Discussion: [NetSuite Professionals](https://archive.netsuiteprofessionals.com/)

---

# Pre-Caching & Drag-Drop Optimization

## The Challenge

Without optimization, a typical financial report with 100 accounts × 12 months = **1,200 individual API calls**, resulting in hours of waiting.

NetSuite has strict limits:
- **Concurrency:** Max 5 simultaneous API requests
- **Row Limit:** 1,000 rows per query response
- **Rate Limiting:** Too many requests = 429 errors

## Our Solution: Intelligent Pre-Caching

### Build Mode Detection

When users drag formulas across cells, Excel creates formulas nearly simultaneously. We detect this pattern:

```
User drags formula across 12 months:
  → Formula 1: triggers Build Mode (3+ formulas in 500ms)
  → Formula 2-12: queued, show #BUSY placeholder
  → User stops dragging
  → 800ms passes (settle time)
  → Single optimized batch request for ALL data
  → All cells update simultaneously
```

**Detection Criteria:**
```javascript
const BUILD_MODE_THRESHOLD = 3;       // Formulas to trigger
const BUILD_MODE_WINDOW_MS = 500;     // Detection window
const BUILD_MODE_SETTLE_MS = 800;     // Wait after last formula
```

### Pivoted Query Optimization (Periods as Columns)

The **key innovation** is returning multiple periods as columns in a single row, rather than separate rows:

**Traditional Approach (Slow):**
```
12 queries, one per month:
  Query 1: Get Jan 2025 balance for Account 4010
  Query 2: Get Feb 2025 balance for Account 4010
  ... (10 more queries)
```

**Our Approach (Fast):**
```sql
-- Single query returns ALL months as columns
SELECT
  a.acctnumber,
  SUM(CASE WHEN TO_CHAR(ap.startdate,'YYYY-MM')='2025-01' THEN amount ELSE 0 END) AS jan_2025,
  SUM(CASE WHEN TO_CHAR(ap.startdate,'YYYY-MM')='2025-02' THEN amount ELSE 0 END) AS feb_2025,
  SUM(CASE WHEN TO_CHAR(ap.startdate,'YYYY-MM')='2025-03' THEN amount ELSE 0 END) AS mar_2025,
  -- ... all 12 months
FROM TransactionAccountingLine tal
  JOIN Transaction t ON t.id = tal.transaction
  JOIN Account a ON a.id = tal.account
  JOIN AccountingPeriod ap ON ap.id = t.postingperiod
WHERE t.posting = 'T'
  AND a.accttype IN ('Income', 'Expense', ...)
  AND EXTRACT(YEAR FROM ap.startdate) = 2025
GROUP BY a.acctnumber
```

**Result:** One query returns 200 accounts × 12 months = 2,400 data points.

### Full Year Refresh Endpoint

When Build Mode detects 6+ months for the same fiscal year, it triggers `/batch/full_year_refresh`:

```javascript
// Endpoint automatically:
// 1. Fetches ALL P&L accounts for the entire year
// 2. Returns pivoted data (periods as columns)
// 3. Caches everything for instant subsequent lookups

POST /batch/full_year_refresh
{
  "year": 2025,
  "subsidiary": "1",
  "accountingBook": "1"
}

// Response: ~200 accounts × 12 months in one response
```

### Smart Period Expansion

When dragging formulas, we automatically pre-cache adjacent months:

```
User requests: Jan 2025, Feb 2025, Mar 2025
System fetches: Dec 2024, Jan 2025, Feb 2025, Mar 2025, Apr 2025

Why?
- User likely to scroll left/right
- Minimal extra cost (same query complexity)
- Instant response when they do
```

### Three-Tier Caching Architecture

```
┌─────────────────────────────────────────────────────────┐
│  TIER 1: In-Memory Cache (functions.js)                 │
│  - Speed: Microseconds                                  │
│  - Scope: Current session                               │
│  - Size: Unlimited (Map structure)                      │
├─────────────────────────────────────────────────────────┤
│  TIER 2: localStorage Cache                             │
│  - Speed: Milliseconds                                  │
│  - Scope: Persists across taskpane refreshes            │
│  - TTL: 5 minutes                                       │
│  - Shared: Between taskpane and custom functions        │
├─────────────────────────────────────────────────────────┤
│  TIER 3: Backend Cache (server.py)                      │
│  - Speed: Avoids NetSuite roundtrip                     │
│  - TTL: 5 minutes                                       │
│  - Benefit: Shared across all users                     │
└─────────────────────────────────────────────────────────┘
```

### Explicit Zero Caching

**Problem:** Accounts with $0 balance return no rows from NetSuite (no transactions = no data).

**Solution:** After fetching, explicitly cache `$0` for any requested account/period NOT in the response:

```javascript
// NetSuite returns:
{ "4220": { "Jan 2025": 50000, "Feb 2025": 45000 } }
// Note: Mar 2025 missing = $0 balance

// We explicitly cache:
cache.set("4220:Mar 2025", 0);  // Now cached as $0, not a miss
```

This prevents repeated queries for zero-balance accounts.

### Performance Results

| Scenario | Without Optimization | With Optimization |
|----------|---------------------|-------------------|
| Single formula | 2-5 sec | 2-5 sec |
| 20 formulas (batch) | 40-100 sec | 5-10 sec |
| Drag 12 months | 60-180 sec + timeouts | 15-20 sec first, instant after |
| Full sheet refresh | Hours + errors | 30-60 sec |
| Second request (cached) | Same as first | **Instant** |

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

## Segment Filters (Class, Department, Location)

**CRITICAL:** Class, Department, and Location fields are on `TransactionLine`, NOT `TransactionAccountingLine`!

```sql
-- WRONG (will fail with "Field 'class' not found"):
WHERE tal.class = 85

-- CORRECT (join to TransactionLine):
FROM TransactionAccountingLine tal
  JOIN Transaction t ON t.id = tal.transaction
  JOIN TransactionLine tl ON t.id = tl.transaction AND tal.transactionline = tl.id
WHERE tl.class = 85
```

| Field | Correct Table | Alias |
|-------|---------------|-------|
| `class` | TransactionLine | `tl.class` |
| `department` | TransactionLine | `tl.department` |
| `location` | TransactionLine | `tl.location` |
| `subsidiary` | Transaction | `t.subsidiary` |
| `account` | TransactionAccountingLine | `tal.account` |
| `amount` | TransactionAccountingLine | `tal.amount` |

When filtering by class/department/location, always add the TransactionLine join.

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

## Special Account Sign Handling (sspecacct)

### The Problem

Certain NetSuite accounts require **additional sign inversion** when displaying amounts on financial statements. Specifically, accounts with a `sspecacct` (Special Account) field value that starts with "Matching" need their calculated amounts inverted to match NetSuite's native financial report presentation.

### Background

NetSuite uses "Matching" special accounts as **contra/offset entries for currency revaluation**. For example:

| Account | sspecacct | Display Behavior |
|---------|-----------|------------------|
| 89100 - Unrealized Gain/Loss | `UnrERV` | Normal sign logic |
| 89201 - Unrealized Matching Gain/Loss | `MatchingUnrERV` | **Inverted sign** |

Both accounts have the same `accttype` (e.g., `OthExpense`), so you **cannot rely on account type alone** for sign logic.

### The Solution

Apply an additional sign inversion for any account where `sspecacct LIKE 'Matching%'`.

**SQL Pattern (applied to all P&L queries):**
```sql
SUM(amount) 
    * CASE WHEN a.accttype IN ('Income', 'OthIncome') THEN -1 ELSE 1 END
    * CASE WHEN a.sspecacct LIKE 'Matching%' THEN -1 ELSE 1 END
```

**How it works:**

| Account Type | Is Matching? | First Multiplier | Second Multiplier | Net Effect |
|--------------|--------------|------------------|-------------------|------------|
| Income/Revenue | No | -1 | 1 | -1 (flip to positive) |
| Expense | No | 1 | 1 | 1 (keep as is) |
| Income/Revenue | Yes | -1 | -1 | 1 (double flip) |
| **OthExpense (Matching)** | **Yes** | 1 | **-1** | **-1 (flip for display)** |

### Why This Works

- The `sspecacct` field is a NetSuite system field that identifies special-purpose accounts
- "Matching" is NetSuite's naming convention for contra accounts used in currency revaluation eliminations
- This approach is **universal** and will automatically handle any future "Matching" special accounts NetSuite may add
- **No hardcoded account numbers required**

### Testing

Verify against NetSuite's native Income Statement for any account with `sspecacct LIKE 'Matching%'`:

```sql
SELECT id, acctnumber, fullname, accttype, sspecacct 
FROM account 
WHERE sspecacct LIKE 'Matching%'
```

Current Matching accounts in this instance:
- **89201** - Unrealized Matching Gain/Loss (`MatchingUnrERV`)

---

# AWS Migration Roadmap

## Current State (Local + Cloudflare Tunnel)

```
┌─────────────┐     ┌─────────────┐     ┌─────────────────────┐
│ Excel Add-in │────▶│ Cloudflare  │────▶│ Cloudflare Tunnel   │
│             │     │ Worker      │     │ (Quick Tunnel)      │
└─────────────┘     │ (CORS Proxy)│     │                     │
                    └─────────────┘     └──────────┬──────────┘
                                                   │
                                                   ▼
                                        ┌─────────────────────┐
                                        │ Local Flask Server  │
                                        │ localhost:5002      │
                                        │ (Developer Machine) │
                                        └─────────────────────┘
```

**Limitations:**
- Requires developer machine running 24/7
- Tunnel URL changes on restart (must update Worker)
- No redundancy or scalability
- Single point of failure

## Target State (AWS)

```
┌─────────────┐     ┌─────────────┐     ┌─────────────────────┐
│ Excel Add-in │────▶│ AWS API     │────▶│ AWS Lambda          │
│             │     │ Gateway     │     │ (or ECS/Fargate)    │
└─────────────┘     │ (HTTPS)     │     │                     │
                    └─────────────┘     └──────────┬──────────┘
                                                   │
                                                   ▼
                                        ┌─────────────────────┐
                                        │ AWS Secrets Manager │
                                        │ (NetSuite Creds)    │
                                        └─────────────────────┘
```

## Migration Steps

### Phase 1: Backend Containerization
1. **Dockerize Flask app**
   ```dockerfile
   FROM python:3.11-slim
   WORKDIR /app
   COPY requirements.txt .
   RUN pip install -r requirements.txt
   COPY . .
   CMD ["gunicorn", "-b", "0.0.0.0:5002", "server:app"]
   ```

2. **Move credentials to environment variables**
   ```python
   # Current: File-based
   config = json.load(open('netsuite_config.json'))
   
   # AWS: Environment variables / Secrets Manager
   config = {
       'account_id': os.environ['NETSUITE_ACCOUNT_ID'],
       'consumer_key': os.environ['NETSUITE_CONSUMER_KEY'],
       # ...
   }
   ```

### Phase 2: AWS Deployment
| Option | Pros | Cons | Cost |
|--------|------|------|------|
| **Lambda + API Gateway** | Serverless, auto-scale | Cold starts, 15min timeout | ~$5-20/month |
| **ECS Fargate** | No cold starts, long-running | Always-on cost | ~$30-50/month |
| **EC2** | Full control | Must manage server | ~$20-40/month |

**Recommendation:** Start with **ECS Fargate** for production reliability.

### Phase 3: Infrastructure Changes

**What Changes:**
| Component | Current | AWS |
|-----------|---------|-----|
| Backend URL | Cloudflare Worker → Tunnel | API Gateway HTTPS endpoint |
| Credentials | Local JSON file | AWS Secrets Manager |
| CORS | Cloudflare Worker | API Gateway CORS config |
| SSL/TLS | Cloudflare | AWS Certificate Manager |
| Logging | Console | CloudWatch Logs |

**What Stays the Same:**
- Excel Add-in code (just update `SERVER_URL`)
- SuiteQL queries
- Caching logic
- All custom functions

### Phase 4: Remove Cloudflare Dependency

```javascript
// functions.js - Update SERVER_URL
// Current:
const SERVER_URL = 'https://netsuite-proxy.chris-corcoran.workers.dev';

// AWS:
const SERVER_URL = 'https://api.xavi.cloudextend.io';
```

The Cloudflare Worker becomes unnecessary - API Gateway handles CORS natively.

### Cost Comparison

| Item | Current (Cloudflare) | AWS (Fargate) |
|------|---------------------|---------------|
| Compute | $0 (local machine) | ~$30/month |
| Tunnel | $0 (free tier) | N/A |
| API Gateway | N/A | ~$5/month |
| Secrets Manager | N/A | ~$1/month |
| **Total** | **$0** (but unreliable) | **~$36/month** |

---

# CEFI Integration (CloudExtend Federated Integration)

## Overview

CEFI (CloudExtend Federated Integration) is our authentication and tenant management system. It will replace the current static credential model.

## Current Authentication Model

```
┌─────────────┐     ┌─────────────┐     ┌─────────────────────┐
│ Excel Add-in │────▶│ Backend     │────▶│ Single NetSuite     │
│             │     │ (static creds)    │ Account             │
└─────────────┘     └─────────────┘     └─────────────────────┘
```

**Limitations:**
- One NetSuite account per deployment
- Credentials hardcoded in backend
- No user-level permissions
- No multi-tenant support

## Target Model with CEFI

```
┌─────────────┐     ┌─────────────┐     ┌─────────────────────┐
│ Excel Add-in │────▶│ CEFI Auth   │────▶│ Token Service       │
│ (User Login) │     │ Portal      │     │                     │
└─────────────┘     └─────────────┘     └──────────┬──────────┘
                                                   │
                                                   ▼
                                        ┌─────────────────────┐
                                        │ Multi-Tenant        │
                                        │ Credential Store    │
                                        │                     │
                                        │ Customer A → NS Acct│
                                        │ Customer B → NS Acct│
                                        │ Customer C → NS Acct│
                                        └─────────────────────┘
```

## CEFI Components

### 1. Authentication Flow
```
1. User opens Excel Add-in
2. Add-in checks for valid CEFI token
3. If no token: Redirect to CEFI login portal
4. User logs in with SSO (Google, Microsoft, SAML)
5. CEFI returns JWT token
6. Add-in stores token, sends with all API requests
7. Backend validates token, retrieves tenant-specific NetSuite credentials
```

### 2. Token Structure
```json
{
  "sub": "user@company.com",
  "tenant_id": "customer-abc-123",
  "netsuite_account": "589861",
  "roles": ["viewer", "editor"],
  "exp": 1735689600,
  "iss": "cefi.cloudextend.io"
}
```

### 3. Backend Changes

```python
# Current: Static credentials
def get_netsuite_client():
    config = json.load(open('netsuite_config.json'))
    return NetSuiteClient(config)

# CEFI: Tenant-specific credentials
def get_netsuite_client(cefi_token):
    # Validate token
    payload = jwt.decode(cefi_token, CEFI_PUBLIC_KEY)
    tenant_id = payload['tenant_id']
    
    # Fetch tenant credentials from secure store
    credentials = secrets_manager.get_secret(f'netsuite/{tenant_id}')
    
    return NetSuiteClient(credentials)
```

### 4. Frontend Changes

```javascript
// functions.js - Add CEFI token to requests
async function fetchWithAuth(url, options = {}) {
    const cefiToken = await getCEFIToken();
    
    if (!cefiToken) {
        // Redirect to login
        window.location.href = 'https://cefi.cloudextend.io/login?redirect=' + 
            encodeURIComponent(window.location.href);
        return;
    }
    
    return fetch(url, {
        ...options,
        headers: {
            ...options.headers,
            'Authorization': `Bearer ${cefiToken}`
        }
    });
}
```

## Benefits of CEFI

| Feature | Current | With CEFI |
|---------|---------|-----------|
| Multi-tenant | ❌ Single account | ✅ Unlimited customers |
| User management | ❌ None | ✅ Full RBAC |
| SSO | ❌ None | ✅ Google, Microsoft, SAML |
| Audit logging | ❌ Basic | ✅ Per-user activity |
| Credential rotation | ❌ Manual | ✅ Automated |
| Billing integration | ❌ None | ✅ Usage tracking |

## Implementation Timeline

| Phase | Tasks | Duration |
|-------|-------|----------|
| **Phase 1** | CEFI portal setup, JWT infrastructure | 2-3 weeks |
| **Phase 2** | Backend token validation, secrets integration | 1-2 weeks |
| **Phase 3** | Frontend login flow, token management | 1-2 weeks |
| **Phase 4** | Multi-tenant credential store | 1 week |
| **Phase 5** | Testing, migration | 1-2 weeks |

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
| 1.5.37.0 | Dec 2025 | Fix tooltips (lighter style), remove unnecessary error messages |
| 1.5.36.0 | Dec 2025 | Remove duplicate tooltips |
| 1.5.35.0 | Dec 2025 | TYPE formula batching for drag operations |
| 1.5.34.0 | Dec 2025 | Fix class/dept/location filters (use TransactionLine not TransactionAccountingLine) |
| 1.5.33.0 | Dec 2025 | Separate Refresh Accounts from RE/NI/CTA |
| 1.5.32.0 | Dec 2025 | Fix account type spellings (DeferExpense, DeferRevenue) |
| 1.5.31.0 | Dec 2025 | Fix Department/Class/Location lookups (direct table queries) |

---

*Document Version: 2.1*
*Last Updated: December 2025*
