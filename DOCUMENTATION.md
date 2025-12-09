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

# Why SuiteQL Over ODBC

## Executive Summary

We chose SuiteQL (REST API) over NetSuite's ODBC driver (SuiteAnalytics Connect) for three primary reasons:

| Factor | ODBC | SuiteQL |
|--------|------|---------|
| **Annual Cost** | $5,000 - $10,000+ | $0 (included) |
| **Performance** | Slower for complex queries | Optimized for aggregations |
| **Consolidation** | Manual currency translation | `BUILTIN.CONSOLIDATE` built-in |

## Cost Analysis

### ODBC Driver Costs
NetSuite's ODBC driver (SuiteAnalytics Connect) requires additional licensing:

| Cost Component | Annual Cost |
|----------------|-------------|
| SuiteAnalytics Connect License | **$3,000 - $6,000/year** |
| Additional user seats (if required) | $500 - $1,000/seat |
| Third-party connector tools | $1,000 - $3,000/year |
| **Total** | **$5,000 - $10,000+/year** |

> *Source: User reports from NetSuite Professionals community (2024) indicate ODBC licenses costing approximately $500/month ($6,000/year).*

### SuiteQL Costs
- **License Cost:** $0 - Included with all NetSuite subscriptions
- **API Calls:** Included in standard governance limits
- **Infrastructure:** Only backend server costs

### ROI Calculation

For an organization with 10 Excel users:
```
ODBC Approach:
  License: $6,000/year
  Connector: $2,000/year
  Total: $8,000/year

XAVI with SuiteQL:
  License: $0
  AWS hosting: ~$50/month = $600/year
  Total: $600/year

Annual Savings: $7,400 (92% reduction)
```

## Performance Comparison

### ODBC Limitations

1. **Connection Overhead:** Each query establishes a new database connection
2. **Query Complexity:** Limited JOIN support, no native aggregation functions
3. **Row Limits:** Must paginate manually for large result sets
4. **No Consolidation:** Currency translation must be done client-side

### SuiteQL Advantages

1. **Rich SQL Support:** Complex JOINs, GROUP BY, aggregations
2. **BUILTIN Functions:** `BUILTIN.CONSOLIDATE` handles multi-currency automatically
3. **Optimized for Analytics:** Designed for reporting workloads
4. **Batch Operations:** Multiple queries can share authentication overhead

### Benchmark Results (Typical)

| Query Type | ODBC | SuiteQL |
|------------|------|---------|
| Single account balance | 2-4 sec | 1-2 sec |
| Full year P&L (200 accounts) | 45-90 sec | 15-30 sec |
| Multi-subsidiary consolidation | N/A (manual) | 20-40 sec |

## Security Advantages

| Aspect | ODBC | SuiteQL |
|--------|------|---------|
| Authentication | Username/Password | OAuth 1.0 (HMAC-SHA256) |
| Permissions | Database-level access | NetSuite role-based |
| Audit Trail | Limited | Full NetSuite logging |
| Firewall | Requires DB port open | Standard HTTPS (443) |

## Technical Limitations Avoided

### ODBC Pain Points We Avoid:

1. **Driver Installation:** Users don't need to install ODBC drivers
2. **Connection Strings:** No complex DSN configuration
3. **Firewall Rules:** No special ports to open (HTTPS only)
4. **Version Compatibility:** No driver version conflicts

### Why Not Both?

Some tools use ODBC for bulk data and API for real-time. We use SuiteQL exclusively because:
- **Consistency:** Same query language everywhere
- **Simplicity:** One integration point to maintain
- **Cost:** Zero additional licensing

## References

- NetSuite SuiteQL Documentation: [docs.oracle.com/netsuite](https://docs.oracle.com/en/cloud/saas/netsuite/ns-online-help/chapter_157108952762.html)
- ODBC Cost Reports: [NetSuite Professionals Archive](https://archive.netsuiteprofessionals.com/t/439676/what-is-the-advantage-of-using-suiteql-in-suitescript-n-quer)
- Cost Comparison Analysis: [Coefficient.io](https://coefficient.io/use-cases/cost-comparison-netsuite-excel-integration-tools-trials)

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
| 1.5.36.0 | Dec 2025 | Remove duplicate tooltips |
| 1.5.35.0 | Dec 2025 | TYPE formula batching for drag operations |
| 1.5.34.0 | Dec 2025 | Improved Recalculate Retained Earnings UX |
| 1.5.33.0 | Dec 2025 | Separate Refresh Accounts from RE/NI/CTA |
| 1.5.32.0 | Dec 2025 | Fix account type spellings (DeferExpense, DeferRevenue) |

---

*Document Version: 2.0*
*Last Updated: December 2025*
