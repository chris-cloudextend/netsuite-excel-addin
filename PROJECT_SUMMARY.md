# NetSuite Excel Add-in - Technical Summary

> **Note:** For complete documentation including CPA perspective, see [DOCUMENTATION.md](DOCUMENTATION.md)

## Project Overview

**XAVI for NetSuite** is an Excel Add-in that provides custom functions to retrieve financial data from NetSuite via SuiteQL queries.

### Key Capabilities
- Multi-Book Accounting support
- Multi-currency consolidation via `BUILTIN.CONSOLIDATE`
- Smart caching with intelligent invalidation
- Build Mode detection for drag-drop batching
- Optimized queries for P&L vs Balance Sheet accounts

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
└─────────────────────┘                              └─────────────────┘
```

---

## Custom Functions Reference

### Core Functions

| Function | Purpose | Example |
|----------|---------|---------|
| `XAVI.BALANCE` | GL account balance | `=XAVI.BALANCE("4010", "Jan 2025", "Jan 2025")` |
| `XAVI.BUDGET` | Budget amount | `=XAVI.BUDGET("4010", "Jan 2025", "Dec 2025")` |
| `XAVI.NAME` | Account name | `=XAVI.NAME("4010")` |
| `XAVI.TYPE` | Account type | `=XAVI.TYPE("4010")` |
| `XAVI.PARENT` | Parent account | `=XAVI.PARENT("4010-1")` |

### Special Formulas (Calculated Values)

| Function | What it Calculates |
|----------|-------------------|
| `XAVI.RETAINEDEARNINGS` | Cumulative P&L through prior fiscal year |
| `XAVI.NETINCOME` | Current fiscal year P&L through target period |
| `XAVI.CTA` | Cumulative Translation Adjustment (plug method) |

---

## Key Technical Decisions

### 1. BUILTIN.CONSOLIDATE Period Parameter
```sql
-- CORRECT: Use report period for all translations
BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT', 
                    {target_sub}, {target_period_id}, 'DEFAULT')
```
All Balance Sheet amounts must be translated at the **same period-end rate**.

### 2. Account Type Spellings (CRITICAL)
```python
DEFERRED_EXPENSE = 'DeferExpense'    # NOT 'DeferExpens'
DEFERRED_REVENUE = 'DeferRevenue'    # NOT 'DeferRevenu'
CRED_CARD = 'CredCard'               # NOT 'CreditCard'
```
Typos silently exclude accounts from queries.

### 3. P&L vs Balance Sheet
- **P&L:** Filter by `ap.periodname IN (periods)` - activity only
- **BS:** Filter by `ap.enddate <= period_end` - cumulative balance

### 4. Sign Conventions
- Assets: × 1 (stored positive)
- Liabilities/Equity/Income: × -1 (stored negative, flip for display)
- Expenses: × 1 (stored positive)

### 5. CTA Plug Method
```
CTA = (Total Assets - Total Liabilities) - Posted Equity - RE - NI
```
This is the only way to get 100% accuracy because NetSuite calculates additional translation adjustments at runtime.

---

## Security

| Area | Status |
|------|--------|
| Credentials in git | ✅ `.gitignore` blocks config files |
| Transport security | ✅ HTTPS via Cloudflare |
| OAuth authentication | ✅ HMAC-SHA256 signatures |
| SQL injection | ✅ `escape_sql()` function |

---

## Backend Endpoints

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/batch/full_year_refresh` | POST | All P&L for fiscal year |
| `/batch/bs_periods` | POST | All BS for specific periods |
| `/batch/balance` | POST | Specific accounts/periods |
| `/retained-earnings` | POST | Calculate RE |
| `/net-income` | POST | Calculate NI |
| `/cta` | POST | Calculate CTA |
| `/lookups/all` | GET | Filter dropdowns |

---

## File Structure

```
├── backend/
│   ├── server.py              # Flask API server
│   ├── constants.py           # Account type constants
│   └── netsuite_config.json   # Credentials (gitignored)
├── docs/
│   ├── functions.js           # Custom Excel functions + caching
│   ├── functions.json         # Function metadata for Excel
│   └── taskpane.html          # Taskpane UI + JavaScript
├── excel-addin/
│   └── manifest-claude.xml    # Excel add-in manifest
└── CLOUDFLARE-WORKER-CODE.js  # Proxy for production
```

---

## Deployment Checklist

- [ ] Update Cloudflare Worker with new tunnel URL
- [ ] Verify GitHub Pages deployment (~1 minute after push)
- [ ] Bump manifest version for cache-busting
- [ ] Test multi-currency consolidation accuracy
- [ ] Verify special formulas (RE, NI, CTA)

---

*Current Version: 1.5.37.0*
*Last Updated: December 2025*
