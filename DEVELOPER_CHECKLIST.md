# XAVI Developer Checklist

**Purpose:** This checklist ensures all integration points are updated when adding or modifying formulas/features.

---

## ⚠️ CRITICAL: Universal NetSuite Compatibility

**All code MUST work across ALL NetSuite accounts.** Never make assumptions about:

| ❌ DON'T Assume | ✅ DO Instead |
|-----------------|---------------|
| Account number prefixes indicate type (1xxx=BS, 4xxx=P&L) | Query NetSuite for actual `accttype` field |
| Specific account numbers exist | Use wildcards or let users specify |
| Subsidiary structure (OneWorld vs non-OneWorld) | Use `BUILTIN.CONSOLIDATE` which works universally |
| Chart of accounts numbering scheme | Query account metadata from NetSuite |
| Fiscal year starts in January | Query `get_fiscal_year_for_period()` |
| Currency is USD | Use subsidiary's currency or consolidate |
| Specific period names exist | Validate periods exist before querying |

### Examples of Universal vs Non-Universal Code:

```python
# ❌ BAD - Assumes account prefix indicates type
if account.startswith('1') or account.startswith('2'):
    is_bs_account = True

# ✅ GOOD - Queries actual account type from NetSuite
type_result = query_netsuite("SELECT accttype FROM Account WHERE acctnumber = '...'")
is_bs_account = is_balance_sheet_account(type_result['accttype'])
```

```python
# ❌ BAD - Assumes fiscal year starts January 1
fy_start = f"Jan {year}"

# ✅ GOOD - Queries actual fiscal year boundaries
fy_info = get_fiscal_year_for_period(period_name, accounting_book)
fy_start = fy_info['fy_start']
```

---

## 🔧 Adding or Modifying a Formula

### 1. Frontend - Function Implementation
| File | Location | What to Update |
|------|----------|----------------|
| `docs/functions.js` | Function definition | Core logic, parameters, return values |
| `docs/functions.js` | `convertToMonthYear()` | If new date handling needed |
| `docs/functions.js` | Cache logic | Cache key format, localStorage keys |
| `docs/functions.js` | `__CLEARCACHE__` handler | If formula has special cache needs |
| `docs/functions.js` | Build mode | If formula should batch with others |

### 2. Frontend - Excel Registration ⚠️ CRITICAL
| File | Location | What to Update |
|------|----------|----------------|
| `docs/functions.json` | Function entry | `id`, `name`, `description` |
| `docs/functions.json` | Parameters array | Names, descriptions, types, optional flags |
| `docs/functions.json` | Options | `stream`, `cancelable`, `volatile` settings |
| `docs/functions.js` | **`CustomFunctions.associate()`** | **MUST add function binding (~line 4986)** |

> ⚠️ **CRITICAL #1**: Missing `CustomFunctions.associate('FUNCTIONNAME', FUNCTIONNAME)` will cause  
> the entire add-in to fail with "We can't start this add-in because it isn't set up properly."

> ⚠️ **CRITICAL #2**: **Optional parameters MUST come AFTER required parameters!**  
> Excel will silently fail to load the add-in if an `optional: true` parameter appears before a required one.  
> If you need a "skippable" parameter in the middle, mark ALL subsequent parameters as `optional: true`  
> and validate required values in JavaScript instead.

### 3. Frontend - Taskpane Integration
| File | Location | What to Update |
|------|----------|----------------|
| `docs/taskpane.html` | `refreshSelected()` | Add formula type detection (~line 11283) |
| `docs/taskpane.html` | `refreshCurrentSheet()` | If formula needs special handling |
| `docs/taskpane.html` | `recalculateSpecialFormulas()` | For RE/NI/CTA type formulas |
| `docs/taskpane.html` | `clearCache()` | If new localStorage keys used |
| `docs/taskpane.html` | UI buttons | If new action buttons needed |
| `docs/taskpane.html` | Tooltips/help text | User-facing descriptions |
| `docs/taskpane.html` | Error messages | Toast notifications, status messages |

### 4. Backend - Server Implementation
| File | Location | What to Update |
|------|----------|----------------|
| `backend/server.py` | New `@app.route` | API endpoint for the formula |
| `backend/server.py` | Query logic | SuiteQL query construction |
| `backend/server.py` | Default handling | `default_subsidiary_id` usage |
| `backend/server.py` | Consolidation | `get_subsidiaries_in_hierarchy()` if needed |
| `backend/server.py` | Name-to-ID conversion | `convert_name_to_id()` for filters |
| `backend/server.py` | Response format | JSON structure returned |

### 5. Manifest & Versioning ⚠️ CRITICAL SHARED RUNTIME
| File | Location | What to Update |
|------|----------|----------------|
| `excel-addin/manifest-claude.xml` | `<Version>` tag | Main version (line ~22) |
| `excel-addin/manifest-claude.xml` | ALL `?v=X.X.X.X` URLs | Cache-busting parameters |
| `docs/taskpane.html` | Footer version | Hardcoded display (~line 2292) |

> ⚠️ **CRITICAL SHARED RUNTIME CONFIGURATION:**  
> - `<Runtime resid>` MUST point to `Taskpane.Url` (NOT a separate functions.html)  
> - CustomFunctions `<Page>` MUST also use `Taskpane.Url`  
> - `taskpane.html` MUST include `<script src="functions.js">` tag  
> - All components (taskpane, functions, commands) use the SAME HTML file  
> 
> **Violating this causes "We can't start this add-in because it isn't set up properly" error.**

### 6. Documentation
| File | What to Update |
|------|----------------|
| `README.md` | Version number, feature list |
| `docs/README.md` | Version number |
| `docs/USER_GUIDE_TYPEBALANCE.md` | TYPEBALANCE usage, account types |
| `docs/SPECIAL_ACCOUNT_TYPES.md` | Special account type reference |
| `QA_TEST_PLAN.md` | Test cases for new feature |
| `PROJECT_SUMMARY.md` | Version number |

---

## 🏷️ Account Type vs Special Account Type (TYPEBALANCE)

When building formulas that filter by account classification, understand the difference:

### Account Type (`accttype` field)
Standard financial categories for financial statements:
- `Bank`, `AcctRec`, `OthCurrAsset`, `FixedAsset`, `OthAsset`
- `AcctPay`, `CredCard`, `OthCurrLiab`, `LongTermLiab`
- `Equity`, `RetainedEarnings`
- `Income`, `COGS`, `Expense`, `OthIncome`, `OthExpense`

**Use for:** Financial reporting, summarizing by category, Balance Sheet/P&L totals

### Special Account Type (`sspecacct` field)
System-assigned tags for accounts with special internal roles:
- `AcctRec`, `AcctPay` - AR/AP control accounts
- `InvtAsset` - Inventory asset account
- `UndepFunds` - Undeposited funds clearing
- `DeferRevenue`, `DeferExpense` - Deferred items
- `RetEarnings`, `CumulTransAdj` - Equity system accounts
- `RealizedERV`, `UnrERV` - FX gain/loss accounts

**Use for:** Identifying system control accounts, troubleshooting posting behavior, understanding transaction flows

### Backend Implementation Pattern

```python
# Determine which field to filter on
account_field = 'a.sspecacct' if use_special_account else 'a.accttype'

# Query uses the dynamic field
query = f"... WHERE {account_field} = '{account_type}' ..."
```

### Frontend Implementation Pattern

```javascript
// Parameter at position 9 controls which field
const useSpecial = useSpecialAccount === 1 || useSpecialAccount === '1';

// Different validation sets for each mode
if (useSpecial) {
    // Validate against BS_SPECIAL_TYPES and PL_SPECIAL_TYPES
} else {
    // Validate against BS_TYPES and PL_TYPES
}
```

### Documentation
- User guide: `docs/USER_GUIDE_TYPEBALANCE.md`
- Reference: `docs/SPECIAL_ACCOUNT_TYPES.md`

---

## 🎯 Special Formula Checklist (NETINCOME, RETAINEDEARNINGS, CTA)

These formulas have additional integration points:

- [ ] `functions.js` - Uses `acquireSpecialFormulaLock()` / `releaseSpecialFormulaLock()`
- [ ] `functions.js` - Uses `broadcastToast()` for progress notifications
- [ ] `taskpane.html` - Listed in `recalculateSpecialFormulas()` 
- [ ] `taskpane.html` - Detected in `refreshSelected()` special formulas array
- [ ] `server.py` - Uses `get_fiscal_year_for_period()` for date boundaries
- [ ] `server.py` - Uses `BUILTIN.CONSOLIDATE()` for multi-currency

---

## 📋 Pre-Commit Checklist

Before committing changes:

- [ ] All version numbers synchronized (manifest, taskpane footer)
- [ ] Console logging added for debugging
- [ ] Error handling returns appropriate codes (#ERROR#, #TIMEOUT#, #SYNTAX#)
- [ ] Cache keys are unique and descriptive
- [ ] Backwards compatibility maintained (or breaking changes documented)
- [ ] Git commit message includes version number

---

## 🔄 When to Update This Checklist

Update this file when:
1. New integration points are discovered
2. Architecture changes (new files, restructured code)
3. New formula types are added
4. New caching mechanisms are introduced

---

## 📝 Version History

| Date | Version | Changes |
|------|---------|---------|
| 2025-12-15 | 3.0.5.81 | Initial checklist created |
| 2025-12-15 | 3.0.5.90 | Added "Universal NetSuite Compatibility" section |
| 2025-12-15 | 3.0.5.96 | Added CustomFunctions.associate() warning |
| 2025-12-15 | 3.0.5.98 | Added CRITICAL shared runtime configuration notes |
| 2025-12-15 | 3.0.5.107 | Added Account Type vs Special Account Type section |

---

*Last updated: December 15, 2025*

