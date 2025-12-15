# XAVI Developer Checklist

**Purpose:** This checklist ensures all integration points are updated when adding or modifying formulas/features.

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

### 2. Frontend - Excel Registration
| File | Location | What to Update |
|------|----------|----------------|
| `docs/functions.json` | Function entry | `id`, `name`, `description` |
| `docs/functions.json` | Parameters array | Names, descriptions, types, optional flags |
| `docs/functions.json` | Options | `stream`, `cancelable`, `volatile` settings |

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

### 5. Manifest & Versioning
| File | Location | What to Update |
|------|----------|----------------|
| `excel-addin/manifest-claude.xml` | `<Version>` tag | Main version (line ~22) |
| `excel-addin/manifest-claude.xml` | ALL `?v=X.X.X.X` URLs | Cache-busting parameters |
| `docs/taskpane.html` | Footer version | Hardcoded display (~line 2292) |

### 6. Documentation
| File | What to Update |
|------|----------------|
| `README.md` | Version number, feature list |
| `docs/README.md` | Version number |
| `USER_GUIDE.md` | Usage examples, version |
| `QA_TEST_PLAN.md` | Test cases for new feature |
| `PROJECT_SUMMARY.md` | Version number |

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

---

*Last updated: December 15, 2025*

