# Special Formulas Reference (RETAINEDEARNINGS, NETINCOME, CTA)

This document explains how the three special formulas work, their relationship to BALANCE formulas, and the refresh sequencing logic.

---

## Overview

These three formulas calculate values that **NetSuite computes dynamically at runtime** - there are no account numbers to query directly:

| Formula | What it Calculates | Backend Endpoint |
|---------|-------------------|------------------|
| `XAVI.RETAINEDEARNINGS` | Cumulative P&L from inception through prior fiscal year end | `/retained-earnings` |
| `XAVI.NETINCOME` | Current fiscal year P&L through target period | `/net-income` |
| `XAVI.CTA` | Cumulative Translation Adjustment (multi-currency plug) | `/cta` |

---

## Why These Are "Special"

Unlike `XAVI.BALANCE` which queries a specific account number, these formulas:
1. **Have no account number** - NetSuite calculates them on-the-fly
2. **Require complex queries** - Each makes its own backend API call
3. **Depend on BALANCE data** - Conceptually, they complete the Balance Sheet after BALANCE accounts are loaded

---

## Formula Signatures

### RETAINEDEARNINGS
```javascript
XAVI.RETAINEDEARNINGS(period, [subsidiary], [accountingBook], [classId], [department], [location])
```
- **period**: Required. e.g., "Mar 2025"
- **subsidiary**: Optional. Subsidiary ID or name
- **accountingBook**: Optional. Defaults to Primary Book
- **classId, department, location**: Optional segment filters

**Backend Logic (server.py):**
```
RE = Sum of all P&L from inception through prior fiscal year end
   + Any manual journal entries posted directly to RetainedEarnings accounts
```

### NETINCOME
```javascript
XAVI.NETINCOME(period, [subsidiary], [accountingBook], [classId], [department], [location])
```
Same signature as RETAINEDEARNINGS.

**Backend Logic:**
```
NI = Sum of all P&L transactions from FY start through target period end
```

### CTA (Cumulative Translation Adjustment)
```javascript
XAVI.CTA(period, [subsidiary], [accountingBook])
```
- **Note**: CTA omits segment filters because translation adjustments apply at entity level only.

**Backend Logic (PLUG METHOD for 100% accuracy):**
```
CTA = (Total Assets - Total Liabilities) - Posted Equity - RE - NI
```

---

## Caching Architecture

All three formulas use the same caching infrastructure as BALANCE:

### Cache Storage
```javascript
// In functions.js
const cache = {
    balance: new Map(),  // Also stores special formula results!
    title: new Map(),
    budget: new Map(),
    type: new Map(),
    parent: new Map()
};
```

### Cache Keys
Each formula type uses a distinct prefix:
```javascript
// RETAINEDEARNINGS
cacheKey = `retainedearnings:${period}:${subsidiary}:${accountingBook}:${classId}:${department}:${location}`;

// NETINCOME  
cacheKey = `netincome:${period}:${subsidiary}:${accountingBook}:${classId}:${department}:${location}`;

// CTA (no segment filters)
cacheKey = `cta:${period}:${subsidiary}:${accountingBook}`;
```

### In-Flight Request Deduplication
Because these formulas make expensive API calls, we prevent duplicate concurrent requests:

```javascript
// In functions.js
const inFlightRequests = new Map();

// Example from RETAINEDEARNINGS:
if (inFlightRequests.has(cacheKey)) {
    console.log(`⏳ Waiting for in-flight request [retained earnings]: ${period}`);
    return await inFlightRequests.get(cacheKey);
}

// Store promise BEFORE awaiting
const requestPromise = (async () => {
    try {
        const response = await fetch(`${SERVER_URL}/retained-earnings`, {...});
        // ... process response ...
    } finally {
        inFlightRequests.delete(cacheKey);  // Remove when done
    }
})();

inFlightRequests.set(cacheKey, requestPromise);
return await requestPromise;
```

---

## Refresh Sequencing

### The Problem (Before Fix)
When "Refresh All" ran, ALL formulas (BALANCE + special) would fire simultaneously. This meant special formulas weren't guaranteed to have fresh BALANCE data.

### The Solution (Current Implementation)

**"Refresh All" in taskpane.html now follows this sequence:**

```
STEP 1: Scan sheet for ALL XAVI formulas
        ├── XAVI.BALANCE formulas → stored in cellsToUpdate[]
        └── XAVI.RETAINEDEARNINGS/NETINCOME/CTA → stored in specialFormulas[]

STEP 2: Classify accounts (P&L vs Balance Sheet)

STEP 3: Clear all caches (including inFlightRequests)

STEP 4: Fetch P&L accounts (fast, ~30s/year)

STEP 5: Fetch Balance Sheet accounts (slower, 2-3 min)

STEP 6: Re-evaluate BALANCE formulas (in batches of 100)
        └── Forces Excel to recalculate with fresh cache

*** 500ms PAUSE to ensure cache is fully populated ***

STEP 7: Re-evaluate SPECIAL formulas (in batches of 50)
        └── These run AFTER all BALANCE data is loaded
        └── Forces fresh API calls to /retained-earnings, /net-income, /cta
```

### Key Code: Scanning for Special Formulas

```javascript
// In taskpane.html refreshCurrentSheet()
let cellsToUpdate = [];   // BALANCE formulas
let specialFormulas = []; // RETAINEDEARNINGS, NETINCOME, CTA formulas

for (let row = 0; row < rowCount; row++) {
    for (let col = 0; col < colCount; col++) {
        const formula = formulas[row][col];
        if (formula && typeof formula === 'string') {
            const upperFormula = formula.toUpperCase();
            
            // BALANCE formulas - refresh first
            if (upperFormula.includes('XAVI.BALANCE')) {
                cellsToUpdate.push({ row, col, formula });
                // ... extract account for classification ...
            }
            
            // Special formulas - refresh AFTER BALANCE data is loaded
            if (upperFormula.includes('XAVI.RETAINEDEARNINGS') ||
                upperFormula.includes('XAVI.NETINCOME') ||
                upperFormula.includes('XAVI.CTA')) {
                const formulaType = upperFormula.includes('RETAINEDEARNINGS') ? 'RETAINEDEARNINGS' :
                                   upperFormula.includes('NETINCOME') ? 'NETINCOME' : 'CTA';
                specialFormulas.push({ row, col, formula, type: formulaType });
            }
        }
    }
}
```

### Key Code: Sequential Refresh

```javascript
// After BALANCE formulas are refreshed (Step 6)...

// ============================================
// STEP 7: Refresh special formulas AFTER BALANCE data is loaded
// These depend on BALANCE data, so they must run second
// ============================================
let specialFormulasRefreshed = 0;

if (specialFormulas.length > 0) {
    console.log(`📊 Refreshing ${specialFormulas.length} special formulas...`);
    
    // Small delay to ensure BALANCE cache is fully populated
    await new Promise(resolve => setTimeout(resolve, 500));
    
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        
        const BATCH_SIZE = 50; // Smaller batch for special formulas
        
        for (let batchNum = 0; batchNum < totalBatches; batchNum++) {
            const batch = specialFormulas.slice(batchStart, batchEnd);
            
            for (const { row, col, formula, type } of batch) {
                const cell = usedRange.getCell(row, col);
                cell.formulas = [[formula]];
                specialFormulasRefreshed++;
            }
            
            await context.sync();
            await new Promise(resolve => setTimeout(resolve, 100));
        }
        
        usedRange.calculate();
        await context.sync();
    });
}
```

### Key Code: Cache Clearing

```javascript
// In functions.js window.clearAllCaches()
window.clearAllCaches = function() {
    console.log('🗑️  CLEARING ALL CACHES...');
    
    cache.balance.clear();
    cache.title.clear();
    cache.budget.clear();
    cache.type.clear();
    cache.parent.clear();
    
    // Clear in-flight requests for special formulas (RETAINEDEARNINGS, NETINCOME, CTA)
    // This ensures fresh API calls will be made when formulas re-evaluate
    if (inFlightRequests && inFlightRequests.size > 0) {
        console.log(`  Clearing ${inFlightRequests.size} in-flight requests...`);
        inFlightRequests.clear();
    }
    
    // Reset stats
    cacheStats.hits = 0;
    cacheStats.misses = 0;
    
    console.log('✅ ALL CACHES CLEARED');
    return true;
};
```

---

## Registration

All three functions are registered with Excel's CustomFunctions API:

```javascript
// In functions.js
if (typeof CustomFunctions !== 'undefined') {
    CustomFunctions.associate('RETAINEDEARNINGS', RETAINEDEARNINGS);
    CustomFunctions.associate('NETINCOME', NETINCOME);
    CustomFunctions.associate('CTA', CTA);
    // ... other functions ...
}
```

---

## Drag/Copy Behavior

When users **drag or copy** formulas (without using Refresh All):

1. Excel triggers each formula independently
2. BALANCE formulas check cache first (usually a hit after initial load)
3. Special formulas also check cache first
4. If cache miss, each makes its own API call
5. In-flight deduplication prevents duplicate concurrent calls

**Important:** There is no guaranteed ordering when dragging. However:
- If BALANCE data was previously loaded (cached), special formulas will get fresh data
- If not cached, all formulas fire in parallel (Excel's native behavior)

**Recommendation:** For consistent results, use "Refresh All" to ensure BALANCE data loads before special formulas evaluate.

---

## Summary Status Display

After refresh, the status shows all formula types:
```javascript
if (specialFormulasRefreshed > 0) {
    summaryParts.push(`Special: ${specialFormulasRefreshed}`);
}

// Example: "P&L: 2 year(s) • Balance Sheet: 45 accounts • Special: 3"
```

---

## Files Modified

| File | Changes |
|------|---------|
| `docs/taskpane.html` | Scan for special formulas; refresh them after BALANCE data |
| `docs/functions.js` | Clear inFlightRequests in clearAllCaches() |
| `backend/server.py` | Contains `/retained-earnings`, `/net-income`, `/cta` endpoints |

---

## Version

These changes were implemented in manifest version **1.5.8.0**.

