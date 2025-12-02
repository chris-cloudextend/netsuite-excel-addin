# 🚀 PERFORMANCE OPTIMIZATION PLAN

**Date:** December 2, 2025  
**Backup:** v1.0.0.87-working (tag) / backup-v1.0.0.87-before-optimization (branch)

---

## 🎯 GOALS

1. **Stop Recalculation on Open** - No @ symbols when opening sheet
2. **Improve Performance** - Faster data retrieval and caching

---

## 🔍 ROOT CAUSE ANALYSIS

### Issue 1: Recalculation on Open

**Problem:**
- Functions show @ symbol (recalculating) every time sheet opens
- Even though `volatile: false` is set

**Root Cause:**
- `streaming: true` functions ALWAYS recalculate in Excel
- Streaming functions use `invocation.setResult()` / `invocation.close()`
- Excel treats these as "live" data connections
- `volatile: false` is IGNORED for streaming functions

**Solution:**
- Convert GLABAL and GLABUD to NON-STREAMING async functions
- Use aggressive client-side caching
- Return cached values instantly on open
- Only fetch from server when explicitly refreshed

### Issue 2: Performance

**Problem:**
- Batching is working but slower than expected
- BUILTIN.CONSOLIDATE called per-account is expensive

**Root Cause:**
- Current approach: Multiple batches of 30 accounts each
- Each batch calls BUILTIN.CONSOLIDATE per account/period
- Batching helps but still multiple queries

**Solution:**
- "Load Everything Once" approach
- Single large query fetches ALL needed data
- BUILTIN.CONSOLIDATE in efficient subquery
- Client caches EVERYTHING
- Subsequent opens = instant (from cache)

---

## 📋 IMPLEMENTATION STRATEGY

### Phase 1: Fix Recalculation (CRITICAL)

**Change GLABAL and GLABUD to NON-STREAMING:**

```javascript
// OLD (Streaming):
function GLABAL(account, fromPeriod, toPeriod, ...) {
    // Uses invocation.setResult() and invocation.close()
    // streaming: true in functions.json
}

// NEW (Non-Streaming):
async function GLABAL(account, fromPeriod, toPeriod, ...) {
    // Returns Promise<number>
    // streaming: false in functions.json
    // Uses aggressive cache
    return cachedValue || await fetchAndCache();
}
```

**Benefits:**
- ✅ No recalculation on open
- ✅ Instant results from cache
- ✅ True non-volatile behavior
- ✅ User explicitly refreshes when needed

**Trade-off:**
- Initial load still needs to fetch data
- But subsequent opens = instant

### Phase 2: Optimize Backend Query

**Current Approach:**
```sql
-- Batch 1: Accounts 1-30
SELECT account, period, BUILTIN.CONSOLIDATE(amount, ...) 
FROM ... WHERE account IN (acc1, acc2, ..., acc30)

-- Batch 2: Accounts 31-60
SELECT account, period, BUILTIN.CONSOLIDATE(amount, ...) 
FROM ... WHERE account IN (acc31, acc32, ..., acc60)
```

**New Approach:**
```sql
-- Single query: ALL accounts for the sheet
WITH consolidated_amounts AS (
    SELECT 
        account,
        period,
        BUILTIN.CONSOLIDATE(amount, ...) as cons_amount
    FROM transactionaccountingline
    -- ... joins ...
    WHERE account IN (ALL_ACCOUNTS_FROM_SHEET)
)
SELECT account, period, SUM(cons_amount)
FROM consolidated_amounts
GROUP BY account, period
```

**Benefits:**
- ✅ One query instead of many
- ✅ CONSOLIDATE called once per account/period
- ✅ Results cached on client
- ✅ Faster overall

**Trade-off:**
- Larger initial query
- But only runs ONCE per session

### Phase 3: Smart Caching Strategy

**Cache Levels:**

1. **Session Cache (In-Memory):**
   - Survives Excel session
   - Cleared on workbook close
   - Fast lookups

2. **Sheet-Level Pre-Loading:**
   - On first formula evaluation, scan sheet
   - Identify ALL accounts/periods needed
   - Fetch ALL data in ONE request
   - Populate cache
   - All subsequent formulas = instant (cache hits)

3. **Refresh Strategy:**
   - "Refresh All" button → clears cache + refetches
   - "Refresh Selected" → clears cache for selection + refetches
   - Auto-refresh = OFF (user controlled)

---

## 🛠️ IMPLEMENTATION STEPS

### Step 1: Convert to Non-Streaming ✅

**Files to Update:**
- `docs/functions.js`
  - Change GLABAL to async function (not streaming)
  - Change GLABUD to async function (not streaming)
  - Remove invocation handling
  - Add Promise return

- `docs/functions.json`
  - Change `"stream": true` → `"stream": false`
  - Keep `"volatile": false`
  - Keep `"cancelable": true`

**Testing:**
- Open sheet → formulas should NOT recalculate
- Click "Refresh All" → formulas should update
- Reopen sheet → formulas show cached values (no @)

### Step 2: Optimize Batching ✅

**Files to Update:**
- `docs/functions.js`
  - Implement "sheet scan" to find all accounts/periods
  - Batch ALL formulas in ONE request
  - Populate cache with all results
  - Return values from cache

- `backend/server.py`
  - Keep current batch_balance endpoint
  - Already optimized with BUILTIN.CONSOLIDATE
  - Just needs larger batch size support

**Testing:**
- Sheet with 100 formulas
- Should make 1-2 requests max (not 100)
- Cache hit rate > 90%

### Step 3: Add Cache Statistics ✅

**For Debugging:**
```javascript
console.log('Cache Stats:');
console.log(`  Hits: ${cacheStats.hits}`);
console.log(`  Misses: ${cacheStats.misses}`);
console.log(`  Hit Rate: ${hitRate}%`);
console.log(`  Cache Size: ${cacheStats.size()}`);
```

---

## 📊 EXPECTED RESULTS

### Before Optimization:
```
Sheet Open:
  • 100 formulas × 100ms each = 10 seconds
  • Every open = recalculation = slow
  • @ symbols everywhere

Performance:
  • Multiple small batches
  • Moderate speed
```

### After Optimization:
```
Sheet Open (First Time):
  • 100 formulas
  • Single batch request = 2 seconds
  • All data cached

Sheet Open (Subsequent):
  • 100 formulas × 0ms (cached) = instant
  • No @ symbols
  • No network requests

Performance:
  • ONE large batch
  • Much faster overall
```

---

## ⚠️ RISKS & MITIGATION

### Risk 1: Large Query Timeout

**Risk:** Single query for 1000+ accounts might timeout

**Mitigation:**
- Implement "smart chunking" - divide into 2-3 large chunks
- Each chunk = 500 accounts max
- Still better than 100 small requests

### Risk 2: Memory Usage

**Risk:** Caching 10,000+ values might use too much memory

**Mitigation:**
- LRU (Least Recently Used) cache eviction
- Max cache size = 10,000 entries
- Monitor with `cacheStats.size()`

### Risk 3: Stale Data

**Risk:** Users forget to refresh, see old data

**Mitigation:**
- Clear visual indicator in task pane
- "Last Refreshed: 10 minutes ago"
- Auto-suggest refresh after 1 hour
- Red warning after 24 hours

---

## 🔄 ROLLBACK PLAN

If optimization causes issues:

```bash
# Revert to working version
git checkout backup-v1.0.0.87-before-optimization

# Or use tag
git checkout v1.0.0.87-working

# Push to revert GitHub
git push origin main --force  # (only if necessary)
```

**Backup Locations:**
- Branch: `backup-v1.0.0.87-before-optimization`
- Tag: `v1.0.0.87-working`
- All code is safe

---

## ✅ SUCCESS CRITERIA

1. **No Recalculation on Open**
   - Open sheet → No @ symbols
   - Formulas show values instantly

2. **Fast Performance**
   - First open: < 3 seconds for 100 formulas
   - Subsequent opens: < 0.5 seconds (cached)

3. **User Control**
   - Refresh only when user clicks button
   - Clear feedback on last refresh time

4. **Reliability**
   - No #VALUE# errors
   - Cache hit rate > 90%
   - Works for sheets with 1000+ formulas

---

## 🚀 IMPLEMENTATION ORDER

1. ✅ Create backup
2. ⏳ Convert GLABAL/GLABUD to non-streaming
3. ⏳ Update functions.json
4. ⏳ Test recalculation fix
5. ⏳ Optimize batching strategy
6. ⏳ Add cache statistics
7. ⏳ Test performance
8. ⏳ Update manifest version
9. ⏳ Deploy and document

---

**Ready to implement! Let's start with Phase 1: Fix Recalculation** 🚀

