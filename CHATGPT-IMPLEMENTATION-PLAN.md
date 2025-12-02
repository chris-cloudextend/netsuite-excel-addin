# ChatGPT Two-Mode Architecture Implementation Plan

## Executive Summary

ChatGPT identified that our SuiteQL query structure was the bottleneck, NOT `BUILTIN.CONSOLIDATE` itself. The user's working query returns ALL accounts × 12 months in **< 30 seconds** because it applies consolidation in a subquery before grouping, rather than inside `SUM()`.

**Goal:** Implement two-mode operation:
- **Mode 1:** Small batches (current approach) for individual formulas
- **Mode 2:** Full-sheet refresh (one big query) when "Refresh All" is clicked

**Expected Result:** Full refresh time drops from 6-8 minutes to ~30-60 seconds.

---

## Phase 2A: Optimize Backend Query Structure

### Current Query (SLOW):
```sql
SELECT 
    a.acctnumber,
    ap.periodname,
    SUM(
        CASE WHEN sub_count > 1 THEN
            TO_NUMBER(BUILTIN.CONSOLIDATE(...))  -- Inside SUM()!
        ELSE tal.amount END
    ) AS total_amount
FROM TransactionAccountingLine tal
...
GROUP BY a.acctnumber, ap.periodname
```

**Problem:** `BUILTIN.CONSOLIDATE` is called inside `SUM()`, forcing NetSuite to consolidate aggregated rows (inefficient).

### Optimized Query (FAST):
```sql
WITH sub_cte AS (
  SELECT COUNT(*) AS subs_count
  FROM subsidiary
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
            {{TARGET_SUBSIDIARY_ID}},
            t.postingperiod,
            'DEFAULT'
          )
        )
      ELSE tal.amount
    END
    * CASE WHEN a.accttype IN ('Income','OthIncome') THEN -1 ELSE 1 END
    AS cons_amt
  FROM transactionaccountingline tal
  JOIN transaction t ON t.id = tal.transaction
  JOIN account a ON a.id = tal.account
  JOIN accountingperiod ap ON ap.id = t.postingperiod
  CROSS JOIN sub_cte
  WHERE t.posting = 'T'
    AND tal.posting = 'T'
    AND tal.accountingbook = 1
    AND ap.isyear = 'F'
    AND ap.isquarter = 'F'
    AND EXTRACT(YEAR FROM ap.startdate) = {{FISCAL_YEAR}}
    AND COALESCE(a.eliminate, 'F') = 'F'
    AND a.accttype IN ('Income','COGS','Cost of Goods Sold','Expense','OthIncome','OthExpense')
    AND a.acctnumber IN ({{ACCOUNT_LIST}})
    -- Optional filters
)
SELECT
  a.acctnumber AS account_number,
  TO_CHAR(ap.startdate,'YYYY-MM') AS month,
  SUM(b.cons_amt) AS amount
FROM base b
JOIN accountingperiod ap ON ap.id = b.period_id
JOIN account a ON a.id = b.account_id
GROUP BY a.acctnumber, ap.startdate
HAVING SUM(b.cons_amt) <> 0
ORDER BY a.acctnumber, ap.startdate
```

**Why This is Fast:**
1. `BUILTIN.CONSOLIDATE` applied ONCE per transaction line in the CTE
2. Result stored as `cons_amt` (pre-consolidated)
3. Outer query just groups pre-consolidated values
4. NetSuite's query optimizer can process this efficiently

---

## Phase 2B: Implement Two-Mode Operation

### Frontend Changes (docs/functions.js):

#### 1. Add Mode Detection

```javascript
// Global state
let isFullRefreshMode = false;
let fullRefreshResolver = null;

// Called by task pane when "Refresh All" is clicked
window.enterFullRefreshMode = () => {
    console.log('🚀 ENTERING FULL REFRESH MODE');
    isFullRefreshMode = true;
    
    // Clear cache to force fresh data
    for (const key in cache) {
        if (cache[key] instanceof Map) {
            cache[key].clear();
        }
    }
    
    // Create a Promise that will resolve when full refresh is complete
    return new Promise((resolve) => {
        fullRefreshResolver = resolve;
    });
};

window.exitFullRefreshMode = () => {
    console.log('✅ EXITING FULL REFRESH MODE');
    isFullRefreshMode = false;
    if (fullRefreshResolver) {
        fullRefreshResolver();
        fullRefreshResolver = null;
    }
};
```

#### 2. Modify GLABAL/GLABUD to Check Mode

```javascript
async function GLABAL(account, fromPeriod, toPeriod, subsidiary, department, location, classId) {
    // ... normalize parameters ...
    const cacheKey = getCacheKey('balance', params);

    // Check cache first
    if (cache.balance.has(cacheKey)) {
        return cache.balance.get(cacheKey);
    }

    // If in full refresh mode, queue for bulk fetch
    if (isFullRefreshMode) {
        return new Promise((resolve, reject) => {
            pendingRequests.balance.set(cacheKey, { params, resolve, reject });
        });
    }

    // Mode 1: Small batch (current behavior)
    return new Promise((resolve, reject) => {
        pendingRequests.balance.set(cacheKey, { params, resolve, reject });
        startBatchTimer();
    });
}
```

#### 3. Add Full Refresh Batch Processor

```javascript
async function processFullRefresh() {
    console.log('🚀 PROCESSING FULL REFRESH');
    
    const allRequests = Array.from(pendingRequests.balance.entries());
    
    // Collect all unique accounts and periods
    const allAccounts = new Set();
    const allPeriods = new Set();
    const filters = {}; // Assume same filters for all (or pick most common)
    
    for (const [cacheKey, request] of allRequests) {
        allAccounts.add(request.params.account);
        if (request.params.fromPeriod) allPeriods.add(request.params.fromPeriod);
        if (request.params.toPeriod) allPeriods.add(request.params.toPeriod);
        
        // Use filters from first request (or validate all same)
        if (Object.keys(filters).length === 0) {
            filters.subsidiary = request.params.subsidiary || '';
            filters.department = request.params.department || '';
            filters.location = request.params.location || '';
            filters.class = request.params.classId || '';
        }
    }
    
    console.log(`📊 Full Refresh: ${allAccounts.size} accounts × ${allPeriods.size} periods`);
    
    // Call new backend endpoint
    const payload = {
        accounts: Array.from(allAccounts),
        periods: Array.from(allPeriods),
        ...filters
    };
    
    try {
        const response = await fetch(`${API_BASE_URL}/batch/full_refresh`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }
        
        const data = await response.json();
        const balances = data.balances || {};
        
        // Populate cache with ALL results
        for (const account in balances) {
            for (const period in balances[account]) {
                const cacheKey = getCacheKey('balance', { 
                    account, 
                    fromPeriod: period, 
                    toPeriod: period,
                    ...filters
                });
                cache.balance.set(cacheKey, balances[account][period]);
            }
        }
        
        console.log(`✅ Cached ${Object.keys(balances).length} accounts`);
        
        // Resolve all pending requests
        for (const [cacheKey, request] of allRequests) {
            const account = request.params.account;
            const fromPeriod = request.params.fromPeriod;
            const toPeriod = request.params.toPeriod;
            
            // Sum requested period range
            let total = 0;
            if (fromPeriod === toPeriod) {
                total = (balances[account] && balances[account][fromPeriod]) || 0;
            } else {
                // Sum multiple periods if range specified
                const periodRange = expandPeriodRange(fromPeriod, toPeriod);
                for (const period of periodRange) {
                    total += (balances[account] && balances[account][period]) || 0;
                }
            }
            
            request.resolve(total);
        }
        
        pendingRequests.balance.clear();
        
    } catch (error) {
        console.error('❌ Full refresh failed:', error);
        
        // Reject all pending requests
        for (const [cacheKey, request] of allRequests) {
            request.reject(error);
        }
        
        pendingRequests.balance.clear();
    } finally {
        window.exitFullRefreshMode();
    }
}
```

---

### Backend Changes (backend/server.py):

#### 1. Add New Endpoint

```python
@app.route('/batch/full_refresh', methods=['POST'])
def batch_full_refresh():
    """
    Full-sheet refresh: One big query for all accounts and periods.
    Expected to return in ~30 seconds for 100 accounts × 12 months.
    """
    try:
        data = request.json
        accounts = data.get('accounts', [])
        periods = data.get('periods', [])
        subsidiary = data.get('subsidiary', '')
        department = data.get('department', '')
        location = data.get('location', '')
        class_id = data.get('class', '')
        
        if not accounts or not periods:
            return jsonify({'error': 'accounts and periods required'}), 400
        
        # Extract year from first period
        # Assume all periods are same year (or handle multiple years)
        fiscal_year = extract_year_from_period(periods[0])
        
        # Build optimized query using ChatGPT's pattern
        query = build_full_refresh_query(
            accounts, fiscal_year, subsidiary, department, location, class_id
        )
        
        # Execute query
        results = execute_suiteql_query(query)
        
        # Transform results to nested dict: { account: { period: value } }
        balances = {}
        for row in results:
            account = row.get('account_number')
            month = row.get('month')  # 'YYYY-MM' format
            amount = float(row.get('amount', 0))
            
            # Convert 'YYYY-MM' to 'Mon YYYY' format
            period_name = convert_to_period_name(month)
            
            if account not in balances:
                balances[account] = {}
            balances[account][period_name] = amount
        
        return jsonify({'balances': balances})
    
    except Exception as e:
        print(f"ERROR in full_refresh: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


def build_full_refresh_query(accounts, fiscal_year, subsidiary, department, location, class_id):
    """
    Build optimized full-refresh query using ChatGPT's CTE pattern.
    """
    # Get target subsidiary
    target_sub = get_target_subsidiary(subsidiary)
    target_sub_str = str(target_sub) if target_sub else '1'
    
    # Build account list
    account_list = ','.join([f"'{acc}'" for acc in accounts])
    
    # Build optional filters
    filters = []
    if subsidiary:
        filters.append(f"t.subsidiary = {subsidiary}")
    if department:
        filters.append(f"tal.department = {department}")
    if location:
        filters.append(f"tal.location = {location}")
    if class_id:
        filters.append(f"tal.class = {class_id}")
    
    filter_clause = " AND " + " AND ".join(filters) if filters else ""
    
    query = f"""
    WITH sub_cte AS (
      SELECT COUNT(*) AS subs_count
      FROM subsidiary
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
                {target_sub_str},
                t.postingperiod,
                'DEFAULT'
              )
            )
          ELSE tal.amount
        END
        * CASE WHEN a.accttype IN ('Income','OthIncome') THEN -1 ELSE 1 END
        AS cons_amt
      FROM transactionaccountingline tal
      JOIN transaction t ON t.id = tal.transaction
      JOIN account a ON a.id = tal.account
      JOIN accountingperiod ap ON ap.id = t.postingperiod
      CROSS JOIN sub_cte
      WHERE t.posting = 'T'
        AND tal.posting = 'T'
        AND tal.accountingbook = 1
        AND ap.isyear = 'F'
        AND ap.isquarter = 'F'
        AND EXTRACT(YEAR FROM ap.startdate) = {fiscal_year}
        AND COALESCE(a.eliminate, 'F') = 'F'
        AND a.accttype IN ('Income','COGS','Cost of Goods Sold','Expense','OthIncome','OthExpense')
        AND a.acctnumber IN ({account_list})
        {filter_clause}
    )
    SELECT
      a.acctnumber AS account_number,
      TO_CHAR(ap.startdate,'YYYY-MM') AS month,
      SUM(b.cons_amt) AS amount
    FROM base b
    JOIN accountingperiod ap ON ap.id = b.period_id
    JOIN account a ON a.id = b.account_id
    GROUP BY a.acctnumber, ap.startdate
    HAVING SUM(b.cons_amt) <> 0
    ORDER BY a.acctnumber, ap.startdate
    """
    
    return query


def extract_year_from_period(period_name):
    """Extract year from 'Jan 2024' format"""
    parts = period_name.split()
    if len(parts) == 2:
        return int(parts[1])
    return 2024  # Default


def convert_to_period_name(month_str):
    """Convert 'YYYY-MM' to 'Mon YYYY' format"""
    from datetime import datetime
    dt = datetime.strptime(month_str, '%Y-%m')
    return dt.strftime('%b %Y')
```

---

### Task Pane Changes (docs/taskpane.html):

```javascript
async function refreshCurrentSheet() {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = '🔄 Starting full refresh...';
    statusDiv.className = 'status-info';
    
    try {
        await Excel.run(async context => {
            // Enter full refresh mode
            await window.enterFullRefreshMode();
            
            statusDiv.textContent = '📊 Collecting formulas...';
            
            // Trigger recalculation
            context.workbook.application.calculate(Excel.CalculationType.recalculate);
            await context.sync();
            
            statusDiv.textContent = '⏳ Fetching data from NetSuite...';
            
            // Wait for full refresh to complete
            // (processFullRefresh will call exitFullRefreshMode when done)
            await new Promise(resolve => setTimeout(resolve, 100)); // Give formulas time to queue
            await window.processFullRefresh();
            
            statusDiv.textContent = '✅ Refresh complete!';
            statusDiv.className = 'status-success';
        });
    } catch (error) {
        statusDiv.textContent = `❌ Error: ${error.message}`;
        statusDiv.className = 'status-error';
        window.exitFullRefreshMode();
    }
}
```

---

## Implementation Steps

### Step 1: Test Query Optimization (Backend Only)
1. Update `build_pl_query` in `backend/server.py` to use CTE pattern
2. Test with 100 accounts × 12 months
3. **Expected:** ~30 seconds (vs. current 226 seconds for 6 accounts)

### Step 2: Implement Full Refresh Endpoint
1. Add `/batch/full_refresh` endpoint to `backend/server.py`
2. Test with Postman/curl
3. Verify returns correct JSON structure

### Step 3: Implement Mode Detection (Frontend)
1. Add `enterFullRefreshMode()` and `exitFullRefreshMode()` to `docs/functions.js`
2. Modify `GLABAL`/`GLABUD` to check `isFullRefreshMode`
3. Add `processFullRefresh()` function

### Step 4: Update Task Pane
1. Modify `refreshCurrentSheet()` to call `enterFullRefreshMode()`
2. Test full refresh flow

### Step 5: Keep Mode 1 Intact
1. Ensure individual formula entry still uses period-by-period batching
2. Test drag-and-fill of one row

### Step 6: Handle Balance Sheet Accounts
1. Create separate `build_bs_full_refresh_query()` for Balance Sheet accounts
2. Apply same CTE pattern
3. Use `t.trandate <= period_end` for cumulative logic

---

## Expected Performance

### Before (Current):
- **Full Refresh:** 6-8 minutes for 100 accounts × 12 months
- **Individual Formula:** Fast (< 1 second)

### After (Optimized):
- **Full Refresh:** ~30-60 seconds for 100 accounts × 12 months
- **Individual Formula:** Fast (< 1 second) - unchanged
- **Backend Query:** One query instead of 24
- **Cache Population:** Instant formula resolution after first refresh

---

## Success Criteria

1. ✅ Full refresh completes in < 1 minute for 100 accounts × 12 months
2. ✅ Individual formula entry remains fast (< 1 second)
3. ✅ Backend makes ONE SuiteQL query per full refresh
4. ✅ All formulas resolve correctly from cache
5. ✅ No `$0` errors
6. ✅ No `N/A` errors for account names
7. ✅ Console shows clear "FULL REFRESH MODE" messages

---

## Rollback Plan

If full refresh fails:
1. Exit full refresh mode
2. Fall back to Mode 1 (period-by-period batching)
3. Return existing cached values if available

---

## Next Steps

1. **Immediate:** Test query optimization on backend
2. **Then:** Implement `/batch/full_refresh` endpoint
3. **Then:** Add mode detection to frontend
4. **Then:** Update task pane
5. **Finally:** Full QA and deployment

**Expected Timeline:** 2-3 hours for full implementation and testing.

