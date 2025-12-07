/**
 * NetSuite Custom Functions - SIMPLIFIED & BULLETPROOF
 * 
 * KEY DESIGN PRINCIPLES:
 * 1. Cache AGGRESSIVELY - never clear unless user clicks button
 * 2. Batch CONSERVATIVELY - small batches, long delays
 * 3. Single cell updates = individual API call (fast)
 * 4. Bulk updates (drag/insert row) = smart batching
 * 5. Deduplication - never make same request twice
 */

const SERVER_URL = 'https://netsuite-proxy.chris-corcoran.workers.dev';
const REQUEST_TIMEOUT = 30000;  // 30 second timeout for NetSuite queries

// ============================================================================
// STATUS BROADCAST - Communicate progress to taskpane via localStorage
// ============================================================================
function broadcastStatus(message, progress = 0, type = 'info') {
    try {
        localStorage.setItem('netsuite_status', JSON.stringify({
            message,
            progress,
            type,
            timestamp: Date.now()
        }));
    } catch (e) {
        // localStorage not available - ignore
    }
}

function clearStatus() {
    try {
        localStorage.removeItem('netsuite_status');
    } catch (e) {}
}

// ============================================================================
// PERIOD EXPANSION - Intelligently expand period ranges for better caching
// ============================================================================
const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                     'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/**
 * Parse "Mon YYYY" string into {month: 0-11, year: YYYY}
 */
function parsePeriod(periodStr) {
    if (!periodStr || typeof periodStr !== 'string') return null;
    const match = periodStr.match(/^([A-Za-z]{3})\s+(\d{4})$/);
    if (!match) return null;
    const monthIndex = MONTH_NAMES.findIndex(m => m.toLowerCase() === match[1].toLowerCase());
    if (monthIndex === -1) return null;
    return { month: monthIndex, year: parseInt(match[2]) };
}

/**
 * Convert {month, year} back to "Mon YYYY" string
 */
function formatPeriod(month, year) {
    return `${MONTH_NAMES[month]} ${year}`;
}

/**
 * Expand a list of periods to include adjacent months for better cache coverage.
 * This ensures that when dragging formulas, nearby periods are pre-fetched.
 * 
 * @param {string[]} periods - Array of "Mon YYYY" strings (e.g., ["Jan 2025", "Feb 2025"])
 * @param {number} expandBefore - Number of months to add before the earliest period (default: 1)
 * @param {number} expandAfter - Number of months to add after the latest period (default: 1)
 * @returns {string[]} Expanded array of periods
 */
function expandPeriodRange(periods, expandBefore = 1, expandAfter = 1) {
    if (!periods || periods.length === 0) return periods;
    
    // Parse all periods
    const parsed = periods.map(parsePeriod).filter(p => p !== null);
    if (parsed.length === 0) return periods;
    
    // Find min and max dates
    let minMonth = parsed[0].month;
    let minYear = parsed[0].year;
    let maxMonth = parsed[0].month;
    let maxYear = parsed[0].year;
    
    for (const p of parsed) {
        const pTotal = p.year * 12 + p.month;
        const minTotal = minYear * 12 + minMonth;
        const maxTotal = maxYear * 12 + maxMonth;
        
        if (pTotal < minTotal) {
            minMonth = p.month;
            minYear = p.year;
        }
        if (pTotal > maxTotal) {
            maxMonth = p.month;
            maxYear = p.year;
        }
    }
    
    // Expand backward
    for (let i = 0; i < expandBefore; i++) {
        minMonth--;
        if (minMonth < 0) {
            minMonth = 11;
            minYear--;
        }
    }
    
    // Expand forward
    for (let i = 0; i < expandAfter; i++) {
        maxMonth++;
        if (maxMonth > 11) {
            maxMonth = 0;
            maxYear++;
        }
    }
    
    // Generate all periods in the expanded range
    const expanded = [];
    let currentMonth = minMonth;
    let currentYear = minYear;
    
    while (currentYear < maxYear || (currentYear === maxYear && currentMonth <= maxMonth)) {
        expanded.push(formatPeriod(currentMonth, currentYear));
        currentMonth++;
        if (currentMonth > 11) {
            currentMonth = 0;
            currentYear++;
        }
    }
    
    console.log(`   📅 Period expansion: [${periods.join(', ')}] → [${expanded.join(', ')}]`);
    return expanded;
}

// ============================================================================
// CACHE - Never expires, persists entire Excel session
// ============================================================================
const cache = {
    balance: new Map(),    // balance cache
    title: new Map(),      // account title cache  
    budget: new Map(),     // budget cache
    type: new Map(),       // account type cache
    parent: new Map()      // parent account cache
};

// In-flight request tracking for expensive calculations (RE, NI, CTA)
// This prevents duplicate concurrent API calls for the same period
const inFlightRequests = new Map();

// Track last access time to implement LRU if needed
const cacheStats = {
    hits: 0,
    misses: 0,
    size: () => cache.balance.size + cache.title.size + cache.budget.size + cache.type.size + cache.parent.size
};

// ============================================================================
// GLOBAL CACHE CONTROL - Accessible from taskpane
// ============================================================================
window.clearAllCaches = function() {
    console.log('🗑️  CLEARING ALL CACHES...');
    console.log(`  Before: ${cache.balance.size} balances, ${cache.title.size} titles, ${cache.budget.size} budgets`);
    
    cache.balance.clear();
    cache.title.clear();
    cache.budget.clear();
    cache.type.clear();
    cache.parent.clear();
    
    // Reset stats
    cacheStats.hits = 0;
    cacheStats.misses = 0;
    
    console.log('✅ ALL CACHES CLEARED');
    return true;
};

/**
 * Selectively clear cache for specific account/period combinations
 * Use this for "Refresh Selected" to avoid clearing ALL cached data
 * Clears from: 1) in-memory cache, 2) fullYearCache, 3) localStorage
 * @param {Array<{account: string, period: string}>} items - Array of {account, period} to clear
 * @returns {number} Number of cache entries cleared
 */
window.clearCacheForItems = function(items) {
    console.log(`🎯 SELECTIVE CACHE CLEAR: ${items.length} items`);
    let cleared = 0;
    
    // Also clear from localStorage - this is critical!
    let localStorageCleared = 0;
    try {
        const STORAGE_KEY = 'netsuite_balance_cache';
        const stored = localStorage.getItem(STORAGE_KEY);
        if (stored) {
            const balanceData = JSON.parse(stored);
            let modified = false;
            
            for (const item of items) {
                const acct = String(item.account);
                if (balanceData[acct] && balanceData[acct][item.period] !== undefined) {
                    delete balanceData[acct][item.period];
                    localStorageCleared++;
                    modified = true;
                    console.log(`   ✓ Cleared localStorage: ${acct}/${item.period}`);
                }
            }
            
            if (modified) {
                localStorage.setItem(STORAGE_KEY, JSON.stringify(balanceData));
                console.log(`   💾 Updated localStorage (removed ${localStorageCleared} entries)`);
            }
        }
    } catch (e) {
        console.warn('   ⚠️ Error clearing localStorage:', e);
    }
    
    for (const item of items) {
        // Use getCacheKey to ensure exact same format as BALANCE
        const cacheKey = getCacheKey('balance', {
            account: String(item.account),
            fromPeriod: item.period,
            toPeriod: item.period,
            subsidiary: item.subsidiary || '',
            department: item.department || '',
            location: item.location || '',
            classId: item.classId || ''
        });
        
        if (cache.balance.has(cacheKey)) {
            cache.balance.delete(cacheKey);
            cleared++;
            console.log(`   ✓ Cleared in-memory: ${item.account}/${item.period}`);
        }
        
        // Also clear from fullYearCache if it exists
        if (fullYearCache && fullYearCache[item.account]) {
            if (fullYearCache[item.account][item.period] !== undefined) {
                delete fullYearCache[item.account][item.period];
                console.log(`   ✓ Cleared fullYearCache: ${item.account}/${item.period}`);
            }
        }
    }
    
    const totalCleared = cleared + localStorageCleared;
    console.log(`   📊 Cleared ${totalCleared} total cache entries (${cleared} in-memory, ${localStorageCleared} localStorage)`);
    return totalCleared;
};

// ============================================================================
// FULL REFRESH MODE - Optimized for bulk sheet refresh
// ============================================================================
let isFullRefreshMode = false;
let fullRefreshResolver = null;
let fullRefreshYear = null;

// CACHE READY SEMAPHORE - Prevents premature formula evaluation
// Formulas check this before returning values during full refresh
window.__NS_CACHE_READY = true;  // Default to true for normal operation

window.enterFullRefreshMode = function(year) {
    console.log('🚀 ENTERING FULL REFRESH MODE');
    console.log(`   Year: ${year || 'auto-detect'}`);
    isFullRefreshMode = true;
    fullRefreshYear = year || null;
    window.__NS_CACHE_READY = false;  // Block premature evaluations
    
    // Clear cache to force fresh data
    window.clearAllCaches();
    
    // Return a Promise that resolves when full refresh completes
    return new Promise((resolve) => {
        fullRefreshResolver = resolve;
    });
};

window.exitFullRefreshMode = function() {
    console.log('✅ EXITING FULL REFRESH MODE');
    isFullRefreshMode = false;
    fullRefreshYear = null;
    window.__NS_CACHE_READY = true;  // Allow normal evaluation
    if (fullRefreshResolver) {
        fullRefreshResolver();
        fullRefreshResolver = null;
    }
};

// Mark cache as ready (called by taskpane after populating cache)
window.markCacheReady = function() {
    console.log('✅ CACHE MARKED AS READY');
    window.__NS_CACHE_READY = true;
};

// ============================================================================
// BUILD MODE - Instant drag-and-drop performance
// When user drags formulas rapidly, we defer NetSuite calls until they stop
// 
// KEY INSIGHT: We DON'T return 0 placeholder - that looks like real data!
// Instead, we return a Promise that resolves after the batch completes.
// This shows #BUSY briefly but ensures correct values.
// ============================================================================
let buildMode = false;
let buildModeLastEvent = 0;
let buildModePending = [];  // Collect pending requests: { cacheKey, params, resolve, reject }
let buildModeTimer = null;
let formulaCreationCount = 0;
let formulaCountResetTimer = null;

const BUILD_MODE_THRESHOLD = 2;     // Enter build mode after 2+ rapid formulas (trigger earlier)
const BUILD_MODE_SETTLE_MS = 500;   // Wait 500ms after last formula before batch
const BUILD_MODE_WINDOW_MS = 800;   // Time window to count formulas (wider!)

function enterBuildMode() {
    if (!buildMode) {
        console.log('🔨 ENTERING BUILD MODE (rapid formula creation detected)');
        buildMode = true;
        
        // CRITICAL: Cancel the regular batch timer to prevent race condition!
        if (batchTimer) {
            console.log('   ⏹️ Cancelled regular batch timer');
            clearTimeout(batchTimer);
            batchTimer = null;
        }
        
        // Move any already-pending requests into build mode queue
        // This prevents them from being processed by the regular batch
        for (const [cacheKey, request] of pendingRequests.balance.entries()) {
            buildModePending.push({
                cacheKey,
                params: request.params,
                resolve: request.resolve,
                reject: request.reject
            });
        }
        if (pendingRequests.balance.size > 0) {
            console.log(`   📦 Moved ${pendingRequests.balance.size} pending requests to build mode`);
            pendingRequests.balance.clear();
        }
    }
}

function exitBuildModeAndProcess() {
    if (!buildMode) return;
    
    const count = buildModePending.length;
    console.log(`🔨 EXITING BUILD MODE (${count} formulas queued)`);
    buildMode = false;
    formulaCreationCount = 0;
    
    // Process all queued formulas in one batch
    if (count > 0) {
        runBuildModeBatch();
    }
}

// ============================================================================
// NETSUITE ACCOUNT TYPES - Complete Reference
// ============================================================================
// 
// BALANCE SHEET ACCOUNTS:
// -----------------------
// ASSETS (Natural Debit Balance - stored POSITIVE in NetSuite)
//   SuiteQL Value     | Description              | Sign in Report
//   ------------------|--------------------------|----------------
//   Bank              | Bank/Cash accounts       | + (no flip)
//   AcctRec           | Accounts Receivable      | + (no flip)
//   OthCurrAsset      | Other Current Asset      | + (no flip)
//   FixedAsset        | Fixed Asset              | + (no flip)
//   OthAsset          | Other Asset              | + (no flip)
//   DeferExpense      | Deferred Expense         | + (no flip)
//   UnbilledRec       | Unbilled Receivable      | + (no flip)
//
// LIABILITIES (Natural Credit Balance - stored NEGATIVE in NetSuite)
//   SuiteQL Value     | Description              | Sign in Report
//   ------------------|--------------------------|----------------
//   AcctPay           | Accounts Payable         | + (flip × -1)
//   CredCard          | Credit Card              | + (flip × -1)
//   OthCurrLiab       | Other Current Liability  | + (flip × -1)
//   LongTermLiab      | Long Term Liability      | + (flip × -1)
//   DeferRevenue      | Deferred Revenue         | + (flip × -1)
//
// EQUITY (Natural Credit Balance - stored NEGATIVE in NetSuite)
//   SuiteQL Value     | Description              | Sign in Report
//   ------------------|--------------------------|----------------
//   Equity            | Equity accounts          | + (flip × -1)
//   RetainedEarnings  | Retained Earnings        | + (flip × -1)
//
// PROFIT & LOSS ACCOUNTS:
// -----------------------
// INCOME (Natural Credit Balance - stored NEGATIVE in NetSuite)
//   SuiteQL Value     | Description              | Sign in Report
//   ------------------|--------------------------|----------------
//   Income            | Revenue/Sales            | + (flip × -1)
//   OthIncome         | Other Income             | + (flip × -1)
//
// EXPENSES (Natural Debit Balance - stored POSITIVE in NetSuite)
//   SuiteQL Value     | Description              | Sign in Report
//   ------------------|--------------------------|----------------
//   COGS              | Cost of Goods Sold       | + (no flip)
//   Expense           | Operating Expense        | + (no flip)
//   OthExpense        | Other Expense            | + (no flip)
//
// OTHER ACCOUNT TYPES:
// --------------------
//   NonPosting        | Non-posting/Statistical  | N/A (no transactions)
//   Stat              | Statistical accounts     | N/A (no transactions)
//
// ============================================================================

// Helper: Check if account type is Balance Sheet
function isBalanceSheetType(acctType) {
    if (!acctType) return false;
    // All Balance Sheet account types (Assets, Liabilities, Equity)
    const bsTypes = [
        // Assets (Debit balance)
        'Bank',           // Bank/Cash accounts
        'AcctRec',        // Accounts Receivable
        'OthCurrAsset',   // Other Current Asset
        'FixedAsset',     // Fixed Asset
        'OthAsset',       // Other Asset
        'DeferExpense',   // Deferred Expense (prepaid expenses)
        'UnbilledRec',    // Unbilled Receivable
        // Liabilities (Credit balance)
        'AcctPay',        // Accounts Payable
        'CredCard',       // Credit Card (NOT 'CreditCard')
        'OthCurrLiab',    // Other Current Liability
        'LongTermLiab',   // Long Term Liability
        'DeferRevenue',   // Deferred Revenue (unearned revenue)
        // Equity (Credit balance)
        'Equity',         // Equity accounts
        'RetainedEarnings' // Retained Earnings
    ];
    return bsTypes.includes(acctType);
}

// Helper: Check if account type needs sign flip for Balance Sheet display
// Liabilities and Equity are stored as negative credits but display as positive
function needsSignFlip(acctType) {
    if (!acctType) return false;
    const flipTypes = [
        // Liabilities (stored negative, display positive)
        'AcctPay',        // Accounts Payable
        'CredCard',       // Credit Card
        'OthCurrLiab',    // Other Current Liability
        'LongTermLiab',   // Long Term Liability
        'DeferRevenue',   // Deferred Revenue
        // Equity (stored negative, display positive)
        'Equity',         // Equity
        'RetainedEarnings' // Retained Earnings
    ];
    return flipTypes.includes(acctType);
}

// Helper: Get account type from cache or fetch it
async function getAccountType(account) {
    const cacheKey = getCacheKey('type', { account });
    if (cache.type.has(cacheKey)) {
        return cache.type.get(cacheKey);
    }
    
    try {
        // Use POST to avoid exposing account numbers in URLs/logs
        const response = await fetch(`${SERVER_URL}/account/type`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ account: String(account) })
        });
        if (response.ok) {
            const type = await response.text();
            cache.type.set(cacheKey, type);
            return type;
        }
    } catch (e) {
        console.warn(`   ⚠️ Failed to get type for ${account}:`, e.message);
    }
    return null;
}

// Helper: Batch get account types (much faster than individual calls)
async function batchGetAccountTypes(accounts) {
    const result = {};
    const uncached = [];
    
    // First check cache
    for (const acct of accounts) {
        const cacheKey = getCacheKey('type', { account: acct });
        if (cache.type.has(cacheKey)) {
            result[acct] = cache.type.get(cacheKey);
        } else {
            uncached.push(acct);
        }
    }
    
    // Fetch uncached in batch
    if (uncached.length > 0) {
        try {
            const response = await fetch(`${SERVER_URL}/batch/account_types`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ accounts: uncached })  // Backend expects 'accounts'
            });
            if (response.ok) {
                const data = await response.json();
                const types = data.types || {};  // Backend returns 'types'
                for (const acct of uncached) {
                    const type = types[acct] || null;
                    result[acct] = type;
                    // Cache for future use
                    const cacheKey = getCacheKey('type', { account: acct });
                    cache.type.set(cacheKey, type);
                }
            }
        } catch (e) {
            console.warn(`   ⚠️ Batch account types failed:`, e.message);
        }
    }
    
    return result;
}

// Helper function to create a filter key for grouping
function getFilterKey(params) {
    const sub = String(params.subsidiary || '').trim();
    const dept = String(params.department || '').trim();
    const loc = String(params.location || '').trim();
    const cls = String(params.classId || '').trim();
    return `${sub}|${dept}|${loc}|${cls}`;
}

// Helper function to parse filter key back to filter object
function parseFilterKey(filterKey) {
    const parts = filterKey.split('|');
    return {
        subsidiary: parts[0] || '',
        department: parts[1] || '',
        location: parts[2] || '',
        classId: parts[3] || ''
    };
}

async function runBuildModeBatch() {
    const batchStartTime = Date.now();
    const pending = buildModePending.slice();
    buildModePending = [];
    
    if (pending.length === 0) return;
    
    console.log(`🔨 BUILD MODE BATCH: ${pending.length} formulas (started at ${new Date().toLocaleTimeString()})`);
    broadcastStatus(`Processing ${pending.length} formulas...`, 5, 'info');
    
    // STEP 1: Group pending formulas by their filter combination
    // This handles cases where formulas have different subsidiaries, etc.
    const filterGroups = new Map(); // filterKey -> array of pending items
    
    for (const item of pending) {
        const filterKey = getFilterKey(item.params);
        if (!filterGroups.has(filterKey)) {
            filterGroups.set(filterKey, []);
        }
        filterGroups.get(filterKey).push(item);
    }
    
    const groupCount = filterGroups.size;
    if (groupCount > 1) {
        console.log(`   📋 Detected ${groupCount} different filter combinations - processing each separately`);
    }
    
    // Collect ALL unique accounts to detect types (shared across all filter groups)
    const allAccountsSet = new Set();
    for (const item of pending) {
        allAccountsSet.add(item.params.account);
    }
    const allAccountsArray = Array.from(allAccountsSet);
    
    // STEP 2: Detect account types ONCE (account types don't depend on filters)
    console.log(`   🔍 Detecting account types for ${allAccountsArray.length} accounts...`);
    broadcastStatus(`Detecting account types...`, 5, 'info');
    const accountTypes = await batchGetAccountTypes(allAccountsArray);
    
    // STEP 3: Process each filter group separately
    let groupIndex = 0;
    let totalResolved = 0;
    let totalZeros = 0;
    
    for (const [filterKey, groupItems] of filterGroups) {
        groupIndex++;
        const filters = parseFilterKey(filterKey);
        
        if (groupCount > 1) {
            console.log(`\n   📦 Processing filter group ${groupIndex}/${groupCount}: ${groupItems.length} formulas`);
            console.log(`      Filters: sub="${filters.subsidiary}", dept="${filters.department}", loc="${filters.location}", class="${filters.classId}"`);
        }
        
        // Collect unique accounts and periods for THIS filter group
        const accounts = new Set();
        const periods = new Set();
        
        for (const item of groupItems) {
            const p = item.params;
            accounts.add(p.account);
            if (p.fromPeriod && p.fromPeriod !== '') {
                periods.add(p.fromPeriod);
            }
            if (p.toPeriod && p.toPeriod !== '' && p.toPeriod !== p.fromPeriod) {
                periods.add(p.toPeriod);
            }
            if (!p.fromPeriod && p.toPeriod) {
                periods.add(p.toPeriod);
            }
        }
        
        const periodsArray = Array.from(periods).filter(p => p && p !== '');
        const accountsArray = Array.from(accounts);
        
        console.log(`   Accounts: ${accountsArray.join(', ')}`);
        console.log(`   Periods (${periodsArray.length}): ${periodsArray.join(', ')}`);
        
        const allBalances = {};
        let hasError = false;
        
        // Detect years from periods
        const years = new Set(periodsArray.filter(p => p && p.includes(' ')).map(p => p.split(' ')[1]));
        const yearsArray = Array.from(years).filter(y => y && !isNaN(parseInt(y)));
        
        // Classify accounts for this group
        const plAccounts = [];
        const bsAccounts = [];
        for (const acct of accountsArray) {
            if (isBalanceSheetType(accountTypes[acct])) {
                bsAccounts.push(acct);
            } else {
                plAccounts.push(acct);
            }
        }
        console.log(`   📊 Account split: ${plAccounts.length} P&L, ${bsAccounts.length} Balance Sheet`);
        
        const usePLFullYear = yearsArray.length > 0 && plAccounts.length >= 5;
        
        // STEP 4: Fetch Balance Sheet accounts for this filter group
        if (bsAccounts.length > 0 && periodsArray.length > 0) {
            // SMART PERIOD EXPANSION: Include adjacent months for better cache coverage
            // This ensures that when user drags Jan→Feb, we also fetch Dec for them
            const expandedBSPeriods = expandPeriodRange(periodsArray, 1, 1);
            console.log(`   📅 BS periods expanded: ${periodsArray.length} → ${expandedBSPeriods.length}`);
            
            // CHECK CACHE FIRST (using EXPANDED periods)
            let allBSInCache = true;
            let cachedBSValues = {};
            
            for (const acct of bsAccounts) {
                cachedBSValues[acct] = {};
                for (const period of expandedBSPeriods) {
                    const ck = getCacheKey('balance', {
                        account: acct,
                        fromPeriod: period,
                        toPeriod: period,
                        subsidiary: filters.subsidiary,
                        department: filters.department,
                        location: filters.location,
                        classId: filters.classId
                    });
                    
                    if (cache.balance.has(ck)) {
                        cachedBSValues[acct][period] = cache.balance.get(ck);
                    } else {
                        allBSInCache = false;
                        break;
                    }
                }
                if (!allBSInCache) break;
            }
            
            if (allBSInCache) {
                console.log(`   ✅ BS CACHE HIT: All ${bsAccounts.length} accounts × ${expandedBSPeriods.length} periods found in cache!`);
                broadcastStatus(`Using cached Balance Sheet data`, 20, 'info');
                
                for (const acct of bsAccounts) {
                    if (!allBalances[acct]) allBalances[acct] = {};
                    for (const period of expandedBSPeriods) {
                        allBalances[acct][period] = cachedBSValues[acct][period];
                    }
                }
            } else {
                console.log(`   📊 Fetching Balance Sheet accounts (${expandedBSPeriods.length} periods, expanded from ${periodsArray.length})...`);
                broadcastStatus(`Fetching Balance Sheet data...`, 15, 'info');
                
                const bsStartTime = Date.now();
                try {
                    const response = await fetch(`${SERVER_URL}/batch/bs_periods`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            periods: expandedBSPeriods,  // Use expanded periods!
                            subsidiary: filters.subsidiary,
                            department: filters.department,
                            location: filters.location,
                            class: filters.classId,
                            accountingbook: filters.accountingBook || ''  // Multi-Book Accounting support
                        })
                    });
                    
                    if (response.ok) {
                        const data = await response.json();
                        const bsBalances = data.balances || {};
                        const bsTime = ((Date.now() - bsStartTime) / 1000).toFixed(1);
                        const bsAccountCount = Object.keys(bsBalances).length;
                        console.log(`   ✅ BS: ${bsAccountCount} accounts in ${bsTime}s`);
                        
                        // Cache ALL Balance Sheet accounts with THIS filter group's filters
                        let bsCached = 0;
                        for (const acct in bsBalances) {
                            if (!allBalances[acct]) allBalances[acct] = {};
                            for (const period in bsBalances[acct]) {
                                allBalances[acct][period] = bsBalances[acct][period];
                                const ck = getCacheKey('balance', {
                                    account: acct,
                                    fromPeriod: period,
                                    toPeriod: period,
                                    subsidiary: filters.subsidiary,
                                    department: filters.department,
                                    location: filters.location,
                                    classId: filters.classId
                                });
                                cache.balance.set(ck, bsBalances[acct][period]);
                                bsCached++;
                            }
                        }
                        console.log(`   💾 Cached ${bsCached} BS values with filters: sub="${filters.subsidiary}"`);
                    } else {
                        console.error(`   ❌ BS multi-period error: ${response.status}`);
                        hasError = true;
                    }
                } catch (error) {
                    console.error(`   ❌ BS fetch error:`, error);
                    hasError = true;
                }
            }
        }
        
        // STEP 5: Fetch P&L accounts for this filter group
        if (plAccounts.length > 0 && yearsArray.length > 0) {
            // CHECK CACHE FIRST for P&L accounts
            let allPLInCache = true;
            let cachedPLValues = {};
            
            for (const acct of plAccounts) {
                cachedPLValues[acct] = {};
                for (const period of periodsArray) {
                    const ck = getCacheKey('balance', {
                        account: acct,
                        fromPeriod: period,
                        toPeriod: period,
                        subsidiary: filters.subsidiary,
                        department: filters.department,
                        location: filters.location,
                        classId: filters.classId
                    });
                    
                    if (cache.balance.has(ck)) {
                        cachedPLValues[acct][period] = cache.balance.get(ck);
                    } else {
                        allPLInCache = false;
                        break;
                    }
                }
                if (!allPLInCache) break;
            }
            
            if (allPLInCache) {
                console.log(`   ✅ P&L CACHE HIT: All ${plAccounts.length} accounts × ${periodsArray.length} periods found in cache!`);
                broadcastStatus(`Using cached P&L data`, 70, 'info');
                
                for (const acct of plAccounts) {
                    if (!allBalances[acct]) allBalances[acct] = {};
                    for (const period of periodsArray) {
                        allBalances[acct][period] = cachedPLValues[acct][period];
                    }
                }
            } else if (usePLFullYear) {
                console.log(`   ⚡ P&L FAST PATH: full_year_refresh for ${yearsArray.length} year(s)`);
                broadcastStatus(`Fetching P&L data for ${yearsArray.join(', ')}...`, 60, 'info');
                
                try {
                    for (const year of yearsArray) {
                        const yearStartTime = Date.now();
                        console.log(`   📡 Fetching P&L year ${year}...`);
                        
                        const response = await fetch(`${SERVER_URL}/batch/full_year_refresh`, {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({
                                year: parseInt(year),
                                subsidiary: filters.subsidiary,
                                department: filters.department,
                                location: filters.location,
                                class: filters.classId,
                                skip_bs: true,
                                accountingbook: filters.accountingBook || ''  // Multi-Book Accounting support
                            })
                        });
                        
                        if (response.ok) {
                            const data = await response.json();
                            const yearBalances = data.balances || {};
                            const yearTime = ((Date.now() - yearStartTime) / 1000).toFixed(1);
                            console.log(`   ✅ P&L Year ${year}: ${Object.keys(yearBalances).length} accounts in ${yearTime}s`);
                            
                            // Cache with THIS filter group's filters
                            let plCached = 0;
                            for (const acct in yearBalances) {
                                if (!allBalances[acct]) allBalances[acct] = {};
                                for (const period in yearBalances[acct]) {
                                    allBalances[acct][period] = yearBalances[acct][period];
                                    const ck = getCacheKey('balance', {
                                        account: acct,
                                        fromPeriod: period,
                                        toPeriod: period,
                                        subsidiary: filters.subsidiary,
                                        department: filters.department,
                                        location: filters.location,
                                        classId: filters.classId
                                    });
                                    cache.balance.set(ck, yearBalances[acct][period]);
                                    plCached++;
                                }
                            }
                            console.log(`   💾 Cached ${plCached} P&L values`);
                        } else {
                            console.error(`   ❌ P&L Year ${year} error: ${response.status}`);
                            hasError = true;
                        }
                    }
                } catch (error) {
                    console.error(`   ❌ P&L full_year_refresh error:`, error);
                    hasError = true;
                }
            } else {
                // SMART PERIOD EXPANSION: Same as BS, include adjacent months
                const expandedPLPeriods = expandPeriodRange(periodsArray, 1, 1);
                console.log(`   📦 P&L: Using batch/balance for ${plAccounts.length} accounts (${expandedPLPeriods.length} periods, expanded from ${periodsArray.length})`);
                broadcastStatus(`Fetching P&L data...`, 60, 'info');
                
                try {
                    const response = await fetch(`${SERVER_URL}/batch/balance`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            accounts: plAccounts,
                            periods: expandedPLPeriods,  // Use expanded periods!
                            subsidiary: filters.subsidiary,
                            department: filters.department,
                            location: filters.location,
                            class: filters.classId,
                            accountingbook: filters.accountingBook || ''  // Multi-Book Accounting support
                        })
                    });
                    
                    if (response.ok) {
                        const data = await response.json();
                        const balances = data.balances || {};
                        console.log(`   ✅ P&L batch: ${Object.keys(balances).length} accounts`);
                        
                        // Cache with THIS filter group's filters
                        for (const acct in balances) {
                            if (!allBalances[acct]) allBalances[acct] = {};
                            for (const period in balances[acct]) {
                                allBalances[acct][period] = balances[acct][period];
                                const ck = getCacheKey('balance', {
                                    account: acct,
                                    fromPeriod: period,
                                    toPeriod: period,
                                    subsidiary: filters.subsidiary,
                                    department: filters.department,
                                    location: filters.location,
                                    classId: filters.classId
                                });
                                cache.balance.set(ck, balances[acct][period]);
                            }
                        }
                        
                        // Cache $0 for P&L accounts not returned (for expanded periods)
                        for (const acct of plAccounts) {
                            if (!allBalances[acct]) allBalances[acct] = {};
                            for (const period of expandedPLPeriods) {
                                if (allBalances[acct][period] === undefined) {
                                    allBalances[acct][period] = 0;
                                    const ck = getCacheKey('balance', {
                                        account: acct,
                                        fromPeriod: period,
                                        toPeriod: period,
                                        subsidiary: filters.subsidiary,
                                        department: filters.department,
                                        location: filters.location,
                                        classId: filters.classId
                                    });
                                    cache.balance.set(ck, 0);
                                }
                            }
                        }
                    } else {
                        console.error(`   ❌ P&L batch error: ${response.status}`);
                        hasError = true;
                    }
                } catch (error) {
                    console.error(`   ❌ P&L batch fetch error:`, error);
                    hasError = true;
                }
            }
        }
        
        // Ensure all requested BS accounts have values (even if 0)
        // IMPORTANT: Also cache $0 values with the normalized key so future lookups find them!
        let zeroCached = 0;
        for (const acct of bsAccounts) {
            if (!allBalances[acct]) allBalances[acct] = {};
            for (const period of periodsArray) {
                if (allBalances[acct][period] === undefined) {
                    console.log(`   💰 BS account ${acct} period ${period} = $0 (not in response)`);
                    allBalances[acct][period] = 0;
                    
                    // Cache $0 with the normalized key (fromPeriod = period, toPeriod = period)
                    // This ensures the next drag finds it in cache!
                    const ck = getCacheKey('balance', {
                        account: acct,
                        fromPeriod: period,
                        toPeriod: period,
                        subsidiary: filters.subsidiary,
                        department: filters.department,
                        location: filters.location,
                        classId: filters.classId
                    });
                    cache.balance.set(ck, 0);
                    zeroCached++;
                }
            }
        }
        if (zeroCached > 0) {
            console.log(`   💾 Cached ${zeroCached} zero-balance BS values`);
        }
        
        console.log(`   📊 Total accounts with data: ${Object.keys(allBalances).join(', ') || 'none'}`);
        
        // Track which periods had successful responses
        const successfulPeriods = new Set();
        for (const acct in allBalances) {
            for (const period in allBalances[acct]) {
                successfulPeriods.add(period);
            }
        }
        
        // STEP 6: Resolve all pending promises for THIS filter group
        for (const item of groupItems) {
            const { params, resolve, cacheKey } = item;
            const { account, fromPeriod, toPeriod } = params;
            
            const lookupPeriod = (fromPeriod && fromPeriod !== '') ? fromPeriod : toPeriod;
            
            if (allBalances[account] && allBalances[account][lookupPeriod] !== undefined) {
                const value = allBalances[account][lookupPeriod];
                console.log(`   ✅ ${account}/${lookupPeriod} = ${value}`);
                
                // Cache with the ORIGINAL request's cacheKey (includes its own filters)
                cache.balance.set(cacheKey, value);
                
                resolve(value);
                totalResolved++;
            } else if (hasError && !successfulPeriods.has(lookupPeriod)) {
                console.log(`   ❌ ${account}/${lookupPeriod} = "" (request failed)`);
                resolve('');
                totalZeros++;
            } else {
                console.log(`   💰 ${account}/${lookupPeriod} = 0 (no transactions)`);
                cache.balance.set(cacheKey, 0);
                resolve(0);
                totalZeros++;
            }
        }
    } // End of filter group loop
    
    const totalTime = ((Date.now() - batchStartTime) / 1000).toFixed(1);
    console.log(`   📊 Resolved: ${totalResolved} with values, ${totalZeros} zeros/errors`);
    console.log(`   ⏱️ TOTAL BUILD MODE TIME: ${totalTime}s`);
    
    // Calculate totals for user-friendly status message
    const requestedCells = pending.length;  // What user actually asked for
    // Note: We can't easily count total preloaded across filter groups, so just report requested cells
    
    // Broadcast completion with helpful info
    const anyError = totalZeros > 0 && totalResolved === 0;
    if (anyError) {
        broadcastStatus(`Completed with errors (${totalTime}s)`, 100, 'error');
    } else {
        // User-friendly message
        let msg = `✅ Updated ${requestedCells} cells`;
        if (groupCount > 1) {
            msg += ` (${groupCount} filter groups)`;
        }
        msg += ` (${totalTime}s)`;
        broadcastStatus(msg, 100, 'success');
    }
    // Clear status after delay
    setTimeout(clearStatus, 10000);  // Extended to 10s so user can read the helpful info
}

// Resolve ALL pending balance requests from cache (called by taskpane after cache is ready)
window.resolvePendingRequests = function() {
    console.log('🔄 RESOLVING ALL PENDING REQUESTS FROM CACHE...');
    let resolved = 0;
    let failed = 0;
    
    for (const [cacheKey, request] of Array.from(pendingRequests.balance.entries())) {
        const { params, resolve } = request;
        const { account, fromPeriod, toPeriod } = params;
        
        // For cumulative queries (empty fromPeriod), use toPeriod for lookup
        const lookupPeriod = (fromPeriod && fromPeriod !== '') ? fromPeriod : toPeriod;
        
        // Try to get value from localStorage cache
        let value = checkLocalStorageCache(account, fromPeriod, toPeriod);
        
        // Fallback to fullYearCache
        if (value === null) {
            value = checkFullYearCache(account, lookupPeriod);
        }
        
        if (value !== null) {
            resolve(value);
            cache.balance.set(cacheKey, value);
            resolved++;
        } else {
            // No value found - resolve with 0 (account has no transactions)
            resolve(0);
            failed++;
        }
        
        pendingRequests.balance.delete(cacheKey);
    }
    
    console.log(`   Resolved: ${resolved}, Not in cache (set to 0): ${failed}`);
    console.log(`   Remaining pending: ${pendingRequests.balance.size}`);
    return { resolved, failed };
};

// ============================================================================
// SHARED STORAGE CACHE - Uses localStorage for cross-context communication
// This works even when Shared Runtime is NOT active!
// ============================================================================
const STORAGE_KEY = 'netsuite_balance_cache';
const STORAGE_TIMESTAMP_KEY = 'netsuite_balance_cache_timestamp';
const STORAGE_TTL = 300000; // 5 minutes in milliseconds

// In-memory cache that can be populated via window function
// This is populated by taskpane when full_year_refresh completes
let fullYearCache = null;
let fullYearCacheTimestamp = null;

// Preload coordination - prevents formulas from making redundant queries while Prep Data is running
// Uses localStorage for cross-context communication (works between taskpane iframe and custom functions)
const PRELOAD_STATUS_KEY = 'netsuite_preload_status';
const PRELOAD_TIMESTAMP_KEY = 'netsuite_preload_timestamp';

function isPreloadInProgress() {
    try {
        const status = localStorage.getItem(PRELOAD_STATUS_KEY);
        const timestamp = localStorage.getItem(PRELOAD_TIMESTAMP_KEY);
        
        // Only 'running' means preload is in progress
        // 'complete', 'error', or anything else means done
        if (status === 'running' && timestamp) {
            // Check if preload started within last 3 minutes (avoid stale flags)
            const elapsed = Date.now() - parseInt(timestamp);
            if (elapsed < 180000) { // 3 minutes max wait
                return true;
            }
            // Stale preload flag - clear it
            console.log('⚠️ Stale preload flag detected - clearing');
            localStorage.removeItem(PRELOAD_STATUS_KEY);
        }
        return false;
    } catch (e) {
        return false;
    }
}

// Wait for preload to complete (polls localStorage)
async function waitForPreload(maxWaitMs = 120000) {
    const startTime = Date.now();
    const pollInterval = 500; // Check every 500ms
    
    while (isPreloadInProgress()) {
        if (Date.now() - startTime > maxWaitMs) {
            console.log('⏰ Preload wait timeout - proceeding with formula');
            return false; // Timeout - proceed anyway
        }
        await new Promise(r => setTimeout(r, pollInterval));
    }
    return true; // Preload completed
}

// These are called by taskpane via localStorage (cross-context compatible)
window.startPreload = function() {
    console.log('========================================');
    console.log('🔄 PRELOAD STARTED - formulas will wait for cache');
    console.log('========================================');
    try {
        localStorage.setItem(PRELOAD_STATUS_KEY, 'running');
        localStorage.setItem(PRELOAD_TIMESTAMP_KEY, Date.now().toString());
    } catch (e) {
        console.warn('Could not set preload status:', e);
    }
    return true;
};

window.finishPreload = function() {
    console.log('========================================');
    console.log('✅ PRELOAD FINISHED - formulas can proceed');
    console.log('========================================');
    try {
        localStorage.setItem(PRELOAD_STATUS_KEY, 'complete');
    } catch (e) {
        console.warn('Could not set preload status:', e);
    }
    return true;
};

// Function to populate the cache from taskpane (via Shared Runtime if available)
window.setFullYearCache = function(balances) {
    console.log('========================================');
    console.log('📦 SETTING FULL YEAR CACHE IN FUNCTIONS.JS');
    console.log(`   Accounts: ${Object.keys(balances).length}`);
    console.log('========================================');
    fullYearCache = balances;
    fullYearCacheTimestamp = Date.now();
    return true;
};

// Function to populate the account TYPE cache from taskpane
// This ensures XAVI.TYPE formulas resolve instantly from cache
window.setAccountTypeCache = function(accountTypes) {
    console.log('========================================');
    console.log('📦 SETTING ACCOUNT TYPE CACHE IN FUNCTIONS.JS');
    console.log(`   Account types: ${Object.keys(accountTypes).length}`);
    console.log('========================================');
    
    // Clear existing type cache to prevent stale data
    cache.type.clear();
    
    // Populate type cache with fresh data
    for (const acctNum in accountTypes) {
        const cacheKey = getCacheKey('type', { account: acctNum });
        cache.type.set(cacheKey, accountTypes[acctNum]);
    }
    
    console.log(`   Type cache now has ${cache.type.size} entries`);
    return true;
};

// Function to populate the account NAME (title) cache from taskpane
// This prevents 35+ parallel requests and NetSuite 429 errors!
window.setAccountNameCache = function(accountNames) {
    console.log('========================================');
    console.log('📦 SETTING ACCOUNT NAME CACHE IN FUNCTIONS.JS');
    console.log(`   Account names: ${Object.keys(accountNames).length}`);
    console.log('========================================');
    
    // Clear existing title cache to prevent stale data
    cache.title.clear();
    
    // Populate title cache with fresh data
    for (const acctNum in accountNames) {
        const cacheKey = getCacheKey('title', { account: acctNum });
        cache.title.set(cacheKey, accountNames[acctNum]);
    }
    
    console.log(`   Title cache now has ${cache.title.size} entries`);
    return true;
};

// Check localStorage for cached data - THIS WORKS!
// Structure: { "4220": { "Apr 2024": 123.45, ... }, ... }
function checkLocalStorageCache(account, period, toPeriod = null) {
    try {
        // DIAGNOSTIC: Log every localStorage check to debug cache issues
        console.log(`🔍 localStorage CHECK for ${account}/${period || toPeriod}`);
        
        const timestamp = localStorage.getItem(STORAGE_TIMESTAMP_KEY);
        if (!timestamp) {
            console.log(`   ⚠️ No timestamp in localStorage (key: ${STORAGE_TIMESTAMP_KEY})`);
            return null;
        }
        
        const cacheAge = Date.now() - parseInt(timestamp);
        const cacheAgeSeconds = Math.round(cacheAge / 1000);
        console.log(`   📅 Cache age: ${cacheAgeSeconds}s (TTL: ${STORAGE_TTL/1000}s)`);
        
        if (cacheAge > STORAGE_TTL) {
            console.log(`   ⏰ Cache EXPIRED (${cacheAgeSeconds}s > ${STORAGE_TTL/1000}s)`);
            return null;
        }
        
        const cached = localStorage.getItem(STORAGE_KEY);
        if (!cached) {
            console.log(`   ⚠️ No cached data in localStorage (key: ${STORAGE_KEY})`);
            return null;
        }
        
        const balances = JSON.parse(cached);
        const accountCount = Object.keys(balances).length;
        console.log(`   📊 Found ${accountCount} accounts in cache`);
        
        // For cumulative queries (empty fromPeriod), use toPeriod for lookup
        const lookupPeriod = (period && period !== '') ? period : toPeriod;
        
        // Debug: Log what we're looking for vs what's available
        console.log(`   🔍 Looking for: account=${account}, period="${lookupPeriod}"`);
        if (balances[account]) {
            const availablePeriods = Object.keys(balances[account]);
            console.log(`   Available periods for ${account}: ${availablePeriods.join(', ')}`);
        } else {
            console.log(`   ⚠️ Account ${account} NOT found in cache. Sample accounts: ${Object.keys(balances).slice(0, 5).join(', ')}`);
        }
        
        // ONLY return if we have an explicit value for this account+period
        // Don't assume $0 for missing periods - the query may have been truncated!
        if (lookupPeriod && balances[account] && balances[account][lookupPeriod] !== undefined) {
            console.log(`   ✅ HIT: ${balances[account][lookupPeriod]}`);
            return balances[account][lookupPeriod];
        }
        
        console.log(`   ❌ MISS: period "${lookupPeriod}" not found for account ${account}`);
        // Period not found - return null to trigger batch processing
        // (Could be missing due to 1000 row limit, not because it's truly $0)
        return null;
        
    } catch (e) {
        console.error('localStorage read error:', e);
        return null;
    }
}

// Check in-memory cache (backup for Shared Runtime)
function checkFullYearCache(account, period) {
    console.log(`🔍 fullYearCache CHECK for ${account}/${period}`);
    
    if (!fullYearCache || !fullYearCacheTimestamp) {
        console.log(`   ⚠️ fullYearCache not set (cache=${!!fullYearCache}, timestamp=${!!fullYearCacheTimestamp})`);
        return null;
    }
    
    const cacheAge = Date.now() - fullYearCacheTimestamp;
    console.log(`   📅 fullYearCache age: ${Math.round(cacheAge/1000)}s`);
    
    // Cache expires after 5 minutes
    if (cacheAge > 300000) {
        console.log(`   ⏰ fullYearCache EXPIRED`);
        fullYearCache = null;
        fullYearCacheTimestamp = null;
        return null;
    }
    
    const accountCount = Object.keys(fullYearCache).length;
    console.log(`   📊 fullYearCache has ${accountCount} accounts`);
    
    // ONLY return if we have an explicit value for this account+period
    if (fullYearCache[account] && fullYearCache[account][period] !== undefined) {
        console.log(`   ✅ fullYearCache HIT: ${fullYearCache[account][period]}`);
        return fullYearCache[account][period];
    }
    
    if (!fullYearCache[account]) {
        console.log(`   ⚠️ Account ${account} NOT in fullYearCache. Sample: ${Object.keys(fullYearCache).slice(0,5).join(', ')}`);
    } else {
        console.log(`   ⚠️ Period ${period} NOT in fullYearCache[${account}]. Available: ${Object.keys(fullYearCache[account]).join(', ')}`);
    }
    
    // Not found - return null to trigger batch processing
    return null;
}

// Save balances to localStorage (called by taskpane via window function)
window.saveBalancesToLocalStorage = function(balances) {
    try {
        console.log('💾 Saving balances to localStorage...');
        localStorage.setItem(STORAGE_KEY, JSON.stringify(balances));
        localStorage.setItem(STORAGE_TIMESTAMP_KEY, Date.now().toString());
        console.log(`✅ Saved ${Object.keys(balances).length} accounts to localStorage`);
        return true;
    } catch (e) {
        console.error('localStorage write error:', e);
        return false;
    }
};

// Also keep the window function for Shared Runtime compatibility
window.populateFrontendCache = function(balances, filters = {}) {
    console.log('========================================');
    console.log('📦 POPULATING FRONTEND CACHE');
    console.log('========================================');
    
    const subsidiary = filters.subsidiary || '';
    const department = filters.department || '';
    const location = filters.location || '';
    const classId = filters.class || '';
    
    let cacheCount = 0;
    let resolvedCount = 0;
    
    // First, populate the in-memory cache
    // CRITICAL: Use getCacheKey to ensure format matches formula lookups!
    for (const [account, periods] of Object.entries(balances)) {
        for (const [period, amount] of Object.entries(periods)) {
            const cacheKey = getCacheKey('balance', {
                account: account,
                fromPeriod: period,
                toPeriod: period,
                subsidiary: subsidiary,
                department: department,
                location: location,
                classId: classId
            });
            cache.balance.set(cacheKey, amount);
            cacheCount++;
        }
    }
    
    // Also save to localStorage for cross-context access
    window.saveBalancesToLocalStorage(balances);
    
    console.log(`✅ Cached ${cacheCount} values in frontend`);
    
    // Resolve pending promises
    console.log(`\n🔄 Checking ${pendingRequests.balance.size} pending requests...`);
    
    for (const [cacheKey, request] of Array.from(pendingRequests.balance.entries())) {
        const { account, fromPeriod } = request.params;
        let value = 0;
        
        if (balances[account] && balances[account][fromPeriod] !== undefined) {
            value = balances[account][fromPeriod];
        }
        
        console.log(`   ✅ Resolving: ${account} = ${value}`);
        try {
            request.resolve(value);
            pendingRequests.balance.delete(cacheKey);
            resolvedCount++;
        } catch (err) {
            console.error(`   ❌ Failed:`, err);
        }
    }
    
    console.log(`✅ Resolved ${resolvedCount} pending requests`);
    console.log('========================================');
    
    return { cacheCount, resolvedCount };
};

// ============================================================================
// REQUEST QUEUE - Collects requests for intelligent batching (Phase 3)
// ============================================================================
const pendingRequests = {
    balance: new Map(),    // Map<cacheKey, {params, resolve, reject}>
    budget: new Map()
};

let batchTimer = null;  // Timer reference for batching
const BATCH_DELAY = 500;           // Wait 500ms to collect multiple requests (matches build mode settle)
const CHUNK_SIZE = 50;             // Max 50 accounts per batch (balances NetSuite limits)
const MAX_PERIODS_PER_BATCH = 3;   // Max 3 periods per batch (prevents backend timeout for high-volume accounts)
const CHUNK_DELAY = 300;           // Wait 300ms between chunks (prevent rate limiting)
const MAX_RETRIES = 2;             // Retry 429 errors up to 2 times
const RETRY_DELAY = 2000;          // Wait 2s before retrying 429 errors

// ============================================================================
// UTILITY: Convert date or date serial to "Mon YYYY" format
// ============================================================================
function convertToMonthYear(value) {
    // If empty, return empty string
    if (!value || value === '') return '';
    
    // If already in "Mon YYYY" format, return as-is
    if (typeof value === 'string' && /^[A-Za-z]{3}\s+\d{4}$/.test(value.trim())) {
        return value.trim();
    }
    
    let date;
    
    // Handle different input types
    if (typeof value === 'number') {
        // Excel date serial number (days since 1/1/1900)
        // Convert to JavaScript Date
        const excelEpoch = new Date(1899, 11, 30); // Excel's epoch is Dec 30, 1899
        date = new Date(excelEpoch.getTime() + value * 24 * 60 * 60 * 1000);
    } else if (value instanceof Date) {
        // Already a Date object
        date = value;
    } else if (typeof value === 'string') {
        // Try to parse as date string
        date = new Date(value);
        if (isNaN(date.getTime())) {
            // Not a valid date, return original
            return String(value);
        }
    } else {
        // Unknown type, return original
        return String(value);
    }
    
    // Convert Date to "Mon YYYY" format
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                       'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const month = monthNames[date.getMonth()];
    const year = date.getFullYear();
    
    return `${month} ${year}`;
}

// ============================================================================
// UTILITY: Normalize account number to string
// ============================================================================
function normalizeAccountNumber(account) {
    // Excel might pass account as a number (e.g., 15000 instead of "15000-1")
    // Always convert to string and trim
    if (account === null || account === undefined) return '';
    return String(account).trim();
}

// ============================================================================
// UTILITY: Generate cache key
// ============================================================================
function getCacheKey(type, params) {
    if (type === 'title') {
        return `title:${normalizeAccountNumber(params.account)}`;
    } else if (type === 'type') {
        // FIX: Account type cache key was missing! All accounts shared '' key!
        return `type:${normalizeAccountNumber(params.account)}`;
    } else if (type === 'balance' || type === 'budget') {
        return JSON.stringify({
            type,
            account: normalizeAccountNumber(params.account),
            fromPeriod: params.fromPeriod,
            toPeriod: params.toPeriod,
            subsidiary: params.subsidiary || '',
            department: params.department || '',
            location: params.location || '',
            class: params.classId || ''
        });
    }
    return '';
}

// ============================================================================
// NAME - Get Account Name
// ============================================================================
/**
 * @customfunction NAME
 * @param {any} accountNumber The account number
 * @param {CustomFunctions.Invocation} invocation Invocation object
 * @returns {Promise<string>} Account name
 * @requiresAddress
 * @cancelable
 */
async function NAME(accountNumber, invocation) {
    const account = normalizeAccountNumber(accountNumber);
    if (!account) return '#N/A';
    
    const cacheKey = getCacheKey('title', { account });
    
    // Check in-memory cache FIRST
    if (cache.title.has(cacheKey)) {
        cacheStats.hits++;
        console.log(`⚡ CACHE HIT [title]: ${account}`);
        return cache.title.get(cacheKey);
    }
    
    // Check localStorage name cache as fallback (prevents 35+ parallel requests!)
    try {
        const nameCache = localStorage.getItem('netsuite_name_cache');
        if (nameCache) {
            const names = JSON.parse(nameCache);
            if (names[account]) {
                // Populate in-memory cache too
                cache.title.set(cacheKey, names[account]);
                cacheStats.hits++;
                console.log(`⚡ LOCALSTORAGE HIT [title]: ${account} → ${names[account]}`);
                return names[account];
            }
        }
    } catch (e) {
        console.warn('localStorage name cache read error:', e.message);
    }
    
    cacheStats.misses++;
    console.log(`📥 CACHE MISS [title]: ${account}`);
    
    // Single request - make with retry for rate limiting
    const MAX_RETRIES = 3;
    const RETRY_DELAYS = [1000, 2000, 4000]; // Exponential backoff
    
    try {
        const controller = new AbortController();
        const signal = controller.signal;
        
        // Listen for cancellation
        if (invocation) {
            invocation.onCanceled = () => {
                console.log(`Title request canceled for ${account}`);
                controller.abort();
            };
        }
        
        for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
            try {
                // Use POST to avoid exposing account numbers in URLs/logs
                const response = await fetch(`${SERVER_URL}/account/name`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ account: String(account) }),
                    signal
                });
                
                if (response.ok) {
                    const title = await response.text();
                    cache.title.set(cacheKey, title);
                    console.log(`💾 Cached title: ${account} → "${title}"`);
                    return title;
                }
                
                // Retry on 429 (rate limit) or 500 (server error from rate limit)
                if ((response.status === 429 || response.status === 500) && attempt < MAX_RETRIES) {
                    console.warn(`⏳ Title rate limited for ${account}, retry ${attempt + 1}/${MAX_RETRIES} in ${RETRY_DELAYS[attempt]}ms`);
                    await new Promise(r => setTimeout(r, RETRY_DELAYS[attempt]));
                    continue;
                }
                
                console.error(`Title API error: ${response.status}`);
                return '#N/A';
                
            } catch (fetchError) {
                if (fetchError.name === 'AbortError') throw fetchError;
                if (attempt < MAX_RETRIES) {
                    console.warn(`⏳ Title fetch error for ${account}, retry ${attempt + 1}/${MAX_RETRIES}`);
                    await new Promise(r => setTimeout(r, RETRY_DELAYS[attempt]));
                    continue;
                }
                throw fetchError;
            }
        }
        
        return '#N/A';
        
    } catch (error) {
        if (error.name === 'AbortError') {
            console.log('Title request was canceled');
            return '#N/A';
        }
        console.error('Title fetch error:', error);
        return '#N/A';
    }
}

// ============================================================================
// TYPE - Get Account Type
// ============================================================================
/**
 * @customfunction TYPE
 * @param {any} accountNumber The account number
 * @param {CustomFunctions.Invocation} invocation Invocation object
 * @returns {Promise<string>} Account type (e.g., "Income", "Expense")
 * @requiresAddress
 * @cancelable
 */
async function TYPE(accountNumber, invocation) {
    const account = normalizeAccountNumber(accountNumber);
    if (!account) return '#N/A';
    
    const cacheKey = getCacheKey('type', { account });
    
    // Check in-memory cache FIRST
    if (!cache.type) cache.type = new Map();
    if (cache.type.has(cacheKey)) {
        cacheStats.hits++;
        console.log(`⚡ CACHE HIT [type]: ${account}`);
        return cache.type.get(cacheKey);
    }
    
    // Check localStorage type cache as fallback
    try {
        const typeCache = localStorage.getItem('netsuite_type_cache');
        if (typeCache) {
            const types = JSON.parse(typeCache);
            if (types[account]) {
                // Populate in-memory cache too
                cache.type.set(cacheKey, types[account]);
                cacheStats.hits++;
                console.log(`⚡ LOCALSTORAGE HIT [type]: ${account} → ${types[account]}`);
                return types[account];
            }
        }
    } catch (e) {
        console.warn('localStorage type cache read error:', e.message);
    }
    
    cacheStats.misses++;
    console.log(`📥 CACHE MISS [type]: ${account}`);
    
    // Single request - make immediately
    try {
        const controller = new AbortController();
        const signal = controller.signal;
        
        // Listen for cancellation
        if (invocation) {
            invocation.onCanceled = () => {
                console.log(`Type request canceled for ${account}`);
                controller.abort();
            };
        }
        
        // Use POST to avoid exposing account numbers in URLs/logs
        const response = await fetch(`${SERVER_URL}/account/type`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ account: String(account) }),
            signal
        });
        if (!response.ok) {
            console.error(`Type API error: ${response.status}`);
            return '#N/A';
        }
        
        const type = await response.text();
        cache.type.set(cacheKey, type);
        console.log(`💾 Cached type: ${account} → "${type}"`);
        return type;
        
    } catch (error) {
        if (error.name === 'AbortError') {
            console.log('Type request was canceled');
            return '#N/A';
        }
        console.error('Type fetch error:', error);
        return '#N/A';
    }
}

// ============================================================================
// PARENT - Get Parent Account
// ============================================================================
/**
 * @customfunction PARENT
 * @param {any} accountNumber The account number
 * @param {CustomFunctions.Invocation} invocation Invocation object
 * @returns {Promise<string>} Parent account number
 * @requiresAddress
 * @cancelable
 */
async function PARENT(accountNumber, invocation) {
    const account = normalizeAccountNumber(accountNumber);
    if (!account) return '#N/A';
    
    const cacheKey = getCacheKey('parent', { account });
    
    // Check cache FIRST
    if (!cache.parent) cache.parent = new Map();
    if (cache.parent.has(cacheKey)) {
        cacheStats.hits++;
        console.log(`⚡ CACHE HIT [parent]: ${account}`);
        return cache.parent.get(cacheKey);
    }
    
    cacheStats.misses++;
    console.log(`📥 CACHE MISS [parent]: ${account}`);
    
    // Single request - make immediately
    try {
        const controller = new AbortController();
        const signal = controller.signal;
        
        // Listen for cancellation
        if (invocation) {
            invocation.onCanceled = () => {
                console.log(`Parent request canceled for ${account}`);
                controller.abort();
            };
        }
        
        // Use POST to avoid exposing account numbers in URLs/logs
        const response = await fetch(`${SERVER_URL}/account/parent`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ account: String(account) }),
            signal
        });
        if (!response.ok) {
            console.error(`Parent API error: ${response.status}`);
            return '#N/A';
        }
        
        const parent = await response.text();
        cache.parent.set(cacheKey, parent);
        console.log(`💾 Cached parent: ${account} → "${parent}"`);
        return parent;
        
    } catch (error) {
        if (error.name === 'AbortError') {
            console.log('Parent request was canceled');
            return '#N/A';
        }
        console.error('Parent fetch error:', error);
        return '#N/A';
    }
}

// ============================================================================
// BALANCE - Get GL Account Balance (NON-STREAMING WITH BATCHING)
// ============================================================================
/**
 * @customfunction BALANCE
 * @param {any} account Account number
 * @param {any} fromPeriod Starting period (e.g., "Jan 2025" or 1/1/2025)
 * @param {any} toPeriod Ending period (e.g., "Mar 2025" or 3/1/2025)
 * @param {any} [subsidiary] Subsidiary filter (optional)
 * @param {any} [department] Department filter (optional)
 * @param {any} [location] Location filter (optional)
 * @param {any} [classId] Class filter (optional)
 * @param {any} [accountingBook] Accounting Book ID (optional, defaults to Primary Book). For Multi-Book Accounting.
 * @returns {Promise<number>} Account balance
 * @requiresAddress
 */
async function BALANCE(account, fromPeriod, toPeriod, subsidiary, department, location, classId, accountingBook) {
    // ================================================================
    // DEBUG: Log every BALANCE call to understand what's happening
    // ================================================================
    console.log(`📥 BALANCE called: account="${account}", fromPeriod="${fromPeriod}"`);
    
    try {
        // ================================================================
        // SPECIAL COMMAND: __CLEARCACHE__ - Clear caches from taskpane
        // Usage: =XAVI.BALANCE("__CLEARCACHE__", "60032:May 2025,60032:Jun 2025", "")
        // The second parameter contains comma-separated account:period pairs to clear
        // Returns: Number of items cleared
        // ================================================================
        const rawAccount = String(account || '').trim();
        console.log(`   rawAccount="${rawAccount}", is __CLEARCACHE__: ${rawAccount === '__CLEARCACHE__'}`);
        
        if (rawAccount === '__CLEARCACHE__') {
            console.log('🔧 __CLEARCACHE__ MATCHED! Starting cache clear...');
            const itemsStr = String(fromPeriod || '').trim();
            console.log('🔧 __CLEARCACHE__ command received:', itemsStr || 'ALL');
            
            let cleared = 0;
            
            if (!itemsStr || itemsStr === 'ALL') {
                // Clear EVERYTHING - all caches including localStorage
                cleared = cache.balance.size;
                cache.balance.clear();
                cache.title.clear();
                cache.budget.clear();
                cache.type.clear();
                cache.parent.clear();
                
                if (fullYearCache) {
                    for (const k in fullYearCache) {
                        delete fullYearCache[k];
                    }
                }
                
                // CRITICAL: Also clear localStorage from this context!
                try {
                    localStorage.removeItem('netsuite_balance_cache');
                    localStorage.removeItem('netsuite_balance_cache_timestamp');
                    console.log('   ✓ Cleared localStorage (functions context)');
                } catch (e) {
                    console.warn('   ⚠️ localStorage clear failed:', e.message);
                }
                
                console.log(`🗑️ Cleared ALL caches (${cleared} balance entries)`);
            } else {
                // Clear SPECIFIC items only - parse "60032:May 2025,60032:Jun 2025" format
                const items = itemsStr.split(',').map(s => {
                    const [account, period] = s.trim().split(':');
                    return { account, period };
                });
                
                console.log(`   Clearing ${items.length} specific items...`);
                
                // Clear from localStorage (functions context)
                try {
                    const stored = localStorage.getItem('netsuite_balance_cache');
                    if (stored) {
                        const balanceData = JSON.parse(stored);
                        let modified = false;
                        
                        for (const item of items) {
                            if (balanceData[item.account] && balanceData[item.account][item.period] !== undefined) {
                                delete balanceData[item.account][item.period];
                                cleared++;
                                modified = true;
                                console.log(`   ✓ Cleared localStorage: ${item.account}/${item.period}`);
                            }
                        }
                        
                        if (modified) {
                            localStorage.setItem('netsuite_balance_cache', JSON.stringify(balanceData));
                        }
                    }
                } catch (e) {
                    console.warn('   ⚠️ localStorage parse error:', e.message);
                }
                
                // Clear from in-memory caches
                for (const item of items) {
                    // Clear from cache.balance
                    const cacheKey = getCacheKey('balance', {
                        account: item.account,
                        fromPeriod: item.period,
                        toPeriod: item.period,
                        subsidiary: '',
                        department: '',
                        location: '',
                        classId: ''
                    });
                    
                    if (cache.balance.has(cacheKey)) {
                        cache.balance.delete(cacheKey);
                        cleared++;
                        console.log(`   ✓ Cleared cache.balance: ${item.account}/${item.period}`);
                    }
                    
                    // Clear from fullYearCache (check for null AND undefined)
                    if (fullYearCache && fullYearCache[item.account]) {
                        if (fullYearCache[item.account][item.period] !== undefined) {
                            delete fullYearCache[item.account][item.period];
                            cleared++;
                            console.log(`   ✓ Cleared fullYearCache: ${item.account}/${item.period}`);
                        }
                    }
                }
                
                console.log(`🗑️ Cleared ${cleared} items from caches`);
            }
            
            return cleared;
        }
        
        // Normalize business parameters
        account = normalizeAccountNumber(account);
        
        if (!account) {
            return 0;
        }
        
        // Convert date values to "Mon YYYY" format (supports both dates and period strings)
        fromPeriod = convertToMonthYear(fromPeriod);
        toPeriod = convertToMonthYear(toPeriod);
        
        // Other parameters as strings
        subsidiary = String(subsidiary || '').trim();
        department = String(department || '').trim();
        location = String(location || '').trim();
        classId = String(classId || '').trim();
        
        // Multi-Book Accounting support - default to empty (uses Primary Book on backend)
        accountingBook = String(accountingBook || '').trim();
        
        const params = { account, fromPeriod, toPeriod, subsidiary, department, location, classId, accountingBook };
        const cacheKey = getCacheKey('balance', params);
        
        // ================================================================
        // PRELOAD COORDINATION: If Prep Data is running, wait for it
        // Uses localStorage for cross-context communication
        // ================================================================
        if (isPreloadInProgress()) {
            console.log(`⏳ Preload in progress - waiting for cache (${account}/${fromPeriod || toPeriod})`);
            await waitForPreload();
            console.log(`✅ Preload complete - checking cache`);
            
            // After preload completes, check caches - should be populated now!
            // Check in-memory cache
            if (cache.balance.has(cacheKey)) {
                console.log(`✅ Post-preload cache hit (memory): ${account}`);
                cacheStats.hits++;
                return cache.balance.get(cacheKey);
            }
            
            // Check localStorage cache
            const localStorageValue = checkLocalStorageCache(account, fromPeriod, toPeriod);
            if (localStorageValue !== null) {
                console.log(`✅ Post-preload cache hit (localStorage): ${account}`);
                cacheStats.hits++;
                cache.balance.set(cacheKey, localStorageValue);
                return localStorageValue;
            }
            
            // Check fullYearCache
            const fyValue = checkFullYearCache(account, fromPeriod || toPeriod);
            if (fyValue !== null) {
                console.log(`✅ Post-preload cache hit (fullYearCache): ${account}`);
                cacheStats.hits++;
                cache.balance.set(cacheKey, fyValue);
                return fyValue;
            }
            
            console.log(`⚠️ Post-preload cache miss - will query NetSuite: ${account}`);
        }
        
        // ================================================================
        // BUILD MODE DETECTION: Detect rapid formula creation (drag/paste)
        // More aggressive detection - lower threshold, wider time window
        // ================================================================
        const now = Date.now();
        buildModeLastEvent = now;
        
        // Count formulas created in the current window
        formulaCreationCount++;
        
        // Reset counter after inactivity
        if (formulaCountResetTimer) clearTimeout(formulaCountResetTimer);
        formulaCountResetTimer = setTimeout(() => {
            formulaCreationCount = 0;
        }, BUILD_MODE_WINDOW_MS);
        
        // Enter build mode if we see rapid formula creation
        if (!buildMode && formulaCreationCount >= BUILD_MODE_THRESHOLD) {
            console.log(`🔨 BUILD MODE: Detected ${formulaCreationCount} formulas in ${BUILD_MODE_WINDOW_MS}ms`);
            enterBuildMode();
        }
        
        // Reset the settle timer on every formula (we'll process after user stops)
        if (buildModeTimer) {
            clearTimeout(buildModeTimer);
        }
        buildModeTimer = setTimeout(() => {
            buildModeTimer = null;
            formulaCreationCount = 0;
            if (buildMode) {
                exitBuildModeAndProcess();
            }
        }, BUILD_MODE_SETTLE_MS);
        
        // ================================================================
        // CHECK FOR CACHE INVALIDATION SIGNAL (from Refresh Selected)
        // ================================================================
        const lookupPeriod = fromPeriod || toPeriod;
        const invalidateKey = 'netsuite_cache_invalidate';
        try {
            const invalidateData = localStorage.getItem(invalidateKey);
            if (invalidateData) {
                const { items, timestamp } = JSON.parse(invalidateData);
                // Only honor signals from last 30 seconds
                if (Date.now() - timestamp < 30000) {
                    const itemKey = `${account}:${lookupPeriod}`;
                    if (items && items.includes(itemKey)) {
                        console.log(`🔄 INVALIDATED: ${itemKey} - clearing from in-memory cache`);
                        // Clear this specific item from in-memory caches
                        cache.balance.delete(cacheKey);
                        if (fullYearCache && fullYearCache[account]) {
                            delete fullYearCache[account][lookupPeriod];
                        }
                        // Remove this item from the invalidation list
                        const newItems = items.filter(i => i !== itemKey);
                        if (newItems.length > 0) {
                            localStorage.setItem(invalidateKey, JSON.stringify({ items: newItems, timestamp }));
                        } else {
                            localStorage.removeItem(invalidateKey);
                        }
                    }
                } else {
                    // Stale invalidation signal - remove it
                    localStorage.removeItem(invalidateKey);
                }
            }
        } catch (e) {
            // Ignore invalidation check errors
        }
        
        // ================================================================
        // CACHE CHECKS (same priority order as before)
        // ================================================================
        
        // Check in-memory cache FIRST - return immediately if found
        if (cache.balance.has(cacheKey)) {
            cacheStats.hits++;
            return cache.balance.get(cacheKey);
        }
        
        // DEBUG: Log cache miss details to help diagnose caching issues
        console.log(`📭 CACHE MISS: ${account}/${fromPeriod || toPeriod}`);
        console.log(`   Key: ${cacheKey.substring(0, 100)}...`);
        console.log(`   Cache size: ${cache.balance.size}`);
        // Show a sample of what IS in cache for comparison
        if (cache.balance.size > 0 && cache.balance.size <= 10) {
            console.log(`   Sample cached keys:`);
            for (const k of cache.balance.keys()) {
                console.log(`     ${k.substring(0, 100)}...`);
            }
        }
        
        // Check localStorage cache (THIS WORKS - proven by user data!)
        // Pass both fromPeriod and toPeriod - for cumulative queries, we lookup by toPeriod
        const localStorageValue = checkLocalStorageCache(account, fromPeriod, toPeriod);
        if (localStorageValue !== null) {
            cacheStats.hits++;
            // Also save to in-memory cache for next time
            cache.balance.set(cacheKey, localStorageValue);
            return localStorageValue;
        }
        
        // Check in-memory full year cache (backup for Shared Runtime)
        const fullYearValue = checkFullYearCache(account, fromPeriod);
        if (fullYearValue !== null) {
            cacheStats.hits++;
            cache.balance.set(cacheKey, fullYearValue);
            return fullYearValue;
        }
        
        // ================================================================
        // BUILD MODE: Queue with Promise (will be resolved after batch)
        // We return a Promise, not 0 - this shows #BUSY briefly but ensures
        // correct values. The batch will resolve all promises at once.
        // ================================================================
        if (buildMode) {
            console.log(`🔨 BUILD MODE: Queuing ${account}/${fromPeriod}`);
            return new Promise((resolve, reject) => {
                buildModePending.push({ cacheKey, params, resolve, reject });
            });
        }
        
        // ================================================================
        // NORMAL MODE: Cache miss - add to batch queue and return Promise
        // ================================================================
        cacheStats.misses++;
        
        // In full refresh mode, queue silently (task pane will trigger processFullRefresh)
        if (!isFullRefreshMode) {
            console.log(`📥 CACHE MISS [balance]: ${account} (${fromPeriod} to ${toPeriod}) → queuing`);
        }
        
        // Return a Promise that will be resolved by the batch processor
        return new Promise((resolve, reject) => {
            console.log(`📥 QUEUED: ${account} for ${fromPeriod}`);
            
            pendingRequests.balance.set(cacheKey, {
                params,
                resolve,
                reject,
                timestamp: Date.now()
            });
            
            console.log(`   Queue size now: ${pendingRequests.balance.size}`);
            console.log(`   isFullRefreshMode: ${isFullRefreshMode}`);
            console.log(`   batchTimer exists: ${!!batchTimer}`);
            
            // In full refresh mode, DON'T start the batch timer
            // The task pane will explicitly call processFullRefresh() when ready
            if (!isFullRefreshMode) {
                // Start batch timer if not already running (Mode 1: small batches)
                if (!batchTimer) {
                    console.log(`⏱️ STARTING batch timer (${BATCH_DELAY}ms)`);
                    batchTimer = setTimeout(() => {
                        console.log('⏱️ Batch timer FIRED!');
                        batchTimer = null;
                        processBatchQueue().catch(err => {
                            console.error('❌ Batch processing error:', err);
                        });
                    }, BATCH_DELAY);
                } else {
                    console.log('   Timer already running, request will be batched');
                }
            } else {
                console.log('   Full refresh mode - NOT starting timer');
            }
        });
        
    } catch (error) {
        console.error('BALANCE error:', error);
        return 0;
    }
}

// ============================================================================
// BUDGET - Get Budget Amount (NON-STREAMING WITH BATCHING)
// ============================================================================
/**
 * @customfunction BUDGET
 * @param {any} account Account number
 * @param {any} fromPeriod Starting period (e.g., "Jan 2025" or 1/1/2025)
 * @param {any} toPeriod Ending period (e.g., "Mar 2025" or 3/1/2025)
 * @param {any} [subsidiary] Subsidiary filter (optional)
 * @param {any} [department] Department filter (optional)
 * @param {any} [location] Location filter (optional)
 * @param {any} [classId] Class filter (optional)
 * @param {any} [accountingBook] Accounting Book ID (optional, defaults to Primary Book). For Multi-Book Accounting.
 * @returns {Promise<number>} Budget amount
 * @requiresAddress
 */
async function BUDGET(account, fromPeriod, toPeriod, subsidiary, department, location, classId, accountingBook) {
    try {
        // Normalize inputs
        account = normalizeAccountNumber(account);
        
        if (!account) {
            return 0;
        }
        
        // Convert date values to "Mon YYYY" format (supports both dates and period strings)
        fromPeriod = convertToMonthYear(fromPeriod);
        toPeriod = convertToMonthYear(toPeriod);
        
        // Other parameters as strings
        subsidiary = String(subsidiary || '').trim();
        department = String(department || '').trim();
        location = String(location || '').trim();
        classId = String(classId || '').trim();
        
        // Multi-Book Accounting support - default to empty (uses Primary Book on backend)
        accountingBook = String(accountingBook || '').trim();
        
        const params = { account, fromPeriod, toPeriod, subsidiary, department, location, classId, accountingBook };
        const cacheKey = getCacheKey('budget', params);
        
        // Check cache FIRST - return immediately if found (NO @ SYMBOL!)
        if (cache.budget.has(cacheKey)) {
            cacheStats.hits++;
            return cache.budget.get(cacheKey);
        }
        
        // Cache miss - use individual fetch (budget endpoint doesn't support batching)
        cacheStats.misses++;
        console.log(`📥 CACHE MISS [budget]: ${account} (${fromPeriod} to ${toPeriod})`);
        
        try {
            const url = new URL(`${SERVER_URL}/budget`);
            url.searchParams.append('account', account);
            if (fromPeriod) url.searchParams.append('from_period', fromPeriod);
            if (toPeriod) url.searchParams.append('to_period', toPeriod);
            if (subsidiary) url.searchParams.append('subsidiary', subsidiary);
            if (department) url.searchParams.append('department', department);
            if (location) url.searchParams.append('location', location);
            if (classId) url.searchParams.append('class', classId);
            if (accountingBook) url.searchParams.append('accountingbook', accountingBook);
            
            const response = await fetch(url.toString());
            if (!response.ok) {
                const errorText = await response.text();
                console.error(`Budget API error: ${response.status}`, errorText);
                return 0;
            }
            
            const text = await response.text();
            const budget = parseFloat(text);
            const finalValue = isNaN(budget) ? 0 : budget;
            
            // Cache the result
            cache.budget.set(cacheKey, finalValue);
            console.log(`💾 Cached budget: ${account} → ${finalValue}`);
            
            return finalValue;
            
        } catch (error) {
            console.error('Budget fetch error:', error);
            return 0;
        }
        
    } catch (error) {
        console.error('BUDGET error:', error);
        return 0;
    }
}

// ============================================================================
// BATCH PROCESSING - Non-Streaming with Promise Resolution (Phase 3)
// ============================================================================
// ============================================================================
// FULL REFRESH PROCESSOR - ONE big query for ALL accounts
// ============================================================================
async function processFullRefresh() {
    console.log('');
    console.log('════════════════════════════════════════════════════════════════════');
    console.log('🚀 PROCESSING FULL REFRESH');
    console.log('════════════════════════════════════════════════════════════════════');
    
    const allRequests = Array.from(pendingRequests.balance.entries());
    
    if (allRequests.length === 0) {
        console.log('⚠️  No requests to process');
        window.exitFullRefreshMode();
        return;
    }
    
    // Extract year from requests (or use provided year)
    let year = fullRefreshYear;
    if (!year && allRequests.length > 0) {
        const firstPeriod = allRequests[0][1].params.fromPeriod;
        if (firstPeriod) {
            const match = firstPeriod.match(/\d{4}/);
            year = match ? parseInt(match[0]) : new Date().getFullYear();
        } else {
            year = new Date().getFullYear();
        }
    }
    
    // Get filters from first request (assume all same filters)
    const filters = {};
    if (allRequests.length > 0) {
        const firstRequest = allRequests[0][1];
        filters.subsidiary = firstRequest.params.subsidiary || '';
        filters.department = firstRequest.params.department || '';
        filters.location = firstRequest.params.location || '';
        filters.class = firstRequest.params.classId || '';
        filters.accountingbook = firstRequest.params.accountingBook || '';  // Multi-Book Accounting support
    }
    
    console.log(`📊 Full Refresh Request:`);
    console.log(`   Formulas: ${allRequests.length}`);
    console.log(`   Year: ${year}`);
    console.log(`   Filters:`, filters);
    console.log('');
    
    try {
        // Call optimized backend endpoint
        const payload = {
            year: year,
            ...filters
        };
        
        console.log('📤 Fetching ALL accounts for entire year...');
        const start = Date.now();
        
        const response = await fetch(`${SERVER_URL}/batch/full_year_refresh`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${await response.text()}`);
        }
        
        const data = await response.json();
        const balances = data.balances || {};
        const queryTime = data.query_time || 0;
        const elapsed = ((Date.now() - start) / 1000).toFixed(2);
        
        console.log('');
        console.log(`✅ DATA RECEIVED`);
        console.log(`   Backend Query Time: ${queryTime.toFixed(2)}s`);
        console.log(`   Total Time: ${elapsed}s`);
        console.log(`   Accounts: ${Object.keys(balances).length}`);
        console.log('');
        
        // Populate cache with ALL results
        console.log('💾 Populating cache...');
        let cachedCount = 0;
        for (const account in balances) {
            for (const period in balances[account]) {
                // Create cache key for this account-period combination
                const cacheKey = getCacheKey('balance', {
                    account: account,
                    fromPeriod: period,
                    toPeriod: period,
                    ...filters
                });
                cache.balance.set(cacheKey, balances[account][period]);
                cachedCount++;
            }
        }
        console.log(`   Cached ${cachedCount} account-period combinations`);
        console.log('');
        
        // Resolve ALL pending requests from cache
        console.log('📝 Resolving formulas...');
        let resolvedCount = 0;
        let errorCount = 0;
        
        for (const [cacheKey, request] of allRequests) {
            try {
                const account = request.params.account;
                const fromPeriod = request.params.fromPeriod;
                const toPeriod = request.params.toPeriod;
                
                // Sum requested period range
                let total = 0;
                if (fromPeriod === toPeriod) {
                    // Single period
                    if (balances[account] && balances[account][fromPeriod] !== undefined) {
                        total = balances[account][fromPeriod];
                    }
                } else {
                    // Multiple periods - sum them
                    const periodRange = expandPeriodRange(fromPeriod, toPeriod);
                    for (const period of periodRange) {
                        if (balances[account] && balances[account][period] !== undefined) {
                            total += balances[account][period];
                        }
                    }
                }
                
                request.resolve(total);
                resolvedCount++;
                
            } catch (error) {
                console.error(`❌ Error resolving ${request.params.account}:`, error);
                request.reject(error);
                errorCount++;
            }
        }
        
        console.log(`   ✅ Resolved: ${resolvedCount} formulas`);
        if (errorCount > 0) {
            console.log(`   ❌ Errors: ${errorCount} formulas`);
        }
        console.log('');
        console.log('════════════════════════════════════════════════════════════════════');
        console.log(`✅ FULL REFRESH COMPLETE (${elapsed}s)`);
        console.log('════════════════════════════════════════════════════════════════════');
        console.log('');
        
        pendingRequests.balance.clear();
        
    } catch (error) {
        console.error('');
        console.error('════════════════════════════════════════════════════════════════════');
        console.error('❌ FULL REFRESH FAILED');
        console.error('════════════════════════════════════════════════════════════════════');
        console.error(error);
        console.error('');
        
        // Reject all pending requests
        for (const [cacheKey, request] of allRequests) {
            request.reject(error);
        }
        
        pendingRequests.balance.clear();
        
    } finally {
        window.exitFullRefreshMode();
    }
}

// Make it globally accessible for taskpane
window.processFullRefresh = processFullRefresh;

async function processBatchQueue() {
    const batchStartTime = Date.now();
    batchTimer = null;  // Reset timer reference
    
    console.log('========================================');
    console.log(`🔄 processBatchQueue() CALLED at ${new Date().toLocaleTimeString()}`);
    console.log('========================================');
    
    // CHECK: If build mode was entered, defer to it instead
    // This handles the race condition where timer fires just as build mode starts
    if (buildMode) {
        console.log('⏸️ Build mode is active - deferring to build mode batch');
        // Move any pending requests to build mode queue
        for (const [cacheKey, request] of pendingRequests.balance.entries()) {
            buildModePending.push({
                cacheKey,
                params: request.params,
                resolve: request.resolve,
                reject: request.reject
            });
        }
        if (pendingRequests.balance.size > 0) {
            console.log(`   📦 Moved ${pendingRequests.balance.size} requests to build mode`);
            pendingRequests.balance.clear();
        }
        return; // Let build mode handle everything
    }
    
    if (pendingRequests.balance.size === 0) {
        console.log('❌ No balance requests in queue - exiting');
        return;
    }
    
    const requestCount = pendingRequests.balance.size;
    console.log(`✅ Found ${requestCount} pending requests`);
    console.log(`📊 Cache stats: ${cacheStats.hits} hits / ${cacheStats.misses} misses`);
    
    // Extract requests and clear queue
    const requests = Array.from(pendingRequests.balance.entries());
    pendingRequests.balance.clear();
    
    // Group by filters ONLY (not periods) - this allows smart batching
    // Example: 1 account × 12 months = 1 batch (not 12 batches)
    // Example: 100 accounts × 1 month = 2 batches (chunked by accounts)
    // Example: 100 accounts × 12 months = 2 batches (all periods together)
    const groups = new Map();
    for (const [cacheKey, request] of requests) {
        const {params} = request;
        const filterKey = JSON.stringify({
            subsidiary: params.subsidiary || '',
            department: params.department || '',
            location: params.location || '',
            class: params.classId || ''
            // Note: NOT grouping by periods - this is the key optimization!
        });
        
        if (!groups.has(filterKey)) {
            groups.set(filterKey, []);
        }
        groups.get(filterKey).push({ cacheKey, request });
    }
    
    console.log(`📦 Grouped into ${groups.size} batch(es) by filters only`);
    
    // Process each group
    for (const [filterKey, groupRequests] of groups.entries()) {
        const filters = JSON.parse(filterKey);
        const accounts = [...new Set(groupRequests.map(r => r.request.params.account))];
        
        // Collect ALL unique periods from ALL requests in this group
        // This allows us to fetch multiple periods in one API call
        const allPeriods = groupRequests.flatMap(r => [r.request.params.fromPeriod, r.request.params.toPeriod]);
        const periods = [...new Set(allPeriods.filter(p => p))];
        
        console.log(`  Batch: ${accounts.length} accounts × ${periods.length} period(s)`);
        
        // Split into chunks to avoid overwhelming NetSuite
        // Chunk by BOTH accounts AND periods to prevent backend timeouts
        const accountChunks = [];
        for (let i = 0; i < accounts.length; i += CHUNK_SIZE) {
            accountChunks.push(accounts.slice(i, i + CHUNK_SIZE));
        }
        
        const periodChunks = [];
        for (let i = 0; i < periods.length; i += MAX_PERIODS_PER_BATCH) {
            periodChunks.push(periods.slice(i, i + MAX_PERIODS_PER_BATCH));
        }
        
        console.log(`  Split into ${accountChunks.length} account chunk(s) × ${periodChunks.length} period chunk(s) = ${accountChunks.length * periodChunks.length} total batches`);
        
        // Track which requests have been resolved to avoid double-resolution
        const resolvedRequests = new Set();
        
        // For each request, track which period chunks need to be processed
        // and accumulate the total across chunks
        const requestAccumulators = new Map();
        for (const {cacheKey, request} of groupRequests) {
            requestAccumulators.set(cacheKey, {
                total: 0,
                periodsNeeded: new Set([request.params.fromPeriod, request.params.toPeriod].filter(p => p)),
                periodsProcessed: new Set()
            });
        }
        
        // Process chunks sequentially (both accounts AND periods)
        let chunkIndex = 0;
        const totalChunks = accountChunks.length * periodChunks.length;
        
        for (let ai = 0; ai < accountChunks.length; ai++) {
            for (let pi = 0; pi < periodChunks.length; pi++) {
                chunkIndex++;
                const accountChunk = accountChunks[ai];
                const periodChunk = periodChunks[pi];
                const chunkStartTime = Date.now();
                console.log(`  📤 Chunk ${chunkIndex}/${totalChunks}: ${accountChunk.length} accounts × ${periodChunk.length} periods (fetching...)`);
            
                try {
                    // Make batch API call
                    const response = await fetch(`${SERVER_URL}/batch/balance`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            accounts: accountChunk,
                            periods: periodChunk,
                            subsidiary: filters.subsidiary || '',
                            department: filters.department || '',
                            location: filters.location || '',
                            class: filters.class || '',
                            accountingbook: filters.accountingBook || ''  // Multi-Book Accounting support
                        })
                    });
                
                    if (!response.ok) {
                        console.error(`  ❌ API error: ${response.status}`);
                        // Reject all promises in this chunk
                        for (const {cacheKey, request} of groupRequests) {
                            if (accountChunk.includes(request.params.account) && !resolvedRequests.has(cacheKey)) {
                                request.reject(new Error(`API error: ${response.status}`));
                                resolvedRequests.add(cacheKey);
                            }
                        }
                        continue;
                    }
                
                const data = await response.json();
                const balances = data.balances || {};
                const chunkTime = ((Date.now() - chunkStartTime) / 1000).toFixed(1);
                
                console.log(`  ✅ Received data for ${Object.keys(balances).length} accounts in ${chunkTime}s`);
                console.log(`  📦 Raw response:`, JSON.stringify(data, null, 2).substring(0, 500));
                console.log(`  📦 Balances object:`, JSON.stringify(balances, null, 2).substring(0, 500));
                
                    // Distribute results to waiting Promises
                    for (const {cacheKey, request} of groupRequests) {
                        // Skip if already resolved
                        if (resolvedRequests.has(cacheKey)) {
                            continue;
                        }
                        
                        const account = request.params.account;
                        
                        // Only process accounts in this chunk
                        if (!accountChunk.includes(account)) {
                            continue;
                        }
                        
                        const fromPeriod = request.params.fromPeriod;
                        const toPeriod = request.params.toPeriod;
                        const accountBalances = balances[account] || {};
                        const accum = requestAccumulators.get(cacheKey);
                        
                        console.log(`    🔍 Account ${account}: accountBalances =`, JSON.stringify(accountBalances));
                        
                        // Process each period in this chunk that this request needs
                        for (const period of periodChunk) {
                            // Check if this request needs this period
                            if (accum.periodsNeeded.has(period) && !accum.periodsProcessed.has(period)) {
                                accum.total += accountBalances[period] || 0;
                                accum.periodsProcessed.add(period);
                            }
                        }
                        
                        // Cache the accumulated result
                        cache.balance.set(cacheKey, accum.total);
                        
                        // Check if all needed periods are now processed
                        const allPeriodsProcessed = [...accum.periodsNeeded].every(p => accum.periodsProcessed.has(p));
                        
                        if (allPeriodsProcessed) {
                            console.log(`    🎯 RESOLVING: ${account} = ${accum.total}`);
                            try {
                                request.resolve(accum.total);
                                console.log(`    ✅ RESOLVED: ${account}`);
                            } catch (resolveErr) {
                                console.error(`    ❌ RESOLVE ERROR for ${account}:`, resolveErr);
                            }
                            resolvedRequests.add(cacheKey);
                        }
                    }
                
                } catch (error) {
                    console.error(`  ❌ Fetch error:`, error);
                    // Reject all promises in this chunk
                    for (const {cacheKey, request} of groupRequests) {
                        if (accountChunk.includes(request.params.account) && !resolvedRequests.has(cacheKey)) {
                            request.reject(error);
                            resolvedRequests.add(cacheKey);
                        }
                    }
                }
                
                // Delay between chunks to avoid rate limiting
                if (chunkIndex < totalChunks) {
                    console.log(`  ⏱️  Waiting ${CHUNK_DELAY}ms before next chunk...`);
                    await new Promise(resolve => setTimeout(resolve, CHUNK_DELAY));
                }
            }
        }
        
        // CRITICAL: Resolve any remaining unresolved requests with their accumulated totals
        // This catches edge cases where periods didn't align perfectly with chunks
        console.log(`\n🔍 Checking for unresolved requests (resolved so far: ${resolvedRequests.size}/${groupRequests.length})`);
        
        let unresolvedCount = 0;
        for (const {cacheKey, request} of groupRequests) {
            if (!resolvedRequests.has(cacheKey)) {
                const accum = requestAccumulators.get(cacheKey);
                console.log(`  ⚠️ FORCE-RESOLVING: ${request.params.account} = ${accum.total}`);
                console.log(`     periodsNeeded: ${[...accum.periodsNeeded].join(', ')}`);
                console.log(`     periodsProcessed: ${[...accum.periodsProcessed].join(', ')}`);
                try {
                    request.resolve(accum.total);
                    console.log(`  ✅ Force-resolved successfully`);
                } catch (err) {
                    console.error(`  ❌ Force-resolve FAILED:`, err);
                }
                resolvedRequests.add(cacheKey);
                unresolvedCount++;
            }
        }
        
        console.log(`📊 Final stats: ${resolvedRequests.size} resolved, ${unresolvedCount} force-resolved`);
    }
    
    const totalBatchTime = ((Date.now() - batchStartTime) / 1000).toFixed(1);
    console.log('========================================');
    console.log(`✅ BATCH PROCESSING COMPLETE in ${totalBatchTime}s`);
    console.log('========================================\n');
}

// ============================================================================
// OLD STREAMING CODE - REMOVED (kept for reference)
// ============================================================================
/*
async function fetchBatchBalances(accounts, periods, filters, allRequests, retryCount = 0) {
    try {
        const payload = {
            accounts,
            periods,
            subsidiary: filters.subsidiary || '',
            department: filters.department || '',
            location: filters.location || '',
            class: filters.class || ''
        };
        
        const response = await fetch(`${SERVER_URL}/batch/balance`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        
        if (response.status === 429) {
            // NetSuite concurrency limit hit
            console.warn(`⚠️  429 ERROR: NetSuite concurrency limit (retry ${retryCount + 1}/${MAX_RETRIES})`);
            
            if (retryCount < MAX_RETRIES) {
                console.log(`  Waiting ${RETRY_DELAY}ms before retry...`);
                await delay(RETRY_DELAY);
                return await fetchBatchBalances(accounts, periods, filters, allRequests, retryCount + 1);
            } else {
                console.error(`  ❌ Max retries reached, returning blanks`);
                // Finish all invocations with 0
                for (const { key, req } of allRequests) {
                    if (accounts.includes(req.params.account)) {
                        safeFinishInvocation(req.invocation, 0);
                    }
                }
                return;
            }
        }
        
        if (!response.ok) {
            console.error(`Batch API error: ${response.status}`);
            // Finish all invocations with 0
            for (const { key, req } of allRequests) {
                if (accounts.includes(req.params.account)) {
                    safeFinishInvocation(req.invocation, 0);
                }
            }
            return;
        }
        
        const data = await response.json();
        const balances = data.balances || {};
        
        console.log(`  ✅ Received balances for ${Object.keys(balances).length} accounts`);
        
        // Track which invocations we've successfully finished
        // This prevents closing them again with 0 if there's an error later
        const finishedInvocations = new Set();
        
        // Distribute results to invocations and close them
        // Backend now returns period-by-period breakdown
        // Each cell extracts and sums ONLY the periods it requested
        for (const { key, req } of allRequests) {
            try {
                // ✅ CRITICAL FIX: Only process accounts that are in THIS batch
                // Don't finish invocations for accounts not in this batch - they're in other batches!
                if (!accounts.includes(req.params.account)) {
                    console.log(`ℹ️  Account ${req.params.account} not in this batch, skipping...`);
                    continue;  // Leave invocation open for next batch
                }
                
                const accountBalances = balances[req.params.account] || {};
                
                // Expand THIS cell's period range
                const cellPeriods = expandPeriodRange(req.params.fromPeriod, req.params.toPeriod);
                
                // Sum only the periods THIS cell requested
                let total = 0;
                for (const period of cellPeriods) {
                    total += accountBalances[period] || 0;
                }
                
                // Cache the result and finish the invocation
                cache.balance.set(key, total);
                console.log(`💾 Cached ${req.params.account} (${cellPeriods.join(', ')}): ${total}`);
                console.log(`   → Finishing invocation for ${req.params.account}:`, {
                    hasInvocation: !!req.invocation,
                    hasSetResult: !!(req.invocation && req.invocation.setResult),
                    hasClose: !!(req.invocation && req.invocation.close),
                    total: total
                });
                safeFinishInvocation(req.invocation, total);
                finishedInvocations.add(key);  // Mark as finished
                
            } catch (error) {
                console.error('Error distributing result:', error, key);
                // ❌ DO NOT cache 0 on error - this causes cached failures!
                // Just finish the invocation with 0, don't pollute cache
                safeFinishInvocation(req.invocation, 0);
                finishedInvocations.add(key);  // Mark as finished (even with 0)
            }
        }
        
    } catch (error) {
        console.error('❌ Batch fetch error:', error);
        // DEFENSIVE: Only close invocations that we HAVEN'T already finished
        // This prevents overwriting correct values with 0!
        console.log(`⚠️  Closing unfinished invocations due to error...`);
        for (const { key, req } of allRequests) {
            try {
                // ✅ CRITICAL FIX: Only close if we haven't finished it yet
                if (req.invocation && !finishedInvocations.has(key)) {
                    console.log(`  → Closing unfinished invocation for ${req.params.account} with 0`);
                    safeFinishInvocation(req.invocation, 0);
                    // Mark as closed in tracker
                    if (invocationTracker.has(key)) {
                        invocationTracker.get(key).closed = true;
                    }
                }
            } catch (closeError) {
                console.error('Error closing invocation:', closeError);
            }
        }
    }
}

// ============================================================================
// HELPER: Expand period range (e.g., "Jan 2025" to "Mar 2025" → all months)
// ============================================================================
function expandPeriodRange(fromPeriod, toPeriod) {
    if (!fromPeriod) {
        return [];
    }
    if (!toPeriod || fromPeriod === toPeriod) {
        return [fromPeriod];
    }
    
    try {
        // Parse month and year from "Jan 2025" format
        const parseMonthYear = (period) => {
            const match = period.match(/^([A-Za-z]+)\s+(\d{4})$/);
            if (!match) return null;
            const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
            const monthIndex = monthNames.findIndex(m => m === match[1]);
            if (monthIndex === -1) return null;
            return { month: monthIndex, year: parseInt(match[2]) };
        };
        
        const from = parseMonthYear(fromPeriod);
        const to = parseMonthYear(toPeriod);
        
        if (!from || !to) {
            // Can't parse - return original periods
            return [fromPeriod, toPeriod];
        }
        
        // Generate all months in range
        const result = [];
        const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                           'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        
        let currentMonth = from.month;
        let currentYear = from.year;
        
        while (currentYear < to.year || (currentYear === to.year && currentMonth <= to.month)) {
            result.push(`${monthNames[currentMonth]} ${currentYear}`);
            currentMonth++;
            if (currentMonth > 11) {
                currentMonth = 0;
                currentYear++;
            }
        }
        
        return result;
        
    } catch (error) {
        console.error('Error expanding period range:', error);
        return [fromPeriod, toPeriod];
    }
}
*/

// (Old streaming functions removed - not needed for Phase 3 non-streaming async)

// ============================================================================
// RETAINEDEARNINGS - Calculate prior years' cumulative P&L (no account number)
// NetSuite calculates this dynamically at report runtime
// ============================================================================
/**
 * Get calculated retained earnings (prior years' cumulative P&L).
 * NetSuite calculates this dynamically - there is no account number to query.
 * RE = Sum of all P&L from inception through prior fiscal year end + posted RE adjustments.
 * 
 * @customfunction RETAINEDEARNINGS
 * @param {any} period Accounting period (e.g., "Mar 2025")
 * @param {any} [subsidiary] Subsidiary ID (optional)
 * @param {any} [accountingBook] Accounting Book ID (optional, defaults to Primary Book)
 * @param {any} [classId] Class filter (optional)
 * @param {any} [department] Department filter (optional)
 * @param {any} [location] Location filter (optional)
 * @returns {Promise<number>} Retained earnings value
 */
async function RETAINEDEARNINGS(period, subsidiary, accountingBook, classId, department, location) {
    try {
        // Convert date values to "Mon YYYY" format
        period = convertToMonthYear(period);
        
        if (!period) {
            console.error('RETAINEDEARNINGS: period is required');
            return 0;
        }
        
        // Normalize optional parameters
        subsidiary = String(subsidiary || '').trim();
        accountingBook = String(accountingBook || '').trim();
        classId = String(classId || '').trim();
        department = String(department || '').trim();
        location = String(location || '').trim();
        
        // Build cache key
        const cacheKey = `retainedearnings:${period}:${subsidiary}:${accountingBook}:${classId}:${department}:${location}`;
        
        // Check cache first
        if (cache.balance.has(cacheKey)) {
            cacheStats.hits++;
            console.log(`📥 CACHE HIT [retained earnings]: ${period}`);
            return cache.balance.get(cacheKey);
        }
        
        // Check if there's already a request in-flight for this exact key
        // This prevents duplicate API calls when Excel evaluates the formula multiple times
        if (inFlightRequests.has(cacheKey)) {
            console.log(`⏳ Waiting for in-flight request [retained earnings]: ${period}`);
            return await inFlightRequests.get(cacheKey);
        }
        
        cacheStats.misses++;
        console.log(`📥 Calculating Retained Earnings for ${period}...`);
        
        // Create the promise and store it BEFORE awaiting
        const requestPromise = (async () => {
            try {
                const response = await fetch(`${SERVER_URL}/retained-earnings`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        period,
                        subsidiary,
                        accountingBook,
                        classId,
                        department,
                        location
                    })
                });
                
                if (!response.ok) {
                    const errorText = await response.text();
                    console.error(`Retained Earnings API error: ${response.status}`, errorText);
                    return 0;
                }
                
                const data = await response.json();
                const value = parseFloat(data.value) || 0;
                
                // Cache the result
                cache.balance.set(cacheKey, value);
                console.log(`✅ Retained Earnings (${period}): ${value.toLocaleString()}`);
                
                return value;
                
            } catch (error) {
                console.error('Retained Earnings fetch error:', error);
                return 0;
            } finally {
                // Remove from in-flight after completion
                inFlightRequests.delete(cacheKey);
            }
        })();
        
        // Store the promise for deduplication
        inFlightRequests.set(cacheKey, requestPromise);
        
        return await requestPromise;
        
    } catch (error) {
        console.error('RETAINEDEARNINGS error:', error);
        return 0;
    }
}

// ============================================================================
// NETINCOME - Calculate current fiscal year net income (no account number)
// NetSuite calculates this dynamically at report runtime
// ============================================================================
/**
 * Get current fiscal year net income through target period.
 * NetSuite calculates this dynamically - there is no account number to query.
 * NI = Sum of all P&L from fiscal year start through target period end.
 * 
 * @customfunction NETINCOME
 * @param {any} period Accounting period (e.g., "Mar 2025")
 * @param {any} [subsidiary] Subsidiary ID (optional)
 * @param {any} [accountingBook] Accounting Book ID (optional, defaults to Primary Book)
 * @param {any} [classId] Class filter (optional)
 * @param {any} [department] Department filter (optional)
 * @param {any} [location] Location filter (optional)
 * @returns {Promise<number>} Net income value
 */
async function NETINCOME(period, subsidiary, accountingBook, classId, department, location) {
    try {
        // Convert date values to "Mon YYYY" format
        period = convertToMonthYear(period);
        
        if (!period) {
            console.error('NETINCOME: period is required');
            return 0;
        }
        
        // Normalize optional parameters
        subsidiary = String(subsidiary || '').trim();
        accountingBook = String(accountingBook || '').trim();
        classId = String(classId || '').trim();
        department = String(department || '').trim();
        location = String(location || '').trim();
        
        // Build cache key
        const cacheKey = `netincome:${period}:${subsidiary}:${accountingBook}:${classId}:${department}:${location}`;
        
        // Check cache first
        if (cache.balance.has(cacheKey)) {
            cacheStats.hits++;
            console.log(`📥 CACHE HIT [net income]: ${period}`);
            return cache.balance.get(cacheKey);
        }
        
        // Check if there's already a request in-flight for this exact key
        if (inFlightRequests.has(cacheKey)) {
            console.log(`⏳ Waiting for in-flight request [net income]: ${period}`);
            return await inFlightRequests.get(cacheKey);
        }
        
        cacheStats.misses++;
        console.log(`📥 Calculating Net Income for ${period}...`);
        
        // Create the promise and store it BEFORE awaiting
        const requestPromise = (async () => {
            try {
                const response = await fetch(`${SERVER_URL}/net-income`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        period,
                        subsidiary,
                        accountingBook,
                        classId,
                        department,
                        location
                    })
                });
                
                if (!response.ok) {
                    const errorText = await response.text();
                    console.error(`Net Income API error: ${response.status}`, errorText);
                    return 0;
                }
                
                const data = await response.json();
                const value = parseFloat(data.value) || 0;
                
                // Cache the result
                cache.balance.set(cacheKey, value);
                console.log(`✅ Net Income (${period}): ${value.toLocaleString()}`);
                
                return value;
                
            } catch (error) {
                console.error('Net Income fetch error:', error);
                return 0;
            } finally {
                inFlightRequests.delete(cacheKey);
            }
        })();
        
        inFlightRequests.set(cacheKey, requestPromise);
        return await requestPromise;
        
    } catch (error) {
        console.error('NETINCOME error:', error);
        return 0;
    }
}

// ============================================================================
// CTA - Calculate Cumulative Translation Adjustment (multi-currency plug)
// This is the balancing figure after currency translation in consolidation
// ============================================================================
/**
 * Get cumulative translation adjustment for consolidated multi-currency reports.
 * This is a "plug" figure that forces the Balance Sheet to balance after currency translation.
 * Note: CTA omits segment filters because translation adjustments apply at entity level.
 * 
 * @customfunction CTA
 * @param {any} period Accounting period (e.g., "Mar 2025")
 * @param {any} [subsidiary] Subsidiary ID (optional)
 * @param {any} [accountingBook] Accounting Book ID (optional, defaults to Primary Book)
 * @returns {Promise<number>} CTA value
 */
async function CTA(period, subsidiary, accountingBook) {
    try {
        // Convert date values to "Mon YYYY" format
        period = convertToMonthYear(period);
        
        if (!period) {
            console.error('CTA: period is required');
            return 0;
        }
        
        // Normalize optional parameters
        subsidiary = String(subsidiary || '').trim();
        accountingBook = String(accountingBook || '').trim();
        
        // Build cache key (no segment filters for CTA - entity level only)
        const cacheKey = `cta:${period}:${subsidiary}:${accountingBook}`;
        
        // Check cache first
        if (cache.balance.has(cacheKey)) {
            cacheStats.hits++;
            console.log(`📥 CACHE HIT [CTA]: ${period}`);
            return cache.balance.get(cacheKey);
        }
        
        // Check if there's already a request in-flight for this exact key
        if (inFlightRequests.has(cacheKey)) {
            console.log(`⏳ Waiting for in-flight request [CTA]: ${period}`);
            return await inFlightRequests.get(cacheKey);
        }
        
        cacheStats.misses++;
        console.log(`📥 Calculating CTA for ${period}...`);
        
        // Create the promise and store it BEFORE awaiting
        const requestPromise = (async () => {
            try {
                const response = await fetch(`${SERVER_URL}/cta`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        period,
                        subsidiary,
                        accountingBook
                    })
                });
                
                if (!response.ok) {
                    const errorText = await response.text();
                    console.error(`CTA API error: ${response.status}`, errorText);
                    return 0;
                }
                
                const data = await response.json();
                const value = parseFloat(data.value) || 0;
                
                // Cache the result
                cache.balance.set(cacheKey, value);
                console.log(`✅ CTA (${period}): ${value.toLocaleString()}`);
                
                return value;
                
            } catch (error) {
                console.error('CTA fetch error:', error);
                return 0;
            } finally {
                inFlightRequests.delete(cacheKey);
            }
        })();
        
        inFlightRequests.set(cacheKey, requestPromise);
        return await requestPromise;
        
    } catch (error) {
        console.error('CTA error:', error);
        return 0;
    }
}

// ============================================================================
// CLEARCACHE - Internal function to clear caches from taskpane
// Called via Excel.evaluate("=NS.CLEARCACHE(items)") from taskpane
// ============================================================================
// Track last CLEARCACHE time to prevent repeated clearing during formula evaluation
let lastClearCacheTime = 0;
const CLEARCACHE_DEBOUNCE_MS = 5000; // 5 second debounce

/**
 * Internal function - clears specified items from in-memory cache
 * @customfunction CLEARCACHE
 * @param {string} [itemsJson] JSON string of items to clear, or empty for all
 * @returns {string} Status message
 */
function CLEARCACHE(itemsJson) {
    console.log('🔧 CLEARCACHE called with:', itemsJson);
    
    try {
        // IMPORTANT: Only clear ALL caches when explicitly requested with "ALL"
        // This prevents accidental cache clearing during calculations
        if (itemsJson === 'ALL') {
            // DEBOUNCE: Prevent repeated "ALL" clears within 5 seconds
            // This happens when Excel re-evaluates =XAVI.CLEARCACHE("ALL") during formula calculations
            const now = Date.now();
            if (now - lastClearCacheTime < CLEARCACHE_DEBOUNCE_MS) {
                console.log(`⚠️ CLEARCACHE("ALL") debounced - last clear was ${Math.round((now - lastClearCacheTime)/1000)}s ago`);
                return 'DEBOUNCED';
            }
            lastClearCacheTime = now;
            
            // Clear ALL caches - explicit request only
            cache.balance.clear();
            cache.title.clear();
            cache.budget.clear();
            cache.type.clear();
            cache.parent.clear();
            if (fullYearCache) {
                Object.keys(fullYearCache).forEach(k => delete fullYearCache[k]);
            }
            console.log('🗑️ Cleared ALL in-memory caches (explicit ALL request)');
            return 'CLEARED_ALL';
        } else if (!itemsJson || itemsJson === '' || itemsJson === null) {
            // Empty/null call - do nothing (prevents accidental clearing)
            console.log('⚠️ CLEARCACHE called with empty/null - ignoring (use "ALL" to clear everything)');
            return 'IGNORED';
        } else {
            // Clear specific items
            const items = JSON.parse(itemsJson);
            let cleared = 0;
            
            for (const item of items) {
                const account = String(item.account);
                const period = item.period;
                
                // Use getCacheKey to ensure exact same format as BALANCE
                // Key order MUST match: type, account, fromPeriod, toPeriod, subsidiary, department, location, class
                const exactKey = getCacheKey('balance', {
                    account: account,
                    fromPeriod: period,
                    toPeriod: period,
                    subsidiary: '',
                    department: '',
                    location: '',
                    classId: ''
                });
                
                console.log(`   🔍 Looking for key: ${exactKey.substring(0, 80)}...`);
                
                if (cache.balance.has(exactKey)) {
                    cache.balance.delete(exactKey);
                    cleared++;
                    console.log(`   ✓ Cleared cache.balance: ${account}/${period}`);
                } else {
                    console.log(`   ⚠️ Key not found in cache.balance`);
                }
                
                // Clear from fullYearCache
                if (fullYearCache && fullYearCache[account]) {
                    if (fullYearCache[account][period] !== undefined) {
                        delete fullYearCache[account][period];
                        cleared++;
                        console.log(`   ✓ Cleared fullYearCache: ${account}/${period}`);
                    }
                }
            }
            
            console.log(`🗑️ Cleared ${cleared} items from in-memory cache`);
            return `CLEARED_${cleared}`;
        }
    } catch (e) {
        console.error('CLEARCACHE error:', e);
        return 'ERROR';
    }
}

// ============================================================================
// REGISTER FUNCTIONS WITH EXCEL
// ============================================================================
// CRITICAL: The manifest ALREADY defines namespace 'NS'
// We just register individual functions - Excel adds the XAVI. prefix automatically!
if (typeof CustomFunctions !== 'undefined') {
    CustomFunctions.associate('NAME', NAME);
    CustomFunctions.associate('TYPE', TYPE);
    CustomFunctions.associate('PARENT', PARENT);
    CustomFunctions.associate('BALANCE', BALANCE);
    CustomFunctions.associate('BUDGET', BUDGET);
    CustomFunctions.associate('RETAINEDEARNINGS', RETAINEDEARNINGS);
    CustomFunctions.associate('NETINCOME', NETINCOME);
    CustomFunctions.associate('CTA', CTA);
    CustomFunctions.associate('CLEARCACHE', CLEARCACHE);
    console.log('✅ Custom functions registered with Excel');
} else {
    console.error('❌ CustomFunctions not available!');
}

