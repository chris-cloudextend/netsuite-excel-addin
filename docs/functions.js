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
// CACHE - Never expires, persists entire Excel session
// ============================================================================
const cache = {
    balance: new Map(),    // balance cache
    title: new Map(),      // account title cache  
    budget: new Map(),     // budget cache
    type: new Map(),       // account type cache
    parent: new Map()      // parent account cache
};

// Track last access time to implement LRU if needed
const cacheStats = {
    hits: 0,
    misses: 0,
    size: () => cache.balance.size + cache.title.size + cache.budget.size + cache.type.size + cache.parent.size
};

// ============================================================================
// REQUEST QUEUE - Collects requests for intelligent batching
// ============================================================================
const requestQueue = {
    balance: new Map(),    // Pending balance requests
    title: new Map(),      // Pending title requests
    budget: new Map()      // Pending budget requests
};

let batchTimer = null;  // Timer reference for batching
const BATCH_DELAY = 100;  // Reduced delay for streaming (was 500ms)
const MAX_CONCURRENT = 3;          // Max 3 concurrent requests to NetSuite
const CHUNK_SIZE = 30;             // Max 30 accounts per batch (increased from 10 for better performance)
const RETRY_DELAY = 2000;          // Wait 2s before retrying 429 errors
const MAX_RETRIES = 2;             // Retry 429 errors up to 2 times

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
// GLATITLE - Get Account Name
// ============================================================================
/**
 * @customfunction GLATITLE
 * @param {any} accountNumber The account number
 * @param {CustomFunctions.Invocation} invocation Invocation object
 * @returns {Promise<string>} Account name
 * @requiresAddress
 * @cancelable
 */
async function GLATITLE(accountNumber, invocation) {
    const account = normalizeAccountNumber(accountNumber);
    if (!account) return '#N/A';
    
    const cacheKey = getCacheKey('title', { account });
    
    // Check cache FIRST
    if (cache.title.has(cacheKey)) {
        cacheStats.hits++;
        console.log(`⚡ CACHE HIT [title]: ${account}`);
        return cache.title.get(cacheKey);
    }
    
    cacheStats.misses++;
    console.log(`📥 CACHE MISS [title]: ${account}`);
    
    // Single request - make immediately (don't batch titles)
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
        
        const response = await fetch(`${SERVER_URL}/account/${account}/name`, { signal });
        if (!response.ok) {
            console.error(`Title API error: ${response.status}`);
            return '#N/A';
        }
        
        const title = await response.text();
        cache.title.set(cacheKey, title);
        console.log(`💾 Cached title: ${account} → "${title}"`);
        return title;
        
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
// GLACCTTYPE - Get Account Type
// ============================================================================
/**
 * @customfunction GLACCTTYPE
 * @param {any} accountNumber The account number
 * @param {CustomFunctions.Invocation} invocation Invocation object
 * @returns {Promise<string>} Account type (e.g., "Income", "Expense")
 * @requiresAddress
 * @cancelable
 */
async function GLACCTTYPE(accountNumber, invocation) {
    const account = normalizeAccountNumber(accountNumber);
    if (!account) return '#N/A';
    
    const cacheKey = getCacheKey('type', { account });
    
    // Check cache FIRST
    if (!cache.type) cache.type = new Map();
    if (cache.type.has(cacheKey)) {
        cacheStats.hits++;
        console.log(`⚡ CACHE HIT [type]: ${account}`);
        return cache.type.get(cacheKey);
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
        
        const response = await fetch(`${SERVER_URL}/account/${account}/type`, { signal });
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
// GLAPARENT - Get Parent Account
// ============================================================================
/**
 * @customfunction GLAPARENT
 * @param {any} accountNumber The account number
 * @param {CustomFunctions.Invocation} invocation Invocation object
 * @returns {Promise<string>} Parent account number
 * @requiresAddress
 * @cancelable
 */
async function GLAPARENT(accountNumber, invocation) {
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
        
        const response = await fetch(`${SERVER_URL}/account/${account}/parent`, { signal });
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
// GLABAL - Get GL Account Balance (NON-STREAMING ASYNC)
// ============================================================================
/**
 * @customfunction GLABAL
 * @param {any} account Account number
 * @param {any} fromPeriod Starting period (e.g., "Jan 2025" or 1/1/2025)
 * @param {any} toPeriod Ending period (e.g., "Mar 2025" or 3/1/2025)
 * @param {any} [subsidiary] Subsidiary filter (optional)
 * @param {any} [department] Department filter (optional)
 * @param {any} [location] Location filter (optional)
 * @param {any} [classId] Class filter (optional)
 * @returns {Promise<number>} Account balance
 * @requiresAddress
 */
async function GLABAL(account, fromPeriod, toPeriod, subsidiary, department, location, classId) {
    try {
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
        
        const params = { account, fromPeriod, toPeriod, subsidiary, department, location, classId };
        const cacheKey = getCacheKey('balance', params);
        
        // Check cache FIRST - return immediately if found (NO @ SYMBOL!)
        if (cache.balance.has(cacheKey)) {
            cacheStats.hits++;
            const value = cache.balance.get(cacheKey);
            // Silent cache hit (no console.log for performance)
            return value;
        }
        
        // Cache miss - need to fetch from backend
        cacheStats.misses++;
        console.log(`📥 CACHE MISS [balance]: ${account} (${fromPeriod} to ${toPeriod})`);
        
        // For Phase 1: Direct fetch (we'll optimize batching in Phase 3)
        // This ensures we don't timeout (no 5-second limit in async functions)
        try {
            const response = await fetch(`${SERVER_URL}/batch/balance`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    requests: [{
                        account,
                        fromPeriod,
                        toPeriod,
                        subsidiary,
                        department,
                        location,
                        classId
                    }]
                })
            });
            
            if (!response.ok) {
                console.error(`Balance API error: ${response.status}`);
                return 0;
            }
            
            const data = await response.json();
            const result = data.results?.[0];
            
            if (result) {
                const value = result.balance || 0;
                // Cache the result
                cache.balance.set(cacheKey, value);
                console.log(`💾 Cached balance: ${account} → ${value}`);
                return value;
            }
            
            return 0;
            
        } catch (error) {
            console.error('Balance fetch error:', error);
            return 0;
        }
        
    } catch (error) {
        console.error('GLABAL error:', error);
        return 0;
    }
}

// ============================================================================
// GLABUD - Get Budget Amount (NON-STREAMING ASYNC)
// ============================================================================
/**
 * @customfunction GLABUD
 * @param {any} account Account number
 * @param {any} fromPeriod Starting period (e.g., "Jan 2025" or 1/1/2025)
 * @param {any} toPeriod Ending period (e.g., "Mar 2025" or 3/1/2025)
 * @param {any} [subsidiary] Subsidiary filter (optional)
 * @param {any} [department] Department filter (optional)
 * @param {any} [location] Location filter (optional)
 * @param {any} [classId] Class filter (optional)
 * @returns {Promise<number>} Budget amount
 * @requiresAddress
 */
async function GLABUD(account, fromPeriod, toPeriod, subsidiary, department, location, classId) {
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
        
        const params = { account, fromPeriod, toPeriod, subsidiary, department, location, classId };
        const cacheKey = getCacheKey('budget', params);
        
        // Check cache FIRST - return immediately if found (NO @ SYMBOL!)
        if (cache.budget.has(cacheKey)) {
            cacheStats.hits++;
            const value = cache.budget.get(cacheKey);
            // Silent cache hit (no console.log for performance)
            return value;
        }
        
        // Cache miss - fetch from backend
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
            
            const response = await fetch(url.toString());
            if (!response.ok) {
                console.error(`Budget API error: ${response.status}`);
                return 0;
            }
            
            const text = await response.text();
            const budget = parseFloat(text);
            const finalValue = isNaN(budget) ? 0 : budget;
            
            // Cache the result
            if (finalValue !== 0) {
                cache.budget.set(cacheKey, finalValue);
                console.log(`💾 Cached budget: ${account} → ${finalValue}`);
            }
            
            return finalValue;
            
        } catch (error) {
            console.error('Budget fetch error:', error);
            return 0;
        }
        
    } catch (error) {
        console.error('GLABUD error:', error);
        return 0;
    }
}

// ============================================================================
// BATCH PROCESSING - Streaming Model (Immediate Start)
// ============================================================================
async function processBatchQueue() {
    batchTimer = null;  // Reset timer reference
    
    if (requestQueue.balance.size === 0) {
        console.log('No requests in queue');
        return;
    }
    
    console.log(`\n🔄 Processing batch queue: ${requestQueue.balance.size} requests`);
    console.log(`📊 Cache stats: ${cacheStats.hits} hits / ${cacheStats.misses} misses / ${cacheStats.size()} entries`);
    
    // Convert queue to array (no invocation objects stored anymore!)
    const requests = Array.from(requestQueue.balance.entries());
    requestQueue.balance.clear();
    
    // Group by filters AND periods (correct, working solution)
    // Each unique period gets its own batch for accurate results
    const groups = new Map();
    for (const [key, req] of requests) {
        const filterKey = JSON.stringify({
            subsidiary: req.params.subsidiary,
            department: req.params.department,
            location: req.params.location,
            class: req.params.classId,
            fromPeriod: req.params.fromPeriod || '',  // Include in grouping
            toPeriod: req.params.toPeriod || ''       // Include in grouping
        });
        
        if (!groups.has(filterKey)) {
            groups.set(filterKey, []);
        }
        groups.get(filterKey).push({ key, req });
    }
    
    console.log(`📦 Grouped into ${groups.size} filter+period group(s)`);
    
    // Process each group
    for (const [filterKey, groupRequests] of groups.entries()) {
        const filters = JSON.parse(filterKey);
        const accounts = [...new Set(groupRequests.map(r => r.req.params.account))];
        
        // All requests in this group have the SAME period range
        // Expand it once for the entire group
        const firstReq = groupRequests[0].req;
        const expandedPeriods = expandPeriodRange(firstReq.params.fromPeriod, firstReq.params.toPeriod);
        const periods = expandedPeriods;
        
        console.log(`  Group: ${accounts.length} accounts × ${periods.length} periods = ${accounts.length * periods.length} data points`);
        
        // Split into small chunks to avoid NetSuite 429 errors
        const accountChunks = [];
        for (let i = 0; i < accounts.length; i += CHUNK_SIZE) {
            accountChunks.push(accounts.slice(i, i + CHUNK_SIZE));
        }
        
        console.log(`  Split into ${accountChunks.length} chunk(s) of max ${CHUNK_SIZE} accounts`);
        
        // Process chunks sequentially with delays
        for (let i = 0; i < accountChunks.length; i++) {
            const chunk = accountChunks[i];
            console.log(`  📤 Processing chunk ${i + 1}/${accountChunks.length} (${chunk.length} accounts)...`);
            
            await fetchBatchBalances(chunk, periods, filters, groupRequests);
            
            // Delay between chunks to avoid 429 errors
            if (i < accountChunks.length - 1) {
                console.log(`  ⏱️  Waiting 1 second before next chunk...`);
                await delay(1000);
            }
        }
    }
    
    console.log('✅ Batch processing complete - results cached\n');
}

// ============================================================================
// FETCH BATCH BALANCES - with 429 retry logic
// ============================================================================
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

// ============================================================================
// HELPER: Safely finish an invocation
// ============================================================================
function safeFinishInvocation(invocation, value) {
    if (!invocation) {
        console.warn("⚠️  safeFinishInvocation called with null invocation");
        return;
    }
    
    try {
        if (typeof invocation.setResult === "function") {
            console.log(`  ✅ setResult(${value})`);
            invocation.setResult(value);
        } else {
            console.warn("  ⚠️  invocation.setResult is not a function!");
        }
        
        // Only call close if it exists (Mac Excel uses preview invocations without close)
        if (typeof invocation.close === "function") {
            console.log(`  ✅ close()`);
            invocation.close();
        } else {
            console.log("  ℹ️  No close() method (preview invocation)");
        }
    } catch (e) {
        console.error("❌ Error finishing invocation:", e);
    }
}

// ============================================================================
// UTILITY: Delay helper
// ============================================================================
function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// ============================================================================
// REGISTER FUNCTIONS WITH EXCEL
// ============================================================================
// CRITICAL: The manifest ALREADY defines namespace 'NS'
// We just register individual functions - Excel adds the NS. prefix automatically!
if (typeof CustomFunctions !== 'undefined') {
    CustomFunctions.associate('GLATITLE', GLATITLE);
    CustomFunctions.associate('GLACCTTYPE', GLACCTTYPE);
    CustomFunctions.associate('GLAPARENT', GLAPARENT);
    CustomFunctions.associate('GLABAL', GLABAL);
    CustomFunctions.associate('GLABUD', GLABUD);
    console.log('✅ Custom functions registered with Excel');
} else {
    console.error('❌ CustomFunctions not available!');
}

