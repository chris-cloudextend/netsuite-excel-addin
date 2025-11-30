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
    budget: new Map()      // budget cache
};

// Track last access time to implement LRU if needed
const cacheStats = {
    hits: 0,
    misses: 0,
    size: () => cache.balance.size + cache.title.size + cache.budget.size
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
const CHUNK_SIZE = 10;             // Max 10 accounts per batch
const RETRY_DELAY = 2000;          // Wait 2s before retrying 429 errors
const MAX_RETRIES = 2;             // Retry 429 errors up to 2 times

// ============================================================================
// UTILITY: Generate cache key
// ============================================================================
function getCacheKey(type, params) {
    if (type === 'title') {
        return `title:${params.account}`;
    } else if (type === 'balance' || type === 'budget') {
        return JSON.stringify({
            type,
            account: params.account,
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
    const account = String(accountNumber || '').trim();
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
// GLABAL - Get GL Account Balance (WITH SMART BATCHING)
// ============================================================================
/**
 * @customfunction GLABAL
 * @streaming
 * @cancelable
 */
function GLABAL(account, fromPeriod, toPeriod, subsidiary, department, location, classId, invocation) {
    // CRITICAL: Thin registration function - NO POLLING, NO WAITING
    // The batch processor will call invocation.setResult() and invocation.close()
    
    try {
        // Normalize inputs
        account = String(account || '').trim();
        fromPeriod = String(fromPeriod || '').trim();
        toPeriod = String(toPeriod || '').trim();
        subsidiary = String(subsidiary || '').trim();
        department = String(department || '').trim();
        location = String(location || '').trim();
        classId = String(classId || '').trim();
        
        if (!account) {
            invocation.setResult(0);
            invocation.close();
            return;
        }
        
        const params = { account, fromPeriod, toPeriod, subsidiary, department, location, classId };
        const cacheKey = getCacheKey('balance', params);
        
        // Check cache FIRST - return immediately if found
        if (cache.balance.has(cacheKey)) {
            cacheStats.hits++;
            const value = cache.balance.get(cacheKey);
            console.log(`⚡ CACHE HIT [balance]: ${account} → ${value}`);
            invocation.setResult(value);
            invocation.close();
            return;
        }
        
        // Cache miss → queue this invocation for batching
        cacheStats.misses++;
        console.log(`📥 CACHE MISS [balance]: ${account} → queuing`);
        
        requestQueue.balance.set(cacheKey, {
            params,
            invocation,  // Store the invocation - batch processor will use it
            retries: 0
        });
        
        // Handle cancellation
        invocation.onCanceled = () => {
            console.log(`⏹ Canceled [balance]: ${account}`);
            requestQueue.balance.delete(cacheKey);
        };
        
        // Start batch processing in a microtask if not already running
        if (!batchTimer) {
            batchTimer = true;
            Promise.resolve().then(() => {
                batchTimer = null;
                processBatchQueue().catch(err => {
                    console.error("processBatchQueue error:", err);
                });
            });
        }
        
        // NO return value. Streaming completes when batch processor calls invocation.close()
        
    } catch (error) {
        console.error('GLABAL synchronous error:', error);
        invocation.setResult(0);
        invocation.close();
    }
}

// ============================================================================
// GLABUD - Get Budget Amount (SAME LOGIC AS GLABAL)
// ============================================================================
/**
 * @customfunction GLABUD
 * @streaming
 * @cancelable
 */
function GLABUD(account, fromPeriod, toPeriod, subsidiary, department, location, classId, invocation) {
    // CRITICAL: Outer function must be SYNCHRONOUS (not async)
    // No return values allowed - only invocation.setResult() + close()
    
    try {
        // Normalize inputs safely
        account = String(account || '').trim();
        fromPeriod = String(fromPeriod || '').trim();
        toPeriod = String(toPeriod || '').trim();
        subsidiary = String(subsidiary || '').trim();
        department = String(department || '').trim();
        location = String(location || '').trim();
        classId = String(classId || '').trim();
        
        if (!account) {
            invocation.setResult(0);  // Return 0 instead of empty for number type
            invocation.close();
            return;  // Early exit is OK (no value returned)
        }
        
        const params = { account, fromPeriod, toPeriod, subsidiary, department, location, classId };
        const cacheKey = getCacheKey('budget', params);
        
        // Check cache FIRST - return immediately if found
        if (cache.budget.has(cacheKey)) {
            cacheStats.hits++;
            invocation.setResult(cache.budget.get(cacheKey));
            invocation.close();
            return;  // Early exit is OK (no value returned)
        }
        
        cacheStats.misses++;
        
        // Handle cancellation
        const controller = new AbortController();
        invocation.onCanceled = () => {
            console.log('Budget request canceled');
            controller.abort();
        };
        
        // Async work wrapped in IIFE (immediately invoked async function)
        (async () => {
            try {
                const url = new URL(`${SERVER_URL}/budget`);
                url.searchParams.append('account', account);
                if (fromPeriod) url.searchParams.append('from_period', fromPeriod);
                if (toPeriod) url.searchParams.append('to_period', toPeriod);
                if (subsidiary) url.searchParams.append('subsidiary', subsidiary);
                if (department) url.searchParams.append('department', department);
                if (location) url.searchParams.append('location', location);
                if (classId) url.searchParams.append('class', classId);
                
                const response = await fetch(url.toString(), { signal: controller.signal });
                if (!response.ok) {
                    console.error(`Budget API error: ${response.status}`);
                    invocation.setResult(0);  // Return 0 for number type
                    invocation.close();
                    return;
                }
                
                const text = await response.text();
                const budget = parseFloat(text);
                const finalValue = isNaN(budget) ? 0 : budget;  // Return 0 instead of empty
                
                if (finalValue !== 0) {
                    cache.budget.set(cacheKey, finalValue);
                }
                
                invocation.setResult(finalValue);
                invocation.close();
                
            } catch (error) {
                if (error.name !== 'AbortError') {
                    console.error('Budget fetch error:', error);
                }
                invocation.setResult(0);  // Return 0 for number type
                invocation.close();
            }
        })();
        
        // NO return statement - streaming function keeps running
        
    } catch (error) {
        // Handle any synchronous errors
        console.error('GLABUD synchronous error:', error);
        invocation.setResult(0);  // Return 0 for number type
        invocation.close();
        return;  // Early exit is OK (no value returned)
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
    
    // Group by filters (batch only identical filter sets)
    const groups = new Map();
    for (const [key, req] of requests) {
        const filterKey = JSON.stringify({
            subsidiary: req.params.subsidiary,
            department: req.params.department,
            location: req.params.location,
            class: req.params.classId
        });
        
        if (!groups.has(filterKey)) {
            groups.set(filterKey, []);
        }
        groups.get(filterKey).push({ key, req });
    }
    
    console.log(`📦 Grouped into ${groups.size} filter group(s)`);
    
    // Process each group
    for (const [filterKey, groupRequests] of groups.entries()) {
        const filters = JSON.parse(filterKey);
        const accounts = [...new Set(groupRequests.map(r => r.req.params.account))];
        const periods = [...new Set(groupRequests.flatMap(r => {
            const p = [];
            if (r.req.params.fromPeriod) p.push(r.req.params.fromPeriod);
            if (r.req.params.toPeriod && r.req.params.toPeriod !== r.req.params.fromPeriod) {
                p.push(r.req.params.toPeriod);
            }
            return p;
        }))];
        
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
                // Resolve all with blank
                for (const { key, req } of allRequests) {
                    if (accounts.includes(req.params.account)) {
                        req.resolve('');
                    }
                }
                return;
            }
        }
        
        if (!response.ok) {
            console.error(`Batch API error: ${response.status}`);
            // Resolve all with blank
            for (const { key, req } of allRequests) {
                if (accounts.includes(req.params.account)) {
                    req.resolve('');
                }
            }
            return;
        }
        
        const data = await response.json();
        const balances = data.balances || {};
        
        console.log(`  ✅ Received balances for ${Object.keys(balances).length} accounts`);
        
        // Distribute results to invocations and close them
        for (const { key, req } of allRequests) {
            try {
                if (!accounts.includes(req.params.account)) {
                    console.warn(`⚠️  Account ${req.params.account} not in response`);
                    safeFinishInvocation(req.invocation, 0);
                    continue;
                }
                
                const accountBalances = balances[req.params.account] || {};
                let total = 0;
                const periodList = Object.keys(accountBalances).sort();
                
                if (req.params.fromPeriod && req.params.toPeriod) {
                    // Sum range
                    for (const period of periodList) {
                        if (period >= req.params.fromPeriod && period <= req.params.toPeriod) {
                            total += accountBalances[period] || 0;
                        }
                    }
                } else if (req.params.fromPeriod) {
                    // Single period
                    total = accountBalances[req.params.fromPeriod] || 0;
                }
                
                // Cache the result and finish the invocation
                cache.balance.set(key, total);
                console.log(`💾 Cached result for ${req.params.account}: ${total}`);
                safeFinishInvocation(req.invocation, total);
                
            } catch (error) {
                console.error('Error distributing result:', error, key);
                cache.balance.set(key, 0);
                safeFinishInvocation(req.invocation, 0);
            }
        }
        
    } catch (error) {
        console.error('❌ Batch fetch error:', error);
        // DEFENSIVE: ALWAYS close all invocations on error (ChatGPT requirement)
        console.log(`⚠️  Closing ${allRequests.length} invocations due to error...`);
        for (const { key, req } of allRequests) {
            try {
                if (req.invocation && typeof req.invocation.setResult === 'function') {
                    console.log(`  → Closing invocation for ${req.params.account} with 0`);
                    req.invocation.setResult(0);  // Return 0 for number type
                    req.invocation.close();
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
// UTILITY: Delay helper
// ============================================================================
function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// ============================================================================
// REGISTER FUNCTIONS WITH EXCEL
// ============================================================================
if (typeof CustomFunctions !== 'undefined') {
    // Associate each function by its ID from functions.json
    // The namespace "NS" comes from the manifest, not the association!
    CustomFunctions.associate('GLATITLE', GLATITLE);
    CustomFunctions.associate('GLABAL', GLABAL);
    CustomFunctions.associate('GLABUD', GLABUD);
    console.log('✅ Custom functions registered with Excel');
} else {
    console.error('❌ CustomFunctions not available!');
}

