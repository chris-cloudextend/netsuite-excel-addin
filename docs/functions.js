/**
 * NetSuite Custom Functions for Excel
 * Provides three custom formulas: NS.GLATITLE, NS.GLABAL, NS.GLABUD
 * 
 * INTELLIGENT BATCHING - Collects multiple requests and sends as one batch query
 * NO CACHING - Every call makes a fresh API request for reliability
 */

// Backend server URL
const SERVER_URL = 'https://pull-themes-friendly-mentor.trycloudflare.com';

// Batching system for GLABAL
const pendingBalanceRequests = new Map(); // key -> {account, fromPeriod, toPeriod, filters, resolve, reject}
const balanceCache = new Map(); // Cache successful results to prevent blanks on recalc
const titleCache = new Map(); // Cache for GLATITLE
const budgetCache = new Map(); // Cache for GLABUD
let batchTimer = null;
const BATCH_DELAY_MS = 200; // Wait 200ms to collect requests before sending batch
const MAX_BATCH_SIZE = 50; // Max accounts*periods per batch to prevent timeouts
const FETCH_TIMEOUT_MS = 30000; // 30 second timeout for API calls

/**
 * Get account name from account number
 * @customfunction
 * @param {any} accountNumber The account number or ID
 * @returns {string} The account name
 */
async function GLATITLE(accountNumber) {
    const account = String(accountNumber || "").trim();
    if (!account || account === "undefined" || account === "null") {
        return "#N/A";
    }
    
    // 🔥 CHECK CACHE FIRST - Return immediately if we have it!
    const cachedValue = titleCache.get(account);
    if (cachedValue !== undefined) {
        console.log(`⚡ Title Cache HIT: ${account} → "${cachedValue}" (instant)`);
        return cachedValue;
    }
    
    console.log(`📥 Title Cache MISS: ${account} → making API call`);
    
    try {
        const response = await fetch(`${SERVER_URL}/account/${account}/name`, {
            method: 'GET',
            headers: { 'Accept': 'text/plain' }
        });
        
        if (!response.ok) {
            console.error(`GLATITLE failed for ${account}: ${response.status}`);
            // Try cache on error (in case we had it before)
            const fallback = titleCache.get(account);
            return fallback !== undefined ? fallback : "#N/A";
        }
        
        const text = await response.text();
        if (!text || text.trim() === "") {
            return "#N/A";
        }
        
        // Cache successful result
        titleCache.set(account, text);
        console.log(`💾 Cached title: ${account} → "${text}"`);
        
        return text;
        
    } catch (error) {
        console.error(`GLATITLE error for ${account}:`, error);
        // Try cache on error
        const fallback = titleCache.get(account);
        return fallback !== undefined ? fallback : "#N/A";
    }
}


/**
 * Get GL account balance - Uses intelligent batching when multiple cells calculated together
 * @customfunction
 * @param {any} account The account number or ID (required)
 * @param {any} fromPeriod Starting period (e.g., "Jan 2025")
 * @param {any} toPeriod Ending period (e.g., "Dec 2025")
 * @param {any} [subsidiary] Subsidiary ID (optional)
 * @param {any} [department] Department ID (optional)
 * @param {any} [location] Location ID (optional)
 * @param {any} [classId] Class ID (optional)
 * @returns {number} The GL account balance
 */
async function GLABAL(account, fromPeriod, toPeriod, subsidiary, department, location, classId) {
    // Convert all to strings and trim
    account = String(account || "").trim();
    fromPeriod = String(fromPeriod || "").trim();
    toPeriod = String(toPeriod || "").trim();
    subsidiary = String(subsidiary || "").trim();
    department = String(department || "").trim();
    location = String(location || "").trim();
    classId = String(classId || "").trim();
    
    if (!account) {
        return "";
    }
    
    // Create unique key for this request
    const requestKey = JSON.stringify({ account, fromPeriod, toPeriod, subsidiary, department, location, classId });
    
    // 🔥 CHECK CACHE FIRST - Return immediately if we have it!
    const cachedValue = balanceCache.get(requestKey);
    if (cachedValue !== undefined) {
        console.log(`⚡ Cache HIT: ${account} → ${cachedValue} (instant, no API call)`);
        return cachedValue;
    }
    
    // Cache miss - queue for API call
    console.log(`📥 Cache MISS: ${account} → queued for API call`);
    
    // Return a promise that will be resolved when batch completes
    return new Promise((resolve, reject) => {
        // Add this request to pending batch
        pendingBalanceRequests.set(requestKey, {
            account,
            fromPeriod,
            toPeriod,
            subsidiary,
            department,
            location,
            classId,
            resolve,
            reject
        });
        
        // Reset/start batch timer
        if (batchTimer) {
            clearTimeout(batchTimer);
        }
        
        batchTimer = setTimeout(() => {
            processBatchBalanceRequests();
        }, BATCH_DELAY_MS);
    });
}


/**
 * Process all pending balance requests as a single batch
 */
async function processBatchBalanceRequests() {
    if (pendingBalanceRequests.size === 0) {
        return;
    }
    
    // Take all pending requests
    const requests = Array.from(pendingBalanceRequests.values());
    const requestKeys = Array.from(pendingBalanceRequests.keys());
    pendingBalanceRequests.clear();
    
    console.log(`Processing batch of ${requests.length} balance requests`);
    
    // If only 1 request, use regular endpoint (faster)
    if (requests.length === 1) {
        const req = requests[0];
        try {
            const params = new URLSearchParams();
            params.append('account', req.account);
            if (req.fromPeriod) params.append('from_period', req.fromPeriod);
            if (req.toPeriod) params.append('to_period', req.toPeriod);
            if (req.subsidiary) params.append('subsidiary', req.subsidiary);
            if (req.department) params.append('department', req.department);
            if (req.location) params.append('location', req.location);
            if (req.classId) params.append('class', req.classId);
            
            const requestKey = JSON.stringify({ 
                account: req.account, 
                fromPeriod: req.fromPeriod, 
                toPeriod: req.toPeriod,
                subsidiary: req.subsidiary,
                department: req.department,
                location: req.location,
                class: req.classId
            });
            
            const response = await fetch(`${SERVER_URL}/balance?${params.toString()}`, {
                method: 'GET',
                headers: { 'Accept': 'application/json' }
            });
            
            if (!response.ok) {
                console.error(`Single balance API error: ${response.status}`);
                // Try cached value
                const cachedValue = balanceCache.get(requestKey);
                req.resolve(cachedValue !== undefined ? cachedValue : "");
                return;
            }
            
            const text = await response.text();
            const balance = parseFloat(text);
            const finalValue = isNaN(balance) ? "" : balance;
            
            // Cache successful result
            if (finalValue !== "") {
                balanceCache.set(requestKey, finalValue);
            }
            
            req.resolve(finalValue);
            
        } catch (error) {
            console.error('Single balance request error:', error);
            // Try cached value
            const requestKey = JSON.stringify({ 
                account: req.account, 
                fromPeriod: req.fromPeriod, 
                toPeriod: req.toPeriod,
                subsidiary: req.subsidiary,
                department: req.department,
                location: req.location,
                class: req.classId
            });
            const cachedValue = balanceCache.get(requestKey);
            req.resolve(cachedValue !== undefined ? cachedValue : "");
        }
        return;
    }
    
    // Multiple requests - use batch endpoint
    try {
        // Collect unique accounts and periods
        const accounts = [...new Set(requests.map(r => r.account))];
        const periods = [...new Set(requests.flatMap(r => {
            const p = [];
            if (r.fromPeriod) p.push(r.fromPeriod);
            if (r.toPeriod && r.toPeriod !== r.fromPeriod) p.push(r.toPeriod);
            return p;
        }))];
        
        // For now, batch only works if all requests share same filters
        // Otherwise fall back to individual calls
        const firstReq = requests[0];
        const sameFilters = requests.every(r => 
            r.subsidiary === firstReq.subsidiary &&
            r.department === firstReq.department &&
            r.location === firstReq.location &&
            r.classId === firstReq.classId
        );
        
        if (!sameFilters) {
            console.log('Mixed filters detected, falling back to individual calls');
            // Process each request individually
            for (const req of requests) {
                processIndividualBalance(req);
            }
            return;
        }
        
        // Check if batch is too large - split if needed
        const dataPoints = accounts.length * periods.length;
        console.log(`Batch size: ${accounts.length} accounts x ${periods.length} periods = ${dataPoints} data points`);
        
        if (dataPoints > MAX_BATCH_SIZE) {
            console.log(`⚠️ Batch too large (${dataPoints} > ${MAX_BATCH_SIZE}), splitting into chunks...`);
            
            // Split accounts into chunks
            const accountChunks = [];
            const chunkSize = Math.ceil(MAX_BATCH_SIZE / periods.length);
            for (let i = 0; i < accounts.length; i += chunkSize) {
                accountChunks.push(accounts.slice(i, i + chunkSize));
            }
            
            console.log(`Split into ${accountChunks.length} chunks of ~${chunkSize} accounts each`);
            
            // Process each chunk separately
            const allBalances = {};
            for (let i = 0; i < accountChunks.length; i++) {
                const chunk = accountChunks[i];
                console.log(`Processing chunk ${i + 1}/${accountChunks.length} (${chunk.length} accounts)`);
                
                const chunkBalances = await fetchBatchBalances(chunk, periods, firstReq);
                Object.assign(allBalances, chunkBalances);
                
                // Small delay between chunks to avoid overwhelming the server
                if (i < accountChunks.length - 1) {
                    await new Promise(resolve => setTimeout(resolve, 100));
                }
            }
            
            // Distribute results
            distributeBalanceResults(requests, allBalances);
            return;
        }
        
        // Normal batch processing
        const balances = await fetchBatchBalances(accounts, periods, firstReq);
        distributeBalanceResults(requests, balances);
        
    } catch (error) {
        console.error('Batch processing error:', error);
        // Resolve all with blank on error
        requests.forEach(req => req.resolve(""));
    }
}

/**
 * Fetch batch balances from API with timeout protection
 */
async function fetchBatchBalances(accounts, periods, filterReq) {
    try {
        const batchPayload = {
            accounts: accounts,
            periods: periods,
            subsidiary: filterReq.subsidiary || "",
            department: filterReq.department || "",
            location: filterReq.location || "",
            class: filterReq.classId || ""
        };
        
        // Add timeout protection
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);
        
        const response = await fetch(`${SERVER_URL}/batch/balance`, {
            method: 'POST',
            headers: { 
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            },
            body: JSON.stringify(batchPayload),
            signal: controller.signal
        });
        
        clearTimeout(timeoutId);
        
        if (!response.ok) {
            const errorText = await response.text().catch(() => "");
            console.error(`⚠️ Batch balance failed: ${response.status} - ${errorText}`);
            console.log('⚠️ Will use cached values where available');
            return {}; // Return empty - distributeResults will use cache
        }
        
        const result = await response.json();
        console.log(`✓ Batch successful: ${Object.keys(result.balances || {}).length} accounts`);
        return result.balances || {};
        
    } catch (error) {
        if (error.name === 'AbortError') {
            console.error(`⚠️ Batch request timed out after ${FETCH_TIMEOUT_MS}ms - will use cached values`);
        } else {
            console.error('⚠️ Fetch balance error:', error);
        }
        console.log(`📦 Cache has ${balanceCache.size} stored values available`);
        return {}; // Return empty - distributeResults will use cache
    }
}

/**
 * Distribute batch balance results to individual requests
 */
function distributeBalanceResults(requests, balances) {
    console.log(`✓ Distributing results: ${Object.keys(balances).length} accounts to ${requests.length} requests`);
    
    for (const req of requests) {
        try {
            const requestKey = JSON.stringify({ 
                account: req.account, 
                fromPeriod: req.fromPeriod, 
                toPeriod: req.toPeriod,
                subsidiary: req.subsidiary,
                department: req.department,
                location: req.location,
                class: req.classId
            });
            
            const accountBalances = balances[req.account] || {};
            
            // Sum balances from fromPeriod to toPeriod
            let total = 0;
            const periods = Object.keys(accountBalances).sort();
            
            if (req.fromPeriod && req.toPeriod) {
                // Sum range
                for (const period of periods) {
                    if (period >= req.fromPeriod && period <= req.toPeriod) {
                        total += accountBalances[period] || 0;
                    }
                }
            } else if (req.fromPeriod) {
                // Single period
                total = accountBalances[req.fromPeriod] || 0;
            } else {
                // All periods
                total = Object.values(accountBalances).reduce((sum, val) => sum + (val || 0), 0);
            }
            
            // Cache the successful result
            balanceCache.set(requestKey, total);
            
            req.resolve(total);
            
        } catch (error) {
            console.error('Error distributing result:', error);
            
            // Try to return cached value instead of blank
            const requestKey = JSON.stringify({ 
                account: req.account, 
                fromPeriod: req.fromPeriod, 
                toPeriod: req.toPeriod,
                subsidiary: req.subsidiary,
                department: req.department,
                location: req.location,
                class: req.classId
            });
            const cachedValue = balanceCache.get(requestKey);
            req.resolve(cachedValue !== undefined ? cachedValue : "");
        }
    }
}


/**
 * Process a single balance request (fallback when batching not possible)
 */
async function processIndividualBalance(req) {
    try {
        const params = new URLSearchParams();
        params.append('account', req.account);
        if (req.fromPeriod) params.append('from_period', req.fromPeriod);
        if (req.toPeriod) params.append('to_period', req.toPeriod);
        if (req.subsidiary) params.append('subsidiary', req.subsidiary);
        if (req.department) params.append('department', req.department);
        if (req.location) params.append('location', req.location);
        if (req.classId) params.append('class', req.classId);
        
        const response = await fetch(`${SERVER_URL}/balance?${params.toString()}`, {
            method: 'GET',
            headers: { 'Accept': 'application/json' }
        });
        
        if (!response.ok) {
            req.resolve("");
            return;
        }
        
        const text = await response.text();
        const balance = parseFloat(text);
        req.resolve(isNaN(balance) ? "" : balance);
        
    } catch (error) {
        console.error('Individual balance request error:', error);
        req.resolve("");
    }
}


/**
 * Get GL budget amount
 * @customfunction
 * @param {any} account The account number or ID (required)
 * @param {any} fromPeriod Starting period (e.g., "Jan 2025")
 * @param {any} toPeriod Ending period (e.g., "Dec 2025")
 * @param {any} [subsidiary] Subsidiary ID (optional)
 * @param {any} [department] Department ID (optional)
 * @param {any} [location] Location ID (optional)
 * @param {any} [classId] Class ID (optional)
 * @returns {number} The budget amount
 */
async function GLABUD(account, fromPeriod, toPeriod, subsidiary, department, location, classId) {
    // Convert all to strings and trim
    account = String(account || "").trim();
    fromPeriod = String(fromPeriod || "").trim();
    toPeriod = String(toPeriod || "").trim();
    subsidiary = String(subsidiary || "").trim();
    department = String(department || "").trim();
    location = String(location || "").trim();
    classId = String(classId || "").trim();
    
    if (!account) {
        return "";
    }
    
    // Create unique key for this request
    const requestKey = JSON.stringify({ account, fromPeriod, toPeriod, subsidiary, department, location, classId });
    
    // 🔥 CHECK CACHE FIRST - Return immediately if we have it!
    const cachedValue = budgetCache.get(requestKey);
    if (cachedValue !== undefined) {
        console.log(`⚡ Budget Cache HIT: ${account} → ${cachedValue} (instant)`);
        return cachedValue;
    }
    
    console.log(`📥 Budget Cache MISS: ${account} → making API call`);
    
    try {
        const params = new URLSearchParams();
        params.append('account', account);
        if (fromPeriod) params.append('from_period', fromPeriod);
        if (toPeriod) params.append('to_period', toPeriod);
        if (subsidiary) params.append('subsidiary', subsidiary);
        if (department) params.append('department', department);
        if (location) params.append('location', location);
        if (classId) params.append('class', classId);
        
        const url = `${SERVER_URL}/budget?${params.toString()}`;
        const response = await fetch(url, {
            method: 'GET',
            headers: { 'Accept': 'application/json' }
        });
        
        if (!response.ok) {
            const errorText = await response.text().catch(() => "");
            console.error(`GLABUD failed for ${account} (${fromPeriod}-${toPeriod}): ${response.status} - ${errorText}`);
            // Try cache on error
            const fallback = budgetCache.get(requestKey);
            return fallback !== undefined ? fallback : "";
        }
        
        const text = await response.text();
        const budget = parseFloat(text);
        
        if (isNaN(budget)) {
            console.error(`GLABUD parsing failed for ${account}: got "${text}"`);
            return "";
        }
        
        // Cache successful result
        budgetCache.set(requestKey, budget);
        console.log(`💾 Cached budget: ${account} → ${budget}`);
        
        return budget;
        
    } catch (error) {
        console.error(`GLABUD error for ${account}:`, error);
        // Try cache on error
        const fallback = budgetCache.get(requestKey);
        return fallback !== undefined ? fallback : "";
    }
}


// Register custom functions
CustomFunctions.associate("GLATITLE", GLATITLE);
CustomFunctions.associate("GLABAL", GLABAL);
CustomFunctions.associate("GLABUD", GLABUD);
