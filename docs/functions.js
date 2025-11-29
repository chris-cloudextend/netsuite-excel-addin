/**
 * NetSuite Custom Functions for Excel
 * Provides three custom formulas: NS.GLATITLE, NS.GLABAL, NS.GLABUD
 * 
 * INTELLIGENT BATCHING - Collects multiple requests and sends as one batch query
 * NO CACHING - Every call makes a fresh API request for reliability
 */

// Backend server URL
const SERVER_URL = 'https://load-scanner-nathan-targeted.trycloudflare.com';

// Batching system for GLABAL
const pendingBalanceRequests = new Map(); // key -> {account, fromPeriod, toPeriod, filters, resolve, reject}
let batchTimer = null;
const BATCH_DELAY_MS = 200; // Wait 200ms to collect requests before sending batch

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
    
    try {
        const response = await fetch(`${SERVER_URL}/account/${account}/name`, {
            method: 'GET',
            headers: { 'Accept': 'text/plain' }
        });
        
        if (!response.ok) {
            console.error(`GLATITLE failed for ${account}: ${response.status}`);
            return "#N/A";
        }
        
        const text = await response.text();
        if (!text || text.trim() === "") {
            return "#N/A";
        }
        
        return text;
        
    } catch (error) {
        console.error(`GLATITLE error for ${account}:`, error);
        return "#N/A";
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
            
            const response = await fetch(`${SERVER_URL}/balance?${params.toString()}`, {
                method: 'GET',
                headers: { 'Accept': 'application/json' }
            });
            
            if (!response.ok) {
                req.reject(new Error(`API error: ${response.status}`));
                return;
            }
            
            const text = await response.text();
            const balance = parseFloat(text);
            req.resolve(isNaN(balance) ? "" : balance);
            
        } catch (error) {
            console.error('Single balance request error:', error);
            req.resolve("");
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
        
        // Build batch request
        const batchPayload = {
            accounts: accounts,
            periods: periods,
            subsidiary: firstReq.subsidiary || "",
            department: firstReq.department || "",
            location: firstReq.location || "",
            class: firstReq.classId || ""
        };
        
        console.log('Sending batch request:', batchPayload);
        
        const response = await fetch(`${SERVER_URL}/batch/balance`, {
            method: 'POST',
            headers: { 
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            },
            body: JSON.stringify(batchPayload)
        });
        
        if (!response.ok) {
            const errorText = await response.text().catch(() => "");
            console.error(`Batch balance failed: ${response.status} - ${errorText}`);
            // Resolve all with blank
            requests.forEach(req => req.resolve(""));
            return;
        }
        
        const result = await response.json();
        const balances = result.balances || {};
        
        console.log('Batch response received:', Object.keys(balances).length, 'accounts');
        
        // Distribute results back to individual requests
        for (const req of requests) {
            try {
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
                
                req.resolve(total);
                
            } catch (error) {
                console.error('Error processing batch result for request:', error);
                req.resolve("");
            }
        }
        
    } catch (error) {
        console.error('Batch balance request failed:', error);
        // Resolve all with blank on error
        requests.forEach(req => req.resolve(""));
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
            return "";
        }
        
        const text = await response.text();
        const budget = parseFloat(text);
        
        if (isNaN(budget)) {
            console.error(`GLABUD parsing failed for ${account}: got "${text}"`);
            return "";
        }
        
        return budget;
        
    } catch (error) {
        console.error(`GLABUD error for ${account}:`, error);
        return "";
    }
}


// Register custom functions
CustomFunctions.associate("GLATITLE", GLATITLE);
CustomFunctions.associate("GLABAL", GLABAL);
CustomFunctions.associate("GLABUD", GLABUD);
