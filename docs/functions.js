/**
 * NetSuite Custom Functions for Excel
 * Provides three custom formulas: NS.GLATITLE, NS.GLABAL, NS.GLABUD
 * 
 * Uses intelligent batching to handle drag-and-drop efficiently
 */

// Backend server URL
const SERVER_URL = 'https://attention-birthday-cherry-shuttle.trycloudflare.com';

// Cache for function results
const functionCache = new Map();
const CACHE_DURATION = 120000; // 2 minutes

// Batch processing system
const batchBuffer = new Map();
const BATCH_DELAY = 150; // Wait 150ms to collect batch requests

/**
 * Generate cache key from parameters
 */
function getCacheKey(functionName, ...params) {
    return `${functionName}:${params.join('|')}`;
}

/**
 * Get cached result if available and not expired
 */
function getCachedResult(cacheKey) {
    const cached = functionCache.get(cacheKey);
    if (cached && Date.now() - cached.timestamp < CACHE_DURATION) {
        return cached.value;
    }
    return null;
}

/**
 * Store result in cache
 */
function setCachedResult(cacheKey, value) {
    functionCache.set(cacheKey, {
        value: value,
        timestamp: Date.now()
    });
}

/**
 * Get or create a batch for this account/department combination
 */
function getBatch(account, subsidiary, department, location, classId) {
    const batchKey = `${account}|${subsidiary}|${department}|${location}|classId`;
    
    if (!batchBuffer.has(batchKey)) {
        batchBuffer.set(batchKey, {
            account: account,
            subsidiary: subsidiary,
            department: department,
            location: location,
            classId: classId,
            periods: new Map(), // period -> array of resolve functions
            timer: null
        });
    }
    
    return batchBuffer.get(batchKey);
}

/**
 * Process a batch request
 */
async function processBatch(batchKey) {
    const batch = batchBuffer.get(batchKey);
    if (!batch || batch.periods.size === 0) {
        batchBuffer.delete(batchKey);
        return;
    }
    
    const periods = Array.from(batch.periods.keys());
    const requests = Array.from(batch.periods.values());
    
    // Clear the batch
    batchBuffer.delete(batchKey);
    
    try {
        // Make ONE API call for all periods
        const response = await fetch(`${SERVER_URL}/batch/balance`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                account: batch.account,
                periods: periods,
                subsidiary: batch.subsidiary || "",
                department: batch.department || "",
                location: batch.location || "",
                class: batch.classId || ""
            })
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }
        
        const data = await response.json();
        const balances = data.balances || {};
        
        // Distribute results to all waiting cells
        for (let i = 0; i < periods.length; i++) {
            const period = periods[i];
            const resolvers = requests[i];
            const balance = balances[period] || 0;
            
            // Cache the result
            const cacheKey = getCacheKey('GLABAL', batch.account, period, period, 
                                        batch.subsidiary, batch.department, batch.location, batch.classId);
            setCachedResult(cacheKey, balance);
            
            // Resolve all promises for this period
            resolvers.forEach(resolve => resolve(balance));
        }
        
    } catch (error) {
        console.error(`Batch request failed: ${error.message}`);
        // Return 0 for all waiting cells on error
        requests.forEach(resolvers => {
            resolvers.forEach(resolve => resolve(0));
        });
    }
}


/**
 * Get account name from account number
 * @customfunction
 * @param {string} accountNumber The account number or ID
 * @param {string} [subsidiary] Subsidiary ID (optional)
 * @param {string} [department] Department ID (optional)
 * @param {string} [location] Location ID (optional)
 * @param {string} [classId] Class ID (optional)
 * @returns {string} The account name
 */
async function GLATITLE(accountNumber, subsidiary, department, location, classId) {
    if (!accountNumber) {
        throw new Error("Account number is required");
    }
    
    // Check cache first
    const cacheKey = getCacheKey('GLATITLE', accountNumber);
    const cached = getCachedResult(cacheKey);
    if (cached !== null) {
        return cached;
    }
    
    try {
        const response = await fetch(`${SERVER_URL}/account/${accountNumber}/name`);
        
        if (!response.ok) {
            return "Error";
        }
        
        const text = await response.text();
        setCachedResult(cacheKey, text);
        return text;
        
    } catch (error) {
        console.error(`Error getting account name: ${error.message}`);
        return "Error";
    }
}


/**
 * Get GL account balance
 * @customfunction
 * @param {string} account The account number or ID (required)
 * @param {string} fromPeriod Starting period name (e.g., "Jan 2025")
 * @param {string} toPeriod Ending period name (e.g., "Dec 2025")  
 * @param {string} [subsidiary] Subsidiary ID (optional)
 * @param {string} [department] Department ID (optional)
 * @param {string} [location] Location ID (optional)
 * @param {string} [classId] Class ID (optional)
 * @returns {number} The GL account balance
 */
async function GLABAL(account, fromPeriod, toPeriod, subsidiary, department, location, classId) {
    if (!account) {
        return 0;
    }
    
    // Convert to strings and handle defaults
    account = String(account || "");
    fromPeriod = String(fromPeriod || "");
    toPeriod = String(toPeriod || "");
    subsidiary = String(subsidiary || "");
    department = String(department || "");
    location = String(location || "");
    classId = String(classId || "");
    
    // For range queries (from != to), call API directly (can't batch easily)
    if (fromPeriod !== toPeriod) {
        const cacheKey = getCacheKey('GLABAL', account, fromPeriod, toPeriod, subsidiary, department, location, classId);
        const cached = getCachedResult(cacheKey);
        if (cached !== null) {
            return cached;
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
            
            const response = await fetch(`${SERVER_URL}/balance?${params.toString()}`);
            if (!response.ok) return 0;
            
            const result = await response.json();
            const balance = typeof result === 'number' ? result : parseFloat(result) || 0;
            setCachedResult(cacheKey, balance);
            return balance;
        } catch (error) {
            console.error(`Balance error: ${error.message}`);
            return 0;
        }
    }
    
    // For single period (from == to), use batching
    const period = fromPeriod;
    const cacheKey = getCacheKey('GLABAL', account, period, period, subsidiary, department, location, classId);
    
    // Check cache first
    const cached = getCachedResult(cacheKey);
    if (cached !== null) {
        return cached;
    }
    
    // Add to batch
    return new Promise((resolve) => {
        const batch = getBatch(account, subsidiary, department, location, classId);
        
        if (!batch.periods.has(period)) {
            batch.periods.set(period, []);
        }
        
        batch.periods.get(period).push(resolve);
        
        // Set timer to process batch after collecting more requests
        if (batch.timer) {
            clearTimeout(batch.timer);
        }
        
        batch.timer = setTimeout(() => {
            const batchKey = `${account}|${subsidiary}|${department}|${location}|${classId}`;
            processBatch(batchKey);
        }, BATCH_DELAY);
    });
}


/**
 * Get GL budget amount
 * @customfunction
 * @param {string} account The account number or ID (required)
 * @param {string} fromPeriod Starting period name (e.g., "Jan 2025")
 * @param {string} toPeriod Ending period name (e.g., "Dec 2025")
 * @param {string} [subsidiary] Subsidiary ID (optional)
 * @param {string} [department] Department ID (optional)
 * @param {string} [location] Location ID (optional)
 * @param {string} [classId] Class ID (optional)
 * @returns {number} The budget amount
 */
async function GLABUD(account, fromPeriod, toPeriod, subsidiary, department, location, classId) {
    if (!account) {
        return 0;
    }
    
    // Convert to strings
    account = String(account || "");
    fromPeriod = String(fromPeriod || "");
    toPeriod = String(toPeriod || "");
    subsidiary = String(subsidiary || "");
    department = String(department || "");
    location = String(location || "");
    classId = String(classId || "");
    
    const cacheKey = getCacheKey('GLABUD', account, fromPeriod, toPeriod, subsidiary, department, location, classId);
    const cached = getCachedResult(cacheKey);
    if (cached !== null) {
        return cached;
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
        
        const response = await fetch(`${SERVER_URL}/budget?${params.toString()}`);
        if (!response.ok) return 0;
        
        const result = await response.json();
        const budget = typeof result === 'number' ? result : parseFloat(result) || 0;
        setCachedResult(cacheKey, budget);
        return budget;
        
    } catch (error) {
        console.error(`Budget error: ${error.message}`);
        return 0;
    }
}


// Register custom functions with Excel
CustomFunctions.associate("GLATITLE", GLATITLE);
CustomFunctions.associate("GLABAL", GLABAL);
CustomFunctions.associate("GLABUD", GLABUD);
