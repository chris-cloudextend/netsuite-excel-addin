/**
 * NetSuite Custom Functions for Excel
 * Provides three custom formulas: NS.GLATITLE, NS.GLABAL, NS.GLABUD
 */

// Backend server URL
const SERVER_URL = 'https://attention-birthday-cherry-shuttle.trycloudflare.com';

// Cache for function results to improve performance
const functionCache = new Map();
const CACHE_DURATION = 30000; // 30 seconds

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
 * Get account name from account number
 * @customfunction
 * @param {string} accountNumber The account number or ID
 * @param {string} [subsidiary] Subsidiary ID (optional, use "" to ignore)
 * @param {string} [department] Department ID (optional, use "" to ignore)
 * @param {string} [location] Location ID (optional, use "" to ignore)
 * @param {string} [classId] Class ID (optional, use "" to ignore)
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
            const error = await response.json();
            throw new Error(error.error || `Server error: ${response.status}`);
        }
        
        const text = await response.text();
        
        // Cache the result
        setCachedResult(cacheKey, text);
        
        return text;
        
    } catch (error) {
        throw new Error(`Error getting account name: ${error.message}`);
    }
}


/**
 * Get GL account balance
 * @customfunction
 * @param {string} account The account number or ID (required)
 * @param {string} fromPeriod Starting period name (e.g., "Jan 2025")
 * @param {string} toPeriod Ending period name (e.g., "Dec 2025")
 * @param {string} [subsidiary] Subsidiary ID (optional, use "" to ignore)
 * @param {string} [department] Department ID (optional, use "" to ignore)
 * @param {string} [location] Location ID (optional, use "" to ignore)
 * @param {string} [classId] Class ID (optional, use "" to ignore)
 * @returns {number} The GL account balance
 */
async function GLABAL(account, fromPeriod, toPeriod, subsidiary, department, location, classId) {
    if (!account) {
        throw new Error("Account number is required");
    }
    
    // Convert parameters to strings and handle defaults
    subsidiary = subsidiary || "";
    department = department || "";
    location = location || "";
    classId = classId || "";
    fromPeriod = fromPeriod || "";
    toPeriod = toPeriod || "";
    
    // Check cache
    const cacheKey = getCacheKey('GLABAL', account, fromPeriod, toPeriod, subsidiary, department, location, classId);
    const cached = getCachedResult(cacheKey);
    if (cached !== null) {
        return cached;
    }
    
    try {
        // Build query string
        const params = new URLSearchParams();
        params.append('account', account);
        
        if (subsidiary && subsidiary !== "") params.append('subsidiary', subsidiary);
        if (fromPeriod && fromPeriod !== "") params.append('from_period', fromPeriod);
        if (toPeriod && toPeriod !== "") params.append('to_period', toPeriod);
        if (classId && classId !== "") params.append('class', classId);
        if (department && department !== "") params.append('department', department);
        if (location && location !== "") params.append('location', location);
        
        const response = await fetch(`${SERVER_URL}/balance?${params.toString()}`);
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || `Server error: ${response.status}`);
        }
        
        const result = await response.json();
        const balance = typeof result === 'number' ? result : parseFloat(result);
        
        // Cache the result
        setCachedResult(cacheKey, balance);
        
        return balance;
        
    } catch (error) {
        throw new Error(`Error getting balance: ${error.message}`);
    }
}


/**
 * Get GL budget amount
 * @customfunction
 * @param {string} account The account number or ID (required)
 * @param {string} fromPeriod Starting period name (e.g., "Jan 2025")
 * @param {string} toPeriod Ending period name (e.g., "Dec 2025")
 * @param {string} [subsidiary] Subsidiary ID (optional, use "" to ignore)
 * @param {string} [department] Department ID (optional, use "" to ignore)
 * @param {string} [location] Location ID (optional, use "" to ignore)
 * @param {string} [classId] Class ID (optional, use "" to ignore)
 * @returns {number} The budget amount
 */
async function GLABUD(account, fromPeriod, toPeriod, subsidiary, department, location, classId) {
    if (!account) {
        throw new Error("Account number is required");
    }
    
    // Convert parameters to strings and handle defaults
    subsidiary = subsidiary || "";
    department = department || "";
    location = location || "";
    classId = classId || "";
    fromPeriod = fromPeriod || "";
    toPeriod = toPeriod || "";
    
    // Check cache
    const cacheKey = getCacheKey('GLABUD', account, fromPeriod, toPeriod, subsidiary, department, location, classId);
    const cached = getCachedResult(cacheKey);
    if (cached !== null) {
        return cached;
    }
    
    try {
        // Build query string
        const params = new URLSearchParams();
        params.append('account', account);
        
        if (subsidiary && subsidiary !== "") params.append('subsidiary', subsidiary);
        if (fromPeriod && fromPeriod !== "") params.append('from_period', fromPeriod);
        if (toPeriod && toPeriod !== "") params.append('to_period', toPeriod);
        if (classId && classId !== "") params.append('class', classId);
        if (department && department !== "") params.append('department', department);
        if (location && location !== "") params.append('location', location);
        
        const response = await fetch(`${SERVER_URL}/budget?${params.toString()}`);
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || `Server error: ${response.status}`);
        }
        
        const result = await response.json();
        const budget = typeof result === 'number' ? result : parseFloat(result);
        
        // Cache the result
        setCachedResult(cacheKey, budget);
        
        return budget;
        
    } catch (error) {
        throw new Error(`Error getting budget: ${error.message}`);
    }
}


// Register custom functions with Excel
CustomFunctions.associate("GLATITLE", GLATITLE);
CustomFunctions.associate("GLABAL", GLABAL);
CustomFunctions.associate("GLABUD", GLABUD);
