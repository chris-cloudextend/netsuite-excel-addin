/**
 * NetSuite Custom Functions for Excel
 * Provides three custom formulas: NS.GLATITLE, NS.GLABAL, NS.GLABUD
 */

// Backend server URL
const SERVER_URL = 'https://attention-birthday-cherry-shuttle.trycloudflare.com';

// Simple cache - keep results for 5 minutes
const cache = {};
const CACHE_TTL = 300000; // 5 minutes

function getFromCache(key) {
    const item = cache[key];
    if (item && (Date.now() - item.time) < CACHE_TTL) {
        return item.value;
    }
    return undefined;
}

function setCache(key, value) {
    cache[key] = { value: value, time: Date.now() };
}

/**
 * Get account name from account number
 * @customfunction
 * @param {any} accountNumber The account number or ID
 * @returns {string} The account name
 */
async function GLATITLE(accountNumber) {
    // Convert to string and check if empty
    const account = String(accountNumber || "").trim();
    if (!account || account === "undefined" || account === "null") {
        return "Error: Account required";
    }
    
    const key = `TITLE:${account}`;
    const cached = getFromCache(key);
    if (cached !== undefined) return cached;
    
    try {
        const response = await fetch(`${SERVER_URL}/account/${account}/name`, {
            method: 'GET',
            headers: { 'Accept': 'text/plain' }
        });
        
        if (!response.ok) {
            return "Error";
        }
        
        const text = await response.text();
        setCache(key, text);
        return text;
        
    } catch (error) {
        return "Error";
    }
}


/**
 * Get GL account balance
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
    
    if (!account) return 0;
    
    // Create cache key
    const key = `BAL:${account}:${fromPeriod}:${toPeriod}:${subsidiary}:${department}:${location}:${classId}`;
    const cached = getFromCache(key);
    if (cached !== undefined) return cached;
    
    try {
        const params = new URLSearchParams();
        params.append('account', account);
        if (fromPeriod) params.append('from_period', fromPeriod);
        if (toPeriod) params.append('to_period', toPeriod);
        if (subsidiary) params.append('subsidiary', subsidiary);
        if (department) params.append('department', department);
        if (location) params.append('location', location);
        if (classId) params.append('class', classId);
        
        const url = `${SERVER_URL}/balance?${params.toString()}`;
        const response = await fetch(url, {
            method: 'GET',
            headers: { 'Accept': 'application/json' }
        });
        
        if (!response.ok) {
            return 0;
        }
        
        const text = await response.text();
        const balance = parseFloat(text) || 0;
        setCache(key, balance);
        return balance;
        
    } catch (error) {
        return 0;
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
    
    if (!account) return 0;
    
    // Create cache key
    const key = `BUD:${account}:${fromPeriod}:${toPeriod}:${subsidiary}:${department}:${location}:${classId}`;
    const cached = getFromCache(key);
    if (cached !== undefined) return cached;
    
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
            return 0;
        }
        
        const text = await response.text();
        const budget = parseFloat(text) || 0;
        setCache(key, budget);
        return budget;
        
    } catch (error) {
        return 0;
    }
}


// Register custom functions
CustomFunctions.associate("GLATITLE", GLATITLE);
CustomFunctions.associate("GLABAL", GLABAL);
CustomFunctions.associate("GLABUD", GLABUD);
