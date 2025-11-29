/**
 * NetSuite Custom Functions for Excel
 * Provides three custom formulas: NS.GLATITLE, NS.GLABAL, NS.GLABUD
 * 
 * NO CACHING - Every call makes a fresh API request for reliability
 * WITH REQUEST THROTTLING - Limit concurrent calls to prevent overwhelming backend
 */

// Backend server URL
const SERVER_URL = 'https://attention-birthday-cherry-shuttle.trycloudflare.com';

// Request queue to prevent overwhelming the backend
const requestQueue = [];
let activeRequests = 0;
const MAX_CONCURRENT_REQUESTS = 3; // Allow 3 simultaneous calls

async function queuedFetch(url, options) {
    // Add to queue and wait for turn
    return new Promise((resolve, reject) => {
        requestQueue.push({ url, options, resolve, reject });
        processQueue();
    });
}

async function processQueue() {
    // If we're at max concurrent or queue is empty, wait
    if (activeRequests >= MAX_CONCURRENT_REQUESTS || requestQueue.length === 0) {
        return;
    }
    
    // Take next request from queue
    const { url, options, resolve, reject } = requestQueue.shift();
    activeRequests++;
    
    try {
        const response = await fetch(url, options);
        resolve(response);
    } catch (error) {
        reject(error);
    } finally {
        activeRequests--;
        // Process next item in queue
        processQueue();
    }
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
        return "#N/A";
    }
    
    try {
        const response = await queuedFetch(`${SERVER_URL}/account/${account}/name`, {
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
        
        const url = `${SERVER_URL}/balance?${params.toString()}`;
        const response = await queuedFetch(url, {
            method: 'GET',
            headers: { 'Accept': 'application/json' }
        });
        
        if (!response.ok) {
            const errorText = await response.text().catch(() => "");
            console.error(`GLABAL failed for ${account} (${fromPeriod}-${toPeriod}): ${response.status} - ${errorText}`);
            return "";
        }
        
        const text = await response.text();
        const balance = parseFloat(text);
        
        if (isNaN(balance)) {
            console.error(`GLABAL parsing failed for ${account}: got "${text}"`);
            return "";
        }
        
        return balance;
        
    } catch (error) {
        console.error(`GLABAL error for ${account}:`, error);
        return "";
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
        const response = await queuedFetch(url, {
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
