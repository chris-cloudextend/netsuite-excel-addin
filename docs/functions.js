/**
 * NetSuite Custom Functions for Excel
 * Provides three custom formulas: NS.GLATITLE, NS.GLABAL, NS.GLABUD
 * 
 * NO CACHING - Every call makes a fresh API request for reliability
 */

// Backend server URL
const SERVER_URL = 'https://attention-birthday-cherry-shuttle.trycloudflare.com';

/**
 * Get account name from account number
 * @customfunction
 * @param {any} accountNumber The account number or ID
 * @returns {string} The account name
 */
async function GLATITLE(accountNumber, invocation) {
    // Convert to string and check if empty
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
            return "#N/A";
        }
        
        const text = await response.text();
        if (!text || text.trim() === "") {
            return "#N/A";
        }
        
        return text;
        
    } catch (error) {
        // Return #N/A text (Excel will recognize it)
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
async function GLABAL(account, fromPeriod, toPeriod, subsidiary, department, location, classId, invocation) {
    // Convert all to strings and trim
    account = String(account || "").trim();
    fromPeriod = String(fromPeriod || "").trim();
    toPeriod = String(toPeriod || "").trim();
    subsidiary = String(subsidiary || "").trim();
    department = String(department || "").trim();
    location = String(location || "").trim();
    classId = String(classId || "").trim();
    
    if (!account) {
        // Return empty string for blank cell - won't break SUM
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
        const response = await fetch(url, {
            method: 'GET',
            headers: { 'Accept': 'application/json' }
        });
        
        if (!response.ok) {
            const errorText = await response.text().catch(() => "");
            console.error(`Balance API error: ${response.status} - ${errorText}`);
            // Return empty string (blank cell) - SUM will ignore it
            return "";
        }
        
        const text = await response.text();
        const balance = parseFloat(text);
        
        // If parsing failed, return empty string (blank)
        if (isNaN(balance)) {
            return "";
        }
        
        return balance;
        
    } catch (error) {
        console.error(`Balance fetch error: ${error.message}`);
        // Return empty string (blank cell) - won't break SUM formulas
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
async function GLABUD(account, fromPeriod, toPeriod, subsidiary, department, location, classId, invocation) {
    // Convert all to strings and trim
    account = String(account || "").trim();
    fromPeriod = String(fromPeriod || "").trim();
    toPeriod = String(toPeriod || "").trim();
    subsidiary = String(subsidiary || "").trim();
    department = String(department || "").trim();
    location = String(location || "").trim();
    classId = String(classId || "").trim();
    
    if (!account) {
        // Return empty string for blank cell - won't break SUM
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
            console.error(`Budget API error: ${response.status} - ${errorText}`);
            // Return empty string (blank cell) - SUM will ignore it
            return "";
        }
        
        const text = await response.text();
        const budget = parseFloat(text);
        
        // If parsing failed, return empty string (blank)
        if (isNaN(budget)) {
            return "";
        }
        
        return budget;
        
    } catch (error) {
        console.error(`Budget fetch error: ${error.message}`);
        // Return empty string (blank cell) - won't break SUM formulas
        return "";
    }
}


// Register custom functions
CustomFunctions.associate("GLATITLE", GLATITLE);
CustomFunctions.associate("GLABAL", GLABAL);
CustomFunctions.associate("GLABUD", GLABUD);
