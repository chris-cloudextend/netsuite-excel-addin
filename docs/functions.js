/**
 * NetSuite Custom Functions for Excel
 * Provides three custom formulas: NS.GLATITLE, NS.GLABAL, NS.GLABUD
 */

// Backend server URL
const SERVER_URL = 'https://attention-birthday-cherry-shuttle.trycloudflare.com';


/**
 * Get account name from account number
 * @customfunction
 * @param {string} accountNumber The account number or ID
 * @param {string} subsidiary Subsidiary ID (optional, use "" to ignore)
 * @param {string} department Department ID (optional, use "" to ignore)
 * @param {string} location Location ID (optional, use "" to ignore)
 * @param {string} classId Class ID (optional, use "" to ignore)
 * @returns {string} The account name
 */
async function GLATITLE(accountNumber, subsidiary, department, location, classId) {
    if (!accountNumber) {
        throw new Error("Account number is required");
    }
    
    try {
        const response = await fetch(`${SERVER_URL}/account/${accountNumber}/name`);
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || `Server error: ${response.status}`);
        }
        
        const text = await response.text();
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
 * @param {string} subsidiary Subsidiary ID (optional, use "" to ignore)
 * @param {string} department Department ID (optional, use "" to ignore)
 * @param {string} location Location ID (optional, use "" to ignore)
 * @param {string} classId Class ID (optional, use "" to ignore)
 * @returns {number} The account balance
 */
async function GLABAL(account, fromPeriod, toPeriod, subsidiary, department, location, classId) {
    if (!account) {
        throw new Error("Account number is required");
    }
    
    try {
        // Build query string
        const params = new URLSearchParams();
        params.append('account', account);
        
        if (subsidiary) params.append('subsidiary', subsidiary);
        if (fromPeriod) params.append('from_period', fromPeriod);
        if (toPeriod) params.append('to_period', toPeriod);
        if (classId) params.append('class', classId);
        if (department) params.append('department', department);
        if (location) params.append('location', location);
        
        const response = await fetch(`${SERVER_URL}/balance?${params.toString()}`);
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || `Server error: ${response.status}`);
        }
        
        const text = await response.text();
        const balance = parseFloat(text);
        
        if (isNaN(balance)) {
            throw new Error(`Invalid balance returned: ${text}`);
        }
        
        return balance;
        
    } catch (error) {
        throw new Error(`Error getting balance: ${error.message}`);
    }
}


/**
 * Get budget amount
 * @customfunction
 * @param {string} account The account number or ID (required)
 * @param {string} fromPeriod Starting period name (e.g., "Jan 2025")
 * @param {string} toPeriod Ending period name (e.g., "Dec 2025")
 * @param {string} budgetCategory Budget category name (optional, e.g., "Operating")
 * @param {string} subsidiary Subsidiary ID (optional, use "" to ignore)
 * @param {string} department Department ID (optional, use "" to ignore)
 * @param {string} location Location ID (optional, use "" to ignore)
 * @param {string} classId Class ID (optional, use "" to ignore)
 * @returns {number} The budget amount
 */
async function GLABUD(account, fromPeriod, toPeriod, budgetCategory, subsidiary, department, location, classId) {
    if (!account) {
        throw new Error("Account number is required");
    }
    
    try {
        // Build query string
        const params = new URLSearchParams();
        params.append('account', account);
        
        if (subsidiary) params.append('subsidiary', subsidiary);
        if (budgetCategory) params.append('budget_category', budgetCategory);
        if (fromPeriod) params.append('from_period', fromPeriod);
        if (toPeriod) params.append('to_period', toPeriod);
        if (classId) params.append('class', classId);
        if (department) params.append('department', department);
        if (location) params.append('location', location);
        
        const response = await fetch(`${SERVER_URL}/budget?${params.toString()}`);
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || `Server error: ${response.status}`);
        }
        
        const text = await response.text();
        const budget = parseFloat(text);
        
        if (isNaN(budget)) {
            throw new Error(`Invalid budget returned: ${text}`);
        }
        
        return budget;
        
    } catch (error) {
        throw new Error(`Error getting budget: ${error.message}`);
    }
}


// Register functions (for Office.js)
CustomFunctions.associate("GLATITLE", GLATITLE);
CustomFunctions.associate("GLABAL", GLABAL);
CustomFunctions.associate("GLABUD", GLABUD);

