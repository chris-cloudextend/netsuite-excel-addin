/**
 * NetSuite Excel Add-in - Configuration
 * 
 * Centralized configuration for all deployment modes.
 * Change DEPLOYMENT_MODE to switch between environments.
 */

// ============================================================================
// DEPLOYMENT MODE - Change this to switch environments
// ============================================================================
// Options: 'azure' | 'cloudflare' | 'local'
const DEPLOYMENT_MODE = 'azure';

// ============================================================================
// ENVIRONMENT CONFIGURATIONS
// ============================================================================
const ENVIRONMENTS = {
    // Azure Functions (Consumption Plan) - Recommended for production
    azure: {
        name: 'Azure Functions',
        serverUrl: 'https://netsuite-excel-func.azurewebsites.net',
        frontendUrl: 'https://netsuiteexcelweb.z13.web.core.windows.net',
        portalUrl: 'https://portal.azure.com',
        description: 'Serverless backend on Azure (Consumption Plan)',
        statusLabel: 'Azure Functions',
        icon: '☁️'
    },
    
    // Cloudflare Workers + Tunnel - Alternative deployment
    cloudflare: {
        name: 'Cloudflare',
        serverUrl: 'https://netsuite-proxy.chris-corcoran.workers.dev',
        frontendUrl: 'https://chris-cloudextend.github.io/netsuite-excel-addin',
        portalUrl: 'https://dash.cloudflare.com',
        description: 'Cloudflare Workers with tunnel to local server',
        statusLabel: 'Cloudflare Tunnel',
        icon: '🔗'
    },
    
    // Local development
    local: {
        name: 'Local',
        serverUrl: 'http://localhost:5002',
        frontendUrl: 'http://localhost:3000',
        portalUrl: null,
        description: 'Local Flask development server',
        statusLabel: 'Local Server',
        icon: '💻'
    }
};

// ============================================================================
// ACTIVE CONFIGURATION (based on DEPLOYMENT_MODE)
// ============================================================================
const CONFIG = {
    // Current deployment mode
    mode: DEPLOYMENT_MODE,
    
    // Active environment settings
    env: ENVIRONMENTS[DEPLOYMENT_MODE],
    
    // Convenience accessors
    get serverUrl() { return this.env.serverUrl; },
    get frontendUrl() { return this.env.frontendUrl; },
    get portalUrl() { return this.env.portalUrl; },
    get name() { return this.env.name; },
    get description() { return this.env.description; },
    get statusLabel() { return this.env.statusLabel; },
    get icon() { return this.env.icon; },
    
    // All available environments (for UI dropdowns, etc.)
    environments: ENVIRONMENTS,
    
    // Version info
    version: '2.0.0.0',
    
    // Request settings
    requestTimeout: 30000,  // 30 seconds
    healthCheckTimeout: 15000,  // 15 seconds (Azure cold start can take time)
    
    // Feature flags
    features: {
        showToasts: true,
        enableDebugLogs: false,
        enableBatching: true
    },
    
    // API endpoints (relative to serverUrl)
    endpoints: {
        health: '/health',
        balance: '/api/balance',
        budget: '/api/budget',
        accountName: '/api/account/name',
        accountType: '/api/account/type',
        accounts: '/api/accounts',
        periods: '/api/periods',
        suiteql: '/api/suiteql',
        transactions: '/transactions',
        subsidiaries: '/subsidiaries',
        departments: '/departments',
        classes: '/classes',
        locations: '/locations',
        fullYearRefresh: '/full_year_refresh',
        bsFullYearRefresh: '/bs_full_year_refresh',
        refreshAccounts: '/refresh_accounts',
        retainedEarnings: '/retained_earnings',
        netIncome: '/net_income',
        cta: '/cta'
    },
    
    // Helper method to build full URL
    buildUrl(endpoint, params = {}) {
        const url = new URL(this.serverUrl + endpoint);
        Object.entries(params).forEach(([key, value]) => {
            if (value !== undefined && value !== null && value !== '') {
                url.searchParams.append(key, value);
            }
        });
        return url.toString();
    },
    
    // Helper to check if using Azure
    isAzure() {
        return this.mode === 'azure';
    },
    
    // Helper to check if using Cloudflare
    isCloudflare() {
        return this.mode === 'cloudflare';
    },
    
    // Helper to check if local
    isLocal() {
        return this.mode === 'local';
    },
    
    // Get error message based on deployment mode
    getConnectionErrorMessage() {
        switch (this.mode) {
            case 'azure':
                return 'Azure Functions may be cold starting or experiencing issues. Check Azure Portal.';
            case 'cloudflare':
                return 'Cloudflare tunnel may have expired. Check terminal for new URL.';
            case 'local':
                return 'Local server may not be running. Start with: python server.py';
            default:
                return 'Connection error. Check server status.';
        }
    },
    
    // Get timeout error message
    getTimeoutErrorMessage() {
        switch (this.mode) {
            case 'azure':
                return 'Request timeout. Azure Functions may be cold starting.';
            case 'cloudflare':
                return 'Tunnel timeout. The connection may have expired.';
            case 'local':
                return 'Request timeout. Check if server is responding.';
            default:
                return 'Request timeout.';
        }
    }
};

// Log active configuration on load
console.log(`📋 Config loaded: ${CONFIG.icon} ${CONFIG.name} (${CONFIG.mode})`);
console.log(`   Server: ${CONFIG.serverUrl}`);

// Export for use in other files (works in both browser and module contexts)
if (typeof window !== 'undefined') {
    window.CONFIG = CONFIG;
    window.ENVIRONMENTS = ENVIRONMENTS;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = { CONFIG, ENVIRONMENTS };
}

