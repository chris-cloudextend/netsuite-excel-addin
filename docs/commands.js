/**
 * Commands.js - Handles ribbon and context menu actions
 * This file is loaded separately from the task pane for ExecuteFunction actions
 */

// Use Cloudflare Worker URL (not tunnel directly)
const SERVER_URL = 'https://netsuite-proxy.chris-corcoran.workers.dev';

console.log('✅ commands.js loaded');
console.log('   SERVER_URL:', SERVER_URL);

/**
 * Drill down from context menu (right-click)
 */
async function drillDownFromContextMenu(event) {
    console.log('=== CONTEXT MENU DRILL-DOWN START ===');
    
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(['formulas', 'values', 'address']);
            await context.sync();

            const formula = range.formulas[0][0];
            const address = range.address;

            console.log('Selected cell:', address);
            console.log('Formula:', formula);
            console.log('Value:', range.values[0][0]);

            // Check for XAVI.BALANCE formula
            if (!formula || !formula.toUpperCase().includes('XAVI.BALANCE')) {
                console.warn('Not an XAVI.BALANCE formula');
                if (event && event.completed) event.completed();
                return;
            }

            // Parse cell references from formula
            const cellRefs = parseFormulaCellRefs(formula);
            console.log('Cell references found:', cellRefs);

            // Resolve cell references to actual values
            const params = await resolveFormulaParams(context, cellRefs);
            console.log('Resolved parameters:', params);

            if (!params.account || !params.period) {
                console.error('Missing account or period');
                if (event && event.completed) event.completed();
                return;
            }

            // Construct API URL
            const queryParams = new URLSearchParams({
                account: params.account,
                period: params.period
            });
            
            if (params.subsidiary) queryParams.append('subsidiary', params.subsidiary);
            if (params.department) queryParams.append('department', params.department);
            if (params.location) queryParams.append('location', params.location);
            if (params.class) queryParams.append('class', params.class);

            const fetchUrl = `${SERVER_URL}/transactions?${queryParams}`;
            console.log('Fetching transactions from:', fetchUrl);

            // Fetch transaction data
            const response = await fetch(fetchUrl);
            console.log('Response status:', response.status);
            
            if (!response.ok) {
                console.error('API error:', response.status);
                if (event && event.completed) event.completed();
                return;
            }

            const data = await response.json();
            console.log('Transactions received:', data.transactions?.length || 0);

            if (!data.transactions || data.transactions.length === 0) {
                console.log('No transactions found');
                if (event && event.completed) event.completed();
                return;
            }

            // Create drill-down sheet
            await createDrillDownSheet(context, data.transactions, params);
            
        });
        
    } catch (error) {
        console.error('=== CONTEXT MENU DRILL-DOWN ERROR ===');
        console.error('Error:', error);
    } finally {
        if (event && event.completed) {
            event.completed();
        }
        console.log('=== CONTEXT MENU DRILL-DOWN END ===');
    }
}

/**
 * Parse cell references from XAVI.BALANCE formula
 */
function parseFormulaCellRefs(formula) {
    try {
        // Extract content between parentheses
        const match = formula.match(/XAVI\.BALANCE\s*\((.*)\)/i);
        if (!match) return null;
        
        const paramsStr = match[1];
        
        // Split by comma, handling quotes and nested functions
        const params = [];
        let current = '';
        let inQuotes = false;
        let parenDepth = 0;
        
        for (let i = 0; i < paramsStr.length; i++) {
            const char = paramsStr[i];
            
            if (char === '"') {
                inQuotes = !inQuotes;
                current += char;
            } else if (char === '(' && !inQuotes) {
                parenDepth++;
                current += char;
            } else if (char === ')' && !inQuotes) {
                parenDepth--;
                current += char;
            } else if (char === ',' && !inQuotes && parenDepth === 0) {
                params.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        if (current.trim()) {
            params.push(current.trim());
        }
        
        return {
            accountRef: params[0] || '',
            fromPeriodRef: params[1] || '',
            toPeriodRef: params[2] || '',
            subsidiaryRef: params[3] || '',
            departmentRef: params[4] || '',
            locationRef: params[5] || '',
            classRef: params[6] || ''
        };
        
    } catch (error) {
        console.error('Error parsing formula:', error);
        return null;
    }
}

/**
 * Resolve cell references to actual values
 */
async function resolveFormulaParams(context, cellRefs) {
    if (!cellRefs) return {};
    
    const cleanParam = (p) => {
        if (!p || p === '""' || p === '') return '';
        return p.replace(/^["']|["']$/g, '');
    };
    
    const isCellRef = (str) => {
        return /^[\$]?[A-Z]+[\$]?\d+$/.test(str);
    };
    
    const getValue = async (ref) => {
        if (!ref || ref === '""' || ref === '') return '';
        
        const cleaned = cleanParam(ref);
        
        // Check if it's a TEXT() formula - extract the cell reference
        const textMatch = cleaned.match(/^TEXT\s*\(\s*([\$]?[A-Z]+[\$]?\d+)\s*,/i);
        if (textMatch) {
            const cellRef = textMatch[1];
            console.log(`  Detected TEXT formula, extracting cell ref: ${cellRef}`);
            try {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const refRange = sheet.getRange(cellRef);
                refRange.load('values');
                await context.sync();
                
                const value = refRange.values[0][0];
                console.log(`  Resolved ${cellRef} to:`, value);
                return String(value || '');
            } catch (e) {
                console.warn(`Could not resolve TEXT formula cell reference ${cellRef}:`, e);
                return '';
            }
        }
        
        // If it's a simple cell reference
        if (isCellRef(cleaned)) {
            try {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const refRange = sheet.getRange(cleaned);
                refRange.load('values');
                await context.sync();
                
                const value = refRange.values[0][0];
                console.log(`  Resolved ${cleaned} to:`, value);
                return String(value || '');
            } catch (e) {
                console.warn(`Could not resolve cell reference ${cleaned}:`, e);
                return '';
            }
        }
        
        return cleaned;
    };
    
    const convertToPeriodName = (value) => {
        if (!value) return '';
        
        // If already a period string like "Jan 2024"
        if (typeof value === 'string' && /^[A-Za-z]{3}\s+\d{4}$/.test(value.trim())) {
            return value.trim();
        }
        
        // If it's an Excel date serial number
        const num = Number(value);
        if (!isNaN(num) && num > 40000 && num < 60000) {
            const excelEpoch = new Date(1899, 11, 30);
            const jsDate = new Date(excelEpoch.getTime() + num * 24 * 60 * 60 * 1000);
            
            const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
            const periodName = `${monthNames[jsDate.getMonth()]} ${jsDate.getFullYear()}`;
            console.log(`  📅 Converted Excel date ${value} to period: ${periodName}`);
            return periodName;
        }
        
        return String(value);
    };
    
    // Resolve all parameters
    const account = await getValue(cellRefs.accountRef);
    const fromPeriodRaw = await getValue(cellRefs.fromPeriodRef);
    const subsidiary = await getValue(cellRefs.subsidiaryRef);
    const department = await getValue(cellRefs.departmentRef);
    const location = await getValue(cellRefs.locationRef);
    const classId = await getValue(cellRefs.classRef);
    
    const period = convertToPeriodName(fromPeriodRaw);
    
    return {
        account: account,
        period: period,
        subsidiary: subsidiary,
        department: department,
        location: location,
        class: classId
    };
}

/**
 * Create drill-down sheet with transactions
 */
async function createDrillDownSheet(context, transactions, params) {
    const sheetName = `Drill_${params.account}_${params.period.replace(/\s+/g, '')}`.substring(0, 31);
    
    // Delete existing sheet if it exists
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    
    const existingSheet = sheets.items.find(s => s.name === sheetName);
    if (existingSheet) {
        existingSheet.delete();
        await context.sync();
    }
    
    // Create new sheet
    const newSheet = sheets.add(sheetName);
    newSheet.activate();
    await context.sync();
    
    // Prepare data
    const headers = ['Date', 'Type', 'Number', 'Entity', 'Memo', 'Debit', 'Credit', 'Net Amount'];
    const rows = transactions.map(t => [
        t.transaction_date || '',
        t.transaction_type || '',
        t.transaction_number || '',
        t.entity_name || '',
        t.memo || '',
        t.debit || 0,
        t.credit || 0,
        t.net_amount || 0
    ]);
    
    const allData = [headers, ...rows];
    
    // Write data
    const dataRange = newSheet.getRangeByIndexes(0, 0, allData.length, headers.length);
    dataRange.values = allData;
    
    // Format headers
    const headerRange = newSheet.getRangeByIndexes(0, 0, 1, headers.length);
    headerRange.format.fill.color = '#09235C';
    headerRange.format.font.color = 'white';
    headerRange.format.font.bold = true;
    
    // Add hyperlinks to transaction numbers
    for (let i = 0; i < transactions.length; i++) {
        if (transactions[i].netsuite_url) {
            const cell = newSheet.getRangeByIndexes(i + 1, 2, 1, 1);
            cell.hyperlink = {
                address: transactions[i].netsuite_url,
                screenTip: 'Open in NetSuite'
            };
            cell.format.font.color = '#0563C1';
            cell.format.font.underline = 'Single';
        }
    }
    
    // Format number columns
    const debitCreditRange = newSheet.getRangeByIndexes(1, 5, transactions.length, 3);
    debitCreditRange.numberFormat = [['#,##0.00']];
    
    // Auto-fit columns
    dataRange.format.autofitColumns();
    
    await context.sync();
    console.log('✅ Drill-down sheet created:', sheetName);
}

// Make function globally available for Office.js ExecuteFunction
if (typeof window !== 'undefined') {
    window.drillDownFromContextMenu = drillDownFromContextMenu;
}

// Also expose at global scope for ExecuteFunction
if (typeof globalThis !== 'undefined') {
    globalThis.drillDownFromContextMenu = drillDownFromContextMenu;
}

// Initialize Office.js
Office.onReady(() => {
    console.log('✅ commands.js ready - context menu function registered');
    console.log('   drillDownFromContextMenu available:', typeof drillDownFromContextMenu);
});
