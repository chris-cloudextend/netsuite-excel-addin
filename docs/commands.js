/**
 * XAVI for NetSuite - Ribbon and Context Menu Commands
 * 
 * Copyright (c) 2025 Celigo, Inc.
 * All rights reserved.
 * 
 * This source code is proprietary and confidential. Unauthorized copying,
 * modification, distribution, or use of this software, via any medium,
 * is strictly prohibited without the express written permission of Celigo, Inc.
 * 
 * For licensing inquiries, contact: legal@celigo.com
 * 
 * ---
 * 
 * This file handles ribbon and context menu actions.
 * It is loaded separately from the task pane for ExecuteFunction actions.
 */

// Use Cloudflare Worker URL (not tunnel directly)
const SERVER_URL = 'https://netsuite-proxy.chris-corcoran.workers.dev';

console.log('✅ commands.js loaded');
console.log('   SERVER_URL:', SERVER_URL);

/**
 * Drill down from context menu (right-click)
 * This is an ExecuteFunction action - must call event.completed() IMMEDIATELY
 * The debug window appears if event.completed() is delayed
 */
function drillDownFromContextMenu(event) {
    console.log('=== CONTEXT MENU DRILL-DOWN START ===');
    
    // CRITICAL: Call event.completed() IMMEDIATELY to close the function dialog
    // This tells Office "I've received the command" - actual work continues async
    if (event && event.completed) {
        event.completed();
        console.log('✅ event.completed() called immediately');
    }
    
    // Now do the actual work asynchronously
    performDrillDown();
}

/**
 * Perform the actual drill-down work (called after event.completed)
 */
async function performDrillDown() {
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

            const upperFormula = String(formula || '').toUpperCase();
            
            // Check for supported formula types
            const hasBalance = upperFormula.includes('XAVI.BALANCE');
            const hasTypeBalance = upperFormula.includes('XAVI.TYPEBALANCE');
            
            if (!hasBalance && !hasTypeBalance) {
                console.warn('Not a supported XAVI formula - skipping drill-down');
                // Show user feedback via alert since we don't have taskpane access
                return;
            }
            
            // For TYPEBALANCE, redirect to taskpane-style drill-down
            if (hasTypeBalance) {
                console.log('📊 TYPEBALANCE detected - using two-step drill-down');
                await handleTypeBalanceDrillDown(context, formula);
                return;
            }

            // Parse cell references from formula
            const cellRefs = parseFormulaCellRefs(formula);
            console.log('Cell references found:', cellRefs);

            if (!cellRefs) {
                console.error('Could not parse formula parameters');
                return;
            }

            // Resolve cell references to actual values
            const params = await resolveFormulaParams(context, cellRefs);
            console.log('Resolved parameters:', params);

            if (!params.account || !params.period) {
                console.error('Missing account or period - cannot drill down');
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
                return;
            }

            const data = await response.json();
            console.log('Transactions received:', data.transactions?.length || 0);

            if (!data.transactions || data.transactions.length === 0) {
                console.log('No transactions found for this period');
                return;
            }

            // Create drill-down sheet
            await createDrillDownSheet(context, data.transactions, params);
            console.log('✅ Drill-down sheet created successfully');
        });
    } catch (error) {
        console.error('Drill-down error:', error);
    }
    console.log('=== CONTEXT MENU DRILL-DOWN END ===');
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
    // Sanitize account for sheet name (Excel doesn't allow * in sheet names)
    const sanitizedAccount = params.account.replace(/\*/g, 'ALL');
    const periodShort = params.period.replace(/\s+/g, '');
    const sheetName = `Drill_${sanitizedAccount}_${periodShort}`.substring(0, 31);
    const isWildcard = params.account.includes('*');
    
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
    
    // Prepare headers - include Account column for wildcard drill-downs
    const headers = isWildcard
        ? ['Account', 'Date', 'Type', 'Number', 'Entity', 'Memo', 'Debit', 'Credit', 'Net Amount']
        : ['Date', 'Type', 'Number', 'Entity', 'Memo', 'Debit', 'Credit', 'Net Amount'];
    
    // Prepare data rows
    const rows = isWildcard
        ? transactions.map(t => [
            t.account_number || '',
            t.transaction_date || '',
            t.transaction_type || '',
            t.transaction_number || '',
            t.entity_name || '',
            t.memo || '',
            t.debit || 0,
            t.credit || 0,
            t.net_amount || 0
        ])
        : transactions.map(t => [
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
    
    // Add hyperlinks to transaction numbers (column varies based on wildcard)
    const linkColIndex = isWildcard ? 3 : 2;  // Number column
    for (let i = 0; i < transactions.length; i++) {
        if (transactions[i].netsuite_url) {
            const cell = newSheet.getRangeByIndexes(i + 1, linkColIndex, 1, 1);
            cell.hyperlink = {
                address: transactions[i].netsuite_url,
                screenTip: 'Open in NetSuite'
            };
            cell.format.font.color = '#0563C1';
            cell.format.font.underline = 'Single';
        }
    }
    
    // Format number columns (last 3 columns are Debit, Credit, Net Amount)
    const numColStartIndex = headers.length - 3;
    if (transactions.length > 0) {
        const numberRange = newSheet.getRangeByIndexes(1, numColStartIndex, transactions.length, 3);
        numberRange.numberFormat = [['#,##0.00']];
    }
    
    // Auto-fit columns
    dataRange.format.autofitColumns();
    
    await context.sync();
    console.log('✅ Drill-down sheet created:', sheetName);
}

/**
 * Handle TYPEBALANCE drill-down (shows accounts first, then transactions)
 */
async function handleTypeBalanceDrillDown(context, formula) {
    try {
        // Parse TYPEBALANCE formula parameters
        const match = formula.match(/XAVI\.TYPEBALANCE\s*\((.*)\)/i);
        if (!match) {
            console.error('Could not parse TYPEBALANCE formula');
            return;
        }
        
        const paramsStr = match[1];
        const params = parseTypeBalanceParams(paramsStr);
        const resolved = await resolveTypeBalanceParams(context, params);
        
        console.log('TYPEBALANCE params:', resolved);
        
        if (!resolved.accountType) {
            console.error('Could not determine account type');
            return;
        }
        
        // Query backend for all accounts of this type with their balances
        const queryParams = new URLSearchParams({
            account_type: resolved.accountType,
            to_period: resolved.toPeriod || '',
            subsidiary: resolved.subsidiary || '',
            use_special: resolved.useSpecialAccountType ? '1' : '0'
        });
        
        console.log('Fetching accounts by type:', queryParams.toString());
        
        const response = await fetch(`${SERVER_URL}/accounts/by-type?${queryParams}`);
        
        if (!response.ok) {
            console.error('API error:', response.status);
            return;
        }
        
        const data = await response.json();
        const accounts = data.accounts || [];
        
        if (accounts.length === 0) {
            console.log('No accounts found for type:', resolved.accountType);
            return;
        }
        
        console.log(`Found ${accounts.length} ${resolved.accountType} accounts`);
        
        // Create accounts sheet
        await createTypeBalanceAccountsSheet(context, accounts, resolved);
        console.log('✅ TYPEBALANCE accounts sheet created');
        
    } catch (error) {
        console.error('TYPEBALANCE drill-down error:', error);
    }
}

/**
 * Parse TYPEBALANCE formula parameters
 */
function parseTypeBalanceParams(paramsStr) {
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
    if (current.trim()) params.push(current.trim());
    
    return {
        accountTypeRef: params[0] || '',
        fromPeriodRef: params[1] || '',
        toPeriodRef: params[2] || '',
        subsidiaryRef: params[3] || '',
        departmentRef: params[4] || '',
        locationRef: params[5] || '',
        classRef: params[6] || '',
        accountingBookRef: params[7] || '',
        useSpecialRef: params[8] || ''
    };
}

/**
 * Resolve TYPEBALANCE parameter references to values
 */
async function resolveTypeBalanceParams(context, refs) {
    const cleanParam = (p) => {
        if (!p || p === '""' || p === '') return '';
        return p.replace(/^["']|["']$/g, '');
    };
    
    const isCellRef = (str) => /^[\$]?[A-Z]+[\$]?\d+$/.test(str);
    
    const getValue = async (ref) => {
        if (!ref || ref === '""' || ref === '') return '';
        const cleaned = cleanParam(ref);
        
        // Check for TEXT() formula - extract the cell reference
        const textMatch = cleaned.match(/^TEXT\s*\(\s*([\$]?[A-Z]+[\$]?\d+)\s*,/i);
        if (textMatch) {
            const cellRef = textMatch[1];
            try {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const refRange = sheet.getRange(cellRef);
                refRange.load('values');
                await context.sync();
                return String(refRange.values[0][0] || '');
            } catch (e) {
                return '';
            }
        }
        
        if (isCellRef(cleaned)) {
            try {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const refRange = sheet.getRange(cleaned);
                refRange.load('values');
                await context.sync();
                return String(refRange.values[0][0] || '');
            } catch (e) {
                return '';
            }
        }
        return cleaned;
    };
    
    return {
        accountType: await getValue(refs.accountTypeRef),
        fromPeriod: await getValue(refs.fromPeriodRef),
        toPeriod: await getValue(refs.toPeriodRef),
        subsidiary: await getValue(refs.subsidiaryRef),
        department: await getValue(refs.departmentRef),
        location: await getValue(refs.locationRef),
        classId: await getValue(refs.classRef),
        accountingBook: await getValue(refs.accountingBookRef),
        useSpecialAccountType: (await getValue(refs.useSpecialRef)) === '1'
    };
}

/**
 * Create TYPEBALANCE accounts drill-down sheet
 */
async function createTypeBalanceAccountsSheet(context, accounts, params) {
    const sheetName = `DrillDown_${params.accountType}`.substring(0, 31);
    
    // Delete existing sheet if present
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    
    const existingSheet = sheets.items.find(s => s.name === sheetName);
    if (existingSheet) {
        existingSheet.delete();
        await context.sync();
    }
    
    const drillSheet = sheets.add(sheetName);
    drillSheet.activate();
    await context.sync();
    
    // Header
    drillSheet.getRange('A1').values = [['TYPEBALANCE DRILL-DOWN']];
    drillSheet.getRange('A1').format.font.bold = true;
    drillSheet.getRange('A1').format.font.size = 14;
    
    const useSpecialLabel = params.useSpecialAccountType ? ' (Special Account Type)' : '';
    drillSheet.getRange('A2').values = [[`Account Type: ${params.accountType}${useSpecialLabel} | Period: ${params.toPeriod || 'All Time'}`]];
    drillSheet.getRange('A2').format.font.bold = true;
    
    // Store drill-down context for second-level drill-down
    // Extract year from toPeriod for full-year transaction queries
    const yearMatch = (params.toPeriod || '').match(/\d{4}/);
    const drilldownPeriod = yearMatch ? yearMatch[0] : params.toPeriod || '';
    drillSheet.getRange('A3').values = [[`DRILLDOWN_CONTEXT:${drilldownPeriod}:${params.subsidiary || ''}`]];
    drillSheet.getRange('A3').format.font.color = 'white'; // Hide it
    
    // Column headers
    const headers = [['Account', 'Account Name', 'Balance']];
    drillSheet.getRange('A5:C5').values = headers;
    drillSheet.getRange('A5:C5').format.font.bold = true;
    drillSheet.getRange('A5:C5').format.fill.color = '#667eea';
    drillSheet.getRange('A5:C5').format.font.color = 'white';
    
    if (accounts.length > 0) {
        const dataRows = accounts.map(acc => [
            acc.account_number || acc.acctnumber || '',
            acc.account_name || acc.accountname || '',
            acc.balance || 0
        ]);
        drillSheet.getRange(`A6:C${5 + accounts.length}`).values = dataRows;
        
        // Format balance column as currency
        drillSheet.getRange(`C6:C${5 + accounts.length}`).numberFormat = [['$#,##0.00']];
        
        // Make account numbers blue (clickable visual cue)
        drillSheet.getRange(`A6:A${5 + accounts.length}`).format.font.color = '#0ea5e9';
        
        // Add total row
        const totalRow = 6 + accounts.length;
        drillSheet.getRange(`A${totalRow}`).values = [['TOTAL']];
        drillSheet.getRange(`A${totalRow}`).format.font.bold = true;
        drillSheet.getRange(`C${totalRow}`).formulas = [[`=SUM(C6:C${totalRow - 1})`]];
        drillSheet.getRange(`C${totalRow}`).format.font.bold = true;
        drillSheet.getRange(`C${totalRow}`).numberFormat = [['$#,##0.00']];
    }
    
    drillSheet.getRange('A:C').format.autofitColumns();
    await context.sync();
}

// Register the function for Office.js ExecuteFunction
Office.onReady((info) => {
    console.log('✅ Office.onReady fired in commands.js');
    console.log('   Host:', info.host);
    console.log('   Platform:', info.platform);
    
    // Register the function at global scope
    if (typeof Office !== 'undefined' && Office.actions) {
        Office.actions.associate("drillDownFromContextMenu", drillDownFromContextMenu);
        console.log('✅ Function registered via Office.actions.associate');
    } else {
        console.log('⚠️ Office.actions not available, using window assignment');
    }
});

// Also make function globally available as fallback
if (typeof window !== 'undefined') {
    window.drillDownFromContextMenu = drillDownFromContextMenu;
}
if (typeof globalThis !== 'undefined') {
    globalThis.drillDownFromContextMenu = drillDownFromContextMenu;
}

console.log('✅ drillDownFromContextMenu registered:', typeof drillDownFromContextMenu);
