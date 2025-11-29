/**
 * Commands.js - Handles ribbon and context menu actions
 * This file is loaded separately from the task pane
 */

// Backend server URL
const SERVER_URL = 'https://pull-themes-friendly-mentor.trycloudflare.com';

console.log('✅ commands.js loaded');

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

            if (!formula || !formula.toUpperCase().includes('NS.GLABAL')) {
                console.error('Not a NS.GLABAL formula');
                event.completed();
                return;
            }

            // Parse cell references from formula
            const cellRefs = parseFormulaCellRefs(formula);
            console.log('Cell references found:', cellRefs);

            // Resolve cell references to actual values
            const params = await resolveFormulaParams(context, cellRefs);
            console.log('Resolved parameters:', params);

            // Construct API URL
            const url = new URL(`${SERVER_URL}/transactions`);
            url.searchParams.append('account', params.account || '');
            url.searchParams.append('period', params.fromPeriod || '');
            if (params.subsidiary) url.searchParams.append('subsidiary', params.subsidiary);
            if (params.department) url.searchParams.append('department', params.department);
            if (params.location) url.searchParams.append('location', params.location);
            if (params.classId) url.searchParams.append('class', params.classId);

            console.log('Fetching transactions from:', url.toString());

            // Fetch transaction data
            const response = await fetch(url.toString());
            if (!response.ok) {
                console.error('API error:', response.status);
                event.completed();
                return;
            }

            const data = await response.json();
            console.log('Transactions received:', data.transactions?.length || 0);

            if (!data.transactions || data.transactions.length === 0) {
                console.log('No transactions found');
                event.completed();
                return;
            }

            // Create drill-down sheet
            await createDrillDownSheet(context, data.transactions, params.account, params.fromPeriod);
            
        });
        
    } catch (error) {
        console.error('=== CONTEXT MENU DRILL-DOWN ERROR ===');
        console.error('Error:', error);
    } finally {
        if (event && event.completed) {
            event.completed();
        }
    }
}

/**
 * Parse cell references from formula text
 */
function parseFormulaCellRefs(formulaText) {
    const refs = {};
    const parts = formulaText.match(/=NS\.GLABAL\((.*)\)/i);
    if (!parts || !parts[1]) return refs;
    
    const args = parts[1].split(',').map(s => s.trim());
    
    if (args[0]) refs.account = args[0].replace(/['"$]/g, '');
    if (args[1]) refs.fromPeriod = args[1].replace(/['"$]/g, '');
    if (args[2]) refs.toPeriod = args[2].replace(/['"$]/g, '');
    if (args[3]) refs.subsidiary = args[3].replace(/['"$]/g, '');
    if (args[4]) refs.department = args[4].replace(/['"$]/g, '');
    if (args[5]) refs.location = args[5].replace(/['"$]/g, '');
    if (args[6]) refs.classId = args[6].replace(/['"$]/g, '');
    
    return refs;
}

/**
 * Resolve cell references to actual values
 */
async function resolveFormulaParams(context, cellRefs) {
    const params = {};
    
    for (const [key, ref] of Object.entries(cellRefs)) {
        if (!ref) continue;
        
        if (ref.match(/^[A-Z]+\$?\d+$/i)) {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(ref);
            range.load('values');
            await context.sync();
            params[key] = String(range.values[0][0] || '');
        } else {
            params[key] = ref;
        }
    }
    
    return params;
}

/**
 * Create drill-down sheet with transactions
 */
async function createDrillDownSheet(context, transactions, account, period) {
    const sheetName = `DrillDown_${account}_${period}`.substring(0, 31);
    
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
    const headers = ['Date', 'Type', 'Number', 'Entity', 'Memo', 'Debit', 'Credit'];
    const rows = transactions.map(t => [
        t.transaction_date || '',
        t.transaction_type || '',
        t.transaction_number || '',
        t.entity_name || '',
        t.memo || '',
        t.debit || 0,
        t.credit || 0
    ]);
    
    const allData = [headers, ...rows];
    
    // Write data
    const dataRange = newSheet.getRangeByIndexes(0, 0, allData.length, headers.length);
    dataRange.values = allData;
    
    // Format headers
    const headerRange = newSheet.getRangeByIndexes(0, 0, 1, headers.length);
    headerRange.format.fill.color = '#4472C4';
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
    
    // Auto-fit columns
    dataRange.format.autofitColumns();
    
    await context.sync();
    console.log('✅ Drill-down sheet created:', sheetName);
}

// Make function globally available
if (typeof window !== 'undefined') {
    window.drillDownFromContextMenu = drillDownFromContextMenu;
}

// Initialize Office.js
Office.onReady(() => {
    console.log('✅ commands.js ready - context menu function available');
});

