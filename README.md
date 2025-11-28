# NetSuite Excel Formulas

A clean, modern Excel add-in that provides custom formulas for querying NetSuite data using SuiteQL.

## Features

This add-in provides three custom Excel formulas:

1. **=NS.GLATITLE(accountNumber)** - Get account name from account number
2. **=NS.GLABAL(subsidiary, account, fromPeriod, toPeriod, class, dept, location)** - Get GL account balance
3. **=NS.GLABUD(subsidiary, budgetCategory, account, fromPeriod, toPeriod, class, dept, location)** - Get budget amount

## Architecture

- **Flask Backend Server**: Handles OAuth authentication and SuiteQL queries to NetSuite
- **Excel Office.js Add-in**: Provides custom functions that call the Flask server
- **No VBA Required**: Modern web-based add-in that works on Windows, Mac, and Excel Online

## Quick Start

### 1. Install Python Dependencies

```bash
cd backend
pip install -r requirements.txt
```

### 2. Configure NetSuite Credentials

Edit `backend/netsuite_config.json` with your NetSuite OAuth credentials:

```json
{
    "account_id": "YOUR_ACCOUNT_ID",
    "consumer_key": "YOUR_CONSUMER_KEY",
    "consumer_secret": "YOUR_CONSUMER_SECRET",
    "token_id": "YOUR_TOKEN_ID",
    "token_secret": "YOUR_TOKEN_SECRET"
}
```

### 3. Start the Backend Server

```bash
cd backend
python server.py
```

The server will start on `http://localhost:5001`

### 4. Load the Excel Add-in

#### Option A: Excel Desktop (Sideloading)

1. Open Excel
2. Go to Insert > Add-ins > My Add-ins
3. Click "Upload My Add-in" 
4. Select `excel-addin/manifest.xml`

#### Option B: Excel Online

1. Upload `excel-addin/manifest.xml` to your OneDrive
2. In Excel Online: Insert > Add-ins > Upload My Add-in
3. Select the manifest file

### 5. Use the Formulas

Once the add-in is loaded, use the custom functions in your Excel cells:

```excel
=NS.GLATITLE("1000")
=NS.GLABAL(1, "4000", "Jan 2025", "Dec 2025", "", "", "")
=NS.GLABUD(1, "Operating", "6000", "Jan 2025", "Dec 2025", "", "", "")
```

## Formula Reference

### NS.GLATITLE(accountNumber)

Returns the account name for a given account number.

**Parameters:**
- `accountNumber`: Account number or internal ID (required)

**Example:**
```excel
=NS.GLATITLE("1000")
// Returns: "Cash - Operating Account"
```

### NS.GLABAL(subsidiary, account, fromPeriod, toPeriod, class, dept, location)

Returns the GL account balance for specified parameters.

**Parameters:**
- `subsidiary`: Subsidiary ID (use "" for all)
- `account`: Account number or ID (required)
- `fromPeriod`: Starting period name (e.g., "Jan 2025")
- `toPeriod`: Ending period name (e.g., "Dec 2025")
- `class`: Class ID (optional)
- `dept`: Department ID (optional)
- `location`: Location ID (optional)

**Example:**
```excel
=NS.GLABAL(1, "4000", "Jan 2025", "Dec 2025", "", "", "")
// Returns: 150000.00
```

### NS.GLABUD(subsidiary, budgetCategory, account, fromPeriod, toPeriod, class, dept, location)

Returns the budget amount for specified parameters.

**Parameters:**
- `subsidiary`: Subsidiary ID (use "" for all)
- `budgetCategory`: Budget category name (e.g., "Operating")
- `account`: Account number or ID (required)
- `fromPeriod`: Starting period name (e.g., "Jan 2025")
- `toPeriod`: Ending period name (e.g., "Dec 2025")
- `class`: Class ID (optional)
- `dept`: Department ID (optional)
- `location`: Location ID (optional)

**Example:**
```excel
=NS.GLABUD(1, "Operating", "6000", "Jan 2025", "Dec 2025", "", "", "")
// Returns: 120000.00
```

## Troubleshooting

### Server won't start
- Check that port 5000 is not in use
- Verify Python 3.7+ is installed

### Formulas return #NAME?
- Ensure the add-in is loaded (Insert > Add-ins > My Add-ins)
- Check that the server is running

### Formulas return errors
- Verify NetSuite credentials in `netsuite_config.json`
- Check server logs for authentication errors
- Ensure your NetSuite account has SuiteQL enabled

## Development

### Project Structure

```
NetSuite Formulas Revised/
├── README.md
├── backend/
│   ├── server.py              # Flask server
│   ├── requirements.txt       # Python dependencies
│   ├── netsuite_config.json  # NetSuite credentials
│   └── netsuite_config.template.json
└── excel-addin/
    ├── manifest.xml           # Add-in manifest
    ├── functions.html         # Custom functions page
    ├── functions.js           # Custom functions code
    └── taskpane.html          # Task pane UI (optional)
```

### Testing with Real NetSuite Data

The configuration file includes credentials for NetSuite test drive account (TSTDRV2320150).

Test queries:
- Account "1000" - typically a cash account
- Account "4000" - typically a revenue account
- Periods like "Jan 2025", "Feb 2025", etc.

## License

MIT License

