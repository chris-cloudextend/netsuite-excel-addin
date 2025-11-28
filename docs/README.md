# NetSuite Excel Add-in

This folder contains the NetSuite Excel Custom Functions add-in files, hosted via GitHub Pages.

## Files

- `functions.js` - Custom functions implementation
- `functions.json` - Functions metadata
- `functions.html` - Functions page
- `taskpane.html` - Task pane UI
- `index.html` - Landing page
- Icon files for the add-in

## Custom Functions

- `NS.GLATITLE(account)` - Get account title
- `NS.GLABAL(account, fromPeriod, toPeriod, ...)` - Get account balance
- `NS.GLABUD(account, fromPeriod, toPeriod, ...)` - Get budget amount

## Backend

These functions connect to a Flask backend server running locally that handles NetSuite authentication and SuiteQL queries.

The backend must be running on `http://localhost:5002` for the formulas to work.

