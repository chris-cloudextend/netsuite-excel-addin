# XAVI for NetSuite - Excel Add-in Files

This folder contains the Excel Add-in files hosted via GitHub Pages.

## Files

| File | Purpose |
|------|---------|
| `functions.js` | Custom function implementations + caching |
| `functions.json` | Function metadata for Excel |
| `functions.html` | Functions runtime page |
| `taskpane.html` | Task pane UI + refresh logic |
| `commands.html` | Ribbon commands page |
| `commands.js` | Command implementations |
| `index.html` | Landing page |
| `icon-*.png` | Add-in icons (16, 32, 64, 80px) |

## Custom Functions

| Function | Description |
|----------|-------------|
| `XAVI.BALANCE` | Get GL account balance |
| `XAVI.BUDGET` | Get budget amount |
| `XAVI.NAME` | Get account name |
| `XAVI.TYPE` | Get account type |
| `XAVI.PARENT` | Get parent account |
| `XAVI.RETAINEDEARNINGS` | Calculate Retained Earnings |
| `XAVI.NETINCOME` | Calculate Net Income |
| `XAVI.CTA` | Calculate CTA (multi-currency) |

## Backend Connection

Functions connect to a Flask backend via Cloudflare tunnel. The backend handles:
- NetSuite OAuth 1.0 authentication
- SuiteQL query execution
- Multi-currency consolidation via `BUILTIN.CONSOLIDATE`

## Deployment

Files are served from GitHub Pages. After pushing changes:
1. Wait ~1 minute for GitHub Pages to deploy
2. Bump manifest version for cache-busting
3. Reload the add-in in Excel

---

*Current Version: 3.0.5.75*
