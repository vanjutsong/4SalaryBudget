# Mobile Budget Tracker Web App

A separate Apps Script project for mobile transaction entry and dashboard.

## Quick Start

1. Create a new Apps Script project at https://script.google.com
2. Copy the Script ID from Project Settings
3. Paste it in `.clasp.json` (replace `YOUR_WEBAPP_SCRIPT_ID_HERE`)
4. Run `clasp push` from this folder
5. Deploy as web app
6. Update URLs in HTML files
7. Add to your phone's home screen

**Full instructions:** See `SETUP_INSTRUCTIONS.md`

## Files

- `WebApp.gs` - Backend code
- `TransactionForm.html` - Add transactions
- `Dashboard.html` - View balances
- `appsscript.json` - Project manifest
- `.clasp.json` - Clasp configuration

## Clasp Commands

```bash
clasp push   # Upload to Apps Script
clasp open   # Open in browser
clasp pull   # Download from Apps Script
```
