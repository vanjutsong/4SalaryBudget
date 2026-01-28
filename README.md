# 4SalaryBudget

This repository contains the source code for the **4SalaryBudget** project.

## Getting started

1. Clone this repository (once you push it to a remote like GitHub).
2. Open the project folder `4SalaryBudget` in your editor (e.g. Cursor).

## Git workflow (basic)

- Make changes to files.
- Check status:
  - `git status`
- Stage files:
  - `git add .`
- Commit your work:
  - `git commit -m "Short description of what you changed"`

## Notes

- The `.gitignore` file is configured to ignore common OS, editor, and build artifacts.

# Salary Budget Tracker

A Google Apps Script-based budget tracking system with a separate mobile web app.

## Project Structure

This project contains **two separate Apps Script projects**:

### 1. Main Budget System (Root)
The core budget management system that handles recurring transactions, schedule generation, and balance calculations.

```
4SalaryBudget/
├── 4salarybudget.gs   # Main budget logic
├── schema.gs          # Data definitions
├── .clasp.json        # Clasp config (main project)
└── appsscript.json    # Manifest
```

**Push with:** `clasp push`

### 2. Mobile Web App (webapp/)
A separate mobile-friendly web app for quick transaction entry and dashboard viewing.

```
webapp/
├── WebApp.gs              # Web app backend
├── TransactionForm.html   # Transaction entry form
├── Dashboard.html         # Balance dashboard
├── .clasp.json            # Clasp config (separate project)
└── appsscript.json        # Manifest
```

**Push with:** `cd webapp && clasp push`

## Setup

### Main Budget System
Already configured. Just run `clasp push` from the root folder.

### Mobile Web App
1. Create a new Apps Script project
2. Copy the Script ID
3. Update `webapp/.clasp.json` with your Script ID
4. Run `cd webapp && clasp push`
5. Deploy as web app

See `webapp/SETUP_INSTRUCTIONS.md` for detailed steps.

## How They Work Together

- Both projects use the **same Google Sheet** (via SPREADSHEET_ID)
- Main system: Handles recurring transactions, generates FinalTracker
- Web app: Adds transactions to VariableExpenses, reads balances from FinalTracker
- They run as separate Apps Script projects but share data

## Quick Commands

```bash
# Main budget system
clasp push              # Upload main system
clasp open              # Open in browser

# Mobile web app
cd webapp
clasp push              # Upload web app
clasp open              # Open in browser
```
