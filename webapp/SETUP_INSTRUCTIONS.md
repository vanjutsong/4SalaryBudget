# Mobile Budget Tracker Web App - Setup Instructions

A simple mobile-friendly web app for tracking transactions and viewing balances. This runs as a **separate** Apps Script project from your main budget system.

## Quick Setup (15 minutes)

### Step 1: Create a New Apps Script Project

1. Go to https://script.google.com
2. Click **"New Project"** (or the **+** button)
3. Name it: **"Budget Tracker Mobile"**
4. You'll get a new project with an empty script

### Step 2: Get the Script ID

1. In your new Apps Script project, click **Project Settings** (gear icon ⚙️)
2. Copy the **Script ID** (looks like: `1abc123...xyz`)
3. Open `webapp/.clasp.json` in this folder
4. Replace `YOUR_WEBAPP_SCRIPT_ID_HERE` with your Script ID
5. Save the file

### Step 3: Push Files with Clasp

Open a terminal/PowerShell in the `webapp` folder and run:

```bash
cd webapp
clasp push
```

This uploads:
- `WebApp.gs` - Backend code
- `TransactionForm.html` - Transaction form
- `Dashboard.html` - Dashboard
- `appsscript.json` - Project settings

### Step 4: Verify in Apps Script

```bash
clasp open
```

You should see:
- `WebApp.gs` (code file)
- `TransactionForm` (HTML file)
- `Dashboard` (HTML file)
- `appsscript.json` (manifest)

### Step 5: Deploy as Web App

1. In Apps Script editor, click **Deploy → New deployment**
2. Click the gear icon ⚙️ → Select **Web app**
3. Fill in:
   - **Description**: "Budget Tracker v1"
   - **Execute as**: "Me"
   - **Who has access**: "Anyone"
4. Click **Deploy**
5. **Click "Authorize access"** and follow the prompts
6. **Copy the Web App URL** (save this!)

### Step 6: Update HTML Files with URL

1. In Apps Script, open `TransactionForm`
2. Find: `const scriptUrl = 'YOUR_APPS_SCRIPT_WEB_APP_URL';`
3. Replace `YOUR_APPS_SCRIPT_WEB_APP_URL` with your Web App URL
4. Save

5. Do the same for `Dashboard`:
   - Find and replace the same line with your URL
   - Save

### Step 7: Redeploy

After updating URLs:
1. Click **Deploy → Manage deployments**
2. Click the pencil icon ✏️
3. Under "Version", select **New version**
4. Click **Deploy**

### Step 8: Test It!

1. Open your Web App URL in a browser
2. You should see the Dashboard
3. Click "Add Transaction" to test the form
4. Try adding a test transaction

### Step 9: Add to Phone Home Screen

**On Android (Chrome):**
1. Open the Dashboard URL in Chrome
2. Tap menu (3 dots) → **"Add to Home screen"**
3. Name it "Budget Tracker"

**On iPhone (Safari):**
1. Open the Dashboard URL in Safari
2. Tap Share button → **"Add to Home Screen"**
3. Name it "Budget Tracker"

Now it works like an app!

## How to Use

### Adding a Transaction

1. Open the app
2. Click **"Add Transaction"**
3. Fill in:
   - Date (defaults to today)
   - Description
   - Payment Mode
   - Category
   - Income or Expense
   - Amount
4. Click **Save Transaction**

### Viewing Dashboard

1. Open the app
2. See balances for:
   - Today
   - Tomorrow
   - Day After Tomorrow
   - End of This Week
   - End of Next Week
3. Click **Refresh** to update

## Files in This Folder

```
webapp/
├── .clasp.json           # Clasp config (put your Script ID here)
├── .claspignore          # Files to ignore
├── appsscript.json       # Apps Script manifest
├── WebApp.gs             # Backend code
├── TransactionForm.html  # Transaction form
├── Dashboard.html        # Dashboard
└── SETUP_INSTRUCTIONS.md # This file
```

## Clasp Commands

```bash
# Push local files to Apps Script
clasp push

# Pull from Apps Script to local
clasp pull

# Open in Apps Script editor
clasp open

# Check what will be pushed
clasp status
```

## Troubleshooting

### "Script ID not found"
- Make sure you updated `.clasp.json` with your actual Script ID
- Check there are no extra spaces or quotes

### "Permission denied" when saving transactions
- In Apps Script, click **Run** on any function once
- Click **Authorize access** and allow permissions

### "FinalTracker sheet not found"
- Make sure your main Google Sheet has a FinalTracker sheet
- Run "Generate Schedule" in your main budget system first

### Dashboard shows wrong data
- Verify the SPREADSHEET_ID in WebApp.gs matches your sheet
- Check that FinalTracker has RunningBalance data

## Two Separate Projects

You now have two Apps Script projects:

1. **Main Budget System** (original)
   - `4salarybudget.gs`, `schema.gs`
   - Runs from: `C:\Users\User\Code\4SalaryBudget`
   - Push with: `clasp push`

2. **Mobile Web App** (new)
   - `WebApp.gs`, HTML files
   - Runs from: `C:\Users\User\Code\4SalaryBudget\webapp`
   - Push with: `cd webapp && clasp push`

They share the same Google Sheet but run as separate Apps Script projects.

## Security Note

- The web app URL is accessible to anyone (if set to "Anyone")
- Anyone with the URL can add transactions
- For more security, use "Anyone with Google account" and only share with yourself
