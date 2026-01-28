# Diagnostics Export Guide

## Quick Method (Recommended)

1. **Run the diagnostic function:**
   - Open your Google Sheet
   - Go to **Extensions → Apps Script**
   - Select `inspectSheetStructure` from the function dropdown
   - Click **Run** (▶️)
   - Or use the menu: **Recurring Transactions → Inspect Sheet Structure**

2. **Copy the output:**
   - A "Diagnostics" sheet will be created/updated
   - Column F contains the full text output
   - Select all of column F and copy it
   - Paste it into a new file: `diagnostics.txt` in this folder

## Automated Method (Optional)

If you want to automate the export:

1. **Install dependencies:**
   ```bash
   npm install googleapis
   ```

2. **Run the export script:**
   ```bash
   node export-diagnostics.js
   ```

   This will automatically read the Diagnostics sheet and create `diagnostics.txt` in this folder.

## What Gets Exported

The diagnostics include:
- All sheet names and their row/column counts
- Headers for each sheet
- Sample data (first 3 rows) from each sheet
- Special previews for `FinalTracker` and `CurrentBalance` sheets

This information helps understand your sheet structure for implementing features like the daily budget calculation.
