/**
 * Script to export Diagnostics sheet to a local file
 * Run: node export-diagnostics.js
 * 
 * Requires: npm install googleapis
 */

const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

const SPREADSHEET_ID = '1FQzuRQwlFrGGu10N8-ne3HYfbl9tbIoaWA2bqlE6bKo';
const DIAGNOSTICS_SHEET_NAME = 'Diagnostics';
const OUTPUT_FILE = 'diagnostics.txt';

async function exportDiagnostics() {
  try {
    // Use the same auth as clasp (from ~/.clasprc.json)
    const auth = new google.auth.GoogleAuth({
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });
    
    const sheets = google.sheets({ version: 'v4', auth });
    
    // Read the Diagnostics sheet
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${DIAGNOSTICS_SHEET_NAME}!F:F`, // Column F has the full output
    });
    
    const rows = response.data.values || [];
    const output = rows.map(row => row[0] || '').join('\n');
    
    // Write to local file
    const outputPath = path.join(__dirname, OUTPUT_FILE);
    fs.writeFileSync(outputPath, output, 'utf8');
    
    console.log(`✅ Diagnostics exported to: ${outputPath}`);
    console.log(`📄 ${rows.length} lines written`);
    
  } catch (error) {
    console.error('❌ Error exporting diagnostics:', error.message);
    console.log('\n💡 Make sure you have:');
    console.log('   1. Run "inspectSheetStructure" in Apps Script first');
    console.log('   2. Installed: npm install googleapis');
    console.log('   3. Authenticated with: clasp login');
    process.exit(1);
  }
}

exportDiagnostics();
