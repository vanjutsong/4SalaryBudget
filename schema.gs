/**
 * Schema Management for Budget Tracker
 * Handles sheet structure: headers, columns, and validation
 */

// Sheet header definitions - single source of truth for all sheet structures
const SCHEMA = {
  RecurExpenses: {
    headers: ['Description', 'Amount', 'Mode', 'Category', 'Frequency', 'Month', 'Day', 'StartDate', 'EndDate', 'Active'],
    description: 'Recurring expense definitions'
  },
  RecurIncome: {
    headers: ['Description', 'Amount', 'Mode', 'Category', 'Frequency', 'Month', 'Day', 'StartDate', 'EndDate', 'Active'],
    description: 'Recurring income definitions'
  },
  VariableExpenses: {
    headers: ['Date', 'Description', 'Mode', 'Category', 'Income', 'Debits', 'CCTransactionDate'],
    description: 'One-time and variable transactions'
  },
  FinalTracker: {
    headers: ['Date', 'Description', 'Mode', 'Category', 'Income', 'Debits', 'CCTransactionDate', 'RunningBalance', 'Edited'],
    description: 'Combined transaction tracker with running balance and edit tracking'
  },
  CurrentBalance: {
    headers: ['Account', 'Amount', 'MaintainBalance', 'SavingsAmount', 'Main'],
    description: 'Current account balances with reserved funds (mark one account Main=TRUE)'
  },
  Variables: {
    headers: ['Name', 'Value', 'LastUpdated'],
    description: 'Calculated values and settings'
  },
  Lookups: {
    headers: ['Category', 'Mode', 'Frequency', 'HighlightColumn', 'MatchType', 'MatchValue', 'HighlightColor'],
    description: 'Dropdown values and highlight configuration'
  },
  SavingsGoals: {
    headers: ['Description', 'TargetAmount', 'Frequency', 'Day', 'Mode', 'AllotmentAmount', 'TargetDate', 'CurrentProgress', 'Calculated', 'Active', 'Account'],
    description: 'Savings goals with allotment or deadline tracking, optionally linked to account'
  },
  Dashboard: {
    headers: [], // Freeform layout - formulas in column B
    description: 'Visual dashboard with daily budget formulas'
  },
  EmergencyFund: {
    headers: [], // Freeform layout - formulas in column B
    description: 'Emergency fund calculator based on monthly salary percentage'
  },
  SavingsChecklist: {
    headers: ['Date', 'Goal Description', 'Amount', 'Account', 'FinalTrackerRow'],
    description: 'Savings occurrences checklist from FinalTracker'
  }
};

// Default values for Lookups sheet (used when initializing)
const DEFAULT_LOOKUPS = {
  Category: ['Salary', 'Bonus', 'Groceries', 'Fast Food', 'Restaurant', 'Coffee Shop', 'Convenience Store', 'Food Delivery', 'Transportation', 'Fuel', 'Parking', 'Utilities', 'Entertainment', 'Digital Purchase', 'Shopping', 'Healthcare', 'Insurance', 'Subscription', 'Rent', 'Loan', 'Savings', 'Investment', 'Share', 'Adjustment', 'Other'],
  Mode: ['Cash', 'Bank', 'CreditCard', 'GCash', 'Maya'],
  Frequency: ['WEEKLY', 'N_DAY_IN_MONTH', 'LAST_DAY_OF_MONTH', 'ANNUAL', 'EVERY_N_DAYS']
};

// Valid values for specific columns (non-lookup values)
const VALID_VALUES = {
  Active: [true, false, 'TRUE', 'FALSE', 1, 0]
};

// ============================================================================
// MENU FUNCTIONS
// ============================================================================

/**
 * Adds Schema menu to the spreadsheet
 * Called by onOpen() in main file, or can be called standalone
 */
function addSchemaMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Schema')
    .addItem('View All Headers', 'viewAllHeaders')
    .addItem('Validate All Sheets', 'validateAllSheets')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Initialize Sheet')
      .addItem('RecurExpenses', 'initRecurExpenses')
      .addItem('RecurIncome', 'initRecurIncome')
      .addItem('VariableExpenses', 'initVariableExpenses')
      .addItem('FinalTracker', 'initFinalTracker')
      .addItem('CurrentBalance', 'initCurrentBalance')
      .addItem('Variables', 'initVariables')
      .addItem('Lookups', 'initLookups')
      .addItem('SavingsGoals', 'initSavingsGoals')
      .addItem('Dashboard', 'initDashboard')
      .addItem('EmergencyFund', 'initEmergencyFund'))
    .addItem('Initialize All Sheets', 'initializeAllSheets')
    .addSeparator()
    .addItem('Add Column...', 'promptAddColumn')
    .addItem('Rename Header...', 'promptRenameHeader')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Lookups')
      .addItem('View Lookup Values', 'viewLookupValues')
      .addItem('Add Lookup Value...', 'promptAddLookupValue')
      .addSeparator()
      .addItem('Setup All Dropdowns', 'setupAllDropdowns')
      .addItem('Clear All Dropdowns', 'clearAllDropdowns')
      .addSeparator()
      .addItem('Reset to Defaults', 'initLookups'))
    .addToUi();
}

// ============================================================================
// VIEW FUNCTIONS
// ============================================================================

/**
 * Display all expected headers for each sheet
 */
function viewAllHeaders() {
  let message = 'Expected Headers by Sheet:\n\n';
  
  for (const [sheetName, config] of Object.entries(SCHEMA)) {
    message += `📋 ${sheetName}\n`;
    message += `   ${config.description}\n`;
    message += `   Headers: ${config.headers.join(', ')}\n\n`;
  }
  
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Get the current headers from a sheet
 */
function getSheetHeaders(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return { exists: false, headers: [] };
  }
  
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    return { exists: true, headers: [] };
  }
  
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return { exists: true, headers: headers };
}

/**
 * View headers for a specific sheet with comparison to expected
 */
function viewSheetHeaders(sheetName) {
  const expected = SCHEMA[sheetName];
  if (!expected) {
    SpreadsheetApp.getUi().alert(`Unknown sheet: ${sheetName}`);
    return;
  }
  
  const current = getSheetHeaders(sheetName);
  
  let message = `Sheet: ${sheetName}\n`;
  message += `${expected.description}\n\n`;
  
  if (!current.exists) {
    message += `⚠️ Sheet does not exist!\n\n`;
    message += `Expected headers:\n${expected.headers.join(', ')}`;
  } else if (current.headers.length === 0) {
    message += `⚠️ Sheet is empty (no headers)\n\n`;
    message += `Expected headers:\n${expected.headers.join(', ')}`;
  } else {
    message += `Current headers:\n${current.headers.join(', ')}\n\n`;
    message += `Expected headers:\n${expected.headers.join(', ')}\n\n`;
    
    // Compare
    const missing = expected.headers.filter(h => !current.headers.includes(h));
    const extra = current.headers.filter(h => !expected.headers.includes(h));
    
    if (missing.length === 0 && extra.length === 0) {
      message += `✅ Headers match expected schema`;
    } else {
      if (missing.length > 0) {
        message += `❌ Missing: ${missing.join(', ')}\n`;
      }
      if (extra.length > 0) {
        message += `➕ Extra: ${extra.join(', ')}`;
      }
    }
  }
  
  SpreadsheetApp.getUi().alert(message);
}

// ============================================================================
// VALIDATION FUNCTIONS
// ============================================================================

/**
 * Validate all sheets against expected schema
 */
function validateAllSheets() {
  const results = [];
  
  for (const sheetName of Object.keys(SCHEMA)) {
    const result = validateSheet(sheetName);
    results.push(result);
  }
  
  let message = 'Schema Validation Results:\n\n';
  
  results.forEach(r => {
    const icon = r.valid ? '✅' : '❌';
    message += `${icon} ${r.sheetName}: ${r.status}\n`;
    if (r.issues.length > 0) {
      r.issues.forEach(issue => {
        message += `   • ${issue}\n`;
      });
    }
  });
  
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Validate a single sheet against expected schema
 */
function validateSheet(sheetName) {
  const expected = SCHEMA[sheetName];
  const result = {
    sheetName: sheetName,
    valid: true,
    status: 'OK',
    issues: []
  };
  
  if (!expected) {
    result.valid = false;
    result.status = 'Unknown sheet';
    return result;
  }
  
  const current = getSheetHeaders(sheetName);
  
  if (!current.exists) {
    result.valid = false;
    result.status = 'Sheet not found';
    result.issues.push('Sheet does not exist');
    return result;
  }
  
  if (current.headers.length === 0) {
    result.valid = false;
    result.status = 'Empty';
    result.issues.push('No headers found');
    return result;
  }
  
  // Check for missing headers
  const missing = expected.headers.filter(h => !current.headers.includes(h));
  if (missing.length > 0) {
    result.valid = false;
    result.issues.push(`Missing: ${missing.join(', ')}`);
  }
  
  // Check for extra headers (warning, not error)
  const extra = current.headers.filter(h => !expected.headers.includes(h));
  if (extra.length > 0) {
    result.issues.push(`Extra columns: ${extra.join(', ')}`);
  }
  
  // Check header order
  let orderCorrect = true;
  for (let i = 0; i < expected.headers.length && i < current.headers.length; i++) {
    if (expected.headers[i] !== current.headers[i]) {
      orderCorrect = false;
      break;
    }
  }
  if (!orderCorrect && missing.length === 0) {
    result.issues.push('Header order differs from expected');
  }
  
  result.status = result.valid ? 'OK' : 'Issues found';
  return result;
}

// ============================================================================
// INITIALIZATION FUNCTIONS
// ============================================================================

/**
 * Initialize a sheet with expected headers
 * Creates sheet if it doesn't exist, sets headers if empty
 */
function initializeSheet(sheetName, overwrite = false) {
  const expected = SCHEMA[sheetName];
  if (!expected) {
    SpreadsheetApp.getUi().alert(`Unknown sheet: ${sheetName}`);
    return false;
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  
  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  
  // Check if sheet has data
  if (lastRow > 0 && !overwrite) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Sheet has data',
      `${sheetName} already has data. Do you want to overwrite the headers?\n\n` +
      `This will only change row 1 (headers), not your data.`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return false;
    }
  }
  
  // Set headers
  const headers = expected.headers;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  // Clear any extra columns in header row
  if (lastCol > headers.length) {
    sheet.getRange(1, headers.length + 1, 1, lastCol - headers.length).clearContent();
  }
  
  SpreadsheetApp.getUi().alert(`✅ ${sheetName} headers initialized!\n\nHeaders: ${headers.join(', ')}`);
  return true;
}

/**
 * Initialize all sheets
 */
function initializeAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Initialize All Sheets',
    'This will set up headers for all sheets:\n\n' +
    Object.keys(SCHEMA).join(', ') + '\n\n' +
    'Existing data will be preserved. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  const results = [];
  for (const sheetName of Object.keys(SCHEMA)) {
    const success = initializeSheetSilent(sheetName);
    results.push({ sheetName, success });
  }
  
  let message = 'Initialization Results:\n\n';
  results.forEach(r => {
    const icon = r.success ? '✅' : '⚠️';
    message += `${icon} ${r.sheetName}\n`;
  });
  
  ui.alert(message);
}

/**
 * Initialize sheet without prompts (for batch operations)
 */
function initializeSheetSilent(sheetName) {
  const expected = SCHEMA[sheetName];
  if (!expected) return false;
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  const headers = expected.headers;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  return true;
}

// Individual sheet initializers for menu
function initRecurExpenses() { initializeSheet('RecurExpenses'); }
function initRecurIncome() { initializeSheet('RecurIncome'); }
function initVariableExpenses() { initializeSheet('VariableExpenses'); }
function initFinalTracker() { initializeSheet('FinalTracker'); }
function initCurrentBalance() { initializeSheet('CurrentBalance'); }
function initVariables() { initializeSheet('Variables'); }
function initLookups() { initializeLookupsSheet(); }
function initSavingsGoals() { initializeSheet('SavingsGoals'); }
function initDashboard() { setupDashboard(); }
function initEmergencyFund() { setupEmergencyFundCalculator(); }
function initSavingsChecklist() { initializeSheet('SavingsChecklist'); }

// ============================================================================
// COLUMN MANAGEMENT FUNCTIONS
// ============================================================================

/**
 * Add a new column to a sheet
 */
function addColumn(sheetName, headerName, position = -1) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet not found: ${sheetName}`);
    return false;
  }
  
  const lastCol = sheet.getLastColumn();
  
  // If position is -1 or greater than lastCol, add at end
  if (position === -1 || position > lastCol) {
    position = lastCol + 1;
  }
  
  // Insert column at position
  if (position <= lastCol) {
    sheet.insertColumnBefore(position);
  }
  
  // Set header
  sheet.getRange(1, position).setValue(headerName);
  sheet.getRange(1, position).setFontWeight('bold');
  
  return true;
}

/**
 * Prompt user to add a column
 */
function promptAddColumn() {
  const ui = SpreadsheetApp.getUi();
  
  // Get sheet name
  const sheetResponse = ui.prompt(
    'Add Column',
    'Enter sheet name:\n(' + Object.keys(SCHEMA).join(', ') + ')',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (sheetResponse.getSelectedButton() !== ui.Button.OK) return;
  const sheetName = sheetResponse.getResponseText().trim();
  
  // Get header name
  const headerResponse = ui.prompt(
    'Add Column',
    'Enter new column header name:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (headerResponse.getSelectedButton() !== ui.Button.OK) return;
  const headerName = headerResponse.getResponseText().trim();
  
  if (!headerName) {
    ui.alert('Header name cannot be empty.');
    return;
  }
  
  // Get position
  const posResponse = ui.prompt(
    'Add Column',
    'Enter column position (1, 2, 3...) or leave empty to add at end:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (posResponse.getSelectedButton() !== ui.Button.OK) return;
  const posText = posResponse.getResponseText().trim();
  const position = posText ? parseInt(posText, 10) : -1;
  
  if (addColumn(sheetName, headerName, position)) {
    ui.alert(`✅ Column "${headerName}" added to ${sheetName}`);
  }
}

/**
 * Rename a header in a sheet
 */
function renameHeader(sheetName, oldName, newName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet not found: ${sheetName}`);
    return false;
  }
  
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    SpreadsheetApp.getUi().alert('Sheet has no headers.');
    return false;
  }
  
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colIndex = headers.indexOf(oldName);
  
  if (colIndex === -1) {
    SpreadsheetApp.getUi().alert(`Header "${oldName}" not found in ${sheetName}.`);
    return false;
  }
  
  sheet.getRange(1, colIndex + 1).setValue(newName);
  return true;
}

/**
 * Prompt user to rename a header
 */
function promptRenameHeader() {
  const ui = SpreadsheetApp.getUi();
  
  // Get sheet name
  const sheetResponse = ui.prompt(
    'Rename Header',
    'Enter sheet name:\n(' + Object.keys(SCHEMA).join(', ') + ')',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (sheetResponse.getSelectedButton() !== ui.Button.OK) return;
  const sheetName = sheetResponse.getResponseText().trim();
  
  // Show current headers
  const current = getSheetHeaders(sheetName);
  if (!current.exists) {
    ui.alert(`Sheet not found: ${sheetName}`);
    return;
  }
  
  // Get old name
  const oldResponse = ui.prompt(
    'Rename Header',
    `Current headers: ${current.headers.join(', ')}\n\nEnter header to rename:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (oldResponse.getSelectedButton() !== ui.Button.OK) return;
  const oldName = oldResponse.getResponseText().trim();
  
  // Get new name
  const newResponse = ui.prompt(
    'Rename Header',
    `Rename "${oldName}" to:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (newResponse.getSelectedButton() !== ui.Button.OK) return;
  const newName = newResponse.getResponseText().trim();
  
  if (!newName) {
    ui.alert('New header name cannot be empty.');
    return;
  }
  
  if (renameHeader(sheetName, oldName, newName)) {
    ui.alert(`✅ Renamed "${oldName}" to "${newName}" in ${sheetName}`);
  }
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Get the expected schema for a sheet
 */
function getSchema(sheetName) {
  return SCHEMA[sheetName] || null;
}

/**
 * Get all valid values for a column
 */
function getValidValues(columnName) {
  return VALID_VALUES[columnName] || null;
}

/**
 * Get column index by header name (1-based)
 */
function getColumnIndex(sheetName, headerName) {
  const current = getSheetHeaders(sheetName);
  if (!current.exists) return -1;
  
  const index = current.headers.indexOf(headerName);
  return index === -1 ? -1 : index + 1; // Convert to 1-based
}

/**
 * Check if a value is valid for a column
 */
function isValidValue(columnName, value) {
  const valid = VALID_VALUES[columnName];
  if (!valid) return true; // No restrictions
  return valid.includes(value);
}

// ============================================================================
// LOOKUPS SHEET FUNCTIONS
// ============================================================================

/**
 * Initialize the Lookups sheet with default values
 */
function initializeLookupsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Lookups');
  
  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('Lookups');
  }
  
  const lastRow = sheet.getLastRow();
  
  // Check if sheet has data
  if (lastRow > 1) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Lookups sheet has data',
      'Do you want to overwrite with default values?\n\n' +
      'This will replace lookup values in columns A–C. Existing highlight rules in columns E–H will be kept.',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return false;
    }
  }
  
  // Save existing E–H (highlight rules) before clearing
  let savedHighlightRules = [];
  if (lastRow >= 2) {
    const numDataRows = lastRow - 1;
    savedHighlightRules = sheet.getRange(2, 5, numDataRows, 4).getValues();
  }
  
  // Clear existing content
  sheet.clear();
  
  // Set headers
  const headers = SCHEMA.Lookups.headers;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  // Find the longest column
  const maxRows = Math.max(
    DEFAULT_LOOKUPS.Category.length,
    DEFAULT_LOOKUPS.Mode.length,
    DEFAULT_LOOKUPS.Frequency.length
  );
  
  // Build data array
  const data = [];
  for (let i = 0; i < maxRows; i++) {
    data.push([
      DEFAULT_LOOKUPS.Category[i] || '',
      DEFAULT_LOOKUPS.Mode[i] || '',
      DEFAULT_LOOKUPS.Frequency[i] || ''
    ]);
  }
  
  // Write lookup data (columns A–C)
  sheet.getRange(2, 1, data.length, 3).setValues(data);
  
  // Restore E–H (highlight rules) without overwriting
  if (savedHighlightRules.length > 0) {
    sheet.getRange(2, 5, savedHighlightRules.length, 4).setValues(savedHighlightRules);
  }
  
  sheet.autoResizeColumns(1, 8);
  
  SpreadsheetApp.getUi().alert(
    `✅ Lookups sheet initialized!\n\n` +
    `Categories: ${DEFAULT_LOOKUPS.Category.length}\n` +
    `Modes: ${DEFAULT_LOOKUPS.Mode.length}\n` +
    `Frequencies: ${DEFAULT_LOOKUPS.Frequency.length}\n\n` +
    `Set highlight rules directly in columns E–H (HighlightColumn, MatchType, MatchValue, HighlightColor).`
  );
  
  return true;
}

/**
 * Get lookup values from the Lookups sheet
 * Returns { Category: [...], Mode: [...], Frequency: [...] }
 */
function getLookupValues() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Lookups');
  
  const result = {
    Category: [],
    Mode: [],
    Frequency: []
  };
  
  if (!sheet) {
    // Return defaults if Lookups sheet doesn't exist
    return DEFAULT_LOOKUPS;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return DEFAULT_LOOKUPS;
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  
  data.forEach(row => {
    if (row[0] && row[0].toString().trim()) result.Category.push(row[0].toString().trim());
    if (row[1] && row[1].toString().trim()) result.Mode.push(row[1].toString().trim());
    if (row[2] && row[2].toString().trim()) result.Frequency.push(row[2].toString().trim());
  });
  
  // If any column is empty, use defaults
  if (result.Category.length === 0) result.Category = DEFAULT_LOOKUPS.Category;
  if (result.Mode.length === 0) result.Mode = DEFAULT_LOOKUPS.Mode;
  if (result.Frequency.length === 0) result.Frequency = DEFAULT_LOOKUPS.Frequency;
  
  return result;
}

/**
 * Clear all data validations (dropdowns) from a sheet
 */
function clearDropdownsFromSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return false;
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow > 1 && lastCol > 0) {
    // Clear all data validations from data rows (row 2 onwards)
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearDataValidations();
  }
  
  return true;
}

/**
 * Set up dropdown validation for a specific sheet
 * Clears old dropdowns first, then applies new ones to correct columns
 */
function setupDropdownsForSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet not found: ${sheetName}`);
    return false;
  }
  
  // Clear ALL existing dropdowns first
  clearDropdownsFromSheet(sheetName);
  
  const lookups = getLookupValues();
  const dataLastRow = sheet.getLastRow();
  
  // Apply dropdowns to existing rows + 10 extra rows for new entries
  // If no data, start from row 2 and add 10 rows
  const startRow = 2;
  const lastRow = Math.max(dataLastRow, 1) + 10;
  
  // Define which columns get dropdowns based on sheet type
  // RecurExpenses/RecurIncome: Description(A), Amount(B), Mode(C), Category(D), Frequency(E), ...
  // VariableExpenses: Date(A), Description(B), Mode(C), Category(D), ...
  // SavingsGoals: Description(A), TargetAmount(B), Frequency(C), Day(D), Mode(E), ..., Active(J), Account(K)
  const dropdownConfig = {
    'RecurExpenses': { Mode: 3, Category: 4, Frequency: 5 },     // C, D, E
    'RecurIncome': { Mode: 3, Category: 4, Frequency: 5 },       // C, D, E
    'VariableExpenses': { Mode: 3, Category: 4 },                // C, D
    'SavingsGoals': { Frequency: 3, GoalMode: 5, Active: 10, Account: 11 }  // C, E, J, K
  };
  
  const config = dropdownConfig[sheetName];
  if (!config) {
    SpreadsheetApp.getUi().alert(`Dropdown configuration not found for: ${sheetName}`);
    return false;
  }
  
  // Apply dropdowns to correct columns
  let appliedCount = 0;
  
  if (config.Mode) {
    const modeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(lookups.Mode, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, config.Mode, lastRow - 1, 1).setDataValidation(modeRule);
    appliedCount++;
  }
  
  if (config.Category) {
    const categoryRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(lookups.Category, true)
      .setAllowInvalid(true) // Allow typing custom values
      .build();
    sheet.getRange(2, config.Category, lastRow - 1, 1).setDataValidation(categoryRule);
    appliedCount++;
  }
  
  if (config.Frequency) {
    const frequencyRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(lookups.Frequency, true)
      .setAllowInvalid(false) // Frequency must be from list
      .build();
    sheet.getRange(2, config.Frequency, lastRow - 1, 1).setDataValidation(frequencyRule);
    appliedCount++;
  }
  
  // SavingsGoals-specific dropdowns
  if (config.GoalMode) {
    const goalModeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['allotment', 'deadline'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, config.GoalMode, lastRow - 1, 1).setDataValidation(goalModeRule);
    appliedCount++;
  }
  
  if (config.Active) {
    const activeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['TRUE', 'FALSE'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, config.Active, lastRow - 1, 1).setDataValidation(activeRule);
    appliedCount++;
  }
  
  if (config.Account) {
    // Get account names from CurrentBalance sheet
    const currentBalanceSheet = ss.getSheetByName('CurrentBalance');
    let accountNames = [''];  // Empty option for no linked account
    if (currentBalanceSheet && currentBalanceSheet.getLastRow() > 1) {
      const cbData = currentBalanceSheet.getRange(2, 1, currentBalanceSheet.getLastRow() - 1, 1).getValues();
      cbData.forEach(row => {
        const name = (row[0] || '').toString().trim();
        if (name) accountNames.push(name);
      });
    }
    
    if (accountNames.length > 1) {
      const accountRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(accountNames, true)
        .setAllowInvalid(true)  // Allow empty or custom
        .build();
      sheet.getRange(2, config.Account, lastRow - 1, 1).setDataValidation(accountRule);
      appliedCount++;
    }
  }
  
  return appliedCount;
}

/**
 * Set up dropdowns for all applicable sheets
 */
function setupAllDropdowns() {
  const sheets = ['RecurExpenses', 'RecurIncome', 'VariableExpenses', 'SavingsGoals'];
  const results = [];
  
  sheets.forEach(sheetName => {
    const count = setupDropdownsForSheet(sheetName);
    results.push({ sheet: sheetName, dropdowns: count });
  });
  
  let message = 'Dropdowns set up!\n\n';
  results.forEach(r => {
    const icon = r.dropdowns > 0 ? '✅' : '⚠️';
    message += `${icon} ${r.sheet}: ${r.dropdowns} dropdown column(s)\n`;
  });
  
  message += '\nValues are pulled from the Lookups sheet.\n(Old dropdowns were cleared first)';
  
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Clear all dropdowns from all applicable sheets
 */
function clearAllDropdowns() {
  const sheets = ['RecurExpenses', 'RecurIncome', 'VariableExpenses', 'SavingsGoals'];
  
  sheets.forEach(sheetName => {
    clearDropdownsFromSheet(sheetName);
  });
  
  SpreadsheetApp.getUi().alert(
    `Dropdowns cleared from:\n\n` +
    sheets.join('\n') +
    `\n\nRun "Setup All Dropdowns" to re-apply them.`
  );
}

/**
 * Add a new value to a lookup column
 */
function addLookupValue(column, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Lookups');
  
  if (!sheet) {
    // Initialize if doesn't exist
    initializeLookupsSheet();
    sheet = ss.getSheetByName('Lookups');
  }
  
  const colIndex = { 'Category': 1, 'Mode': 2, 'Frequency': 3 }[column];
  if (!colIndex) {
    SpreadsheetApp.getUi().alert(`Invalid column: ${column}`);
    return false;
  }
  
  // Find the last used row in this column
  const lastRow = sheet.getLastRow();
  const colData = sheet.getRange(2, colIndex, Math.max(lastRow - 1, 1), 1).getValues();
  
  // Check if value already exists
  for (let i = 0; i < colData.length; i++) {
    if (colData[i][0] && colData[i][0].toString().toLowerCase() === value.toLowerCase()) {
      SpreadsheetApp.getUi().alert(`"${value}" already exists in ${column}.`);
      return false;
    }
  }
  
  // Find first empty row in this column
  let insertRow = -1;
  for (let i = 0; i < colData.length; i++) {
    if (!colData[i][0] || colData[i][0].toString().trim() === '') {
      insertRow = i + 2; // +2 because array is 0-indexed and data starts at row 2
      break;
    }
  }
  
  if (insertRow === -1) {
    insertRow = lastRow + 1;
  }
  
  sheet.getRange(insertRow, colIndex).setValue(value);
  return true;
}

/**
 * Prompt to add a new lookup value
 */
function promptAddLookupValue() {
  const ui = SpreadsheetApp.getUi();
  
  // Get column
  const colResponse = ui.prompt(
    'Add Lookup Value',
    'Enter column name:\n(Category, Mode, Frequency)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (colResponse.getSelectedButton() !== ui.Button.OK) return;
  const column = colResponse.getResponseText().trim();
  
  if (!['Category', 'Mode', 'Frequency'].includes(column)) {
    ui.alert('Invalid column. Must be Category, Mode, or Frequency.');
    return;
  }
  
  // Get value
  const valResponse = ui.prompt(
    'Add Lookup Value',
    `Enter new value for ${column}:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (valResponse.getSelectedButton() !== ui.Button.OK) return;
  const value = valResponse.getResponseText().trim();
  
  if (!value) {
    ui.alert('Value cannot be empty.');
    return;
  }
  
  if (addLookupValue(column, value)) {
    ui.alert(`✅ Added "${value}" to ${column}.\n\nRun "Setup All Dropdowns" to update the dropdown lists.`);
  }
}

/**
 * View current lookup values
 */
function viewLookupValues() {
  const lookups = getLookupValues();
  
  let message = 'Current Lookup Values\n';
  message += '='.repeat(40) + '\n\n';
  
  message += `📁 Categories (${lookups.Category.length}):\n`;
  message += lookups.Category.join(', ') + '\n\n';
  
  message += `💳 Modes (${lookups.Mode.length}):\n`;
  message += lookups.Mode.join(', ') + '\n\n';
  
  message += `🔄 Frequencies (${lookups.Frequency.length}):\n`;
  message += lookups.Frequency.join(', ');
  
  SpreadsheetApp.getUi().alert(message);
}
