// Replace with your Google Sheet ID (from the URL: .../spreadsheets/d/SPREADSHEET_ID/edit)
const SPREADSHEET_ID = '1FQzuRQwlFrGGu10N8-ne3HYfbl9tbIoaWA2bqlE6bKo';

/**
 * Helper function to generate recurring transactions from RecurExpenses and RecurIncome
 * Returns an array of transactions (does NOT write to any sheet)
 * Used by generateRecurring() and duplicate checking
 */
/**
 * Generate savings goal transactions from SavingsGoals sheet
 * Creates recurring transactions based on Frequency, Day, and AllotmentAmount
 */
function generateSavingsGoalsTransactions(savingsGoalsSheet, accountSavingsMap) {
  const transactions = [];
  
  if (!savingsGoalsSheet || savingsGoalsSheet.getLastRow() <= 1) return transactions;
  
  // Read all 11 columns
  const data = savingsGoalsSheet.getRange(2, 1, savingsGoalsSheet.getLastRow() - 1, 11).getValues();
  
  // Get date range (today to 1 year from now, or use a reasonable default)
  const today = new Date();
  today.setHours(12, 0, 0, 0);
  const endDate = new Date(today);
  endDate.setFullYear(endDate.getFullYear() + 1);
  
  data.forEach(row => {
    const description = row[0];
    const targetAmount = Number(row[1]) || 0;
    const frequency = (row[2] || '').toString().toUpperCase();
    const day = row[3];
    const goalMode = (row[4] || '').toString().toLowerCase(); // allotment or deadline
    const allotmentAmount = Number(row[5]) || 0;
    const targetDate = row[6];
    let currentProgress = Number(row[7]) || 0;
    // Handle various Active formats
    const activeValue = row[9];
    const isActive = activeValue === true || activeValue === 'TRUE' || activeValue === 'true' || activeValue === 1 || activeValue === '1';
    const linkedAccount = (row[10] || '').toString().trim().toLowerCase();
    
    // Skip if not active or no description
    if (!isActive || !description) return;
    
    // If account is linked, use account's SavingsAmount as initial progress
    if (linkedAccount && accountSavingsMap && accountSavingsMap[linkedAccount] !== undefined) {
      currentProgress = accountSavingsMap[linkedAccount];
    }
    
    // For allotment mode, need AllotmentAmount; for deadline mode, calculate from target
    let amountPerOccurrence = allotmentAmount;
    if (goalMode === 'deadline' && targetDate && allotmentAmount <= 0) {
      // Calculate amount per occurrence from deadline
      const remaining = targetAmount - currentProgress;
      if (remaining <= 0) return;
      const occurrences = countOccurrences(today, new Date(targetDate), frequency, day);
      if (occurrences > 0) {
        amountPerOccurrence = remaining / occurrences;
      }
    }
    
    // Skip if no amount to save
    if (amountPerOccurrence <= 0) return;
    
    // Skip if goal is already reached
    if (currentProgress >= targetAmount) return;
    
    // Generate dates based on frequency
    const dates = generateDates(frequency, day, null, today, endDate);
    
    // Calculate how many transactions until goal is reached
    const remaining = targetAmount - currentProgress;
    const maxTransactions = Math.ceil(remaining / amountPerOccurrence);
    
    dates.slice(0, maxTransactions).forEach(date => {
      transactions.push([
        date,           // Date
        description,    // Description (goal name)
        'Bank',         // Mode (default to Bank for savings)
        'Savings',      // Category
        '',             // Income (empty for savings)
        amountPerOccurrence, // Debits (savings reduces available balance)
        ''              // CCTransactionDate
      ]);
    });
  });
  
  return transactions;
}

function generateRecurringTransactions(recurExpensesSheet, recurIncomeSheet) {
  const scheduledTransactions = [];

  // Helper to process either income or expenses
  const processRecurringSheet = (sheet, isIncome) => {
    if (!sheet || sheet.getLastRow() <= 1) return;

    // A–J = 10 columns (Description, Amount, Mode, Category, Frequency, Month, Day, StartDate, EndDate, Active)
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();

    data.forEach(row => {
      const [description, amount, mode, category, frequency, month, day, startDate, endDate, active] = row;

      if (!active) return;
      if (!startDate || !endDate) return;

      const start = toSafeDate(startDate);
      const end = toSafeDate(endDate);
      if (start > end) return;

      const dates = generateDates(frequency, day, month, start, end);

      dates.forEach(date => {
        const isCC = mode === 'CreditCard';
        const txnDate = date; // original transaction date
        let ccDueDate = null;
        
        if (isCC) {
          const ccStatementDate = computeCcStatementDate(txnDate);
          if (ccStatementDate) {
            ccDueDate = computeCcDueDate(ccStatementDate);
          }
        }

        scheduledTransactions.push([
          isCC && ccDueDate ? ccDueDate : date,  // Date (due date for CC, txn date otherwise)
          description,                  // Description
          mode,                         // Mode
          category,                     // Category
          isIncome ? amount : '',       // Income
          isIncome ? '' : amount,       // Debits
          isCC ? txnDate : ''           // CCTransactionDate (keep original txn date)
        ]);
      });
    });
  };

  processRecurringSheet(recurExpensesSheet, false);
  processRecurringSheet(recurIncomeSheet, true);

  return scheduledTransactions;
}

/**
 * Main function to update FinalTracker with all transactions
 * Combines RecurExpenses, RecurIncome, VariableExpenses, and SavingsGoals into FinalTracker
 * Also normalizes CreditCard dates before processing
 */
function updateFinalTracker() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const recurExpensesSheet = ss.getSheetByName('RecurExpenses');
  const recurIncomeSheet = ss.getSheetByName('RecurIncome');
  const variableSheet = ss.getSheetByName('VariableExpenses') || ss.getSheetByName('VariableExpences');
  const finalTrackerSheet = ss.getSheetByName('FinalTracker');

  if (!finalTrackerSheet) {
    SpreadsheetApp.getUi().alert('FinalTracker sheet not found.');
    return;
  }
  
  // Normalize CreditCard dates in VariableExpenses before processing
  // This ensures CC transactions have correct due dates
  // Check for incomplete entries and warn if found
  const ccNormalized = normalizeVariableExpensesCcDatesWithCheck(ss, variableSheet);

  const today = toSafeDate(new Date());
  
  // Ensure headers are set correctly (9 columns including Edited)
  const headers = ['Date', 'Description', 'Mode', 'Category', 'Income', 'Debits', 'CCTransactionDate', 'RunningBalance', 'Edited'];
  finalTrackerSheet.getRange(1, 1, 1, 9).setValues([headers]);
  finalTrackerSheet.getRange(1, 1, 1, 9).setFontWeight('bold');

  // Check if row 2 is the "Starting Balance" row
  const finalLastRow = finalTrackerSheet.getLastRow();
  const row2Description = finalLastRow >= 2 ? finalTrackerSheet.getRange(2, 2).getValue() : '';
  const hasStartingBalanceRow = row2Description === 'Starting Balance';
  
  const dataStartRow = hasStartingBalanceRow ? 3 : 2;
  
  // Read existing FinalTracker data (including RunningBalance and Edited columns)
  const preservedRows = [];
  let existingStartingBalanceRow = null;
  
  if (finalLastRow >= 2) {
    // Read all 9 columns (or as many as exist)
    const numCols = Math.min(finalTrackerSheet.getLastColumn(), 9);
    const existingData = finalTrackerSheet.getRange(2, 1, finalLastRow - 1, numCols).getValues();
    
    existingData.forEach((row, idx) => {
      const rowDate = toSafeDate(row[0]);
      const description = row[1];
      const runningBalance = row[7] || ''; // Column H
      const edited = row[8]; // Column I (Edited)
      
      // Preserve Starting Balance row
      if (description === 'Starting Balance') {
        existingStartingBalanceRow = row;
        return;
      }
      
      // Skip rows without dates
      if (!rowDate) return;
      
      // Check if row should be preserved:
      // 1. Past date (before today)
      // 2. Edited = TRUE (or any truthy value)
      const isPast = rowDate < today;
      const isEdited = edited === true || edited === 'TRUE' || edited === true || edited === 1;
      
      if (isPast || isEdited) {
        // Preserve this row (ensure it has 9 columns)
        while (row.length < 9) row.push('');
        preservedRows.push(row);
      }
    });
  }

  // Build edited-preserved list: preserved rows with Edited=TRUE (user changed date/amount etc.)
  // Used to suppress the matching recurring instance and avoid duplicates (e.g. Salary B moved to next day)
  const editedPreserved = [];
  preservedRows.forEach(preserved => {
    const edited = preserved[8];
    const isEdited = edited === true || edited === 'TRUE' || edited === 'true' || edited === 1;
    if (!isEdited) return;
    const d = toSafeDate(preserved[0]);
    if (!d) return;
    const desc = (preserved[1] || '').toString().trim();
    const mode = (preserved[2] || '').toString().trim();
    const income = Math.round((Number(preserved[4]) || 0) * 100);
    const debits = Math.round((Number(preserved[5]) || 0) * 100);
    const sig = `${desc}|${mode}|${income}|${debits}`;
    editedPreserved.push({ sig, date: d, desc, mode });
  });

  // Generate recurring transactions for ALL dates (we'll filter later)
  const allRecurringTransactions = generateRecurringTransactions(recurExpensesSheet, recurIncomeSheet);
  
  // Filter to only future recurring transactions (today and later)
  const futureRecurringRaw = [];
  allRecurringTransactions.forEach(row => {
    const rowDate = toSafeDate(row[0]);
    if (rowDate && rowDate >= today) {
      // Add empty columns for RunningBalance and Edited (will be set by formulas/user)
      while (row.length < 7) row.push('');
      row.push(''); // RunningBalance (column 8) - will be set by formula
      row.push(''); // Edited (column 9) - empty by default
      futureRecurringRaw.push(row);
    }
  });

  // Skip recurring instances that match an edited-preserved row (same desc/mode/income/debits)
  // to avoid duplicates when user e.g. moves Salary B to the next day. Match by "closest date" within 5 days.
  const EDIT_MATCH_DAYS = 5;
  const recurringWithMeta = futureRecurringRaw.map((row, i) => {
    const d = toSafeDate(row[0]);
    const desc = (row[1] || '').toString().trim();
    const mode = (row[2] || '').toString().trim();
    const income = Math.round((Number(row[4]) || 0) * 100);
    const debits = Math.round((Number(row[5]) || 0) * 100);
    const sig = `${desc}|${mode}|${income}|${debits}`;
    return { row, date: d, sig, index: i };
  });
  const skipRecurringIndices = new Set();
  editedPreserved.forEach(ep => {
    const msPerDay = 24 * 60 * 60 * 1000;
    let best = { index: -1, diff: Infinity };
    // Consider all matches (same sig, within 5 days); skip the closest by date, not the first found
    recurringWithMeta.forEach(m => {
      if (m.sig !== ep.sig || skipRecurringIndices.has(m.index)) return;
      const diff = Math.abs(m.date.getTime() - ep.date.getTime());
      if (diff <= EDIT_MATCH_DAYS * msPerDay && diff < best.diff) {
        best = { index: m.index, diff };
      }
    });
    if (best.index >= 0) skipRecurringIndices.add(best.index);
  });
  const futureRecurringTransactions = futureRecurringRaw.filter((_, i) => !skipRecurringIndices.has(i));

  // Build a Set of preserved transaction keys for fast O(1) duplicate checking
  // Only need to check duplicates for past transactions (future ones won't be in preserved)
  const preservedKeys = new Set();
  preservedRows.forEach(preserved => {
    const preservedDate = toSafeDate(preserved[0]);
    if (preservedDate) {
      const preservedDesc = (preserved[1] || '').toString().trim();
      const preservedMode = (preserved[2] || '').toString().trim();
      const preservedIncome = Math.round((Number(preserved[4]) || 0) * 100); // Round to cents
      const preservedDebits = Math.round((Number(preserved[5]) || 0) * 100);
      // Create a unique key for fast lookup
      const key = `${preservedDate.getTime()}|${preservedDesc}|${preservedMode}|${preservedIncome}|${preservedDebits}`;
      preservedKeys.add(key);
    }
  });

  // Read VariableExpenses (past and future - unlike recurring, variable expenses can be logged late)
  const variableRaw = [];
  if (variableSheet && variableSheet.getLastRow() > 1) {
    const varData = variableSheet.getRange(2, 1, variableSheet.getLastRow() - 1, 7).getValues();
    varData.forEach(row => {
      const rowDate = toSafeDate(row[0]);
      if (rowDate) {
        // Only check duplicates for past transactions (exact key match with preserved)
        let isDuplicate = false;
        if (rowDate < today) {
          const rowDesc = (row[1] || '').toString().trim();
          const rowMode = (row[2] || '').toString().trim();
          const rowIncome = Math.round((Number(row[4]) || 0) * 100);
          const rowDebits = Math.round((Number(row[5]) || 0) * 100);
          const key = `${rowDate.getTime()}|${rowDesc}|${rowMode}|${rowIncome}|${rowDebits}`;
          isDuplicate = preservedKeys.has(key);
        }
        if (!isDuplicate) {
          while (row.length < 7) row.push('');
          row.push(''); // RunningBalance
          row.push(''); // Edited
          variableRaw.push(row);
        }
      }
    });
  }

  // Skip variable rows that match an edited-preserved row (same desc/mode/income/debits),
  // same "closest date within 5 days" rule, to avoid duplicates when user edits a variable-sourced row.
  const variableWithMeta = variableRaw.map((row, i) => {
    const d = toSafeDate(row[0]);
    const desc = (row[1] || '').toString().trim();
    const mode = (row[2] || '').toString().trim();
    const income = Math.round((Number(row[4]) || 0) * 100);
    const debits = Math.round((Number(row[5]) || 0) * 100);
    const sig = `${desc}|${mode}|${income}|${debits}`;
    return { row, date: d, sig, index: i };
  });
  const skipVariableIndices = new Set();
  const msPerDayVar = 24 * 60 * 60 * 1000;
  editedPreserved.forEach(ep => {
    let best = { index: -1, diff: Infinity };
    // Consider all matches; skip the closest by date, not the first found
    variableWithMeta.forEach(m => {
      if (m.sig !== ep.sig || skipVariableIndices.has(m.index)) return;
      const diff = Math.abs(m.date.getTime() - ep.date.getTime());
      if (diff <= EDIT_MATCH_DAYS * msPerDayVar && diff < best.diff) {
        best = { index: m.index, diff };
      }
    });
    if (best.index >= 0) skipVariableIndices.add(best.index);
  });
  const variableTransactions = variableRaw.filter((_, i) => !skipVariableIndices.has(i));

  // Build account savings map from CurrentBalance for SavingsGoals
  const currentBalanceSheet = ss.getSheetByName('CurrentBalance');
  const accountSavingsMap = {};
  if (currentBalanceSheet && currentBalanceSheet.getLastRow() > 1) {
    const cbData = currentBalanceSheet.getRange(2, 1, currentBalanceSheet.getLastRow() - 1, 4).getValues();
    cbData.forEach(row => {
      const accountName = (row[0] || '').toString().trim().toLowerCase();
      const savingsAmount = Number(row[3]) || 0;
      if (accountName) {
        accountSavingsMap[accountName] = savingsAmount;
      }
    });
  }

  // Generate savings goals transactions (account link provides initial progress)
  const savingsGoalsSheet = ss.getSheetByName('SavingsGoals');
  const allSavingsGoalsTransactions = generateSavingsGoalsTransactions(savingsGoalsSheet, accountSavingsMap);
  
  // Filter savings goals to future only
  const futureSavingsGoalsRaw = [];
  allSavingsGoalsTransactions.forEach(row => {
    const rowDate = toSafeDate(row[0]);
    if (rowDate && rowDate >= today) {
      while (row.length < 7) row.push('');
      row.push(''); // RunningBalance
      row.push(''); // Edited
      futureSavingsGoalsRaw.push(row);
    }
  });

  // Skip savings goal instances that match an edited-preserved row (same desc/mode/income/debits),
  // same "closest date within 5 days" rule as recurring; match by desc|mode only (amounts can vary per occurrence).
  const savingsWithMeta = futureSavingsGoalsRaw.map((row, i) => {
    const d = toSafeDate(row[0]);
    const desc = (row[1] || '').toString().trim();
    const mode = (row[2] || '').toString().trim();
    return { row, date: d, desc, mode, index: i };
  });
  const skipSavingsIndices = new Set();
  const msPerDay = 24 * 60 * 60 * 1000;
  editedPreserved.forEach(ep => {
    let best = { index: -1, diff: Infinity };
    // Match by desc|mode only (amounts vary per occurrence). Skip closest by date within 5 days.
    savingsWithMeta.forEach(m => {
      if (m.desc !== ep.desc || m.mode !== ep.mode || skipSavingsIndices.has(m.index)) return;
      const diff = Math.abs(m.date.getTime() - ep.date.getTime());
      if (diff <= EDIT_MATCH_DAYS * msPerDay && diff < best.diff) {
        best = { index: m.index, diff };
      }
    });
    if (best.index >= 0) skipSavingsIndices.add(best.index);
  });
  const futureSavingsGoalsTransactions = futureSavingsGoalsRaw.filter((_, i) => !skipSavingsIndices.has(i));

  // Combine: preserved rows + future recurring + variable (past & future) + future savings goals
  const allTransactions = [...preservedRows, ...futureRecurringTransactions, ...variableTransactions, ...futureSavingsGoalsTransactions];

  if (allTransactions.length === 0 && !existingStartingBalanceRow) {
    SpreadsheetApp.getUi().alert('No transactions found.');
    return;
  }

  // Sort by date, then income-first for same dates, with "Unaccounted" entries always last
  allTransactions.sort((a, b) => {
    const dateA = toSafeDate(a[0]);
    const dateB = toSafeDate(b[0]);
    
    // Handle null dates
    if (!dateA && !dateB) return 0;
    if (!dateA) return 1;
    if (!dateB) return -1;
    
    // Sort by date first
    if (dateA.getTime() !== dateB.getTime()) {
      return dateA - dateB;
    }
    
    // Same date: Check for "Unaccounted" entries (should always be last)
    const descA = (a[1] || '').toString();
    const descB = (b[1] || '').toString();
    const isUnaccountedA = descA === 'Unaccounted Expenses' || descA === 'Unaccounted Income';
    const isUnaccountedB = descB === 'Unaccounted Expenses' || descB === 'Unaccounted Income';
    
    // If one is unaccounted and the other isn't, unaccounted goes last
    if (isUnaccountedA && !isUnaccountedB) return 1;  // A is unaccounted -> A goes after B
    if (isUnaccountedB && !isUnaccountedA) return -1; // B is unaccounted -> B goes after A
    
    // Same date, neither or both are unaccounted: Income entries come first
    const incomeA = Number(a[4]) || 0;  // Income column (index 4)
    const incomeB = Number(b[4]) || 0;
    
    if (incomeA > 0 && incomeB === 0) return -1;  // A has income, B doesn't -> A first
    if (incomeB > 0 && incomeA === 0) return 1;   // B has income, A doesn't -> B first
    
    return 0;
  });

  // Clear existing data (except header row and RunningBalance column)
  // Clear columns 1-7 (Date through CCTransactionDate) and column 9 (Edited)
  // Preserve column 8 (RunningBalance) formulas
  if (finalLastRow >= 2) {
    finalTrackerSheet.getRange(2, 1, finalLastRow - 1, 7).clearContent();  // Columns A-G
    finalTrackerSheet.getRange(2, 9, finalLastRow - 1, 1).clearContent();  // Column I (Edited)
  }

  // Write Starting Balance row if it existed
  let writeRow = 2;
  if (existingStartingBalanceRow) {
    while (existingStartingBalanceRow.length < 9) existingStartingBalanceRow.push('');
    // Write columns 1-7 and 9 separately to preserve RunningBalance formula
    finalTrackerSheet.getRange(2, 1, 1, 7).setValues([existingStartingBalanceRow.slice(0, 7)]);
    finalTrackerSheet.getRange(2, 9, 1, 1).setValues([[existingStartingBalanceRow[8] || '']]);
    writeRow = 3;
  }

  // Write all transactions (columns 1-7 and 9, preserve column 8 RunningBalance)
  if (allTransactions.length > 0) {
    // Extract columns 1-7 (Date through CCTransactionDate)
    const dataColumns1to7 = allTransactions.map(row => row.slice(0, 7));
    finalTrackerSheet.getRange(writeRow, 1, allTransactions.length, 7).setValues(dataColumns1to7);
    
    // Extract column 9 (Edited)
    const dataColumn9 = allTransactions.map(row => [row[8] || '']);
    finalTrackerSheet.getRange(writeRow, 9, allTransactions.length, 1).setValues(dataColumn9);
  }

  // Highlight Salary rows yellow
  highlightRows(finalTrackerSheet, writeRow, allTransactions);

  // Count stats
  const preservedCount = preservedRows.length;
  const newRecurringCount = futureRecurringTransactions.length;
  const newVariableCount = variableTransactions.length;
  const newSavingsGoalsCount = futureSavingsGoalsTransactions.length;

  // Show summary
  const message = 
    `FinalTracker updated!\n\n` +
    `Preserved past/edited rows: ${preservedCount}\n` +
    `New recurring (future): ${newRecurringCount}\n` +
    `New variable (future): ${newVariableCount}\n` +
    `New savings goals (future): ${newSavingsGoalsCount}\n` +
    `Total: ${allTransactions.length} transactions.\n\n` +
    (existingStartingBalanceRow ? `Starting Balance row preserved.\n` : '') +
    `Remember to run "Setup Running Balance" to update formulas.`;

  SpreadsheetApp.getUi().alert(message);
}

/**
 * Get highlight rules from Lookups sheet (columns E-H)
 * @returns {Array} Array of highlight rule objects
 */
function getHighlightRules() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const lookupsSheet = ss.getSheetByName('Lookups');
  
  if (!lookupsSheet) {
    return [];
  }
  
  const lastRow = lookupsSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  
  // Read columns E-H (HighlightColumn, MatchType, MatchValue, HighlightColor)
  const data = lookupsSheet.getRange(2, 5, lastRow - 1, 4).getValues();
  
  const rules = [];
  data.forEach(row => {
    const column = (row[0] || '').toString().trim();
    const matchType = (row[1] || '').toString().trim().toLowerCase();
    const value = (row[2] || '').toString().trim();
    const color = (row[3] || '').toString().trim();
    
    // Only add valid rules (must have all fields)
    if (column && matchType && value && color) {
      rules.push({
        column: column,
        matchType: matchType,
        value: value,
        color: color
      });
    }
  });
  
  return rules;
}

/**
 * Highlight rows in FinalTracker based on configurable rules from Lookups sheet
 * @param {Sheet} sheet - The FinalTracker sheet
 * @param {number} startRow - The row where transactions start
 * @param {Array} transactions - Array of transaction rows
 */
function highlightRows(sheet, startRow, transactions) {
  if (!transactions || transactions.length === 0) return;
  
  const white = '#FFFFFF';
  
  // First, reset all transaction rows to white background (columns A-H, excluding Edited)
  sheet.getRange(startRow, 1, transactions.length, 8).setBackground(white);
  
  // Get highlight rules from Lookups sheet
  const rules = getHighlightRules();
  
  if (rules.length === 0) return;
  
  // Column mapping for FinalTracker: Date=0, Description=1, Mode=2, Category=3
  const columnMap = {
    'description': 1,
    'mode': 2,
    'category': 3
  };
  
  // Build a map of row index -> color (last rule wins if multiple matches)
  const rowColors = new Map();
  
  // Apply each rule and collect which rows need which colors
  rules.forEach(rule => {
    const colIndex = columnMap[rule.column.toLowerCase()];
    if (colIndex === undefined) return;
    
    transactions.forEach((row, idx) => {
      const cellValue = (row[colIndex] || '').toString().toLowerCase();
      const matchValue = rule.value.toLowerCase();
      
      let isMatch = false;
      if (rule.matchType === 'exact') {
        isMatch = cellValue === matchValue;
      } else if (rule.matchType === 'contains') {
        isMatch = cellValue.includes(matchValue);
      }
      
      if (isMatch) {
        rowColors.set(idx, rule.color);
      }
    });
  });
  
  // Batch apply colors: group rows by color and apply in batches
  if (rowColors.size === 0) return;
  
  // Group row indices by color
  const colorGroups = new Map();
  rowColors.forEach((color, idx) => {
    if (!colorGroups.has(color)) {
      colorGroups.set(color, []);
    }
    colorGroups.get(color).push(idx);
  });
  
  // Apply each color group in batches (process in chunks to avoid too many API calls)
  colorGroups.forEach((indices, color) => {
    // Sort indices to process consecutive rows together
    indices.sort((a, b) => a - b);
    
    // Process in batches of consecutive rows
    let batchStart = indices[0];
    let batchEnd = indices[0];
    
    for (let i = 1; i < indices.length; i++) {
      if (indices[i] === batchEnd + 1) {
        // Consecutive, extend batch
        batchEnd = indices[i];
      } else {
        // Gap found, apply current batch
        const batchLength = batchEnd - batchStart + 1;
        sheet.getRange(startRow + batchStart, 1, batchLength, 8).setBackground(color);
        // Start new batch
        batchStart = indices[i];
        batchEnd = indices[i];
      }
    }
    // Apply final batch
    const batchLength = batchEnd - batchStart + 1;
    sheet.getRange(startRow + batchStart, 1, batchLength, 8).setBackground(color);
  });
}
  
  // Forces a stable "date-only" behavior (midday avoids timezone date shifts)
  function toSafeDate(value) {
    if (!value) return null;
    const d = new Date(value);
    if (isNaN(d.getTime())) return null; // Invalid date
    d.setHours(12, 0, 0, 0);
    return d;
  }

  /**
   * Given a credit card transaction date, compute the statement date
   * using the 26th cutoff rule:
   * - If txnDate day < 26  → statement is the 26th of that month
   * - If txnDate day >= 26 → statement is the 26th of the next month
   * Statement period: 26th of previous month to 25th of current month
   * Transactions on the 26th belong to the next statement period
   */
  function computeCcStatementDate(txnDate) {
    const d = toSafeDate(txnDate);
    const year = d.getFullYear();
    const monthIndex = d.getMonth(); // 0..11
    const day = d.getDate();

    if (day < 26) {
      // Same month, 26th
      return new Date(year, monthIndex, 26, 12, 0, 0, 0);
    }

    // Next month, 26th (Date handles year rollover)
    // Transactions on 26th and later belong to next statement period
    return new Date(year, monthIndex + 1, 26, 12, 0, 0, 0);
  }

  /**
   * Given a statement date (typically the 26th), compute the due date
   * as the 15th of the month AFTER the statement month.
   */
  function computeCcDueDate(statementDate) {
    const s = toSafeDate(statementDate);
    const year = s.getFullYear();
    const monthIndex = s.getMonth(); // 0..11

    // Month after the statement month, day = 15
    return new Date(year, monthIndex + 1, 15, 12, 0, 0, 0);
  }

  /**
   * Get balance breakdown from CurrentBalance sheet
   * Returns: { totalAmount, totalMaintainBalance, totalSavings, availableBalance, accounts[] }
   * Available Balance = Amount - MaintainBalance - SavingsAmount
   */
  function getBalanceBreakdown() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const currentBalanceSheet = ss.getSheetByName('CurrentBalance');
    
    const result = {
      totalAmount: 0,
      totalMaintainBalance: 0,
      totalSavings: 0,
      availableBalance: 0,
      accounts: []
    };
    
    if (!currentBalanceSheet) {
      return result;
    }
    
    const lastRow = currentBalanceSheet.getLastRow();
    if (lastRow < 2) {
      return result;
    }
    
    // Read columns A-D: Account, Amount, MaintainBalance, SavingsAmount
    const data = currentBalanceSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    
    data.forEach(row => {
      const [account, amount, maintainBalance, savingsAmount] = row;
      const amt = Number(amount) || 0;
      const maintain = Number(maintainBalance) || 0;
      const savings = Number(savingsAmount) || 0;
      const available = amt - maintain - savings;
      
      result.totalAmount += amt;
      result.totalMaintainBalance += maintain;
      result.totalSavings += savings;
      result.accounts.push({
        account: account || 'Unknown',
        amount: amt,
        maintainBalance: maintain,
        savingsAmount: savings,
        available: available
      });
    });
    
    result.availableBalance = result.totalAmount - result.totalMaintainBalance - result.totalSavings;
    return result;
  }

  /**
   * Get available balance (quick helper for functions that just need the number)
   * Available = Total Amount - MaintainBalance - SavingsAmount
   */
  function getAvailableBalance() {
    return getBalanceBreakdown().availableBalance;
  }

  /**
   * Display balance summary showing Total vs Available breakdown
   */
  function viewBalanceSummary() {
    const breakdown = getBalanceBreakdown();
    
    if (breakdown.accounts.length === 0) {
      SpreadsheetApp.getUi().alert('CurrentBalance sheet not found or empty.');
      return;
    }
    
    let message = `Balance Summary\n`;
    message += `${'='.repeat(40)}\n\n`;
    
    // Per-account breakdown
    message += `By Account:\n`;
    breakdown.accounts.forEach(acc => {
      message += `\n${acc.account}:\n`;
      message += `  Amount: ${acc.amount.toFixed(2)}\n`;
      if (acc.maintainBalance > 0) {
        message += `  - MaintainBalance: ${acc.maintainBalance.toFixed(2)}\n`;
      }
      if (acc.savingsAmount > 0) {
        message += `  - Savings: ${acc.savingsAmount.toFixed(2)}\n`;
      }
      message += `  = Available: ${acc.available.toFixed(2)}\n`;
    });
    
    // Totals
    message += `\n${'='.repeat(40)}\n`;
    message += `TOTALS:\n`;
    message += `  Total Amount: ${breakdown.totalAmount.toFixed(2)}\n`;
    message += `  - MaintainBalance: ${breakdown.totalMaintainBalance.toFixed(2)}\n`;
    message += `  - Savings: ${breakdown.totalSavings.toFixed(2)}\n`;
    message += `  = Available Balance: ${breakdown.availableBalance.toFixed(2)}\n`;
    
    SpreadsheetApp.getUi().alert(message);
  }
 
/**
 * Helper function to normalize CC dates and check for incomplete entries
 * Called internally by updateFinalTracker
 * Returns object with normalized count and incomplete entries info
 */
function normalizeVariableExpensesCcDatesWithCheck(ss, variableSheet) {
  if (!variableSheet) return { normalized: 0, incomplete: [] };
  
  const lastRow = variableSheet.getLastRow();
  if (lastRow <= 1) return { normalized: 0, incomplete: [] };
  
  const data = variableSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  // Quick check: skip if no CreditCard transactions at all
  const hasCreditCard = data.some(row => row[2] === 'CreditCard');
  if (!hasCreditCard) return { normalized: 0, incomplete: [] };
  
  let updatedCount = 0;
  const incompleteEntries = [];
  
  data.forEach((row, idx) => {
    const [existingDate, description, mode, , , , ccTxnDate] = row;
    
    if (mode !== 'CreditCard') return;
    
    // Check for incomplete CC entries (missing CCTransactionDate)
    if (!ccTxnDate) {
      incompleteEntries.push({
        row: idx + 2, // 1-based row number
        description: description || '(no description)'
      });
      return;
    }
    
    const txnDate = toSafeDate(ccTxnDate);
    const statementDate = computeCcStatementDate(txnDate);
    const dueDate = computeCcDueDate(statementDate);
    
    const currentDate = existingDate ? toSafeDate(existingDate) : null;
    if (!currentDate || currentDate.getTime() !== dueDate.getTime()) {
      row[0] = dueDate;
      updatedCount += 1;
    }
  });
  
  // Warn about incomplete CC entries
  if (incompleteEntries.length > 0) {
    const details = incompleteEntries.slice(0, 5).map(e => `  Row ${e.row}: ${e.description}`).join('\n');
    const moreText = incompleteEntries.length > 5 ? `\n  ...and ${incompleteEntries.length - 5} more` : '';
    SpreadsheetApp.getUi().alert(
      `Warning: ${incompleteEntries.length} CreditCard row(s) are missing CCTransactionDate:\n\n` +
      details + moreText + '\n\n' +
      `Please fill in the CCTransactionDate column for accurate due date calculation.`
    );
  }
  
  if (updatedCount > 0) {
    const dateColumn = data.map(r => [r[0]]);
    variableSheet.getRange(2, 1, dateColumn.length, 1).setValues(dateColumn);
  }
  
  return { normalized: updatedCount, incomplete: incompleteEntries };
}

// Normalize VariableExpenses credit card rows so Date reflects the due date
// If silent=true, skips alerts and just returns the count
function normalizeVariableExpensesCcDates(silent = false) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const variableSheet =
    ss.getSheetByName('VariableExpenses') || ss.getSheetByName('VariableExpences');

  if (!variableSheet) {
    if (!silent) SpreadsheetApp.getUi().alert('Variable expenses sheet not found.');
    return 0;
  }

  const lastRow = variableSheet.getLastRow();
  if (lastRow <= 1) {
    if (!silent) SpreadsheetApp.getUi().alert('No variable expenses to normalize.');
    return 0;
  }

  // Expect columns: Date, Description, Mode, Category, Income, Debits, CCTransactionDate
  const data = variableSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  let updatedCount = 0;
  let incompleteCount = 0;

  data.forEach((row, idx) => {
    const [existingDate, description, mode, , , , ccTxnDate] = row;

    // Skip non-CreditCard rows
    if (mode !== 'CreditCard') return;
    
    // Check for incomplete CC entries (missing CCTransactionDate)
    if (!ccTxnDate) {
      incompleteCount += 1;
      return;
    }

    const txnDate = toSafeDate(ccTxnDate);
    const statementDate = computeCcStatementDate(txnDate);
    const dueDate = computeCcDueDate(statementDate);

    const currentDate = existingDate ? toSafeDate(existingDate) : null;
    if (!currentDate || currentDate.getTime() !== dueDate.getTime()) {
      row[0] = dueDate; // Update Date column to due date
      updatedCount += 1;
    }
  });

  // Warn about incomplete CC entries
  if (incompleteCount > 0 && !silent) {
    SpreadsheetApp.getUi().alert(
      `Warning: ${incompleteCount} CreditCard row(s) are missing CCTransactionDate.\n` +
      `Please fill in the CCTransactionDate column for these entries.`
    );
  }

  if (updatedCount === 0) {
    if (!silent) SpreadsheetApp.getUi().alert('All CreditCard rows are already normalized.');
    return 0;
  }

  // Write back only the Date column (col 1) for all rows
  const dateColumn = data.map(r => [r[0]]);
  variableSheet.getRange(2, 1, dateColumn.length, 1).setValues(dateColumn);

  if (!silent) {
    SpreadsheetApp.getUi().alert(`Normalized ${updatedCount} CreditCard variable expenses.`);
  }
  
  return updatedCount;
}
  
  function generateDates(frequency, day, month, startDate, endDate) {
    const dates = [];
  
    switch (frequency) {
      case 'WEEKLY': {
        // day: 1=Mon ... 7=Sun (coerce to number)
        const numericDay = Number(day);
        if (isNaN(numericDay) || numericDay < 1 || numericDay > 7) break;
        const target = (numericDay === 7) ? 0 : numericDay; // JS: 0=Sun..6=Sat
        let d = new Date(startDate);
  
        // move forward to the first target weekday on/after startDate
        const diff = (target - d.getDay() + 7) % 7;
        d.setDate(d.getDate() + diff);
  
        while (d <= endDate) {
          dates.push(new Date(d));
          d.setDate(d.getDate() + 7);
        }
        break;
      }
  
      case 'N_DAY_IN_MONTH': {
        // day: single value (e.g., "15") or comma-separated (e.g., "15, 28")
        const dayStr = (day || '').toString();
        let daysArray = [];
        
        if (dayStr.includes(',')) {
          // Multiple days: "15, 28"
          daysArray = dayStr.split(',').map(d => Number(d.trim())).filter(d => !isNaN(d) && d >= 1 && d <= 31).sort((a, b) => a - b);
        } else {
          // Single day
          const numericDay = Number(day);
          if (!isNaN(numericDay) && numericDay >= 1 && numericDay <= 31) {
            daysArray = [numericDay];
          }
        }
        
        if (daysArray.length === 0) break;
        
        let y = startDate.getFullYear();
        let m = startDate.getMonth();

        while (true) {
          const last = lastDayOfMonth(y, m);
          
          // Generate dates for each day in the array
          for (const numericDay of daysArray) {
            const dd = Math.min(numericDay, last);
            const d = new Date(y, m, dd, 12, 0, 0, 0);
    
            if (d >= startDate && d <= endDate) dates.push(d);
          }
          
          if (new Date(y, m + 1, 1) > endDate) break;

          m += 1;
          if (m > 11) { m = 0; y += 1; }
        }
        break;
      }
  
      case 'ANNUAL': {
        // month: 1..12, day: 1..31 (clamp) - coerce to numbers
        const numericMonth = Number(month);
        const numericDay = Number(day);
        if (isNaN(numericMonth) || numericMonth < 1 || numericMonth > 12) break;
        if (isNaN(numericDay) || numericDay < 1 || numericDay > 31) break;
        let y = startDate.getFullYear();

        while (true) {
          const m = numericMonth - 1;
          const last = lastDayOfMonth(y, m);
          const dd = Math.min(numericDay, last);
          const d = new Date(y, m, dd, 12, 0, 0, 0);
  
          if (d >= startDate && d <= endDate) dates.push(d);
          if (new Date(y + 1, 0, 1) > endDate) break;
  
          y += 1;
        }
        break;
      }

      case 'EVERY_N_DAYS': {
        // day: number of days between each occurrence (e.g., 14 = every 2 weeks)
        const intervalDays = Number(day);
        if (isNaN(intervalDays) || intervalDays < 1) break;
        
        let d = new Date(startDate);
        d.setHours(12, 0, 0, 0);

        while (d <= endDate) {
          dates.push(new Date(d));
          d.setDate(d.getDate() + intervalDays);
        }
        break;
      }
    }
  
    return dates;
  }
  
  function lastDayOfMonth(year, monthIndex) {
    // monthIndex: 0..11
    return new Date(year, monthIndex + 1, 0).getDate();
  }
  
  /**
   * Sets up the RunningBalance column in FinalTracker sheet
   * Adds/updates the "Starting Balance" row as the first data row
   * Uses Available Balance (Amount - MaintainBalance - SavingsAmount)
   * Formula for subsequent rows: Previous row's RunningBalance + current Income - current Debits
   */
  function setupRunningBalance() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const finalTrackerSheet = ss.getSheetByName('FinalTracker');
    
    if (!finalTrackerSheet) {
      SpreadsheetApp.getUi().alert('FinalTracker sheet not found.');
      return;
    }
    
    const RUNNING_BALANCE_COL = 8; // Column H
    const lastRow = finalTrackerSheet.getLastRow();
    
    // Ensure headers are set (9 columns including Edited)
    const headers = ['Date', 'Description', 'Mode', 'Category', 'Income', 'Debits', 'CCTransactionDate', 'RunningBalance', 'Edited'];
    finalTrackerSheet.getRange(1, 1, 1, 9).setValues([headers]);
    finalTrackerSheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    
    // Get balance breakdown (available = amount - maintainBalance - savings)
    const balanceBreakdown = getBalanceBreakdown();
    if (balanceBreakdown.accounts.length === 0) {
      SpreadsheetApp.getUi().alert('CurrentBalance sheet not found or empty. Cannot set up RunningBalance formulas.');
      return;
    }
    
    // Check if row 2 exists and is the "Starting Balance" row
    const today = toSafeDate(new Date());
    let needsStartingBalanceRow = false;
    
    if (lastRow <= 1) {
      // No data rows, need to insert starting balance row
      needsStartingBalanceRow = true;
    } else {
      // Check if row 2 is the starting balance row
      const row2Description = finalTrackerSheet.getRange(2, 2).getValue(); // Description column
      if (row2Description !== 'Starting Balance') {
        // Insert a new row at position 2
        finalTrackerSheet.insertRowAfter(1);
        needsStartingBalanceRow = true;
      // Starting Balance row already exists - don't update the date
      // Use updateStartingBalance() to update the date manually
    }
    }
    
    // Set up the Starting Balance row (row 2)
    if (needsStartingBalanceRow) {
      finalTrackerSheet.getRange(2, 1, 1, 7).setValues([[
        today,              // Date: today
        'Starting Balance', // Description
        '',                 // Mode
        '',                 // Category
        0,                  // Income: 0
        0,                  // Debits: 0
        ''                  // CCTransactionDate
      ]]);
    }
    
    // Set Available Balance as static value for row 2
    const availableBalance = balanceBreakdown.availableBalance;
    finalTrackerSheet.getRange(2, RUNNING_BALANCE_COL).setValue(availableBalance);
    
    // Get the current last row (may have changed if we inserted a row)
    const currentLastRow = finalTrackerSheet.getLastRow();
    
    // If there are more rows (row 3 and beyond), set up their formulas
    if (currentLastRow > 2) {
      const formulas = [];
      for (let i = 3; i <= currentLastRow; i++) {
        // Formula: Previous row's RunningBalance + current Income - current Debits
        // =H{i-1}+E{i}-F{i}
        const formula = `=H${i-1}+E${i}-F${i}`;
        formulas.push([formula]);
      }
      
      // Write formulas for rows 3 onwards
      finalTrackerSheet.getRange(3, RUNNING_BALANCE_COL, formulas.length, 1).setFormulas(formulas);
    }
    
    SpreadsheetApp.getUi().alert(
      `RunningBalance column set up!\n\n` +
      `Starting Balance: ${availableBalance.toFixed(2)} (Available)\n` +
      `  Total Amount: ${balanceBreakdown.totalAmount.toFixed(2)}\n` +
      `  - MaintainBalance: ${balanceBreakdown.totalMaintainBalance.toFixed(2)}\n` +
      `  - Savings: ${balanceBreakdown.totalSavings.toFixed(2)}\n\n` +
      `RunningBalance formulas applied to ${currentLastRow - 1} row(s).\n\n` +
      `Note: Starting Balance is static. Use "Update Starting Balance" to refresh.`
    );
  }

  /**
   * Updates the Starting Balance row's date to today and recalculates the balance
   * Call this when you want to "reset" your starting point to today
   */
  function updateStartingBalance() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const finalTrackerSheet = ss.getSheetByName('FinalTracker');

    if (!finalTrackerSheet) {
      SpreadsheetApp.getUi().alert('FinalTracker sheet not found.');
      return;
    }

    // Get balance breakdown
    const balanceBreakdown = getBalanceBreakdown();
    if (balanceBreakdown.accounts.length === 0) {
      SpreadsheetApp.getUi().alert('CurrentBalance sheet not found or empty.');
      return;
    }

    // Check if Starting Balance row exists
    const lastRow = finalTrackerSheet.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('No Starting Balance row found. Run "Setup Running Balance" first.');
      return;
    }

    const row2Description = finalTrackerSheet.getRange(2, 2).getValue();
    if (row2Description !== 'Starting Balance') {
      SpreadsheetApp.getUi().alert('Starting Balance row not found at row 2. Run "Setup Running Balance" first.');
      return;
    }

    // Update date to today
    const today = toSafeDate(new Date());
    finalTrackerSheet.getRange(2, 1).setValue(today);

    // Update the RunningBalance value with Available Balance
    const RUNNING_BALANCE_COL = 8;
    const availableBalance = balanceBreakdown.availableBalance;
    finalTrackerSheet.getRange(2, RUNNING_BALANCE_COL).setValue(availableBalance);

    const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    SpreadsheetApp.getUi().alert(
      `Starting Balance updated!\n\n` +
      `Date: ${formattedDate}\n` +
      `Available Balance: ${availableBalance.toFixed(2)}\n\n` +
      `Breakdown:\n` +
      `  Total Amount: ${balanceBreakdown.totalAmount.toFixed(2)}\n` +
      `  - MaintainBalance: ${balanceBreakdown.totalMaintainBalance.toFixed(2)}\n` +
      `  - Savings: ${balanceBreakdown.totalSavings.toFixed(2)}`
    );
  }

  /**
   * Refresh the Starting Balance VALUE only (keeps the existing date)
   * Use this when you update MaintainBalance or SavingsAmount in CurrentBalance
   * Does NOT reset the tracker - just recalculates the available balance
   */
  function refreshStartingBalanceValue() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const finalTrackerSheet = ss.getSheetByName('FinalTracker');

    if (!finalTrackerSheet) {
      SpreadsheetApp.getUi().alert('FinalTracker sheet not found.');
      return;
    }

    // Get balance breakdown
    const balanceBreakdown = getBalanceBreakdown();
    if (balanceBreakdown.accounts.length === 0) {
      SpreadsheetApp.getUi().alert('CurrentBalance sheet not found or empty.');
      return;
    }

    // Check if Starting Balance row exists
    const lastRow = finalTrackerSheet.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('No Starting Balance row found. Run "Setup Running Balance" first.');
      return;
    }

    const row2Description = finalTrackerSheet.getRange(2, 2).getValue();
    if (row2Description !== 'Starting Balance') {
      SpreadsheetApp.getUi().alert('Starting Balance row not found at row 2. Run "Setup Running Balance" first.');
      return;
    }

    // Get the existing date (DON'T change it)
    const existingDate = finalTrackerSheet.getRange(2, 1).getValue();
    const oldValue = finalTrackerSheet.getRange(2, 8).getValue(); // Current RunningBalance value

    // Update ONLY the RunningBalance value with new Available Balance
    const RUNNING_BALANCE_COL = 8;
    const availableBalance = balanceBreakdown.availableBalance;
    finalTrackerSheet.getRange(2, RUNNING_BALANCE_COL).setValue(availableBalance);

    const formattedDate = existingDate ? Utilities.formatDate(toSafeDate(existingDate), Session.getScriptTimeZone(), 'MM/dd/yyyy') : 'N/A';
    const difference = availableBalance - (Number(oldValue) || 0);
    
    SpreadsheetApp.getUi().alert(
      `Starting Balance VALUE refreshed!\n\n` +
      `Date: ${formattedDate} (unchanged)\n` +
      `Old Value: ${(Number(oldValue) || 0).toFixed(2)}\n` +
      `New Value: ${availableBalance.toFixed(2)}\n` +
      `Change: ${difference >= 0 ? '+' : ''}${difference.toFixed(2)}\n\n` +
      `Breakdown:\n` +
      `  Total Amount: ${balanceBreakdown.totalAmount.toFixed(2)}\n` +
      `  - MaintainBalance: ${balanceBreakdown.totalMaintainBalance.toFixed(2)}\n` +
      `  - Savings: ${balanceBreakdown.totalSavings.toFixed(2)}`
    );
  }

  /**
   * Reconciles Available Balance with today's running balance in FinalTracker
   * If there's a difference, adds an entry to VariableExpenses as "Unaccounted Expenses" or "Unaccounted Income"
   * This ensures Available Balance and today's running balance match
   */
  function reconcileBalance() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const finalTrackerSheet = ss.getSheetByName('FinalTracker');
    const variableSheet = ss.getSheetByName('VariableExpenses') || ss.getSheetByName('VariableExpences');
    
    if (!finalTrackerSheet) {
      SpreadsheetApp.getUi().alert('FinalTracker sheet not found.');
      return;
    }
    
    if (!variableSheet) {
      SpreadsheetApp.getUi().alert('VariableExpenses sheet not found.');
      return;
    }
    
    // Get available balance (amount - maintainBalance - savings)
    const balanceBreakdown = getBalanceBreakdown();
    if (balanceBreakdown.accounts.length === 0) {
      SpreadsheetApp.getUi().alert('CurrentBalance sheet not found or empty.');
      return;
    }
    const availableBalance = balanceBreakdown.availableBalance;
    
    // Find today's running balance in FinalTracker
    const today = toSafeDate(new Date());
    const lastRow = finalTrackerSheet.getLastRow();
    
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('No data in FinalTracker. Please run "Update FinalTracker" first.');
      return;
    }
    
    // Get all data rows (columns: Date, Description, Mode, Category, Income, Debits, CCTransactionDate, RunningBalance)
    const data = finalTrackerSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    // Build array of valid transactions (on or before today) with their dates and running balances
    // This filters out future transactions and rows without valid dates
    const validTransactions = [];
    for (let i = 0; i < data.length; i++) {
      const rowDate = toSafeDate(data[i][0]);
      if (rowDate && rowDate <= today) {
        const runningBalance = Number(data[i][7]) || 0;
        validTransactions.push({
          date: rowDate,
          runningBalance: runningBalance,
          index: i,
          row: data[i]
        });
      }
    }
    
    if (validTransactions.length === 0) {
      SpreadsheetApp.getUi().alert('Could not find today\'s running balance in FinalTracker. Please ensure there are transactions on or before today.');
      return;
    }
    
    // Sort by date (ascending), then by index (to preserve order for same dates)
    // This ensures we can find the LAST transaction on today's date
    validTransactions.sort((a, b) => {
      const dateDiff = a.date.getTime() - b.date.getTime();
      if (dateDiff !== 0) return dateDiff;
      return a.index - b.index; // If same date, preserve original order
    });
    
    // Find the LAST transaction on today's date (if exists), otherwise the last transaction before today
    // Since data is sorted ascending, the last transaction on today will be at the end of today's group
    let todayRunningBalance = null;
    let todayRowIndex = -1;
    let foundDate = null;
    
    // First, try to find the last transaction on today's date
    for (let i = validTransactions.length - 1; i >= 0; i--) {
      const txn = validTransactions[i];
      if (txn.date.getTime() === today.getTime()) {
        // Found today's date - take the last one (since we're iterating backwards)
        todayRunningBalance = txn.runningBalance;
        todayRowIndex = txn.index;
        foundDate = txn.date;
        break; // This is the last transaction on today
      }
    }
    
    // If no transaction found on today, use the most recent transaction before today
    if (todayRunningBalance === null) {
      // Since data is sorted, the last item is the most recent before today
      const lastTxn = validTransactions[validTransactions.length - 1];
      todayRunningBalance = lastTxn.runningBalance;
      todayRowIndex = lastTxn.index;
      foundDate = lastTxn.date;
    }
    
    // Calculate difference (available balance vs running balance)
    const difference = availableBalance - todayRunningBalance;
    
    // Check if an "Unaccounted" entry already exists for today
    const varLastRow = variableSheet.getLastRow();
    let existingEntryIndex = -1;
    
    if (varLastRow > 1) {
      const varData = variableSheet.getRange(2, 1, varLastRow - 1, 7).getValues();
      for (let i = 0; i < varData.length; i++) {
        const varDate = toSafeDate(varData[i][0]);
        const varDesc = varData[i][1];
        if (varDate && varDate.getTime() === today.getTime() && 
            (varDesc === 'Unaccounted Expenses' || varDesc === 'Unaccounted Income')) {
          existingEntryIndex = i;
          break;
        }
      }
    }
    
    // If difference is negligible (less than 1 cent)
    if (Math.abs(difference) < 0.01) {
      // If there's an existing unaccounted entry, remove it
      if (existingEntryIndex >= 0) {
        variableSheet.deleteRow(existingEntryIndex + 2);
        SpreadsheetApp.getUi().alert(
          `Balances match!\n\n` +
          `Available Balance: ${availableBalance.toFixed(2)}\n` +
          `Today's Running Balance: ${todayRunningBalance.toFixed(2)}\n` +
          `Difference: ${difference.toFixed(2)}\n\n` +
          `Removed existing unaccounted entry.`
        );
      } else {
        SpreadsheetApp.getUi().alert(
          `Balances match!\n\n` +
          `Available Balance: ${availableBalance.toFixed(2)}\n` +
          `Today's Running Balance: ${todayRunningBalance.toFixed(2)}\n` +
          `Difference: ${difference.toFixed(2)}`
        );
      }
      return;
    }
    
    // If there's an existing entry, update it
    if (existingEntryIndex >= 0) {
      const response = SpreadsheetApp.getUi().alert(
        `Unaccounted entry already exists for today.\n\n` +
        `Available Balance: ${availableBalance.toFixed(2)}\n` +
        `Today's Running Balance: ${todayRunningBalance.toFixed(2)}\n` +
        `Difference: ${difference.toFixed(2)}\n\n` +
        `Would you like to update it?`,
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      
      if (response !== SpreadsheetApp.getUi().Button.YES) {
        return;
      }
      
      // Update the existing entry
      if (difference < 0) {
        // CurrentBalance is less, so it's an expense
        variableSheet.getRange(existingEntryIndex + 2, 1, 1, 7).setValues([[
          today,
          'Unaccounted Expenses',
          'Cash',
          'Adjustment',
          '',
          Math.abs(difference),
          ''
        ]]);
      } else {
        // CurrentBalance is more, so it's income
        variableSheet.getRange(existingEntryIndex + 2, 1, 1, 7).setValues([[
          today,
          'Unaccounted Income',
          'Cash',
          'Adjustment',
          difference,
          '',
          ''
        ]]);
      }
      SpreadsheetApp.getUi().alert(`Updated unaccounted entry for today.`);
      return;
    }
    
    // Add new unaccounted entry
    const newRow = [];
    if (difference < 0) {
      // CurrentBalance is less than running balance, so it's an expense
      newRow.push([
        today,
        'Unaccounted Expenses',
        'Cash',
        'Adjustment',
        '',
        Math.abs(difference),
        ''
      ]);
    } else {
      // CurrentBalance is more than running balance, so it's income
      newRow.push([
        today,
        'Unaccounted Income',
        'Cash',
        'Adjustment',
        difference,
        '',
        ''
      ]);
    }
    
    // Append to VariableExpenses
    const insertRow = varLastRow > 1 ? varLastRow + 1 : 2;
    variableSheet.getRange(insertRow, 1, 1, 7).setValues(newRow);
    
    const entryType = difference < 0 ? 'Unaccounted Expenses' : 'Unaccounted Income';
    const amount = Math.abs(difference);
    
    SpreadsheetApp.getUi().alert(
      `Reconciliation complete!\n\n` +
      `Available Balance: ${availableBalance.toFixed(2)}\n` +
      `Today's Running Balance: ${todayRunningBalance.toFixed(2)}\n` +
      `Difference: ${difference.toFixed(2)}\n\n` +
      `Added "${entryType}" entry of ${amount.toFixed(2)} to VariableExpenses.\n\n` +
      `Please run "Update FinalTracker" to include this in the tracker.`
    );
  }
  
  /**
   * Calculate daily budget over the next 4 salary receipts
   * Uses Available Balance (Amount - MaintainBalance - SavingsAmount)
   * Formula: (projectedBalanceAt4thSalary - availableBalanceToday) / daysUntil4thSalary
   */
  function recalcDailyBudget() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const finalTrackerSheet = ss.getSheetByName('FinalTracker');
    
    if (!finalTrackerSheet) {
      SpreadsheetApp.getUi().alert('FinalTracker sheet not found.');
      return;
    }
    
    // Get available balance (amount - maintainBalance - savings)
    const balanceBreakdown = getBalanceBreakdown();
    if (balanceBreakdown.accounts.length === 0) {
      SpreadsheetApp.getUi().alert('CurrentBalance sheet not found or empty.');
      return;
    }
    const availableBalanceToday = balanceBreakdown.availableBalance;
    
    // Find salary rows in FinalTracker (Category == "Salary" and Income > 0)
    const lastRow = finalTrackerSheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('No data in FinalTracker. Please generate schedule first.');
      return;
    }
    
    // Columns: Date, Description, Mode, Category, Income, Debits, CCTransactionDate, RunningBalance
    const data = finalTrackerSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const today = toSafeDate(new Date());
    
    // Find all future salary dates (where Category == "Salary" and Date >= today)
    const salaryRows = [];
    data.forEach((row, idx) => {
      const [date, , , category, income] = row;
      if (category === 'Salary' && income && Number(income) > 0) {
        const rowDate = toSafeDate(date);
        if (rowDate && rowDate >= today) {
          salaryRows.push({
            date: rowDate,
            runningBalance: row[7] || null, // RunningBalance column (index 7)
            rowIndex: idx + 2 // Actual row number in sheet
          });
        }
      }
    });
    
    if (salaryRows.length === 0) {
      SpreadsheetApp.getUi().alert('No future salary dates found in FinalTracker.');
      return;
    }
    
    // Sort by date and get the 4th one
    salaryRows.sort((a, b) => a.date - b.date);
    const fourthSalary = salaryRows.length >= 4 ? salaryRows[3] : salaryRows[salaryRows.length - 1];
    
    if (fourthSalary.runningBalance === null || fourthSalary.runningBalance === undefined || fourthSalary.runningBalance === '') {
      SpreadsheetApp.getUi().alert('RunningBalance not found for 4th salary. Please run "Setup Running Balance" first.');
      return;
    }
    
    const projectedBalanceAt4thSalary = Number(fourthSalary.runningBalance) || 0;
    const fourthSalaryDate = fourthSalary.date;
    
    // Calculate days until 4th salary
    const daysUntil = Math.max(1, Math.floor((fourthSalaryDate - today) / (1000 * 60 * 60 * 24)));
    
    // Calculate daily budget
    const dailyBudget = (projectedBalanceAt4thSalary - availableBalanceToday) / daysUntil;
    
    // Write to Variables sheet (or create it)
    let varsSheet = ss.getSheetByName('Variables');
    if (!varsSheet) {
      varsSheet = ss.insertSheet('Variables');
      varsSheet.getRange(1, 1, 1, 3).setValues([['Metric', 'Value', 'Notes']]);
      varsSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    }
    
    // Clear existing budget data and write new values
    const budgetData = [
      ['Today', today, ''],
      ['Total Amount', balanceBreakdown.totalAmount, 'Sum of all accounts'],
      ['MaintainBalance', balanceBreakdown.totalMaintainBalance, 'Bank required minimum'],
      ['Savings', balanceBreakdown.totalSavings, 'Set aside for savings'],
      ['Available Balance', availableBalanceToday, 'Amount - Maintain - Savings'],
      ['', '', ''],
      ['4th Salary Date', fourthSalaryDate, ''],
      ['Projected Balance at 4th Salary', projectedBalanceAt4thSalary, ''],
      ['Days Until 4th Salary', daysUntil, ''],
      ['Daily Budget', dailyBudget, ''],
      ['', '', ''],
      ['Last Updated', new Date(), '']
    ];
    
    varsSheet.getRange(2, 1, budgetData.length, 3).setValues(budgetData);
    varsSheet.autoResizeColumns(1, 3);
    
    SpreadsheetApp.getUi().alert(
      `Daily Budget Calculated!\n\n` +
      `Available Balance: ${availableBalanceToday.toFixed(2)}\n` +
      `  (Total: ${balanceBreakdown.totalAmount.toFixed(2)} - Maintain: ${balanceBreakdown.totalMaintainBalance.toFixed(2)} - Savings: ${balanceBreakdown.totalSavings.toFixed(2)})\n\n` +
      `4th Salary Date: ${Utilities.formatDate(fourthSalaryDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')}\n` +
      `Projected Balance: ${projectedBalanceAt4thSalary.toFixed(2)}\n` +
      `Days Until: ${daysUntil}\n` +
      `Daily Budget: ${dailyBudget.toFixed(2)}\n\n` +
      `Results written to Variables sheet.`
    );
  }

  /**
   * Alias for updateFinalTracker() - kept for backward compatibility
   */
  function generateRecurring() {
    updateFinalTracker();
  }

  /**
   * Parse CSV text into array of transaction objects
   * CSV format: Date, Description, Amount, Category, Mode
   */
  function parseCsvTransactions(csvText) {
    if (!csvText || csvText.trim() === '') {
      return { transactions: [], errors: ['CSV input is empty'] };
    }

    // Normalize line endings (handle \r\n, \r, and \n)
    const normalizedText = csvText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    
    // Split by newline and filter out empty lines
    const lines = normalizedText.split('\n')
      .map(line => line.trim())
      .filter(line => line.length > 0);
    
    const transactions = [];
    const errors = [];

    lines.forEach((line, lineNum) => {
      try {
        // Simple CSV parsing (handles quoted fields)
        const fields = [];
        let currentField = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
          const char = line[i];
          
          if (char === '"') {
            inQuotes = !inQuotes;
          } else if (char === ',' && !inQuotes) {
            fields.push(currentField.trim());
            currentField = '';
          } else {
            currentField += char;
          }
        }
        fields.push(currentField.trim()); // Add last field

        if (fields.length !== 7) {
          errors.push(`Line ${lineNum + 1}: Expected 7 fields, got ${fields.length}`);
          return;
        }

        const [dateStr, description, mode, category, incomeStr, debitsStr, ccTxnDateStr] = fields;

        transactions.push({
          date: dateStr,
          description: description,
          mode: mode,
          category: category,
          income: incomeStr,
          debits: debitsStr,
          ccTransactionDate: ccTxnDateStr,
          lineNumber: lineNum + 1
        });
      } catch (e) {
        errors.push(`Line ${lineNum + 1}: Parse error - ${e.message}`);
      }
    });

    return { transactions, errors };
  }

  /**
   * Validate a transaction object
   * Returns { valid: boolean, errors: string[] }
   */
  function validateTransaction(txn) {
    const errors = [];

    // Validate date
    const date = toSafeDate(txn.date);
    if (!date) {
      errors.push(`Invalid date: ${txn.date}`);
    }

    // Validate description
    if (!txn.description || txn.description.trim() === '') {
      errors.push('Description is required');
    }

    // Validate mode
    if (!txn.mode || txn.mode.trim() === '') {
      errors.push('Mode is required');
    }

    // Validate category
    if (!txn.category || txn.category.trim() === '') {
      errors.push('Category is required');
    }

    // Validate income and debits - at least one must be provided and numeric
    const income = txn.income ? Number(txn.income) : 0;
    const debits = txn.debits ? Number(txn.debits) : 0;
    
    if (isNaN(income) || income < 0) {
      errors.push(`Invalid income: ${txn.income}`);
    }
    if (isNaN(debits) || debits < 0) {
      errors.push(`Invalid debits: ${txn.debits}`);
    }
    if (income === 0 && debits === 0) {
      errors.push('Either Income or Debits must be greater than 0');
    }
    if (income > 0 && debits > 0) {
      errors.push('Cannot have both Income and Debits (one must be 0 or empty)');
    }

    // Validate CCTransactionDate if mode is CreditCard (handle both "CreditCard" and "Credit Card")
    let ccTxnDate = null;
    const normalizedMode = (txn.mode || '').trim().toLowerCase().replace(/\s+/g, '');
    if (normalizedMode === 'creditcard') {
      if (txn.ccTransactionDate && txn.ccTransactionDate.trim() !== '') {
        ccTxnDate = toSafeDate(txn.ccTransactionDate);
        if (!ccTxnDate) {
          errors.push(`Invalid CCTransactionDate: ${txn.ccTransactionDate}`);
        }
      }
    }

    return {
      valid: errors.length === 0,
      errors: errors,
      validatedData: errors.length === 0 ? {
        date: date,
        description: txn.description.trim(),
        mode: txn.mode.trim(),
        category: txn.category.trim(),
        income: income,
        debits: debits,
        ccTransactionDate: ccTxnDate
      } : null
    };
  }

  /**
   * Check if a transaction is a duplicate in recurring transactions
   * recurringData: Array of recurring transaction rows (generated from RecurExpenses/RecurIncome)
   * Returns { isDuplicate: boolean, matchInfo: string }
   */
  function checkDuplicateInRecurring(transaction, recurringData) {
    if (!recurringData || recurringData.length === 0) {
      return { isDuplicate: false, matchInfo: null };
    }

    const txnAmount = transaction.income || transaction.debits || 0;
    const txnDesc = transaction.description.toLowerCase().trim();

    // Check each recurring transaction
    for (let i = 0; i < recurringData.length; i++) {
      const recurRow = recurringData[i];
      const [recurDate, recurDesc, recurMode, recurCategory, recurIncome, recurDebits, recurCcTxnDate] = recurRow;

      // Check if modes match (case-insensitive)
      if ((recurMode || '').toLowerCase() !== (transaction.mode || '').toLowerCase()) {
        continue;
      }

      // Get the amount (either from Income or Debits)
      const recurAmount = Number(recurIncome) || Number(recurDebits) || 0;
      
      // Check if amounts match (within tolerance)
      if (Math.abs(recurAmount - txnAmount) > 0.01) {
        continue;
      }

      // Check if descriptions are similar (case-insensitive, exact or contains)
      const recurDescLower = (recurDesc || '').toLowerCase().trim();
      if (recurDescLower === '' || !txnDesc.includes(recurDescLower) && !recurDescLower.includes(txnDesc)) {
        // Try fuzzy match - check if key words match
        const txnWords = txnDesc.split(/\s+/);
        const recurWords = recurDescLower.split(/\s+/);
        const commonWords = txnWords.filter(word => word.length > 3 && recurWords.includes(word));
        if (commonWords.length === 0) {
          continue;
        }
      }

      // For credit card transactions, check CCTransactionDate (original txn date)
      // For non-CC, check the Date field
      let dateToCompare = null;
      let transactionDateToCompare = transaction.date;
      
      // Normalize mode for comparison (handle "CreditCard" and "Credit Card")
      const normalizedTxnMode = (transaction.mode || '').trim().toLowerCase().replace(/\s+/g, '');
      const isCreditCard = normalizedTxnMode === 'creditcard';
      
      // If transaction has CCTransactionDate, use that for comparison
      if (isCreditCard && transaction.ccTransactionDate) {
        transactionDateToCompare = transaction.ccTransactionDate;
      }
      
      if (isCreditCard && recurCcTxnDate) {
        // For CC transactions, compare transaction's CCTransactionDate with recurring's CCTransactionDate
        dateToCompare = toSafeDate(recurCcTxnDate);
      } else {
        // For non-CC, compare with Date
        dateToCompare = toSafeDate(recurDate);
      }

      if (dateToCompare) {
        const txnDateObj = toSafeDate(transactionDateToCompare);
        if (txnDateObj) {
          const dateDiff = Math.abs((txnDateObj - dateToCompare) / (1000 * 60 * 60 * 24)); // days difference
          
          // If it's a credit card transaction, check if the transaction date matches exactly
          // For other modes, allow some flexibility (within 3 days)
          const maxDaysDiff = isCreditCard ? 0 : 3;
          
          if (dateDiff <= maxDaysDiff) {
            const displayDate = Utilities.formatDate(dateToCompare, Session.getScriptTimeZone(), 'MM/dd/yyyy');
            return {
              isDuplicate: true,
              matchInfo: `Matches recurring: "${recurDesc}" (${recurAmount}) on ${displayDate}`
            };
          }
        }
      }
    }

    return { isDuplicate: false, matchInfo: null };
  }

  /**
   * Log credit card transactions from CSV batch input
   * CSV format: Date,Description,Mode,Category,Income,Debits,CCTransactionDate
   */
  function logCreditCardTransactions() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const variableSheet = ss.getSheetByName('VariableExpenses') || ss.getSheetByName('VariableExpences');

    if (!variableSheet) {
      SpreadsheetApp.getUi().alert('VariableExpenses sheet not found.');
      return;
    }

    // Get CSV input from user using HTML dialog for multi-line support
    const htmlOutput = HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            label { display: block; margin-bottom: 10px; font-weight: bold; }
            textarea { width: 100%; height: 400px; font-family: monospace; padding: 10px; box-sizing: border-box; }
            .example { background: #f5f5f5; padding: 10px; margin: 10px 0; font-size: 12px; font-family: monospace; }
            .buttons { margin-top: 20px; text-align: right; }
            button { padding: 10px 20px; margin-left: 10px; cursor: pointer; }
            .info { font-size: 12px; color: #666; margin-top: 10px; }
          </style>
        </head>
        <body>
          <label>Paste CSV transactions (one per line):</label>
          <div class="example">
            Format: Date,Description,Mode,Category,Income,Debits,CCTransactionDate<br>
            Example:<br>
            2026-02-15,Starbucks Coffee,CreditCard,Food,,150,2026-01-15<br>
            2026-02-15,Gas Station,CreditCard,Transportation,,2000,2026-01-20
          </div>
          <textarea id="csvInput" placeholder="Paste your CSV data here..."></textarea>
          <div class="info">Note: Leave Income or Debits empty (use one or the other, not both)</div>
          <div class="buttons">
            <button onclick="cancel()">Cancel</button>
            <button onclick="submit()" style="background: #4285f4; color: white; border: none;">Submit</button>
          </div>
          <script>
            function submit() {
              const csvText = document.getElementById('csvInput').value.trim();
              if (!csvText) {
                alert('Please enter CSV data.');
                return;
              }
              google.script.host.setHeight(600);
              google.script.run.withSuccessHandler(onSuccess).withFailureHandler(onFailure).processCsvInput(csvText);
            }
            function cancel() {
              google.script.host.close();
            }
            function onSuccess(result) {
              google.script.host.close();
              if (result && result.error) {
                alert(result.error);
              }
            }
            function onFailure(error) {
              alert('Error: ' + error.message);
            }
          </script>
        </body>
      </html>
    `)
    .setWidth(800)
    .setHeight(600);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Log Credit Card Transactions');
  }

  /**
   * Process CSV input from HTML dialog
   * This function is called by the HTML dialog
   */
  function processCsvInput(csvText) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ui = SpreadsheetApp.getUi();
    const variableSheet = ss.getSheetByName('VariableExpenses') || ss.getSheetByName('VariableExpences');
    const recurExpensesSheet = ss.getSheetByName('RecurExpenses');
    const recurIncomeSheet = ss.getSheetByName('RecurIncome');

    if (!variableSheet) {
      return { error: 'VariableExpenses sheet not found.' };
    }

    if (!csvText || csvText.trim() === '') {
      return { error: 'No input provided.' };
    }

    // Parse CSV
    const parseResult = parseCsvTransactions(csvText);
    if (parseResult.errors.length > 0 && parseResult.transactions.length === 0) {
      return { error: 'CSV parsing errors:\n\n' + parseResult.errors.join('\n') };
    }

    // Generate recurring transactions in memory for duplicate checking
    const recurringData = generateRecurringTransactions(recurExpensesSheet, recurIncomeSheet);

    // Validate and process transactions
    const validTransactions = [];
    const invalidTransactions = [];
    const duplicateTransactions = [];

    parseResult.transactions.forEach(txn => {
      const validation = validateTransaction(txn);
      
      if (!validation.valid) {
        invalidTransactions.push({
          transaction: txn,
          errors: validation.errors
        });
        return;
      }

      const validatedTxn = validation.validatedData;
      
      // Check for duplicates against recurring transactions
      const duplicateCheck = checkDuplicateInRecurring(validatedTxn, recurringData);
      if (duplicateCheck.isDuplicate) {
        duplicateTransactions.push({
          transaction: validatedTxn,
          matchInfo: duplicateCheck.matchInfo
        });
        return;
      }

      validTransactions.push(validatedTxn);
    });

    // Build preview message
    let previewMessage = `Preview:\n\n`;
    previewMessage += `Valid transactions to add: ${validTransactions.length}\n`;
    previewMessage += `Duplicates found: ${duplicateTransactions.length}\n`;
    previewMessage += `Invalid transactions: ${invalidTransactions.length}\n\n`;

    if (duplicateTransactions.length > 0) {
      previewMessage += `Duplicates (will be skipped):\n`;
      duplicateTransactions.forEach((dup, idx) => {
        previewMessage += `${idx + 1}. ${dup.transaction.description} - ${dup.matchInfo}\n`;
      });
      previewMessage += '\n';
    }

    if (invalidTransactions.length > 0) {
      previewMessage += `Invalid transactions (will be skipped):\n`;
      invalidTransactions.forEach((inv, idx) => {
        previewMessage += `${idx + 1}. Line ${inv.transaction.lineNumber}: ${inv.errors.join(', ')}\n`;
      });
      previewMessage += '\n';
    }

    if (validTransactions.length === 0) {
      return { error: previewMessage + '\nNo valid transactions to add.' };
    }

    previewMessage += `Valid transactions:\n`;
    validTransactions.slice(0, 10).forEach((txn, idx) => {
      const amount = txn.income > 0 ? `Income: ${txn.income}` : `Debits: ${txn.debits}`;
      previewMessage += `${idx + 1}. ${txn.description} - ${amount} (${txn.category}, ${txn.mode})\n`;
    });
    if (validTransactions.length > 10) {
      previewMessage += `... and ${validTransactions.length - 10} more\n`;
    }

    // Confirm with user
    const confirmResponse = ui.alert(
      'Confirm Transaction Logging',
      previewMessage + '\n\nProceed with adding these transactions?',
      ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
      return;
    }

    // Prepare rows for VariableExpenses
    const rowsToAdd = validTransactions.map(txn => {
      let date = txn.date;
      let ccTxnDate = txn.ccTransactionDate || '';

      // Normalize mode to handle both "CreditCard" and "Credit Card"
      const normalizedMode = (txn.mode || '').trim().toLowerCase().replace(/\s+/g, '');
      const isCreditCard = normalizedMode === 'creditcard';

      // For credit card transactions:
      if (isCreditCard) {
        if (txn.ccTransactionDate) {
          // If CCTransactionDate is provided, check if Date is already the due date
          const txnDateObj = toSafeDate(txn.ccTransactionDate);
          const providedDateObj = toSafeDate(txn.date);
          
          if (txnDateObj && providedDateObj) {
            // Compute expected due date from CCTransactionDate
            const statementDate = computeCcStatementDate(txn.ccTransactionDate);
            const expectedDueDate = computeCcDueDate(statementDate);
            const expectedDueDateObj = toSafeDate(expectedDueDate);
            
            // If Date and CCTransactionDate are different, assume Date is already the due date
            // If they're the same, compute the due date from CCTransactionDate
            const dateDiff = Math.abs((providedDateObj - txnDateObj) / (1000 * 60 * 60 * 24));
            
            if (dateDiff > 1) {
              // Dates are different - Date is likely already the due date, use it as-is
              date = txn.date;
            } else {
              // Dates are the same or very close - Date is likely the transaction date
              // Use computed due date instead
              date = expectedDueDate;
            }
          } else {
            // Fallback: compute due date from CCTransactionDate
            const statementDate = computeCcStatementDate(txn.ccTransactionDate);
            date = computeCcDueDate(statementDate);
          }
          // ccTxnDate is already set above
        } else {
          // If CC mode but no CCTransactionDate provided, assume Date is transaction date
          const statementDate = computeCcStatementDate(txn.date);
          const dueDate = computeCcDueDate(statementDate);
          date = dueDate;
          ccTxnDate = txn.date; // Store original as CCTransactionDate
        }
      }
      
      // Normalize mode to "CreditCard" (remove spaces) for consistency
      const normalizedModeForSheet = isCreditCard ? 'CreditCard' : txn.mode;

      // VariableExpenses columns: Date, Description, Mode, Category, Income, Debits, CCTransactionDate
      return [
        date,                    // Date (due date for CC if CCTransactionDate provided, otherwise computed)
        txn.description,        // Description
        normalizedModeForSheet,  // Mode (normalized to "CreditCard" if credit card)
        txn.category,           // Category
        txn.income || '',       // Income
        txn.debits || '',       // Debits
        ccTxnDate               // CCTransactionDate
      ];
    });

    // Write to VariableExpenses
    const lastRow = variableSheet.getLastRow();
    const insertRow = lastRow > 1 ? lastRow + 1 : 2;
    
    // Ensure headers exist
    if (lastRow <= 1) {
      const headers = ['Date', 'Description', 'Mode', 'Category', 'Income', 'Debits', 'CCTransactionDate'];
      variableSheet.getRange(1, 1, 1, 7).setValues([headers]);
      variableSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    }

    variableSheet.getRange(insertRow, 1, rowsToAdd.length, 7).setValues(rowsToAdd);

    // Show summary
    let summary = `Transactions logged successfully!\n\n`;
    summary += `Added: ${validTransactions.length} transaction(s)\n`;
    if (duplicateTransactions.length > 0) {
      summary += `Skipped (duplicates): ${duplicateTransactions.length}\n`;
    }
    if (invalidTransactions.length > 0) {
      summary += `Skipped (invalid): ${invalidTransactions.length}\n`;
    }
    summary += `\nRemember to run "Update FinalTracker" to include these in the tracker.`;

    ui.alert(summary);
    return { success: true, added: validTransactions.length, duplicates: duplicateTransactions.length, invalid: invalidTransactions.length };
  }

  /**
   * Mark selected rows in FinalTracker as Edited
   * Edited rows will not be overwritten when running Update FinalTracker
   */
  function markSelectedAsEdited() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    if (sheet.getName() !== 'FinalTracker') {
      SpreadsheetApp.getUi().alert('Please select rows in the FinalTracker sheet.');
      return;
    }
    
    const selection = ss.getSelection();
    const activeRange = selection.getActiveRange();
    
    if (!activeRange) {
      SpreadsheetApp.getUi().alert('Please select one or more rows first.');
      return;
    }
    
    const startRow = activeRange.getRow();
    const numRows = activeRange.getNumRows();
    
    // Don't allow editing header row or Starting Balance row
    if (startRow < 2) {
      SpreadsheetApp.getUi().alert('Cannot mark header row as edited.');
      return;
    }
    
    // Check if any selected row is Starting Balance
    const descriptions = sheet.getRange(startRow, 2, numRows, 1).getValues();
    for (let i = 0; i < descriptions.length; i++) {
      if (descriptions[i][0] === 'Starting Balance') {
        SpreadsheetApp.getUi().alert('Cannot mark Starting Balance row as edited.');
        return;
      }
    }
    
    // Set Edited column (column I = 9) to TRUE for selected rows
    const editedValues = [];
    for (let i = 0; i < numRows; i++) {
      editedValues.push([true]);
    }
    
    sheet.getRange(startRow, 9, numRows, 1).setValues(editedValues);
    
    SpreadsheetApp.getUi().alert(`Marked ${numRows} row(s) as Edited.\n\nThese rows will be preserved when running Update FinalTracker.`);
  }

  /**
   * Clear Edited flag from selected rows in FinalTracker
   */
  function clearEditedFlag() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    if (sheet.getName() !== 'FinalTracker') {
      SpreadsheetApp.getUi().alert('Please select rows in the FinalTracker sheet.');
      return;
    }
    
    const selection = ss.getSelection();
    const activeRange = selection.getActiveRange();
    
    if (!activeRange) {
      SpreadsheetApp.getUi().alert('Please select one or more rows first.');
      return;
    }
    
    const startRow = activeRange.getRow();
    const numRows = activeRange.getNumRows();
    
    if (startRow < 2) {
      SpreadsheetApp.getUi().alert('Cannot modify header row.');
      return;
    }
    
    // Clear Edited column (column I = 9) for selected rows
    sheet.getRange(startRow, 9, numRows, 1).clearContent();
    
    SpreadsheetApp.getUi().alert(`Cleared Edited flag from ${numRows} row(s).`);
  }

  // ============================================================================
  // SAVINGS GOALS TRACKER
  // ============================================================================

  /**
   * Calculate progress for each savings goal
   * If Account is specified, uses that account's SavingsAmount from CurrentBalance
   * Otherwise, looks for transactions where Category = "Savings" and Description contains the goal name
   */
  function calculateGoalProgress() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const goalsSheet = ss.getSheetByName('SavingsGoals');
    const finalTrackerSheet = ss.getSheetByName('FinalTracker');
    const currentBalanceSheet = ss.getSheetByName('CurrentBalance');
    
    if (!goalsSheet) {
      SpreadsheetApp.getUi().alert('SavingsGoals sheet not found. Please create it first.');
      return;
    }
    
    const goalsLastRow = goalsSheet.getLastRow();
    if (goalsLastRow < 2) {
      SpreadsheetApp.getUi().alert('No goals found in SavingsGoals sheet.');
      return;
    }
    
    // Read goals (columns A-K, including Account)
    const goalsData = goalsSheet.getRange(2, 1, goalsLastRow - 1, 11).getValues();
    
    // Build account savings map from CurrentBalance
    const accountSavings = {};
    if (currentBalanceSheet && currentBalanceSheet.getLastRow() > 1) {
      // Columns: Account(0), Amount(1), MaintainBalance(2), SavingsAmount(3)
      const cbData = currentBalanceSheet.getRange(2, 1, currentBalanceSheet.getLastRow() - 1, 4).getValues();
      cbData.forEach(row => {
        const accountName = (row[0] || '').toString().trim().toLowerCase();
        const savingsAmount = Number(row[3]) || 0;
        if (accountName) {
          accountSavings[accountName] = savingsAmount;
        }
      });
    }
    
    // Read FinalTracker transactions (for goals without Account specified)
    let ftData = [];
    if (finalTrackerSheet && finalTrackerSheet.getLastRow() > 1) {
      // Columns: Date(0), Description(1), Mode(2), Category(3), Income(4), Debits(5)
      ftData = finalTrackerSheet.getRange(2, 1, finalTrackerSheet.getLastRow() - 1, 6).getValues();
    }
    
    // Calculate progress for each goal
    const progressUpdates = [];
    
    goalsData.forEach((goal, idx) => {
      const goalDescription = (goal[0] || '').toString().trim().toLowerCase();
      const isActive = goal[9] === true || goal[9] === 'TRUE' || goal[9] === 1;
      const linkedAccount = (goal[10] || '').toString().trim().toLowerCase();
      
      if (!goalDescription || !isActive) {
        progressUpdates.push([goal[7] || 0]); // Keep existing progress
        return;
      }
      
      let totalSaved = 0;
      
      // If Account is specified, use that account's SavingsAmount
      if (linkedAccount && accountSavings.hasOwnProperty(linkedAccount)) {
        totalSaved = accountSavings[linkedAccount];
      } else {
        // Otherwise, sum transactions where Category = "Savings" and Description contains goal name
        ftData.forEach(tx => {
          const txDescription = (tx[1] || '').toString().toLowerCase();
          const txCategory = (tx[3] || '').toString().toLowerCase();
          const txIncome = Number(tx[4]) || 0;
          const txDebit = Number(tx[5]) || 0;
          
          // Check if this transaction is for this goal
          if (txCategory === 'savings' && txDescription.includes(goalDescription)) {
            // Income adds to savings, Debits subtract (withdrawal)
            totalSaved += txIncome - txDebit;
          }
        });
      }
      
      progressUpdates.push([totalSaved]);
    });
    
    // Write progress to column H (CurrentProgress)
    if (progressUpdates.length > 0) {
      goalsSheet.getRange(2, 8, progressUpdates.length, 1).setValues(progressUpdates);
    }
    
    return progressUpdates.length;
  }

  /**
   * Update calculated values for each savings goal
   * - Allotment mode: Calculate estimated completion date
   * - Deadline mode: Calculate required amount per occurrence
   */
  function updateGoalsCalculations() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const goalsSheet = ss.getSheetByName('SavingsGoals');
    
    if (!goalsSheet) {
      SpreadsheetApp.getUi().alert('SavingsGoals sheet not found. Please create it first.');
      return;
    }
    
    const goalsLastRow = goalsSheet.getLastRow();
    if (goalsLastRow < 2) {
      SpreadsheetApp.getUi().alert('No goals found in SavingsGoals sheet.');
      return;
    }
    
    // First, update progress from transactions
    calculateGoalProgress();
    
    // Re-read goals with updated progress
    const goalsData = goalsSheet.getRange(2, 1, goalsLastRow - 1, 11).getValues();
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const calculatedUpdates = [];
    
    goalsData.forEach((goal, idx) => {
      const description = goal[0];
      const targetAmount = Number(goal[1]) || 0;
      const frequency = (goal[2] || '').toString().toUpperCase();
      const day = goal[3]; // Keep as-is to support "15, 28" format
      const mode = (goal[4] || '').toString().toLowerCase();
      const allotmentAmount = Number(goal[5]) || 0;
      const targetDate = goal[6] ? new Date(goal[6]) : null;
      const currentProgress = Number(goal[7]) || 0;
      const isActive = goal[9] === true || goal[9] === 'TRUE' || goal[9] === 1;
      
      if (!description || !isActive || targetAmount <= 0) {
        calculatedUpdates.push(['']);
        return;
      }
      
      const remaining = targetAmount - currentProgress;
      
      if (remaining <= 0) {
        calculatedUpdates.push(['GOAL REACHED!']);
        return;
      }
      
      if (mode === 'allotment' && allotmentAmount > 0) {
        // Calculate estimated completion date
        const occurrencesNeeded = Math.ceil(remaining / allotmentAmount);
        const estimatedDate = calculateFutureDate(today, frequency, day, occurrencesNeeded);
        calculatedUpdates.push([estimatedDate ? Utilities.formatDate(estimatedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'Unable to calculate']);
      } else if (mode === 'deadline' && targetDate) {
        // Calculate required amount per occurrence
        const occurrencesUntilDeadline = countOccurrences(today, targetDate, frequency, day);
        if (occurrencesUntilDeadline > 0) {
          const requiredPerOccurrence = remaining / occurrencesUntilDeadline;
          calculatedUpdates.push([requiredPerOccurrence.toFixed(2) + ' per occurrence']);
        } else {
          calculatedUpdates.push(['Deadline passed or no occurrences']);
        }
      } else {
        calculatedUpdates.push(['Set mode and amount/date']);
      }
    });
    
    // Write calculated values to column I
    if (calculatedUpdates.length > 0) {
      goalsSheet.getRange(2, 9, calculatedUpdates.length, 1).setValues(calculatedUpdates);
    }
    
    SpreadsheetApp.getUi().alert(`Updated ${calculatedUpdates.length} goal(s).`);
  }

  /**
   * Parse day value which can be a single number or comma-separated (e.g., "15" or "15, 28")
   * Returns sorted array of day numbers
   */
  function parseDays(day) {
    const dayStr = (day || '').toString();
    if (dayStr.includes(',')) {
      return dayStr.split(',').map(d => Number(d.trim())).filter(d => !isNaN(d) && d >= 1 && d <= 31).sort((a, b) => a - b);
    }
    const num = Number(day);
    return (!isNaN(num) && num >= 1 && num <= 31) ? [num] : [1];
  }

  /**
   * Calculate a future date based on frequency and number of occurrences
   * Supports multiple days for N_DAY_IN_MONTH (e.g., "15, 28")
   */
  function calculateFutureDate(startDate, frequency, day, occurrences) {
    const result = new Date(startDate);
    
    if (frequency === 'N_DAY_IN_MONTH') {
      const days = parseDays(day);
      let count = 0;
      
      while (count < occurrences) {
        // Find next occurrence
        let found = false;
        const currentDay = result.getDate();
        const currentMonth = result.getMonth();
        const currentYear = result.getFullYear();
        
        // Check remaining days in current month
        for (const d of days) {
          const lastDayOfMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
          const actualDay = Math.min(d, lastDayOfMonth);
          if (actualDay > currentDay) {
            result.setDate(actualDay);
            count++;
            found = true;
            if (count >= occurrences) break;
          }
        }
        
        // Move to next month if needed
        if (!found || count < occurrences) {
          result.setMonth(result.getMonth() + 1);
          result.setDate(1);
          
          // Add occurrences for the new month
          if (count < occurrences) {
            const lastDayOfMonth = new Date(result.getFullYear(), result.getMonth() + 1, 0).getDate();
            for (const d of days) {
              if (count >= occurrences) break;
              const actualDay = Math.min(d, lastDayOfMonth);
              result.setDate(actualDay);
              count++;
            }
          }
        }
      }
      
      return result;
    }
    
    // Other frequencies (unchanged)
    for (let i = 0; i < occurrences; i++) {
      switch (frequency) {
        case 'WEEKLY':
          result.setDate(result.getDate() + 7);
          break;
        case 'ANNUAL':
          result.setFullYear(result.getFullYear() + 1);
          break;
        case 'EVERY_N_DAYS':
          result.setDate(result.getDate() + (Number(day) || 1));
          break;
        default:
          // Default to monthly
          result.setMonth(result.getMonth() + 1);
      }
    }
    
    return result;
  }

  /**
   * Count number of occurrences between two dates based on frequency
   * Supports multiple days for N_DAY_IN_MONTH (e.g., "15, 28")
   */
  function countOccurrences(startDate, endDate, frequency, day) {
    if (endDate <= startDate) return 0;
    
    const diffMs = endDate - startDate;
    const diffDays = diffMs / (1000 * 60 * 60 * 24);
    
    switch (frequency) {
      case 'WEEKLY':
        return Math.floor(diffDays / 7);
      case 'N_DAY_IN_MONTH': {
        // Count occurrences for multiple days per month
        const days = parseDays(day);
        let count = 0;
        const current = new Date(startDate);
        current.setHours(12, 0, 0, 0);
        
        while (current < endDate) {
          const lastDayOfMonth = new Date(current.getFullYear(), current.getMonth() + 1, 0).getDate();
          for (const d of days) {
            const actualDay = Math.min(d, lastDayOfMonth);
            const checkDate = new Date(current.getFullYear(), current.getMonth(), actualDay, 12, 0, 0, 0);
            if (checkDate > startDate && checkDate <= endDate) {
              count++;
            }
          }
          current.setMonth(current.getMonth() + 1);
          current.setDate(1);
        }
        return count;
      }
      case 'ANNUAL':
        return Math.floor(diffDays / 365);
      case 'EVERY_N_DAYS':
        return Math.floor(diffDays / (Number(day) || 1));
      default:
        return Math.floor(diffDays / 30); // Default to monthly
    }
  }

  /**
   * Display a summary of all savings goals
   */
  function viewGoalsSummary() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const goalsSheet = ss.getSheetByName('SavingsGoals');
    
    if (!goalsSheet) {
      SpreadsheetApp.getUi().alert('SavingsGoals sheet not found. Please create it first.');
      return;
    }
    
    const goalsLastRow = goalsSheet.getLastRow();
    if (goalsLastRow < 2) {
      SpreadsheetApp.getUi().alert('No goals found in SavingsGoals sheet.');
      return;
    }
    
    // Update calculations first
    calculateGoalProgress();
    
    // Re-read goals
    const goalsData = goalsSheet.getRange(2, 1, goalsLastRow - 1, 11).getValues();
    
    let message = '=== SAVINGS GOALS SUMMARY ===\n\n';
    let totalTarget = 0;
    let totalProgress = 0;
    let activeGoals = 0;
    
    goalsData.forEach((goal, idx) => {
      const description = goal[0];
      const targetAmount = Number(goal[1]) || 0;
      const mode = (goal[4] || '').toString();
      const currentProgress = Number(goal[7]) || 0;
      const calculated = goal[8] || '';
      const isActive = goal[9] === true || goal[9] === 'TRUE' || goal[9] === 1;
      const linkedAccount = (goal[10] || '').toString().trim();
      
      if (!description || !isActive) return;
      
      activeGoals++;
      totalTarget += targetAmount;
      totalProgress += currentProgress;
      
      const percentage = targetAmount > 0 ? ((currentProgress / targetAmount) * 100).toFixed(1) : 0;
      const remaining = targetAmount - currentProgress;
      
      message += `${description}\n`;
      if (linkedAccount) {
        message += `  Account: ${linkedAccount}\n`;
      }
      message += `  Target: ${targetAmount.toFixed(2)}\n`;
      message += `  Progress: ${currentProgress.toFixed(2)} (${percentage}%)\n`;
      message += `  Remaining: ${remaining.toFixed(2)}\n`;
      message += `  Mode: ${mode}\n`;
      if (calculated) {
        message += `  ${mode === 'allotment' ? 'Est. Completion' : 'Required'}: ${calculated}\n`;
      }
      message += '\n';
    });
    
    if (activeGoals > 0) {
      const totalPercentage = totalTarget > 0 ? ((totalProgress / totalTarget) * 100).toFixed(1) : 0;
      message += `--- TOTALS ---\n`;
      message += `Active Goals: ${activeGoals}\n`;
      message += `Total Target: ${totalTarget.toFixed(2)}\n`;
      message += `Total Progress: ${totalProgress.toFixed(2)} (${totalPercentage}%)\n`;
      message += `Total Remaining: ${(totalTarget - totalProgress).toFixed(2)}\n`;
    } else {
      message = 'No active goals found.';
    }
    
    SpreadsheetApp.getUi().alert(message);
  }

  // ============================================================================
  // DASHBOARD - DAILY BUDGET
  // ============================================================================

  /**
   * Setup Dashboard sheet with formula-based Daily Budget calculations
   * Uses Option B + C: Daily Rate (fixed) + Cumulative Budget + Available to Spend (1:1 impact)
   */
  function setupDashboard() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let dashboardSheet = ss.getSheetByName('Dashboard');
    
    if (!dashboardSheet) {
      dashboardSheet = ss.insertSheet('Dashboard');
    }
    
    // Clear existing content
    dashboardSheet.clear();
    
    // Labels in column A, formulas in column B
    const labels = [
      ['DAILY BUDGET DASHBOARD'],      // Row 1
      [''],                            // Row 2
      ['Cycle 1 Start:'],              // Row 3
      ['Target Date:'],                // Row 4
      ['Total Days:'],                 // Row 5
      ['Days Elapsed:'],               // Row 6
      [''],                            // Row 7
      ['Projected Balance:'],          // Row 8
      ['Daily Rate:'],                 // Row 9
      ['Cumulative Budget:'],          // Row 10
      [''],                            // Row 11
      ['Starting Balance:'],           // Row 12
      ['Current Balance:'],            // Row 13
      ['Actual Spent:'],               // Row 14
      [''],                            // Row 15
      ['Reconciliation Status:'],      // Row 16
      [''],                            // Row 17
      [''],                            // Row 18 // Second blank row for spacing
      ['Today Budget:'],               // Row 19 - will be overwritten with formula
      ['Tomorrow Budget:'],            // Row 20 - will be overwritten with formula
      ['Day After Budget:'],           // Row 21 - will be overwritten with formula
      [''],                            // Row 22 // blank separator
      ['WEEKLY BUDGET'],               // Row 23 // WEEKLY BUDGET title
      [''],                            // Row 24 // blank (title row, no label needed)
      ['Weekly Rate:'],                // Row 25
      [''],                            // Row 26 // blank space between Weekly Rate and Current Week
      ['Current Week:'],               // Row 27 - will be overwritten with formula
      ['Next Week: Mon-Sun']           // Row 28
    ];
    
    // Formulas in column B (with error handling for fresh start)
    const formulas = [
      [''], // Row 1: title
      [''], // Row 2: blank
      ['=IFERROR(LARGE(FILTER(FinalTracker!A:A, FinalTracker!B:B="Salary B", FinalTracker!A:A<=TODAY()), 1), INDEX(FinalTracker!A:A, MATCH("Starting Balance", FinalTracker!B:B, 0)))'], // Row 3: Cycle 1 Start (fallback to Starting Balance date)
      ['=SMALL(FILTER(FinalTracker!A:A, FinalTracker!B:B="Salary B", FinalTracker!A:A>TODAY()), 2) - 1'], // Row 4: Target Date
      ['=IFERROR(B4 - B3 + 1, "")'], // Row 5: Total Days
      ['=IFERROR(TODAY() - B3 + 1, "")'], // Row 6: Days Elapsed
      [''], // Row 7: blank
      ['=IFERROR(INDEX(FinalTracker!H:H, MATCH(B4, FinalTracker!A:A, 1)), "")'], // Row 8: Projected Balance (approximate match - finds row on or before target date)
      ['=IFERROR(B8 / B5, "")'], // Row 9: Daily Rate
      ['=IFERROR(B9 * B6, "")'], // Row 10: Cumulative Budget
      [''], // Row 11: blank
      ['=IFERROR(INDEX(FinalTracker!H:H, MATCH(B3, FinalTracker!A:A, 0)), INDEX(FinalTracker!H:H, MATCH("Starting Balance", FinalTracker!B:B, 0)))'], // Row 12: Starting Balance (fallback to Starting Balance row)
      ['=SUMPRODUCT(CurrentBalance!B2:B) - SUMPRODUCT(CurrentBalance!C2:C) - SUMPRODUCT(CurrentBalance!D2:D)'], // Row 13: Current Balance
      ['=IFERROR(B12 - B13, "")'], // Row 14: Actual Spent
      [''], // Row 15: blank
      ['=IFERROR(IF(ABS(B13 - INDEX(FinalTracker!H:H, MAX(ARRAYFORMULA(IF((ISNUMBER(FinalTracker!A2:A10000))*(INT(FinalTracker!A2:A10000)<=INT(TODAY())), ROW(FinalTracker!A2:A10000), 0))))) < 0.01, "Reconciled", "Not Reconciled: " & TEXT(B13 - INDEX(FinalTracker!H:H, MAX(ARRAYFORMULA(IF((ISNUMBER(FinalTracker!A2:A10000))*(INT(FinalTracker!A2:A10000)<=INT(TODAY())), ROW(FinalTracker!A2:A10000), 0)))), "+#,##0.00;-#,##0.00")), "Not Reconciled: N/A")'], // Row 16: Reconciliation Status - INDEX/MAX/ARRAYFORMULA to get LAST valid date entry (A2:A10000 covers up to 10,000 rows, data already sorted by updateFinalTracker)
      [''], // Row 17: first blank spacing
      [''], // Row 18: second blank spacing
      ['=IFERROR(B9 * B6 - B14, "")'], // Row 19: Today Budget (Daily Rate × Days Elapsed - Actual Spent) - numeric only
      ['=IFERROR(B9 * (B6 + 1) - B14, "")'], // Row 20: Tomorrow Budget - numeric only
      ['=IFERROR(B9 * (B6 + 2) - B14, "")'], // Row 21: Day After Budget - numeric only
      [''], // Row 22: blank separator
      [''], // Row 23: WEEKLY BUDGET title (blank in formulas, label only)
      [''], // Row 24: blank (title row, no formula)
      ['=IFERROR(B8 * 7 / B5, "")'], // Row 25: Weekly Rate = Projected / (Total Days / 7) - value in B25
      [''], // Row 26: blank space
      ['=IFERROR(B25 * (B6 / 7) - B14, "")'], // Row 27: Current Week Budget - numeric only
      ['=IFERROR(B25 * ((B6 / 7) + 1) - B14, "")'] // Row 28: Next Week Budget - numeric only
    ];
    
    // Write labels
    dashboardSheet.getRange(1, 1, labels.length, 1).setValues(labels);
    
    // Write formulas
    dashboardSheet.getRange(1, 2, formulas.length, 1).setFormulas(formulas);
    
    // Set dynamic day labels in column A (formulas for rows 19, 20, 21, 27)
    dashboardSheet.getRange('A19').setFormula('="Today Budget: " & TEXT(TODAY(), "ddd")');
    dashboardSheet.getRange('A20').setFormula('="Tomorrow Budget: " & TEXT(TODAY()+1, "ddd")');
    dashboardSheet.getRange('A21').setFormula('="Day After Budget: " & TEXT(TODAY()+2, "ddd")');
    dashboardSheet.getRange('A27').setFormula('="Current Week: " & TEXT(TODAY(), "ddd") & "-Sun"');
    
    // Set number format for budget values (rows 19-21, 27-28)
    dashboardSheet.getRange('B19').setNumberFormat('#,##0.00');
    dashboardSheet.getRange('B20').setNumberFormat('#,##0.00');
    dashboardSheet.getRange('B21').setNumberFormat('#,##0.00');
    dashboardSheet.getRange('B27').setNumberFormat('#,##0.00');
    dashboardSheet.getRange('B28').setNumberFormat('#,##0.00');
    
    // Formatting
    // Title row (daily budget)
    dashboardSheet.getRange(1, 1, 1, 2).merge().setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
    
    // Weekly budget title (row 23, merged 2 cells like row 1)
    dashboardSheet.getRange(23, 1, 1, 2).merge().setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
    
    // Value columns - number format
    dashboardSheet.getRange(3, 2, 1, 1).setNumberFormat('yyyy-mm-dd'); // Cycle 1 Start
    dashboardSheet.getRange(4, 2, 1, 1).setNumberFormat('yyyy-mm-dd'); // Target Date
    dashboardSheet.getRange(5, 2, 1, 1).setNumberFormat('0'); // Total Days
    dashboardSheet.getRange(6, 2, 1, 1).setNumberFormat('0'); // Days Elapsed
    dashboardSheet.getRange(8, 2, 1, 1).setNumberFormat('#,##0.00'); // Projected Balance
    dashboardSheet.getRange(9, 2, 1, 1).setNumberFormat('#,##0.00'); // Daily Rate
    dashboardSheet.getRange(10, 2, 1, 1).setNumberFormat('#,##0.00'); // Cumulative Budget
    dashboardSheet.getRange(12, 2, 1, 1).setNumberFormat('#,##0.00'); // Starting Balance
    dashboardSheet.getRange(13, 2, 1, 1).setNumberFormat('#,##0.00'); // Current Balance
    dashboardSheet.getRange(14, 2, 1, 1).setNumberFormat('#,##0.00'); // Actual Spent
    
    // Reconciliation Status formatting (row 16)
    const reconciliationStatusCell = dashboardSheet.getRange(16, 2, 1, 1);
    reconciliationStatusCell.setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
    // Apply conditional background color based on status text
    // Green for "Reconciled", red for "Not Reconciled"
    // Use whenTextContains for text matching
    const greenRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([reconciliationStatusCell])
      .whenTextEqualTo('Reconciled')
      .setBackground('#1FC71F')
      .build();
    const redRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([reconciliationStatusCell])
      .whenTextContains('Not Reconciled')
      .setBackground('#DF0000')
      .build();
    // Get existing rules and add new ones
    const existingRules = dashboardSheet.getConditionalFormatRules();
    existingRules.push(greenRule, redRule);
    dashboardSheet.setConditionalFormatRules(existingRules);
    
    // Reconcile button - user will create drawing button manually
    // Space is reserved below reconciliation status (rows 16-17 are blank)
    
    // Today Budget - emphasized (row 19)
    dashboardSheet.getRange(19, 2, 1, 1).setFontWeight('bold').setFontSize(14);
    dashboardSheet.getRange(19, 1, 1, 1).setFontWeight('bold').setFontSize(12);
    
    // Tomorrow and Day After - less emphasis (rows 20-21)
    dashboardSheet.getRange(20, 2, 1, 1).setFontSize(10); // Tomorrow Budget
    dashboardSheet.getRange(21, 2, 1, 1).setFontSize(10); // Day After Budget
    
    // Weekly budget formatting
    dashboardSheet.getRange(25, 2, 1, 1).setNumberFormat('#,##0.00'); // Weekly Rate (B25)
    dashboardSheet.getRange(27, 2, 1, 1).setFontWeight('bold').setFontSize(12); // Current Week Budget (B27)
    dashboardSheet.getRange(28, 2, 1, 1).setFontSize(10); // Next Week Budget (B28)
    
    // Bold labels
    dashboardSheet.getRange(3, 1, 16, 1).setFontWeight('bold'); // Updated to include reconciliation status row
    dashboardSheet.getRange(25, 1, 1, 1).setFontWeight('bold'); // Weekly Rate label (row 25)
    dashboardSheet.getRange(27, 1, 2, 1).setFontWeight('bold'); // Current Week and Next Week labels (rows 27-28)
    
    // Set column widths
    dashboardSheet.setColumnWidth(1, 200);
    dashboardSheet.setColumnWidth(2, 150);
    
    // ============================================================================
    // SAVINGS GOALS DASHBOARD (Columns D-G)
    // ============================================================================
    
    // Title row (row 1)
    dashboardSheet.getRange(1, 4, 1, 4).merge().setValue('SAVINGS GOALS').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
    
    // Row 2 is empty (spacing)
    
    // Header row (row 3)
    dashboardSheet.getRange(3, 4, 1, 1).setValue('Description').setFontWeight('bold');
    dashboardSheet.getRange(3, 5, 1, 1).setValue('Progress %').setFontWeight('bold');
    dashboardSheet.getRange(3, 6, 1, 1).setValue('Progress Bar').setFontWeight('bold');
    dashboardSheet.getRange(3, 7, 1, 1).setValue('Date of Completion').setFontWeight('bold');
    
    // Array formulas to get active savings goals starting from row 4
    // Column D: Description (filter active goals)
    const descriptionFormula = '=FILTER(SavingsGoals!A2:A, SavingsGoals!J2:J=TRUE, SavingsGoals!A2:A<>"")';
    
    // Column E: Progress Percentage (as decimal 0.0-1.0, will be formatted as %)
    const progressPercentFormula = '=FILTER(SavingsGoals!H2:H/SavingsGoals!B2:B, SavingsGoals!J2:J=TRUE, SavingsGoals!A2:A<>"")';
    
    // Column F: Progress Bar (using REPT to create visual bar, references column E)
    const progressBarFormula = '=ARRAYFORMULA(IF(ISBLANK(D4:D), "", REPT("█", ROUND(E4:E*20, 0)) & REPT("░", 20-ROUND(E4:E*20, 0))))';
    
    // Column G: Date of Completion
    // For deadline mode: TargetDate (G) has the deadline date
    // For allotment mode: Calculated (I) has the estimated completion date in "yyyy-MM-dd" format
    // Use TargetDate if it exists, otherwise use Calculated (which may be date or "per occurrence" text)
    // We'll format as date, so non-date text will show as error - but that's acceptable
    const completionDateFormula = '=ARRAYFORMULA(IF(ISBLANK(D4:D), "", IF(FILTER(SavingsGoals!G2:G, SavingsGoals!J2:J=TRUE, SavingsGoals!A2:A<>"")<>"", FILTER(SavingsGoals!G2:G, SavingsGoals!J2:J=TRUE, SavingsGoals!A2:A<>""), FILTER(SavingsGoals!I2:I, SavingsGoals!J2:J=TRUE, SavingsGoals!A2:A<>""))))';
    
    // Apply array formulas starting from row 4
    dashboardSheet.getRange(4, 4, 1, 1).setFormula(descriptionFormula);
    dashboardSheet.getRange(4, 5, 1, 1).setFormula(progressPercentFormula);
    dashboardSheet.getRange(4, 6, 1, 1).setFormula(progressBarFormula);
    dashboardSheet.getRange(4, 7, 1, 1).setFormula(completionDateFormula);
    
    // Formatting for Savings Goals section
    dashboardSheet.getRange(3, 4, 1, 4).setFontWeight('bold').setBackground('#E8E8E8');
    dashboardSheet.getRange(4, 5, 48, 1).setNumberFormat('0.0%'); // Progress Percentage as percentage
    dashboardSheet.getRange(4, 6, 48, 1).setFontFamily('Courier New'); // Monospace font for progress bar
    dashboardSheet.getRange(4, 7, 48, 1).setNumberFormat('yyyy-mm-dd'); // Date format
    
    // Set column widths for Savings Goals
    dashboardSheet.setColumnWidth(4, 200); // Description
    dashboardSheet.setColumnWidth(5, 100); // Progress %
    dashboardSheet.setColumnWidth(6, 250); // Progress Bar
    dashboardSheet.setColumnWidth(7, 150); // Date of Completion
    
    SpreadsheetApp.getUi().alert('Dashboard initialized!\n\nFormulas are set up and will auto-update.\n\nKey metric: "AVAILABLE TO SPEND" shows 1:1 impact of your spending.');
  }

  // ============================================================================
  // EMERGENCY FUND CALCULATOR
  // ============================================================================

  /**
   * Setup Emergency Fund Calculator sheet with formula-based calculations
   * Calculates target based on: Monthly Salary × Percentage × Target Months
   */
  function setupEmergencyFundCalculator() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let emergencyFundSheet = ss.getSheetByName('EmergencyFund');
    
    if (!emergencyFundSheet) {
      emergencyFundSheet = ss.insertSheet('EmergencyFund');
    }
    
    // Clear existing content
    emergencyFundSheet.clear();
    
    // Labels in column A, formulas/values in column B
    const labels = [
      ['EMERGENCY FUND CALCULATOR'],
      [''],
      ['Monthly Salary:'],
      ['Percentage:'],
      ['Target Months:'],
      [''],
      ['Emergency Fund Target:'],
      [''],
      ['Current Progress:'],
      ['Remaining:'],
      ['Progress %:'],
      [''],
      ['Monthly Savings Needed:']
    ];
    
    // Formulas in column B
    // Monthly Salary: Convert biweekly salary to monthly
    // Biweekly pay: 26 pay periods per year, so monthly = (biweekly × 26) / 12
    // Average all salary transactions (biweekly amounts) and convert to monthly
    const monthlySalaryFormula = '=IFERROR(AVERAGEIF(FinalTracker!D:D, "Salary", FinalTracker!E:E) * 26 / 12, "")';
    
    const formulas = [
      [''], // Row 1 - title
      [''], // Row 2 - blank
      [monthlySalaryFormula], // Monthly Salary (converted from biweekly)
      [''], // Row 4 - Percentage (user input, default 100)
      [''], // Row 5 - Target Months (user input, default 6)
      [''], // Row 6 - blank
      ['=IFERROR(B3 * IF(B4>0, B4/100, 1) * IF(B5>0, B5, 6), "")'], // Emergency Fund Target
      [''], // Row 8 - blank
      ['=IFERROR(INDEX(FILTER(SavingsGoals!H:H, ISNUMBER(SEARCH("Emergency Fund", SavingsGoals!A:A))), 1), 0)'], // Current Progress
      ['=IFERROR(B7 - B9, "")'], // Remaining
      ['=IFERROR(B9 / B7 * 100, "")'], // Progress %
      [''], // Row 12 - blank
      ['=IFERROR(B10 / 12, "")'] // Monthly Savings Needed (assuming 12 months to reach target)
    ];
    
    // Write labels
    emergencyFundSheet.getRange(1, 1, labels.length, 1).setValues(labels);
    
    // Write formulas
    emergencyFundSheet.getRange(1, 2, formulas.length, 1).setFormulas(formulas);
    
    // Set default values for user inputs
    emergencyFundSheet.getRange(4, 2, 1, 1).setValue(100); // Percentage default: 100%
    emergencyFundSheet.getRange(5, 2, 1, 1).setValue(6); // Target Months default: 6
    
    // Formatting
    // Title row
    emergencyFundSheet.getRange(1, 1, 1, 2).merge().setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
    
    // Value columns - number format
    emergencyFundSheet.getRange(3, 2, 1, 1).setNumberFormat('#,##0.00'); // Monthly Salary
    emergencyFundSheet.getRange(4, 2, 1, 1).setNumberFormat('0'); // Percentage
    emergencyFundSheet.getRange(5, 2, 1, 1).setNumberFormat('0'); // Target Months
    emergencyFundSheet.getRange(7, 2, 1, 1).setNumberFormat('#,##0.00').setFontWeight('bold').setFontSize(12); // Emergency Fund Target
    emergencyFundSheet.getRange(9, 2, 1, 1).setNumberFormat('#,##0.00'); // Current Progress
    emergencyFundSheet.getRange(10, 2, 1, 1).setNumberFormat('#,##0.00'); // Remaining
    emergencyFundSheet.getRange(11, 2, 1, 1).setNumberFormat('0.0%'); // Progress %
    emergencyFundSheet.getRange(13, 2, 1, 1).setNumberFormat('#,##0.00'); // Monthly Savings Needed
    
    // Bold labels
    emergencyFundSheet.getRange(3, 1, 11, 1).setFontWeight('bold');
    
    // Set column widths
    emergencyFundSheet.setColumnWidth(1, 200);
    emergencyFundSheet.setColumnWidth(2, 150);
    
    SpreadsheetApp.getUi().alert('Emergency Fund Calculator initialized!\n\nDefault: 100% of monthly salary × 6 months\n\nYou can adjust Percentage and Target Months in the sheet.');
  }

  function onOpen() {
    // Budget Tracker menu (features)
    SpreadsheetApp.getUi()
      .createMenu('Budget Tracker')
      .addItem('Update FinalTracker', 'generateRecurring')
      .addItem('Log Transactions', 'logCreditCardTransactions')
      .addItem('Normalize CC Dates', 'normalizeVariableExpensesCcDates')
      .addSeparator()
      .addItem('Mark Selected as Edited', 'markSelectedAsEdited')
      .addItem('Clear Edited Flag', 'clearEditedFlag')
      .addSeparator()
      .addItem('View Balance Summary', 'viewBalanceSummary')
      .addItem('Refresh Starting Balance Value', 'refreshStartingBalanceValue')
      .addSeparator()
      .addItem('Setup Running Balance', 'setupRunningBalance')
      .addItem('Reset Starting Balance (Date + Value)', 'updateStartingBalance')
      .addItem('Reconcile Balance', 'reconcileBalance')
      .addItem('Recalculate Daily Budget', 'recalcDailyBudget')
      .addSeparator()
      .addItem('Update Goals Progress', 'updateGoalsCalculations')
      .addItem('View Goals Summary', 'viewGoalsSummary')
      .addSeparator()
      .addItem('Inspect Sheet Structure', 'inspectSheetStructure')
      .addItem('Export Diagnostics', 'exportDiagnostics')
      .addToUi();
    
    // Schema menu (sheet structure management)
    addSchemaMenu();
  }

  /**
   * Diagnostic function to inspect sheet structure and help understand the current state
   * Writes output to a "Diagnostics" sheet and also logs to console
   * Run this from Apps Script editor to see sheet information
   */
  function inspectSheetStructure() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets();
    
    // Create or get Diagnostics sheet
    let diagSheet = ss.getSheetByName('Diagnostics');
    if (!diagSheet) {
      diagSheet = ss.insertSheet('Diagnostics');
    } else {
      diagSheet.clear();
    }
    
    let output = '=== SHEET STRUCTURE ===\n\n';
    const rows = [['Sheet Name', 'Rows', 'Columns', 'Headers', 'Sample Data']];
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      output += `Sheet: "${name}"\n`;
      output += `  Rows: ${lastRow}, Columns: ${lastCol}\n`;
      
      let headersStr = '';
      let sampleDataStr = '';
      
      if (lastRow > 0 && lastCol > 0) {
        // Get headers
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        headersStr = headers.join(' | ');
        output += `  Headers: ${headersStr}\n`;
        
        // Get sample data (first 3 rows after header)
        if (lastRow > 1) {
          const sampleRows = Math.min(3, lastRow - 1);
          const sampleData = sheet.getRange(2, 1, sampleRows, lastCol).getValues();
          sampleDataStr = sampleData.map((row, idx) => `Row ${idx + 2}: ${row.join(' | ')}`).join('\n');
          output += `  Sample data (${sampleRows} rows):\n`;
          sampleData.forEach((row, idx) => {
            output += `    Row ${idx + 2}: ${row.join(' | ')}\n`;
          });
        }
      }
      
      // Check for specific sheets mentioned in code
      if (name === 'FinalTracker') {
        const data = sheet.getRange(1, 1, Math.min(5, lastRow), lastCol).getValues();
        output += `  FinalTracker preview:\n`;
        data.forEach((row, idx) => {
          output += `    Row ${idx + 1}: ${row.join(' | ')}\n`;
        });
      }
      
      if (name === 'CurrentBalance') {
        const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
        output += `  CurrentBalance full data:\n`;
        data.forEach((row, idx) => {
          output += `    Row ${idx + 1}: ${row.join(' | ')}\n`;
        });
      }
      
      rows.push([name, lastRow, lastCol, headersStr, sampleDataStr]);
      output += '\n';
    });
    
    // Write to Diagnostics sheet
    diagSheet.getRange(1, 1, rows.length, 5).setValues(rows);
    diagSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    diagSheet.autoResizeColumns(1, 5);
    
    // Also write full text output to column F
    const outputLines = output.split('\n');
    diagSheet.getRange(1, 6, outputLines.length, 1).setValues(outputLines.map(line => [line]));
    diagSheet.setColumnWidth(6, 200);
    
    // Log to console and show alert
    console.log(output);
    SpreadsheetApp.getUi().alert('Diagnostics written to "Diagnostics" sheet! Check column F for full output.');
  }
  
  /**
   * Export diagnostics to a downloadable format
   * This creates a text representation that can be easily copied
   */
  function exportDiagnostics() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const diagSheet = ss.getSheetByName('Diagnostics');
    
    if (!diagSheet) {
      SpreadsheetApp.getUi().alert('Please run inspectSheetStructure first!');
      return;
    }
    
    const fullOutput = diagSheet.getRange(1, 6, diagSheet.getLastRow(), 1).getValues()
      .map(row => row[0])
      .join('\n');
    
    // Copy to clipboard (via alert that user can copy)
    SpreadsheetApp.getUi().alert('Full diagnostics output:\n\n' + fullOutput.substring(0, 20000));
  }
