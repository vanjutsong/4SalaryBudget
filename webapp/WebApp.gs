// Web App Backend for Mobile Budget Tracker
// Deploy this as a web app in Google Apps Script

const SPREADSHEET_ID = '1FQzuRQwlFrGGu10N8-ne3HYfbl9tbIoaWA2bqlE6bKo';

// Main budget script web app URL. Deploy the root project (4salarybudget) as a web app and set this
// so that opening the dashboard triggers updateFinalTracker and refreshes balances.
const MAIN_SCRIPT_WEB_APP_URL = ''; // e.g. 'https://script.google.com/macros/s/YOUR_MAIN_SCRIPT_DEPLOY_ID/exec'

/**
 * Handle GET requests (for dashboard)
 */
function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'getBalances') {
    return ContentService.createTextOutput(JSON.stringify(getBalances()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Return HTML for dashboard or form
  const page = e.parameter.page || 'dashboard';
  if (page === 'form') {
    return HtmlService.createTemplateFromFile('TransactionForm')
      .evaluate()
      .setTitle('Add Transaction')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // Default: dashboard — trigger updateFinalTracker when user opens dashboard (e.g. back from form)
  if (MAIN_SCRIPT_WEB_APP_URL) {
    try {
      UrlFetchApp.fetch(MAIN_SCRIPT_WEB_APP_URL + '?action=updateFinalTracker', { muteHttpExceptions: true });
    } catch (err) {
      // Continue to serve dashboard even if refresh fails
    }
  }
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('Budget Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handle POST requests (for saving transactions)
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = saveTransaction(data);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Save a transaction to VariableExpenses sheet
 */
function saveTransaction(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const variableSheet = ss.getSheetByName('VariableExpenses') || ss.getSheetByName('VariableExpences');
  
  if (!variableSheet) {
    throw new Error('VariableExpenses sheet not found');
  }
  
  // Format date
  const date = new Date(data.date);
  const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  // Prepare row data: Date, Description, Mode, Category, Income, Debits, CCTransactionDate
  let rowData;
  
  if (data.type === 'Income') {
    rowData = [
      formattedDate,
      data.description,
      data.mode,
      data.category,
      data.amount,  // Income
      '',           // Debits (empty)
      data.ccTransactionDate ? Utilities.formatDate(new Date(data.ccTransactionDate), Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''
    ];
  } else {
    rowData = [
      formattedDate,
      data.description,
      data.mode,
      data.category,
      '',           // Income (empty)
      data.amount,  // Debits
      data.ccTransactionDate ? Utilities.formatDate(new Date(data.ccTransactionDate), Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''
    ];
  }
  
  // Handle credit card date normalization if needed
  if (data.mode === 'CreditCard' && data.ccTransactionDate) {
    // Use the credit card date calculation functions from your main script
    const ccTxnDate = new Date(data.ccTransactionDate);
    const statementDate = computeCcStatementDate(ccTxnDate);
    const dueDate = computeCcDueDate(statementDate);
    rowData[0] = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  
  // Append to sheet
  variableSheet.appendRow(rowData);
  
  return {
    success: true,
    message: 'Transaction saved successfully'
  };
}

/**
 * Get balances for dashboard
 * Reads directly from Dashboard sheet cells:
 * - Today's budget: B19 (numeric)
 * - Tomorrow budget: B20 (numeric)
 * - Day After: B21 (numeric)
 * - Current Week: B27 (numeric)
 * - Next Week: B28 (numeric)
 * Labels with day names are in column A
 */
function getBalances() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dashboardSheet = ss.getSheetByName('Dashboard');
  
  if (!dashboardSheet) {
    return { error: 'Dashboard sheet not found' };
  }
  
  // Read values directly from the Dashboard sheet
  // Column A has labels with day names, Column B has numeric values
  const todayLabel = dashboardSheet.getRange('A19').getValue();
  const todayBudget = dashboardSheet.getRange('B19').getValue();
  const tomorrowLabel = dashboardSheet.getRange('A20').getValue();
  const tomorrowBudget = dashboardSheet.getRange('B20').getValue();
  const dayAfterLabel = dashboardSheet.getRange('A21').getValue();
  const dayAfterBudget = dashboardSheet.getRange('B21').getValue();
  const currentWeekLabel = dashboardSheet.getRange('A27').getValue();
  const currentWeekBudget = dashboardSheet.getRange('B27').getValue();
  const nextWeekLabel = dashboardSheet.getRange('A28').getValue();
  const nextWeekBudget = dashboardSheet.getRange('B28').getValue();
  
  return {
    today: Number(todayBudget) || 0,
    todayLabel: todayLabel ? todayLabel.toString() : 'Today Budget',
    tomorrow: Number(tomorrowBudget) || 0,
    tomorrowLabel: tomorrowLabel ? tomorrowLabel.toString() : 'Tomorrow Budget',
    dayAfter: Number(dayAfterBudget) || 0,
    dayAfterLabel: dayAfterLabel ? dayAfterLabel.toString() : 'Day After Budget',
    thisWeek: Number(currentWeekBudget) || 0,
    thisWeekLabel: currentWeekLabel ? currentWeekLabel.toString() : 'Current Week',
    nextWeek: Number(nextWeekBudget) || 0,
    nextWeekLabel: nextWeekLabel ? nextWeekLabel.toString() : 'Next Week'
  };
}

/**
 * Credit card date calculation functions (from your main script)
 */
function computeCcStatementDate(txnDate) {
  const d = new Date(txnDate);
  const year = d.getFullYear();
  const monthIndex = d.getMonth();
  const day = d.getDate();

  if (day < 26) {
    return new Date(year, monthIndex, 26, 12, 0, 0, 0);
  }

  return new Date(year, monthIndex + 1, 26, 12, 0, 0, 0);
}

function computeCcDueDate(statementDate) {
  const s = new Date(statementDate);
  const year = s.getFullYear();
  const monthIndex = s.getMonth();
  return new Date(year, monthIndex + 1, 15, 12, 0, 0, 0);
}

/**
 * Get categories from Lookups sheet (column A only; does not touch Highlight section E–H)
 */
function getCategories() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const lookupsSheet = ss.getSheetByName('Lookups');

  if (!lookupsSheet) {
    return [];
  }

  const lastRow = lookupsSheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const categories = lookupsSheet.getRange(2, 1, lastRow, 1).getValues()
    .map(row => row[0])
    .filter(cat => cat && cat.toString().trim() !== '');

  return [...new Set(categories)];
}
