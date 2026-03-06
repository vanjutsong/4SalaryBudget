const CC_RECON_DEFAULT_DATE_WINDOW_DAYS = 2;
const CC_RECON_FX_VARIANCE_RATIO = 0.25;
const CC_RECON_FX_MIN_ABS = 50;
const CC_RECON_MANUAL_CANDIDATE_VARIANCE_RATIO = 0.35;
const CC_RECON_MANUAL_CANDIDATE_MIN_ABS = 80;

/**
 * Open unified credit card reconciliation dialog (ChatGPT CSV workflow).
 */
function reconcileCreditCardBillFromScreenshot() {
  const promptTemplate = getChatGptStatementToCsvPrompt();
  const htmlOutput = HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <style>
            body { font-family: Arial, sans-serif; padding: 16px; }
            h3 { margin: 0 0 12px 0; }
            label { display: block; font-weight: bold; margin-top: 10px; margin-bottom: 6px; }
            textarea { width: 100%; box-sizing: border-box; font-family: monospace; }
            #promptTemplate { height: 180px; }
            #csvInput { height: 240px; }
            #summary { height: 200px; }
            #csvOutput { height: 180px; }
            .row { display: flex; gap: 12px; align-items: center; margin: 8px 0; }
            .buttons { margin-top: 12px; text-align: right; }
            .muted { color: #666; font-size: 12px; }
            .panel { margin-top: 12px; padding: 10px; border: 1px solid #ddd; border-radius: 6px; background: #fafafa; }
            .warning { color: #b06000; white-space: pre-wrap; font-size: 12px; }
            .review-list { max-height: 320px; overflow: auto; border: 1px solid #ddd; padding: 8px; background: #fff; }
            .review-item { border-bottom: 1px dashed #ddd; padding: 8px 0; }
            .review-item:last-child { border-bottom: none; }
            .review-head { font-size: 12px; margin-bottom: 6px; }
            .review-actions { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
            .desc-input { min-width: 280px; }
            button { padding: 8px 14px; cursor: pointer; }
          </style>
        </head>
        <body>
          <h3>Credit Card Reconciliation</h3>
          <div class="muted">Use ChatGPT to convert screenshot to CSV, then paste CSV below for reconciliation.</div>

          <label for="promptTemplate">Prompt for ChatGPT (copy and use with screenshot)</label>
          <textarea id="promptTemplate" readonly></textarea>
          <div class="buttons" style="margin-top:6px;">
            <button onclick="copyPromptTemplate()">Copy Prompt</button>
          </div>

          <label for="csvInput">Paste CSV output from ChatGPT</label>
          <textarea id="csvInput" placeholder="Format per row: Date,Description,Mode,Category,Income,Debits,CCTransactionDate"></textarea>

          <div class="row">
            <label for="dateWindow" style="margin:0;">Date Window (days)</label>
            <input id="dateWindow" type="number" min="0" max="7" value="${CC_RECON_DEFAULT_DATE_WINDOW_DAYS}" />
            <span class="muted">Used for matching against CCTransactionDate</span>
          </div>

          <div class="buttons">
            <button onclick="cancelDialog()">Close</button>
            <button onclick="runReconciliation()" style="background:#4285f4; color:#fff; border:none;">Reconcile</button>
          </div>

          <div id="resultPanel" class="panel" style="display:none;">
            <label for="summary">Reconciliation Summary</label>
            <textarea id="summary" readonly></textarea>

            <label for="csvOutput">CSV for Statement-only Rows (schema-matching)</label>
            <textarea id="csvOutput" readonly></textarea>

            <div id="warnings" class="warning"></div>
          </div>

          <div id="resolvePanel" class="panel" style="display:none;">
            <label>Resolve Missing Statement Items</label>
            <div class="muted">Choose Add / Match existing / Skip per item. Add action enforces duplicate checks.</div>
            <div id="reviewList" class="review-list"></div>
            <div class="buttons">
              <button onclick="applyDecisions()" style="background:#0b8043; color:#fff; border:none;">Apply Decisions</button>
            </div>
          </div>

          <script>
            let lastReviewItems = [];
            const promptTemplate = ${JSON.stringify(promptTemplate)};

            document.getElementById('promptTemplate').value = promptTemplate;

            function cancelDialog() {
              google.script.host.close();
            }

            function copyPromptTemplate() {
              const text = promptTemplate || '';
              if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(text).then(function() {
                  alert('Prompt copied.');
                }).catch(function() {
                  fallbackCopyPrompt(text);
                });
                return;
              }
              fallbackCopyPrompt(text);
            }

            function fallbackCopyPrompt(text) {
              const box = document.getElementById('promptTemplate');
              box.value = text;
              box.focus();
              box.select();
              document.execCommand('copy');
              alert('Prompt copied.');
            }

            function runReconciliation() {
              const csvText = document.getElementById('csvInput').value.trim();
              const dateWindow = Number(document.getElementById('dateWindow').value || 0);
              if (!csvText) {
                alert('Please paste CSV output first.');
                return;
              }

              const handleResult = function(result) {
                if (!result) {
                  alert('No result returned.');
                  return;
                }
                if (result.error) {
                  alert(result.error);
                  return;
                }
                document.getElementById('resultPanel').style.display = 'block';
                document.getElementById('summary').value = result.summaryText || '';
                document.getElementById('csvOutput').value = result.statementOnlyCsv || '';
                document.getElementById('warnings').textContent = (result.warnings && result.warnings.length > 0)
                  ? ('Warnings:\\n' + result.warnings.join('\\n'))
                  : '';

                lastReviewItems = result.reviewItems || [];
                renderReviewItems(lastReviewItems);
              };

              const handleFailure = function(error) {
                alert('Error: ' + error.message);
              };

              google.script.run
                .withSuccessHandler(handleResult)
                .withFailureHandler(handleFailure)
                .processCreditCardReconciliationCsvInput(csvText, dateWindow);
            }

            function renderReviewItems(items) {
              const panel = document.getElementById('resolvePanel');
              const list = document.getElementById('reviewList');
              list.innerHTML = '';

              if (!items || items.length === 0) {
                panel.style.display = 'none';
                return;
              }

              panel.style.display = 'block';
              items.forEach(function(item, idx) {
                const wrap = document.createElement('div');
                wrap.className = 'review-item';

                const head = document.createElement('div');
                head.className = 'review-head';
                head.textContent = (idx + 1) + '. ' + item.dateText + ' | ' + item.description + ' | ' + item.amountText;
                wrap.appendChild(head);

                const actions = document.createElement('div');
                actions.className = 'review-actions';

                const select = document.createElement('select');
                select.id = 'action-' + item.id;
                const hasCandidates = item.candidates && item.candidates.length > 0;
                const defaultAction = hasCandidates ? 'match' : 'add';
                select.innerHTML =
                  '<option value="add"' + (defaultAction === 'add' ? ' selected' : '') + '>Add as new transaction</option>' +
                  '<option value="match"' + (defaultAction === 'match' ? ' selected' : '') + '>Match existing transaction</option>' +
                  '<option value="skip">Skip</option>';
                actions.appendChild(select);

                const candidateSelect = document.createElement('select');
                candidateSelect.id = 'candidate-' + item.id;
                const candidateOptions = (item.candidates || []).map(function(c) {
                  return '<option value="' + c.rowNumber + '">Row ' + c.rowNumber + ' | ' + c.dateText + ' | ' + c.amountText + ' | ' + c.description + '</option>';
                }).join('');
                candidateSelect.innerHTML = candidateOptions || '<option value="">No candidates</option>';
                candidateSelect.disabled = !hasCandidates;
                actions.appendChild(candidateSelect);

                const descInput = document.createElement('input');
                descInput.type = 'text';
                descInput.id = 'desc-' + item.id;
                descInput.className = 'desc-input';
                descInput.value = item.description || '';
                descInput.placeholder = 'Description to use if Add is selected';
                actions.appendChild(descInput);

                wrap.appendChild(actions);
                list.appendChild(wrap);
              });
            }

            function applyDecisions() {
              if (!lastReviewItems || lastReviewItems.length === 0) {
                alert('No missing statement items to resolve.');
                return;
              }

              const decisions = lastReviewItems.map(function(item) {
                const action = (document.getElementById('action-' + item.id) || {}).value || 'skip';
                const candidateRow = (document.getElementById('candidate-' + item.id) || {}).value || '';
                const customDescription = (document.getElementById('desc-' + item.id) || {}).value || '';
                return {
                  itemId: item.id,
                  action: action,
                  candidateRowNumber: candidateRow ? Number(candidateRow) : null,
                  customDescription: customDescription,
                  statementDateIso: item.dateIso,
                  statementAmount: item.amountValue,
                  statementDescription: item.description
                };
              });

              const dateWindow = Number(document.getElementById('dateWindow').value || 0);
              google.script.run
                .withSuccessHandler(function(result) {
                  if (result && result.error) {
                    alert(result.error);
                    return;
                  }
                  alert((result && result.message) || 'Decisions applied.');
                })
                .withFailureHandler(function(error) {
                  alert('Error applying decisions: ' + error.message);
                })
                .applyCreditCardReconciliationDecisions(decisions, dateWindow);
            }
          </script>
        </body>
      </html>
    `)
    .setWidth(900)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Reconcile Credit Card Bill');
}

function getChatGptStatementToCsvPrompt() {
  const lookupCategories = getLookupCategoriesForPrompt();
  const categoryListText = lookupCategories.length > 0
    ? lookupCategories.join(', ')
    : 'Other';

  return [
    'You are a transaction extraction assistant.',
    '',
    'I will upload a screenshot of my credit card statement. Extract transactions and return ONLY CSV rows (no explanation), matching this exact schema and column order:',
    '',
    'Date,Description,Mode,Category,Income,Debits,CCTransactionDate',
    '',
    'Rules:',
    '1) Each output row must have exactly 7 columns.',
    '2) Mode must always be: CreditCard',
    '3) Category must be selected from this allowed list (from my Lookups sheet):',
    `   ${categoryListText}`,
    '   - Choose the single best matching category per transaction.',
    '   - If uncertain, use "Other" (if present in the list), and append [CHECK] at end of Description.',
    '4) Income must be empty unless it is a refund/credit; normal purchases go to Debits.',
    '5) Debits must be positive numeric value with up to 2 decimals, no currency symbols.',
    '6) CCTransactionDate = transaction date from statement (purchase date).',
    '7) Date = due-date-based date computed from CCTransactionDate using this rule:',
    '   - statement cutoff is the 26th',
    '   - if CCTransactionDate day < 26: statement date = 26th of same month',
    '   - if CCTransactionDate day >= 26: statement date = 26th of next month',
    '   - due date = 15th of month after statement date month',
    '   - output Date as YYYY-MM-DD',
    '8) CCTransactionDate format must be YYYY-MM-DD.',
    '9) Description must use the merchant name exactly as it appears on the bill line (keep merchant tokens/symbols; do not rewrite brand names).',
    '10) If a line has both foreign amount and billed local amount, use billed local amount (rightmost final billed amount).',
    '11) If a transaction is ambiguous, still output best effort and add [CHECK] at end of Description.',
    '12) Do not include headers, markdown, code block, numbering, or extra text.',
    '13) Preserve one row per transaction.',
    '',
    'Now process the uploaded screenshot and return only CSV rows.'
  ].join('\\n');
}

function getLookupCategoriesForPrompt() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const lookupsSheet = ss.getSheetByName('Lookups');
    if (!lookupsSheet) return [];

    const lastRow = lookupsSheet.getLastRow();
    if (lastRow < 2) return [];

    const values = lookupsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const unique = [];
    const seen = {};

    values.forEach(row => {
      const raw = (row[0] || '').toString().trim();
      if (!raw) return;
      const key = raw.toLowerCase();
      if (seen[key]) return;
      seen[key] = true;
      unique.push(raw);
    });

    return unique;
  } catch (e) {
    return [];
  }
}

/**
 * Parse CSV text (from ChatGPT) and reconcile against CreditCard rows.
 */
function processCreditCardReconciliationCsvInput(csvText, dateWindowDays) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const variableSheet = ss.getSheetByName('VariableExpenses') || ss.getSheetByName('VariableExpences');
  if (!variableSheet) {
    return { error: 'VariableExpenses sheet not found.' };
  }

  if (!csvText || csvText.trim() === '') {
    return { error: 'No CSV text provided.' };
  }

  const resolvedDateWindow = Number.isFinite(Number(dateWindowDays))
    ? Math.max(0, Math.min(7, Math.floor(Number(dateWindowDays))))
    : CC_RECON_DEFAULT_DATE_WINDOW_DAYS;

  const parsed = parseStatementCsvTransactions(csvText);
  if (parsed.transactions.length === 0) {
    return {
      error: 'No parseable CSV rows found.\n\n' +
        (parsed.warnings.length > 0 ? parsed.warnings.slice(0, 20).join('\n') : 'Please check CSV formatting.')
    };
  }

  const ledgerTransactions = getLedgerCreditCardTransactions(variableSheet);
  const reconciliation = reconcileCreditCardStatementTransactions(parsed.transactions, ledgerTransactions, resolvedDateWindow);
  const summaryText = buildCreditCardReconciliationSummary(parsed, reconciliation, resolvedDateWindow);
  const statementOnlyCsv = buildSchemaCsvForStatementRows(reconciliation.statementOnly);
  const reviewItems = buildManualReviewItems(reconciliation.statementOnly, ledgerTransactions, resolvedDateWindow);

  return {
    success: true,
    summaryText,
    statementOnlyCsv,
    reviewItems,
    warnings: parsed.warnings
  };
}

// Backward compatibility wrapper
function processCreditCardReconciliationInput(ocrText, dateWindowDays) {
  return processCreditCardReconciliationCsvInput(ocrText, dateWindowDays);
}

function parseStatementCsvTransactions(csvText) {
  const normalized = (csvText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = normalized.split('\n').map(l => l.trim()).filter(Boolean);
  const transactions = [];
  const warnings = [];

  lines.forEach((line, idx) => {
    const lineNo = idx + 1;
    const fields = parseCsvLineFields(line);
    if (fields.length === 0) return;

    if (lineNo === 1 && fields.length >= 7 && String(fields[0]).toLowerCase() === 'date' && String(fields[6]).toLowerCase().includes('cctransactiondate')) {
      return; // skip header if present
    }

    if (fields.length !== 7) {
      warnings.push(`Line ${lineNo}: Expected 7 CSV fields, got ${fields.length}`);
      return;
    }

    const description = (fields[1] || '').toString().trim();
    const income = Number(String(fields[4] || '').replace(/,/g, '').trim() || 0);
    const debits = Number(String(fields[5] || '').replace(/,/g, '').trim() || 0);
    const ccDate = toSafeDate(fields[6] || fields[0]);

    if (!description) {
      warnings.push(`Line ${lineNo}: Missing Description`);
      return;
    }
    if (!ccDate) {
      warnings.push(`Line ${lineNo}: Invalid CCTransactionDate`);
      return;
    }
    if ((isNaN(income) ? 0 : income) <= 0 && (isNaN(debits) ? 0 : debits) <= 0) {
      warnings.push(`Line ${lineNo}: Income/Debits missing or invalid`);
      return;
    }

    const amountSigned = debits > 0 ? Math.abs(debits) : -Math.abs(income);
    transactions.push({
      lineNumber: lineNo,
      date: ccDate,
      description,
      amountSigned
    });
  });

  return { transactions, warnings };
}

function parseCsvLineFields(line) {
  const fields = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && i + 1 < line.length && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === ',' && !inQuotes) {
      fields.push(current.trim());
      current = '';
    } else {
      current += ch;
    }
  }
  fields.push(current.trim());
  return fields;
}

function applyCreditCardReconciliationDecisions(decisions, dateWindowDays) {
  if (!decisions || !Array.isArray(decisions) || decisions.length === 0) {
    return { error: 'No decisions provided.' };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const variableSheet = ss.getSheetByName('VariableExpenses') || ss.getSheetByName('VariableExpences');
  if (!variableSheet) {
    return { error: 'VariableExpenses sheet not found.' };
  }

  const resolvedDateWindow = Number.isFinite(Number(dateWindowDays))
    ? Math.max(0, Math.min(7, Math.floor(Number(dateWindowDays))))
    : CC_RECON_DEFAULT_DATE_WINDOW_DAYS;

  const ledgerTransactions = getLedgerCreditCardTransactions(variableSheet);
  const rowsToAppend = [];
  let addedCount = 0;
  let matchedCount = 0;
  let skippedCount = 0;
  let blockedDuplicateCount = 0;

  decisions.forEach(decision => {
    const action = ((decision && decision.action) || 'skip').toString().toLowerCase();
    if (action === 'skip') {
      skippedCount += 1;
      return;
    }

    if (action === 'match') {
      const candidateRow = Number(decision.candidateRowNumber) || 0;
      if (candidateRow > 0) {
        // Manual confirm: treat as reconciled without adding any row.
        matchedCount += 1;
      } else {
        skippedCount += 1;
      }
      return;
    }

    if (action !== 'add') {
      skippedCount += 1;
      return;
    }

    const ccTxnDate = toSafeDate(decision.statementDateIso);
    const amountValue = Math.abs(Number(decision.statementAmount) || 0);
    const finalDescription = ((decision.customDescription || decision.statementDescription || '').toString().trim());
    if (!ccTxnDate || amountValue <= 0 || !finalDescription) {
      skippedCount += 1;
      return;
    }

    if (hasDuplicateCreditCardByDateAmount(ledgerTransactions, ccTxnDate, amountValue, resolvedDateWindow)) {
      blockedDuplicateCount += 1;
      return;
    }

    const dueDate = computeCcDueDate(computeCcStatementDate(ccTxnDate));
    rowsToAppend.push([
      dueDate,
      finalDescription,
      'CreditCard',
      'Other',
      '',
      amountValue,
      ccTxnDate
    ]);

    // Keep in-memory ledger up to date so duplicates in the same batch are blocked too.
    ledgerTransactions.push({
      rowNumber: -1,
      date: ccTxnDate,
      description: finalDescription,
      amountSigned: amountValue
    });
    addedCount += 1;
  });

  if (rowsToAppend.length > 0) {
    const lastRow = variableSheet.getLastRow();
    const insertRow = lastRow > 1 ? lastRow + 1 : 2;
    variableSheet.getRange(insertRow, 1, rowsToAppend.length, 7).setValues(rowsToAppend);
  }

  const message =
    'Reconciliation decisions applied.\n\n' +
    `Added new transactions: ${addedCount}\n` +
    `Matched existing (no add): ${matchedCount}\n` +
    `Skipped: ${skippedCount}\n` +
    `Blocked as duplicates: ${blockedDuplicateCount}\n\n` +
    'Tip: run "Update FinalTracker" to include newly added transactions.';

  return { success: true, message };
}

function hasDuplicateCreditCardByDateAmount(ledgerTransactions, ccTxnDate, amountValue, dateWindowDays) {
  if (!ledgerTransactions || ledgerTransactions.length === 0) return false;
  const targetAmount = Math.abs(Number(amountValue) || 0);
  if (targetAmount <= 0) return false;

  for (let i = 0; i < ledgerTransactions.length; i++) {
    const txn = ledgerTransactions[i];
    const txnAmount = Math.abs(Number(txn.amountSigned) || 0);
    if (Math.abs(txnAmount - targetAmount) > 0.01) continue;

    const txnDate = toSafeDate(txn.date);
    if (!txnDate || !ccTxnDate) continue;
    const dateDiffDays = Math.abs(txnDate.getTime() - ccTxnDate.getTime()) / (1000 * 60 * 60 * 24);
    if (dateDiffDays <= dateWindowDays) return true;
  }
  return false;
}

function buildManualReviewItems(statementOnly, ledgerTransactions, dateWindowDays) {
  if (!statementOnly || statementOnly.length === 0) return [];
  const tz = Session.getScriptTimeZone();
  return statementOnly.map((stmt, idx) => {
    const stmtDate = toSafeDate(stmt.date);
    const dateIso = stmtDate ? Utilities.formatDate(stmtDate, tz, 'yyyy-MM-dd') : '';
    const amountValue = Math.abs(Number(stmt.amountSigned) || 0);
    const candidates = findManualMatchCandidates(stmt, ledgerTransactions, dateWindowDays);
    return {
      id: `s${idx + 1}_${dateIso}_${Math.round(amountValue * 100)}`,
      dateIso,
      dateText: stmtDate ? Utilities.formatDate(stmtDate, tz, 'MM/dd/yyyy') : 'N/A',
      description: stmt.description || '',
      amountValue,
      amountText: amountValue.toFixed(2),
      candidates
    };
  });
}

function findManualMatchCandidates(statementTxn, ledgerTransactions, dateWindowDays) {
  const candidates = [];
  const stmtAmount = Math.abs(Number(statementTxn.amountSigned) || 0);
  if (!stmtAmount || !ledgerTransactions || ledgerTransactions.length === 0) return candidates;

  const maxVariance = Math.max(CC_RECON_MANUAL_CANDIDATE_MIN_ABS, stmtAmount * CC_RECON_MANUAL_CANDIDATE_VARIANCE_RATIO);
  const tz = Session.getScriptTimeZone();

  ledgerTransactions.forEach(ledger => {
    const ledgerAmount = Math.abs(Number(ledger.amountSigned) || 0);
    if (!ledgerAmount) return;

    const amountDiff = Math.abs(ledgerAmount - stmtAmount);
    if (amountDiff > maxVariance) return;

    const stmtDate = toSafeDate(statementTxn.date);
    const ledgerDate = toSafeDate(ledger.date);
    if (!stmtDate || !ledgerDate) return;

    const dateDiffDays = Math.abs(ledgerDate.getTime() - stmtDate.getTime()) / (1000 * 60 * 60 * 24);
    if (dateDiffDays > Math.max(dateWindowDays, 5)) return;

    const descriptionScore = scoreDescriptionSimilarity(statementTxn.description, ledger.description);
    if (descriptionScore < 0.2 && amountDiff > 0.01) return;

    candidates.push({
      rowNumber: ledger.rowNumber,
      dateIso: Utilities.formatDate(ledgerDate, tz, 'yyyy-MM-dd'),
      dateText: Utilities.formatDate(ledgerDate, tz, 'MM/dd/yyyy'),
      description: ledger.description || '',
      amountValue: ledgerAmount,
      amountText: ledgerAmount.toFixed(2),
      amountDiff,
      dateDiffDays,
      descriptionScore
    });
  });

  candidates.sort((a, b) => {
    if (a.amountDiff !== b.amountDiff) return a.amountDiff - b.amountDiff;
    if (a.dateDiffDays !== b.dateDiffDays) return a.dateDiffDays - b.dateDiffDays;
    return b.descriptionScore - a.descriptionScore;
  });

  return candidates.slice(0, 3);
}

function getLedgerCreditCardTransactions(variableSheet) {
  const lastRow = variableSheet.getLastRow();
  if (lastRow <= 1) return [];

  const rows = variableSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const txns = [];

  rows.forEach((row, idx) => {
    const mode = (row[2] || '').toString().trim().toLowerCase().replace(/\s+/g, '');
    if (mode !== 'creditcard') return;

    const ccDate = toSafeDate(row[6]);
    const dueDate = toSafeDate(row[0]);
    const effectiveDate = ccDate || dueDate;
    if (!effectiveDate) return;

    const description = (row[1] || '').toString().trim();
    const income = Number(row[4]) || 0;
    const debits = Number(row[5]) || 0;
    const amountSigned = debits > 0 ? debits : (income > 0 ? -income : 0);
    if (amountSigned === 0) return;

    txns.push({
      rowNumber: idx + 2,
      date: effectiveDate,
      description,
      amountSigned
    });
  });

  return txns;
}

function reconcileCreditCardStatementTransactions(statementTransactions, ledgerTransactions, dateWindowDays) {
  const usedLedgerIndices = {};
  const matched = [];
  const fxAdjustedMatches = [];
  const statementOnly = [];

  const sortedStatement = statementTransactions.slice().sort((a, b) => a.date.getTime() - b.date.getTime());

  sortedStatement.forEach(stmt => {
    let bestIndex = -1;
    let bestScore = null;

    for (let i = 0; i < ledgerTransactions.length; i++) {
      if (usedLedgerIndices[i]) continue;

      const ledger = ledgerTransactions[i];
      const amountDiff = Math.abs((ledger.amountSigned || 0) - (stmt.amountSigned || 0));
      if (amountDiff > 0.01) continue;

      const dateDiffDays = Math.abs(ledger.date.getTime() - stmt.date.getTime()) / (1000 * 60 * 60 * 24);
      if (dateDiffDays > dateWindowDays) continue;

      const descriptionScore = scoreDescriptionSimilarity(stmt.description, ledger.description);
      const score = {
        dateDiffDays,
        descriptionScore,
        rowNumber: ledger.rowNumber
      };

      const isBetter = !bestScore ||
        score.dateDiffDays < bestScore.dateDiffDays ||
        (score.dateDiffDays === bestScore.dateDiffDays && score.descriptionScore > bestScore.descriptionScore) ||
        (score.dateDiffDays === bestScore.dateDiffDays &&
          score.descriptionScore === bestScore.descriptionScore &&
          score.rowNumber < bestScore.rowNumber);

      if (isBetter) {
        bestIndex = i;
        bestScore = score;
      }
    }

    if (bestIndex >= 0) {
      usedLedgerIndices[bestIndex] = true;
      matched.push({
        statement: stmt,
        ledger: ledgerTransactions[bestIndex],
        dateDiffDays: bestScore.dateDiffDays
      });
      return;
    }

    const fxCandidate = findFxAdjustedCandidate(stmt, ledgerTransactions, usedLedgerIndices, dateWindowDays);
    if (fxCandidate) {
      usedLedgerIndices[fxCandidate.index] = true;
      fxAdjustedMatches.push({
        statement: stmt,
        ledger: ledgerTransactions[fxCandidate.index],
        dateDiffDays: fxCandidate.dateDiffDays,
        amountDiff: fxCandidate.amountDiff
      });
      return;
    }

    statementOnly.push(stmt);
  });

  let periodStart = null;
  let periodEnd = null;
  if (statementTransactions.length > 0) {
    periodStart = new Date(Math.min.apply(null, statementTransactions.map(t => t.date.getTime())));
    periodEnd = new Date(Math.max.apply(null, statementTransactions.map(t => t.date.getTime())));
  }

  const ledgerOnly = [];
  for (let i = 0; i < ledgerTransactions.length; i++) {
    if (usedLedgerIndices[i]) continue;
    const txn = ledgerTransactions[i];
    if (!periodStart || !periodEnd) continue;
    if (txn.date >= periodStart && txn.date <= periodEnd) {
      ledgerOnly.push(txn);
    }
  }

  return { matched, fxAdjustedMatches, statementOnly, ledgerOnly, periodStart, periodEnd };
}

function findFxAdjustedCandidate(statementTxn, ledgerTransactions, usedLedgerIndices, dateWindowDays) {
  const stmtAmount = Math.abs(Number(statementTxn.amountSigned) || 0);
  if (stmtAmount <= 0) return null;

  const maxVariance = Math.max(CC_RECON_FX_MIN_ABS, stmtAmount * CC_RECON_FX_VARIANCE_RATIO);
  let best = null;

  for (let i = 0; i < ledgerTransactions.length; i++) {
    if (usedLedgerIndices[i]) continue;

    const ledger = ledgerTransactions[i];
    const ledgerAmount = Math.abs(Number(ledger.amountSigned) || 0);
    if (ledgerAmount <= 0) continue;

    const amountDiff = Math.abs(ledgerAmount - stmtAmount);
    if (amountDiff <= 0.01 || amountDiff > maxVariance) continue;

    const dateDiffDays = Math.abs(ledger.date.getTime() - statementTxn.date.getTime()) / (1000 * 60 * 60 * 24);
    if (dateDiffDays > dateWindowDays) continue;

    const descriptionScore = scoreDescriptionSimilarity(statementTxn.description, ledger.description);
    if (descriptionScore < 0.45) continue;

    const score = { index: i, amountDiff, dateDiffDays, descriptionScore };
    const isBetter = !best ||
      score.descriptionScore > best.descriptionScore ||
      (score.descriptionScore === best.descriptionScore && score.amountDiff < best.amountDiff) ||
      (score.descriptionScore === best.descriptionScore && score.amountDiff === best.amountDiff && score.dateDiffDays < best.dateDiffDays);

    if (isBetter) best = score;
  }

  return best;
}

function scoreDescriptionSimilarity(a, b) {
  const left = (a || '').toLowerCase().replace(/[^a-z0-9 ]/g, ' ').split(/\s+/).filter(w => w.length >= 3);
  const right = (b || '').toLowerCase().replace(/[^a-z0-9 ]/g, ' ').split(/\s+/).filter(w => w.length >= 3);
  if (left.length === 0 || right.length === 0) return 0;

  const rightLookup = {};
  right.forEach(word => { rightLookup[word] = true; });

  let overlap = 0;
  left.forEach(word => {
    if (rightLookup[word]) overlap += 1;
  });

  return overlap / left.length;
}

function buildCreditCardReconciliationSummary(parsed, reconciliation, dateWindowDays) {
  const tz = Session.getScriptTimeZone();
  const formatDate = d => Utilities.formatDate(d, tz, 'MM/dd/yyyy');
  const amountText = n => Math.abs(Number(n) || 0).toFixed(2);

  let summary = 'Credit Card Reconciliation Result\n\n';
  summary += `Parsed statement rows: ${parsed.transactions.length}\n`;
  summary += `Matched (exact amount): ${reconciliation.matched.length}\n`;
  summary += `Matched (FX-adjusted recurring): ${reconciliation.fxAdjustedMatches.length}\n`;
  summary += `Statement-only (missing in ledger): ${reconciliation.statementOnly.length}\n`;
  summary += `Ledger-only (within statement period): ${reconciliation.ledgerOnly.length}\n`;
  summary += `Date window used: +/- ${dateWindowDays} day(s)\n`;
  summary += `FX variance allowance: max(${CC_RECON_FX_MIN_ABS.toFixed(2)}, ${(CC_RECON_FX_VARIANCE_RATIO * 100).toFixed(0)}% of amount)\n`;

  if (reconciliation.periodStart && reconciliation.periodEnd) {
    summary += `Statement period from CSV rows: ${formatDate(reconciliation.periodStart)} to ${formatDate(reconciliation.periodEnd)}\n`;
  }

  if (reconciliation.statementOnly.length > 0) {
    summary += '\nStatement-only details (top 15):\n';
    reconciliation.statementOnly.slice(0, 15).forEach((txn, idx) => {
      summary += `${idx + 1}. ${formatDate(txn.date)} | ${txn.description} | ${amountText(txn.amountSigned)}\n`;
    });
    if (reconciliation.statementOnly.length > 15) {
      summary += `... and ${reconciliation.statementOnly.length - 15} more\n`;
    }
  }

  if (reconciliation.fxAdjustedMatches.length > 0) {
    summary += '\nFX-adjusted recurring matches (top 15):\n';
    reconciliation.fxAdjustedMatches.slice(0, 15).forEach((match, idx) => {
      summary += `${idx + 1}. ${formatDate(match.statement.date)} | ${match.statement.description} | stmt ${amountText(match.statement.amountSigned)} vs ledger ${amountText(match.ledger.amountSigned)} (diff ${amountText(match.amountDiff)})\n`;
    });
    if (reconciliation.fxAdjustedMatches.length > 15) {
      summary += `... and ${reconciliation.fxAdjustedMatches.length - 15} more\n`;
    }
  }

  if (reconciliation.ledgerOnly.length > 0) {
    summary += '\nLedger-only details (top 15):\n';
    reconciliation.ledgerOnly.slice(0, 15).forEach((txn, idx) => {
      summary += `${idx + 1}. Row ${txn.rowNumber} | ${formatDate(txn.date)} | ${txn.description} | ${amountText(txn.amountSigned)}\n`;
    });
    if (reconciliation.ledgerOnly.length > 15) {
      summary += `... and ${reconciliation.ledgerOnly.length - 15} more\n`;
    }
  }

  return summary;
}

function buildSchemaCsvForStatementRows(statementRows) {
  if (!statementRows || statementRows.length === 0) return '';

  const tz = Session.getScriptTimeZone();
  const lines = [];

  statementRows.forEach(txn => {
    const ccTxnDate = toSafeDate(txn.date);
    if (!ccTxnDate) return;

    const ccTxnDateText = Utilities.formatDate(ccTxnDate, tz, 'yyyy-MM-dd');
    const dueDate = computeCcDueDate(computeCcStatementDate(ccTxnDate));
    const dueDateText = Utilities.formatDate(dueDate, tz, 'yyyy-MM-dd');

    const amount = Number(txn.amountSigned) || 0;
    const income = amount < 0 ? Math.abs(amount).toFixed(2) : '';
    const debits = amount >= 0 ? Math.abs(amount).toFixed(2) : '';

    const fields = [
      dueDateText,
      txn.description || '',
      'CreditCard',
      'Other',
      income,
      debits,
      ccTxnDateText
    ].map(escapeCsvField);

    lines.push(fields.join(','));
  });

  return lines.join('\n');
}

function escapeCsvField(value) {
  const text = (value === null || value === undefined) ? '' : String(value);
  if (text.includes('"') || text.includes(',') || text.includes('\n')) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}
