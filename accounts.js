// ================================
// ⚙️ TEST/MOCK CONFIGURATION
// ================================
const TEST_MODE_ENABLED = true; // Set to 'true' to test on single sheet, 'false' for production
const TEST_ACCOUNT_NAME = "VN - Tran Van Giang"; // Specify which account to test with

// ===============================
// 📒 ACCOUNT MANAGER SCRIPT
// ===============================

const updateAllAccounts = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("VN - Master Ledger");
  const data = masterSheet.getDataRange().getValues();
  const richData = masterSheet.getDataRange().getRichTextValues();
  const headers = data[0];

  const accountCol = headers.indexOf("Account");
  if (accountCol === -1) throw new Error("No 'Account' column found in VN - Master Ledger");

  // --- Collect all rows by account ---
  const accountMap = {};
  data.slice(1).forEach((row, i) => {
    const account = row[accountCol];
    if (!account) return;
    if (!accountMap[account]) accountMap[account] = [];
    accountMap[account].push({ row, rich: richData[i + 1] });
  });

  // --- Copy header formatting and column widths ---
  const headerRange = masterSheet.getRange(1, 1, 1, headers.length);
  const columnWidths = Array.from({ length: headers.length }, (_, i) =>
    masterSheet.getColumnWidth(i + 1)
  );

  // --- Loop through accounts (with test mode support) ---
  const accountsToProcess = TEST_MODE_ENABLED 
    ? Object.entries(accountMap).filter(([account]) => account === TEST_ACCOUNT_NAME)
    : Object.entries(accountMap);
  
  if (TEST_MODE_ENABLED && accountsToProcess.length === 0) {
    SpreadsheetApp.getUi().alert(`⚠️ Test Mode: Account '${TEST_ACCOUNT_NAME}' not found in master data.`);
    return;
  }
  
  accountsToProcess.forEach(([account, entries]) => {
    const cleanName = account.replace(/[\\\/\?\*\[\]]/g, " ");
    let targetSheet = ss.getSheetByName(cleanName);
    if (!targetSheet) targetSheet = ss.insertSheet(cleanName);
    targetSheet.clear();

    // --- Header row ---
    targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRangeTarget = targetSheet.getRange(1, 1, 1, headers.length);
    headerRangeTarget
      .setFontSize(12)
      .setFontWeight("bold")
      .setFontFamily("Arial")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setBackground("#0b5394") // deep blue
      .setFontColor("#ffffff");

    // --- Match column widths ---
    columnWidths.forEach((width, i) => targetSheet.setColumnWidth(i + 1, width));

    // --- Write data (sorted by date) ---
    const rows = entries.map(e => e.row);
    const richRows = entries.map(e => e.rich);
    
    // Sort by date column (column B - dd/mm/YY)
    const dateCol = 1; // Column B is index 1
    const sortedIndices = rows.map((row, index) => ({ row, rich: richRows[index], originalIndex: index }))
      .sort((a, b) => {
        const dateA = new Date(a.row[dateCol]);
        const dateB = new Date(b.row[dateCol]);
        return dateA - dateB; // Oldest to newest
      });
    
    // Write sorted data
    const sortedRows = sortedIndices.map(item => item.row);
    const sortedRich = sortedIndices.map(item => item.rich);
    
    targetSheet.getRange(2, 1, sortedRows.length, headers.length).setValues(sortedRows);

    // --- Center + style data ---
    targetSheet.getRange(2, 1, sortedRows.length, headers.length)
      .setFontSize(11)
      .setFontFamily("Arial")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // --- Reapply hyperlinks (with sorted data) ---
    sortedRich.forEach((richRow, rIdx) => {
      richRow.forEach((rich, cIdx) => {
        if (rich?.getLinkUrl()) {
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(rich.getText())
            .setLinkUrl(rich.getLinkUrl())
            .build();
          targetSheet.getRange(rIdx + 2, cIdx + 1).setRichTextValue(richText);
        }
      });
    });

    // --- Format date column (B) ---
    targetSheet.getRange(2, 2, sortedRows.length, 1).setNumberFormat("dd/mm/yy");

    // --- Compute totals (Debit = I, Credit = J) ---
    const debitCol = 9;
    const creditCol = 10;
    const totalDebit = sortedRows.reduce((sum, r) => sum + cleanCurrency(r[debitCol - 1]), 0);
    const totalCredit = sortedRows.reduce((sum, r) => sum + cleanCurrency(r[creditCol - 1]), 0);

    const totalRow = sortedRows.length + 2;
    
    // --- Add yellow separator row ---
    targetSheet.getRange(totalRow, 1, 1, headers.length)
      .setBackground("#FFF9C4") // Light yellow background
      .setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
    
    const totalsStartRow = totalRow + 1;
    
    // --- Create totals section exactly like your example ---
    // Labels row (dark blue background, white text) - ALIGNED UNDER CORRECT COLUMNS
    targetSheet.getRange(totalsStartRow, debitCol).setValue("Total Debit")
      .setBackground("#0b5394")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    
    targetSheet.getRange(totalsStartRow, creditCol).setValue("Total Credit")
      .setBackground("#0b5394")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    
    targetSheet.getRange(totalsStartRow, creditCol + 1).setValue("Remaining funds")
      .setBackground("#0b5394")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // Values row with CORRECT COLORS - Green for Debit, Orange for Credit, Yellow for Remaining
    const valuesRow = totalsStartRow + 1;
    targetSheet.getRange(valuesRow, debitCol).setValue(totalDebit)
      .setBackground("#d9ead3") // Light green for debit
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setNumberFormat("#,##0");
    
    targetSheet.getRange(valuesRow, creditCol).setValue(totalCredit)
      .setBackground("#fce5cd") // Light orange for credit
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setNumberFormat("#,##0");
    
    const remainingFunds = totalCredit - totalDebit;
    targetSheet.getRange(valuesRow, creditCol + 1).setValue(remainingFunds)
      .setBackground("#FFF9C4") // Light yellow for remaining
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setNumberFormat("#,##0");

    // --- Hide column N (if exists) ---
    const lastCol = targetSheet.getLastColumn();
    if (lastCol >= 14) targetSheet.hideColumns(14);

    // --- Set row height (21 px) ---
    Array.from({ length: valuesRow }, (_, i) => i + 1).forEach(r =>
      targetSheet.setRowHeight(r, 21)
    );

    // --- Format Debit/Credit columns ---
    targetSheet.getRange(2, debitCol, sortedRows.length, 1).setBackground("#d9ead3"); // green tint
    targetSheet.getRange(2, creditCol, sortedRows.length, 1).setBackground("#fce5cd"); // orange tint

    // --- Format Bill/Red Bill/Doc columns (yellow) ---
    const billCols = ["O", "P", "Q"].map(c => c.charCodeAt(0) - 64);
    billCols.forEach(col => {
      if (col <= headers.length)
        targetSheet.getRange(2, col, sortedRows.length, 1).setBackground("#fff2cc");
    });

    // --- Transaction number column (blue tint) ---
    targetSheet.getRange(2, 1, sortedRows.length, 1).setBackground("#c9daf8");

    // --- Remove gridlines, add outer borders ---
    targetSheet.setFrozenRows(1);
    targetSheet.setHiddenGridlines(true);
    const totalRowIndex = targetSheet.getLastRow();
    const totalColIndex = targetSheet.getLastColumn();
    targetSheet
      .getRange(1, 1, totalRowIndex, totalColIndex)
      .setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
  });

  const successMessage = TEST_MODE_ENABLED 
    ? `🧪 Test Mode: Account '${TEST_ACCOUNT_NAME}' updated successfully!`
    : "✅ All account sheets updated successfully!";
  SpreadsheetApp.getUi().alert(successMessage);
};

// ===============================
// ⚡ QUICK UPDATE — Refresh existing account sheets only
// ===============================
const quickUpdateAccounts = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("VN - Master Ledger");
  if (!masterSheet) throw new Error("❌ 'VN - Master Ledger' sheet not found.");

  const data = masterSheet.getDataRange().getValues();
  const headers = data[0];
  const accountCol = headers.indexOf("Account");
  const debitCol = headers.indexOf("Debit (VND)");
  const creditCol = headers.indexOf("Credit (VND)");
  if (accountCol === -1) throw new Error("❌ No 'Account' column found.");

  const accountMap = {};
  data.slice(1).forEach(row => {
    const account = row[accountCol];
    if (!account) return;
    if (!accountMap[account]) accountMap[account] = [];
    accountMap[account].push(row);
  });

  const accountsToProcess = TEST_MODE_ENABLED 
    ? Object.keys(accountMap).filter(account => account === TEST_ACCOUNT_NAME)
    : Object.keys(accountMap);
  
  if (TEST_MODE_ENABLED && accountsToProcess.length === 0) {
    SpreadsheetApp.getUi().alert(`⚠️ Test Mode: Account '${TEST_ACCOUNT_NAME}' not found in master data.`);
    return;
  }
  
  accountsToProcess.forEach(account => {
    const sheet = ss.getSheetByName(account);
    if (sheet) {
      const rows = accountMap[account];
      
      // Sort by date column (column B - dd/mm/YY)
      const dateCol = 1; // Column B is index 1
      rows.sort((a, b) => {
        const dateA = new Date(a[dateCol]);
        const dateB = new Date(b[dateCol]);
        return dateA - dateB; // Oldest to newest
      });
      
      const lastCol = sheet.getLastColumn();
      sheet.getRange(2, 1, sheet.getLastRow(), lastCol).clearContent();
      sheet.getRange(2, 1, rows.length, lastCol).setValues(rows);

      // --- Recompute totals with EXACT SAME APPROACH as full rebuild ---
      const debitColFixed = 9; // Same as full rebuild
      const creditColFixed = 10; // Same as full rebuild
      const totalDebit = rows.reduce((sum, r) => sum + cleanCurrency(r[debitColFixed - 1]), 0);
      const totalCredit = rows.reduce((sum, r) => sum + cleanCurrency(r[creditColFixed - 1]), 0);
      const totalRow = rows.length + 2;
      
      // --- Add yellow separator row ---
      sheet.getRange(totalRow, 1, 1, sheet.getLastColumn())
        .setBackground("#FFF9C4") // Light yellow background
        .setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
      
      const totalsStartRow = totalRow + 1;
      
      // Labels row (dark blue background, white text) - EXACT SAME as full rebuild
      sheet.getRange(totalsStartRow, debitColFixed).setValue("Total Debit")
        .setBackground("#0b5394")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
      
      sheet.getRange(totalsStartRow, creditColFixed).setValue("Total Credit")
        .setBackground("#0b5394")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
      
      sheet.getRange(totalsStartRow, creditColFixed + 1).setValue("Remaining funds")
        .setBackground("#0b5394")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");

      // Values row with EXACT SAME COLORS as full rebuild
      const valuesRow = totalsStartRow + 1;
      sheet.getRange(valuesRow, debitColFixed).setValue(totalDebit)
        .setBackground("#d9ead3") // Light green for debit
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setNumberFormat("#,##0");
      
      sheet.getRange(valuesRow, creditColFixed).setValue(totalCredit)
        .setBackground("#fce5cd") // Light orange for credit
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setNumberFormat("#,##0");
      
      const remainingFunds = totalCredit - totalDebit;
      sheet.getRange(valuesRow, creditColFixed + 1).setValue(remainingFunds)
        .setBackground("#FFF9C4") // Light yellow for remaining
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setNumberFormat("#,##0");
    }
  });

  const successMessage = TEST_MODE_ENABLED 
    ? `🧪 Test Mode: Quick update completed for '${TEST_ACCOUNT_NAME}'!`
    : "⚡ Quick update completed — existing account sheets refreshed!";
  SpreadsheetApp.getUi().alert(successMessage);
};

// ===============================
// 🚀 SMART UPDATE — Update only new/changed data
// ===============================
const smartUpdateAccounts = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("VN - Master Ledger");
  if (!masterSheet) throw new Error("❌ 'VN - Master Ledger' sheet not found.");

  const data = masterSheet.getDataRange().getValues();
  const headers = data[0];
  const accountCol = headers.indexOf("Account");
  const transactionCol = headers.indexOf("N-Transaction Numb");
  const debitCol = headers.indexOf("Debit (VND)");
  const creditCol = headers.indexOf("Credit (VND)");
  
  if (accountCol === -1 || transactionCol === -1) 
    throw new Error("❌ Missing required columns.");

  const accountMap = {};
  data.slice(1).forEach(row => {
    const account = row[accountCol];
    const transactionId = row[transactionCol];
    if (!account || !transactionId) return;
    if (!accountMap[account]) accountMap[account] = [];
    accountMap[account].push({ row, transactionId });
  });

  let updatedCount = 0;
  Object.entries(accountMap).forEach(([account, entries]) => {
    const sheet = ss.getSheetByName(account);
    if (!sheet) return;

    // Get existing transaction IDs from the sheet
    const existingData = sheet.getDataRange().getValues();
    const existingTransactions = new Set();
    for (let i = 1; i < existingData.length; i++) {
      const transactionId = existingData[i][transactionCol];
      if (transactionId) existingTransactions.add(transactionId);
    }

    // Find new transactions
    const newEntries = entries.filter(entry => 
      !existingTransactions.has(entry.transactionId)
    );

    if (newEntries.length > 0) {
      // Add new rows
      const newRows = newEntries.map(e => e.row);
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, headers.length).setValues(newRows);
      
      // Sort the entire sheet by date
      const allData = sheet.getDataRange().getValues();
      const headerRow = allData[0];
      const dataRows = allData.slice(1);
      
      dataRows.sort((a, b) => {
        const dateA = new Date(a[1]); // Column B
        const dateB = new Date(b[1]);
        return dateA - dateB;
      });
      
      // Rewrite the sorted data
      sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
      
      // Update totals with proper formatting
      const totalDebit = dataRows.reduce((sum, r) => sum + cleanCurrency(r[debitCol]), 0);
      const totalCredit = dataRows.reduce((sum, r) => sum + cleanCurrency(r[creditCol]), 0);
      const totalRow = dataRows.length + 2;
      
      // --- Add yellow separator row ---
      sheet.getRange(totalRow, 1, 1, sheet.getLastColumn())
        .setBackground("#FFF9C4") // Light yellow background
        .setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
      
      const totalsStartRow = totalRow + 1;
      
      // Labels row (dark blue background, white text) - ALIGNED UNDER CORRECT COLUMNS
      sheet.getRange(totalsStartRow, debitCol).setValue("Total Debit")
        .setBackground("#0b5394")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
      
      sheet.getRange(totalsStartRow, creditCol).setValue("Total Credit")
        .setBackground("#0b5394")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
      
      sheet.getRange(totalsStartRow, creditCol + 1).setValue("Remaining funds")
        .setBackground("#0b5394")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");

      // Values row with CORRECT COLORS - Green for Debit, Orange for Credit, Yellow for Remaining
      const valuesRow = totalsStartRow + 1;
      sheet.getRange(valuesRow, debitCol).setValue(totalDebit)
        .setBackground("#d9ead3") // Light green for debit
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setNumberFormat("#,##0");
      
      sheet.getRange(valuesRow, creditCol).setValue(totalCredit)
        .setBackground("#fce5cd") // Light orange for credit
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setNumberFormat("#,##0");
      
      const remainingFunds = totalCredit - totalDebit;
      sheet.getRange(valuesRow, creditCol + 1).setValue(remainingFunds)
        .setBackground("#FFF9C4") // Light yellow for remaining
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setNumberFormat("#,##0");
      
      updatedCount += newEntries.length;
    }
  });

  const successMessage = TEST_MODE_ENABLED 
    ? `🧪 Test Mode: Smart update completed — ${updatedCount} new transactions added to '${TEST_ACCOUNT_NAME}'!`
    : `🚀 Smart update completed — ${updatedCount} new transactions added across all account sheets!`;
  SpreadsheetApp.getUi().alert(successMessage);
};

// --- Helper: Clean currency values ---
const cleanCurrency = value => {
  if (!value) return 0;
  if (typeof value === "number") return value;
  const str = value.toString().replace(/[^\d\-.,]/g, "").replace(",", ".");
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
};