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
  
    // --- Loop through accounts ---
    Object.entries(accountMap).forEach(([account, entries]) => {
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
      
      // Sort by date column (column B - dd/mm/YY)
      const dateCol = 1; // Column B is index 1
      rows.sort((a, b) => {
        const dateA = new Date(a[dateCol]);
        const dateB = new Date(b[dateCol]);
        return dateA - dateB; // Oldest to newest
      });
      
      targetSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  
      // --- Center + style data ---
      targetSheet.getRange(2, 1, rows.length, headers.length)
        .setFontSize(11)
        .setFontFamily("Arial")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  
      // --- Reapply hyperlinks ---
      entries.forEach((entry, rIdx) => {
        entry.rich.forEach((rich, cIdx) => {
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
      targetSheet.getRange(2, 2, rows.length, 1).setNumberFormat("dd/mm/yy");
  
      // --- Compute totals (Debit = I, Credit = J) ---
      const debitCol = 9;
      const creditCol = 10;
      const totalDebit = rows.reduce((sum, r) => sum + cleanCurrency(r[debitCol - 1]), 0);
      const totalCredit = rows.reduce((sum, r) => sum + cleanCurrency(r[creditCol - 1]), 0);
  
      const totalRow = rows.length + 2;
      targetSheet.getRange(totalRow, debitCol - 1).setValue("TOTAL:");
      targetSheet.getRange(totalRow, debitCol).setValue(totalDebit);
      targetSheet.getRange(totalRow, creditCol).setValue(totalCredit);
  
      // --- Style total row ---
      const totalRange = targetSheet.getRange(totalRow, 1, 1, headers.length);
      totalRange
        .setBackground("#FFF9C4")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
      targetSheet.getRange(totalRow, debitCol, 1, 2).setNumberFormat("#,##0");
  
      // --- Hide column N (if exists) ---
      const lastCol = targetSheet.getLastColumn();
      if (lastCol >= 14) targetSheet.hideColumns(14);
  
      // --- Set row height (21 px) ---
      Array.from({ length: totalRow }, (_, i) => i + 1).forEach(r =>
        targetSheet.setRowHeight(r, 21)
      );
  
      // --- Format Debit/Credit columns ---
      targetSheet.getRange(2, debitCol, rows.length, 1).setBackground("#d9ead3"); // green tint
      targetSheet.getRange(2, creditCol, rows.length, 1).setBackground("#fce5cd"); // orange tint
  
      // --- Format Bill/Red Bill/Doc columns (yellow) ---
      const billCols = ["O", "P", "Q"].map(c => c.charCodeAt(0) - 64);
      billCols.forEach(col => {
        if (col <= headers.length)
          targetSheet.getRange(2, col, rows.length, 1).setBackground("#fff2cc");
      });
  
      // --- Transaction number column (blue tint) ---
      targetSheet.getRange(2, 1, rows.length, 1).setBackground("#c9daf8");
  
      // --- Remove gridlines, add outer borders ---
      targetSheet.setFrozenRows(1);
      targetSheet.setHiddenGridlines(true);
      const totalRowIndex = targetSheet.getLastRow();
      const totalColIndex = targetSheet.getLastColumn();
      targetSheet
        .getRange(1, 1, totalRowIndex, totalColIndex)
        .setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
    });
  
    SpreadsheetApp.getUi().alert("✅ All account sheets updated successfully!");
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

    Object.keys(accountMap).forEach(account => {
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

        // --- Recompute totals ---
        const totalDebit = rows.reduce((sum, r) => sum + cleanCurrency(r[debitCol]), 0);
        const totalCredit = rows.reduce((sum, r) => sum + cleanCurrency(r[creditCol]), 0);
        const totalRow = rows.length + 2;
        sheet.getRange(totalRow, debitCol).setValue(totalDebit);
        sheet.getRange(totalRow, creditCol).setValue(totalCredit);
        sheet.getRange(totalRow, debitCol, 1, 2).setNumberFormat("#,##0");
      }
    });

    SpreadsheetApp.getUi().alert("⚡ Quick update completed — existing account sheets refreshed!");
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
        
        // Update totals
        const totalDebit = dataRows.reduce((sum, r) => sum + cleanCurrency(r[debitCol]), 0);
        const totalCredit = dataRows.reduce((sum, r) => sum + cleanCurrency(r[creditCol]), 0);
        const totalRow = dataRows.length + 2;
        sheet.getRange(totalRow, debitCol - 1).setValue("TOTAL:");
        sheet.getRange(totalRow, debitCol).setValue(totalDebit);
        sheet.getRange(totalRow, creditCol).setValue(totalCredit);
        sheet.getRange(totalRow, debitCol, 1, 2).setNumberFormat("#,##0");
        
        updatedCount += newEntries.length;
      }
    });

    SpreadsheetApp.getUi().alert(`🚀 Smart update completed — ${updatedCount} new transactions added across all account sheets!`);
  };

  // --- Helper: Clean currency values ---
  const cleanCurrency = value => {
    if (!value) return 0;
    if (typeof value === "number") return value;
    const str = value.toString().replace(/[^\d\-.,]/g, "").replace(",", ".");
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  };