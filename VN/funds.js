// ================================
// 💰 FUND MANAGER SCRIPT
// ================================

const updateAllFunds = () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName("VN - Master Ledger");
    const data = masterSheet.getDataRange().getValues();
    const richData = masterSheet.getDataRange().getRichTextValues();
    const headers = data[0];
  
    const fundCol = headers.indexOf("Funds");
    const accountCol = headers.indexOf("Account");
    
    if (fundCol === -1) throw new Error("No 'Funds' column found in VN - Master Ledger");
    if (accountCol === -1) throw new Error("No 'Account' column found in VN - Master Ledger");
  
    // --- Filter: Only include rows where Account is "VN - Expenses" or "VN - Revenues" ---
    const allowedAccounts = ["VN - Expenses", "VN - Revenues"];
    
    // --- Collect all rows by fund (filtered by allowed accounts) ---
    const fundMap = {};
    data.slice(1).forEach((row, i) => {
      const account = row[accountCol];
      const fund = row[fundCol];
      
      // Only process if account is allowed and fund exists
      if (!fund || !account) return;
      if (!allowedAccounts.includes(account)) return;
      
      if (!fundMap[fund]) fundMap[fund] = [];
      fundMap[fund].push({ row, rich: richData[i + 1] });
    });
  
    // --- Copy header formatting and column widths ---
    const headerRange = masterSheet.getRange(1, 1, 1, headers.length);
    const columnWidths = Array.from({ length: headers.length }, (_, i) =>
      masterSheet.getColumnWidth(i + 1)
    );
  
    // --- Loop through funds ---
    Object.entries(fundMap).forEach(([fund, entries]) => {
      const cleanName = fund.replace(/[\\\/\?\*\[\]]/g, " ");
      const sheetName = `Fund - ${cleanName}`;
      let targetSheet = ss.getSheetByName(sheetName);
      if (!targetSheet) targetSheet = ss.insertSheet(sheetName);
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
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Enable text wrapping
        .setBackground("#0b5394") // deep blue
        .setFontColor("#ffffff")
        .setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID); // White borders between columns
  
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
  
      // --- Define column positions first ---
      const debitCol = 9;
      const creditCol = 10;
      
      // --- Format date column (B) ---
      targetSheet.getRange(2, 2, sortedRows.length, 1).setNumberFormat("dd/mm/yy");
      
      // --- Format Debit and Credit columns with thousands separators ---
      targetSheet.getRange(2, debitCol, sortedRows.length, 1).setNumberFormat("#,##0");
      targetSheet.getRange(2, creditCol, sortedRows.length, 1).setNumberFormat("#,##0");
  
      // --- Compute totals (Debit = I, Credit = J) ---
      const totalDebit = sortedRows.reduce((sum, r) => sum + cleanCurrency(r[debitCol - 1]), 0);
      const totalCredit = sortedRows.reduce((sum, r) => sum + cleanCurrency(r[creditCol - 1]), 0);
  
      const totalRow = sortedRows.length + 2;
      
      // --- Add yellow separator row ---
      targetSheet.getRange(totalRow, 1, 1, headers.length)
        .setBackground("#FFF9C4") // Light yellow background
        .setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
      
      const totalsStartRow = totalRow + 1;
      
      // --- Create totals section ---
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
      
      const remainingFunds = totalDebit - totalCredit; // Debit (money in) - Credit (expenses) = Remaining
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
  
      // --- Add "Last update" timestamp at bottom right ---
      const lastUpdateRow = totalRowIndex + 2;
      const lastUpdateCol = totalColIndex;
      const now = new Date();
      const timestamp = `Last update: ${now.toLocaleString()}`;
      
      targetSheet.getRange(lastUpdateRow, lastUpdateCol).setValue(timestamp)
        .setFontSize(9)
        .setFontColor("#666666")
        .setFontStyle("italic")
        .setHorizontalAlignment("right")
        .setVerticalAlignment("bottom");
      
      // --- Update row height for timestamp row ---
      targetSheet.setRowHeight(lastUpdateRow, 21);
    });
  
    SpreadsheetApp.getUi().alert("✅ All fund sheets updated successfully!");
  };
  
  // ===============================
  // 🎯 REBUILD CURRENT FUND — Update only the fund sheet you're currently viewing
  // ===============================
  const rebuildCurrentFund = () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentSheet = ss.getActiveSheet();
    const currentSheetName = currentSheet.getName();
    
    // Check if current sheet is a fund sheet
    if (!currentSheetName.startsWith("Fund - ")) {
      SpreadsheetApp.getUi().alert("⚠️ Cannot rebuild this sheet. Please open a fund sheet (starting with 'Fund - ') and try again.");
      return;
    }
    
    // Skip if trying to rebuild the master ledger or summary sheets
    if (currentSheetName === "VN - Master Ledger" || currentSheetName === "Summary") {
      SpreadsheetApp.getUi().alert("⚠️ Cannot rebuild this sheet. Please open a fund sheet and try again.");
      return;
    }
    
    const masterSheet = ss.getSheetByName("VN - Master Ledger");
    if (!masterSheet) {
      SpreadsheetApp.getUi().alert("❌ 'VN - Master Ledger' sheet not found.");
      return;
    }
    
    const data = masterSheet.getDataRange().getValues();
    const richData = masterSheet.getDataRange().getRichTextValues();
    const headers = data[0];
    const fundCol = headers.indexOf("Funds");
    const accountCol = headers.indexOf("Account");
    
    if (fundCol === -1) {
      throw new Error("No 'Funds' column found in VN - Master Ledger");
    }
    if (accountCol === -1) {
      throw new Error("No 'Account' column found in VN - Master Ledger");
    }
    
    // Extract fund name from sheet name (remove "Fund - " prefix)
    const fundName = currentSheetName.replace(/^Fund - /, "");
    
    // --- Filter: Only include rows where Account is "VN - Expenses" or "VN - Revenues" ---
    const allowedAccounts = ["VN - Expenses", "VN - Revenues"];
    
    // --- Find all rows for this fund (filtered by allowed accounts) ---
    const fundEntries = [];
    data.slice(1).forEach((row, i) => {
      const account = row[accountCol];
      const fund = row[fundCol];
      
      // Only process if account is allowed, fund matches, and fund exists
      if (!fund || !account) return;
      if (!allowedAccounts.includes(account)) return;
      
      // Match by exact fund name or by cleaned name (handle special characters)
      const cleanFundName = fund.replace(/[\\\/\?\*\[\]]/g, " ");
      if (fund === fundName || cleanFundName === fundName) {
        fundEntries.push({ row, rich: richData[i + 1] });
      }
    });
    
    if (fundEntries.length === 0) {
      SpreadsheetApp.getUi().alert(`⚠️ No data found for fund '${fundName}' in the master ledger (filtered by VN - Expenses and VN - Revenues accounts).`);
      return;
    }
    
    // --- Clear and rebuild the sheet ---
    const targetSheet = currentSheet;
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
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Enable text wrapping
      .setBackground("#0b5394")
      .setFontColor("#ffffff")
      .setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID); // White borders between columns
    
    // --- Match column widths from master ---
    const columnWidths = Array.from({ length: headers.length }, (_, i) =>
      masterSheet.getColumnWidth(i + 1)
    );
    columnWidths.forEach((width, i) => targetSheet.setColumnWidth(i + 1, width));
    
    // --- Write data (sorted by date) ---
    const rows = fundEntries.map(e => e.row);
    const richRows = fundEntries.map(e => e.rich);
    
    // Sort by date column (column B - dd/mm/YY)
    const dateCol = 1;
    const sortedIndices = rows.map((row, index) => ({ row, rich: richRows[index], originalIndex: index }))
      .sort((a, b) => {
        const dateA = new Date(a.row[dateCol]);
        const dateB = new Date(b.row[dateCol]);
        return dateA - dateB; // Oldest to newest
      });
    
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
    
    // --- Reapply hyperlinks ---
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
    
    // --- Define column positions ---
    const debitCol = 9;
    const creditCol = 10;
    
    // --- Format date column (B) ---
    targetSheet.getRange(2, 2, sortedRows.length, 1).setNumberFormat("dd/mm/yy");
    
    // --- Format Debit and Credit columns ---
    targetSheet.getRange(2, debitCol, sortedRows.length, 1).setNumberFormat("#,##0");
    targetSheet.getRange(2, creditCol, sortedRows.length, 1).setNumberFormat("#,##0");
    
    // --- Compute totals ---
    const totalDebit = sortedRows.reduce((sum, r) => sum + cleanCurrency(r[debitCol - 1]), 0);
    const totalCredit = sortedRows.reduce((sum, r) => sum + cleanCurrency(r[creditCol - 1]), 0);
    const totalRow = sortedRows.length + 2;
    
    // --- Add yellow separator row ---
    targetSheet.getRange(totalRow, 1, 1, headers.length)
      .setBackground("#FFF9C4")
      .setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
    
    const totalsStartRow = totalRow + 1;
    
    // Labels row
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
    
    // Values row
    const valuesRow = totalsStartRow + 1;
    targetSheet.getRange(valuesRow, debitCol).setValue(totalDebit)
      .setBackground("#d9ead3")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setNumberFormat("#,##0");
    
    targetSheet.getRange(valuesRow, creditCol).setValue(totalCredit)
      .setBackground("#fce5cd")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setNumberFormat("#,##0");
    
    const remainingFunds = totalDebit - totalCredit;
    targetSheet.getRange(valuesRow, creditCol + 1).setValue(remainingFunds)
      .setBackground("#FFF9C4")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setNumberFormat("#,##0");
    
    // --- Hide column N (if exists) ---
    const lastCol = targetSheet.getLastColumn();
    if (lastCol >= 14) targetSheet.hideColumns(14);
    
    // --- Gridlines and borders ---
    targetSheet.setFrozenRows(1);
    targetSheet.setHiddenGridlines(true);
    const lastRow = targetSheet.getLastRow();
    targetSheet
      .getRange(1, 1, lastRow, lastCol)
      .setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
  
    // --- Add "Last update" timestamp at bottom right ---
    const lastUpdateRow = lastRow + 2;
    const lastUpdateCol = lastCol;
    const now = new Date();
    const timestamp = `Last update: ${now.toLocaleString()}`;
    
    targetSheet.getRange(lastUpdateRow, lastUpdateCol).setValue(timestamp)
      .setFontSize(9)
      .setFontColor("#666666")
      .setFontStyle("italic")
      .setHorizontalAlignment("right")
      .setVerticalAlignment("bottom");
  
    // --- Set row height (21 px) for all rows including timestamp ---
    const totalRows = lastUpdateRow;
    for (let r = 1; r <= totalRows; r++) {
      targetSheet.setRowHeight(r, 21);
    }
    
    // --- Format columns ---
    targetSheet.getRange(2, debitCol, sortedRows.length, 1).setBackground("#d9ead3");
    targetSheet.getRange(2, creditCol, sortedRows.length, 1).setBackground("#fce5cd");
    
    // Bill columns (yellow)
    const billCols = ["O", "P", "Q"].map(c => c.charCodeAt(0) - 64);
    billCols.forEach(col => {
      if (col <= headers.length)
        targetSheet.getRange(2, col, sortedRows.length, 1).setBackground("#fff2cc");
    });
    
    // Transaction number column (blue)
    targetSheet.getRange(2, 1, sortedRows.length, 1).setBackground("#c9daf8");
    
    SpreadsheetApp.getUi().alert(`✅ Fund '${fundName}' updated successfully!`);
  };
  
  
  