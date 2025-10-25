// ===============================
// 💰 FUND MANAGER SCRIPT
// ===============================

// ===============================
// 🏗️ FULL REBUILD — Create or overwrite all FUND sheets
// ===============================
function updateAllFunds() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName("VN - Master Ledger");
    if (!masterSheet) throw new Error("❌ 'VN - Master Ledger' sheet not found.");
  
    const data = masterSheet.getDataRange().getValues();
    const richData = masterSheet.getDataRange().getRichTextValues();
    const headers = data[0];
  
    const fundCol = headers.indexOf("Funds");
    if (fundCol === -1) throw new Error("❌ No 'Funds' column found in VN - Master Ledger");
  
    // --- Collect all rows by fund ---
    const fundMap = {};
    data.slice(1).forEach((row, i) => {
      const fund = row[fundCol];
      if (!fund) return;
      if (!fundMap[fund]) fundMap[fund] = [];
      fundMap[fund].push({ row, rich: richData[i + 1] });
    });
  
    // --- Copy column widths for consistency ---
    const columnWidths = Array.from({ length: headers.length }, (_, i) =>
      masterSheet.getColumnWidth(i + 1)
    );
  
    // --- Loop through funds ---
    Object.entries(fundMap).forEach(([fund, entries]) => {
      const cleanName = "Fund - " + fund.replace(/[\\\/\?\*\[\]]/g, " ");
      let targetSheet = ss.getSheetByName(cleanName);
      if (!targetSheet) targetSheet = ss.insertSheet(cleanName);
      targetSheet.clear();
  
      // --- Header row ---
      targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      targetSheet.getRange(1, 1, 1, headers.length)
        .setFontSize(12)
        .setFontWeight("bold")
        .setFontFamily("Arial")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setBackground("#134f5c")
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
  
      // --- Style data ---
      targetSheet.getRange(2, 1, rows.length, headers.length)
        .setFontSize(11)
        .setFontFamily("Arial")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
  
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
  
      // --- Totals ---
      const debitCol = 9;
      const creditCol = 10;
      const totalDebit = rows.reduce((sum, r) => sum + cleanCurrency(r[debitCol - 1]), 0);
      const totalCredit = rows.reduce((sum, r) => sum + cleanCurrency(r[creditCol - 1]), 0);
      const totalRow = rows.length + 2;
  
      targetSheet.getRange(totalRow, debitCol - 1).setValue("TOTAL:");
      targetSheet.getRange(totalRow, debitCol).setValue(totalDebit);
      targetSheet.getRange(totalRow, creditCol).setValue(totalCredit);
      targetSheet.getRange(totalRow, debitCol, 1, 2).setNumberFormat("#,##0");
      targetSheet.getRange(totalRow, 1, 1, headers.length)
        .setBackground("#D9EAD3")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
  
      // --- Borders, polish ---
      targetSheet.setFrozenRows(1);
      targetSheet.setHiddenGridlines(true);
      targetSheet.getRange(1, 1, targetSheet.getLastRow(), targetSheet.getLastColumn())
        .setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
    });
  
    SpreadsheetApp.getUi().alert("✅ All fund sheets created/updated successfully!");
  }
  
  // ===============================
  // ⚡ QUICK UPDATE — Refresh existing fund sheets only
  // ===============================
  function quickUpdateFunds() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName("VN - Master Ledger");
    if (!masterSheet) throw new Error("❌ 'VN - Master Ledger' sheet not found.");
  
    const data = masterSheet.getDataRange().getValues();
    const headers = data[0];
    const fundCol = headers.indexOf("Funds");
    const debitCol = headers.indexOf("Debit (VND)");
    const creditCol = headers.indexOf("Credit (VND)");
    if (fundCol === -1) throw new Error("❌ No 'Funds' column found.");
  
    const fundMap = {};
    data.slice(1).forEach(row => {
      const fund = row[fundCol];
      if (!fund) return;
      if (!fundMap[fund]) fundMap[fund] = [];
      fundMap[fund].push(row);
    });
  
    Object.keys(fundMap).forEach(fund => {
      const sheet = ss.getSheetByName("Fund - " + fund);
      if (sheet) {
        const rows = fundMap[fund];
        
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
  
    SpreadsheetApp.getUi().alert("⚡ Quick update completed — existing fund sheets refreshed!");
  }