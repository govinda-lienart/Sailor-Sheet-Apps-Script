// ===============================
// 💰 FUND MANAGER SCRIPT
// ===============================

// --- Helper: Safely convert text/number currency to number ---
const cleanCurrency = value => {
    if (!value) return 0;
    if (typeof value === "number") return value;
    const str = value.toString().replace(/[^\d\-.,]/g, "").replace(",", ".");
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  };
  
  // ===============================
  // 🔹 FULL REBUILD — Create or overwrite all fund sheets
  // ===============================
  function updateAllFunds() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName("VN - Master Ledger");
    if (!masterSheet) throw new Error("❌ 'VN - Master Ledger' sheet not found.");
  
    const data = masterSheet.getDataRange().getValues();
    const richData = masterSheet.getDataRange().getRichTextValues();
    const headers = data[0];
  
    const fundCol = headers.indexOf("Funds");
    const accountCol = headers.indexOf("Account");
    if (fundCol === -1 || accountCol === -1)
      throw new Error("❌ Missing 'Funds' or 'Account' column in master sheet.");
  
    // Only show revenues and expenses
    const relevantAccounts = ["VN - Expenses", "VN - Revenues"];
  
    // --- Collect all rows by fund ---
    const fundMap = {};
    data.slice(1).forEach((row, i) => {
      const fund = row[fundCol];
      const account = row[accountCol];
      if (!fund || !relevantAccounts.includes(account)) return;
      if (!fundMap[fund]) fundMap[fund] = [];
      fundMap[fund].push({ row, rich: richData[i + 1] });
    });
  
    // --- Copy header formatting and column widths ---
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
      const headerRangeTarget = targetSheet.getRange(1, 1, 1, headers.length);
      headerRangeTarget
        .setFontSize(12)
        .setFontWeight("bold")
        .setFontFamily("Arial")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setBackground("#0b5394")
        .setFontColor("#ffffff");
  
      // --- Match column widths ---
      columnWidths.forEach((width, i) => targetSheet.setColumnWidth(i + 1, width));
  
      // --- Write data ---
      const rows = entries.map(e => e.row);
      targetSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  
      // --- Basic styling ---
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
  
      // --- Format date column (C = dd/mm/yy) ---
      targetSheet.getRange(2, 3, rows.length, 1).setNumberFormat("dd/mm/yy");
  
      // --- Compute totals (Debit = I, Credit = J) ---
      const debitCol = 9;
      const creditCol = 10;
      const totalDebit = rows.reduce((sum, r) => sum + cleanCurrency(r[debitCol - 1]), 0);
      const totalCredit = rows.reduce((sum, r) => sum + cleanCurrency(r[creditCol - 1]), 0);
      const remaining = totalCredit - totalDebit;
  
      const totalRow = rows.length + 2;
  
      // --- Totals Row ---
      targetSheet.getRange(totalRow, debitCol - 1).setValue("TOTALS:");
      targetSheet.getRange(totalRow, debitCol).setValue(totalDebit);
      targetSheet.getRange(totalRow, creditCol).setValue(totalCredit);
      targetSheet.getRange(totalRow, creditCol + 1)
        .setValue(`Remaining = ${remaining.toLocaleString()} VND`);
  
      // --- Formatting ---
      targetSheet.getRange(totalRow, debitCol, 1, 2).setNumberFormat("#,##0");
      targetSheet.getRange(totalRow, 1, 1, headers.length)
        .setBackground("#FFF9C4")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
  
      // --- Visual polish ---
      targetSheet.setFrozenRows(1);
      targetSheet.setHiddenGridlines(true);
      targetSheet.getRange(1, 1, targetSheet.getLastRow(), targetSheet.getLastColumn())
        .setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
    });
  
    SpreadsheetApp.getUi().alert("✅ All fund sheets updated successfully!");
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
    const accountCol = headers.indexOf("Account");
    const debitCol = headers.indexOf("Debit (VND)");
    const creditCol = headers.indexOf("Credit (VND)");
    if (fundCol === -1 || accountCol === -1)
      throw new Error("❌ Missing 'Funds' or 'Account' column.");
  
    const relevantAccounts = ["VN - Expenses", "VN - Revenues"];
    const fundMap = {};
    data.slice(1).forEach(row => {
      const fund = row[fundCol];
      const account = row[accountCol];
      if (!fund || !relevantAccounts.includes(account)) return;
      if (!fundMap[fund]) fundMap[fund] = [];
      fundMap[fund].push(row);
    });
  
    Object.keys(fundMap).forEach(fund => {
      const sheet = ss.getSheetByName("Fund - " + fund);
      if (sheet) {
        const rows = fundMap[fund];
        const lastCol = sheet.getLastColumn();
        sheet.getRange(2, 1, sheet.getLastRow(), lastCol).clearContent();
        sheet.getRange(2, 1, rows.length, lastCol).setValues(rows);
  
        // --- Recompute totals ---
        const totalDebit = rows.reduce((sum, r) => sum + cleanCurrency(r[debitCol]), 0);
        const totalCredit = rows.reduce((sum, r) => sum + cleanCurrency(r[creditCol]), 0);
        const remaining = totalCredit - totalDebit;
        const totalRow = rows.length + 2;
  
        sheet.getRange(totalRow, debitCol - 1).setValue("TOTALS:");
        sheet.getRange(totalRow, debitCol).setValue(totalDebit);
        sheet.getRange(totalRow, creditCol).setValue(totalCredit);
        sheet.getRange(totalRow, creditCol + 1)
          .setValue(`Remaining = ${remaining.toLocaleString()} VND`);
  
        sheet.getRange(totalRow, debitCol, 1, 2).setNumberFormat("#,##0");
        sheet.getRange(totalRow, 1, 1, headers.length)
          .setBackground("#FFF9C4")
          .setFontWeight("bold")
          .setHorizontalAlignment("center");
      }
    });
  
    SpreadsheetApp.getUi().alert("⚡ Quick update completed — existing fund sheets refreshed!");
  }
  
  
  // ===============================
  // 🔹 MENU BUILDER
  // ===============================
  function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("💰 Fund Manager")
      .addItem("🔁 Full Rebuild (All Funds)", "updateAllFunds")
      .addItem("⚡ Quick Update (Existing Only)", "quickUpdateFunds")
      .addToUi();
  }