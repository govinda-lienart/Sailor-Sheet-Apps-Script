function createOrUpdateAuditSummary() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName("VN - Master Ledger");
    if (!masterSheet) throw new Error("❌ Sheet 'VN - Master Ledger' not found.");
  
    const data = masterSheet.getDataRange().getValues();
    const richData = masterSheet.getDataRange().getRichTextValues();
    const headers = data[0];
    const rows = data.slice(1);
  
    const accountCol = headers.indexOf("Account");
    const fundCol = headers.indexOf("Funds");
    const debitCol = headers.indexOf("Debit (VND)");
    const creditCol = headers.indexOf("Credit (VND)");
  
    if (accountCol === -1 || fundCol === -1 || debitCol === -1 || creditCol === -1)
      throw new Error("❌ Missing required columns in VN - Master Ledger.");
  
    // --- Group by Account ---
    const accountMap = {};
    rows.forEach((row, i) => {
      const account = row[accountCol];
      const debit = parseFloat(row[debitCol]) || 0;
      const credit = parseFloat(row[creditCol]) || 0;
      if (!account) return;
      if (!accountMap[account]) accountMap[account] = { debit: 0, credit: 0, rich: [] };
      accountMap[account].debit += debit;
      accountMap[account].credit += credit;
      accountMap[account].rich.push(richData[i + 1]);
    });
  
    const accountData = Object.entries(accountMap).map(([account, { debit, credit }]) => {
      return [
        account, // Just use the account name directly for now
        debit,
        credit,
        debit - credit,
      ];
    });
  
    // --- Group by Fund (only Revenue and Expense accounts) ---
    const fundMap = {};
    rows.forEach((row, i) => {
      const fund = row[fundCol];
      const account = row[accountCol];
      const debit = parseFloat(row[debitCol]) || 0;
      const credit = parseFloat(row[creditCol]) || 0;
      
      // Only include Revenue and Expense accounts
      if (!fund || !account) return;
      if (!/Revenue|Expense/i.test(account)) return;
      
      if (!fundMap[fund]) fundMap[fund] = { debit: 0, credit: 0, rich: [] };
      fundMap[fund].debit += debit;
      fundMap[fund].credit += credit;
      fundMap[fund].rich.push(richData[i + 1]);
    });
  
    const fundData = Object.entries(fundMap).map(([fund, { debit, credit }]) => {
      return [
        "Fund",
        fund, // Just use the fund name directly for now
        debit,
        credit,
        credit - debit, // Remaining balance
      ];
    });
  
    // --- Categorize accounts into 3 groups ---
    const expensesRevenues = accountData.filter(([a]) => /Expense|Revenue/i.test(a));
    const indovina = accountData.filter(([a]) => /Indovina/i.test(a));
    const custodians = accountData.filter(([a]) => !/Expense|Revenue|Indovina/i.test(a));
  
    // --- Create or clear Summary sheet ---
    const sheetName = "Summary";
    let summary = ss.getSheetByName(sheetName);
    if (!summary) summary = ss.insertSheet(sheetName);
    else summary.clear();
  
    // Move to first position
    ss.setActiveSheet(summary);
    ss.moveActiveSheet(1);
    summary.setHiddenGridlines(true);
  
    // --- Header section ---
    const now = new Date();
    summary.getRange("A1").setValue("THREE MONKEYS WILDLIFE CONSERVANCY");
    summary.getRange("A2").setValue("COMPREHENSIVE FINANCIAL SUMMARY REPORT");
    summary.getRange("A3").setValue("Generated on: " + now.toLocaleString());
  
    summary.getRange("A1").setFontSize(16).setFontWeight("bold").setFontColor("#0b5394");
    summary.getRange("A2").setFontSize(13).setFontStyle("italic").setFontColor("#0b5394");
    summary.getRange("A3").setFontSize(10).setFontColor("#666666");
  
    // Helper function to insert one formatted table
    function insertTable(title, startRow, data, isFundTable = false) {
      const headers = isFundTable ? 
        ["Type", "Name", "Total Debit (VND)", "Total Credit (VND)", "Remaining (VND)"] :
        ["Account", "Total Debit (VND)", "Total Credit (VND)", "Remaining funds"];
      const tableStart = startRow + 1;
      const dataStart = tableStart + 1;
  
      summary.getRange(startRow, 1).setValue(title)
        .setFontWeight("bold").setFontSize(12).setFontColor("#0b5394");
  
      // Header row
      summary.getRange(tableStart, 1, 1, headers.length).setValues([headers])
        .setFontWeight("bold")
        .setFontSize(11)
        .setFontColor("#ffffff")
        .setBackground("#0b5394")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
  
      // Data
      if (data.length > 0) {
        summary.getRange(dataStart, 1, data.length, headers.length).setValues(data);
        summary.getRange(dataStart, isFundTable ? 3 : 2, data.length, 3).setNumberFormat("#,##0");
        summary.getRange(dataStart, 1, data.length, headers.length)
          .setFontSize(10)
          .setFontFamily("Arial")
          .setVerticalAlignment("middle");
        
        // Add hyperlinks using the working method from the older script
        if (isFundTable) {
          // Add hyperlinks for fund names (column B)
          data.forEach((row, index) => {
            const fundName = row[1]; // Fund name is in column B
            const fundSheetName = "Fund - " + fundName;
            const fundSheet = ss.getSheetByName(fundSheetName);
            if (fundSheet) {
              const cell = summary.getRange(dataStart + index, 2);
              const richText = SpreadsheetApp.newRichTextValue()
                .setText(fundName)
                .setLinkUrl(`#gid=${fundSheet.getSheetId()}`)
                .build();
              cell.setRichTextValue(richText);
            }
          });
        } else {
          // Add hyperlinks for account names (column A)
          data.forEach((row, index) => {
            const accountName = row[0]; // Account name is in column A
            let accountSheet = ss.getSheetByName(accountName);
            
            // If exact name not found, try some common variations
            if (!accountSheet) {
              // Try with "Wallet" prefix (common for custodian accounts)
              accountSheet = ss.getSheetByName("VN - Wallet " + accountName.replace("VN - ", ""));
            }
            if (!accountSheet) {
              // Try without "VN - " prefix
              accountSheet = ss.getSheetByName(accountName.replace("VN - ", ""));
            }
            if (!accountSheet) {
              // Try with just the name part
              const namePart = accountName.split(" ").slice(-2).join(" "); // Get last two words
              accountSheet = ss.getSheetByName(namePart);
            }
            
            if (accountSheet) {
              const cell = summary.getRange(dataStart + index, 1);
              const richText = SpreadsheetApp.newRichTextValue()
                .setText(accountName)
                .setLinkUrl(`#gid=${accountSheet.getSheetId()}`)
                .build();
              cell.setRichTextValue(richText);
            } else {
              // Debug: log what sheets are available
              console.log(`Sheet not found for: ${accountName}`);
              const allSheets = ss.getSheets().map(s => s.getName());
              console.log(`Available sheets: ${allSheets.join(", ")}`);
            }
          });
        }
  
        // Totals
        const totalDebit = data.reduce((s, r) => s + r[isFundTable ? 2 : 1], 0);
        const totalCredit = data.reduce((s, r) => s + r[isFundTable ? 3 : 2], 0);
        const totalDiff = isFundTable ? totalCredit - totalDebit : totalDebit - totalCredit;
        const totalRow = dataStart + data.length;
        
        const totalLabel = isFundTable ? "GRAND TOTAL (All Account)" : "TOTAL";
        const totalData = isFundTable ? 
          ["", totalLabel, totalDebit, totalCredit, totalDiff] :
          [totalLabel, totalDebit, totalCredit, totalDiff];
          
        summary.getRange(totalRow, 1, 1, headers.length).setValues([totalData]);
        summary.getRange(totalRow, 1, 1, headers.length)
          .setBackground("#FFF9C4")
          .setFontWeight("bold")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
        summary.getRange(totalRow, isFundTable ? 3 : 2, 1, 3).setNumberFormat("#,##0");
  
        // Border
        const fullRange = summary.getRange(tableStart, 1, data.length + 1, headers.length);
        fullRange.setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
  
        return totalRow + 2; // Next table starts after some spacing
      } else {
        summary.getRange(tableStart + 1, 1).setValue("No data available.")
          .setFontStyle("italic")
          .setFontColor("#888888");
        return tableStart + 3;
      }
    }
  
    // --- Insert Fund Summary first ---
    let nextRow = 7;
    console.log("Fund Data:", fundData); // Debug log
    nextRow = insertTable("📊 FUND SUMMARY", nextRow, fundData, true);
    
    // Add spacing and section header
    nextRow += 3;
    summary.getRange(nextRow, 1).setValue("📊 ACCOUNTS SUMMARY")
      .setFontWeight("bold").setFontSize(14).setFontColor("#0b5394");
    nextRow += 2;
    
    // --- Insert Account Summary tables ---
    console.log("Expenses/Revenues:", expensesRevenues); // Debug log
    nextRow = insertTable("I. Revenues & Expenses", nextRow, expensesRevenues);
    nextRow = insertTable("II. Indovina Bank Accounts", nextRow, indovina);
    nextRow = insertTable("III. Custodian Accounts", nextRow, custodians);
  
    // --- Add explanatory note below custodian section ---
    summary.getRange(nextRow + 1, 1).setValue(
      "Note: The above Custodian Accounts represent project funds temporarily advanced to team members " +
      "for field operations and related expenses. Balances include both cash and bank-based holdings, " +
      "as recorded in the VN - Master Ledger."
    )
      .setFontSize(9)
      .setFontStyle("italic")
      .setFontColor("#777777")
      .setWrap(true)
      .setBackground("#f5f5f5");
  
    summary.autoResizeColumns(1, 5);
    summary.setRowHeights(1, nextRow + 3, 22);
    summary.getRange("A1:A3").setHorizontalAlignment("left");
  
    SpreadsheetApp.getUi().alert("✅ Comprehensive Summary sheet updated successfully!");
  }
  
  
  // ===============================
  // 🔹 HELPER FUNCTIONS
  // ===============================
  
  // --- Function to open the Summary sheet ---
  function openSummarySheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Summary");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("⚠️ The Summary sheet does not exist. Please generate it first.");
      return;
    }
    ss.setActiveSheet(sheet);
    SpreadsheetApp.getUi().alert("📄 Summary sheet opened.");
  }

  // ===============================
  // 📊 FUND SUMMARY CREATOR
  // ===============================
  function createFundSummary() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName("VN - Master Ledger");
    if (!masterSheet) throw new Error("❌ Sheet 'VN - Master Ledger' not found.");
  
    const data = masterSheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
  
    const fundCol = headers.indexOf("Funds");
    const debitCol = headers.indexOf("Debit (VND)");
    const creditCol = headers.indexOf("Credit (VND)");
  
    if (fundCol === -1 || debitCol === -1 || creditCol === -1)
      throw new Error("❌ Missing required columns in VN - Master Ledger.");
  
    // --- Group by Fund ---
    const fundMap = {};
    rows.forEach(row => {
      const fund = row[fundCol];
      const debit = parseFloat(row[debitCol]) || 0;
      const credit = parseFloat(row[creditCol]) || 0;
      if (!fund) return;
      if (!fundMap[fund]) fundMap[fund] = { debit: 0, credit: 0 };
      fundMap[fund].debit += debit;
      fundMap[fund].credit += credit;
    });
  
    const fundData = Object.entries(fundMap).map(([fund, { debit, credit }]) => [
      "Fund",
      fund,
      debit,
      credit,
      credit - debit, // Remaining balance
    ]);
  
    // --- Create or clear Fund Summary sheet ---
    const sheetName = "Fund Summary";
    let summary = ss.getSheetByName(sheetName);
    if (!summary) summary = ss.insertSheet(sheetName);
    else summary.clear();
  
    // Move to second position
    ss.setActiveSheet(summary);
    ss.moveActiveSheet(2);
    summary.setHiddenGridlines(true);
  
    // --- Header section ---
    const now = new Date();
    summary.getRange("A1").setValue("THREE MONKEYS WILDLIFE CONSERVANCY");
    summary.getRange("A2").setValue("FUND SUMMARY REPORT");
    summary.getRange("A3").setValue("Generated on: " + now.toLocaleString());
  
    summary.getRange("A1").setFontSize(16).setFontWeight("bold").setFontColor("#0b5394");
    summary.getRange("A2").setFontSize(13).setFontStyle("italic").setFontColor("#0b5394");
    summary.getRange("A3").setFontSize(10).setFontColor("#666666");
  
    // --- Table headers ---
    const headers_fund = ["Type", "Name", "Total Debit (VND)", "Total Credit (VND)", "Remaining (VND)"];
    const tableStart = 6;
    const dataStart = tableStart + 1;
  
    // Header row
    summary.getRange(tableStart, 1, 1, headers_fund.length).setValues([headers_fund])
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontColor("#ffffff")
      .setBackground("#0b5394")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  
    // Data
    if (fundData.length > 0) {
      summary.getRange(dataStart, 1, fundData.length, headers_fund.length).setValues(fundData);
      summary.getRange(dataStart, 3, fundData.length, 3).setNumberFormat("#,##0");
      summary.getRange(dataStart, 1, fundData.length, headers_fund.length)
        .setFontSize(10)
        .setFontFamily("Arial")
        .setVerticalAlignment("middle");
  
      // Grand Total
      const totalDebit = fundData.reduce((s, r) => s + r[2], 0);
      const totalCredit = fundData.reduce((s, r) => s + r[3], 0);
      const totalRemaining = totalCredit - totalDebit;
      const totalRow = dataStart + fundData.length;
      summary.getRange(totalRow, 1, 1, 5).setValues([
        ["", "GRAND TOTAL (All Account)", totalDebit, totalCredit, totalRemaining],
      ]);
      summary.getRange(totalRow, 1, 1, 5)
        .setBackground("#FFF9C4")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
      summary.getRange(totalRow, 3, 1, 3).setNumberFormat("#,##0");
  
      // Border
      const fullRange = summary.getRange(tableStart, 1, fundData.length + 1, 5);
      fullRange.setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
    }
  
    summary.autoResizeColumns(1, 5);
    summary.setRowHeights(1, totalRow + 3, 22);
    summary.getRange("A1:A3").setHorizontalAlignment("left");
  
    SpreadsheetApp.getUi().alert("✅ Fund Summary sheet updated successfully!");
  }