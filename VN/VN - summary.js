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
        fund, // Just the fund name (no "Type" column)
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
    const startCol = 2; // Start from column B for left margin (column A is empty)
    summary.getRange(1, startCol).setValue("THREE MONKEYS WILDLIFE CONSERVANCY");
    summary.getRange(2, startCol).setValue("COMPREHENSIVE FINANCIAL SUMMARY REPORT");
    summary.getRange(3, startCol).setValue("Generated on: " + now.toLocaleString());
  
    summary.getRange(1, startCol).setFontSize(16).setFontWeight("bold").setFontColor("#0b5394");
    summary.getRange(2, startCol).setFontSize(13).setFontStyle("italic").setFontColor("#0b5394");
    summary.getRange(3, startCol).setFontSize(10).setFontColor("#666666");
  
    // Helper function to insert one formatted table
    // tableType: "fund" = funds (debit=orange, credit=green), "revenue_expense" = same, "asset" = asset accounts (debit=green, credit=orange)
    function insertTable(title, startRow, data, isFundTable = false, tableType = "asset", startCol = 2) {
      const headers = isFundTable ? 
        ["Name", "Total Debit (VND)", "Total Credit (VND)", "Remaining (VND)"] :
        ["Account", "Total Debit (VND)", "Total Credit (VND)", "Remaining funds"];
      const tableStart = startRow + 1;
      const dataStart = tableStart + 1;
  
      if (title) {
        summary.getRange(startRow, startCol).setValue(title)
          .setFontWeight("bold").setFontSize(12).setFontColor("#0b5394");
      }
  
      // Header row
      summary.getRange(tableStart, startCol, 1, headers.length).setValues([headers])
        .setFontWeight("bold")
        .setFontSize(11)
        .setFontColor("#ffffff")
        .setBackground("#0b5394")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
  
      // Data
      if (data.length > 0) {
        summary.getRange(dataStart, startCol, data.length, headers.length).setValues(data);
        summary.getRange(dataStart, startCol + 1, data.length, 3).setNumberFormat("#,##0");
        summary.getRange(dataStart, startCol, data.length, headers.length)
          .setFontSize(10)
          .setFontFamily("Arial")
          .setVerticalAlignment("middle");
        
        // Left-align account/name column, right-align number columns
        summary.getRange(dataStart, startCol, data.length, 1).setHorizontalAlignment("left"); // Account/Name column
        summary.getRange(dataStart, startCol + 1, data.length, 3).setHorizontalAlignment("right"); // Number columns (Debit, Credit, Remaining)
        
        // Color code Debit and Credit columns based on table type
        const debitCol = startCol + 1; // Column C
        const creditCol = startCol + 2; // Column D
        
        // For funds and revenue/expense: Debit=orange (expenses), Credit=green (revenues)
        // For asset accounts: Debit=green (money in), Credit=orange (money out)
        const debitColor = (tableType === "fund" || tableType === "revenue_expense") ? "#fce5cd" : "#d9ead3";
        const creditColor = (tableType === "fund" || tableType === "revenue_expense") ? "#d9ead3" : "#fce5cd";
        
        summary.getRange(dataStart, debitCol, data.length, 1).setBackground(debitColor);
        summary.getRange(dataStart, creditCol, data.length, 1).setBackground(creditColor);
        
        // Add hyperlinks using the working method from the older script
        if (isFundTable) {
          // Add hyperlinks for fund names (column A)
          data.forEach((row, index) => {
            const fundName = row[0]; // Fund name is now in column A
            const fundSheetName = "Fund - " + fundName;
            const fundSheet = ss.getSheetByName(fundSheetName);
            if (fundSheet) {
              const cell = summary.getRange(dataStart + index, startCol);
              const richText = SpreadsheetApp.newRichTextValue()
                .setText(fundName)
                .setLinkUrl(`#gid=${fundSheet.getSheetId()}`)
                .build();
              cell.setRichTextValue(richText);
            }
          });
        } else {
          // Add hyperlinks for account names
          data.forEach((row, index) => {
            const accountName = row[0]; // Account name
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
              const cell = summary.getRange(dataStart + index, startCol);
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
        const totalDebit = data.reduce((s, r) => s + r[isFundTable ? 1 : 1], 0);
        const totalCredit = data.reduce((s, r) => s + r[isFundTable ? 2 : 2], 0);
        const totalDiff = isFundTable ? totalCredit - totalDebit : totalDebit - totalCredit;
        const totalRow = dataStart + data.length;
        
        const totalLabel = isFundTable ? "GRAND TOTAL (All Account)" : "TOTAL";
        const totalData = isFundTable ? 
          [totalLabel, totalDebit, totalCredit, totalDiff] :
          [totalLabel, totalDebit, totalCredit, totalDiff];
          
        summary.getRange(totalRow, startCol, 1, headers.length).setValues([totalData]);
        // Note: Background and alignment will be set separately for label vs values
        summary.getRange(totalRow, startCol, 1, headers.length)
          .setFontWeight("bold")
          .setVerticalAlignment("middle");
        summary.getRange(totalRow, startCol + 1, 1, 3).setNumberFormat("#,##0");
        
        // Style TOTAL row matching screenshot:
        // First cell (label) gets dark blue background with white text
        const totalLabelCell = summary.getRange(totalRow, startCol, 1, 1);
        totalLabelCell.setBackground("#0b5394")
          .setFontColor("#ffffff")
          .setHorizontalAlignment("left"); // left-aligned as per screenshot
        
        // Value cells (Debit, Credit, Remaining) get light yellow/golden background and right alignment
        summary.getRange(totalRow, debitCol, 1, 1)
          .setBackground("#FFF9C4")
          .setHorizontalAlignment("right"); // right-aligned for numbers
        summary.getRange(totalRow, creditCol, 1, 1)
          .setBackground("#FFF9C4")
          .setHorizontalAlignment("right"); // right-aligned for numbers
        summary.getRange(totalRow, creditCol + 1, 1, 1)
          .setBackground("#FFF9C4")
          .setHorizontalAlignment("right"); // right-aligned for numbers
  
        // Set borders matching screenshot exactly:
        // 1. Outer border - thick dark blue around entire table
        const fullTableRange = summary.getRange(tableStart, startCol, data.length + 1, headers.length);
        fullTableRange.setBorder(
          true, true, true, true, // outer: top, left, bottom, right
          false, false, // no inner borders yet (will add separately)
          "#0b5394", // dark blue
          SpreadsheetApp.BorderStyle.SOLID_THICK
        );
        
        // 2. Thick dark blue separator line below header row
        summary.getRange(tableStart + 1, startCol, 1, headers.length)
          .setBorder(true, false, false, false, false, false, // top border only (thick blue)
            "#0b5394", SpreadsheetApp.BorderStyle.SOLID_THICK);
        
        // 3. Thick dark blue separator line above TOTAL/GRAND TOTAL row
        summary.getRange(totalRow, startCol, 1, headers.length)
          .setBorder(true, false, false, false, false, false, // top border only (thick blue)
            "#0b5394", SpreadsheetApp.BorderStyle.SOLID_THICK);
        
        // 4. Thin light gray vertical lines in header row (between columns only)
        summary.getRange(tableStart, startCol, 1, headers.length)
          .setBorder(false, false, false, false, true, false, // vertical internal only
            "#d0d0d0", SpreadsheetApp.BorderStyle.SOLID);
        
        // 5. Thin light gray internal grid lines for data rows (both horizontal and vertical)
        if (data.length > 1) {
          const dataCellsRange = summary.getRange(dataStart, startCol, data.length - 1, headers.length);
          // Set horizontal lines between data rows
          for (let r = dataStart; r < totalRow - 1; r++) {
            summary.getRange(r + 1, startCol, 1, headers.length)
              .setBorder(true, false, false, false, false, false, // top border (horizontal line)
                "#d0d0d0", SpreadsheetApp.BorderStyle.SOLID);
          }
          // Set vertical lines between columns for all data rows
          dataCellsRange.setBorder(
            false, false, false, false,
            true, false, // vertical internal only
            "#d0d0d0", SpreadsheetApp.BorderStyle.SOLID
          );
        }
        
        // 6. Thin light gray vertical lines in TOTAL row (between columns only)
        summary.getRange(totalRow, startCol, 1, headers.length)
          .setBorder(false, false, false, false, true, false, // vertical internal only
            "#d0d0d0", SpreadsheetApp.BorderStyle.SOLID);
        
        // Add horizontal line below the table for section separation
        const lineRow = totalRow + 1;
        summary.getRange(lineRow, startCol, 1, headers.length)
          .setBorder(false, false, false, false, false, true, "#d0d0d0", SpreadsheetApp.BorderStyle.SOLID);
  
        return totalRow + 2; // Next table starts after some spacing
      } else {
        summary.getRange(tableStart + 1, startCol).setValue("No data available.")
          .setFontStyle("italic")
          .setFontColor("#888888");
        return tableStart + 3;
      }
    }
  
    // --- Insert VN - Master Ledger section first ---
    let nextRow = 7;
    summary.getRange(nextRow, startCol).setValue("📋 VN - MASTER LEDGER")
      .setFontWeight("bold").setFontSize(14).setFontColor("#0b5394");
    nextRow += 2;
    
    // Create table for VN - Master Ledger
    const masterTableStart = nextRow;
    const masterDataStart = nextRow + 1;
    
    // Headers
    summary.getRange(masterTableStart, startCol, 1, 1).setValues([["Dataset"]])
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontColor("#ffffff")
      .setBackground("#0b5394")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    
    // Data row - make the VN - Master Ledger text itself clickable
    const masterSheetLink = summary.getRange(masterDataStart, startCol);
    const masterRichText = SpreadsheetApp.newRichTextValue()
      .setText("VN - Master Ledger")
      .setLinkUrl(`#gid=${masterSheet.getSheetId()}`)
      .build();
    masterSheetLink.setRichTextValue(masterRichText)
      .setFontSize(11)
      .setFontColor("#0b5394")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    
    // Style the data row
    summary.getRange(masterDataStart, startCol, 1, 1)
      .setFontSize(11)
      .setFontFamily("Arial")
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
    
    // Border for the whole table
    summary.getRange(masterTableStart, startCol, 2, 1)
      .setBorder(true, true, true, true, true, true, "#bfbfbf", SpreadsheetApp.BorderStyle.SOLID);
    
    nextRow = masterDataStart + 3;
    
    // --- Insert Accounts Summary FIRST ---
    summary.getRange(nextRow, startCol).setValue("📊 ACCOUNTS SUMMARY")
      .setFontWeight("bold").setFontSize(14).setFontColor("#0b5394");
    nextRow += 2;
    
    // --- Insert Account Summary tables ---
    console.log("Expenses/Revenues:", expensesRevenues); // Debug log
    nextRow = insertTable("I. Revenues & Expenses", nextRow, expensesRevenues, false, "revenue_expense", startCol);
    nextRow = insertTable("II. Indovina Bank Accounts", nextRow, indovina, false, "asset", startCol);
    nextRow = insertTable("III. Custodian Accounts", nextRow, custodians, false, "asset", startCol);
    
    // Add spacing and section header with horizontal line separator before Fund Summary (only 2 rows total)
    nextRow += 1; // First spacing row
    summary.getRange(nextRow, startCol, 1, 10)
      .setBorder(false, false, false, false, false, true, "#b0b0b0", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    nextRow += 1; // Second spacing row - Fund Summary heading goes here
    
    // --- Insert Fund Summary AFTER Accounts Summary ---
    summary.getRange(nextRow, startCol).setValue("📊 FUND SUMMARY")
      .setFontWeight("bold").setFontSize(14).setFontColor("#0b5394");
    nextRow += 2;
    console.log("Fund Data:", fundData); // Debug log
    nextRow = insertTable("", nextRow, fundData, true, "fund", startCol); // Empty title since we have separate heading
  
    // Set proper column widths for all tables (column A is empty margin, start from B)
    summary.setColumnWidth(1, 50); // Left margin column A
    summary.setColumnWidth(2, 200); // Name/Account column B
    summary.setColumnWidth(3, 150); // Debit column C
    summary.setColumnWidth(4, 150); // Credit column D
    summary.setColumnWidth(5, 150); // Remaining column E
    summary.setColumnWidth(6, 150); // Extra column
    
    summary.setRowHeights(1, nextRow + 3, 22);
    summary.getRange(1, startCol, 3, 1).setHorizontalAlignment("left");
  
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