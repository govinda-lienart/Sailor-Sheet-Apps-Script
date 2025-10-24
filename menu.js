// ===============================
// 🎯 MAIN MENU SYSTEM
// ===============================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create the main consolidated menu
  ui.createMenu("📊 Financial Manager")
    .addSubMenu(ui.createMenu("📒 Account Manager")
      .addItem("🔁 Full Rebuild (All Accounts)", "updateAllAccounts")
      .addItem("⚡ Quick Update (Existing Only)", "quickUpdateAccounts"))
    .addSubMenu(ui.createMenu("💰 Fund Manager")
      .addItem("🔁 Full Rebuild (All Funds)", "updateAllFunds")
      .addItem("⚡ Quick Update (Existing Funds)", "quickUpdateFunds"))
    .addSubMenu(ui.createMenu("📘 Audit Tools")
      .addItem("📊 Update Comprehensive Summary", "createOrUpdateAuditSummary")
      .addItem("🔧 Simple Summary (No Hyperlinks)", "createSimpleSummary"))
    .addSeparator()
    .addItem("❓ Help & Documentation", "showHelp")
    .addToUi();
}

// ===============================
// 📚 HELP FUNCTION
// ===============================
function showHelp() {
  const helpText = `
📊 FINANCIAL MANAGER HELP

🔹 ACCOUNT MANAGER:
• Full Rebuild: Creates/updates all account sheets (VN - Indovina Bank, VN - Revenues, VN - Expenses)
• Quick Update: Refreshes existing account sheets with new data

🔹 FUND MANAGER:
• Full Rebuild: Creates/updates all fund sheets (Unrestricted Funds, Conference Participation Fee, etc.)
• Quick Update: Refreshes existing fund sheets with new data

🔹 AUDIT TOOLS:
• Update Summary Sheet: Creates/updates audit summary reports

📋 REQUIREMENTS:
• Master sheet must be named "VN - Master Ledger"
• Required columns: Funds, Account, Debit (VND), Credit (VND)
• Data should start from row 2 (row 1 = headers)

💡 TIPS:
• Use "Full Rebuild" when adding new funds/accounts
• Use "Quick Update" for regular data updates
• Check the console for any error messages
  `;
  
  SpreadsheetApp.getUi().alert("📚 Financial Manager Help", helpText, SpreadsheetApp.getUi().ButtonSet.OK);
}