// =====================================
// 🎯 MAIN MENU SYSTEM
// ====================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create the main consolidated menu
  ui.createMenu("📊 Financial Manager")
    .addSubMenu(ui.createMenu("📒 Account Manager")
        .addItem("🔁 Full Rebuild (All Accounts)", "updateAllAccounts")
        .addItem("🔄 Update Current Sheet", "rebuildCurrentAccount"))
    .addSubMenu(ui.createMenu("💰 Fund Manager")
      .addItem("🔁 Full Rebuild (All Funds)", "updateAllFunds")
      .addItem("🔄 Update Current Sheet", "rebuildCurrentFund"))
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
• Full Rebuild (All Accounts): Creates/updates ALL account sheets at once
• Update Current Sheet: Rebuilds ONLY the sheet you're currently viewing
  (Just open any account sheet and use this option - no code editing needed!)

🔹 FUND MANAGER:
• Full Rebuild (All Funds): Creates/updates ALL fund sheets at once
• Update Current Sheet: Rebuilds ONLY the fund sheet you're currently viewing
  (Just open any fund sheet and use this option!)

🔹 AUDIT TOOLS:
• Update Summary Sheet: Creates/updates comprehensive audit summary reports

📋 REQUIREMENTS:
• Master sheet must be named "VN - Master Ledger"
• Required columns: Funds, Account, Debit (VND), Credit (VND)
• Data should start from row 2 (row 1 = headers)

💡 TIPS:
• Use "Full Rebuild" when adding new accounts or doing a complete refresh
• Use "Update Current Sheet" for quick single-account updates
• Always open the account sheet first before using "Update Current Sheet"
  `;
  
  SpreadsheetApp.getUi().alert("📚 Financial Manager Help", helpText, SpreadsheetApp.getUi().ButtonSet.OK);
}