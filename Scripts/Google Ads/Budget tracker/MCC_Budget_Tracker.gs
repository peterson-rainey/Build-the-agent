// --- MCC BUDGET TRACKER CONFIGURATION ---

// 1. MASTER SPREADSHEET URL
//    - This spreadsheet contains the mapping of Account IDs to their individual spreadsheet URLs
//    - Format: Column A = Client Name, Column B = Account ID (10-digit), Column C = Spreadsheet URL
const MASTER_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1bS1xmTFB0LWGvYeZBPcykGMY0SwZ4W-Ny6aKVloy0SM/edit?usp=sharing";

// 2. MASTER SPREADSHEET TAB NAME
//    - The tab name in the master spreadsheet that contains the account mappings
const MASTER_SHEET_NAME = "AccountMappings"; // Default tab name

// 3. TEST CID (Optional)
//    - To run for a single account for testing, enter its Customer ID (e.g., "123-456-7890").
//    - Leave blank ("") to run for all accounts in the master spreadsheet.
const SINGLE_CID_FOR_TESTING = ""; // Example: "123-456-7890" or ""

// 4. CELL LOCATIONS FOR BUDGET DATA
//    - These are the cells where budget data will be written in each account's spreadsheet
const BUDGET_CELLS = {
  TOTAL_SPEND: "F8",
  AVERAGE_DAILY_SPEND: "G8", 
  TOTAL_DAILY_BUDGET: "G6"
};

// --- END OF CONFIGURATION ---

function main() {
  Logger.log("=== Starting MCC Budget Tracker ===");
  
  if (SINGLE_CID_FOR_TESTING) {
    Logger.log(`Running in test mode for CID: ${SINGLE_CID_FOR_TESTING}`);
  } else {
    Logger.log("Running for all accounts in the master spreadsheet.");
  }

  // Get account mappings from master spreadsheet
  const accountMappings = getAccountMappings();
  if (!accountMappings || accountMappings.length === 0) {
    Logger.log("❌ No account mappings found in master spreadsheet.");
    return;
  }

  Logger.log(`Found ${accountMappings.length} account mappings in master spreadsheet.`);

  let processedCount = 0;
  let successCount = 0;
  let errorCount = 0;

  // Process each account
  for (const mapping of accountMappings) {
    const accountId = mapping.accountId;
    const spreadsheetUrl = mapping.spreadsheetUrl;
    const accountName = mapping.accountName || "Unknown";

    // Skip if testing mode and this isn't the test account
    if (SINGLE_CID_FOR_TESTING && accountId !== SINGLE_CID_FOR_TESTING) {
      continue;
    }

    processedCount++;
    Logger.log(`\n--- Processing Account ${processedCount}/${accountMappings.length} ---`);
    Logger.log(`Account: ${accountName} (${accountId})`);
    Logger.log(`Spreadsheet: ${spreadsheetUrl}`);

    try {
      // Switch to the account context
      const account = getAccountById(accountId);
      if (!account) {
        Logger.log(`❌ Account ${accountId} not found or not accessible.`);
        errorCount++;
        continue;
      }

      AdsManagerApp.select(account);
      Logger.log(`✓ Switched to account: ${account.getName()}`);

      // Run budget tracker for this account
      const budgetData = getBudgetDataForAccount();
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, budgetData);
      
      Logger.log(`✓ Successfully updated budget data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC BUDGET TRACKER COMPLETED ===`);
  Logger.log(`Total accounts processed: ${processedCount}`);
  Logger.log(`Successful updates: ${successCount}`);
  Logger.log(`Errors: ${errorCount}`);
}

function getAccountMappings() {
  try {
    Logger.log("Reading account mappings from master spreadsheet...");
    
    const masterSpreadsheet = SpreadsheetApp.openByUrl(MASTER_SPREADSHEET_URL);
    const masterSheet = masterSpreadsheet.getSheetByName(MASTER_SHEET_NAME);
    
    if (!masterSheet) {
      throw new Error(`Tab "${MASTER_SHEET_NAME}" not found in master spreadsheet`);
    }

    const data = masterSheet.getDataRange().getValues();
    const mappings = [];

    // Skip header row, process data rows
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const accountName = row[0]?.toString().trim() || "Unknown";
      const accountId = row[1]?.toString().trim();
      const spreadsheetUrl = row[2]?.toString().trim();

      if (accountId && spreadsheetUrl) {
        mappings.push({
          accountId: accountId,
          spreadsheetUrl: spreadsheetUrl,
          accountName: accountName
        });
      }
    }

    Logger.log(`✓ Found ${mappings.length} valid account mappings`);
    return mappings;

  } catch (error) {
    Logger.log(`❌ Error reading master spreadsheet: ${error.message}`);
    throw error;
  }
}

function getAccountById(accountId) {
  try {
    const accountIterator = AdsManagerApp.accounts().withIds([accountId]).get();
    if (accountIterator.hasNext()) {
      return accountIterator.next();
    }
    return null;
  } catch (error) {
    Logger.log(`Error finding account ${accountId}: ${error.message}`);
    return null;
  }
}

function getBudgetDataForAccount() {
  Logger.log("Getting budget data for current account...");
  
  try {
    // Get total spend for current month
    const totalSpend = getTotalSpendForMonth();
    Logger.log(`Total spend for current month: $${totalSpend.toFixed(2)}`);
    
    // Calculate average daily spend
    const averageDailySpend = calculateAverageDailySpend(totalSpend);
    Logger.log(`Average daily spend: $${averageDailySpend.toFixed(2)}`);
    
    // Get sum of all campaign daily budgets
    const totalDailyBudget = getTotalDailyBudget();
    Logger.log(`Total daily budget: $${totalDailyBudget.toFixed(2)}`);
    
    return {
      totalSpend: totalSpend,
      averageDailySpend: averageDailySpend,
      totalDailyBudget: totalDailyBudget
    };
    
  } catch (error) {
    Logger.log(`❌ Error getting budget data: ${error.message}`);
    throw error;
  }
}

function getTotalSpendForMonth() {
  let totalCostMicros = 0;
  
  try {
    const query = `
      SELECT 
          metrics.cost_micros
      FROM customer
      WHERE segments.date DURING THIS_MONTH
    `;
    
    const rows = AdsApp.search(query);
    
    if (rows.hasNext()) {
      const row = rows.next();
      const metrics = row.metrics || {};
      const costMicros = Number(metrics.costMicros) || 0;
      totalCostMicros = costMicros;
    }
    
    // Convert micros to dollars
    const totalSpend = totalCostMicros / 1000000;
    return totalSpend;
    
  } catch (error) {
    Logger.log(`❌ Error getting total spend: ${error.message}`);
    return 0;
  }
}

function calculateAverageDailySpend(totalSpend) {
  try {
    // Get current date
    const now = new Date();
    const currentDay = now.getDate();
    
    // Calculate average daily spend based on current day of month
    const averageDailySpend = totalSpend / currentDay;
    return averageDailySpend;
    
  } catch (error) {
    Logger.log(`❌ Error calculating average daily spend: ${error.message}`);
    return 0;
  }
}

function getTotalDailyBudget() {
  let totalBudgetMicros = 0;
  let campaignCount = 0;
  
  try {
    Logger.log("Getting daily budgets using GAQL to find ALL campaigns...");
    
    // Use GAQL to get ALL campaigns including experiments and special states
    const query = `
    SELECT 
        campaign.id,
        campaign.name,
        campaign.status,
        campaign.experiment_type,
        campaign_budget.amount_micros
    FROM campaign
    `;
    
    const rows = AdsApp.search(query);
    
    while (rows.hasNext()) {
      const row = rows.next();
      const campaign = row.campaign || {};
      const campaignBudget = row.campaignBudget || {};
      const campaignName = campaign.name || "Unknown Campaign";
      const campaignStatus = campaign.status || "UNKNOWN";
      const experimentType = campaign.experimentType || "NONE";
      const budgetMicros = Number(campaignBudget.amountMicros) || 0;
      
      // Skip only EXPERIMENT campaigns to avoid double counting budgets
      // BASE campaigns are the original campaigns and should be counted
      if (experimentType === "EXPERIMENT") {
        continue;
      }
      
      // Only process ENABLED campaigns
      if (campaignStatus !== "ENABLED") {
        continue;
      }
      
      campaignCount++;
      
      // Get budget directly from GAQL result
      const dailyBudgetAmount = budgetMicros / 1000000; // Convert micros to dollars
      
      Logger.log(`  Campaign ${campaignCount}: ${campaignName} - Budget: $${dailyBudgetAmount.toFixed(2)}`);
      
      totalBudgetMicros += budgetMicros;
    }
    
    Logger.log(`Total ACTIVE campaigns processed: ${campaignCount}`);
    
    // Convert micros to dollars
    const totalDailyBudget = totalBudgetMicros / 1000000;
    Logger.log(`Total daily budget (ACTIVE campaigns only): $${totalDailyBudget.toFixed(2)}`);
    return totalDailyBudget;
    
  } catch (error) {
    Logger.log(`❌ Error getting total daily budget: ${error.message}`);
    return 0;
  }
}

function updateAccountSpreadsheet(spreadsheetUrl, budgetData) {
  try {
    Logger.log("Updating account spreadsheet...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Find the first visible tab (the one that shows up first)
    const sheets = spreadsheet.getSheets();
    let firstVisibleSheet = null;
    let firstVisibleSheetName = "";
    
    // Find the first visible sheet from the beginning (first visible)
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const sheetName = sheet.getName();
      
      // Find the first visible sheet from the beginning (first visible)
      if (!sheet.isSheetHidden() && !firstVisibleSheet) {
        firstVisibleSheet = sheet;
        firstVisibleSheetName = sheetName;
      }
    }
    
    if (!firstVisibleSheet) {
      throw new Error("No visible sheets found in spreadsheet");
    }
    
    Logger.log(`✓ Updating sheet: ${firstVisibleSheetName}`);
    
    // Format values as currency
    const formattedTotalSpend = Utilities.formatString("$%.2f", budgetData.totalSpend);
    const formattedAverageDailySpend = Utilities.formatString("$%.2f", budgetData.averageDailySpend);
    const formattedTotalDailyBudget = Utilities.formatString("$%.2f", budgetData.totalDailyBudget);
    
    // Update cells
    firstVisibleSheet.getRange(BUDGET_CELLS.TOTAL_SPEND).setValue(formattedTotalSpend);
    firstVisibleSheet.getRange(BUDGET_CELLS.AVERAGE_DAILY_SPEND).setValue(formattedAverageDailySpend);
    firstVisibleSheet.getRange(BUDGET_CELLS.TOTAL_DAILY_BUDGET).setValue(formattedTotalDailyBudget);
    
    Logger.log(`✓ Successfully updated budget cells:`);
    Logger.log(`  ${BUDGET_CELLS.TOTAL_SPEND}: ${formattedTotalSpend}`);
    Logger.log(`  ${BUDGET_CELLS.AVERAGE_DAILY_SPEND}: ${formattedAverageDailySpend}`);
    Logger.log(`  ${BUDGET_CELLS.TOTAL_DAILY_BUDGET}: ${formattedTotalDailyBudget}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}
