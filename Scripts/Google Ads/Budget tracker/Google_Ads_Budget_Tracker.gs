const SHEET_URL = 'https://docs.google.com/spreadsheets/d/18PZxaSAwZBv32yCgYkxT-LW5o4hKHNDCfXR51w2ceI0/edit?usp=sharing';

// GAQL query to get total spend for current month across all campaigns
const SPEND_QUERY = `
SELECT 
    metrics.cost_micros
FROM customer
WHERE segments.date DURING THIS_MONTH
`;

// GAQL query to get all campaign daily budgets
const BUDGET_QUERY = `
SELECT 
    campaign.id,
    campaign.name,
    campaign.daily_budget_micros,
    campaign.status
FROM campaign
WHERE campaign.status = 'ENABLED'
`;

function main() {
    try {
        Logger.log("=== Starting Google Ads Budget Tracker ===");
        
        // Test spreadsheet access first
        Logger.log("Testing spreadsheet access...");
        const spreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
        Logger.log("✓ Successfully accessed spreadsheet: " + spreadsheet.getName());
        
        // Get total spend for current month
        Logger.log("Getting total spend for current month...");
        const totalSpend = getTotalSpendForMonth();
        Logger.log("Total spend for current month: $" + totalSpend.toFixed(2));
        
        // Calculate average daily spend
        Logger.log("Calculating average daily spend...");
        const averageDailySpend = calculateAverageDailySpend(totalSpend);
        Logger.log("Average daily spend: $" + averageDailySpend.toFixed(2));
        
        // Get sum of all campaign daily budgets
        Logger.log("Getting total daily budget...");
        const totalDailyBudget = getTotalDailyBudget();
        Logger.log("Total daily budget: $" + totalDailyBudget.toFixed(2));
        
        // Update spreadsheet
        Logger.log("Updating spreadsheet...");
        updateSpreadsheet(totalSpend, averageDailySpend, totalDailyBudget);
        
        Logger.log("=== Budget tracking completed successfully! ===");
        
    } catch (error) {
        Logger.log("❌ ERROR in main function: " + error.toString());
        Logger.log("Stack trace: " + error.stack);
        throw error;
    }
}

function getTotalSpendForMonth() {
    let totalCostMicros = 0;
    
    try {
        Logger.log("Executing spend query...");
        const rows = AdsApp.search(SPEND_QUERY);
        
        if (rows.hasNext()) {
            const row = rows.next();
            Logger.log("Sample row structure: " + JSON.stringify(row));
            
            const metrics = row.metrics || {};
            Logger.log("Metrics object: " + JSON.stringify(metrics));
            
            const costMicros = Number(metrics.costMicros) || 0;
            totalCostMicros = costMicros;
            Logger.log("Cost in micros: " + costMicros);
        } else {
            Logger.log("⚠️ No data returned from spend query");
        }
        
        // Convert micros to dollars
        const totalSpend = totalCostMicros / 1000000;
        Logger.log("Converted to dollars: $" + totalSpend.toFixed(2));
        return totalSpend;
        
    } catch (error) {
        Logger.log("❌ Error getting total spend: " + error.toString());
        return 0;
    }
}

function calculateAverageDailySpend(totalSpend) {
    try {
        // Get current date
        const now = new Date();
        const currentDay = now.getDate();
        Logger.log("Current day of month: " + currentDay);
        
        // Calculate average daily spend based on current day of month
        const averageDailySpend = totalSpend / currentDay;
        Logger.log("Average daily spend calculation: $" + totalSpend.toFixed(2) + " / " + currentDay + " = $" + averageDailySpend.toFixed(2));
        return averageDailySpend;
        
    } catch (error) {
        Logger.log("❌ Error calculating average daily spend: " + error.toString());
        return 0;
    }
}

function getTotalDailyBudget() {
    let totalBudgetMicros = 0;
    let campaignCount = 0;
    let totalCampaignsProcessed = 0;
    
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
            totalCampaignsProcessed++;
            
            const campaign = row.campaign || {};
            const campaignBudget = row.campaignBudget || {};
            const campaignName = campaign.name || "Unknown Campaign";
            const campaignStatus = campaign.status || "UNKNOWN";
            const experimentType = campaign.experimentType || "NONE";
            const budgetMicros = Number(campaignBudget.amountMicros) || 0;
            
            // Log all campaigns for debugging
            Logger.log("Campaign " + totalCampaignsProcessed + ": " + campaignName + " (Status: " + campaignStatus + ", Experiment: " + experimentType + ")");
            
            // Skip only EXPERIMENT campaigns to avoid double counting budgets
            // BASE campaigns are the original campaigns and should be counted
            if (experimentType === "EXPERIMENT") {
                Logger.log("  - Skipping (experiment campaign)");
                continue;
            }
            
            // Only process ENABLED campaigns
            if (campaignStatus !== "ENABLED") {
                Logger.log("  - Skipping (not enabled)");
                continue;
            }
            
            campaignCount++;
            
            // Get budget directly from GAQL result
            const dailyBudgetAmount = budgetMicros / 1000000; // Convert micros to dollars
            
            // Log budget details for debugging
            Logger.log("  - Status: " + campaignStatus);
            Logger.log("  - Budget Amount: $" + dailyBudgetAmount.toFixed(2));
            Logger.log("  - Counting as daily budget: $" + dailyBudgetAmount.toFixed(2));
            
            totalBudgetMicros += budgetMicros;
        }
        
        Logger.log("Total campaigns processed: " + totalCampaignsProcessed);
        Logger.log("Total ACTIVE campaigns processed: " + campaignCount);
        
        // Convert micros to dollars
        const totalDailyBudget = totalBudgetMicros / 1000000;
        Logger.log("Total daily budget (ACTIVE campaigns only): $" + totalDailyBudget.toFixed(2));
        return totalDailyBudget;
        
    } catch (error) {
        Logger.log("❌ Error getting total daily budget: " + error.toString());
        return 0;
    }
}

function updateSpreadsheet(totalSpend, averageDailySpend, totalDailyBudget) {
    try {
        Logger.log("Opening spreadsheet...");
        const spreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
        Logger.log("✓ Spreadsheet opened: " + spreadsheet.getName());
        
        // Find the first visible tab (the one that shows up first)
        const sheets = spreadsheet.getSheets();
        Logger.log("Total sheets in spreadsheet: " + sheets.length);
        
        let firstVisibleSheet = null;
        let firstVisibleSheetName = "";
        
        // Log all sheet names and find the first visible one
        Logger.log("All sheet names:");
        for (let i = 0; i < sheets.length; i++) {
            const sheet = sheets[i];
            const sheetName = sheet.getName();
            Logger.log("  " + (i + 1) + ". " + sheetName + (sheet.isSheetHidden() ? " (HIDDEN)" : " (VISIBLE)"));
            
            // Find the first visible sheet from the beginning (first visible)
            if (!sheet.isSheetHidden() && !firstVisibleSheet) {
                firstVisibleSheet = sheet;
                firstVisibleSheetName = sheetName;
            }
        }
        
        if (!firstVisibleSheet) {
            throw new Error("No visible sheets found in spreadsheet");
        }
        
        Logger.log("✓ First visible sheet: " + firstVisibleSheetName);
        
        // Format values as currency
        const formattedTotalSpend = Utilities.formatString("$%.2f", totalSpend);
        const formattedAverageDailySpend = Utilities.formatString("$%.2f", averageDailySpend);
        const formattedTotalDailyBudget = Utilities.formatString("$%.2f", totalDailyBudget);
        
        Logger.log("Formatted values:");
        Logger.log("  Total Spend: " + formattedTotalSpend);
        Logger.log("  Avg Daily Spend: " + formattedAverageDailySpend);
        Logger.log("  Total Daily Budget: " + formattedTotalDailyBudget);
        
        // Test cell access before updating
        Logger.log("Testing cell access...");
        const testCell = firstVisibleSheet.getRange("A1");
        Logger.log("✓ Can access cell A1: " + testCell.getValue());
        
        // Update cells
        Logger.log("Updating F8...");
        firstVisibleSheet.getRange("F8").setValue(formattedTotalSpend);
        Logger.log("✓ F8 updated");
        
        Logger.log("Updating G8...");
        firstVisibleSheet.getRange("G8").setValue(formattedAverageDailySpend);
        Logger.log("✓ G8 updated");
        
        Logger.log("Updating G6...");
        firstVisibleSheet.getRange("G6").setValue(formattedTotalDailyBudget);
        Logger.log("✓ G6 updated");
        
        // Verify the updates
        Logger.log("Verifying updates...");
        const verifyF8 = firstVisibleSheet.getRange("F8").getValue();
        const verifyG8 = firstVisibleSheet.getRange("G8").getValue();
        const verifyG6 = firstVisibleSheet.getRange("G6").getValue();
        
        Logger.log("✓ Verification - F8: " + verifyF8 + ", G8: " + verifyG8 + ", G6: " + verifyG6);
        
        Logger.log("=== Spreadsheet updated successfully! ===");
        
    } catch (error) {
        Logger.log("❌ Error updating spreadsheet: " + error.toString());
        Logger.log("Error details: " + error.message);
        throw error;
    }
}

// Test function to verify spreadsheet access
function testSpreadsheetAccess() {
    try {
        Logger.log("Testing spreadsheet access...");
        const spreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
        Logger.log("✓ Successfully accessed spreadsheet: " + spreadsheet.getName());
        
        const sheets = spreadsheet.getSheets();
        Logger.log("✓ Found " + sheets.length + " sheets");
        
        for (let i = 0; i < sheets.length; i++) {
            Logger.log("Sheet " + (i + 1) + ": " + sheets[i].getName());
        }
        
        const lastSheet = sheets[sheets.length - 1];
        Logger.log("✓ Most recent sheet: " + lastSheet.getName());
        
        // Test writing to a cell
        const testValue = "Test " + new Date().toLocaleString();
        lastSheet.getRange("A1").setValue(testValue);
        Logger.log("✓ Successfully wrote test value to A1: " + testValue);
        
        return true;
        
    } catch (error) {
        Logger.log("❌ Spreadsheet access test failed: " + error.toString());
        return false;
    }
}
