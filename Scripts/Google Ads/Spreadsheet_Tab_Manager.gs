/**
 * Google Apps Script for Spreadsheet Tab Management
 * This script provides utility functions for managing tabs and can be used
 * in conjunction with the Google Ads Budget Tracker script.
 */

// Configuration
const SPREADSHEET_ID = '18PZxaSAwZBv32yCgYkxT-LW5o4hKHNDCfXR51w2ceI0';

/**
 * Get the most recent tab in the spreadsheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The most recent sheet
 */
function getMostRecentTab() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = spreadsheet.getSheets();
  return sheets[sheets.length - 1];
}

/**
 * Get the name of the most recent tab
 * @returns {string} The name of the most recent sheet
 */
function getMostRecentTabName() {
  const mostRecentSheet = getMostRecentTab();
  return mostRecentSheet.getName();
}

/**
 * Create a new weekly tab with the current date
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The newly created sheet
 */
function createWeeklyTab() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Create tab name with current date (e.g., "Week of 8/12/25")
  const today = new Date();
  const weekStart = getWeekStart(today);
  const tabName = Utilities.formatDate(weekStart, Session.getScriptTimeZone(), "M/d/yy");
  
  // Create new sheet
  const newSheet = spreadsheet.insertSheet(tabName);
  
  // Copy template from first sheet if it exists
  const firstSheet = spreadsheet.getSheets()[0];
  if (firstSheet) {
    copySheetFormatting(firstSheet, newSheet);
  }
  
  Logger.log("Created new weekly tab: " + tabName);
  return newSheet;
}

/**
 * Get the start of the week (Sunday) for a given date
 * @param {Date} date - The date to get the week start for
 * @returns {Date} The start of the week
 */
function getWeekStart(date) {
  const day = date.getDay();
  const diff = date.getDate() - day;
  return new Date(date.setDate(diff));
}

/**
 * Copy formatting from one sheet to another
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet - Source sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet - Target sheet
 */
function copySheetFormatting(sourceSheet, targetSheet) {
  try {
    // Copy headers and basic formatting
    const headerRange = sourceSheet.getRange("A1:AB1");
    const targetRange = targetSheet.getRange("A1:AB1");
    
    // Copy values and formatting
    targetRange.setValues(headerRange.getValues());
    targetRange.setFontWeight("bold");
    targetRange.setBackground("#f3f3f3");
    
    // Set column widths
    for (let i = 1; i <= 28; i++) {
      const columnWidth = sourceSheet.getColumnWidth(i);
      targetSheet.setColumnWidth(i, columnWidth);
    }
    
    Logger.log("Copied formatting from template sheet");
  } catch (error) {
    Logger.log("Error copying formatting: " + error.toString());
  }
}

/**
 * Update budget tracking cells in the most recent tab
 * This function can be called from the Google Ads script
 * @param {number} totalSpend - Total spend for the month
 * @param {number} averageDailySpend - Average daily spend
 * @param {number} totalDailyBudget - Total daily budget
 */
function updateBudgetCells(totalSpend, averageDailySpend, totalDailyBudget) {
  try {
    const mostRecentSheet = getMostRecentTab();
    
    // Format values as currency
    const formattedTotalSpend = Utilities.formatString("$%.2f", totalSpend);
    const formattedAverageDailySpend = Utilities.formatString("$%.2f", averageDailySpend);
    const formattedTotalDailyBudget = Utilities.formatString("$%.2f", totalDailyBudget);
    
    // Update cells
    mostRecentSheet.getRange("F8").setValue(formattedTotalSpend);
    mostRecentSheet.getRange("G8").setValue(formattedAverageDailySpend);
    mostRecentSheet.getRange("G6").setValue(formattedTotalDailyBudget);
    
    Logger.log("Updated budget cells in sheet: " + mostRecentSheet.getName());
    Logger.log("F8 (Total Spend): " + formattedTotalSpend);
    Logger.log("G8 (Avg Daily Spend): " + formattedAverageDailySpend);
    Logger.log("G6 (Total Daily Budget): " + formattedTotalDailyBudget);
    
  } catch (error) {
    Logger.log("Error updating budget cells: " + error.toString());
    throw error;
  }
}

/**
 * Test function to verify the script is working
 */
function testTabManager() {
  try {
    const mostRecentTab = getMostRecentTab();
    Logger.log("Most recent tab: " + mostRecentTab.getName());
    
    // Test updating some sample values
    updateBudgetCells(1234.56, 45.67, 89.12);
    
    Logger.log("Test completed successfully!");
  } catch (error) {
    Logger.log("Test failed: " + error.toString());
  }
}
