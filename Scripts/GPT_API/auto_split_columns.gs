/**
 * Auto-split columns script for Google Sheets
 * Automatically splits text in column A into multiple columns when new data is pasted
 * Uses vertical bar (|) as separator
 */

// Configuration - using different names to avoid conflicts
const AUTO_SPLIT_SPREADSHEET_ID = '1ZhrUxuevNYM2pD_TJ6IH_WRePKKo_QCjCPW-64hKiB4';
const AUTO_SPLIT_DATA_LOG_SHEET = 'ChatGPT_Data_Log';
const AUTO_SPLIT_COMPANY_RULES_SHEET = 'Company_Rules';

/**
 * Trigger function that runs when the sheet is edited
 */
function onEdit(e) {
  // Get the edited range
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  
  // Only process if editing the correct sheets
  if (sheetName !== AUTO_SPLIT_DATA_LOG_SHEET && sheetName !== AUTO_SPLIT_COMPANY_RULES_SHEET) {
    return;
  }
  
  // Only process if editing column A
  if (range.getColumn() !== 1) {
    return;
  }
  
  // Get the value from the cell
  const cellValue = range.getValue();
  
  // Check if the value contains vertical bar separators (|)
  if (typeof cellValue === 'string' && cellValue.includes('|')) {
    console.log('Auto-splitting text in', sheetName, 'row', range.getRow());
    
    // Split the text by vertical bar separator
    const splitValues = cellValue.split('|').map(value => value.trim());
    
    // Clear the original cell
    range.clearContent();
    
    // Set the split values across the row
    sheet.getRange(range.getRow(), 1, 1, splitValues.length).setValues([splitValues]);
  }
}

/**
 * Test function to verify auto-split functionality
 */
function testAutoSplit() {
  const testData = '2025-01-20|123|1|Direct pricing inquiry|E-commerce|Budget inquiry with value proposition|Audit pricing + Call option + Value proposition|What are your rates?||Initial pricing inquiry|Thank you for your interest in our services...';
  
  console.log('Test data:', testData);
  console.log('Split result:', testData.split('|').map(value => value.trim()));
}

/**
 * Process existing data in the sheets
 */
function splitExistingData() {
  const spreadsheet = SpreadsheetApp.openById(AUTO_SPLIT_SPREADSHEET_ID);
  
  // Process Data Log sheet
  const dataLogSheet = spreadsheet.getSheetByName(AUTO_SPLIT_DATA_LOG_SHEET);
  if (dataLogSheet) {
    const dataRange = dataLogSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 0; i < values.length; i++) {
      const cellValue = values[i][0];
      if (typeof cellValue === 'string' && cellValue.includes('|')) {
        const splitValues = cellValue.split('|').map(value => value.trim());
        dataLogSheet.getRange(i + 1, 1, 1, splitValues.length).setValues([splitValues]);
        console.log('Split row', i + 1, 'in Data Log sheet');
      }
    }
  }
  
  // Process Company Rules sheet
  const companyRulesSheet = spreadsheet.getSheetByName(AUTO_SPLIT_COMPANY_RULES_SHEET);
  if (companyRulesSheet) {
    const dataRange = companyRulesSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 0; i < values.length; i++) {
      const cellValue = values[i][0];
      if (typeof cellValue === 'string' && cellValue.includes('|')) {
        const splitValues = cellValue.split('|').map(value => value.trim());
        companyRulesSheet.getRange(i + 1, 1, 1, splitValues.length).setValues([splitValues]);
        console.log('Split row', i + 1, 'in Company Rules sheet');
      }
    }
  }
}

/**
 * Setup function to create the onEdit trigger
 */
function setupAutoSplitTrigger() {
  try {
    // Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onEdit') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create new trigger - CORRECTED SYNTAX
    ScriptApp.newTrigger('onEdit')
      .for(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
    
    console.log('Auto-split trigger created successfully');
    
  } catch (error) {
    console.error('Error setting up trigger:', error);
  }
}