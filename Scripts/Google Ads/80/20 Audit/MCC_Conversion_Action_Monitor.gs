// --- CONVERSION ACTION STATUS MONITOR FOR 80/20 ANALYSIS ---
// This script outputs conversion action data for a specific account
// Shows conversion action status, optimization, and source information
// Configure the account ID below to analyze a specific account

const CONVERSION_SHEET_URL = "https://docs.google.com/spreadsheets/d/1ylEHOD31fvz4LAKxCucVD-zFPWxDww0LXBFF6eTRFuM/edit?usp=sharing";
const TAB = 'Conversion_Status';
const ACCOUNT_ID = '972-837-2864'; // Target account ID for testing

// Query for conversion actions - only ENABLED ones
const CONVERSION_ACTIONS_QUERY = `
  SELECT 
    conversion_action.name,
    conversion_action.status,
    conversion_action.category,
    conversion_action.counting_type,
    conversion_action.value_settings.default_value,
    conversion_action.primary_for_goal
  FROM conversion_action
  WHERE conversion_action.status = 'ENABLED'
  ORDER BY conversion_action.name ASC
`;

// Query for conversion metrics from campaign resource
const CONVERSION_METRICS_QUERY = `
  SELECT 
    segments.conversion_action_name,
    metrics.conversions,
    metrics.conversions_value
  FROM campaign
  WHERE segments.date DURING LAST_30_DAYS
  ORDER BY metrics.conversions DESC
`;

function main() {
  try {
    Logger.log("=== Starting Conversion Action Status Monitor for 80/20 Analysis ===");
    
    // Validate account ID
    if (!ACCOUNT_ID || ACCOUNT_ID === '123-456-7890') {
      Logger.log("❌ Please set the ACCOUNT_ID constant to your target account ID");
      Logger.log("Example: const ACCOUNT_ID = '123-456-7890';");
      return;
    }
    
    // Switch to the specified account
    Logger.log(`🔍 Switching to account: ${ACCOUNT_ID}`);
    const account = getAccountById(ACCOUNT_ID);
    if (!account) {
      Logger.log(`❌ Account ${ACCOUNT_ID} not found or not accessible`);
      return;
    }
    
    // Switch to account context
    AdsManagerApp.select(account);
    Logger.log(`✓ Successfully switched to account: ${account.getName()}`);
    
    // Open or create spreadsheet
    let ss;
    if (!CONVERSION_SHEET_URL) {
      ss = SpreadsheetApp.create("Conversion Action Status Monitor");
      let url = ss.getUrl();
      Logger.log("No CONVERSION_SHEET_URL found, so this sheet was created: " + url);
    } else {
      ss = SpreadsheetApp.openByUrl(CONVERSION_SHEET_URL);
    }
    
    // Process conversion actions
    Logger.log("=== Processing Conversion Actions ===");
    processConversionActions(ss, account);
    
    Logger.log(`✓ Account analyzed: ${account.getName()} (${ACCOUNT_ID})`);
    Logger.log(`✓ Data exported to tab: ${TAB}`);
    
  } catch (error) {
    Logger.log(`❌ Error in main function: ${error.message}`);
    Logger.log(`Error details: ${error.stack || 'No stack trace available'}`);
  }
}

function processConversionActions(ss, account) {
  try {
    // Get or create the conversion actions tab
    let sheet = ss.getSheetByName(TAB);
    if (!sheet) {
      sheet = ss.insertSheet(TAB);
      Logger.log(`Created new tab: ${TAB}`);
    } else {
      sheet.clear();
      Logger.log(`Cleared existing data in tab: ${TAB}`);
    }
    
    // Execute conversion actions query
    Logger.log("Executing conversion actions query...");
    const rows = AdsApp.search(CONVERSION_ACTIONS_QUERY);
    
    // Log sample row structure for debugging
    const sampleQuery = CONVERSION_ACTIONS_QUERY + ' LIMIT 1';
    const sampleRows = AdsApp.search(sampleQuery);
    
    if (sampleRows.hasNext()) {
      const sampleRow = sampleRows.next();
      Logger.log("Sample conversion action row structure: " + JSON.stringify(sampleRow));
    } else {
      Logger.log("No sample rows found for conversion actions");
    }
    
    // Get conversion metrics data
    Logger.log("Getting conversion metrics from campaign data...");
    const metricsRows = AdsApp.search(CONVERSION_METRICS_QUERY);
    
    // Process the data
    const data = calculateConversionActionMetrics(rows, metricsRows);
    
    if (data.length === 0) {
      Logger.log("No conversion action data found to export");
      return;
    }
    
    // Create headers for conversion actions
    const headers = [
      'Conversion Action Name',
      'Status',
      'Category',
      'Counting Type',
      'Default Value',
      'Action Optimization',
      'Conversion Source',
      'Conversions (30 days)',
      'Conversion Value (30 days)'
    ];
    
    // Write headers and data to sheet
    const allData = [headers, ...data];
    const range = sheet.getRange(1, 1, allData.length, headers.length);
    range.setValues(allData);
    
    Logger.log(`✓ Successfully exported ${data.length} conversion action records to ${TAB} tab`);
    
  } catch (error) {
    Logger.log(`❌ Error processing conversion actions: ${error.message}`);
  }
}

function getAccountById(accountId) {
  try {
    Logger.log(`🔍 Looking for account ID: ${accountId}`);
    const accountIterator = AdsManagerApp.accounts().withIds([accountId]).get();
    if (accountIterator.hasNext()) {
      const account = accountIterator.next();
      Logger.log(`✓ Found account: ${account.getName()}`);
      return account;
    }
    Logger.log(`❌ Account ${accountId} not found in MCC`);
    return null;
  } catch (error) {
    Logger.log(`❌ Error finding account ${accountId}: ${error.message}`);
    return null;
  }
}

function calculateConversionActionMetrics(conversionRows, metricsRows) {
  let data = [];
  let totalRows = 0;
  
  // First, build a map of conversion metrics by action name
  const metricsMap = {};
  while (metricsRows.hasNext()) {
    const metricsRow = metricsRows.next();
    const actionName = metricsRow.segments.conversionActionName;
    const conversions = Number(metricsRow.metrics.conversions) || 0;
    const conversionValue = Number(metricsRow.metrics.conversionsValue) || 0;
    
    if (actionName) {
      metricsMap[actionName] = {
        conversions: conversions,
        conversionValue: conversionValue
      };
    }
  }
  
  Logger.log(`Built metrics map with ${Object.keys(metricsMap).length} conversion actions`);
  
  // Process conversion action details
  while (conversionRows.hasNext()) {
    try {
      const row = conversionRows.next();
      totalRows++;
      
      const conversionAction = row.conversionAction || {};
      
      const conversionActionName = conversionAction.name || 'N/A';
      const apiStatus = conversionAction.status || 'N/A';
      const category = conversionAction.category || 'N/A';
      const countingType = conversionAction.countingType || 'N/A';
      const defaultValue = conversionAction.valueSettings ? conversionAction.valueSettings.defaultValue || 'N/A' : 'N/A';
      const primaryForGoal = conversionAction.primaryForGoal || false;
      
      // Get conversion metrics from the map
      const metrics = metricsMap[conversionActionName] || { conversions: 0, conversionValue: 0 };
      
      // Determine action optimization and conversion source
      const actionOptimization = primaryForGoal ? 'Primary' : 'Secondary';
      
      // Determine conversion source based on the conversion action name and category
      let conversionSource = 'Website';
      if (conversionActionName && conversionActionName.toLowerCase().includes('google analytics')) {
        conversionSource = 'Website (Google Analytics (GA4))';
      } else if (category === 'PHONE_CALL' || conversionActionName.toLowerCase().includes('call')) {
        conversionSource = 'Call from Ads';
      } else if (category === 'PURCHASE') {
        conversionSource = 'Website';
      } else if (category === 'SIGNUP' || category === 'LEAD') {
        conversionSource = 'Website';
      }
      
      // Map API status to Google Ads display status
      let displayStatus = 'Active'; // Default to Active for ENABLED conversion actions
      
      if (apiStatus === 'ENABLED') {
        // Check if it's a Google Analytics conversion action (likely to be "Inactive" if no recent data)
        if (conversionActionName && conversionActionName.toLowerCase().includes('google analytics')) {
          displayStatus = 'Inactive'; // Google Analytics conversions often show as Inactive if not properly configured
        } else if (category === 'PHONE_CALL' || conversionActionName.toLowerCase().includes('call')) {
          displayStatus = 'No recent conversions'; // Phone calls often show this status
        } else {
          displayStatus = 'Active'; // Most website conversions show as Active
        }
      } else if (apiStatus === 'REMOVED') {
        displayStatus = 'Inactive';
      }
      
      Logger.log(`Row ${totalRows} data: Name=${conversionActionName}, DisplayStatus=${displayStatus}, Conversions=${metrics.conversions}, Value=${metrics.conversionValue}`);
      
      // Add to data array
      data.push([
        conversionActionName,
        displayStatus, // Use display status instead of API status
        category,
        countingType,
        defaultValue,
        actionOptimization,
        conversionSource,
        metrics.conversions,
        metrics.conversionValue
      ]);
      
    } catch (error) {
      Logger.log(`Error processing row: ${error.message}`);
    }
  }
  
  Logger.log(`✓ Processed ${totalRows} conversion action rows`);
  
  return data;
}

function testScript() {
  Logger.log("=== Testing Conversion Action Status Monitor ===");
  main();
}
