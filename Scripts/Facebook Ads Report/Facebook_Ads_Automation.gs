/**
 * Facebook Ads Data Automation Script
 * Automatically exports Facebook Ads data to Google Sheets daily
 */

// Configuration
const CONFIG = {
  FACEBOOK_APP_ID: 'YOUR_FACEBOOK_APP_ID',
  FACEBOOK_APP_SECRET: 'YOUR_FACEBOOK_APP_SECRET',
  ACCESS_TOKEN: 'YOUR_FACEBOOK_ACCESS_TOKEN',
  AD_ACCOUNT_ID: 'act_YOUR_AD_ACCOUNT_ID',
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID', // Get this from your Google Sheet URL
  RAW_DATA_SHEET: 'Raw Data'
};

/**
 * Main automation function - runs daily
 */
function dailyFacebookAdsExport() {
  try {
    console.log('Starting daily Facebook Ads export...');
    
    // Get the target spreadsheet
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Export data from Facebook Ads API
    const adsData = exportFacebookAdsData();
    
    // Write data to Google Sheets
    writeDataToSheet(spreadsheet, adsData);
    
    // Process the data (run your existing processing script)
    processExportedData(spreadsheet);
    
    console.log('Daily Facebook Ads export completed successfully');
    
  } catch (error) {
    console.error('Error in daily Facebook Ads export:', error);
    sendErrorNotification(error);
  }
}

/**
 * Export data from Facebook Ads API
 */
function exportFacebookAdsData() {
  try {
    console.log('Exporting data from Facebook Ads API...');
    
    // Get yesterday's date (for daily export)
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const dateString = yesterday.toISOString().split('T')[0];
    
    // Facebook Ads API endpoint
    const apiUrl = `https://graph.facebook.com/v18.0/${CONFIG.AD_ACCOUNT_ID}/insights`;
    
    // API parameters
    const params = {
      access_token: CONFIG.ACCESS_TOKEN,
      fields: 'campaign_name,adset_name,date_start,date_stop,impressions,clicks,spend,reach,frequency,results,result_type',
      date_preset: 'yesterday',
      time_increment: 1,
      level: 'ad'
    };
    
    // Make API request
    const response = UrlFetchApp.fetch(apiUrl + '?' + Object.keys(params).map(key => key + '=' + encodeURIComponent(params[key])).join('&'));
    const data = JSON.parse(response.getContentText());
    
    if (data.error) {
      throw new Error('Facebook API Error: ' + data.error.message);
    }
    
    console.log(`Exported ${data.data.length} records from Facebook Ads API`);
    return data.data;
    
  } catch (error) {
    console.error('Error exporting Facebook Ads data:', error);
    throw error;
  }
}

/**
 * Write exported data to Google Sheets
 */
function writeDataToSheet(spreadsheet, data) {
  try {
    console.log('Writing data to Google Sheets...');
    
    let sheet = spreadsheet.getSheetByName(CONFIG.RAW_DATA_SHEET);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.RAW_DATA_SHEET);
    }
    
    // Clear existing data
    sheet.clear();
    
    // Prepare headers
    const headers = [
      'Campaign name',
      'Ad set name', 
      'Date',
      'Result type',
      'Results',
      'Reach',
      'Impressions',
      'Spend',
      'Link clicks',
      'Frequency'
    ];
    
    // Write headers
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // Prepare data rows
    const rows = data.map(record => [
      record.campaign_name || '',
      record.adset_name || '',
      record.date_start || '',
      record.result_type || '',
      record.results || 0,
      record.reach || 0,
      record.impressions || 0,
      record.spend || 0,
      record.clicks || 0,
      record.frequency || 0
    ]);
    
    // Write data
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
    
    console.log(`Wrote ${rows.length} rows to Google Sheets`);
    
  } catch (error) {
    console.error('Error writing data to sheet:', error);
    throw error;
  }
}

/**
 * Process the exported data using existing script
 */
function processExportedData(spreadsheet) {
  try {
    console.log('Processing exported data...');
    
    // Call your existing processing function
    enhancedManualProcess();
    
    console.log('Data processing completed');
    
  } catch (error) {
    console.error('Error processing exported data:', error);
    throw error;
  }
}

/**
 * Set up daily trigger
 */
function setupDailyTrigger() {
  try {
    // Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'dailyFacebookAdsExport') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create new daily trigger (runs at 2 AM)
    ScriptApp.newTrigger('dailyFacebookAdsExport')
      .timeBased()
      .everyDays(1)
      .atHour(2)
      .create();
    
    console.log('Daily trigger set up successfully');
    
  } catch (error) {
    console.error('Error setting up trigger:', error);
    throw error;
  }
}

/**
 * Send error notification
 */
function sendErrorNotification(error) {
  try {
    // You can customize this to send email, Slack, etc.
    console.error('AUTOMATION ERROR:', error.message);
    
    // Example: Send email notification
    // MailApp.sendEmail({
    //   to: 'your-email@example.com',
    //   subject: 'Facebook Ads Export Error',
    //   body: 'The daily Facebook Ads export failed: ' + error.message
    // });
    
  } catch (notificationError) {
    console.error('Error sending notification:', notificationError);
  }
}

/**
 * Test the automation manually
 */
function testAutomation() {
  console.log('Testing Facebook Ads automation...');
  dailyFacebookAdsExport();
}

/**
 * Manual setup function
 */
function setupAutomation() {
  console.log('Setting up Facebook Ads automation...');
  
  // 1. Set up daily trigger
  setupDailyTrigger();
  
  // 2. Test the automation
  testAutomation();
  
  console.log('Automation setup completed!');
}
