/**
 * Facebook Ads Reporting System - Troubleshooting Script
 * Use this script to diagnose and fix common setup issues
 */

/**
 * Check and fix sheet naming issues
 */
function troubleshootSheetNames() {
  try {
    console.log('Starting troubleshooting...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    console.log('Available sheets in the spreadsheet:');
    sheets.forEach((sheet, index) => {
      console.log(`${index + 1}. "${sheet.getName()}"`);
    });
    
    // Check for Raw Data sheet
    const rawDataSheet = findSheetByName(sheets, ['Raw Data', 'RawData', 'raw data', 'rawdata', 'Data', 'data']);
    if (rawDataSheet) {
      console.log(`Found raw data sheet: "${rawDataSheet.getName()}"`);
      
      // Rename if needed
      if (rawDataSheet.getName() !== 'Raw Data') {
        console.log(`Renaming "${rawDataSheet.getName()}" to "Raw Data"`);
        rawDataSheet.setName('Raw Data');
        console.log('Sheet renamed successfully');
      }
    } else {
      console.log('No raw data sheet found. Creating one...');
      createRawDataSheet(spreadsheet);
    }
    
    // Check for other required sheets
    checkAndCreateRequiredSheets(spreadsheet);
    
    console.log('Troubleshooting completed successfully');
    
  } catch (error) {
    console.error('Error during troubleshooting:', error);
  }
}

/**
 * Find sheet by name with multiple possible variations
 */
function findSheetByName(sheets, possibleNames) {
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    if (possibleNames.includes(sheetName)) {
      return sheet;
    }
  }
  return null;
}

/**
 * Create Raw Data sheet with proper headers
 */
function createRawDataSheet(spreadsheet) {
  const rawDataSheet = spreadsheet.insertSheet('Raw Data');
  
  // Add headers
  const headers = [
    'Campaign Name',
    'Ad Set Name',
    'Ad Name',
    'Date',
    'Impressions',
    'Clicks',
    'Spend',
    'Results',
    'Cost per Result',
    'Reach',
    'Frequency',
    'CPM',
    'CPC',
    'CTR',
    'Relevance Score',
    'Quality Ranking',
    'Engagement Rate Ranking',
    'Conversion Rate Ranking',
    'Revenue',
    'ROAS'
  ];
  
  rawDataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = rawDataSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  
  // Add filters
  headerRange.createFilter();
  
  // Set up data validation for date column
  const dateColumn = rawDataSheet.getRange(2, 4, 1000, 1); // Column D, starting from row 2
  const dateValidation = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('Please enter a valid date')
    .build();
  dateColumn.setDataValidation(dateValidation);
  
  console.log('Raw Data sheet created with proper headers and formatting');
}

/**
 * Check and create other required sheets
 */
function checkAndCreateRequiredSheets(spreadsheet) {
  const requiredSheets = [
    { name: 'Processed Data', exists: false },
    { name: 'Report', exists: false },
    { name: 'Dashboard', exists: false }
  ];
  
  const existingSheets = spreadsheet.getSheets();
  
  // Check which sheets exist
  for (const required of requiredSheets) {
    for (const sheet of existingSheets) {
      if (sheet.getName() === required.name) {
        required.exists = true;
        break;
      }
    }
  }
  
  // Create missing sheets
  for (const required of requiredSheets) {
    if (!required.exists) {
      console.log(`Creating "${required.name}" sheet`);
      spreadsheet.insertSheet(required.name);
    }
  }
}

/**
 * Test data processing with sample data
 */
function testWithSampleData() {
  try {
    console.log('Testing with sample data...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const rawDataSheet = spreadsheet.getSheetByName('Raw Data');
    
    if (!rawDataSheet) {
      console.log('Raw Data sheet not found. Running troubleshooting first...');
      troubleshootSheetNames();
      return;
    }
    
    // Check if there's already data
    const existingData = rawDataSheet.getDataRange().getValues();
    if (existingData.length > 1) {
      console.log('Data already exists in Raw Data sheet. Processing existing data...');
      processFacebookAdsData();
      return;
    }
    
    // Add sample data
    const sampleData = [
      ['Summer Sale 2024', 'Women\'s Collection', 'Summer Dress Ad 1', '2024-01-01', 1250, 75, 187.50, 8, 23.44, 980, 1.28, 150.00, 2.50, 6.00, 8, 'Good', 'Above Average', 'Above Average', 320, 1.71],
      ['Summer Sale 2024', 'Women\'s Collection', 'Summer Dress Ad 2', '2024-01-01', 980, 58, 147.00, 6, 24.50, 750, 1.31, 150.00, 2.53, 5.92, 7, 'Good', 'Above Average', 'Average', 240, 1.63],
      ['Brand Awareness', 'General Audience', 'Brand Video Ad 1', '2024-01-01', 2000, 40, 300.00, 2, 150.00, 1600, 1.25, 150.00, 7.50, 2.00, 6, 'Average', 'Average', 'Below Average', 80, 0.27],
      ['Product Launch', 'Early Adopters', 'New Product Ad 1', '2024-01-01', 800, 120, 240.00, 15, 16.00, 600, 1.33, 300.00, 2.00, 15.00, 9, 'Excellent', 'Above Average', 'Above Average', 600, 2.50],
      ['Retargeting', 'Previous Visitors', 'Retargeting Ad 1', '2024-01-01', 400, 80, 120.00, 10, 12.00, 300, 1.33, 300.00, 1.50, 20.00, 9, 'Excellent', 'Above Average', 'Above Average', 400, 3.33]
    ];
    
    // Add sample data starting from row 2 (after headers)
    rawDataSheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    console.log('Sample data added successfully');
    console.log('Now processing the data...');
    
    // Process the data
    processFacebookAdsData();
    
  } catch (error) {
    console.error('Error testing with sample data:', error);
  }
}

/**
 * Check script configuration
 */
function checkConfiguration() {
  console.log('Checking script configuration...');
  
  console.log('CONFIG object:');
  console.log(JSON.stringify(CONFIG, null, 2));
  
  console.log('Available functions:');
  console.log('- troubleshootSheetNames() - Fix sheet naming issues');
  console.log('- testWithSampleData() - Test with sample data');
  console.log('- processFacebookAdsData() - Process existing data');
  console.log('- manualProcess() - Manual processing trigger');
  console.log('- initializeReportingSystem() - Full system initialization');
}

/**
 * Quick setup function
 */
function quickSetup() {
  try {
    console.log('Running quick setup...');
    
    // Step 1: Fix sheet names
    troubleshootSheetNames();
    
    // Step 2: Add sample data and test
    testWithSampleData();
    
    console.log('Quick setup completed successfully!');
    console.log('Check the generated sheets: Processed Data, Report, and Dashboard');
    
  } catch (error) {
    console.error('Error during quick setup:', error);
  }
}

/**
 * Verify data format
 */
function verifyDataFormat() {
  try {
    console.log('Verifying data format...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const rawDataSheet = spreadsheet.getSheetByName('Raw Data');
    
    if (!rawDataSheet) {
      console.log('Raw Data sheet not found');
      return;
    }
    
    const data = rawDataSheet.getDataRange().getValues();
    
    if (data.length === 0) {
      console.log('No data found in Raw Data sheet');
      return;
    }
    
    const headers = data[0];
    console.log('Headers found:', headers);
    
    // Check required headers
    const requiredHeaders = [
      'Campaign Name',
      'Ad Set Name', 
      'Ad Name',
      'Date',
      'Impressions',
      'Clicks',
      'Spend',
      'Results'
    ];
    
    const missingHeaders = [];
    for (const required of requiredHeaders) {
      if (!headers.includes(required)) {
        missingHeaders.push(required);
      }
    }
    
    if (missingHeaders.length > 0) {
      console.log('Missing required headers:', missingHeaders);
      console.log('Please ensure your data has the correct column headers');
    } else {
      console.log('All required headers are present');
    }
    
    // Check data rows
    if (data.length > 1) {
      console.log(`Found ${data.length - 1} data rows`);
      
      // Check first few rows for data quality
      for (let i = 1; i < Math.min(4, data.length); i++) {
        const row = data[i];
        console.log(`Row ${i + 1}:`, row.slice(0, 5)); // Show first 5 columns
      }
    }
    
  } catch (error) {
    console.error('Error verifying data format:', error);
  }
}
