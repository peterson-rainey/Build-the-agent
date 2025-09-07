// --- MCC ASSET PERFORMANCE REPORT CONFIGURATION ---

// 1. MASTER SPREADSHEET URL
//    - This spreadsheet contains the mapping of Account IDs to their individual spreadsheet URLs
//    - Format: Column A = Client Name, Column B = Account ID (10-digit), Column E = Spreadsheet URL
const MASTER_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1bS1xmTFB0LWGvYeZBPcykGMY0SwZ4W-Ny6aKVloy0SM/edit?usp=sharing";

// 2. MASTER SPREADSHEET TAB NAME
//    - The tab name in the master spreadsheet that contains the account mappings
const MASTER_SHEET_NAME = "AccountMappings"; // Default tab name

// 3. TEST CID (Optional)
//    - To run for a single account for testing, enter its Customer ID (e.g., "123-456-7890").
//    - Leave blank ("") to run for all accounts in the master spreadsheet.
const SINGLE_CID_FOR_TESTING = "242-931-3541"; // Example: "123-456-7890" or ""

// 4. REPORT CONFIGURATION
const TAB = 'HeadlinePerformance';

// --- END OF CONFIGURATION ---

// Query for headlines and descriptions using campaign_asset (provides conversion data)
const QUERY = `
  SELECT 
    campaign_asset.asset,
    campaign_asset.field_type,
    asset.text_asset.text,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM campaign_asset 
  WHERE segments.date DURING LAST_30_DAYS
  AND campaign_asset.field_type IN ('HEADLINE', 'DESCRIPTION')
  ORDER BY metrics.impressions DESC
`;

function main() {
  Logger.log("=== Starting MCC Asset Performance Report ===");
  
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

      // Get asset performance data for this account
      const assetData = getAssetPerformanceDataForAccount();
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, assetData, accountName);
      
      Logger.log(`✓ Successfully updated asset performance data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC ASSET PERFORMANCE REPORT COMPLETED ===`);
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
      const spreadsheetUrl = row[4]?.toString().trim(); // Column E (index 4)

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

function getAssetPerformanceDataForAccount() {
  Logger.log("Getting asset performance data for current account...");
  
  try {
    // Log sample row structure for debugging
    const sampleQuery = QUERY + ' LIMIT 1';
    const sampleRows = AdsApp.search(sampleQuery);
    
    if (sampleRows.hasNext()) {
      const sampleRow = sampleRows.next();
      Logger.log("Sample row structure: " + JSON.stringify(sampleRow));
      if (sampleRow.metrics) {
        Logger.log("Sample metrics object: " + JSON.stringify(sampleRow.metrics));
      }
    } else {
      Logger.log("Query returned no rows for sample check.");
    }

    // Process the main query results
    Logger.log("Processing headlines and descriptions from campaign_asset...");
    const rows = AdsApp.search(QUERY);
    const data = calculateAssetMetrics(rows);
    Logger.log(`✓ Found ${data.length} records from campaign_asset`);
    
    // Sort data by impressions (descending) then by status (Enabled, Paused, Removed)
    const sortedData = sortData(data);
    
    Logger.log(`✓ Processed ${data.length} asset performance records for current account`);
    return sortedData;
    
  } catch (error) {
    Logger.log(`❌ Error getting asset performance data: ${error.message}`);
    throw error;
  }
}


function calculateAssetMetrics(rows) {
  let data = [];
  let rowCount = 0;
  let assetTypeCounts = {};

  while (rows.hasNext()) {
    try {
      let row = rows.next();
      rowCount++;

      // Access dimensions using correct nested object structure
      let assetResourceName = row.campaignAsset && row.campaignAsset.asset ? row.campaignAsset.asset : 'N/A';
      let fieldType = row.campaignAsset && row.campaignAsset.fieldType ? row.campaignAsset.fieldType : 'UNKNOWN';
      
      // Extract actual asset text content
      let assetText = 'N/A';
      if (row.asset && row.asset.textAsset) {
        assetText = row.asset.textAsset.text || 'N/A';
      }
      
      // Clean up asset text
      if (assetText === 'N/A' || !assetText.trim()) {
        assetText = assetResourceName.split('/').pop(); // Fallback to asset ID
      }

      // Access metrics nested within the 'metrics' object
      const metrics = row.metrics || {};

      let impressions = Number(metrics.impressions) || 0;
      let clicks = Number(metrics.clicks) || 0;
      let costMicros = Number(metrics.costMicros) || 0;
      let conversions = Number(metrics.conversions) || 0;
      let conversionValue = Number(metrics.conversionsValue) || 0;

      // Convert cost from micros to actual currency
      let cost = costMicros / 1000000;

      // Calculate additional metrics
      let ctr = impressions > 0 ? (clicks / impressions) : 0;
      let averageCpc = clicks > 0 ? (cost / clicks) : 0;
      let costPerConversion = conversions > 0 ? (cost / conversions) : 0;

      // Only include assets with impressions > 0
      if (impressions > 0) {
        // Count asset types for logging
        if (assetTypeCounts[fieldType]) {
          assetTypeCounts[fieldType]++;
        } else {
          assetTypeCounts[fieldType] = 1;
        }

        // Determine asset status and eligibility
        let assetStatus = 'Enabled'; // Default to Enabled
        let status = 'Eligible'; // Default to Eligible
        let statusReason = ''; // Default to empty
        
        // Create one row per asset matching your CSV format
        let newRow = [
          assetStatus, // Asset Status (Enabled/Paused)
          assetText, // Text (actual headline text)
          'Ad', // Level (always "Ad" in your data)
          status, // Status (Eligible/Not eligible)
          statusReason, // Status Reason (empty or " --")
          'Advertiser', // Source (always "Advertiser" in your data)
          ' --', // Last Updated (always " --" in your data)
          clicks, // Clicks
          impressions, // Impressions
          ctr, // CTR
          'USD', // Currency Code (always "USD" in your data)
          averageCpc, // Avg CPC
          cost, // Cost
          conversions, // Conversions
          costPerConversion // Cost / Conv
        ];

        data.push(newRow);
      }

    } catch (e) {
      Logger.log("Error processing row: " + e + " | Row data: " + JSON.stringify(row));
      // Continue with next row
    }
  }

  Logger.log(`✓ Found ${data.length} assets with impressions > 0`);
  Logger.log(`Asset breakdown: ${JSON.stringify(assetTypeCounts)}`);
  return data;
}

function sortData(data) {
  return data.sort((a, b) => {
    // Primary sort: impressions (descending) - now at index 8
    const impressionsA = a[8] || 0;
    const impressionsB = b[8] || 0;
    
    if (impressionsA !== impressionsB) {
      return impressionsB - impressionsA; // Descending order
    }
    
    // Secondary sort: cost (descending) - now at index 12
    const costA = a[12] || 0;
    const costB = b[12] || 0;
    
    return costB - costA; // Descending order
  });
}

function updateAccountSpreadsheet(spreadsheetUrl, assetData, accountName) {
  try {
    Logger.log("Updating account spreadsheet with asset performance data...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Create or get the tab for asset performance data
    let sheet = spreadsheet.getSheetByName(TAB);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(TAB);
      Logger.log(`✓ Created new sheet: ${TAB}`);
    } else {
      // Clear existing data
      sheet.clear();
      Logger.log(`✓ Cleared existing data in sheet: ${TAB}`);
    }
    
    // Create headers (conversions not available in ad_group_ad_asset_view for Search campaigns)
    const headers = [
      'Asset Status',
      'Text',
      'Level',
      'Status',
      'Status Reason',
      'Source',
      'Last Updated',
      'Clicks',
      'Impressions',
      'CTR',
      'Currency Code',
      'Avg CPC',
      'Cost',
      'Conversions',
      'Cost / Conv'
    ];
    
    // Prepare all data for bulk write
    let allData = [headers];
    
    if (assetData.length > 0) {
      allData = allData.concat(assetData);
    } else {
      // Add a row indicating no data found
      allData.push(['No asset performance data found for the last 30 days']);
    }
    
    // Write all data at once using setValues for efficiency
    const range = sheet.getRange(1, 1, allData.length, headers.length);
    range.setValues(allData);
    
    // Format the spreadsheet
    formatSpreadsheet(sheet, allData.length, headers.length);
    
    Logger.log(`✓ Successfully updated ${TAB} sheet with ${assetData.length} asset performance rows`);
    Logger.log(`✓ Account: ${accountName}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}

function formatSpreadsheet(sheet, totalRows, totalCols) {
  try {
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, totalCols);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285F4');
    headerRange.setFontColor('#FFFFFF');
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, totalCols);
    
    // Add borders to all data
    const dataRange = sheet.getRange(1, 1, totalRows, totalCols);
    dataRange.setBorder(true, true, true, true, true, true);
    
    // Format currency columns (Avg CPC, Cost, Cost / Conv)
    const currencyColumns = [11, 12, 14]; // 0-based indices for Avg CPC, Cost, Cost / Conv
    for (const col of currencyColumns) {
      if (col < totalCols) {
        const range = sheet.getRange(2, col + 1, totalRows - 1, 1);
        range.setNumberFormat('$#,##0.00');
      }
    }
    
    // Format percentage columns (CTR)
    const percentageColumns = [9]; // 0-based index for CTR
    for (const col of percentageColumns) {
      if (col < totalCols) {
        const range = sheet.getRange(2, col + 1, totalRows - 1, 1);
        range.setNumberFormat('0.00%');
      }
    }
    
    // Format number columns (Clicks, Impressions, Conversions)
    const numberColumns = [7, 8, 13]; // 0-based indices for Clicks, Impressions, Conversions
    for (const col of numberColumns) {
      if (col < totalCols) {
        const range = sheet.getRange(2, col + 1, totalRows - 1, 1);
        range.setNumberFormat('#,##0');
      }
    }
    
    Logger.log(`✓ Applied formatting to spreadsheet`);
    
  } catch (error) {
    Logger.log(`❌ Error formatting spreadsheet: ${error.message}`);
  }
}
