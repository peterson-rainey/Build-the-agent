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
const TAB = 'AssetPerformance';

// --- END OF CONFIGURATION ---

const QUERY = `
  SELECT 
    campaign.name,
    campaign.status,
    campaign.advertising_channel_type,
    ad_group.name,
    ad_group.status,
    ad_group_ad.ad.id,
    ad_group_ad.ad.name,
    ad_group_ad.status,
    ad_group_ad_asset_view.asset,
    ad_group_ad_asset_view.field_type,
    ad_group_ad_asset_view.performance_label,
    metrics.cost_micros,
    metrics.impressions,
    metrics.clicks,
    metrics.conversions,
    metrics.conversions_value,
    metrics.ctr,
    metrics.average_cpc,
    metrics.cost_per_conversion,
    metrics.value_per_conversion
  FROM ad_group_ad_asset_view 
  WHERE segments.date DURING LAST_30_DAYS
  AND ad_group_ad_asset_view.field_type IN ('HEADLINE', 'DESCRIPTION')
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

    // Create asset text lookup table
    Logger.log("Creating asset text lookup table...");
    const assetTextLookup = createAssetTextLookup();
    Logger.log(`✓ Created lookup table with ${Object.keys(assetTextLookup).length} assets`);

    // Process the main query results
    const rows = AdsApp.search(QUERY);
    const data = calculateAssetMetrics(rows, assetTextLookup);
    
    // Sort data by impressions (descending) then by status (Enabled, Paused, Removed)
    const sortedData = sortData(data);
    
    Logger.log(`✓ Processed ${data.length} asset performance records for current account`);
    return sortedData;
    
  } catch (error) {
    Logger.log(`❌ Error getting asset performance data: ${error.message}`);
    throw error;
  }
}

function createAssetTextLookup() {
  const assetLookup = {};
  
  try {
    // Query all text assets in the account
    const assetQuery = `
      SELECT 
        asset.resource_name,
        asset.text_asset.text
      FROM asset 
      WHERE asset.type = 'TEXT'
    `;
    
    const assetRows = AdsApp.search(assetQuery);
    let assetCount = 0;
    
    while (assetRows.hasNext()) {
      try {
        const assetRow = assetRows.next();
        const resourceName = assetRow.asset.resourceName;
        const text = assetRow.asset.textAsset ? assetRow.asset.textAsset.text : 'N/A';
        
        if (resourceName && text) {
          assetLookup[resourceName] = text;
          assetCount++;
        }
      } catch (error) {
        Logger.log(`Error processing asset row: ${error.message}`);
      }
    }
    
    Logger.log(`✓ Processed ${assetCount} text assets for lookup table`);
    return assetLookup;
    
  } catch (error) {
    Logger.log(`❌ Error creating asset text lookup: ${error.message}`);
    return {};
  }
}

function calculateAssetMetrics(rows, assetTextLookup) {
  let data = [];
  let rowCount = 0;

  while (rows.hasNext()) {
    try {
      let row = rows.next();
      rowCount++;

      // Access dimensions using correct nested object structure
      let campaignName = row.campaign ? row.campaign.name : 'N/A';
      let campaignStatus = row.campaign ? row.campaign.status : 'UNKNOWN';
      let campaignType = row.campaign ? row.campaign.advertisingChannelType : 'UNKNOWN';
      let adGroupName = row.adGroup ? row.adGroup.name : 'N/A';
      let adGroupStatus = row.adGroup ? row.adGroup.status : 'UNKNOWN';
      
      // Access ad information
      let adId = row.adGroupAd && row.adGroupAd.ad ? row.adGroupAd.ad.id : 'N/A';
      let adName = row.adGroupAd && row.adGroupAd.ad ? row.adGroupAd.ad.name : 'N/A';
      let adStatus = row.adGroupAd ? row.adGroupAd.status : 'UNKNOWN';

      // Access individual asset data
      let assetResourceName = 'N/A';
      let fieldType = 'N/A';
      let performanceLabel = 'N/A';
      let assetText = 'N/A';
      
      // Get data from ad_group_ad_asset_view
      if (row.adGroupAdAssetView) {
        fieldType = row.adGroupAdAssetView.fieldType || 'N/A';
        performanceLabel = row.adGroupAdAssetView.performanceLabel || 'N/A';
        assetResourceName = row.adGroupAdAssetView.asset || 'N/A';
      }
      
      // Get asset text content from lookup table
      if (assetResourceName && assetResourceName !== 'N/A' && assetTextLookup[assetResourceName]) {
        assetText = assetTextLookup[assetResourceName];
      } else {
        assetText = 'N/A';
      }
      
      // Extract asset ID from resource name (format: customers/{customer_id}/assets/{asset_id})
      let assetId = 'N/A';
      if (assetResourceName && assetResourceName !== 'N/A') {
        const parts = assetResourceName.split('/');
        if (parts.length > 0) {
          assetId = parts[parts.length - 1];
        }
      }

      // Determine combined status - if any component is paused, show as paused
      let combinedStatus = 'ENABLED';
      if (campaignStatus === 'PAUSED' || adGroupStatus === 'PAUSED' || adStatus === 'PAUSED') {
        combinedStatus = 'PAUSED';
      } else if (campaignStatus === 'REMOVED' || adGroupStatus === 'REMOVED' || adStatus === 'REMOVED') {
        combinedStatus = 'REMOVED';
      }

      // Access metrics nested within the 'metrics' object - these are asset-specific
      const metrics = row.metrics || {};

      let costMicros = Number(metrics.costMicros) || 0;
      let cost = costMicros / 1000000; // Convert from micros to currency units
      let impressions = Number(metrics.impressions) || 0;
      let clicks = Number(metrics.clicks) || 0;
      let conversions = Number(metrics.conversions) || 0;
      let conversionsValue = Number(metrics.conversionsValue) || 0;
      let ctr = Number(metrics.ctr) || 0;
      let averageCpc = Number(metrics.averageCpc) || 0;
      let costPerConversion = Number(metrics.costPerConversion) || 0;
      let valuePerConversion = Number(metrics.valuePerConversion) || 0;

      // Skip assets with zero impressions (optional - remove this if you want all assets)
      if (impressions === 0) {
        continue;
      }

      // Create one row per asset with its individual performance metrics
      let newRow = [
        campaignName,
        campaignType,
        adGroupName,
        adId,
        adName,
        combinedStatus,
        fieldType, // Asset Type (HEADLINE or DESCRIPTION) - moved up for prominence
        assetText, // Actual asset text content - moved up for prominence
        assetId, // Asset ID
        assetResourceName, // Asset resource name
        performanceLabel, // Performance label (BEST, GOOD, LOW, PENDING)
        cost, // Cost in currency units
        impressions, // Impressions
        clicks, // Clicks
        conversions, // Conversions
        conversionsValue, // Conversion value
        ctr, // Click-through rate
        averageCpc, // Average CPC
        costPerConversion, // Cost per conversion
        valuePerConversion // Value per conversion
      ];
      data.push(newRow);

    } catch (e) {
      Logger.log("Error processing row: " + e + " | Row data: " + JSON.stringify(row));
      // Continue with next row
    }
  }

  Logger.log(`Processed ${rowCount} total rows, ${data.length} successful asset rows`);
  return data;
}

function sortData(data) {
  return data.sort((a, b) => {
    // Primary sort: impressions (descending)
    const impressionsA = a[13] || 0; // impressions is now at index 13
    const impressionsB = b[13] || 0;
    
    if (impressionsA !== impressionsB) {
      return impressionsB - impressionsA; // Descending order
    }
    
    // Secondary sort: status (Enabled, Paused, Removed)
    const statusA = a[5] || ''; // status is still at index 5
    const statusB = b[5] || '';
    
    const statusOrder = { 'ENABLED': 1, 'PAUSED': 2, 'REMOVED': 3 };
    const orderA = statusOrder[statusA] || 4;
    const orderB = statusOrder[statusB] || 4;
    
    return orderA - orderB; // Ascending order (Enabled first)
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
    
    // Create headers
    const headers = [
      'Campaign Name',
      'Campaign Type',
      'Ad Group Name',
      'Ad ID',
      'Ad Name',
      'Status',
      'Asset Type',
      'Asset Text',
      'Asset ID',
      'Asset Resource Name',
      'Performance Label',
      'Cost',
      'Impressions',
      'Clicks',
      'Conversions',
      'Conversion Value',
      'CTR',
      'Avg CPC',
      'Cost/Conv',
      'Value/Conv'
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
    
    // Format currency columns (Cost, Avg CPC, Cost/Conv, Value/Conv)
    const currencyColumns = [11, 17, 18, 19]; // 0-based indices (moved due to new columns)
    for (const col of currencyColumns) {
      if (col < totalCols) {
        const range = sheet.getRange(2, col + 1, totalRows - 1, 1);
        range.setNumberFormat('$#,##0.00');
      }
    }
    
    // Format percentage columns (CTR)
    const percentageColumns = [16]; // 0-based index for CTR (moved due to new columns)
    for (const col of percentageColumns) {
      if (col < totalCols) {
        const range = sheet.getRange(2, col + 1, totalRows - 1, 1);
        range.setNumberFormat('0.00%');
      }
    }
    
    // Format number columns (Impressions, Clicks, Conversions)
    const numberColumns = [12, 13, 14]; // 0-based indices (moved due to new columns)
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
