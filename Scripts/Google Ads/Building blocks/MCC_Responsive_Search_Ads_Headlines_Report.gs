// --- MCC RESPONSIVE SEARCH ADS HEADLINES REPORT CONFIGURATION ---

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
const TAB = 'RSAHeadlinesData';

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
    metrics.impressions
  FROM ad_group_ad_asset_view 
  WHERE segments.date DURING LAST_30_DAYS
  AND campaign.advertising_channel_type = "SEARCH"
  AND ad_group_ad.ad.type = "RESPONSIVE_SEARCH_AD"
  AND ad_group_ad_asset_view.field_type = "HEADLINE"
  ORDER BY metrics.impressions DESC
`;

function main() {
  Logger.log("=== Starting MCC Responsive Search Ads Headlines Report ===");
  
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

      // Get responsive search ads headlines data for this account
      const headlinesData = getRSAHeadlinesDataForAccount();
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, headlinesData, accountName);
      
      Logger.log(`✓ Successfully updated RSA headlines data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC RESPONSIVE SEARCH ADS HEADLINES REPORT COMPLETED ===`);
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

function getRSAHeadlinesDataForAccount() {
  Logger.log("Getting responsive search ads headlines data for current account...");
  
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
      if (sampleRow.campaign) {
        Logger.log("Sample campaign object: " + JSON.stringify(sampleRow.campaign));
      }
      if (sampleRow.adGroup) {
        Logger.log("Sample adGroup object: " + JSON.stringify(sampleRow.adGroup));
      }
      if (sampleRow.adGroupAd) {
        Logger.log("Sample adGroupAd object: " + JSON.stringify(sampleRow.adGroupAd));
      }
    } else {
      Logger.log("Query returned no rows for sample check.");
    }

    // Process the main query results
    const rows = AdsApp.search(QUERY);
    const data = calculateHeadlinesMetrics(rows);
    
    // Sort data by conversions (descending) then by status (Enabled, Paused, Removed)
    const sortedData = sortData(data);
    
    Logger.log(`✓ Processed ${data.length} RSA headlines for current account`);
    return sortedData;
    
  } catch (error) {
    Logger.log(`❌ Error getting RSA headlines data: ${error.message}`);
    throw error;
  }
}

function calculateHeadlinesMetrics(rows) {
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

      // Access individual headline asset data
      let assetResourceName = 'N/A';
      let fieldType = 'N/A';
      let performanceLabel = 'N/A';
      
      // Get data from ad_group_ad_asset_view
      if (row.adGroupAdAssetView) {
        fieldType = row.adGroupAdAssetView.fieldType || 'N/A';
        performanceLabel = row.adGroupAdAssetView.performanceLabel || 'N/A';
        assetResourceName = row.adGroupAdAssetView.asset || 'N/A';
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

      // Access metrics nested within the 'metrics' object - these are now headline-specific
      const metrics = row.metrics || {};

      let impressions = Number(metrics.impressions) || 0;

      // Skip headlines with zero impressions
      if (impressions === 0) {
        continue;
      }

      // Create one row per headline with its individual performance metrics
      let newRow = [
        campaignName,
        campaignType,
        adGroupName,
        adId,
        adName,
        combinedStatus,
        assetId, // Asset ID
        assetResourceName, // Asset resource name (can be used to get text later)
        performanceLabel, // Performance label (BEST, GOOD, LOW, PENDING)
        fieldType, // Field type (should be HEADLINE)
        impressions // Headline-specific impressions
      ];
      data.push(newRow);

    } catch (e) {
      Logger.log("Error processing row: " + e + " | Row data: " + JSON.stringify(row));
      // Continue with next row
    }
  }

  Logger.log(`Processed ${rowCount} total rows, ${data.length} successful headline rows`);
  return data;
}

function sortData(data) {
  return data.sort((a, b) => {
    // Primary sort: impressions (descending)
    const impressionsA = a[10] || 0; // impressions is at index 10
    const impressionsB = b[10] || 0;
    
    if (impressionsA !== impressionsB) {
      return impressionsB - impressionsA; // Descending order
    }
    
    // Secondary sort: status (Enabled, Paused, Removed)
    const statusA = a[5] || ''; // status is at index 5
    const statusB = b[5] || '';
    
    const statusOrder = { 'ENABLED': 1, 'PAUSED': 2, 'REMOVED': 3 };
    const orderA = statusOrder[statusA] || 4;
    const orderB = statusOrder[statusB] || 4;
    
    return orderA - orderB; // Ascending order (Enabled first)
  });
}

function updateAccountSpreadsheet(spreadsheetUrl, headlinesData, accountName) {
  try {
    Logger.log("Updating account spreadsheet with RSA headlines data...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Create or get the tab for headlines data
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
      'Asset ID',
      'Asset Resource Name',
      'Performance Label',
      'Field Type',
      'Impressions'
    ];
    
    // Prepare all data for bulk write
    let allData = [headers];
    
    if (headlinesData.length > 0) {
      allData = allData.concat(headlinesData);
    } else {
      // Add a row indicating no data found
      allData.push(['No responsive search ads headlines found for the last 30 days']);
    }
    
    // Write all data at once using setValues for efficiency
    const range = sheet.getRange(1, 1, allData.length, headers.length);
    range.setValues(allData);
    
    // Format the header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#E8E8E8');
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
    
    Logger.log(`✓ Successfully updated ${TAB} sheet with ${headlinesData.length} headline rows`);
    Logger.log(`✓ Account: ${accountName}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}
