// --- MCC ADS REPORT CONFIGURATION ---

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
const TAB = 'AdsData';

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
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM ad_group_ad 
  WHERE segments.date DURING LAST_30_DAYS
  AND campaign.advertising_channel_type != "PERFORMANCE_MAX"
  ORDER BY metrics.cost_micros DESC
`;

function main() {
  Logger.log("=== Starting MCC Ads Report ===");
  
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

      // Get ads data for this account
      const adsData = getAdsDataForAccount();
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, adsData, accountName);
      
      Logger.log(`✓ Successfully updated ads data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC ADS REPORT COMPLETED ===`);
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

function getAdsDataForAccount() {
  Logger.log("Getting ads data for current account...");
  
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
    const data = calculateMetrics(rows);
    
    // Sort data by conversions (descending) then by status (Enabled, Paused, Removed)
    const sortedData = sortData(data);
    
    Logger.log(`✓ Processed ${data.length} ads for current account`);
    return sortedData;
    
  } catch (error) {
    Logger.log(`❌ Error getting ads data: ${error.message}`);
    throw error;
  }
}

function calculateMetrics(rows) {
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

      // Determine combined status - if any component is paused, show as paused
      let combinedStatus = 'ENABLED';
      if (campaignStatus === 'PAUSED' || adGroupStatus === 'PAUSED' || adStatus === 'PAUSED') {
        combinedStatus = 'PAUSED';
      } else if (campaignStatus === 'REMOVED' || adGroupStatus === 'REMOVED' || adStatus === 'REMOVED') {
        combinedStatus = 'REMOVED';
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

      // Skip ads with zero impressions
      if (impressions === 0) {
        continue;
      }

      // Add only the requested metrics to a new row
      let newRow = [
        campaignName,
        campaignType,
        adGroupName,
        adId,
        adName,
        combinedStatus,
        impressions,
        clicks,
        cost,
        conversions,
        conversionValue
      ];

      data.push(newRow);

    } catch (e) {
      Logger.log("Error processing row: " + e + " | Row data: " + JSON.stringify(row));
      // Continue with next row
    }
  }

  Logger.log(`Processed ${rowCount} total rows, ${data.length} successful rows`);
  return data;
}

function sortData(data) {
  return data.sort((a, b) => {
    // Primary sort: conversions (descending)
    const conversionsA = a[9] || 0; // conversions is at index 9
    const conversionsB = b[9] || 0;
    
    if (conversionsA !== conversionsB) {
      return conversionsB - conversionsA; // Descending order
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

function updateAccountSpreadsheet(spreadsheetUrl, adsData, accountName) {
  try {
    Logger.log("Updating account spreadsheet with ads data...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Create or get the tab for ads data
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
      'Impressions',
      'Clicks',
      'Cost',
      'Conversions',
      'Conversion Value'
    ];
    
    // Prepare all data for bulk write
    let allData = [headers];
    
    if (adsData.length > 0) {
      allData = allData.concat(adsData);
    } else {
      // Add a row indicating no data found
      allData.push(['No ads data found for the last 30 days']);
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
    
    Logger.log(`✓ Successfully updated ${TAB} sheet with ${adsData.length} ads`);
    Logger.log(`✓ Account: ${accountName}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}
