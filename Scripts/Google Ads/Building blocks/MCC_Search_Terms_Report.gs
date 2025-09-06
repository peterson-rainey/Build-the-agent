// --- MCC SEARCH TERMS REPORT CONFIGURATION ---

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
const TAB = 'SearchTermsData';

// --- END OF CONFIGURATION ---

const QUERY = `
  SELECT 
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    ad_group.id,
    ad_group.name,
    ad_group.status,
    search_term_view.search_term,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM search_term_view
  WHERE segments.date DURING LAST_30_DAYS
    AND campaign.advertising_channel_type != "PERFORMANCE_MAX"
  ORDER BY metrics.impressions DESC
`;

function main() {
  Logger.log("=== Starting MCC Search Terms Report ===");
  
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

      // Get search terms data for this account
      const searchTermsData = getSearchTermsDataForAccount();
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, searchTermsData, accountName);
      
      Logger.log(`✓ Successfully updated search terms data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC SEARCH TERMS REPORT COMPLETED ===`);
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

function getSearchTermsDataForAccount() {
  Logger.log("Getting search terms data for current account...");
  
  try {
    // Log sample row structure for debugging
    const sampleQuery = QUERY + ' LIMIT 1';
    const sampleRows = AdsApp.search(sampleQuery);
    
    if (sampleRows.hasNext()) {
      const sampleRow = sampleRows.next();
      Logger.log("Sample row structure: " + JSON.stringify(sampleRow));
      if (sampleRow.campaign) {
        Logger.log("Sample campaign object: " + JSON.stringify(sampleRow.campaign));
      }
      if (sampleRow.adGroup) {
        Logger.log("Sample adGroup object: " + JSON.stringify(sampleRow.adGroup));
      }
      if (sampleRow.searchTermView) {
        Logger.log("Sample searchTermView object: " + JSON.stringify(sampleRow.searchTermView));
      }
      if (sampleRow.metrics) {
        Logger.log("Sample metrics object: " + JSON.stringify(sampleRow.metrics));
      }
    } else {
      Logger.log("Query returned no rows for sample check.");
    }

    // Process the main query results
    const rows = AdsApp.search(QUERY);
    const data = calculateMetrics(rows);
    
    // Sort data by conversions (descending) then by status (Enabled, Paused, Removed)
    const sortedData = sortData(data);
    
    Logger.log(`✓ Processed ${data.length} search term records for current account`);
    return sortedData;
    
  } catch (error) {
    Logger.log(`❌ Error getting search terms data: ${error.message}`);
    throw error;
  }
}

function calculateMetrics(rows) {
  let data = [];
  let totalRows = 0;
  
  while (rows.hasNext()) {
    try {
      const row = rows.next();
      const campaign = row.campaign || {};
      const adGroup = row.adGroup || {};
      const searchTermView = row.searchTermView || {};
      const metrics = row.metrics || {};
      
      const campaignId = campaign.id || 'Unknown';
      const campaignName = campaign.name || 'Unknown Campaign';
      const campaignType = campaign.advertisingChannelType || 'Unknown';
      const campaignStatus = campaign.status || 'Unknown';
      const adGroupId = adGroup.id || 'Unknown';
      const adGroupName = adGroup.name || 'Unknown Ad Group';
      const adGroupStatus = adGroup.status || 'Unknown';
      
      // Get search term
      const searchTerm = searchTermView.searchTerm || 'Unknown Search Term';
      
      // Access metrics with proper number conversion
      let impressions = Number(metrics.impressions) || 0;
      let clicks = Number(metrics.clicks) || 0;
      let costMicros = Number(metrics.costMicros) || 0;
      let conversions = Number(metrics.conversions) || 0;
      let conversionValue = Number(metrics.conversionsValue) || 0;
      
      // Convert cost from micros to actual currency
      let cost = costMicros / 1000000;
      
      // Only include rows with more than 1 impression (filter out 1 impression or less)
      if (impressions > 1) {
        // Determine combined status (campaign and ad group)
        let combinedStatus = 'ENABLED';
        if (campaignStatus === 'PAUSED' || adGroupStatus === 'PAUSED') {
          combinedStatus = 'PAUSED';
        } else if (campaignStatus === 'REMOVED' || adGroupStatus === 'REMOVED') {
          combinedStatus = 'REMOVED';
        }
        
        let newRow = [
          campaignName,
          campaignType,
          campaignStatus,
          adGroupName,
          adGroupStatus,
          searchTerm,
          combinedStatus,
          impressions,
          clicks,
          cost,
          conversions,
          conversionValue
        ];
        
        data.push(newRow);
        totalRows++;
      }
      
    } catch (error) {
      Logger.log(`Error processing search term row: ${error.message}`);
    }
  }
  
  Logger.log(`✓ Processed ${totalRows} search term records with impressions > 1`);
  
  return data;
}

function sortData(data) {
  return data.sort((a, b) => {
    // Primary sort: conversions (descending)
    const conversionsA = a[11] || 0; // conversions is at index 11
    const conversionsB = b[11] || 0;
    
    if (conversionsA !== conversionsB) {
      return conversionsB - conversionsA; // Descending order
    }
    
    // Secondary sort: combined status (Enabled, Paused, Removed)
    const statusA = a[6] || ''; // combined status is at index 6
    const statusB = b[6] || '';
    
    const statusOrder = { 'ENABLED': 1, 'PAUSED': 2, 'REMOVED': 3 };
    const orderA = statusOrder[statusA] || 4;
    const orderB = statusOrder[statusB] || 4;
    
    return orderA - orderB; // Ascending order (Enabled first)
  });
}

function updateAccountSpreadsheet(spreadsheetUrl, searchTermsData, accountName) {
  try {
    Logger.log("Updating account spreadsheet with search terms data...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Create or get the tab for search terms data
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
      'Campaign Status',
      'Ad Group Name',
      'Ad Group Status',
      'Search Term',
      'Combined Status',
      'Impressions',
      'Clicks',
      'Cost',
      'Conversions',
      'Conversion Value'
    ];
    
    // Prepare all data for bulk write
    let allData = [headers];
    
    if (searchTermsData.length > 0) {
      allData = allData.concat(searchTermsData);
    } else {
      // Add a row indicating no data found
      allData.push(['No search terms data found for the last 30 days']);
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
    
    Logger.log(`✓ Successfully updated ${TAB} sheet with ${searchTermsData.length} search term records`);
    Logger.log(`✓ Account: ${accountName}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}
