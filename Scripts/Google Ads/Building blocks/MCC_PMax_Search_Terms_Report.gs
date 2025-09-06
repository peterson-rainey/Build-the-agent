// --- MCC PERFORMANCE MAX SEARCH TERMS REPORT CONFIGURATION ---

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
const TAB = 'PMaxSearchTermsData';

// --- END OF CONFIGURATION ---

// Main query to get Performance Max campaigns
const CAMPAIGN_QUERY = `
  SELECT 
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status
  FROM campaign
  WHERE campaign.advertising_channel_type = "PERFORMANCE_MAX"
    AND campaign.status IN ('ENABLED', 'PAUSED')
  ORDER BY campaign.name ASC
`;

function main() {
  Logger.log("=== Starting MCC Performance Max Search Terms Report ===");
  
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

      // Get Performance Max search terms data for this account
      const pmaxSearchTermsData = getPMaxSearchTermsDataForAccount();
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, pmaxSearchTermsData, accountName);
      
      Logger.log(`✓ Successfully updated Performance Max search terms data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC PERFORMANCE MAX SEARCH TERMS REPORT COMPLETED ===`);
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

function getPMaxSearchTermsDataForAccount() {
  try {
    Logger.log("Getting Performance Max search terms data for current account...");
    
    // First get all Performance Max campaigns
    const campaignQuery = AdsApp.search(CAMPAIGN_QUERY);
    const campaigns = [];
    
    while (campaignQuery.hasNext()) {
      const row = campaignQuery.next();
      campaigns.push({
        id: row.campaign.id,
        name: row.campaign.name,
        status: row.campaign.status
      });
    }
    
    Logger.log(`Found ${campaigns.length} Performance Max campaigns`);
    
    const allSearchTerms = [];
    
    // For each Performance Max campaign, get its search terms using campaign_search_term_insight
    for (const campaign of campaigns) {
      Logger.log(`Getting search terms for campaign: ${campaign.name} (ID: ${campaign.id})`);
      
      const searchTermsQuery = `
        SELECT 
          campaign_search_term_insight.category_label,
          metrics.impressions,
          metrics.clicks,
          metrics.conversions,
          metrics.conversions_value
        FROM campaign_search_term_insight
        WHERE segments.date DURING LAST_30_DAYS
          AND campaign_search_term_insight.campaign_id = ${campaign.id}
        ORDER BY metrics.impressions DESC
      `;
      
      try {
        const searchTerms = AdsApp.search(searchTermsQuery);
        let campaignSearchTermCount = 0;
        
        while (searchTerms.hasNext()) {
          const row = searchTerms.next();
          campaignSearchTermCount++;
          
          const impressions = Number(row.metrics.impressions) || 0;
          const clicks = Number(row.metrics.clicks) || 0;
          const conversions = Number(row.metrics.conversions) || 0;
          const conversionValue = Number(row.metrics.conversions_value) || 0;
          
          // Only include search terms with impressions > 1
          if (impressions > 1) {
            // Note: cost_micros is not available for campaign_search_term_insight
            const cost = 0; // Set to 0 since this metric is not available
            
            allSearchTerms.push([
              campaign.name,
              campaign.status,
              row.campaignSearchTermInsight.categoryLabel || 'Unknown',
              impressions,
              clicks,
              cost,
              conversions,
              conversionValue
            ]);
            
            if (campaignSearchTermCount <= 5) { // Log first 5 for debugging
              Logger.log(`  Search Term: "${row.campaignSearchTermInsight.categoryLabel}" - ${impressions} impressions, ${clicks} clicks, ${conversions} conversions`);
            }
          }
        }
        
        Logger.log(`  Found ${campaignSearchTermCount} search terms for campaign ${campaign.name}`);
        
      } catch (error) {
        Logger.log(`  ❌ Error getting search terms for campaign ${campaign.name}: ${error.message}`);
      }
    }
    
    // Sort data by conversions (descending) then by campaign status
    const sortedData = sortData(allSearchTerms);
    
    Logger.log(`✓ Processed ${allSearchTerms.length} Performance Max search term records with impressions > 1 for current account`);
    return sortedData;
    
  } catch (error) {
    Logger.log(`❌ Error getting Performance Max search terms data: ${error.message}`);
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
      const searchTermView = row.searchTermView || {};
      const metrics = row.metrics || {};
      
      const campaignId = campaign.id || 'Unknown';
      const campaignName = campaign.name || 'Unknown Campaign';
      const campaignType = campaign.advertisingChannelType || 'Unknown';
      const campaignStatus = campaign.status || 'Unknown';
      
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
      
      // Include all rows with impressions (Performance Max search terms are rare)
      if (impressions > 0) {
        let newRow = [
          campaignName,
          campaignType,
          campaignStatus,
          searchTerm,
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
      Logger.log(`Error processing Performance Max search term row: ${error.message}`);
    }
  }
  
  Logger.log(`✓ Processed ${totalRows} Performance Max search term records with impressions > 0`);
  
  return data;
}

function sortData(data) {
  return data.sort((a, b) => {
    // Primary sort: conversions (descending)
    const conversionsA = a[6] || 0; // conversions is at index 6
    const conversionsB = b[6] || 0;
    
    if (conversionsA !== conversionsB) {
      return conversionsB - conversionsA; // Descending order
    }
    
    // Secondary sort: campaign status (Enabled, Paused, Removed)
    const statusA = a[1] || ''; // campaign status is at index 1
    const statusB = b[1] || '';
    
    const statusOrder = { 'ENABLED': 1, 'PAUSED': 2, 'REMOVED': 3 };
    const orderA = statusOrder[statusA] || 4;
    const orderB = statusOrder[statusB] || 4;
    
    return orderA - orderB; // Ascending order (Enabled first)
  });
}

function updateAccountSpreadsheet(spreadsheetUrl, pmaxSearchTermsData, accountName) {
  try {
    Logger.log("Updating account spreadsheet with Performance Max search terms data...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Create or get the tab for Performance Max search terms data
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
      'Campaign Status',
      'Search Term',
      'Impressions',
      'Clicks',
      'Cost',
      'Conversions',
      'Conversion Value'
    ];
    
    // Prepare all data for bulk write
    let allData = [headers];
    
    if (pmaxSearchTermsData.length > 0) {
      allData = allData.concat(pmaxSearchTermsData);
    } else {
      // Add a row indicating no data found with proper column count
      const noDataRow = ['No Performance Max search terms data found for the last 30 days'];
      // Pad with empty strings to match header column count
      while (noDataRow.length < headers.length) {
        noDataRow.push('');
      }
      allData.push(noDataRow);
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
    
    Logger.log(`✓ Successfully updated ${TAB} sheet with ${pmaxSearchTermsData.length} Performance Max search term records`);
    Logger.log(`✓ Account: ${accountName}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}
