// --- MCC NEGATIVE KEYWORDS REPORT CONFIGURATION ---

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
const TAB = 'NegativeKeywordsData';

// --- END OF CONFIGURATION ---

// Query 1: Campaign Level Negative Keywords
const CAMPAIGN_NEGATIVES_QUERY = `
  SELECT 
    campaign.id,
    campaign.name,
    campaign.status,
    campaign_criterion.keyword.text,
    campaign_criterion.keyword.match_type,
    campaign_criterion.status
  FROM campaign_criterion 
  WHERE campaign_criterion.negative = TRUE
  AND campaign_criterion.type = 'KEYWORD'
  AND campaign_criterion.status = 'ENABLED'
  AND campaign.status IN ('ENABLED', 'PAUSED')
  ORDER BY campaign.name, campaign_criterion.keyword.text
`;

// Query 2: Ad Group Level Negative Keywords
const ADGROUP_NEGATIVES_QUERY = `
  SELECT 
    campaign.name,
    campaign.status,
    ad_group.name,
    ad_group.status,
    ad_group_criterion.criterion_id,
    ad_group_criterion.keyword.text,
    ad_group_criterion.keyword.match_type,
    ad_group_criterion.status
  FROM ad_group_criterion 
  WHERE ad_group_criterion.type = 'KEYWORD'
  AND ad_group_criterion.negative = TRUE
  AND ad_group_criterion.status = 'ENABLED'
  AND campaign.status IN ('ENABLED', 'PAUSED')
  AND ad_group.status IN ('ENABLED', 'PAUSED')
  ORDER BY campaign.name, ad_group.name, ad_group_criterion.keyword.text
`;

// Note: Negative Keyword Lists are processed using AdsApp.negativeKeywordLists() API
// instead of GAQL queries for better compatibility

function main() {
  Logger.log("=== Starting MCC Negative Keywords Report ===");
  
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

      // Get negative keywords data for this account
      const negativeKeywordsData = getNegativeKeywordsDataForAccount();
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, negativeKeywordsData, accountName);
      
      Logger.log(`✓ Successfully updated negative keywords data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC NEGATIVE KEYWORDS REPORT COMPLETED ===`);
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

function getNegativeKeywordsDataForAccount() {
  Logger.log("Getting negative keywords data for current account...");
  
  try {
    const allData = [];
    
    // Process Campaign Level Negative Keywords
    Logger.log("Getting campaign level negative keywords...");
    try {
      const campaignRows = AdsApp.search(CAMPAIGN_NEGATIVES_QUERY);
      const campaignData = processCampaignNegativesData(campaignRows);
      allData.push(...campaignData);
      Logger.log(`✓ Found ${campaignData.length} campaign negative keywords`);
    } catch (error) {
      Logger.log(`❌ Error getting campaign negative keywords: ${error.message}`);
    }
    
    // Process Ad Group Level Negative Keywords
    Logger.log("Getting ad group level negative keywords...");
    try {
      const adGroupRows = AdsApp.search(ADGROUP_NEGATIVES_QUERY);
      const adGroupData = processAdGroupNegativesData(adGroupRows);
      allData.push(...adGroupData);
      Logger.log(`✓ Found ${adGroupData.length} ad group negative keywords`);
    } catch (error) {
      Logger.log(`❌ Error getting ad group negative keywords: ${error.message}`);
    }
    
    // Note: Negative keyword lists are not processed as individual keywords
    // The Google Ads API doesn't provide a reliable way to extract individual keywords from lists
    // We focus on campaign and ad group level negative keywords instead
    
    // Sort all data
    const sortedData = sortData(allData);
    
    Logger.log(`✓ Processed ${allData.length} total negative keyword records for current account`);
    return sortedData;
    
  } catch (error) {
    Logger.log(`❌ Error getting negative keywords data: ${error.message}`);
    throw error;
  }
}

function processCampaignNegativesData(rows) {
  let data = [];
  let rowCount = 0;

  while (rows.hasNext()) {
    try {
      let row = rows.next();
      rowCount++;

      // Access campaign data
      let campaignName = row.campaign ? row.campaign.name : 'N/A';
      let campaignStatus = row.campaign ? row.campaign.status : 'N/A';

      // Access campaign criterion data
      let criterionId = row.campaignCriterion ? row.campaignCriterion.criterionId : 'N/A';
      let criterionStatus = row.campaignCriterion ? row.campaignCriterion.status : 'N/A';

      // Access keyword data
      let keywordText = 'N/A';
      let matchType = 'N/A';
      if (row.campaignCriterion && row.campaignCriterion.keyword) {
        keywordText = row.campaignCriterion.keyword.text || 'N/A';
        matchType = row.campaignCriterion.keyword.matchType || 'N/A';
      }

      // Log first few keywords for debugging
      if (rowCount <= 5) {
        Logger.log(`  Campaign Keyword ${rowCount}: "${keywordText}"`);
      }

      // Format match type for better readability
      let matchTypeFormatted = formatMatchType(matchType);

      // Add negative keyword data to a new row
      let newRow = [
        'Campaign',
        campaignName,
        campaignStatus,
        'N/A', // No ad group for campaign level
        'N/A', // No ad group status for campaign level
        criterionId,
        keywordText,
        matchTypeFormatted,
        matchType,
        criterionStatus
      ];

      data.push(newRow);

    } catch (e) {
      Logger.log("Error processing campaign negative row: " + e);
    }
  }

  Logger.log(`Processed ${rowCount} campaign negative rows, ${data.length} negative keywords found`);
  return data;
}

function processAdGroupNegativesData(rows) {
  let data = [];
  let rowCount = 0;

  while (rows.hasNext()) {
    try {
      let row = rows.next();
      rowCount++;

      // Access campaign data
      let campaignName = row.campaign ? row.campaign.name : 'N/A';
      let campaignStatus = row.campaign ? row.campaign.status : 'N/A';

      // Access ad group data
      let adGroupName = row.adGroup ? row.adGroup.name : 'N/A';
      let adGroupStatus = row.adGroup ? row.adGroup.status : 'N/A';

      // Access ad group criterion data
      let criterionId = row.adGroupCriterion ? row.adGroupCriterion.criterionId : 'N/A';
      let criterionStatus = row.adGroupCriterion ? row.adGroupCriterion.status : 'N/A';

      // Access keyword data
      let keywordText = 'N/A';
      let matchType = 'N/A';
      if (row.adGroupCriterion && row.adGroupCriterion.keyword) {
        keywordText = row.adGroupCriterion.keyword.text || 'N/A';
        matchType = row.adGroupCriterion.keyword.matchType || 'N/A';
      }

      // Log first few keywords for debugging
      if (rowCount <= 5) {
        Logger.log(`  Ad Group Keyword ${rowCount}: "${keywordText}"`);
      }

      // Format match type for better readability
      let matchTypeFormatted = formatMatchType(matchType);

      // Add negative keyword data to a new row
      let newRow = [
        'Ad Group',
        campaignName,
        campaignStatus,
        adGroupName,
        adGroupStatus,
        criterionId,
        keywordText,
        matchTypeFormatted,
        matchType,
        criterionStatus
      ];

      data.push(newRow);

    } catch (e) {
      Logger.log("Error processing ad group negative row: " + e);
    }
  }

  Logger.log(`Processed ${rowCount} ad group negative rows, ${data.length} negative keywords found`);
  return data;
}

// Note: Negative keyword lists processing removed
// The Google Ads API doesn't provide a reliable way to extract individual keywords from lists
// We focus on campaign and ad group level negative keywords instead

function formatMatchType(matchType) {
  const matchTypeNames = {
    'EXACT': 'Exact',
    'PHRASE': 'Phrase',
    'BROAD': 'Broad'
  };
  return matchTypeNames[matchType] || matchType;
}

function sortData(data) {
  return data.sort((a, b) => {
    // Primary sort: type (Campaign, Ad Group)
    const typeA = a[0] || '';
    const typeB = b[0] || '';
    
    if (typeA !== typeB) {
      return typeA.localeCompare(typeB);
    }
    
    // Secondary sort: campaign name (ascending)
    const campaignA = a[1] || '';
    const campaignB = b[1] || '';
    
    if (campaignA !== campaignB) {
      return campaignA.localeCompare(campaignB);
    }
    
    // Tertiary sort: ad group name (ascending) - only for Ad Group type
    if (typeA === 'Ad Group') {
      const adGroupA = a[3] || '';
      const adGroupB = b[3] || '';
      
      if (adGroupA !== adGroupB) {
        return adGroupA.localeCompare(adGroupB);
      }
    }
    
    // Quaternary sort: negative keyword text (ascending)
    const keywordA = a[6] || '';
    const keywordB = b[6] || '';
    
    return keywordA.localeCompare(keywordB);
  });
}

function updateAccountSpreadsheet(spreadsheetUrl, negativeKeywordsData, accountName) {
  try {
    Logger.log("Updating account spreadsheet with negative keywords data...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Create or get the tab for negative keywords data
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
      'Type',
      'Campaign Name',
      'Campaign Status',
      'Ad Group Name',
      'Ad Group Status',
      'Criterion ID',
      'Negative Keyword',
      'Match Type (Formatted)',
      'Match Type (Raw)',
      'Criterion Status'
    ];
    
    // Prepare all data for bulk write
    let allData = [headers];
    
    if (negativeKeywordsData.length > 0) {
      allData = allData.concat(negativeKeywordsData);
    } else {
      // Add a row indicating no data found with proper column count
      const noDataRow = ['No negative keywords found in the account'];
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
    
    Logger.log(`✓ Successfully updated ${TAB} sheet with ${negativeKeywordsData.length} negative keyword records`);
    Logger.log(`✓ Account: ${accountName}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}
