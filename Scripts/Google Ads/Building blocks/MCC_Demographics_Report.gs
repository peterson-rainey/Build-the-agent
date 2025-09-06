// --- MCC NON-PERFORMANCE MAX DEMOGRAPHICS REPORT CONFIGURATION ---

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
const TAB = 'DemographicsData';

// --- END OF CONFIGURATION ---

// Age demographics query for non-Performance Max campaigns
const AGE_QUERY = `
  SELECT 
    campaign.name,
    campaign.status,
    campaign.advertising_channel_type,
    ad_group.name,
    ad_group.status,
    ad_group_criterion.age_range.type,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM age_range_view
  WHERE segments.date DURING LAST_30_DAYS
    AND campaign.advertising_channel_type != "PERFORMANCE_MAX"
    AND campaign.status IN ('ENABLED', 'PAUSED')
    AND ad_group.status IN ('ENABLED', 'PAUSED')
    AND metrics.impressions > 0
  ORDER BY metrics.impressions DESC
`;

// Gender demographics query for non-Performance Max campaigns
const GENDER_QUERY = `
  SELECT 
    campaign.name,
    campaign.status,
    campaign.advertising_channel_type,
    ad_group.name,
    ad_group.status,
    ad_group_criterion.gender.type,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM gender_view
  WHERE segments.date DURING LAST_30_DAYS
    AND campaign.advertising_channel_type != "PERFORMANCE_MAX"
    AND campaign.status IN ('ENABLED', 'PAUSED')
    AND ad_group.status IN ('ENABLED', 'PAUSED')
    AND metrics.impressions > 0
  ORDER BY metrics.impressions DESC
`;

// Income demographics query for non-Performance Max campaigns
const INCOME_QUERY = `
  SELECT 
    campaign.name,
    campaign.status,
    campaign.advertising_channel_type,
    ad_group.name,
    ad_group.status,
    ad_group_criterion.income_range.type,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM income_range_view
  WHERE segments.date DURING LAST_30_DAYS
    AND campaign.advertising_channel_type != "PERFORMANCE_MAX"
    AND campaign.status IN ('ENABLED', 'PAUSED')
    AND ad_group.status IN ('ENABLED', 'PAUSED')
    AND metrics.impressions > 0
  ORDER BY metrics.impressions DESC
`;

function main() {
  Logger.log("=== Starting MCC Non-Performance Max Demographics Report ===");
  
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

      // Get demographics data for this account
      const demographicsData = getDemographicsDataForAccount();
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, demographicsData, accountName);
      
      Logger.log(`✓ Successfully updated demographics data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC NON-PERFORMANCE MAX DEMOGRAPHICS REPORT COMPLETED ===`);
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

function getDemographicsDataForAccount() {
  Logger.log("Getting demographics data for current account...");
  
  try {
    const allDemographics = [];
    
    // Get age demographics
    Logger.log("Getting age demographics...");
    try {
      const ageRows = AdsApp.search(AGE_QUERY);
      const ageData = calculateDemographicsMetrics(ageRows, 'Age');
      allDemographics.push(...ageData);
      Logger.log(`✓ Found ${ageData.length} age demographic records`);
    } catch (error) {
      Logger.log(`❌ Error getting age demographics: ${error.message}`);
    }
    
    // Get gender demographics
    Logger.log("Getting gender demographics...");
    try {
      const genderRows = AdsApp.search(GENDER_QUERY);
      const genderData = calculateDemographicsMetrics(genderRows, 'Gender');
      allDemographics.push(...genderData);
      Logger.log(`✓ Found ${genderData.length} gender demographic records`);
    } catch (error) {
      Logger.log(`❌ Error getting gender demographics: ${error.message}`);
    }
    
    // Get income demographics
    Logger.log("Getting income demographics...");
    try {
      const incomeRows = AdsApp.search(INCOME_QUERY);
      const incomeData = calculateDemographicsMetrics(incomeRows, 'Income');
      allDemographics.push(...incomeData);
      Logger.log(`✓ Found ${incomeData.length} income demographic records`);
    } catch (error) {
      Logger.log(`❌ Error getting income demographics: ${error.message}`);
    }
    
    // Sort data by conversions (descending) then by status (Enabled, Paused, Removed)
    const sortedData = sortData(allDemographics);
    
    Logger.log(`✓ Processed ${allDemographics.length} total demographics records for current account`);
    return sortedData;
    
  } catch (error) {
    Logger.log(`❌ Error getting demographics data: ${error.message}`);
    throw error;
  }
}

function calculateDemographicsMetrics(rows, demographicType) {
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

      // Access ad_group_criterion for the specific demographic type
      let demographicValue = 'Unknown';
      if (row.adGroupCriterion) {
        if (demographicType === 'Age' && row.adGroupCriterion.ageRange) {
          demographicValue = row.adGroupCriterion.ageRange.type || 'Unknown';
        } else if (demographicType === 'Gender' && row.adGroupCriterion.gender) {
          demographicValue = row.adGroupCriterion.gender.type || 'Unknown';
        } else if (demographicType === 'Income' && row.adGroupCriterion.incomeRange) {
          demographicValue = row.adGroupCriterion.incomeRange.type || 'Unknown';
        }
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

      // Skip records with zero impressions or unknown demographic value
      if (impressions === 0 || demographicValue === 'Unknown') {
        continue;
      }

      // Create row for this demographic record
      let newRow = [
        campaignName,
        campaignType,
        campaignStatus,
        adGroupName,
        adGroupStatus,
        demographicType,
        demographicValue,
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

  Logger.log(`Processed ${rowCount} total ${demographicType} rows, ${data.length} successful records`);
  return data;
}

function sortData(data) {
  return data.sort((a, b) => {
    // Primary sort: conversions (descending)
    const conversionsA = a[10] || 0; // conversions is at index 10
    const conversionsB = b[10] || 0;
    
    if (conversionsA !== conversionsB) {
      return conversionsB - conversionsA; // Descending order
    }
    
    // Secondary sort: campaign status (Enabled, Paused, Removed)
    const statusA = a[2] || ''; // campaign status is at index 2
    const statusB = b[2] || '';
    
    const statusOrder = { 'ENABLED': 1, 'PAUSED': 2, 'REMOVED': 3 };
    const orderA = statusOrder[statusA] || 4;
    const orderB = statusOrder[statusB] || 4;
    
    return orderA - orderB; // Ascending order (Enabled first)
  });
}

function updateAccountSpreadsheet(spreadsheetUrl, demographicsData, accountName) {
  try {
    Logger.log("Updating account spreadsheet with demographics data...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Create or get the tab for demographics data
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
      'Demographic Type',
      'Demographic Value',
      'Impressions',
      'Clicks',
      'Cost',
      'Conversions',
      'Conversion Value'
    ];
    
    // Prepare all data for bulk write
    let allData = [headers];
    
    if (demographicsData.length > 0) {
      allData = allData.concat(demographicsData);
    } else {
      // Add a row indicating no data found with proper column count
      const noDataRow = ['No demographics data found for the last 30 days'];
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
    
    Logger.log(`✓ Successfully updated ${TAB} sheet with ${demographicsData.length} demographics records`);
    Logger.log(`✓ Account: ${accountName}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}