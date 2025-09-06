// --- MCC LOCATIONS REPORT CONFIGURATION ---

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
const TAB = 'LocationsData';

// --- END OF CONFIGURATION ---

// Query for targeted locations (campaign settings)
const QUERY = `
  SELECT 
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    campaign_criterion.location.geo_target_constant,
    campaign_criterion.bid_modifier,
    campaign_criterion.status
  FROM campaign_criterion
  WHERE 
    campaign_criterion.type = 'LOCATION'
    AND campaign_criterion.negative = FALSE
    AND campaign.status IN ('ENABLED', 'PAUSED')
  ORDER BY campaign.name ASC
`;

function main() {
  Logger.log("=== Starting MCC Locations Report ===");
  
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

      // Get locations data for this account
      const locationsData = getLocationsDataForAccount();
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, locationsData, accountName);
      
      Logger.log(`✓ Successfully updated locations data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC LOCATIONS REPORT COMPLETED ===`);
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

function getLocationsDataForAccount() {
  Logger.log("Getting locations data for current account...");
  
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
      if (sampleRow.campaignCriterion) {
        Logger.log("Sample campaignCriterion object: " + JSON.stringify(sampleRow.campaignCriterion));
      }
    } else {
      Logger.log("Query returned no rows for sample check.");
    }

    // Process the main query results
    const rows = AdsApp.search(QUERY);
    const data = calculateMetrics(rows);
    
    // Sort data by campaign name then by criterion status
    const sortedData = sortData(data);
    
    Logger.log(`✓ Processed ${data.length} targeted locations for current account`);
    return sortedData;
    
  } catch (error) {
    Logger.log(`❌ Error getting locations data: ${error.message}`);
    throw error;
  }
}

function calculateMetrics(rows) {
  let data = [];
  let locationIdData = {};
  let totalRows = 0;
  
  // First pass: collect all location IDs to get names
  Logger.log("First pass: collecting targeted location IDs...");
  const tempRows = [];
  
  while (rows.hasNext()) {
    try {
      const row = rows.next();
      tempRows.push(row);
      
      const campaignCriterion = row.campaignCriterion || {};
      const location = campaignCriterion.location || {};
      const locationId = location.geoTargetConstant || 'Unknown';
      
      if (locationId && locationId !== 'Unknown') {
        // Extract just the ID number from geoTargetConstants/2840 format
        const cleanLocationId = locationId.replace('geoTargetConstants/', '');
        locationIdData[cleanLocationId] = true;
      }
      
    } catch (error) {
      Logger.log(`Error in first pass: ${error.message}`);
    }
  }
  
  // Get location names from IDs
  Logger.log(`Getting location names for ${Object.keys(locationIdData).length} targeted location IDs...`);
  const locationNames = getLocationNamesFromIds(Object.keys(locationIdData));
  
  // Second pass: process data with location names
  Logger.log("Second pass: processing targeted location data with names...");
  
  tempRows.forEach(row => {
    try {
      const campaign = row.campaign || {};
      const campaignCriterion = row.campaignCriterion || {};
      const location = campaignCriterion.location || {};
      
      const campaignId = campaign.id || 'Unknown';
      const campaignName = campaign.name || 'Unknown Campaign';
      const campaignType = campaign.advertisingChannelType || 'Unknown';
      const campaignStatus = campaign.status || 'Unknown';
      const locationId = location.geoTargetConstant || 'Unknown';
      const bidModifier = campaignCriterion.bidModifier || 0;
      const criterionStatus = campaignCriterion.status || 'Unknown';
      
      // Extract just the ID number from geoTargetConstants/2840 format
      const cleanLocationId = locationId.replace('geoTargetConstants/', '');
      
      // Get location name
      const locationName = locationNames[cleanLocationId] || `Unknown (ID: ${cleanLocationId})`;
      
      let newRow = [
        campaignName,
        campaignType,
        campaignStatus,
        cleanLocationId,
        locationName,
        bidModifier,
        criterionStatus
      ];
      
      data.push(newRow);
      totalRows++;
      
    } catch (error) {
      Logger.log(`Error processing targeted location row: ${error.message}`);
    }
  });
  
  Logger.log(`✓ Processed ${totalRows} targeted location records`);
  Logger.log(`✓ Total unique targeted locations: ${Object.keys(locationNames).length}`);
  
  return data;
}

function getLocationNamesFromIds(locationIds) {
  try {
    const locationNames = {};
    
    if (locationIds.length === 0) {
      return locationNames;
    }
    
    // Query geo_target_constant to get names from IDs - All location types
    const locationIdsList = locationIds.map(id => `"${id}"`).join(',');
    const query = `
      SELECT
          geo_target_constant.id,
          geo_target_constant.name,
          geo_target_constant.target_type
      FROM geo_target_constant
      WHERE geo_target_constant.id IN (${locationIdsList})
    `;
    
    Logger.log(`Querying geo_target_constant for ${locationIds.length} location IDs...`);
    
    const rows = AdsApp.search(query);
    while (rows.hasNext()) {
      const row = rows.next();
      const geoTargetConstant = row.geoTargetConstant || {};
      const id = geoTargetConstant.id || '';
      const name = geoTargetConstant.name || '';
      const targetType = geoTargetConstant.targetType || '';
      
      // Include all location types
      if (id && name) {
        locationNames[id] = `${name} (${targetType})`;
        Logger.log(`Mapped location ID ${id} to name: ${name} (${targetType})`);
      }
    }
    
    Logger.log(`Successfully mapped ${Object.keys(locationNames).length} location IDs to names`);
    return locationNames;
    
  } catch (error) {
    Logger.log(`Error getting location names from IDs: ${error.message}`);
    return {};
  }
}

function sortData(data) {
  return data.sort((a, b) => {
    // Primary sort: campaign name (ascending)
    const campaignNameA = a[0] || '';
    const campaignNameB = b[0] || '';
    
    if (campaignNameA !== campaignNameB) {
      return campaignNameA.localeCompare(campaignNameB);
    }
    
    // Secondary sort: criterion status (Enabled, Paused, Removed)
    const statusA = a[6] || ''; // criterion status is at index 6
    const statusB = b[6] || '';
    
    const statusOrder = { 'ENABLED': 1, 'PAUSED': 2, 'REMOVED': 3 };
    const orderA = statusOrder[statusA] || 4;
    const orderB = statusOrder[statusB] || 4;
    
    return orderA - orderB; // Ascending order (Enabled first)
  });
}

function updateAccountSpreadsheet(spreadsheetUrl, locationsData, accountName) {
  try {
    Logger.log("Updating account spreadsheet with locations data...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Create or get the tab for locations data
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
      'Location ID',
      'Location Name',
      'Bid Modifier',
      'Criterion Status'
    ];
    
    // Prepare all data for bulk write
    let allData = [headers];
    
    if (locationsData.length > 0) {
      allData = allData.concat(locationsData);
    } else {
      // Add a row indicating no data found
      allData.push(['No targeted locations found']);
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
    
    Logger.log(`✓ Successfully updated ${TAB} sheet with ${locationsData.length} targeted locations`);
    Logger.log(`✓ Account: ${accountName}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}
