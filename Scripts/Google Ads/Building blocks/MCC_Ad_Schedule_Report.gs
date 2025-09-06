// --- MCC AD SCHEDULE REPORT CONFIGURATION ---

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
const TAB = 'AdScheduleData';

// --- END OF CONFIGURATION ---

const QUERY = `
  SELECT 
    campaign.name,
    campaign.status,
    campaign.advertising_channel_type,
    segments.day_of_week,
    segments.hour,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM campaign 
  WHERE segments.date DURING LAST_30_DAYS
  AND campaign.status IN ('ENABLED', 'PAUSED')
  ORDER BY campaign.name, segments.day_of_week, segments.hour
`;

function main() {
  Logger.log("=== Starting MCC Ad Schedule Report ===");
  
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

      // Get ad schedule data for this account
      const adScheduleData = getAdScheduleDataForAccount();
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, adScheduleData, accountName);
      
      Logger.log(`✓ Successfully updated ad schedule data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC AD SCHEDULE REPORT COMPLETED ===`);
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

function getAdScheduleDataForAccount() {
  Logger.log("Getting ad schedule data for current account...");
  
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
      if (sampleRow.segments) {
        Logger.log("Sample segments object: " + JSON.stringify(sampleRow.segments));
      }
    } else {
      Logger.log("Query returned no rows for sample check.");
    }

    // Process the main query results
    const rows = AdsApp.search(QUERY);
    const data = calculateMetrics(rows);
    
    // Sort data by campaign, day of week, then hour
    const sortedData = sortData(data);
    
    Logger.log(`✓ Processed ${data.length} ad schedule records for current account`);
    return sortedData;
    
  } catch (error) {
    Logger.log(`❌ Error getting ad schedule data: ${error.message}`);
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

      // Access day of week and hour segments
      let dayOfWeek = row.segments ? row.segments.dayOfWeek : 'UNKNOWN';
      let hour = row.segments ? row.segments.hour : 'UNKNOWN';

      // Convert day of week number to day name
      let dayName = getDayName(dayOfWeek);
      
      // Format hour for better readability
      let hourFormatted = formatHour(hour);

      // Access metrics nested within the 'metrics' object
      const metrics = row.metrics || {};

      let impressions = Number(metrics.impressions) || 0;
      let clicks = Number(metrics.clicks) || 0;
      let costMicros = Number(metrics.costMicros) || 0;
      let conversions = Number(metrics.conversions) || 0;
      let conversionValue = Number(metrics.conversionsValue) || 0;

      // Convert cost from micros to actual currency
      let cost = costMicros / 1000000;

      // Add ad schedule data to a new row
      let newRow = [
        campaignName,
        campaignType,
        campaignStatus,
        dayName,
        dayOfWeek,
        hourFormatted,
        hour,
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

function getDayName(dayOfWeek) {
  const dayNames = {
    'MONDAY': 'Monday',
    'TUESDAY': 'Tuesday', 
    'WEDNESDAY': 'Wednesday',
    'THURSDAY': 'Thursday',
    'FRIDAY': 'Friday',
    'SATURDAY': 'Saturday',
    'SUNDAY': 'Sunday'
  };
  return dayNames[dayOfWeek] || dayOfWeek;
}

function formatHour(hour) {
  if (hour === 'UNKNOWN' || hour === null || hour === undefined) {
    return 'Unknown';
  }
  
  let hourNum = parseInt(hour);
  if (isNaN(hourNum)) {
    return hour;
  }
  
  // Format as 12-hour time with AM/PM
  if (hourNum === 0) {
    return '12:00 AM';
  } else if (hourNum < 12) {
    return `${hourNum}:00 AM`;
  } else if (hourNum === 12) {
    return '12:00 PM';
  } else {
    return `${hourNum - 12}:00 PM`;
  }
}

function sortData(data) {
  return data.sort((a, b) => {
    // Primary sort: campaign name (ascending)
    const campaignA = a[0] || '';
    const campaignB = b[0] || '';
    
    if (campaignA !== campaignB) {
      return campaignA.localeCompare(campaignB);
    }
    
    // Secondary sort: day of week number (ascending)
    const dayA = a[4] || 0; // day of week number is at index 4
    const dayB = b[4] || 0;
    
    if (dayA !== dayB) {
      return dayA - dayB;
    }
    
    // Tertiary sort: hour (ascending)
    const hourA = a[6] || 0; // hour number is at index 6
    const hourB = b[6] || 0;
    
    return hourA - hourB;
  });
}

function updateAccountSpreadsheet(spreadsheetUrl, adScheduleData, accountName) {
  try {
    Logger.log("Updating account spreadsheet with ad schedule data...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Create or get the tab for ad schedule data
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
      'Day of Week',
      'Day Number',
      'Hour (12h)',
      'Hour (24h)',
      'Impressions',
      'Clicks',
      'Cost',
      'Conversions',
      'Conversion Value'
    ];
    
    // Prepare all data for bulk write
    let allData = [headers];
    
    if (adScheduleData.length > 0) {
      allData = allData.concat(adScheduleData);
    } else {
      // Add a row indicating no data found with proper column count
      const noDataRow = ['No ad schedule data found for the last 30 days'];
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
    
    Logger.log(`✓ Successfully updated ${TAB} sheet with ${adScheduleData.length} ad schedule records`);
    Logger.log(`✓ Account: ${accountName}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}
