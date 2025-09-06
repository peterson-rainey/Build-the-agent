// --- MCC PERFORMANCE MAX CHANNEL PERFORMANCE REPORT CONFIGURATION ---

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
const TAB = 'PMAXChannelPerformance';

// --- END OF CONFIGURATION ---

const QUERY = `
  SELECT 
    campaign.name,
    campaign.status,
    campaign.advertising_channel_type,
    performance_max_placement_view.display_name,
    performance_max_placement_view.placement,
    performance_max_placement_view.placement_type,
    performance_max_placement_view.target_url,
    metrics.impressions
  FROM performance_max_placement_view 
  WHERE segments.date DURING LAST_30_DAYS
  AND campaign.advertising_channel_type = "PERFORMANCE_MAX"
  ORDER BY campaign.name, performance_max_placement_view.display_name
`;

function main() {
  Logger.log("=== Starting MCC Performance Max Channel Performance Report ===");
  Logger.log(`Script version: 1.0`);
  Logger.log(`Timestamp: ${new Date().toISOString()}`);
  
  if (SINGLE_CID_FOR_TESTING) {
    Logger.log(`Running in test mode for CID: ${SINGLE_CID_FOR_TESTING}`);
  } else {
    Logger.log("Running for all accounts in the master spreadsheet.");
  }
  
  // Validate the query before processing accounts
  Logger.log("Validating GAQL query structure...");
  validateQuery();

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

      // Get Performance Max channel performance data for this account
      const channelPerformanceData = getPMAXChannelPerformanceDataForAccount();
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, channelPerformanceData, accountName);
      
      Logger.log(`✓ Successfully updated Performance Max channel performance data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC PERFORMANCE MAX CHANNEL PERFORMANCE REPORT COMPLETED ===`);
  Logger.log(`Total accounts processed: ${processedCount}`);
  Logger.log(`Successful updates: ${successCount}`);
  Logger.log(`Errors: ${errorCount}`);
}

function validateQuery() {
  Logger.log("=== QUERY VALIDATION ===");
  Logger.log("Validating GAQL query structure and fields...");
  
  // Check if required fields are present
  const requiredFields = [
    'campaign.name',
    'campaign.status', 
    'campaign.advertising_channel_type',
    'performance_max_placement_view.display_name',
    'performance_max_placement_view.placement',
    'performance_max_placement_view.placement_type',
    'performance_max_placement_view.target_url',
    'metrics.impressions'
  ];
  
  Logger.log("Required fields for this query:");
  requiredFields.forEach(field => {
    if (QUERY.includes(field)) {
      Logger.log(`✓ ${field} - Found`);
    } else {
      Logger.log(`❌ ${field} - Missing`);
    }
  });
  
  // Check for common query issues
  if (!QUERY.includes('FROM performance_max_placement_view')) {
    Logger.log("❌ Query must use 'FROM performance_max_placement_view' resource");
  } else {
    Logger.log("✓ Using correct 'FROM performance_max_placement_view' resource");
  }
  
  if (!QUERY.includes('segments.date DURING LAST_30_DAYS')) {
    Logger.log("❌ Query must include date range filter");
  } else {
    Logger.log("✓ Date range filter present");
  }
  
  if (!QUERY.includes('campaign.advertising_channel_type = "PERFORMANCE_MAX"')) {
    Logger.log("❌ Query must filter for Performance Max campaigns");
  } else {
    Logger.log("✓ Performance Max filter present");
  }
  
  if (!QUERY.includes('performance_max_placement_view.display_name')) {
    Logger.log("❌ Query must include placement display name for channel data");
  } else {
    Logger.log("✓ Placement display name present");
  }
  
  if (!QUERY.includes('metrics.impressions')) {
    Logger.log("❌ Query must include impressions metric");
  } else {
    Logger.log("✓ Impressions metric present");
  }
  
  Logger.log("=== END QUERY VALIDATION ===");
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

function getPMAXChannelPerformanceDataForAccount() {
  Logger.log("Getting Performance Max channel performance data for current account...");
  Logger.log(`Account: ${AdsApp.currentAccount().getCustomerId()}`);
  Logger.log(`Account Name: ${AdsApp.currentAccount().getName()}`);
  
  try {
    // Log the exact query being executed
    Logger.log("=== EXECUTING GAQL QUERY ===");
    Logger.log(QUERY);
    Logger.log("=== END QUERY ===");
    
    // First, test the query with a sample to check for syntax errors
    Logger.log("Testing query syntax with sample data...");
    const sampleQuery = QUERY + ' LIMIT 1';
    Logger.log("Sample query: " + sampleQuery);
    
    try {
      const sampleRows = AdsApp.search(sampleQuery);
      Logger.log("✓ Query syntax is valid");
      
      if (sampleRows.hasNext()) {
        const sampleRow = sampleRows.next();
        Logger.log("=== SAMPLE ROW STRUCTURE ===");
        Logger.log("Full sample row: " + JSON.stringify(sampleRow));
        
        if (sampleRow.metrics) {
          Logger.log("Sample metrics object: " + JSON.stringify(sampleRow.metrics));
        } else {
          Logger.log("⚠️ No metrics object found in sample row");
        }
        
        if (sampleRow.campaign) {
          Logger.log("Sample campaign object: " + JSON.stringify(sampleRow.campaign));
        } else {
          Logger.log("⚠️ No campaign object found in sample row");
        }
        
        if (sampleRow.segments) {
          Logger.log("Sample segments object: " + JSON.stringify(sampleRow.segments));
        } else {
          Logger.log("⚠️ No segments object found in sample row");
        }
        
        Logger.log("=== END SAMPLE ROW ===");
      } else {
        Logger.log("⚠️ Sample query returned no rows - this might indicate no Performance Max campaigns or no data in the date range");
      }
    } catch (sampleError) {
      Logger.log(`❌ Sample query failed: ${sampleError.message}`);
      Logger.log("This indicates a query syntax error. Please check the GAQL query structure.");
      throw sampleError;
    }

    // Process the main query results
    Logger.log("Executing main query for all data...");
    const rows = AdsApp.search(QUERY);
    Logger.log("✓ Main query executed successfully");
    
    const data = calculateChannelPerformanceMetrics(rows);
    Logger.log(`✓ Processed ${data.length} channel performance rows for current account`);
    
    // Log summary of data found
    if (data.length > 0) {
      const campaigns = [...new Set(data.map(row => row[0]))];
      const channels = [...new Set(data.map(row => row[3]))];
      Logger.log(`Found data for ${campaigns.length} campaigns: ${campaigns.join(', ')}`);
      Logger.log(`Found data for ${channels.length} channels: ${channels.join(', ')}`);
    } else {
      Logger.log("⚠️ No channel performance data found for this account");
    }
    
    // Sort data by campaign name, then by channel
    const sortedData = sortData(data);
    Logger.log("✓ Data sorted successfully");
    
    return sortedData;
    
  } catch (error) {
    Logger.log(`❌ Error getting Performance Max channel performance data: ${error.message}`);
    Logger.log(`Error type: ${error.name}`);
    Logger.log(`Error stack: ${error.stack}`);
    throw error;
  }
}

function calculateChannelPerformanceMetrics(rows) {
  Logger.log("Starting to process channel performance metrics...");
  let data = [];
  let rowCount = 0;
  let skippedRows = 0;
  let errorRows = 0;

  while (rows.hasNext()) {
    try {
      let row = rows.next();
      rowCount++;
      
      if (rowCount <= 3) {
        Logger.log(`=== PROCESSING ROW ${rowCount} ===`);
        Logger.log("Raw row data: " + JSON.stringify(row));
      }

      // Access dimensions using correct nested object structure
      let campaignName = row.campaign ? row.campaign.name : 'N/A';
      let campaignStatus = row.campaign ? row.campaign.status : 'UNKNOWN';
      let campaignType = row.campaign ? row.campaign.advertisingChannelType : 'UNKNOWN';
      
      // Access placement information
      let displayName = 'UNKNOWN';
      let placement = 'UNKNOWN';
      let placementType = 'UNKNOWN';
      let targetUrl = 'UNKNOWN';
      
      if (row.performanceMaxPlacementView) {
        displayName = row.performanceMaxPlacementView.displayName || 'UNKNOWN';
        placement = row.performanceMaxPlacementView.placement || 'UNKNOWN';
        placementType = row.performanceMaxPlacementView.placementType || 'UNKNOWN';
        targetUrl = row.performanceMaxPlacementView.targetUrl || 'UNKNOWN';
      } else {
        Logger.log(`⚠️ Row ${rowCount}: No performanceMaxPlacementView found`);
      }

      // Map placement to friendly channel names
      let channelName = mapPlacementToChannel(displayName, placement, placementType);

      // Access metrics nested within the 'metrics' object
      const metrics = row.metrics || {};
      
      if (rowCount <= 3) {
        Logger.log(`Row ${rowCount} metrics: ${JSON.stringify(metrics)}`);
      }
      
      let impressions = Number(metrics.impressions) || 0;

      if (rowCount <= 3) {
        Logger.log(`Row ${rowCount} processed values: Campaign=${campaignName}, Channel=${channelName}, DisplayName=${displayName}, Placement=${placement}, Impressions=${impressions}`);
      }

      // Skip rows with zero impressions
      if (impressions === 0) {
        skippedRows++;
        if (rowCount <= 10) {
          Logger.log(`Skipping row ${rowCount} - zero impressions`);
        }
        continue;
      }

      // Create one row per placement with its performance metrics
      let newRow = [
        campaignName,
        campaignType,
        campaignStatus,
        channelName,
        displayName,
        placement,
        placementType,
        targetUrl,
        impressions
      ];
      data.push(newRow);

    } catch (e) {
      errorRows++;
      Logger.log(`❌ Error processing row ${rowCount}: ${e.message}`);
      Logger.log("Problematic row data: " + JSON.stringify(row));
      // Continue with next row
    }
  }

  Logger.log(`=== DATA PROCESSING SUMMARY ===`);
  Logger.log(`Total rows processed: ${rowCount}`);
  Logger.log(`Successful rows: ${data.length}`);
  Logger.log(`Skipped rows (zero impressions): ${skippedRows}`);
  Logger.log(`Error rows: ${errorRows}`);
  
  if (data.length > 0) {
    Logger.log(`Sample successful row: ${JSON.stringify(data[0])}`);
  }
  
  return data;
}

function mapPlacementToChannel(displayName, placement, placementType) {
  // Map placement values to friendly channel names based on placement patterns
  const displayStr = displayName.toString().toLowerCase();
  const placementStr = placement.toString().toLowerCase();
  const typeStr = placementType.toString().toLowerCase();
  
  // Check for specific channel patterns in display name, placement, or type
  if (displayStr.includes('search') || placementStr.includes('search') || typeStr.includes('search')) {
    return 'Search';
  } else if (displayStr.includes('youtube') || placementStr.includes('youtube') || typeStr.includes('youtube')) {
    return 'YouTube';
  } else if (displayStr.includes('display') || placementStr.includes('display') || typeStr.includes('display')) {
    return 'Display';
  } else if (displayStr.includes('discover') || placementStr.includes('discover') || typeStr.includes('discover')) {
    return 'Discover';
  } else if (displayStr.includes('gmail') || placementStr.includes('gmail') || typeStr.includes('gmail')) {
    return 'Gmail';
  } else if (displayStr.includes('maps') || placementStr.includes('maps') || typeStr.includes('maps')) {
    return 'Maps';
  } else if (displayStr.includes('shopping') || placementStr.includes('shopping') || typeStr.includes('shopping')) {
    return 'Shopping';
  } else if (displayStr.includes('partner') || placementStr.includes('partner') || typeStr.includes('partner')) {
    return 'Search Partners';
  } else {
    // If we can't determine the channel, return the display name or placement
    return displayName !== 'UNKNOWN' ? displayName : placement;
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
    
    // Secondary sort: channel name (ascending)
    const channelA = a[3] || '';
    const channelB = b[3] || '';
    
    return channelA.localeCompare(channelB);
  });
}

function updateAccountSpreadsheet(spreadsheetUrl, channelPerformanceData, accountName) {
  try {
    Logger.log("Updating account spreadsheet with Performance Max channel performance data...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Create or get the tab for channel performance data
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
      'Channel',
      'Display Name',
      'Placement',
      'Placement Type',
      'Target URL',
      'Impressions'
    ];
    
    // Prepare all data for bulk write
    let allData = [headers];
    
    if (channelPerformanceData.length > 0) {
      allData = allData.concat(channelPerformanceData);
    } else {
      // Add a row indicating no data found with proper column count
      let noDataRow = new Array(headers.length).fill('');
      noDataRow[0] = 'No Performance Max channel performance data found for the last 30 days';
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
    
    Logger.log(`✓ Successfully updated ${TAB} sheet with ${channelPerformanceData.length} channel performance rows`);
    Logger.log(`✓ Account: ${accountName}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}
