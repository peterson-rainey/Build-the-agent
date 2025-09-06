// --- MCC PERFORMANCE MAX ASSET GROUPS REPORT CONFIGURATION ---

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
const TAB = 'AssetGroupsAssetData';

// --- END OF CONFIGURATION ---

const QUERY = `
  SELECT 
    campaign.name,
    campaign.status,
    campaign.advertising_channel_type,
    asset_group.name,
    asset_group.status,
    asset_group_asset.asset,
    asset_group_asset.field_type,
    asset_group_asset.performance_label,
    asset_group_asset.asset.id,
    asset_group_asset.asset.name,
    asset_group_asset.asset.type,
    asset_group_asset.asset.source,
    asset_group_asset.asset.text_asset.text,
    asset_group_asset.asset.image_asset.file_size,
    asset_group_asset.asset.image_asset.mime_type,
    asset_group_asset.asset.image_asset.full_size.url,
    asset_group_asset.asset.image_asset.full_size.height_pixels,
    asset_group_asset.asset.image_asset.full_size.width_pixels,
    asset_group_asset.asset.video_asset.youtube_video_id,
    asset_group_asset.asset.video_asset.duration_millis,
    asset_group_asset.asset.youtube_video_asset.youtube_video_id,
    asset_group_asset.asset.youtube_video_asset.youtube_video_title,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM asset_group_asset 
  WHERE segments.date DURING LAST_30_DAYS
  AND campaign.advertising_channel_type = "PERFORMANCE_MAX"
  AND asset_group_asset.field_type NOT IN ("BUSINESS_NAME", "CALL_TO_ACTION_SELECTION")
  ORDER BY metrics.impressions DESC
`;

function main() {
  Logger.log("=== Starting MCC Performance Max Asset Groups Report ===");
  
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

      // Get Performance Max asset groups data for this account
      const assetGroupsData = getPMAXAssetGroupsDataForAccount();
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, assetGroupsData, accountName);
      
      Logger.log(`✓ Successfully updated Performance Max asset groups data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC PERFORMANCE MAX ASSET GROUPS REPORT COMPLETED ===`);
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

function getAssetContent(asset) {
  // Handle different asset types following Mike Scripts pattern
  switch (asset.type) {
    case 'TEXT':
      return asset.textAsset && asset.textAsset.text ? asset.textAsset.text : '';
    case 'IMAGE':
      return asset.imageAsset && asset.imageAsset.fullSize && asset.imageAsset.fullSize.url ? asset.imageAsset.fullSize.url : '';
    case 'VIDEO':
      if (asset.videoAsset && asset.videoAsset.youtubeVideoId) {
        return `https://www.youtube.com/watch?v=${asset.videoAsset.youtubeVideoId}`;
      }
      return '';
    case 'YOUTUBE_VIDEO':
      if (asset.youtubeVideoAsset && asset.youtubeVideoAsset.youtubeVideoId) {
        return `https://www.youtube.com/watch?v=${asset.youtubeVideoAsset.youtubeVideoId}`;
      }
      return '';
    default:
      return '';
  }
}

function getPMAXAssetGroupsDataForAccount() {
  Logger.log("Getting Performance Max asset groups data for current account...");
  
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
      if (sampleRow.assetGroup) {
        Logger.log("Sample assetGroup object: " + JSON.stringify(sampleRow.assetGroup));
      }
      if (sampleRow.assetGroupAsset) {
        Logger.log("Sample assetGroupAsset object: " + JSON.stringify(sampleRow.assetGroupAsset));
      }
    } else {
      Logger.log("Query returned no rows for sample check.");
    }

    // Process the main query results
    const rows = AdsApp.search(QUERY);
    const data = calculateAssetGroupsMetrics(rows);
    
    // Sort data by impressions (descending) then by status (Enabled, Paused, Removed)
    const sortedData = sortData(data);
    
    Logger.log(`✓ Processed ${data.length} Performance Max assets for current account`);
    return sortedData;
    
  } catch (error) {
    Logger.log(`❌ Error getting Performance Max asset groups data: ${error.message}`);
    throw error;
  }
}

function calculateAssetGroupsMetrics(rows) {
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
      let assetGroupName = row.assetGroup ? row.assetGroup.name : 'N/A';
      let assetGroupStatus = row.assetGroup ? row.assetGroup.status : 'UNKNOWN';

      // Access asset information
      let assetId = 'N/A';
      let assetName = 'N/A';
      let assetType = 'N/A';
      let assetSource = 'N/A';
      let assetContent = 'N/A';
      let fieldType = 'N/A';
      let performanceLabel = 'N/A';

      if (row.assetGroupAsset) {
        fieldType = row.assetGroupAsset.fieldType || 'N/A';
        performanceLabel = row.assetGroupAsset.performanceLabel || 'N/A';
        
        if (row.assetGroupAsset.asset) {
          assetId = row.assetGroupAsset.asset.id || 'N/A';
          assetName = row.assetGroupAsset.asset.name || 'N/A';
          assetType = row.assetGroupAsset.asset.type || 'N/A';
          assetSource = row.assetGroupAsset.asset.source || 'N/A';
          
          // Get asset content based on asset type (following Mike Scripts pattern)
          assetContent = getAssetContent(row.assetGroupAsset.asset);
        }
      }

      // Determine combined status - if any component is paused, show as paused
      let combinedStatus = 'ENABLED';
      if (campaignStatus === 'PAUSED' || assetGroupStatus === 'PAUSED') {
        combinedStatus = 'PAUSED';
      } else if (campaignStatus === 'REMOVED' || assetGroupStatus === 'REMOVED') {
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

      // Skip assets with zero impressions
      if (impressions === 0) {
        continue;
      }

      // Create one row per asset with its individual performance metrics
      let newRow = [
        campaignName,
        campaignType,
        assetGroupName,
        combinedStatus,
        assetId,
        assetName,
        assetType,
        assetSource,
        fieldType,
        performanceLabel,
        impressions,
        clicks,
        cost,
        conversions,
        conversionValue,
        assetContent
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
    const impressionsA = a[10] || 0; // impressions is at index 10
    const impressionsB = b[10] || 0;
    
    if (impressionsA !== impressionsB) {
      return impressionsB - impressionsA; // Descending order
    }
    
    // Secondary sort: status (Enabled, Paused, Removed)
    const statusA = a[3] || ''; // status is at index 3
    const statusB = b[3] || '';
    
    const statusOrder = { 'ENABLED': 1, 'PAUSED': 2, 'REMOVED': 3 };
    const orderA = statusOrder[statusA] || 4;
    const orderB = statusOrder[statusB] || 4;
    
    return orderA - orderB; // Ascending order (Enabled first)
  });
}

function updateAccountSpreadsheet(spreadsheetUrl, assetGroupsData, accountName) {
  try {
    Logger.log("Updating account spreadsheet with Performance Max asset groups data...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Create or get the tab for asset groups data
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
      'Asset Group Name',
      'Status',
      'Asset ID',
      'Asset Name',
      'Asset Type',
      'Asset Source',
      'Field Type',
      'Performance Label',
      'Impressions',
      'Clicks',
      'Cost',
      'Conversions',
      'Conversion Value',
      'Asset Content'
    ];
    
    // Prepare all data for bulk write
    let allData = [headers];
    
    if (assetGroupsData.length > 0) {
      allData = allData.concat(assetGroupsData);
    } else {
      // Add a row indicating no data found
      allData.push(['No Performance Max asset groups data found for the last 30 days']);
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
    
    Logger.log(`✓ Successfully updated ${TAB} sheet with ${assetGroupsData.length} assets`);
    Logger.log(`✓ Account: ${accountName}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}
