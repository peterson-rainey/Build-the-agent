// --- MCC ASSETS REPORT CONFIGURATION ---
// This script reports on specific asset types across the entire account:
// - Sitelinks, Callouts, Structured Snippets
// Excludes: All other asset types including Headlines, Descriptions, Business Name, 
// Business Logo, Message Assets, Call Assets, Lead Form Assets, Location Assets, 
// Price Assets, Promotion Assets

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
const TAB = 'AssetsData';

// --- END OF CONFIGURATION ---

const QUERY = `
  SELECT 
    campaign_asset.asset,
    campaign_asset.field_type,
    asset.text_asset.text,
    asset.sitelink_asset.description1,
    asset.sitelink_asset.description2,
    asset.callout_asset.callout_text,
    asset.structured_snippet_asset.header,
    asset.structured_snippet_asset.values,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM campaign_asset 
  WHERE segments.date DURING LAST_30_DAYS
  AND campaign_asset.field_type IN (
    'SITELINK',
    'CALLOUT',
    'STRUCTURED_SNIPPET'
  )
  ORDER BY metrics.cost_micros DESC
`;

function main() {
  Logger.log("=== Starting MCC Assets Report ===");
  
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

      // Get assets data for this account
      const assetsData = getAssetsDataForAccount();
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, assetsData, accountName);
      
      Logger.log(`✓ Successfully updated assets data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC ASSETS REPORT COMPLETED ===`);
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

    return mappings;

  } catch (error) {
    Logger.log(`❌ Error reading master spreadsheet: ${error.message}`);
    return [];
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
    return null;
  }
}

function getAssetsDataForAccount() {
  try {
    // Process the main query results
    const rows = AdsApp.search(QUERY);
    const data = calculateMetrics(rows);
    
    // Sort data by conversions (descending) then by impressions (descending)
    const sortedData = sortData(data);
    
    Logger.log(`✓ Processed ${data.length} assets with impressions > 0`);
    return sortedData;
    
  } catch (error) {
    Logger.log(`❌ Error getting assets data: ${error.message}`);
    return [];
  }
}

function calculateMetrics(rows) {
  let data = [];
  let rowCount = 0;
  let assetTypeCounts = {};

  while (rows.hasNext()) {
    try {
      let row = rows.next();
      rowCount++;

      // Access dimensions using correct nested object structure
      let assetResourceName = row.campaignAsset && row.campaignAsset.asset ? row.campaignAsset.asset : 'N/A';
      let fieldType = row.campaignAsset && row.campaignAsset.fieldType ? row.campaignAsset.fieldType : 'UNKNOWN';
      
      // Extract actual asset text content based on asset type
      let assetText = 'N/A';
      if (row.asset) {
        if (fieldType === 'SITELINK' && row.asset.sitelinkAsset) {
          const desc1 = row.asset.sitelinkAsset.description1 || '';
          const desc2 = row.asset.sitelinkAsset.description2 || '';
          assetText = [desc1, desc2].filter(text => text.trim()).join(' ');
        } else if (fieldType === 'CALLOUT' && row.asset.calloutAsset) {
          assetText = row.asset.calloutAsset.calloutText || 'N/A';
        } else if (fieldType === 'STRUCTURED_SNIPPET' && row.asset.structuredSnippetAsset) {
          const header = row.asset.structuredSnippetAsset.header || '';
          const values = row.asset.structuredSnippetAsset.values || [];
          assetText = `${header}: ${values.join(', ')}`;
        } else if (row.asset.textAsset) {
          assetText = row.asset.textAsset.text || 'N/A';
        }
      }
      
      // Clean up asset text
      if (assetText === 'N/A' || !assetText.trim()) {
        assetText = assetResourceName.split('/').pop(); // Fallback to asset ID
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

      // Only include assets with impressions > 0
      if (impressions > 0) {
        // Count asset types for logging
        if (assetTypeCounts[fieldType]) {
          assetTypeCounts[fieldType]++;
        } else {
          assetTypeCounts[fieldType] = 1;
        }

        // Add asset data with performance metrics to a new row
        let newRow = [
          assetText,
          fieldType,
          impressions,
          clicks,
          cost,
          conversions,
          conversionValue
        ];

        data.push(newRow);
      }

    } catch (e) {
      // Continue with next row
    }
  }

  Logger.log(`✓ Found ${data.length} assets with impressions > 0`);
  Logger.log(`Asset breakdown: ${JSON.stringify(assetTypeCounts)}`);
  return data;
}

function sortData(data) {
  return data.sort((a, b) => {
    // Primary sort: conversions (descending)
    const conversionsA = a[5] || 0; // conversions is at index 5
    const conversionsB = b[5] || 0;
    
    if (conversionsA !== conversionsB) {
      return conversionsB - conversionsA; // Descending order
    }
    
    // Secondary sort: impressions (descending)
    const impressionsA = a[2] || 0; // impressions is at index 2
    const impressionsB = b[2] || 0;
    
    return impressionsB - impressionsA; // Descending order
  });
}

function updateAccountSpreadsheet(spreadsheetUrl, assetsData, accountName) {
  try {
    Logger.log("Updating account spreadsheet with assets data...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Create or get the tab for assets data
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
      'Asset Text',
      'Asset Type',
      'Impressions',
      'Clicks',
      'Cost',
      'Conversions',
      'Conversion Value'
    ];
    
    // Prepare all data for bulk write
    let allData = [headers];
    
    if (assetsData.length > 0) {
      allData = allData.concat(assetsData);
    } else {
      // Add a row indicating no data found - pad with empty cells to match header count
      const noDataRow = ['No assets data found for the last 30 days'];
      // Pad the row with empty strings to match the number of headers
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
    
    Logger.log(`✓ Updated ${TAB} sheet with ${assetsData.length} assets for ${accountName}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}
