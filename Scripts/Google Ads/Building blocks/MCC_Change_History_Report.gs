// --- MCC CHANGE HISTORY REPORT CONFIGURATION ---

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
const TAB = 'ChangeHistory';
const CAMPAIGNS_TAB = 'CampaignLookup';
const ADGROUPS_TAB = 'AdGroupLookup';

// --- END OF CONFIGURATION ---

const QUERY = `
  SELECT 
    change_event.change_date_time,
    change_event.change_resource_name,
    change_event.change_resource_type,
    change_event.client_type,
    change_event.user_email,
    change_event.resource_change_operation,
    change_event.changed_fields,
    change_event.old_resource,
    change_event.new_resource
  FROM change_event 
  WHERE change_event.change_date_time DURING LAST_7_DAYS
  ORDER BY change_event.change_date_time DESC
  LIMIT 10000
`;

function main() {
  Logger.log("=== Starting MCC Change History Report ===");
  
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

      // Get change history data for this account
      const result = getChangeHistoryDataForAccount();
      const changeHistoryData = result.changeHistory;
      const campaignLookup = result.campaignLookup;
      const adGroupLookup = result.adGroupLookup;
      
      // Update the account's spreadsheet
      updateAccountSpreadsheet(spreadsheetUrl, changeHistoryData, accountName, campaignLookup, adGroupLookup);
      
      Logger.log(`✓ Successfully updated change history data for ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      errorCount++;
    }
  }

  // Summary
  Logger.log(`\n=== MCC CHANGE HISTORY REPORT COMPLETED ===`);
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

function getChangeHistoryDataForAccount() {
  Logger.log("Getting change history data for current account...");
  
  try {
    // Create lookup tables for campaigns and ad groups
    const campaignLookup = createCampaignLookup();
    const adGroupLookup = createAdGroupLookup();
    
    // Get raw change history data
    const report = AdsApp.report(QUERY);
    const rows = report.rows();
    const data = getSimpleChangeData(rows);
    
    Logger.log(`✓ Processed ${data.length} change history records for current account`);
    return {
      changeHistory: data,
      campaignLookup: campaignLookup,
      adGroupLookup: adGroupLookup
    };
    
  } catch (error) {
    Logger.log(`❌ Error getting change history data: ${error.message}`);
    throw error;
  }
}

function getSimpleChangeData(rows) {
  let data = [];
  let count = 0;
  
  while (rows.hasNext()) {
    try {
      const row = rows.next();
      count++;
      
      // Simple data extraction - just get the raw fields
      const changeDateTime = row['change_event.change_date_time'] || '';
      const changeResourceName = row['change_event.change_resource_name'] || '';
      const changeResourceType = row['change_event.change_resource_type'] || '';
      const clientType = row['change_event.client_type'] || '';
      const userEmail = row['change_event.user_email'] || '';
      const resourceChangeOperation = row['change_event.resource_change_operation'] || '';
      const changedFields = row['change_event.changed_fields'] || '';
      
      const newRow = [
        changeDateTime,
        changeResourceName,
        changeResourceType,
        clientType,
        userEmail,
        resourceChangeOperation,
        changedFields
      ];
      
      data.push(newRow);
      
    } catch (error) {
      Logger.log(`Error processing row ${count}: ${error.message}`);
    }
  }
  
  return data;
}

// Fast version without API calls for better performance
function extractResourceInfoFast(resourceName, resourceType) {
  let campaignName = 'Unknown';
  let adGroupName = 'Unknown';
  let resourceNameClean = 'Unknown';
  
  try {
    // Clean up resource name for display
    resourceNameClean = resourceName.replace(/^customers\/\d+\//, '').replace(/\/\d+$/, '');
    
    // Extract IDs without making API calls
    if (resourceType === 'AD_GROUP_CRITERION') {
      const parts = resourceName.split('/');
      if (parts.length >= 4) {
        const adGroupCriterionId = parts[3];
        const adGroupId = adGroupCriterionId.split('~')[0];
        adGroupName = `Ad Group ${adGroupId}`;
        campaignName = `Campaign (ID: ${adGroupId})`;
      }
    } else if (resourceType === 'AD_GROUP') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const adGroupId = parts[2].replace('adGroups/', '');
        adGroupName = `Ad Group ${adGroupId}`;
        campaignName = `Campaign (ID: ${adGroupId})`;
      }
    } else if (resourceType === 'CAMPAIGN') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const campaignId = parts[2].replace('campaigns/', '');
        campaignName = `Campaign ${campaignId}`;
        adGroupName = 'N/A';
      }
    } else if (resourceType === 'AD_GROUP_AD') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const adGroupId = parts[2].replace('adGroups/', '');
        adGroupName = `Ad Group ${adGroupId}`;
        campaignName = `Campaign (ID: ${adGroupId})`;
      }
    } else if (resourceType === 'ASSET_GROUP') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const assetGroupId = parts[2].replace('assetGroups/', '');
        adGroupName = 'Asset Group';
        campaignName = `Performance Max (ID: ${assetGroupId})`;
      }
    } else if (resourceType === 'ASSET_GROUP_ASSET') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const assetGroupId = parts[2].replace('assetGroups/', '');
        adGroupName = 'Asset Group';
        campaignName = `Performance Max (ID: ${assetGroupId})`;
      }
    }
  } catch (e) {
    // Keep defaults if parsing fails
  }
  
  return {
    campaignName: campaignName,
    adGroupName: adGroupName,
    resourceName: resourceNameClean
  };
}

// Helper function to extract meaningful resource information (with API calls - slower)
function extractResourceInfo(resourceName, resourceType) {
  let campaignName = 'Unknown';
  let adGroupName = 'Unknown';
  let resourceNameClean = 'Unknown';
  
  try {
    // Clean up resource name for display
    resourceNameClean = resourceName.replace(/^customers\/\d+\//, '').replace(/\/\d+$/, '');
    
    // Extract campaign and ad group info based on resource type
    if (resourceType === 'AD_GROUP_CRITERION') {
      const parts = resourceName.split('/');
      if (parts.length >= 4) {
        const adGroupCriterionId = parts[3];
        // Extract ad group ID from criterion ID (format: adGroupId~criterionId)
        const adGroupId = adGroupCriterionId.split('~')[0];
        adGroupName = getAdGroupNameById(adGroupId);
        campaignName = getCampaignNameFromAdGroup(adGroupId);
      }
    } else if (resourceType === 'AD_GROUP') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const adGroupId = parts[2].replace('adGroups/', '');
        adGroupName = getAdGroupNameById(adGroupId);
        campaignName = getCampaignNameFromAdGroup(adGroupId);
      }
    } else if (resourceType === 'CAMPAIGN') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const campaignId = parts[2].replace('campaigns/', '');
        campaignName = getCampaignNameById(campaignId);
        adGroupName = 'N/A';
      }
    } else if (resourceType === 'AD_GROUP_AD') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const adGroupId = parts[2].replace('adGroups/', '');
        adGroupName = getAdGroupNameById(adGroupId);
        campaignName = getCampaignNameFromAdGroup(adGroupId);
      }
    } else if (resourceType === 'ASSET_GROUP') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const assetGroupId = parts[2].replace('assetGroups/', '');
        adGroupName = 'Asset Group';
        campaignName = getCampaignNameFromAssetGroup(assetGroupId);
      }
    } else if (resourceType === 'ASSET_GROUP_ASSET') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const assetGroupId = parts[2].replace('assetGroups/', '');
        adGroupName = 'Asset Group';
        campaignName = getCampaignNameFromAssetGroup(assetGroupId);
      }
    }
  } catch (e) {
    // Keep defaults if parsing fails
  }
  
  return {
    campaignName: campaignName,
    adGroupName: adGroupName,
    resourceName: resourceNameClean
  };
}

// Helper function to get ad group name by ID
function getAdGroupNameById(adGroupId) {
  try {
    const adGroupIterator = AdsApp.adGroups()
      .withIds([adGroupId])
      .get();
    
    if (adGroupIterator.hasNext()) {
      const adGroup = adGroupIterator.next();
      return adGroup.getName();
    }
  } catch (e) {
    // Return ID if lookup fails
  }
  return `Ad Group ${adGroupId}`;
}

// Helper function to get campaign name by ID
function getCampaignNameById(campaignId) {
  try {
    const campaignIterator = AdsApp.campaigns()
      .withIds([campaignId])
      .get();
    
    if (campaignIterator.hasNext()) {
      const campaign = campaignIterator.next();
      return campaign.getName();
    }
  } catch (e) {
    // Return ID if lookup fails
  }
  return `Campaign ${campaignId}`;
}

// Helper function to get campaign name from ad group ID
function getCampaignNameFromAdGroup(adGroupId) {
  try {
    const adGroupIterator = AdsApp.adGroups()
      .withIds([adGroupId])
      .get();
    
    if (adGroupIterator.hasNext()) {
      const adGroup = adGroupIterator.next();
      const campaign = adGroup.getCampaign();
      return campaign.getName();
    }
  } catch (e) {
    // Return generic name if lookup fails
  }
  return 'Unknown Campaign';
}

// Helper function to get campaign name from asset group ID
function getCampaignNameFromAssetGroup(assetGroupId) {
  try {
    const assetGroupIterator = AdsApp.assetGroups()
      .withIds([assetGroupId])
      .get();
    
    if (assetGroupIterator.hasNext()) {
      const assetGroup = assetGroupIterator.next();
      const campaign = assetGroup.getCampaign();
      return campaign.getName();
    }
  } catch (e) {
    // Return generic name if lookup fails
  }
  return 'Unknown Campaign';
}

// Helper function to generate a human-readable change summary
function generateChangeSummary(operation, changedFields, oldResource, newResource) {
  let summary = operation;
  
  try {
    if (operation === 'CREATE') {
      summary = 'Created new ' + (changedFields || 'resource');
    } else if (operation === 'UPDATE') {
      const fields = changedFields ? changedFields.split(',').length : 0;
      summary = `Updated ${fields} field(s)`;
    } else if (operation === 'REMOVE') {
      summary = 'Removed ' + (changedFields || 'resource');
    }
    
    // Add specific context for common changes
    if (changedFields && changedFields.includes('status')) {
      summary += ' (Status change)';
    } else if (changedFields && changedFields.includes('bid')) {
      summary += ' (Bid change)';
    } else if (changedFields && changedFields.includes('keyword')) {
      summary += ' (Keyword change)';
    }
  } catch (e) {
    // Keep original operation if parsing fails
  }
  
  return summary;
}


// Helper function to format client type for better readability
function formatClientType(clientType) {
  const typeMap = {
    'GOOGLE_ADS_WEB_CLIENT': 'Web Interface',
    'GOOGLE_ADS_API_CLIENT': 'API',
    'GOOGLE_ADS_SCRIPT': 'Script',
    'GOOGLE_ADS_MOBILE_APP': 'Mobile App',
    'GOOGLE_ADS_BULK_UPLOAD': 'Bulk Upload',
    'GOOGLE_ADS_EDITOR': 'Editor'
  };
  return typeMap[clientType] || clientType;
}

// Helper function to format user email for privacy
function formatUserEmail(email) {
  if (!email || email === 'Unknown') return 'Unknown';
  
  try {
    const parts = email.split('@');
    if (parts.length === 2) {
      return parts[0] + '@***';
    }
  } catch (e) {
    // Keep original if parsing fails
  }
  
  return email;
}

// Create campaign lookup table
function createCampaignLookup() {
  const lookup = {};
  
  try {
    const campaigns = AdsApp.campaigns().get();
    while (campaigns.hasNext()) {
      const campaign = campaigns.next();
      const id = campaign.getId();
      const name = campaign.getName();
      lookup[id] = name;
    }
    Logger.log(`✓ Created campaign lookup with ${Object.keys(lookup).length} campaigns`);
  } catch (error) {
    Logger.log(`⚠️ Error creating campaign lookup: ${error.message}`);
  }
  
  return lookup;
}

// Create ad group lookup table
function createAdGroupLookup() {
  const lookup = {};
  
  try {
    const adGroups = AdsApp.adGroups().get();
    while (adGroups.hasNext()) {
      const adGroup = adGroups.next();
      const id = adGroup.getId();
      const name = adGroup.getName();
      lookup[id] = name;
    }
    Logger.log(`✓ Created ad group lookup with ${Object.keys(lookup).length} ad groups`);
  } catch (error) {
    Logger.log(`⚠️ Error creating ad group lookup: ${error.message}`);
  }
  
  return lookup;
}

// Extract resource info using lookup tables (simplified to show IDs)
function extractResourceInfoWithLookup(resourceName, resourceType, campaignLookup, adGroupLookup) {
  let campaignName = 'Unknown';
  let adGroupName = 'Unknown';
  let resourceNameClean = 'Unknown';
  
  try {
    // Clean up resource name for display
    resourceNameClean = resourceName.replace(/^customers\/\d+\//, '').replace(/\/\d+$/, '');
    
    // Extract IDs (simplified - just show IDs, not names)
    if (resourceType === 'AD_GROUP_CRITERION') {
      const parts = resourceName.split('/');
      if (parts.length >= 4) {
        const adGroupCriterionId = parts[3];
        const adGroupId = adGroupCriterionId.split('~')[0];
        adGroupName = adGroupId;
        campaignName = getCampaignIdFromAdGroupId(adGroupId);
      }
    } else if (resourceType === 'AD_GROUP') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const adGroupId = parts[2].replace('adGroups/', '');
        adGroupName = adGroupId;
        campaignName = getCampaignIdFromAdGroupId(adGroupId);
      }
    } else if (resourceType === 'CAMPAIGN') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const campaignId = parts[2].replace('campaigns/', '');
        campaignName = campaignId;
        adGroupName = 'N/A';
      }
    } else if (resourceType === 'AD_GROUP_AD') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const adGroupId = parts[2].replace('adGroups/', '');
        adGroupName = adGroupId;
        campaignName = getCampaignIdFromAdGroupId(adGroupId);
      }
    } else if (resourceType === 'ASSET_GROUP') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const assetGroupId = parts[2].replace('assetGroups/', '');
        adGroupName = 'Asset Group';
        campaignName = assetGroupId;
      }
    } else if (resourceType === 'ASSET_GROUP_ASSET') {
      const parts = resourceName.split('/');
      if (parts.length >= 3) {
        const assetGroupId = parts[2].replace('assetGroups/', '');
        adGroupName = 'Asset Group';
        campaignName = assetGroupId;
      }
    }
  } catch (e) {
    // Keep defaults if parsing fails
  }
  
  return {
    campaignName: campaignName,
    adGroupName: adGroupName,
    resourceName: resourceNameClean
  };
}

// Get campaign ID from ad group ID (simplified)
function getCampaignIdFromAdGroupId(adGroupId) {
  try {
    const adGroup = AdsApp.adGroups().withIds([adGroupId]).get();
    if (adGroup.hasNext()) {
      const adGroupObj = adGroup.next();
      const campaign = adGroupObj.getCampaign();
      return campaign.getId();
    }
  } catch (e) {
    // Return generic ID if lookup fails
  }
  return adGroupId; // Fallback to ad group ID if campaign lookup fails
}

// Comprehensive spreadsheet formatting function
function formatSpreadsheet(sheet, totalRows, totalCols) {
  try {
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, totalCols);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285F4'); // Google Blue
    headerRange.setFontColor('white');
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Set column widths for better readability
    const columnWidths = [12, 10, 20, 20, 25, 15, 12, 20, 20, 15, 15]; // Date, Time, Campaign, etc.
    for (let i = 0; i < Math.min(columnWidths.length, totalCols); i++) {
      sheet.setColumnWidth(i + 1, columnWidths[i] * 8); // Convert to pixels
    }
    
    // Format data rows with alternating colors
    if (totalRows > 1) {
      const dataRange = sheet.getRange(2, 1, totalRows - 1, totalCols);
      
      // Apply alternating row colors
      for (let row = 2; row <= totalRows; row++) {
        const rowRange = sheet.getRange(row, 1, 1, totalCols);
        if (row % 2 === 0) {
          rowRange.setBackground('#F8F9FA'); // Light gray for even rows
        } else {
          rowRange.setBackground('white'); // White for odd rows
        }
        rowRange.setVerticalAlignment('middle');
      }
      
      // Format specific columns
      // Date column (A)
      if (totalCols >= 1) {
        const dateRange = sheet.getRange(2, 1, totalRows - 1, 1);
        dateRange.setNumberFormat('mm/dd/yyyy');
      }
      
      // Time column (B)
      if (totalCols >= 2) {
        const timeRange = sheet.getRange(2, 2, totalRows - 1, 1);
        timeRange.setNumberFormat('hh:mm:ss AM/PM');
      }
      
      
      // Operation column (G) - add conditional formatting
      if (totalCols >= 7) {
        const operationRange = sheet.getRange(2, 7, totalRows - 1, 1);
        
        // CREATE - green text
        const createRule = SpreadsheetApp.newConditionalFormatRule()
          .setRanges([operationRange])
          .whenTextEqualTo('CREATE')
          .setFontColor('#2E7D32') // Dark green
          .build();
        
        // UPDATE - blue text
        const updateRule = SpreadsheetApp.newConditionalFormatRule()
          .setRanges([operationRange])
          .whenTextEqualTo('UPDATE')
          .setFontColor('#1976D2') // Dark blue
          .build();
        
        // REMOVE - red text
        const removeRule = SpreadsheetApp.newConditionalFormatRule()
          .setRanges([operationRange])
          .whenTextEqualTo('REMOVE')
          .setFontColor('#D32F2F') // Dark red
          .build();
        
        const rules = sheet.getConditionalFormatRules();
        rules.push(createRule, updateRule, removeRule);
        sheet.setConditionalFormatRules(rules);
      }
    }
    
    // Add borders to all cells
    const allRange = sheet.getRange(1, 1, totalRows, totalCols);
    allRange.setBorder(true, true, true, true, true, true);
    
    // Set text wrapping for long content
    const dataRange = sheet.getRange(1, 1, totalRows, totalCols);
    dataRange.setWrap(true);
    
    // Add a summary row at the top (after headers)
    if (totalRows > 1) {
      sheet.insertRowAfter(1);
      const summaryRow = sheet.getRange(2, 1, 1, totalCols);
      summaryRow.merge();
      summaryRow.setValue(`📊 Change History Summary: ${totalRows - 1} changes in the last 7 days`);
      summaryRow.setBackground('#E3F2FD'); // Light blue
      summaryRow.setFontWeight('bold');
      summaryRow.setHorizontalAlignment('center');
      summaryRow.setVerticalAlignment('middle');
      
      // Adjust frozen rows to include summary
      sheet.setFrozenRows(2);
    }
    
    Logger.log('✓ Applied comprehensive spreadsheet formatting');
    
  } catch (error) {
    Logger.log(`⚠️ Error applying formatting: ${error.message}`);
  }
}

function sortData(data) {
  return data.sort((a, b) => {
    // Sort by change date (descending - most recent first)
    const dateA = new Date(a[0]);
    const dateB = new Date(b[0]);
    return dateB - dateA;
  });
}

function updateAccountSpreadsheet(spreadsheetUrl, changeHistoryData, accountName, campaignLookup, adGroupLookup) {
  try {
    Logger.log("Updating account spreadsheet with change history data...");
    
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    Logger.log(`✓ Opened spreadsheet: ${spreadsheet.getName()}`);
    
    // Update main change history sheet
    updateChangeHistorySheet(spreadsheet, changeHistoryData, accountName);
    
    // Create campaign lookup sheet
    updateCampaignLookupSheet(spreadsheet, campaignLookup);
    
    // Create ad group lookup sheet
    updateAdGroupLookupSheet(spreadsheet, adGroupLookup);
    
    Logger.log(`✓ Successfully updated all sheets for ${accountName}`);
    
  } catch (error) {
    Logger.log(`❌ Error updating spreadsheet: ${error.message}`);
    throw error;
  }
}

function updateChangeHistorySheet(spreadsheet, changeHistoryData, accountName) {
  let sheet = spreadsheet.getSheetByName(TAB);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(TAB);
  } else {
    sheet.clear();
  }
  
  // Create headers
  const headers = [
    'Change Date Time',
    'Resource Name',
    'Resource Type',
    'Client Type',
    'User Email',
    'Operation',
    'Changed Fields'
  ];
  
  // Prepare all data for bulk write
  let allData = [headers];
  
  if (changeHistoryData.length > 0) {
    allData = allData.concat(changeHistoryData);
  } else {
    allData.push(['No change history data found for the last 7 days']);
  }
  
  // Write all data at once using setValues for efficiency
  const range = sheet.getRange(1, 1, allData.length, headers.length);
  range.setValues(allData);
  
  // Apply simple formatting
  formatSimpleSpreadsheet(sheet, allData.length, headers.length);
  
  Logger.log(`✓ Successfully updated ${TAB} sheet with ${changeHistoryData.length} change history records`);
}

function updateCampaignLookupSheet(spreadsheet, campaignLookup) {
  let sheet = spreadsheet.getSheetByName(CAMPAIGNS_TAB);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CAMPAIGNS_TAB);
  } else {
    sheet.clear();
  }
  
  // Create headers
  const headers = ['Campaign ID', 'Campaign Name'];
  
  // Prepare data
  let allData = [headers];
  for (const [id, name] of Object.entries(campaignLookup)) {
    allData.push([id, name]);
  }
  
  // Write data
  const range = sheet.getRange(1, 1, allData.length, headers.length);
  range.setValues(allData);
  
  // Format the sheet
  formatLookupSheet(sheet, allData.length, headers.length);
  
  Logger.log(`✓ Created campaign lookup with ${Object.keys(campaignLookup).length} campaigns`);
}

function updateAdGroupLookupSheet(spreadsheet, adGroupLookup) {
  let sheet = spreadsheet.getSheetByName(ADGROUPS_TAB);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(ADGROUPS_TAB);
  } else {
    sheet.clear();
  }
  
  // Create headers
  const headers = ['Ad Group ID', 'Ad Group Name'];
  
  // Prepare data
  let allData = [headers];
  for (const [id, name] of Object.entries(adGroupLookup)) {
    allData.push([id, name]);
  }
  
  // Write data
  const range = sheet.getRange(1, 1, allData.length, headers.length);
  range.setValues(allData);
  
  // Format the sheet
  formatLookupSheet(sheet, allData.length, headers.length);
  
  Logger.log(`✓ Created ad group lookup with ${Object.keys(adGroupLookup).length} ad groups`);
}

function formatSimpleSpreadsheet(sheet, totalRows, totalCols) {
  try {
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, totalCols);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285F4');
    headerRange.setFontColor('white');
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, totalCols);
    
    // Add borders
    const allRange = sheet.getRange(1, 1, totalRows, totalCols);
    allRange.setBorder(true, true, true, true, true, true);
    
  } catch (error) {
    Logger.log(`⚠️ Error applying formatting: ${error.message}`);
  }
}

function formatLookupSheet(sheet, totalRows, totalCols) {
  try {
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, totalCols);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285F4');
    headerRange.setFontColor('white');
    headerRange.setHorizontalAlignment('center');
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Set column widths
    sheet.setColumnWidth(1, 120); // ID column
    sheet.setColumnWidth(2, 300); // Name column
    
    // Add borders
    const allRange = sheet.getRange(1, 1, totalRows, totalCols);
    allRange.setBorder(true, true, true, true, true, true);
    
    // Auto-resize name column
    sheet.autoResizeColumns(2, 1);
    
    Logger.log('✓ Applied lookup sheet formatting');
    
  } catch (error) {
    Logger.log(`⚠️ Error applying lookup sheet formatting: ${error.message}`);
  }
}
