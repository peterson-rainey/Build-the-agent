/**
 * MCC Zero Impressions Monitor
 * 
 * This script monitors all enabled campaigns, ad groups, ads, and asset groups
 * across multiple MCC accounts to identify any content that got zero impressions yesterday.
 * 
 * Features:
 * - MCC-level account iteration
 * - Traditional ads zero impressions monitoring
 * - Performance Max asset group zero impressions monitoring
 * - Email alerts for zero impression content
 * - Comprehensive logging and error handling
 * - Script health monitoring
 * 
 * @author AI Assistant
 * @version 1.0
 * @date 2025-09-03
 */

// Configuration
const MASTER_SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1bS1xmTFB0LWGvYeZBPcykGMY0SwZ4W-Ny6aKVloy0SM/edit';
const MASTER_SHEET_NAME = 'AccountMappings';

/**
 * Main function - Entry point for the script
 */
function main() {
  Logger.log('=== Starting MCC Zero Impressions Monitor ===');
  Logger.log('📋 Filtering: Only active, approved ads from enabled, non-experiment campaigns (excludes experiments, ineligible, ended campaigns, and draft campaigns)');
  
  // Check if today is Monday through Friday (business days only)
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday
  
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    Logger.log('📅 Today is a weekend (Saturday or Sunday). This script only runs Monday through Friday.');
    Logger.log('📅 Script will exit without processing accounts.');
    logScriptHealth(105); // zero impressions monitor - weekend exit
    return;
  }
  
  Logger.log(`📅 Today is a business day (${getDayName(dayOfWeek)}). Proceeding with zero impressions monitoring...`);
  
  try {
    // Read account mappings from master spreadsheet
    const accountMappings = getAccountMappings();
    
    if (!accountMappings || accountMappings.length === 0) {
      Logger.log('❌ No account mappings found. Exiting.');
      logScriptHealth(105); // Error 105: no account mappings
      return;
    }

    Logger.log(`📋 Found ${accountMappings.length} account mappings.`);
    
    // Process each account for zero impressions content
    Logger.log('🔍 Processing each account for zero impressions content...');
    const accountsWithZeroImpressions = [];
    let totalZeroImpressionsAds = 0;
    let totalZeroImpressionsAssetGroups = 0;
    let totalCampaignsChecked = 0;
    let totalAdGroupsChecked = 0;
    let totalAdsChecked = 0;
    let totalAssetGroupsChecked = 0;
    
    accountMappings.forEach((mapping, index) => {
      Logger.log(`\n--- Processing Account ${index + 1}/${accountMappings.length} ---`);
      Logger.log(`Account: ${mapping.accountName} (${mapping.accountId})`);
      
      try {
        // Find and switch to the account
        const account = getAccountById(mapping.accountId);
        if (!account) {
          Logger.log(`❌ Account not found: ${mapping.accountName}`);
          logScriptHealth(105); // ad disapproval check
          return;
        }

        Logger.log(`✓ Account found: ${account.getName()}`);
        
        // Switch to account context
        AdsManagerApp.select(account);
        Logger.log(`✓ Switched to account: ${account.getName()}`);

        // Get zero impressions issues for this account
        Logger.log(`🔍 Getting zero impressions issues for ${mapping.accountName}...`);
        const zeroImpressionsIssues = getZeroImpressionsIssues(mapping.accountName);
        
        if (zeroImpressionsIssues) {
          // Update counters
          totalCampaignsChecked += zeroImpressionsIssues.campaignCount || 0;
          totalAdGroupsChecked += zeroImpressionsIssues.adGroupCount || 0;
          totalAdsChecked += zeroImpressionsIssues.adCount || 0;
          totalAssetGroupsChecked += zeroImpressionsIssues.assetGroupCount || 0;
          totalZeroImpressionsAds += zeroImpressionsIssues.zeroImpressionsAds || 0;
          totalZeroImpressionsAssetGroups += zeroImpressionsIssues.zeroImpressionsAssetGroups || 0;
          
          // Store issues for this account
          if (zeroImpressionsIssues.zeroImpressionsAds > 0 || zeroImpressionsIssues.zeroImpressionsAssetGroups > 0) {
            accountsWithZeroImpressions.push({
              accountName: mapping.accountName,
              accountId: mapping.accountId,
              emailRecipient: mapping.emailRecipient,
              campaignCount: zeroImpressionsIssues.campaignCount || 0,
              adGroupCount: zeroImpressionsIssues.adGroupCount || 0,
              adCount: zeroImpressionsIssues.adCount || 0,
              assetGroupCount: zeroImpressionsIssues.assetGroupCount || 0,
              zeroImpressionsAds: zeroImpressionsIssues.zeroImpressionsAds || 0,
              zeroImpressionsAssetGroups: zeroImpressionsIssues.zeroImpressionsAssetGroups || 0,
              adIssues: zeroImpressionsIssues.adIssues || [],
              assetGroupIssues: zeroImpressionsIssues.assetGroupIssues || []
            });
            
            Logger.log(`⚠️ Found ${zeroImpressionsIssues.zeroImpressionsAds} ads with zero impressions and ${zeroImpressionsIssues.zeroImpressionsAssetGroups} asset groups with zero impressions`);
          } else {
            Logger.log(`✓ No zero impressions content found`);
          }
        }
        
        Logger.log(`✓ Successfully processed ${mapping.accountName}`);

      } catch (error) {
        Logger.log(`❌ Error processing ${mapping.accountName}: ${error.message}`);
      }
    });
    
    // Send individual email alerts to each account's contact person
    if (accountsWithZeroImpressions.length > 0) {
      Logger.log(`📧 Sending individual email alerts to ${accountsWithZeroImpressions.length} accounts with zero impressions content`);
      accountsWithZeroImpressions.forEach(accountData => {
        sendIndividualZeroImpressionsAlertEmail(accountData);
      });
    } else {
      Logger.log(`✓ No email sent - no zero impressions content found`);
    }
    
    // Final summary
    Logger.log('\n=== MCC ZERO IMPRESSIONS MONITOR COMPLETED ===');
    Logger.log('📊 Summary:');
    Logger.log(`  - Total accounts processed: ${accountMappings.length}`);
    Logger.log(`  - Total campaigns checked: ${totalCampaignsChecked}`);
    Logger.log(`  - Total ad groups checked: ${totalAdGroupsChecked}`);
    Logger.log(`  - Total ads checked: ${totalAdsChecked}`);
    Logger.log(`  - Total asset groups checked: ${totalAssetGroupsChecked}`);
    Logger.log(`  - Total ads with zero impressions: ${totalZeroImpressionsAds}`);
    Logger.log(`  - Total asset groups with zero impressions: ${totalZeroImpressionsAssetGroups}`);
    Logger.log(`  - Accounts with zero impressions content: ${accountsWithZeroImpressions.length}`);
    
  } catch (error) {
    Logger.log(`❌ Critical error in main function: ${error.message}`);
    Logger.log(`Error details: ${error.toString()}`);
  }
  
  logScriptHealth(105); // zero impressions monitor completed
}

/**
 * Read account mappings from the master spreadsheet
 */
function getAccountMappings() {
  try {
    Logger.log("📋 Reading account mappings from master spreadsheet...");
    
    const masterSpreadsheet = SpreadsheetApp.openByUrl(MASTER_SPREADSHEET_URL);
    Logger.log(`✓ Opened master spreadsheet: ${masterSpreadsheet.getName()}`);
    
    const masterSheet = masterSpreadsheet.getSheetByName(MASTER_SHEET_NAME);
    
    if (!masterSheet) {
      throw new Error(`Tab "${MASTER_SHEET_NAME}" not found in master spreadsheet`);
    }

    Logger.log(`✓ Found sheet: ${MASTER_SHEET_NAME}`);
    
    const data = masterSheet.getDataRange().getValues();
    Logger.log(`✓ Read ${data.length} rows from master spreadsheet`);
    
    const mappings = [];

    // Skip header row, process data rows
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const accountName = row[0]?.toString().trim() || "Unknown";
      const accountId = row[1]?.toString().trim();
      const spreadsheetUrl = row[2]?.toString().trim();
      const emailRecipient = row[3]?.toString().trim() || ""; // Column D for email

      if (accountId) {
        Logger.log(`📋 Row ${i + 1}: ${accountName} (${accountId}) - Email: ${emailRecipient || 'Not specified'}`);
        
        mappings.push({
          accountId: accountId,
          spreadsheetUrl: spreadsheetUrl,
          accountName: accountName,
          emailRecipient: emailRecipient
        });
      } else {
        Logger.log(`⚠️ Row ${i + 1}: Skipping - no account ID found`);
      }
    }

    Logger.log(`✓ Found ${mappings.length} valid account mappings`);
    return mappings;

  } catch (error) {
    Logger.log(`❌ Error reading master spreadsheet: ${error.message}`);
    Logger.log(`Error details: ${error.stack || 'No stack trace available'}`);
    throw error;
  }
}

/**
 * Find an account by ID within the MCC
 */
function getAccountById(accountId) {
  try {
    Logger.log(`🔍 Looking for account ID: ${accountId}`);
    const accountIterator = AdsManagerApp.accounts().withIds([accountId]).get();
    if (accountIterator.hasNext()) {
      const account = accountIterator.next();
      Logger.log(`✓ Found account: ${account.getName()}`);
      return account;
    }
    Logger.log(`❌ Account ${accountId} not found in MCC`);
    logScriptHealth(105); // ad disapproval check
    return null;
  } catch (error) {
    Logger.log(`❌ Error finding account ${accountId}: ${error.message}`);
    logScriptHealth(105); // ad disapproval check
    return null;
  }
}

/**
 * Get zero impressions issues for a specific account
 */
function getZeroImpressionsIssues(accountName) {
  const startTime = new Date();
  Logger.log(`Function started at: ${startTime.toISOString()}`);
  Logger.log(`Account context: ${accountName}`);
  
  try {
    // Get yesterday's date
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const yesterdayStr = Utilities.formatDate(yesterday, AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
    
    Logger.log(`📅 Checking for zero impressions on: ${yesterdayStr}`);
    
    // Get traditional ads zero impressions issues
    Logger.log('🔍 Using GAQL to get ads with zero impressions...');
    const adZeroImpressionsIssues = getAdZeroImpressionsIssues(accountName, yesterdayStr);
    
    // Get asset group zero impressions information
    Logger.log('🔍 Getting asset groups with zero impressions for Performance Max campaigns...');
    const assetGroupZeroImpressionsIssues = getAssetGroupZeroImpressionsIssues(accountName, yesterdayStr);
    
    // Combine results
    const totalZeroImpressionsAds = (adZeroImpressionsIssues?.zeroImpressionsAds || 0) + (assetGroupZeroImpressionsIssues?.zeroImpressionsAds || 0);
    const totalZeroImpressionsAssetGroups = assetGroupZeroImpressionsIssues?.zeroImpressionsAssetGroups || 0;
    
    const result = {
      campaignCount: adZeroImpressionsIssues?.campaignCount || 0,
      adGroupCount: adZeroImpressionsIssues?.adGroupCount || 0,
      adCount: adZeroImpressionsIssues?.adCount || 0,
      assetGroupCount: assetGroupZeroImpressionsIssues?.assetGroupCount || 0,
      zeroImpressionsAds: totalZeroImpressionsAds,
      zeroImpressionsAssetGroups: totalZeroImpressionsAssetGroups,
      adIssues: adZeroImpressionsIssues?.adIssues || [],
      assetGroupIssues: assetGroupZeroImpressionsIssues?.assetGroupIssues || []
    };
    
    // Log results
    Logger.log(`=== RESULTS FOR ${accountName.toUpperCase()} ===`);
    Logger.log(`Campaigns Checked: ${result.campaignCount}`);
    Logger.log(`Ad Groups Checked: ${result.adGroupCount}`);
    Logger.log(`Ads Checked: ${result.adCount}`);
    Logger.log(`Asset Groups Checked: ${result.assetGroupCount}`);
    Logger.log(`Ads with Zero Impressions: ${result.zeroImpressionsAds}`);
    Logger.log(`Asset Groups with Zero Impressions: ${result.zeroImpressionsAssetGroups}`);
    
    // Log execution summary
    const endTime = new Date();
    Logger.log('=== EXECUTION SUMMARY ===');
    Logger.log(`Function completed at: ${endTime.toISOString()}`);
    Logger.log(`Total GAQL rows processed: ${adZeroImpressionsIssues?.totalRows || 0}`);
    Logger.log(`Total asset group rows processed: ${assetGroupZeroImpressionsIssues?.totalRows || 0}`);
    Logger.log(`Data extraction successful: Yes`);
    
    if (totalZeroImpressionsAds > 0 || totalZeroImpressionsAssetGroups > 0) {
      Logger.log(`⚠️ Found ${totalZeroImpressionsAds} ads with zero impressions and ${totalZeroImpressionsAssetGroups} asset groups with zero impressions`);
    } else {
      Logger.log(`✓ No zero impressions content found`);
    }
    
    return result;
    
  } catch (error) {
    Logger.log(`❌ Error getting zero impressions issues for ${accountName}: ${error.message}`);
    Logger.log(`Error details: ${error.toString()}`);
    return null;
  }
}

/**
 * Get zero impressions issues for traditional ads using GAQL
 */
function getAdZeroImpressionsIssues(accountName, dateStr) {
  try {
    Logger.log('Executing GAQL query for ads with zero impressions...');
    
    const query = `
      SELECT 
        campaign.id,
        campaign.name,
        campaign.status,
        campaign.experiment_type,
        campaign.end_date,
        ad_group.id,
        ad_group.name,
        ad_group_ad.ad.id,
        ad_group_ad.ad.type,
        ad_group_ad.status,
        ad_group_ad.policy_summary.approval_status,
        metrics.impressions
      FROM ad_group_ad
      WHERE 
        campaign.status = 'ENABLED' AND
        ad_group.status = 'ENABLED' AND
        ad_group_ad.status = 'ENABLED' AND
        ad_group_ad.policy_summary.approval_status IN ('APPROVED', 'APPROVED_LIMITED') AND
        segments.date = '${dateStr}'
      ORDER BY campaign.name, ad_group.name
    `;
    
    Logger.log(`Query: ${query}`);
    
    const rows = AdsApp.search(query, { 'apiVersion': 'v20' });
    let totalRows = 0;
    let campaignCount = 0;
    let adGroupCount = 0;
    let adCount = 0;
    let zeroImpressionsAds = 0;
    const adIssues = [];
    const processedCampaigns = new Set();
    const processedAdGroups = new Set();
    
    // Log first row structure for debugging
    if (rows.hasNext()) {
      const firstRow = rows.next();
      const flattenedRow = flattenObject(firstRow);
      Logger.log(`🔍 First row available fields: ${Object.keys(flattenedRow).join(', ')}`);
      Logger.log(`🔍 Sample field values:`);
      Logger.log(`   campaign.id: ${flattenedRow['campaign.id']}`);
      Logger.log(`   campaign.name: ${flattenedRow['campaign.name']}`);
      Logger.log(`   campaign.status: ${flattenedRow['campaign.status']}`);
      Logger.log(`   campaign.experimentType: ${flattenedRow['campaign.experimentType']}`);
      Logger.log(`   campaign.endDate: ${flattenedRow['campaign.endDate']}`);
      Logger.log(`   adGroup.id: ${flattenedRow['adGroup.id']}`);
      Logger.log(`   adGroup.name: ${flattenedRow['adGroup.name']}`);
      Logger.log(`   adGroupAd.ad.id: ${flattenedRow['adGroupAd.ad.id']}`);
      Logger.log(`   adGroupAd.ad.type: ${flattenedRow['adGroupAd.ad.type']}`);
      Logger.log(`   adGroupAd.policySummary.approvalStatus: ${flattenedRow['adGroupAd.policySummary.approvalStatus']}`);
      Logger.log(`   metrics.impressions: ${flattenedRow['metrics.impressions']}`);
    }
    
    Logger.log(`Executing main GAQL query for processing...`);
    
    // Execute the query again for processing (since we can't reset the iterator)
    const processingRows = AdsApp.search(query, { 'apiVersion': 'v20' });
    
    while (processingRows.hasNext()) {
      totalRows++;
      const row = processingRows.next();
      const flattenedRow = flattenObject(row);
      
      Logger.log(`Processing row ${totalRows}: ${JSON.stringify(flattenedRow)}`);
      
      try {
        // Extract data from flattened row
        const campaignId = flattenedRow['campaign.id'] || '';
        const campaignName = flattenedRow['campaign.name'] || '';
        const campaignStatus = flattenedRow['campaign.status'] || '';
        const experimentType = flattenedRow['campaign.experimentType'] || '';
        const campaignEndDate = flattenedRow['campaign.endDate'] || '';
        const adGroupId = flattenedRow['adGroup.id'] || '';
        const adGroupName = flattenedRow['adGroup.name'] || '';
        const adId = flattenedRow['adGroupAd.ad.id'] || '';
        const adType = flattenedRow['adGroupAd.ad.type'] || '';
        const adStatus = flattenedRow['adGroupAd.status'] || '';
        const approvalStatus = flattenedRow['adGroupAd.policySummary.approvalStatus'] || '';
        const impressions = parseInt(flattenedRow['metrics.impressions'] || '0');
        
        // Skip if this is an ended EXPERIMENT campaign (BASE campaigns are always monitored)
        if (experimentType === 'EXPERIMENT' && campaignEndDate && campaignEndDate !== '' && campaignEndDate < dateStr) {
          Logger.log(`⏭️ Skipping ended experiment campaign: ${campaignName} (ended: ${campaignEndDate})`);
          continue;
        }
        
        // Skip if campaign has ended (only if end_date is set and is in the past)
        if (campaignEndDate && campaignEndDate !== '' && campaignEndDate < dateStr) {
          Logger.log(`⏭️ Skipping ended campaign: ${campaignName} (ended: ${campaignEndDate})`);
          continue;
        }
        
        // Count campaigns and ad groups (unique)
        if (!processedCampaigns.has(campaignId)) {
          processedCampaigns.add(campaignId);
          campaignCount++;
          Logger.log(`✓ Campaign counted: ${campaignName}`);
        }
        
        if (!processedAdGroups.has(adGroupId)) {
          processedAdGroups.add(adGroupId);
          adGroupCount++;
          Logger.log(`✓ Ad Group counted: ${adGroupName}`);
        }
        
        adCount++;
        Logger.log(`✓ Ad counted: ${adId} (${adType}) - Impressions: ${impressions}`);
        
        // Check for zero impressions
        if (impressions === 0) {
          zeroImpressionsAds++;
          const issue = {
            campaignName,
            adGroupName,
            adId,
            adType,
            impressions,
            date: dateStr
          };
          adIssues.push(issue);
          Logger.log(`⚠️ ZERO IMPRESSIONS AD: ${adId} (${adType}) - Impressions: ${impressions}`);
        }
        
        Logger.log(`Row ${totalRows} data: Campaign=${campaignName}(${campaignId}), Status=${campaignStatus}, Experiment=${experimentType}, EndDate=${campaignEndDate}, AdGroup=${adGroupName}(${adGroupId}), Ad=${adId}, Type=${adType}, Status=${adStatus}, Approval=${approvalStatus}, Impressions=${impressions}`);
        
      } catch (rowError) {
        Logger.log(`❌ Error processing row ${totalRows}: ${rowError.message}`);
      }
    }
    
    Logger.log(`Finished processing ${totalRows} GAQL rows`);
    
    return {
      totalRows,
      campaignCount,
      adGroupCount,
      adCount,
      zeroImpressionsAds,
      adIssues
    };
    
  } catch (error) {
    Logger.log(`❌ Error getting ad zero impressions issues: ${error.message}`);
    throw error;
  }
}

/**
 * Get zero impressions issues for Performance Max asset groups
 */
function getAssetGroupZeroImpressionsIssues(accountName, dateStr) {
  try {
    Logger.log('Executing asset group query for zero impressions...');
    
    const query = `
      SELECT 
        campaign.id,
        campaign.name,
        campaign.status,
        campaign.experiment_type,
        campaign.end_date,
        asset_group.id,
        asset_group.name,
        asset_group.status,
        metrics.impressions
      FROM asset_group
      WHERE 
        campaign.status = 'ENABLED' AND
        asset_group.status = 'ENABLED' AND
        segments.date = '${dateStr}'
      ORDER BY campaign.name, asset_group.name
    `;
    
    Logger.log(`Executing query: ${query}`);
    
    const rows = AdsApp.search(query, { 'apiVersion': 'v20' });
    let totalRows = 0;
    let assetGroupCount = 0;
    let zeroImpressionsAssetGroups = 0;
    const assetGroupIssues = [];
    
    // Log first row structure for debugging
    if (rows.hasNext()) {
      const firstRow = rows.next();
      const flattenedRow = flattenObject(firstRow);
      Logger.log(`🔍 First asset group row available fields: ${Object.keys(flattenedRow).join(', ')}`);
      Logger.log(`🔍 Sample asset group field values:`);
      Logger.log(`   campaign.id: ${flattenedRow['campaign.id']}`);
      Logger.log(`   campaign.name: ${flattenedRow['campaign.name']}`);
      Logger.log(`   campaign.status: ${flattenedRow['campaign.status']}`);
      Logger.log(`   campaign.experimentType: ${flattenedRow['campaign.experimentType']}`);
      Logger.log(`   campaign.endDate: ${flattenedRow['campaign.endDate']}`);
      Logger.log(`   assetGroup.id: ${flattenedRow['assetGroup.id']}`);
      Logger.log(`   assetGroup.name: ${flattenedRow['assetGroup.name']}`);
      Logger.log(`   assetGroup.status: ${flattenedRow['assetGroup.status']}`);
      Logger.log(`   metrics.impressions: ${flattenedRow['metrics.impressions']}`);
    }
    
    Logger.log(`Executing main asset group query for processing...`);
    
    // Execute the query again for processing (since we can't reset the iterator)
    const processingRows = AdsApp.search(query, { 'apiVersion': 'v20' });
    
    while (processingRows.hasNext()) {
      totalRows++;
      const row = processingRows.next();
      const flattenedRow = flattenObject(row);
      
      try {
        // Extract data from flattened row
        const campaignId = flattenedRow['campaign.id'] || '';
        const campaignName = flattenedRow['campaign.name'] || '';
        const campaignStatus = flattenedRow['campaign.status'] || '';
        const experimentType = flattenedRow['campaign.experimentType'] || '';
        const campaignEndDate = flattenedRow['campaign.endDate'] || '';
        const assetGroupId = flattenedRow['assetGroup.id'] || '';
        const assetGroupName = flattenedRow['assetGroup.name'] || '';
        const assetGroupStatus = flattenedRow['assetGroup.status'] || '';
        const assetGroupType = 'PERFORMANCE_MAX';
        const impressions = parseInt(flattenedRow['metrics.impressions'] || '0');
        
        // Skip if this is an ended EXPERIMENT campaign (BASE campaigns are always monitored)
        if (experimentType === 'EXPERIMENT' && campaignEndDate && campaignEndDate !== '' && campaignEndDate < dateStr) {
          Logger.log(`⏭️ Skipping ended experiment asset group: ${assetGroupName} in ${campaignName} (ended: ${campaignEndDate})`);
          continue;
        }
        
        // Skip if campaign has ended (only if end_date is set and is in the past)
        if (campaignEndDate && campaignEndDate !== '' && campaignEndDate < dateStr) {
          Logger.log(`⏭️ Skipping ended asset group: ${assetGroupName} in ${campaignName} (ended: ${campaignEndDate})`);
          continue;
        }
        
        assetGroupCount++;
        Logger.log(`Asset Group ${totalRows}: Campaign=${campaignName}(${campaignId}), Status=${campaignStatus}, Experiment=${experimentType}, EndDate=${campaignEndDate}, AssetGroup=${assetGroupName}(${assetGroupId}), Type=${assetGroupType}, Status=${assetGroupStatus}, Impressions=${impressions}`);
        Logger.log(`✓ Asset Group counted: ${assetGroupName} (${assetGroupType})`);
        
        // Check for zero impressions
        if (impressions === 0) {
          zeroImpressionsAssetGroups++;
          const issue = {
            campaignName,
            assetGroupName,
            assetGroupId,
            assetGroupType,
            impressions,
            date: dateStr
          };
          assetGroupIssues.push(issue);
          Logger.log(`⚠️ ZERO IMPRESSIONS ASSET GROUP: ${assetGroupName} (${assetGroupType}) - Impressions: ${impressions}`);
        }
        
      } catch (rowError) {
        Logger.log(`❌ Error processing asset group row ${totalRows}: ${rowError.message}`);
      }
    }
    
    if (totalRows === 0) {
      Logger.log(`No asset groups found for ${accountName}`);
    } else {
      Logger.log(`✓ Successfully processed ${totalRows} asset group rows`);
    }
    
    return {
      totalRows,
      assetGroupCount,
      zeroImpressionsAssetGroups,
      assetGroupIssues
    };
    
  } catch (error) {
    Logger.log(`❌ Error getting asset group zero impressions issues: ${error.message}`);
    return {
      totalRows: 0,
      assetGroupCount: 0,
      zeroImpressionsAssetGroups: 0,
      assetGroupIssues: []
    };
  }
}

/**
 * Send individual email alert to specific account contact person
 */
function sendIndividualZeroImpressionsAlertEmail(accountData) {
  try {
    const subject = `🚨 Google Ads Zero Impressions Alert - ${accountData.accountName} - ${new Date().toLocaleDateString()}`;
    
    let htmlBody = `
      <h2>🚨 Google Ads Zero Impressions Alert</h2>
      <p><strong>Date:</strong> ${new Date().toLocaleString()}</p>
      <p><strong>Account:</strong> ${accountData.accountName}</p>
      <p><strong>Monitoring Period:</strong> Yesterday (Business Day Monitoring)</p>
      
      <h3>📊 Account Summary</h3>
      <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
        <tr style="background-color: #f2f2f2;">
          <th>Metric</th>
          <th>Count</th>
        </tr>
        <tr>
          <td><strong>Campaigns Checked</strong></td>
          <td>${accountData.campaignCount}</td>
        </tr>
        <tr>
          <td><strong>Ad Groups Checked</strong></td>
          <td>${accountData.adGroupCount}</td>
        </tr>
        <tr>
          <td><strong>Ads Checked</strong></td>
          <td>${accountData.adCount}</td>
        </tr>
        <tr>
          <td><strong>Asset Groups Checked</strong></td>
          <td>${accountData.assetGroupCount}</td>
        </tr>
        <tr style="color: red;">
          <td><strong>Ads with Zero Impressions</strong></td>
          <td>${accountData.zeroImpressionsAds}</td>
        </tr>
        <tr style="color: red;">
          <td><strong>Asset Groups with Zero Impressions</strong></td>
          <td>${accountData.zeroImpressionsAssetGroups}</td>
        </tr>
      </table>
    `;
    
    // Add detailed zero impressions ads section
    if (accountData.adIssues && accountData.adIssues.length > 0) {
      htmlBody += `<h3>🚫 Ads with Zero Impressions Details</h3>`;
      htmlBody += `<ul>`;
      accountData.adIssues.forEach(issue => {
        htmlBody += `<li><strong>${issue.campaignName}</strong> > <strong>${issue.adGroupName}</strong> > Ad ${issue.adId} (${issue.adType}) - <span style="color: red;">${issue.impressions} impressions</span></li>`;
      });
      htmlBody += `</ul>`;
    }
    
    // Add detailed zero impressions asset groups section
    if (accountData.assetGroupIssues && accountData.assetGroupIssues.length > 0) {
      htmlBody += `<h3>🚫 Asset Groups with Zero Impressions Details</h3>`;
      htmlBody += `<ul>`;
      accountData.assetGroupIssues.forEach(issue => {
        htmlBody += `<li><strong>${issue.campaignName}</strong> > Asset Group ${issue.assetGroupName} (${issue.assetGroupType}) - <span style="color: red;">${issue.impressions} impressions</span></li>`;
      });
      htmlBody += `</ul>`;
    }
    
    htmlBody += `
      <hr>
      <p><em>This alert was generated automatically by the MCC Zero Impressions Monitor script.</em></p>
      <p><em>This email was sent specifically to the contact person for ${accountData.accountName}.</em></p>
      <p><em>Content with zero impressions may indicate issues with targeting, bidding, or ad quality that need attention.</em></p>
      <p><em>Note: This script only runs Monday through Friday (business days only).</em></p>
      <p><em>Filtering: Only active, approved ads from enabled, non-experiment campaigns are monitored (experiments, ineligible, ended campaigns, and draft campaigns are excluded).</em></p>
    `;
    
    // Get the email recipient from the account data
    const emailRecipient = accountData.emailRecipient || 'peterson@creeksidemarketingpros.com';
    
    // Try GmailApp first, fallback to MailApp
    try {
      GmailApp.sendEmail(emailRecipient, subject, '', { htmlBody: htmlBody });
      Logger.log(`✓ Email sent successfully to ${emailRecipient} for ${accountData.accountName} via GmailApp`);
    } catch (gmailError) {
      Logger.log(`⚠️ GmailApp failed for ${accountData.accountName}, trying MailApp: ${gmailError.message}`);
      MailApp.sendEmail(emailRecipient, subject, '', { htmlBody: htmlBody });
      Logger.log(`✓ Email sent successfully to ${emailRecipient} for ${accountData.accountName} via MailApp`);
    }
    
  } catch (error) {
    Logger.log(`❌ Error sending individual zero impressions alert email for ${accountData.accountName}: ${error.message}`);
  }
}

/**
 * Utility function to flatten nested objects
 */
function flattenObject(ob) {
  const toReturn = {};
  
  for (const i in ob) {
    if (!ob.hasOwnProperty(i)) continue;
    
    if ((typeof ob[i]) === 'object' && ob[i] !== null) {
      const flatObject = flattenObject(ob[i]);
      for (const x in flatObject) {
        if (!flatObject.hasOwnProperty(x)) continue;
        toReturn[i + '.' + x] = flatObject[x];
      }
    } else {
      toReturn[i] = ob[i];
    }
  }
  return toReturn;
}

/**
 * Log script health to monitoring spreadsheet
 */
function logScriptHealth(scriptId) {
  const MONITOR_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1nQVIR4Rt6lMA8QDVdTzGEMZSJGozreCdgrQ4s5rh01Q/edit';
  
  try {
    const sheet = SpreadsheetApp.openByUrl(MONITOR_SHEET_URL).getSheetByName('Monitor');
    const dataRange = sheet.getRange(6, 1, sheet.getLastRow() - 5, 1);
    const scriptIds = dataRange.getValues().flat();
    const rowIndex = scriptIds.indexOf(scriptId);
    
    if (rowIndex !== -1) {
      const actualRow = rowIndex + 6;
      sheet.getRange(actualRow, 5).setValue(new Date());
      sheet.getRange(actualRow, 6).setValue(''); // Clear any alert flag
    }
  } catch (e) {
    // Don't let monitoring errors break your script
    Logger.log('Could not update monitor: ' + e.toString());
  }
}

/**
 * Helper function to get day name from day number
 */
function getDayName(dayNumber) {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  return days[dayNumber];
}

/**
 * Test function for debugging
 */
function testScript() {
  Logger.log("=== Testing MCC Zero Impressions Monitor ===");
  main();
}
