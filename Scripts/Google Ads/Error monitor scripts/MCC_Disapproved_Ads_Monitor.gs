// --- MCC DISAPPROVED ADS MONITOR CONFIGURATION ---

// 1. MASTER SPREADSHEET URL
//    - This spreadsheet contains the mapping of Account IDs to their individual spreadsheet URLs
//    - Format: Column A = Client Name, Column B = Account ID (10-digit), Column C = Spreadsheet URL
const MASTER_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1bS1xmTFB0LWGvYeZBPcykGMY0SwZ4W-Ny6aKVloy0SM/edit?usp=sharing";

// 2. MASTER SPREADSHEET TAB NAME
//    - The tab name in the master spreadsheet that contains the account mappings
const MASTER_SHEET_NAME = "AccountMappings"; // Default tab name

// 3. EMAIL CONFIGURATION
// Email recipients will be read from Column D of the master spreadsheet for each account
// Multiple emails can be separated by commas in Column D
const EMAIL_SUBJECT_PREFIX = "[Google Ads Disapproved Ads Alert]";

// 4. TEST CID (Optional)
//    - To run for a single account for testing, enter its Customer ID (e.g., "123-456-7890").
//    - Leave blank ("") to run for all accounts in the master spreadsheet.
const SINGLE_CID_FOR_TESTING = ""; // Example: "123-456-7890" or ""

// 5. DEBUG MODE - Set to true to see detailed logs
const DEBUG_MODE = true;

// 6. TEST MODE - Set to true to simulate finding disapproved ads for testing
const TEST_MODE = false; // Set to true to test email functionality

// 7. RUN MODE - Choose what to run when main() is executed
// Options: "MONITOR" (normal operation), "TEST_EMAIL", "TEST_ACCESS", "COMPREHENSIVE_TEST"
const RUN_MODE = "MONITOR"; // Change this to run different tests

// --- END OF CONFIGURATION ---

function main() {
  Logger.log("=== Starting MCC Disapproved Ads Monitor ===");
  Logger.log(`Configuration:`);
  Logger.log(`  - Master Spreadsheet: ${MASTER_SPREADSHEET_URL}`);
  Logger.log(`  - Master Sheet Name: ${MASTER_SHEET_NAME}`);
  Logger.log(`  - Test Mode: ${TEST_MODE ? 'ENABLED' : 'DISABLED'}`);
  Logger.log(`  - Debug Mode: ${DEBUG_MODE ? 'ENABLED' : 'DISABLED'}`);
  Logger.log(`  - Single CID Testing: ${SINGLE_CID_FOR_TESTING || 'ALL ACCOUNTS'}`);
  Logger.log(`  - Run Mode: ${RUN_MODE}`);
  
  // Choose what to run based on RUN_MODE
  switch (RUN_MODE) {
    case "TEST_EMAIL":
      Logger.log("🔧 Running in TEST_EMAIL mode...");
      testEmailFunctionality();
      break;
    case "TEST_ACCESS":
      Logger.log("🔧 Running in TEST_ACCESS mode...");
      testAccountAccess();
      break;
    case "COMPREHENSIVE_TEST":
      Logger.log("🔧 Running in COMPREHENSIVE_TEST mode...");
      runComprehensiveTest();
      break;
    case "MONITOR":
    default:
      Logger.log("🔧 Running in MONITOR mode (normal operation)...");
      runDisapprovedAdsMonitor();
      break;
  }
}

// Main monitoring function (original main logic)
function runDisapprovedAdsMonitor() {
  Logger.log("=== Starting MCC Disapproved Ads Monitor ===");
  
  if (SINGLE_CID_FOR_TESTING) {
    Logger.log(`Running in test mode for CID: ${SINGLE_CID_FOR_TESTING}`);
  } else {
    Logger.log("Running for all accounts in the master spreadsheet.");
  }

  // Get account mappings from master spreadsheet
  Logger.log("\n📋 Step 1: Reading account mappings from master spreadsheet...");
  const accountMappings = getAccountMappings();
  if (!accountMappings || accountMappings.length === 0) {
    Logger.log("❌ No account mappings found in master spreadsheet.");
    return;
  }

  Logger.log(`✓ Found ${accountMappings.length} account mappings in master spreadsheet.`);
  
  // Log account details for verification
  if (DEBUG_MODE) {
    Logger.log("📋 Account mappings found:");
    accountMappings.forEach((mapping, index) => {
      Logger.log(`  ${index + 1}. ${mapping.accountName} (${mapping.accountId}) - Emails: ${mapping.emailRecipients.join(', ') || 'None'}`);
    });
  }

  let processedCount = 0;
  let successCount = 0;
  let errorCount = 0;
  let totalDisapprovedAds = 0;
  let allDisapprovedAds = [];

  Logger.log("\n🔍 Step 2: Processing each account for disapproved ads...");

  // Process each account
  for (const mapping of accountMappings) {
    const accountId = mapping.accountId;
    const accountName = mapping.accountName || "Unknown";

    // Skip if testing mode and this isn't the test account
    if (SINGLE_CID_FOR_TESTING && accountId !== SINGLE_CID_FOR_TESTING) {
      Logger.log(`⏭️ Skipping ${accountName} (not in test mode)`);
      continue;
    }

    processedCount++;
    Logger.log(`\n--- Processing Account ${processedCount}/${accountMappings.length} ---`);
    Logger.log(`Account: ${accountName} (${accountId})`);
    Logger.log(`Email Recipients: ${mapping.emailRecipients.join(', ') || 'None (will use fallback)'}`);

    try {
      Logger.log(`🔍 Step 2.${processedCount}.1: Finding account in MCC...`);
      // Switch to the account context
      const account = getAccountById(accountId);
      if (!account) {
        Logger.log(`❌ Account ${accountId} not found or not accessible.`);
        errorCount++;
        continue;
      }

      Logger.log(`✓ Account found: ${account.getName()}`);
      Logger.log(`🔍 Step 2.${processedCount}.2: Switching to account context...`);
      AdsManagerApp.select(account);
      Logger.log(`✓ Switched to account: ${account.getName()}`);

      Logger.log(`🔍 Step 2.${processedCount}.3: Checking for disapproved ads...`);
      // Check for disapproved ads in this account
      const disapprovedAds = getDisapprovedAdsForAccount(accountName);
      
      if (disapprovedAds.length > 0) {
        Logger.log(`⚠️ Found ${disapprovedAds.length} disapproved ads in ${accountName}`);
        
        // Send email notification for this specific account if it has email recipients
        if (mapping.emailRecipients && mapping.emailRecipients.length > 0) {
          Logger.log(`📧 Sending email to account-specific recipients: ${mapping.emailRecipients.join(', ')}`);
          sendDisapprovedAdsEmail(disapprovedAds, mapping.emailRecipients, accountName);
          Logger.log(`✓ Email notification sent to ${mapping.emailRecipients.join(', ')} for ${accountName}`);
        } else {
          Logger.log(`⚠️ No email recipients configured for ${accountName} - adding to master list`);
          allDisapprovedAds = allDisapprovedAds.concat(disapprovedAds);
          totalDisapprovedAds += disapprovedAds.length;
        }
      } else {
        Logger.log(`✓ No disapproved ads found in ${accountName}`);
      }
      
      Logger.log(`✓ Successfully checked ${accountName}`);
      successCount++;

    } catch (error) {
      Logger.log(`❌ Error processing account ${accountName} (${accountId}): ${error.message}`);
      Logger.log(`Error details: ${error.stack || 'No stack trace available'}`);
      errorCount++;
    }
  }

  Logger.log("\n📧 Step 3: Processing fallback email notifications...");
  // Send email notification for accounts without specific email recipients
  if (allDisapprovedAds.length > 0) {
    // Send to default email if any accounts don't have specific recipients
    const defaultEmailRecipients = ["petersonrainey@gmail.com"]; // Fallback email
    Logger.log(`📧 Sending fallback email for ${totalDisapprovedAds} disapproved ads from accounts without specific recipients`);
    sendDisapprovedAdsEmail(allDisapprovedAds, defaultEmailRecipients, "Multiple Accounts");
    Logger.log(`✓ Fallback email notification sent for ${totalDisapprovedAds} disapproved ads`);
  } else {
    Logger.log(`📧 No fallback email needed - all accounts had specific recipients or no disapproved ads found`);
  }

  // Summary
  Logger.log(`\n=== MCC DISAPPROVED ADS MONITOR COMPLETED ===`);
  Logger.log(`📊 Summary:`);
  Logger.log(`  - Total accounts processed: ${processedCount}`);
  Logger.log(`  - Successful checks: ${successCount}`);
  Logger.log(`  - Errors: ${errorCount}`);
  Logger.log(`  - Total disapproved ads found: ${totalDisapprovedAds}`);
  Logger.log(`  - Accounts with specific email recipients: ${accountMappings.filter(m => m.emailRecipients && m.emailRecipients.length > 0).length}`);
  Logger.log(`  - Accounts using fallback email: ${accountMappings.filter(m => !m.emailRecipients || m.emailRecipients.length === 0).length}`);
}

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

      if (accountId) {
        // Parse email recipients from Column D (multiple emails separated by commas)
        const emailRecipients = row[3]?.toString().trim() || "";
        const emailList = emailRecipients ? emailRecipients.split(',').map(email => email.trim()).filter(email => email) : [];
        
        Logger.log(`📋 Row ${i + 1}: ${accountName} (${accountId}) - Emails: ${emailList.join(', ') || 'None'}`);
        
        mappings.push({
          accountId: accountId,
          spreadsheetUrl: spreadsheetUrl,
          accountName: accountName,
          emailRecipients: emailList
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

function getAccountById(accountId) {
  try {
    Logger.log(`🔍 Looking for account ID: ${accountId}`);
    const accountIterator = AdsManagerApp.accounts().withIds([accountId]).get();
    if (accountIterator.hasNext()) {
      const account = accountIterator.next();
      Logger.log(`✓ Found account: ${account.getName()} (${account.getCustomerId()})`);
      return account;
    }
    Logger.log(`❌ Account ${accountId} not found in MCC`);
    return null;
  } catch (error) {
    Logger.log(`❌ Error finding account ${accountId}: ${error.message}`);
    return null;
  }
}

function getDisapprovedAdsForAccount(accountName) {
  Logger.log(`🔍 Checking for disapproved ads in ${accountName}...`);
  
  const disapprovedAds = [];
  
  try {
    Logger.log(`📋 Building GAQL queries for ${accountName}...`);
    
    // Query 1: Search, Display, YouTube ads (ad_group_ad)
    const searchDisplayQuery = `
      SELECT 
          campaign.id,
          campaign.name,
          campaign.advertising_channel_type,
          ad_group.id,
          ad_group.name,
          ad_group_ad.ad.id,
          ad_group_ad.ad.name,
          ad_group_ad.status,
          ad_group_ad.ad.final_urls
      FROM ad_group_ad
      WHERE 
          campaign.status = 'ENABLED'
          AND ad_group.status = 'ENABLED'
          AND campaign.advertising_channel_type IN ('SEARCH', 'DISPLAY', 'VIDEO')
      ORDER BY campaign.name ASC, ad_group.name ASC
    `;
    
    // Query 2: Shopping campaigns (product groups)
    const shoppingQuery = `
      SELECT 
          campaign.id,
          campaign.name,
          campaign.advertising_channel_type,
          ad_group.id,
          ad_group.name,
          ad_group_criterion.criterion_id,
          ad_group_criterion.status
      FROM ad_group_criterion
      WHERE 
          campaign.status = 'ENABLED'
          AND campaign.advertising_channel_type = 'SHOPPING'
          AND ad_group.status = 'ENABLED'
      ORDER BY campaign.name ASC, ad_group.name ASC
    `;
    
    // Query 3: Performance Max campaigns (asset groups)
    const pmaxQuery = `
      SELECT 
          campaign.id,
          campaign.name,
          campaign.advertising_channel_type,
          asset_group.id,
          asset_group.name,
          asset_group.status
      FROM asset_group
      WHERE 
          campaign.status = 'ENABLED'
          AND campaign.advertising_channel_type = 'PERFORMANCE_MAX'
      ORDER BY campaign.name ASC, asset_group.name ASC
    `;
    
    // Query 4: Demand Gen campaigns (asset groups)
    const demandGenQuery = `
      SELECT 
          campaign.id,
          campaign.name,
          campaign.advertising_channel_type,
          asset_group.id,
          asset_group.name,
          asset_group.status
      FROM asset_group
      WHERE 
          campaign.status = 'ENABLED'
          AND campaign.advertising_channel_type = 'PERFORMANCE_MAX'
      ORDER BY campaign.name ASC, asset_group.name ASC
    `;
    

    
    // Execute all queries
    let searchDisplayRows = AdsApp.search(searchDisplayQuery);
    processSearchDisplayAds(searchDisplayRows, accountName, disapprovedAds, searchDisplayQuery);
    
    let shoppingRows = AdsApp.search(shoppingQuery);
    processShoppingCampaigns(shoppingRows, accountName, disapprovedAds, shoppingQuery);
    
    let pmaxRows = AdsApp.search(pmaxQuery);
    processPmaxAssets(pmaxRows, accountName, disapprovedAds, pmaxQuery);
    
    let demandGenRows = AdsApp.search(demandGenQuery);
    processDemandGenAssets(demandGenRows, accountName, disapprovedAds, demandGenQuery);
    
    Logger.log(`✓ Found ${disapprovedAds.length} disapproved ads in ${accountName}`);
    return disapprovedAds;
    
  } catch (error) {
    Logger.log(`❌ Error getting disapproved ads for ${accountName}: ${error.message}`);
    Logger.log(`Error details: ${error.stack || 'No stack trace available'}`);
    return [];
  }
}

function sendDisapprovedAdsEmail(disapprovedAds, emailRecipients, accountContext) {
  try {
    Logger.log(`📧 Preparing email notification for ${accountContext}...`);
    Logger.log(`📧 Email recipients: ${emailRecipients.join(', ')}`);
    
    const currentDate = new Date().toLocaleDateString();
    const totalAds = disapprovedAds.length;
    
    Logger.log(`📧 Building email content for ${totalAds} disapproved ads...`);
    
    // Group by account for better organization
    const adsByAccount = {};
    disapprovedAds.forEach(ad => {
      if (!adsByAccount[ad.accountName]) {
        adsByAccount[ad.accountName] = [];
      }
      adsByAccount[ad.accountName].push(ad);
    });
    
    // Build email body
    let emailBody = `
<h2>🚨 Google Ads Disapproved Ads Alert</h2>
<p><strong>Date:</strong> ${currentDate}</p>
<p><strong>Total Disapproved Ads Found:</strong> ${totalAds}</p>
<hr>
`;

    // Add content for each account
    Object.keys(adsByAccount).forEach(accountName => {
      const accountAds = adsByAccount[accountName];
      
      emailBody += `
<h3>📊 Account: ${accountName}</h3>
<p><strong>Disapproved Ads:</strong> ${accountAds.length}</p>
<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
  <tr style="background-color: #f2f2f2;">
    <th>Campaign</th>
    <th>Campaign Type</th>
    <th>Ad Group/Asset Group</th>
    <th>Asset Type</th>
    <th>Ad Text/Asset ID</th>
    <th>Disapproval Reason</th>
    <th>Date Found</th>
  </tr>
`;
      
      accountAds.forEach(ad => {
        emailBody += `
  <tr>
    <td>${ad.campaignName}</td>
    <td>${ad.campaignType || 'Unknown'}</td>
    <td>${ad.adGroupName}</td>
    <td>${ad.assetType || 'Unknown'}</td>
    <td>${ad.adText}</td>
    <td>${ad.disapprovalReason}</td>
    <td>${ad.dateFound}</td>
  </tr>
`;
      });
      
      emailBody += `</table>`;
    });
    
    emailBody += `
<hr>
<p><strong>Action Required:</strong></p>
<ul>
  <li>Review each disapproved ad in Google Ads</li>
  <li>Fix the policy violations mentioned in the disapproval reason</li>
  <li>Resubmit ads for approval</li>
  <li>Monitor for new disapprovals</li>
</ul>

<p><em>This alert was generated by the MCC Disapproved Ads Monitor script.</em>
`;
    
    // Send email
    const subject = `${EMAIL_SUBJECT_PREFIX} ${totalAds} Disapproved Ads Found - ${accountContext} - ${currentDate}`;
    
    Logger.log(`📧 Sending email with subject: ${subject}`);
    
    GmailApp.sendEmail(
      emailRecipients.join(','),
      subject,
      emailBody.replace(/<[^>]*>/g, ''), // Plain text version
      {
        htmlBody: emailBody,
        name: 'Google Ads Disapproved Ads Monitor'
      }
    );
    
    Logger.log(`✓ Email notification sent successfully to ${emailRecipients.join(', ')}`);
    
  } catch (error) {
    Logger.log(`❌ Error sending email: ${error.message}`);
    Logger.log(`Error details: ${error.stack || 'No stack trace available'}`);
  }
}

// Process Search and Display ads
function processSearchDisplayAds(rows, accountName, disapprovedAds, query) {
  let rowCount = 0;
  const statusCounts = {
    ENABLED: 0,
    PAUSED: 0,
    REMOVED: 0,
    PENDING: 0,
    DISAPPROVED: 0,
    UNDER_REVIEW: 0,
    UNKNOWN: 0
  };
  
  // Process Search/Display/YouTube ads silently
  
  // Log sample row structure for debugging (only in extreme debug mode)
  if (DEBUG_MODE && DEBUG_MODE === 'EXTREME' && rows.hasNext()) {
    Logger.log(`🔍 Logging sample Search/Display row structure for ${accountName}...`);
    const sampleRow = rows.next();
    Logger.log("Sample Search/Display row structure: " + JSON.stringify(sampleRow));
    // Reset the iterator by creating a new search
    const resetRows = AdsApp.search(query);
    rows = resetRows;
  }
  
  while (rows.hasNext()) {
    try {
      const row = rows.next();
      rowCount++;
      
      // Access nested objects correctly
      const campaign = row.campaign || {};
      const adGroup = row.adGroup || {};
      const adGroupAd = row.adGroupAd || {};
      const ad = adGroupAd.ad || {};
      
      const campaignId = campaign.id || '';
      const campaignName = campaign.name || 'Unknown Campaign';
      const campaignType = campaign.campaignType || campaign.advertisingChannelType || 'UNKNOWN';
      const adGroupId = adGroup.id || '';
      const adGroupName = adGroup.name || 'Unknown Ad Group';
      const adId = ad.id || '';
      const adName = ad.name || 'Unknown Ad';
      const adStatus = adGroupAd.status || 'UNKNOWN';
      
      // Count status
      if (statusCounts.hasOwnProperty(adStatus)) {
        statusCounts[adStatus]++;
      } else {
        statusCounts.UNKNOWN++;
      }
      
      // Check if this ad is disapproved
      if (adStatus !== 'DISAPPROVED' && adStatus !== 'UNDER_REVIEW') {
        continue;
      }
      
      Logger.log(`📋 Found disapproved ad: ${campaignName} (${campaignType}) > ${adGroupName} > ${adName}`);
      
      // Get disapproval reason - policy summary not available in GAQL
      let disapprovalReason = 'Ad disapproved - check Google Ads interface for specific reason';
      
      Logger.log(`📋 Disapproval reason: ${disapprovalReason}`);
      
      // Get ad text (headline) - try to get from ad name or final URLs
      let adText = adName;
      if (ad.finalUrls && ad.finalUrls.length > 0) {
        adText = `Ad: ${adName} | URL: ${ad.finalUrls[0]}`;
      }
      
      // Determine asset type based on campaign type
      let assetType = 'Search/Display Ad';
      if (campaignType === 'PERFORMANCE_MAX') {
        assetType = 'Demand Gen Ad';
      } else if (campaignType === 'VIDEO') {
        assetType = 'YouTube Ad';
      }
      
      const disapprovedAd = {
        accountName: accountName,
        campaignName: campaignName,
        campaignType: campaignType,
        adGroupName: adGroupName,
        adText: adText,
        disapprovalReason: disapprovalReason,
        dateFound: new Date().toISOString().split('T')[0], // YYYY-MM-DD format
        campaignId: campaignId,
        adGroupId: adGroupId,
        adId: adId,
        assetType: assetType
      };
      
      disapprovedAds.push(disapprovedAd);
      
      if (DEBUG_MODE) {
        Logger.log(`✓ Added disapproved Search/Display ad to list: ${campaignName} > ${adGroupName} > ${adText}`);
      }
      
    } catch (rowError) {
      Logger.log(`❌ Error processing Search/Display row ${rowCount}: ${rowError.message}`);
      // Continue with next row
    }
  }
  
  // Processed Search/Display/YouTube ads
  
  // Log status summary
  logStatusSummary(accountName, 'Search/Display/YouTube Ads', statusCounts);
}

// Process Shopping campaigns
function processShoppingCampaigns(rows, accountName, disapprovedAds, query) {
  let rowCount = 0;
  const statusCounts = {
    ENABLED: 0,
    PAUSED: 0,
    REMOVED: 0,
    PENDING: 0,
    DISAPPROVED: 0,
    UNDER_REVIEW: 0,
    UNKNOWN: 0
  };
  
  // Process Shopping campaigns silently
  
  // Log sample row structure for debugging (only in extreme debug mode)
  if (DEBUG_MODE && DEBUG_MODE === 'EXTREME' && rows.hasNext()) {
    Logger.log(`🔍 Logging sample Shopping row structure for ${accountName}...`);
    const sampleRow = rows.next();
    Logger.log("Sample Shopping row structure: " + JSON.stringify(sampleRow));
    // Reset the iterator by creating a new search
    const resetRows = AdsApp.search(query);
    rows = resetRows;
  }
  
  while (rows.hasNext()) {
    try {
      const row = rows.next();
      rowCount++;
      
      // Access nested objects correctly
      const campaign = row.campaign || {};
      const adGroup = row.adGroup || {};
      const adGroupCriterion = row.adGroupCriterion || {};
      
      const campaignId = campaign.id || '';
      const campaignName = campaign.name || 'Unknown Campaign';
      const campaignType = campaign.advertisingChannelType || 'UNKNOWN';
      const adGroupId = adGroup.id || '';
      const adGroupName = adGroup.name || 'Unknown Ad Group';
      const criterionId = adGroupCriterion.criterionId || '';
      const criterionStatus = adGroupCriterion.status || 'UNKNOWN';
      
      // Count status
      if (statusCounts.hasOwnProperty(criterionStatus)) {
        statusCounts[criterionStatus]++;
      } else {
        statusCounts.UNKNOWN++;
      }
      
      // Check if this criterion is disapproved
      if (criterionStatus !== 'DISAPPROVED' && criterionStatus !== 'UNDER_REVIEW') {
        continue;
      }
      
      Logger.log(`📋 Found disapproved Shopping product: ${campaignName} (${campaignType}) > ${adGroupName} > Product ID: ${criterionId}`);
      
      // Get disapproval reason - policy summary not available in GAQL
      let disapprovalReason = 'Shopping product disapproved - check Google Ads interface for specific reason';
      
      Logger.log(`📋 Disapproval reason: ${disapprovalReason}`);
      
      const disapprovedAd = {
        accountName: accountName,
        campaignName: campaignName,
        campaignType: campaignType,
        adGroupName: adGroupName,
        adText: `Shopping Product ID: ${criterionId}`,
        disapprovalReason: disapprovalReason,
        dateFound: new Date().toISOString().split('T')[0], // YYYY-MM-DD format
        campaignId: campaignId,
        adGroupId: adGroupId,
        adId: criterionId,
        assetType: 'Shopping Product'
      };
      
      disapprovedAds.push(disapprovedAd);
      
      if (DEBUG_MODE) {
        Logger.log(`✓ Added disapproved Shopping product to list: ${campaignName} > ${adGroupName} > Product ID: ${criterionId}`);
      }
      
    } catch (rowError) {
      Logger.log(`❌ Error processing Shopping row ${rowCount}: ${rowError.message}`);
      // Continue with next row
    }
  }
  
  // Processed Shopping campaigns
  
  // Log status summary
  logStatusSummary(accountName, 'Shopping Campaigns', statusCounts);
}

// Process Performance Max assets
function processPmaxAssets(rows, accountName, disapprovedAds, query) {
  let rowCount = 0;
  const statusCounts = {
    ENABLED: 0,
    PAUSED: 0,
    REMOVED: 0,
    PENDING: 0,
    DISAPPROVED: 0,
    UNDER_REVIEW: 0,
    UNKNOWN: 0
  };
  
  // Process Performance Max assets silently
  
  // Log sample row structure for debugging (only in extreme debug mode)
  if (DEBUG_MODE && DEBUG_MODE === 'EXTREME' && rows.hasNext()) {
    Logger.log(`🔍 Logging sample PMax row structure for ${accountName}...`);
    const sampleRow = rows.next();
    Logger.log("Sample PMax row structure: " + JSON.stringify(sampleRow));
    // Reset the iterator by creating a new search
    const resetRows = AdsApp.search(query);
    rows = resetRows;
  }
  
  while (rows.hasNext()) {
    try {
      const row = rows.next();
      rowCount++;
      
      // Access nested objects correctly
      const campaign = row.campaign || {};
      const assetGroup = row.assetGroup || {};
      
      const campaignId = campaign.id || '';
      const campaignName = campaign.name || 'Unknown Campaign';
      const campaignType = campaign.campaignType || campaign.advertisingChannelType || 'UNKNOWN';
      const assetGroupId = assetGroup.id || '';
      const assetGroupName = assetGroup.name || 'Unknown Asset Group';
      const assetStatus = assetGroup.status || 'UNKNOWN';
      
      // Count status
      if (statusCounts.hasOwnProperty(assetStatus)) {
        statusCounts[assetStatus]++;
      } else {
        statusCounts.UNKNOWN++;
      }
      
      // Check if this asset group is disapproved
      if (assetStatus !== 'DISAPPROVED' && assetStatus !== 'UNDER_REVIEW') {
        continue;
      }
      
      Logger.log(`📋 Found disapproved PMax asset: ${campaignName} (${campaignType}) > ${assetGroupName}`);
      
      // Get disapproval reason - policy summary not available in GAQL
      let disapprovalReason = 'Performance Max asset group disapproved - check Google Ads interface for specific reason';
      
      Logger.log(`📋 Disapproval reason: ${disapprovalReason}`);
      
      const disapprovedAd = {
        accountName: accountName,
        campaignName: campaignName,
        campaignType: campaignType,
        adGroupName: assetGroupName,
        adText: `PMax Asset Group: ${assetGroupName}`,
        disapprovalReason: disapprovalReason,
        dateFound: new Date().toISOString().split('T')[0], // YYYY-MM-DD format
        campaignId: campaignId,
        adGroupId: assetGroupId,
        adId: assetGroupId,
        assetType: 'Performance Max Asset Group'
      };
      
      disapprovedAds.push(disapprovedAd);
      
      if (DEBUG_MODE) {
        Logger.log(`✓ Added disapproved PMax asset group to list: ${campaignName} > ${assetGroupName}`);
      }
      
    } catch (rowError) {
      Logger.log(`❌ Error processing PMax row ${rowCount}: ${rowError.message}`);
      // Continue with next row
    }
  }
  
  // Processed Performance Max assets
  
  // Log status summary
  logStatusSummary(accountName, 'Performance Max Assets', statusCounts);
}

// Process Demand Gen asset groups
function processDemandGenAssets(rows, accountName, disapprovedAds, query) {
  let rowCount = 0;
  const statusCounts = {
    ENABLED: 0,
    PAUSED: 0,
    REMOVED: 0,
    PENDING: 0,
    DISAPPROVED: 0,
    UNDER_REVIEW: 0,
    UNKNOWN: 0
  };
  
  // Process Demand Gen asset groups silently
  
  // Log sample row structure for debugging (only in extreme debug mode)
  if (DEBUG_MODE && DEBUG_MODE === 'EXTREME' && rows.hasNext()) {
    Logger.log(`🔍 Logging sample Demand Gen row structure for ${accountName}...`);
    const sampleRow = rows.next();
    Logger.log("Sample Demand Gen row structure: " + JSON.stringify(sampleRow));
    // Reset the iterator by creating a new search
    const resetRows = AdsApp.search(query);
    rows = resetRows;
  }
  
  while (rows.hasNext()) {
    try {
      const row = rows.next();
      rowCount++;
      
      const campaign = row.campaign;
      const assetGroup = row.assetGroup;
      
      const campaignId = campaign.id || '';
      const campaignName = campaign.name || 'Unknown Campaign';
      const campaignType = campaign.advertisingChannelType || 'UNKNOWN';
      const assetGroupId = assetGroup.id || '';
      const assetGroupName = assetGroup.name || 'Unknown Asset Group';
      const assetGroupStatus = assetGroup.status || 'UNKNOWN';

      
      // Count status
      if (statusCounts.hasOwnProperty(assetGroupStatus)) {
        statusCounts[assetGroupStatus]++;
      } else {
        statusCounts.UNKNOWN++;
      }
      
      // Check if this asset group is disapproved
      if (assetGroupStatus !== 'DISAPPROVED' && assetGroupStatus !== 'UNDER_REVIEW') {
        continue;
      }
      
      Logger.log(`📋 Found disapproved Demand Gen asset group: ${campaignName} > ${assetGroupName}`);
      
      // Get disapproval reason - generic message since specific policy details not available
      let disapprovalReason = 'Demand Gen asset group disapproved - check Google Ads interface for specific reason';
      
      Logger.log(`📋 Disapproval reason: ${disapprovalReason}`);
      
      const disapprovedAd = {
        accountName: accountName,
        campaignName: campaignName,
        campaignType: campaignType,
        adGroupName: assetGroupName,
        adText: `Asset Group: ${assetGroupName}`,
        disapprovalReason: disapprovalReason,
        dateFound: new Date().toISOString().split('T')[0], // YYYY-MM-DD format
        campaignId: campaignId,
        adGroupId: assetGroupId,
        adId: assetGroupId,
        assetType: 'Demand Gen Asset Group'
      };
      
      disapprovedAds.push(disapprovedAd);
      
    } catch (error) {
      Logger.log(`❌ Error processing Demand Gen row ${rowCount} for ${accountName}: ${error.message}`);
      if (DEBUG_MODE) {
        Logger.log(`Error details: ${error}`);
      }
    }
  }
  
  // Processed Demand Gen asset groups
  
  // Log status summary
  logStatusSummary(accountName, 'Demand Gen Asset Groups', statusCounts);
  
  // Note: Individual asset disapproval status is not directly queryable via GAQL
  // The disapproval status shown in Google Ads interface is determined by policy checks at runtime
  // This script can only detect asset groups that are disapproved at the group level
  if (rowCount > 0) {
    Logger.log(`ℹ️ Note: Individual asset disapproval within Demand Gen asset groups is not queryable via GAQL.`);
    Logger.log(`ℹ️ Check the Google Ads interface for specific asset disapproval details.`);
  }
}

// Log status summary for each campaign type
function logStatusSummary(accountName, campaignType, statusCounts) {
  const total = Object.values(statusCounts).reduce((a, b) => a + b, 0);
  const disapproved = statusCounts.DISAPPROVED || 0;
  const underReview = statusCounts.UNDER_REVIEW || 0;
  
  if (disapproved > 0 || underReview > 0) {
    Logger.log(`🚨 ${accountName} - ${campaignType}: ${disapproved} disapproved, ${underReview} under review (Total: ${total})`);
  } else {
    Logger.log(`✓ ${accountName} - ${campaignType}: ${total} items, all approved`);
  }
}

// Test function to run the monitor manually
function testDisapprovedAdsMonitor() {
  Logger.log("=== Testing MCC Disapproved Ads Monitor ===");
  main();
}

// Function to check a specific account (for testing)
function checkSpecificAccount(accountId) {
  Logger.log(`=== Checking Specific Account: ${accountId} ===`);
  
  try {
    const account = getAccountById(accountId);
    if (!account) {
      Logger.log(`❌ Account ${accountId} not found or not accessible.`);
      return;
    }

    AdsManagerApp.select(account);
    Logger.log(`✓ Switched to account: ${account.getName()}`);

    const disapprovedAds = getDisapprovedAdsForAccount(account.getName());
    
    if (disapprovedAds.length > 0) {
      Logger.log(`⚠️ Found ${disapprovedAds.length} disapproved ads`);
      // Use default email recipients for testing
      const defaultEmailRecipients = ["petersonrainey@gmail.com"];
      sendDisapprovedAdsEmail(disapprovedAds, defaultEmailRecipients, account.getName());
    } else {
      Logger.log(`✓ No disapproved ads found`);
    }
    
  } catch (error) {
    Logger.log(`❌ Error checking account ${accountId}: ${error.message}`);
  }
}

// Function to test email functionality without disapproved ads
function testEmailFunctionality() {
  Logger.log("=== Testing Email Functionality ===");
  
  // Create test disapproved ads data
  const testDisapprovedAds = [
    {
      accountName: "Test Account 1",
      campaignName: "Test Campaign 1",
      adGroupName: "Test Ad Group 1",
      adText: "Test Ad Headline 1",
      disapprovalReason: "Test disapproval reason 1",
      dateFound: new Date().toISOString().split('T')[0],
      campaignId: "123456789",
      adGroupId: "987654321",
      adId: "111222333"
    },
    {
      accountName: "Test Account 2",
      campaignName: "Test Campaign 2",
      adGroupName: "Test Ad Group 2",
      adText: "Test Ad Headline 2",
      disapprovalReason: "Test disapproval reason 2",
      dateFound: new Date().toISOString().split('T')[0],
      campaignId: "456789123",
      adGroupId: "321654987",
      adId: "444555666"
    }
  ];
  
  Logger.log(`📧 Testing email with ${testDisapprovedAds.length} test disapproved ads`);
  
  // Test email recipients
  const testEmailRecipients = ["petersonrainey@gmail.com"];
  
  // Send test email
  sendDisapprovedAdsEmail(testDisapprovedAds, testEmailRecipients, "TEST ACCOUNT");
  
  Logger.log("✓ Test email sent successfully!");
  Logger.log("📧 Check your email to verify the email functionality works correctly.");
}

// Function to test account access and mapping
function testAccountAccess() {
  Logger.log("=== Testing Account Access and Mapping ===");
  
  try {
    Logger.log("📋 Testing master spreadsheet access...");
    const accountMappings = getAccountMappings();
    
    if (accountMappings && accountMappings.length > 0) {
      Logger.log(`✓ Successfully read ${accountMappings.length} account mappings`);
      
      // Test first account access
      const firstMapping = accountMappings[0];
      Logger.log(`🔍 Testing access to first account: ${firstMapping.accountName} (${firstMapping.accountId})`);
      
      const account = getAccountById(firstMapping.accountId);
      if (account) {
        Logger.log(`✓ Successfully accessed account: ${account.getName()}`);
        Logger.log(`✓ Account customer ID: ${account.getCustomerId()}`);
        Logger.log(`✓ Account status: ${account.getStatus()}`);
        
        // Test switching to account
        AdsManagerApp.select(account);
        Logger.log(`✓ Successfully switched to account context`);
        
        // Test GAQL query execution
        Logger.log(`🔍 Testing GAQL query execution...`);
        const testQuery = `
          SELECT 
              campaign.id,
              campaign.name
          FROM campaign
          LIMIT 1
        `;
        
        const testRows = AdsApp.search(testQuery);
        if (testRows.hasNext()) {
          const testRow = testRows.next();
          Logger.log(`✓ GAQL query executed successfully`);
          Logger.log(`✓ Sample campaign: ${testRow.campaign.name}`);
        } else {
          Logger.log(`⚠️ GAQL query executed but no campaigns found`);
        }
        
      } else {
        Logger.log(`❌ Failed to access account: ${firstMapping.accountId}`);
      }
      
    } else {
      Logger.log(`❌ No account mappings found`);
    }
    
  } catch (error) {
    Logger.log(`❌ Error during account access test: ${error.message}`);
    Logger.log(`Error details: ${error.stack || 'No stack trace available'}`);
  }
}

// Function to run comprehensive testing
function runComprehensiveTest() {
  Logger.log("=== Running Comprehensive Test ===");
  
  Logger.log("🔍 Step 1: Testing account access and mapping...");
  testAccountAccess();
  
  Logger.log("\n📧 Step 2: Testing email functionality...");
  testEmailFunctionality();
  
  Logger.log("\n✅ Comprehensive test completed!");
  Logger.log("📋 Check the logs above for any issues.");
  Logger.log("📧 Check your email for the test email to verify email functionality.");
}
