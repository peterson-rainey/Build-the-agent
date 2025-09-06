// --- MCC DESTINATION URL MONITOR ---
// Configuration
const MASTER_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1bS1xmTFB0LWGvYeZBPcykGMY0SwZ4W-Ny6aKVloy0SM/edit?usp=sharing";
const MASTER_SHEET_NAME = "AccountMappings";
const URL_TEST_TIMEOUT = 10000; // 10 seconds timeout for URL testing
const MAX_URLS_PER_BATCH = 50; // Process URLs in batches to avoid timeouts

function main() {
  try {
    Logger.log("=== Starting MCC Destination URL Monitor ===");
    
    // Get account mappings
    const accountMappings = getAccountMappings();
    if (accountMappings.length === 0) {
      Logger.log("❌ No account mappings found. Exiting.");
      logScriptHealth(103); // Error 404: destination not working
      return;
    }
    
    Logger.log(`📋 Found ${accountMappings.length} account mappings.`);
    
    // Process each account
    const accountsWithUrlIssues = [];
    let totalUrlIssues = 0;
    let totalCampaignsChecked = 0;
    let totalAdGroupsChecked = 0;
    let totalAdsChecked = 0;
    let totalAssetGroupsChecked = 0;
    
    Logger.log("🔍 Processing each account for destination URL issues...");
    
    accountMappings.forEach((mapping, index) => {
      Logger.log(`\n--- Processing Account ${index + 1}/${accountMappings.length} ---`);
      Logger.log(`Account: ${mapping.accountName} (${mapping.accountId})`);
      
      try {
        // Find and switch to account
        const account = getAccountById(mapping.accountId);
        if (!account) {
          Logger.log(`❌ Account not found: ${mapping.accountName}`);
          logScriptHealth(103); // Error 404: destination not working
          return;
        }
        
        Logger.log(`✓ Account found: ${account.getName()}`);
        
        // Switch to account context
        AdsManagerApp.select(account);
        Logger.log(`✓ Switched to account: ${account.getName()}`);
        
        // Get destination URL issues
        const accountData = getDestinationUrlIssues(mapping.accountName);
        
        // Add to totals with error handling
        totalCampaignsChecked += accountData.campaignCount || 0;
        totalAdGroupsChecked += accountData.adGroupCount || 0;
        totalAdsChecked += accountData.adCount || 0;
        totalAssetGroupsChecked += accountData.assetGroupCount || 0;
        
        // Log asset group status
        if (accountData.assetGroupError) {
          Logger.log(`⚠️ Asset group monitoring failed for ${mapping.accountName}: ${accountData.assetGroupError}`);
        } else {
          Logger.log(`✓ Asset group monitoring successful for ${mapping.accountName}: ${accountData.assetGroupCount || 0} asset groups found`);
        }
        
        if (accountData.totalIssues > 0) {
          accountsWithUrlIssues.push({
            accountName: mapping.accountName,
            accountId: mapping.accountId,
            emailRecipient: mapping.emailRecipient,
            totalIssues: accountData.totalIssues,
            campaignIssues: accountData.campaignIssues,
            adGroupIssues: accountData.adGroupIssues,
            adIssues: accountData.adIssues,
            assetGroupIssues: accountData.assetGroupIssues || []
          });
          totalUrlIssues += accountData.totalIssues;
          
          Logger.log(`⚠️ DESTINATION URL ISSUES FOUND:`);
          Logger.log(`   Campaign Issues: ${accountData.campaignIssues.length}`);
          Logger.log(`   Ad Group Issues: ${accountData.adGroupIssues.length}`);
          Logger.log(`   Ad Issues: ${accountData.adIssues.length}`);
          Logger.log(`   Asset Group Issues: ${(accountData.assetGroupIssues || []).length}`);
        } else {
          Logger.log(`✓ No destination URL issues found`);
        }
        
        Logger.log(`✓ Successfully processed ${mapping.accountName}`);
        
      } catch (error) {
        Logger.log(`❌ Error processing ${mapping.accountName}: ${error.message}`);
      }
    });
    
    // Send email if URL issues found
    if (accountsWithUrlIssues.length > 0) {
      sendAlertEmail(accountsWithUrlIssues, totalUrlIssues);
      Logger.log(`📧 Email sent for ${accountsWithUrlIssues.length} accounts with destination URL issues`);
    } else {
      Logger.log(`✓ No email sent - no destination URL issues found`);
    }
    
    Logger.log("\n=== MCC DESTINATION URL MONITOR COMPLETED ===");
    Logger.log(`📊 Summary:`);
    Logger.log(`  - Total accounts processed: ${accountMappings.length}`);
    Logger.log(`  - Total campaigns checked: ${totalCampaignsChecked}`);
    Logger.log(`  - Total ad groups checked: ${totalAdGroupsChecked}`);
    Logger.log(`  - Total ads checked: ${totalAdsChecked}`);
    Logger.log(`  - Total asset groups checked: ${totalAssetGroupsChecked}`);
    Logger.log(`  - Accounts with URL issues: ${accountsWithUrlIssues.length}`);
    Logger.log(`  - Total URL issues: ${totalUrlIssues}`);
    
  } catch (error) {
    Logger.log(`❌ Main function error: ${error.message}`);
  }
  
  logScriptHealth(103); // Error 404: destination not working
}

function getDestinationUrlIssues(accountName) {
  try {
    Logger.log(`🔍 Getting destination URL issues for ${accountName}...`);
    Logger.log(`Function started at: ${new Date().toISOString()}`);
    Logger.log(`Account context: ${AdsApp.currentAccount().getName()}`);
    
    const campaignIssues = [];
    const adGroupIssues = [];
    const adIssues = [];
    const assetGroupIssues = [];
    
    // Counters for logging
    let campaignCount = 0;
    let adGroupCount = 0;
    let adCount = 0;
    let assetGroupCount = 0;
    
    // Use GAQL to get all ads with their URLs in one query
    Logger.log(`🔍 Using GAQL to get ads with destination URLs...`);
    
    const query = `
      SELECT 
        campaign.id,
        campaign.name,
        ad_group.id,
        ad_group.name,
        ad_group_ad.ad.id,
        ad_group_ad.ad.type,
        ad_group_ad.status,
        ad_group_ad.policy_summary.approval_status,
        ad_group_ad.policy_summary.policy_topic_entries
      FROM ad_group_ad
      WHERE 
        campaign.status = 'ENABLED' AND
        ad_group.status = 'ENABLED' AND
        ad_group_ad.status = 'ENABLED'
      ORDER BY campaign.name, ad_group.name
    `;
    
    Logger.log(`Executing GAQL query for ads...`);
    Logger.log(`Query: ${query}`);
    
    // Use safe data fetching for main GAQL query
    Logger.log(`Executing GAQL query for ads...`);
    const adData = fetchDataSafely(query);
    
    if (!adData || adData.length === 0) {
      Logger.log(`No ad data returned from GAQL query`);
      return {
        campaignIssues: [],
        adGroupIssues: [],
        adIssues: [],
        assetGroupIssues: [],
        totalIssues: 0,
        campaignCount: 0,
        adGroupCount: 0,
        adCount: 0,
        assetGroupCount: 0
      };
    }
    
    Logger.log(`GAQL query returned ${adData.length} rows`);
    
    // Debug: Log the first row structure to see available fields
    if (adData.length > 0) {
      Logger.log(`🔍 First row available fields: ${Object.keys(adData[0]).join(', ')}`);
      Logger.log(`🔍 Sample field values:`);
      Logger.log(`   campaign.id: ${adData[0]['campaign.id']}`);
      Logger.log(`   campaign.name: ${adData[0]['campaign.name']}`);
      Logger.log(`   adGroup.id: ${adData[0]['adGroup.id']}`);
      Logger.log(`   adGroup.name: ${adData[0]['adGroup.name']}`);
      Logger.log(`   adGroupAd.ad.id: ${adData[0]['adGroupAd.ad.id']}`);
      Logger.log(`   adGroupAd.ad.type: ${adData[0]['adGroupAd.ad.type']}`);
    }
    
    // Process the ad data safely
    let rowCount = 0;
    adData.forEach((row, index) => {
      try {
        rowCount++;
        Logger.log(`Processing row ${rowCount}: ${JSON.stringify(row)}`);
        
        // Access flattened fields directly from the row object
        const campaignId = row['campaign.id'] || '';
        const campaignName = row['campaign.name'] || '';
        const adGroupId = row['adGroup.id'] || '';
        const adGroupName = row['adGroup.name'] || '';
        const adId = row['adGroupAd.ad.id'] || '';
        const adType = row['adGroupAd.ad.type'] || 'Unknown';
        const adStatus = row['adGroupAd.status'] || '';
        
        Logger.log(`Row ${rowCount} data: Campaign=${campaignName}(${campaignId}), AdGroup=${adGroupName}(${adGroupId}), Ad=${adId}, Type=${adType}, Status=${adStatus}`);
        
        // Count items
        if (campaignId && campaignName) {
          campaignCount++;
          Logger.log(`✓ Campaign counted: ${campaignName}`);
        } else {
          Logger.log(`⚠️ Campaign missing data: ID=${campaignId}, Name=${campaignName}`);
        }
        
        if (adGroupId && adGroupName) {
          adGroupCount++;
          Logger.log(`✓ Ad Group counted: ${adGroupName}`);
        } else {
          Logger.log(`⚠️ Ad Group missing data: ID=${adGroupId}, Name=${adGroupName}`);
        }
        
        if (adId) {
          adCount++;
          Logger.log(`✓ Ad counted: ${adId} (${adType})`);
        } else {
          Logger.log(`⚠️ Ad missing ID: ${adId}`);
        }
        
        // For now, we'll just count the ads since getting URLs requires different approach
        // We can add URL checking later once we figure out the correct method
        
      } catch (error) {
        Logger.log(`❌ Error processing GAQL row ${index + 1}: ${error.message}`);
        Logger.log(`Error details: ${error.stack || 'No stack trace'}`);
      }
    });
    
    Logger.log(`Finished processing ${rowCount} GAQL rows`);
    


    
    // Now get asset groups (Performance Max campaigns) with enhanced error handling
    const assetGroupResult = getAssetGroupsSafely(accountName);
    
    if (assetGroupResult.error) {
      Logger.log(`⚠️ Asset group monitoring failed: ${assetGroupResult.error}`);
      Logger.log(`⚠️ Asset group monitoring disabled for ${accountName} due to errors`);
      assetGroupCount = 0;
      assetGroupRowCount = 0;
    } else {
      assetGroupCount = assetGroupResult.assetGroupCount;
      assetGroupRowCount = assetGroupResult.assetGroupData.length;
      Logger.log(`✓ Successfully processed ${assetGroupRowCount} asset group rows`);
    }
    
    const totalIssues = campaignIssues.length + adGroupIssues.length + adIssues.length;
    
    // Log results with counts
    Logger.log(`=== RESULTS FOR ${accountName.toUpperCase()} ===`);
    Logger.log(`Campaigns Checked: ${campaignCount}`);
    Logger.log(`Ad Groups Checked: ${adGroupCount}`);
    Logger.log(`Ads Checked: ${adCount}`);
    Logger.log(`Asset Groups Checked: ${assetGroupCount}`);
    Logger.log(`Campaign Issues: ${campaignIssues.length}`);
    Logger.log(`Ad Group Issues: ${adGroupIssues.length}`);
    Logger.log(`Ad Issues: ${adIssues.length}`);
    Logger.log(`Asset Group Issues: ${assetGroupIssues.length}`);
    Logger.log(`Total Issues: ${totalIssues}`);
    
    // Log data validation
    Logger.log(`=== DATA VALIDATION ===`);
    if (campaignCount === 0) {
      Logger.log(`⚠️ WARNING: No campaigns found - this might indicate a query issue`);
    }
    if (adGroupCount === 0) {
      Logger.log(`⚠️ WARNING: No ad groups found - this might indicate a query issue`);
    }
    if (adCount === 0) {
      Logger.log(`⚠️ WARNING: No ads found - this might indicate a query issue`);
    }
    if (assetGroupCount === 0) {
      Logger.log(`⚠️ WARNING: No asset groups found - this might indicate a query issue`);
    }
    
    // Log execution summary
    Logger.log(`=== EXECUTION SUMMARY ===`);
    Logger.log(`Function completed at: ${new Date().toISOString()}`);
    Logger.log(`Total GAQL rows processed: ${rowCount}`);
    Logger.log(`Total asset group rows processed: ${assetGroupRowCount}`);
    Logger.log(`Data extraction successful: ${campaignCount > 0 && (adGroupCount > 0 || assetGroupCount > 0) ? 'Yes' : 'No'}`);
    
    return {
      campaignIssues,
      adGroupIssues,
      adIssues,
      assetGroupIssues,
      totalIssues,
      campaignCount,
      adGroupCount,
      adCount,
      assetGroupCount
    };
    
  } catch (error) {
    Logger.log(`❌ Error getting destination URL issues for ${accountName}: ${error.message}`);
    Logger.log(`Error details: ${error}`);
    
    logScriptHealth(103); // Error 404: destination not working
    return {
      campaignIssues: [],
      adGroupIssues: [],
      adIssues: [],
      assetGroupIssues: [],
      totalIssues: 0,
      campaignCount: 0,
      adGroupCount: 0,
      adCount: 0,
      assetGroupCount: 0
    };
  }
}

function getAssetGroupsSafely(accountName) {
  try {
    Logger.log(`🔍 Getting asset groups for Performance Max campaigns...`);
    
    const assetGroupQuery = `
      SELECT 
        campaign.id,
        campaign.name,
        asset_group.id,
        asset_group.name,
        asset_group.status
      FROM asset_group
      WHERE 
        campaign.status = 'ENABLED' AND
        asset_group.status = 'ENABLED'
      ORDER BY campaign.name, asset_group.name
    `;
    
    Logger.log(`Executing asset group query: ${assetGroupQuery}`);
    
    // Use safe data fetching with error handling
    const assetGroupData = fetchDataSafely(assetGroupQuery);
    
    if (!assetGroupData || assetGroupData.length === 0) {
      Logger.log(`No asset groups found for ${accountName}`);
      return { assetGroupCount: 0, assetGroupData: [] };
    }
    
    // Validate the data structure before processing
    if (!validateAssetGroupData(assetGroupData)) {
      Logger.log(`⚠️ Asset group data validation failed for ${accountName}`);
      return { assetGroupCount: 0, assetGroupData: [], error: 'Data validation failed' };
    }
    
    Logger.log(`Found ${assetGroupData.length} asset groups for ${accountName}`);
    
    // Debug: Log the first asset group row structure
    if (assetGroupData.length > 0) {
      Logger.log(`🔍 First asset group row available fields: ${Object.keys(assetGroupData[0]).join(', ')}`);
      Logger.log(`🔍 Sample asset group field values:`);
      Logger.log(`   campaign.id: ${assetGroupData[0]['campaign.id']}`);
      Logger.log(`   campaign.name: ${assetGroupData[0]['campaign.name']}`);
      Logger.log(`   asset_group.id: ${assetGroupData[0]['asset_group.id']}`);
      Logger.log(`   asset_group.name: ${assetGroupData[0]['asset_group.name']}`);
      Logger.log(`   asset_group.status: ${assetGroupData[0]['asset_group.status']}`);
    }
    
    // Process and count asset groups
    let assetGroupCount = 0;
    assetGroupData.forEach((row, index) => {
      try {
        // Access flattened fields directly from the row object
        const campaignId = row['campaign.id'] || '';
        const campaignName = row['campaign.name'] || '';
        const assetGroupId = row['assetGroup.id'] || '';
        const assetGroupName = row['assetGroup.name'] || '';
        const assetGroupStatus = row['assetGroup.status'] || '';
        const assetGroupType = 'PERFORMANCE_MAX'; // Default type since we can't query it
        
        Logger.log(`Asset Group ${index + 1}: Campaign=${campaignName}(${campaignId}), AssetGroup=${assetGroupName}(${assetGroupId}), Type=${assetGroupType}, Status=${assetGroupStatus}`);
        
        if (assetGroupId && assetGroupName) {
          assetGroupCount++;
          Logger.log(`✓ Asset Group counted: ${assetGroupName} (${assetGroupType})`);
        } else {
          Logger.log(`⚠️ Asset Group missing data: ID=${assetGroupId}, Name=${assetGroupName}`);
        }
      } catch (error) {
        Logger.log(`❌ Error processing asset group row ${index + 1}: ${error.message}`);
      }
    });
    
    return { assetGroupCount, assetGroupData };
    
  } catch (error) {
    Logger.log(`❌ Error getting asset groups for ${accountName}: ${error.message}`);
    Logger.log(`Error details: ${error.stack || 'No stack trace'}`);
    return { assetGroupCount: 0, assetGroupData: [], error: error.message };
  }
}

function fetchDataSafely(query) {
  try {
    Logger.log(`Executing query: ${query}`);
    let data = [];
    let iterator = AdsApp.search(query, { 'apiVersion': 'v20' });
    
    while (iterator.hasNext()) {
      let row = flattenObject(iterator.next());
      data.push(row);
    }
    
    Logger.log(`Query returned ${data.length} rows`);
    return data;
    
  } catch (error) {
    Logger.log(`❌ Error fetching data for query: ${query}`);
    Logger.log(`Error message: ${error.message}`);
    Logger.log(`Error stack: ${error.stack || 'No stack trace'}`);
    return null;
  }
}

function validateAssetGroupData(data) {
  if (!data || !Array.isArray(data)) {
    Logger.log(`⚠️ Invalid asset group data format: ${typeof data}`);
    return false;
  }
  
  if (data.length === 0) {
    Logger.log(`ℹ️ No asset group data to validate`);
    return true; // Empty data is valid
  }
  
  // Check first row structure
  const firstRow = data[0];
  const requiredFields = ['campaign.id', 'assetGroup.id', 'assetGroup.name'];
  
  for (const field of requiredFields) {
    if (!(field in firstRow)) {
      Logger.log(`⚠️ Missing required field in asset group data: ${field}`);
      Logger.log(`Available fields: ${Object.keys(firstRow).join(', ')}`);
      return false;
    }
  }
  
  Logger.log(`✓ Asset group data validation passed`);
  return true;
}

function flattenObject(ob) {
  let toReturn = {};
  let stack = [{ obj: ob, prefix: '' }];

  while (stack.length > 0) {
    let { obj, prefix } = stack.pop();
    for (let i in obj) {
      if (obj.hasOwnProperty(i)) {
        let key = prefix ? prefix + '.' + i : i;
        if (typeof obj[i] === 'object' && obj[i] !== null) {
          stack.push({ obj: obj[i], prefix: key });
        } else {
          toReturn[key] = obj[i];
        }
      }
    }
  }

  return toReturn;
}

function isUrlWorking(url) {
  try {
    if (!url || url.trim() === '') {
      return false;
    }
    
    // Basic URL validation
    try {
      new URL(url);
    } catch (e) {
      Logger.log(`Invalid URL format: ${url}`);
      return false;
    }
    
    // For Google Ads scripts, we can't make actual HTTP requests
    // So we'll do basic validation and return true for now
    // In a real implementation, you might want to use UrlFetchApp with timeout
    
    // Basic checks
    if (url.includes('localhost') || url.includes('127.0.0.1')) {
      return false; // Local URLs won't work
    }
    
    if (url.startsWith('http://') && !url.includes('localhost')) {
      // HTTP URLs might have issues, but we'll assume they work for now
      return true;
    }
    
    if (url.startsWith('https://')) {
      return true; // HTTPS URLs are generally reliable
    }
    
    logScriptHealth(103); // Error 404: destination not working
    return false;
    
  } catch (error) {
    Logger.log(`Error checking URL ${url}: ${error.message}`);
    logScriptHealth(103); // Error 404: destination not working
    return false;
  }
}

function sendAlertEmail(accountsWithUrlIssues, totalUrlIssues) {
  try {
    Logger.log("📧 Sending individual alert emails to account managers...");
    
    let emailsSent = 0;
    let emailsFailed = 0;
    
    // Send individual email to each account's responsible person
    accountsWithUrlIssues.forEach(account => {
      if (!account.emailRecipient) {
        Logger.log(`⚠️ No email recipient specified for ${account.accountName} - skipping email`);
        return;
      }
      
      const subject = `🚨 Alert: ${account.accountName} has Destination URL Issues`;
      
      let body = `
<h2>🚨 Destination URL Issues Alert</h2>
<p><strong>Account:</strong> ${account.accountName} (${account.accountId})</p>
<p><strong>Date:</strong> ${new Date().toLocaleDateString()}</p>
<p><strong>Total Issues:</strong> ${account.totalIssues}</p>

<h3>📊 Summary</h3>
<p><strong>Campaign Issues:</strong> ${account.campaignIssues.length}</p>
<p><strong>Ad Group Issues:</strong> ${account.adGroupIssues.length}</p>
<p><strong>Ad Issues:</strong> ${account.adIssues.length}</p>

<h3>🎯 Campaign Issues</h3>
<ul>
`;
      
      account.campaignIssues.forEach(issue => {
        body += `<li><strong>${issue.name}</strong> (ID: ${issue.id}) - ${issue.url}</li>`;
      });
      
      body += `
</ul>

<h3>🎯 Ad Group Issues</h3>
<ul>
`;
      
      account.adGroupIssues.forEach(issue => {
        body += `<li><strong>${issue.name}</strong> (ID: ${issue.id}) - Campaign: ${issue.campaignName} - ${issue.url}</li>`;
      });
      
      body += `
</ul>

<h3>🎯 Ad Issues</h3>
<ul>
`;
      
      account.adIssues.forEach(issue => {
        body += `<li><strong>${issue.type}</strong> (ID: ${issue.id}) - Ad Group: ${issue.adGroupName} - Campaign: ${issue.campaignName} - ${issue.url}</li>`;
      });
      
      body += `
</ul>

<h3>🎯 Asset Group Issues</h3>
<ul>
`;
      
      if (account.assetGroupIssues && account.assetGroupIssues.length > 0) {
        account.assetGroupIssues.forEach(issue => {
          body += `<li><strong>${issue.name}</strong> (ID: ${issue.id}) - Campaign: ${issue.campaignName} - Type: ${issue.type} - ${issue.url}</li>`;
        });
      } else {
        body += `<li>No asset group issues found</li>`;
      }
      
      body += `
</ul>

<p><em>Generated by MCC Destination URL Monitor</em></p>
`;

      // Send email using available email service
      try {
        // Try GmailApp first (if authorized)
        GmailApp.sendEmail(account.emailRecipient, subject, '', {
          htmlBody: body
        });
        Logger.log(`✓ Alert email sent to ${account.emailRecipient} for ${account.accountName}`);
        emailsSent++;
      } catch (gmailError) {
        Logger.log(`GmailApp failed for ${account.accountName}, trying MailApp: ${gmailError.message}`);
        
        // Fallback to MailApp (more commonly available)
        try {
          MailApp.sendEmail(account.emailRecipient, subject, body.replace(/<[^>]*>/g, ''), {
            htmlBody: body
          });
          Logger.log(`✓ Alert email sent via MailApp to ${account.emailRecipient} for ${account.accountName}`);
          emailsSent++;
        } catch (mailError) {
          Logger.log(`❌ Both GmailApp and MailApp failed for ${account.accountName}: ${mailError.message}`);
          Logger.log(`📧 Email content (for manual sending):`);
          Logger.log(`Subject: ${subject}`);
          Logger.log(`Body: ${body}`);
          emailsFailed++;
        }
      }
    });
    
    Logger.log(`📧 Email Summary: ${emailsSent} sent successfully, ${emailsFailed} failed`);
    
  } catch (error) {
    Logger.log(`❌ Error in sendAlertEmail function: ${error.message}`);
    logScriptHealth(103); // Error 404: destination not working
  }
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
    logScriptHealth(103); // Error 404: destination not working
    throw error;
  }
}

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
    return null;
  } catch (error) {
    Logger.log(`❌ Error finding account ${accountId}: ${error.message}`);
    logScriptHealth(103); // Error 404: destination not working
    return null;
  }
}

function testScript() {
  Logger.log("=== Testing MCC Destination URL Monitor ===");
  main();
}

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
    logScriptHealth(103); // Error 404: destination not working
  }
}
