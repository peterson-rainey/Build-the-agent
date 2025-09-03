// --- COUNTRY IMPRESSIONS EMAIL MONITOR ---
// Configuration
const TARGET_COUNTRIES = ["India", "Philippines", "Pakistan", "Brazil", "Saudi Arabia", "Jamaica", "Puerto Rico", "Argentina", "Cambodia", "Costa Rica", "El Salvador", "Nigeria", "Thailand", "Afghanistan", "Albania", "Algeria", "Bangladesh", "Belarus", "China", "Colombia", "Dominican Republic", "Ecuador", "Ethiopia", "Egypt", "Honduras", "Indonesia"];
const CHECK_YESTERDAY_ONLY = true;
const MASTER_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1bS1xmTFB0LWGvYeZBPcykGMY0SwZ4W-Ny6aKVloy0SM/edit?usp=sharing";
const MASTER_SHEET_NAME = "AccountMappings";

// Enable Gmail service for email notifications
// Note: You may need to authorize Gmail access in the script settings

function main() {
  try {
    Logger.log("=== Starting MCC Country Impressions Email Monitor ===");
    
    // Get account mappings
    const accountMappings = getAccountMappings();
    if (accountMappings.length === 0) {
      Logger.log("❌ No account mappings found. Exiting.");
      logScriptHealth(102); // Country Checker
      return;
    }
    
    Logger.log(`📋 Found ${accountMappings.length} account mappings.`);
    
    // Process each account
    const accountsWithTargetCountries = [];
    let totalTargetCountryImpressions = 0;
    
    Logger.log("🔍 Processing each account for country impressions...");
    
    accountMappings.forEach((mapping, index) => {
      Logger.log(`\n--- Processing Account ${index + 1}/${accountMappings.length} ---`);
      Logger.log(`Account: ${mapping.accountName} (${mapping.accountId})`);
      
      try {
        // Find and switch to account
        const account = getAccountById(mapping.accountId);
        if (!account) {
          Logger.log(`❌ Account not found: ${mapping.accountName}`);
          return;
        }
        
        Logger.log(`✓ Account found: ${account.getName()}`);
        
        // Switch to account context
        AdsManagerApp.select(account);
        Logger.log(`✓ Switched to account: ${account.getName()}`);
        
        // Get country impressions with campaign details
        const accountData = getCountryImpressionsWithCampaigns(mapping.accountName);
        
        if (accountData.targetCountryImpressions > 0) {
          accountsWithTargetCountries.push({
            accountName: mapping.accountName,
            accountId: mapping.accountId,
            emailRecipient: mapping.emailRecipient,
            targetImpressions: accountData.targetCountryImpressions,
            targetCountries: accountData.targetCountries,
            campaignDetails: accountData.campaignDetails
          });
          totalTargetCountryImpressions += accountData.targetCountryImpressions;
          
          Logger.log(`⚠️ TARGET COUNTRIES WITH IMPRESSIONS:`);
          accountData.targetCountries.forEach(country => {
            Logger.log(`   ${country}: ${accountData.countryData[country]}`);
          });
        } else {
          Logger.log(`✓ No target country impressions found`);
        }
        
        Logger.log(`✓ Successfully processed ${mapping.accountName}`);
        
      } catch (error) {
        Logger.log(`❌ Error processing ${mapping.accountName}: ${error.message}`);
      }
    });
    
    // Send email if target countries found
    if (accountsWithTargetCountries.length > 0) {
      sendAlertEmail(accountsWithTargetCountries, totalTargetCountryImpressions);
      Logger.log(`📧 Email sent for ${accountsWithTargetCountries.length} accounts with target country impressions`);
    } else {
      Logger.log(`�� No email sent - no target country impressions found`);
    }
    
    Logger.log("\n=== MCC COUNTRY IMPRESSIONS EMAIL MONITOR COMPLETED ===");
    Logger.log(`📊 Summary:`);
    Logger.log(`  - Total accounts processed: ${accountMappings.length}`);
    Logger.log(`  - Accounts with target country impressions: ${accountsWithTargetCountries.length}`);
    Logger.log(`  - Total target country impressions: ${totalTargetCountryImpressions}`);
    
  } catch (error) {
    Logger.log(`❌ Main function error: ${error.message}`);
  }
  
  logScriptHealth(102); // Country Checker
}

function getCountryImpressionsWithCampaigns(accountName) {
  try {
    Logger.log(`🔍 Getting country impressions with campaign details for ${accountName}...`);
    
    // Set date range
    const dateRange = CHECK_YESTERDAY_ONLY 
      ? `segments.date = "${getYesterdayDate()}"`
      : `segments.date DURING LAST_30_DAYS`;
    
    // Query to get impressions by location IDs with campaign info
    const query = `
      SELECT 
          user_location_view.country_criterion_id,
          campaign.id,
          campaign.name,
          metrics.impressions
      FROM user_location_view
      WHERE 
          ${dateRange}
      ORDER BY metrics.impressions DESC
    `;
    
    Logger.log(`Executing query to get country impressions with campaigns...`);
    
    // First, log sample row structure for debugging
    const sampleQuery = query + ' LIMIT 1';
    const sampleRows = AdsApp.search(sampleQuery);
    
    if (sampleRows.hasNext()) {
      const sampleRow = sampleRows.next();
      Logger.log("Sample row structure: " + JSON.stringify(sampleRow));
      if (sampleRow.userLocationView) {
        Logger.log("Sample userLocationView object: " + JSON.stringify(sampleRow.userLocationView));
      }
      if (sampleRow.campaign) {
        Logger.log("Sample campaign object: " + JSON.stringify(sampleRow.campaign));
      }
      if (sampleRow.metrics) {
        Logger.log("Sample metrics object: " + JSON.stringify(sampleRow.metrics));
      }
    } else {
      Logger.log("No sample rows found - query may return no results");
    }
    
    // Execute the main query
    const rows = AdsApp.search(query);
    const locationIdData = {};
    const campaignDetails = {};
    let totalImpressions = 0;
    
    // Process results to get location IDs and impressions with campaign info
    while (rows.hasNext()) {
      try {
        const row = rows.next();
        const userLocationView = row.userLocationView || {};
        const campaign = row.campaign || {};
        const metrics = row.metrics || {};
        
        const locationId = userLocationView.countryCriterionId || 'Unknown';
        const campaignId = campaign.id || 'Unknown';
        const campaignName = campaign.name || 'Unknown Campaign';
        const impressions = Number(metrics.impressions) || 0;
        
        if (locationId && campaignId) {
          // Add to location ID data
          if (!locationIdData[locationId]) {
            locationIdData[locationId] = 0;
          }
          locationIdData[locationId] += impressions;
          totalImpressions += impressions;
          
          // Store campaign details
          if (!campaignDetails[locationId]) {
            campaignDetails[locationId] = {};
          }
          if (!campaignDetails[locationId][campaignId]) {
            campaignDetails[locationId][campaignId] = {
              name: campaignName,
              impressions: 0
            };
          }
          campaignDetails[locationId][campaignId].impressions += impressions;
        }
        
      } catch (error) {
        Logger.log(`Error processing row: ${error.message}`);
      }
    }
    
    // Get location names from IDs
    const locationNames = getLocationNamesFromIds(Object.keys(locationIdData));
    
    // Convert ID data to name data
    const countryData = {};
    const targetCountries = [];
    let targetCountryImpressions = 0;
    
    Object.keys(locationIdData).forEach(locationId => {
      const locationName = locationNames[locationId] || `Unknown (ID: ${locationId})`;
      const impressions = locationIdData[locationId];
      
      countryData[locationName] = impressions;
      
      // Only flag target countries if they have actual impressions
      if (TARGET_COUNTRIES.includes(locationName) && impressions > 0) {
        targetCountryImpressions += impressions;
        if (!targetCountries.includes(locationName)) {
          targetCountries.push(locationName);
        }
      }
    });
    
    // Convert campaign details to use location names
    const campaignDetailsByName = {};
    Object.keys(campaignDetails).forEach(locationId => {
      const locationName = locationNames[locationId] || `Unknown (ID: ${locationId})`;
      if (TARGET_COUNTRIES.includes(locationName)) {
        campaignDetailsByName[locationName] = campaignDetails[locationId];
      }
    });
    
    // Log results
    Logger.log(`=== RESULTS FOR ${accountName.toUpperCase()} ===`);
    Logger.log(`Total Countries: ${Object.keys(countryData).length}`);
    Logger.log(`Total Impressions: ${totalImpressions.toLocaleString()}`);
    Logger.log(`Target Country Impressions: ${targetCountryImpressions.toLocaleString()}`);
    
    if (targetCountryImpressions > 0) {
      const percentage = ((targetCountryImpressions / totalImpressions) * 100).toFixed(2);
      Logger.log(`Target Country %: ${percentage}%`);
    }
    
    return {
      countryData,
      totalImpressions,
      targetCountryImpressions,
      targetCountries,
      campaignDetails: campaignDetailsByName
    };
    
  } catch (error) {
    Logger.log(`❌ Error getting country impressions for ${accountName}: ${error.message}`);
    Logger.log(`Error details: ${error}`);
    
    return {
      countryData: {},
      totalImpressions: 0,
      targetCountryImpressions: 0,
      targetCountries: [],
      campaignDetails: {}
    };
  }
}

function sendAlertEmail(accountsWithTargetCountries, totalTargetCountryImpressions) {
  try {
    Logger.log("📧 Sending individual alert emails to account managers...");
    
    let emailsSent = 0;
    let emailsFailed = 0;
    
    // Send individual email to each account's responsible person
    accountsWithTargetCountries.forEach(account => {
      if (!account.emailRecipient) {
        Logger.log(`⚠️ No email recipient specified for ${account.accountName} - skipping email`);
        return;
      }
      
      const subject = `🚨 Alert: ${account.accountName} has Target Country Impressions`;
      
      let body = `
<h2>🚨 Country Impressions Alert</h2>
<p><strong>Account:</strong> ${account.accountName} (${account.accountId})</p>
<p><strong>Date:</strong> ${new Date().toLocaleDateString()}</p>
<p><strong>Date Range:</strong> ${CHECK_YESTERDAY_ONLY ? 'Yesterday Only' : 'Last 30 Days'}</p>

<h3>📊 Summary</h3>
<p><strong>Total Target Impressions:</strong> ${account.targetImpressions.toLocaleString()}</p>
<p><strong>Problem Countries:</strong> ${account.targetCountries.join(', ')}</p>

<h3>🎯 Campaigns with Issues</h3>
<ul>
`;

      account.targetCountries.forEach(country => {
        if (account.campaignDetails[country]) {
          Object.keys(account.campaignDetails[country]).forEach(campaignId => {
            const campaign = account.campaignDetails[country][campaignId];
            body += `<li><strong>${campaign.name}</strong> - ${country}: ${campaign.impressions.toLocaleString()} impressions</li>`;
          });
        }
      });
      
      body += `
</ul>

<p><em>Generated by MCC Country Impressions Monitor</em></p>
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
    return null;
  }
}

function getLocationNamesFromIds(locationIds) {
  try {
    const locationNames = {};
    
    if (locationIds.length === 0) {
      return locationNames;
    }
    
    // Query geo_target_constant to get names from IDs
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
      
      if (id && name) {
        locationNames[id] = name;
        Logger.log(`Mapped ID ${id} to name: ${name}`);
      }
    }
    
    Logger.log(`Successfully mapped ${Object.keys(locationNames).length} location IDs to names`);
    return locationNames;
    
  } catch (error) {
    Logger.log(`Error getting location names from IDs: ${error.message}`);
    return {};
  }
}

function getYesterdayDate() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  return Utilities.formatDate(yesterday, AdsApp.currentAccount().getTimeZone(), 'yyyyMMdd');
}

function testScript() {
  Logger.log("=== Testing MCC Country Impressions Email Monitor ===");
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
  }
}
