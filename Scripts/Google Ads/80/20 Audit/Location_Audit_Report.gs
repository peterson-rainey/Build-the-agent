// --- LOCATION AUDIT REPORT FOR 80/20 ANALYSIS ---
// This script outputs raw location data by campaign for pivot table analysis
// Shows both targeted locations and actual user locations with performance metrics
// Configure the account ID below to analyze a specific account

const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1ylEHOD31fvz4LAKxCucVD-zFPWxDww0LXBFF6eTRFuM/edit?usp=sharing';
const TARGETED_LOCATIONS_TAB = 'TargetedLocations';
const USER_LOCATIONS_TAB = 'UserLocations';
const ACCOUNT_ID = '972-837-2864'; // Target account ID for testing

// Query for user locations (where ads actually showed) - Country level only, enabled campaigns only
const USER_LOCATIONS_QUERY = `
  SELECT 
    user_location_view.country_criterion_id,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM user_location_view
  WHERE segments.date DURING LAST_30_DAYS
    AND campaign.status = 'ENABLED'
  ORDER BY metrics.impressions DESC
`;

// Query for targeted locations (campaign location targets) - Country level only
// Note: No date segments needed for campaign criteria as they are static settings
const TARGETED_LOCATIONS_QUERY = `
  SELECT 
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign_criterion.location.geo_target_constant,
    campaign_criterion.status
  FROM campaign_criterion
  WHERE 
    campaign_criterion.type = 'LOCATION'
    AND campaign.status = 'ENABLED'
    AND campaign_criterion.status = 'ENABLED'
  ORDER BY campaign.name ASC
`;

function main() {
  try {
    Logger.log("=== Starting Location Audit Report for 80/20 Analysis ===");
    
    // Validate account ID
    if (!ACCOUNT_ID || ACCOUNT_ID === '123-456-7890') {
      Logger.log("❌ Please set the ACCOUNT_ID constant to your target account ID");
      Logger.log("Example: const ACCOUNT_ID = '123-456-7890';");
      return;
    }
    
    // Switch to the specified account
    Logger.log(`🔍 Switching to account: ${ACCOUNT_ID}`);
    const account = getAccountById(ACCOUNT_ID);
    if (!account) {
      Logger.log(`❌ Account ${ACCOUNT_ID} not found or not accessible`);
      return;
    }
    
    // Switch to account context
    AdsManagerApp.select(account);
    Logger.log(`✓ Successfully switched to account: ${account.getName()}`);
    
    // Open or create spreadsheet
    let ss;
    if (!SHEET_URL) {
      ss = SpreadsheetApp.create("Location Audit Report");
      let url = ss.getUrl();
      Logger.log("No SHEET_URL found, so this sheet was created: " + url);
    } else {
      ss = SpreadsheetApp.openByUrl(SHEET_URL);
    }
    
    // Process User Locations (where ads actually showed)
    Logger.log("=== Processing User Locations (where ads showed) ===");
    processUserLocations(ss, account);
    
    // Process Targeted Locations (campaign location targets)
    Logger.log("=== Processing Targeted Locations (campaign targets) ===");
    processTargetedLocations(ss, account);
    
    Logger.log(`✓ Account analyzed: ${account.getName()} (${ACCOUNT_ID})`);
    Logger.log(`✓ Data exported to two tabs: ${USER_LOCATIONS_TAB} and ${TARGETED_LOCATIONS_TAB}`);
    
  } catch (error) {
    Logger.log(`❌ Error in main function: ${error.message}`);
    Logger.log(`Error details: ${error.stack || 'No stack trace available'}`);
  }
}

function processUserLocations(ss, account) {
  try {
    // Get or create the user locations tab
    let sheet = ss.getSheetByName(USER_LOCATIONS_TAB);
    if (!sheet) {
      sheet = ss.insertSheet(USER_LOCATIONS_TAB);
      Logger.log(`Created new tab: ${USER_LOCATIONS_TAB}`);
    } else {
      sheet.clear();
      Logger.log(`Cleared existing data in tab: ${USER_LOCATIONS_TAB}`);
    }
    
    // Execute user locations query
    Logger.log("Executing user locations query...");
    const rows = AdsApp.search(USER_LOCATIONS_QUERY);
    
    // Log sample row structure for debugging
    const sampleQuery = USER_LOCATIONS_QUERY + ' LIMIT 1';
    const sampleRows = AdsApp.search(sampleQuery);
    
    if (sampleRows.hasNext()) {
      const sampleRow = sampleRows.next();
      Logger.log("Sample user location row structure: " + JSON.stringify(sampleRow));
    } else {
      Logger.log("No sample rows found for user locations");
    }
    
    // Process the data
    const data = calculateUserLocationMetrics(rows);
    
    if (data.length === 0) {
      Logger.log("No user location data found to export");
      return;
    }
    
    // Create headers for user locations
    const headers = [
      'Campaign ID',
      'Campaign Name', 
      'Channel Type',
      'Campaign Status',
      'Location ID',
      'Location Name',
      'Impressions',
      'Clicks',
      'Cost',
      'Conversions',
      'Conversion Value',
      'CPC',
      'CTR',
      'Conv Rate',
      'CPA',
      'ROAS',
      'AOV'
    ];
    
    // Write headers and data to sheet
    const allData = [headers, ...data];
    const range = sheet.getRange(1, 1, allData.length, headers.length);
    range.setValues(allData);
    
    Logger.log(`✓ Successfully exported ${data.length} user location records to ${USER_LOCATIONS_TAB} tab`);
    
  } catch (error) {
    Logger.log(`❌ Error processing user locations: ${error.message}`);
  }
}

function processTargetedLocations(ss, account) {
  try {
    // Get or create the targeted locations tab
    let sheet = ss.getSheetByName(TARGETED_LOCATIONS_TAB);
    if (!sheet) {
      sheet = ss.insertSheet(TARGETED_LOCATIONS_TAB);
      Logger.log(`Created new tab: ${TARGETED_LOCATIONS_TAB}`);
    } else {
      sheet.clear();
      Logger.log(`Cleared existing data in tab: ${TARGETED_LOCATIONS_TAB}`);
    }
    
    // First, get campaigns with impressions to filter targeted locations
    Logger.log("Getting campaigns with impressions for filtering...");
    const campaignsWithImpressions = getCampaignsWithImpressions();
    
    // Execute targeted locations query
    Logger.log("Executing targeted locations query...");
    const rows = AdsApp.search(TARGETED_LOCATIONS_QUERY);
    
    // Log sample row structure for debugging
    const sampleQuery = TARGETED_LOCATIONS_QUERY + ' LIMIT 1';
    const sampleRows = AdsApp.search(sampleQuery);
    
    if (sampleRows.hasNext()) {
      const sampleRow = sampleRows.next();
      Logger.log("Sample targeted location row structure: " + JSON.stringify(sampleRow));
    } else {
      Logger.log("No sample rows found for targeted locations");
    }
    
    // Process the data, filtering out campaigns with zero impressions
    const data = processTargetedLocationData(rows, campaignsWithImpressions);
    
    if (data.length === 0) {
      Logger.log("No targeted location data found to export");
      return;
    }
    
    // Create headers for targeted locations
    const headers = [
      'Campaign ID',
      'Campaign Name',
      'Channel Type',
      'Location ID',
      'Location Name',
      'Status'
    ];
    
    // Write headers and data to sheet
    const allData = [headers, ...data];
    const range = sheet.getRange(1, 1, allData.length, headers.length);
    range.setValues(allData);
    
    Logger.log(`✓ Successfully exported ${data.length} targeted location records to ${TARGETED_LOCATIONS_TAB} tab`);
    
  } catch (error) {
    Logger.log(`❌ Error processing targeted locations: ${error.message}`);
  }
}

function calculateUserLocationMetrics(rows) {
  let data = [];
  let locationIdData = {};
  let totalRows = 0;
  
  // First pass: collect all location IDs to get names
  Logger.log("First pass: collecting location IDs...");
  const tempRows = [];
  
  while (rows.hasNext()) {
    try {
      const row = rows.next();
      tempRows.push(row);
      
      const userLocationView = row.userLocationView || {};
      const locationId = userLocationView.countryCriterionId || 'Unknown';
      
      if (locationId && locationId !== 'Unknown') {
        // Extract just the ID number from geoTargetConstants/2840 format if present
        const cleanLocationId = locationId.toString().replace('geoTargetConstants/', '');
        locationIdData[cleanLocationId] = true;
      }
      
    } catch (error) {
      Logger.log(`Error in first pass: ${error.message}`);
    }
  }
  
  // Get location names from IDs
  Logger.log(`Getting location names for ${Object.keys(locationIdData).length} location IDs...`);
  const locationNames = getLocationNamesFromIds(Object.keys(locationIdData));
  
  // Second pass: process data with location names
  Logger.log("Second pass: processing data with location names...");
  
  tempRows.forEach(row => {
    try {
      const userLocationView = row.userLocationView || {};
      const campaign = row.campaign || {};
      const metrics = row.metrics || {};
      
      const locationId = userLocationView.countryCriterionId || 'Unknown';
      const campaignId = campaign.id || 'Unknown';
      const campaignName = campaign.name || 'Unknown Campaign';
      const channelType = campaign.advertisingChannelType || 'Unknown';
      const campaignStatus = campaign.status || 'Unknown';
      
      // Extract just the ID number from geoTargetConstants/2840 format if present
      const cleanLocationId = locationId.toString().replace('geoTargetConstants/', '');
      
      // Get location name
      const locationName = locationNames[cleanLocationId] || `Unknown (ID: ${cleanLocationId})`;
      
      // Access metrics with proper number conversion
      let impressions = Number(metrics.impressions) || 0;
      let clicks = Number(metrics.clicks) || 0;
      let costMicros = Number(metrics.costMicros) || 0;
      let conversions = Number(metrics.conversions) || 0;
      let conversionValue = Number(metrics.conversionsValue) || 0;
      
      // Calculate derived metrics
      let cost = costMicros / 1000000; // Convert micros to actual currency
      let cpc = clicks > 0 ? cost / clicks : 0;
      let ctr = impressions > 0 ? clicks / impressions : 0;
      let convRate = clicks > 0 ? conversions / clicks : 0;
      let cpa = conversions > 0 ? cost / conversions : 0;
      let roas = cost > 0 ? conversionValue / cost : 0;
      let aov = conversions > 0 ? conversionValue / conversions : 0;
      
      // Only include rows with actual impressions (filter out zero data)
      if (impressions > 0) {
        let newRow = [
          campaignId,
          campaignName,
          channelType,
          campaignStatus,
          cleanLocationId,
          locationName,
          impressions,
          clicks,
          cost,
          conversions,
          conversionValue,
          cpc,
          ctr,
          convRate,
          cpa,
          roas,
          aov
        ];
        
        data.push(newRow);
        totalRows++;
      }
      
    } catch (error) {
      Logger.log(`Error processing row: ${error.message}`);
    }
  });
  
  Logger.log(`✓ Processed ${totalRows} rows with impressions > 0`);
  Logger.log(`✓ Total unique locations: ${Object.keys(locationNames).length}`);
  
  return data;
}

function getCampaignsWithImpressions() {
  try {
    Logger.log("Getting list of campaigns with impressions...");
    const campaignsWithImpressions = new Set();
    
    // Use the same query as user locations to get campaigns with impressions
    const query = `
      SELECT 
        campaign.id,
        campaign.status,
        metrics.impressions
      FROM user_location_view
      WHERE segments.date DURING LAST_30_DAYS
        AND campaign.status = 'ENABLED'
        AND metrics.impressions > 0
    `;
    
    const rows = AdsApp.search(query);
    while (rows.hasNext()) {
      const row = rows.next();
      const campaign = row.campaign || {};
      const metrics = row.metrics || {};
      
      const campaignId = campaign.id || 'Unknown';
      const impressions = Number(metrics.impressions) || 0;
      
      if (campaignId && impressions > 0) {
        campaignsWithImpressions.add(campaignId);
      }
    }
    
    Logger.log(`Found ${campaignsWithImpressions.size} campaigns with impressions`);
    return campaignsWithImpressions;
    
  } catch (error) {
    Logger.log(`Error getting campaigns with impressions: ${error.message}`);
    return new Set();
  }
}

function processTargetedLocationData(rows, campaignsWithImpressions) {
  let data = [];
  let locationIdData = {};
  
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
      const channelType = campaign.advertisingChannelType || 'Unknown';
      const locationId = location.geoTargetConstant || 'Unknown';
      const status = campaignCriterion.status || 'Unknown';
      
      // Only include campaigns that have impressions (filter out zero impression campaigns)
      if (campaignsWithImpressions.has(campaignId)) {
        // Extract just the ID number from geoTargetConstants/2840 format
        const cleanLocationId = locationId.replace('geoTargetConstants/', '');
        
        // Get location name
        const locationName = locationNames[cleanLocationId] || `Unknown (ID: ${cleanLocationId})`;
        
        let newRow = [
          campaignId,
          campaignName,
          channelType,
          cleanLocationId,
          locationName,
          status
        ];
        
        data.push(newRow);
      }
      
    } catch (error) {
      Logger.log(`Error processing targeted location row: ${error.message}`);
    }
  });
  
  Logger.log(`✓ Processed ${data.length} targeted location records`);
  Logger.log(`✓ Total unique targeted locations: ${Object.keys(locationNames).length}`);
  
  return data;
}

function getLocationNamesFromIds(locationIds) {
  try {
    const locationNames = {};
    
    if (locationIds.length === 0) {
      return locationNames;
    }
    
    // Query geo_target_constant to get names from IDs - Country level only
    const locationIdsList = locationIds.map(id => `"${id}"`).join(',');
    const query = `
      SELECT
          geo_target_constant.id,
          geo_target_constant.name,
          geo_target_constant.target_type
      FROM geo_target_constant
      WHERE geo_target_constant.id IN (${locationIdsList})
        AND geo_target_constant.target_type = 'Country'
    `;
    
    Logger.log(`Querying geo_target_constant for ${locationIds.length} location IDs...`);
    
    const rows = AdsApp.search(query);
    while (rows.hasNext()) {
      const row = rows.next();
      const geoTargetConstant = row.geoTargetConstant || {};
      const id = geoTargetConstant.id || '';
      const name = geoTargetConstant.name || '';
      const targetType = geoTargetConstant.targetType || '';
      
      // Only include country-level locations
      if (id && name && targetType === 'Country') {
        locationNames[id] = name;
        Logger.log(`Mapped country ID ${id} to name: ${name}`);
      }
    }
    
    Logger.log(`Successfully mapped ${Object.keys(locationNames).length} location IDs to names`);
    return locationNames;
    
  } catch (error) {
    Logger.log(`Error getting location names from IDs: ${error.message}`);
    return {};
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

function testScript() {
  Logger.log("=== Testing Location Audit Report ===");
  main();
}
