// --- DEMOGRAPHIC AUDIT REPORT FOR 80/20 ANALYSIS ---
// This script outputs demographic performance data by campaign across multiple time periods
// Shows gender, age range, and income range demographics with performance metrics for 180, 90, 30, and 14-day periods
// Only includes demographics with >10 impressions in the last 180 days
// Configure the account ID below to analyze a specific account

const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1ylEHOD31fvz4LAKxCucVD-zFPWxDww0LXBFF6eTRFuM/edit?usp=sharing';
const TAB_180_DAYS = 'Demographics_180_Days';
const TAB_90_DAYS = 'Demographics_90_Days';
const TAB_30_DAYS = 'Demographics_30_Days';
const TAB_14_DAYS = 'Demographics_14_Days';
const ACCOUNT_ID = '972-837-2864'; // Target account ID for testing

// Query for gender demographic performance - 180 days (used for filtering)
const GENDER_180_QUERY = `
  SELECT 
    ad_group_criterion.gender.type,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM gender_view
  WHERE segments.date DURING LAST_180_DAYS
    AND campaign.status = 'ENABLED'
  ORDER BY metrics.impressions DESC
`;

// Query for age range demographic performance - 180 days (used for filtering)
const AGE_RANGE_180_QUERY = `
  SELECT 
    ad_group_criterion.age_range.type,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM age_range_view
  WHERE segments.date DURING LAST_180_DAYS
    AND campaign.status = 'ENABLED'
  ORDER BY metrics.impressions DESC
`;

// Query for gender demographic performance - 90 days
const GENDER_90_QUERY = `
  SELECT 
    ad_group_criterion.gender.type,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM gender_view
  WHERE segments.date DURING LAST_90_DAYS
    AND campaign.status = 'ENABLED'
  ORDER BY metrics.impressions DESC
`;

// Query for age range demographic performance - 90 days
const AGE_RANGE_90_QUERY = `
  SELECT 
    ad_group_criterion.age_range.type,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM age_range_view
  WHERE segments.date DURING LAST_90_DAYS
    AND campaign.status = 'ENABLED'
  ORDER BY metrics.impressions DESC
`;

// Query for gender demographic performance - 30 days
const GENDER_30_QUERY = `
  SELECT 
    ad_group_criterion.gender.type,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM gender_view
  WHERE segments.date DURING LAST_30_DAYS
    AND campaign.status = 'ENABLED'
  ORDER BY metrics.impressions DESC
`;

// Query for age range demographic performance - 30 days
const AGE_RANGE_30_QUERY = `
  SELECT 
    ad_group_criterion.age_range.type,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM age_range_view
  WHERE segments.date DURING LAST_30_DAYS
    AND campaign.status = 'ENABLED'
  ORDER BY metrics.impressions DESC
`;

// Query for gender demographic performance - 14 days
const GENDER_14_QUERY = `
  SELECT 
    ad_group_criterion.gender.type,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM gender_view
  WHERE segments.date DURING LAST_14_DAYS
    AND campaign.status = 'ENABLED'
  ORDER BY metrics.impressions DESC
`;

// Query for age range demographic performance - 14 days
const AGE_RANGE_14_QUERY = `
  SELECT 
    ad_group_criterion.age_range.type,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM age_range_view
  WHERE segments.date DURING LAST_14_DAYS
    AND campaign.status = 'ENABLED'
  ORDER BY metrics.impressions DESC
`;

// Query for income range demographic performance - 180 days (used for filtering)
const INCOME_RANGE_180_QUERY = `
  SELECT 
    ad_group_criterion.income_range.type,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM income_range_view
  WHERE segments.date DURING LAST_180_DAYS
    AND campaign.status = 'ENABLED'
  ORDER BY metrics.impressions DESC
`;

// Query for income range demographic performance - 90 days
const INCOME_RANGE_90_QUERY = `
  SELECT 
    ad_group_criterion.income_range.type,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM income_range_view
  WHERE segments.date DURING LAST_90_DAYS
    AND campaign.status = 'ENABLED'
  ORDER BY metrics.impressions DESC
`;

// Query for income range demographic performance - 30 days
const INCOME_RANGE_30_QUERY = `
  SELECT 
    ad_group_criterion.income_range.type,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM income_range_view
  WHERE segments.date DURING LAST_30_DAYS
    AND campaign.status = 'ENABLED'
  ORDER BY metrics.impressions DESC
`;

// Query for income range demographic performance - 14 days
const INCOME_RANGE_14_QUERY = `
  SELECT 
    ad_group_criterion.income_range.type,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM income_range_view
  WHERE segments.date DURING LAST_14_DAYS
    AND campaign.status = 'ENABLED'
  ORDER BY metrics.impressions DESC
`;

function main() {
  try {
    Logger.log("=== Starting Demographic Audit Report for 80/20 Analysis ===");
    
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
      ss = SpreadsheetApp.create("Demographic Audit Report");
      let url = ss.getUrl();
      Logger.log("No SHEET_URL found, so this sheet was created: " + url);
    } else {
      ss = SpreadsheetApp.openByUrl(SHEET_URL);
    }
    
    // First, get demographics with >10 impressions in last 180 days for filtering
    Logger.log("=== Getting demographics with >10 impressions in last 180 days ===");
    const qualifyingDemographics = getQualifyingDemographics();
    
    if (qualifyingDemographics.size === 0) {
      Logger.log("❌ No demographics found with >10 impressions in last 180 days");
      return;
    }
    
    Logger.log(`✓ Found ${qualifyingDemographics.size} qualifying demographics`);
    
    // Process each time period for gender, age range, and income range demographics
    Logger.log("=== Processing 180-day demographic data ===");
    processDemographicData(ss, GENDER_180_QUERY, AGE_RANGE_180_QUERY, INCOME_RANGE_180_QUERY, TAB_180_DAYS, qualifyingDemographics, '180 days');
    
    Logger.log("=== Processing 90-day demographic data ===");
    processDemographicData(ss, GENDER_90_QUERY, AGE_RANGE_90_QUERY, INCOME_RANGE_90_QUERY, TAB_90_DAYS, qualifyingDemographics, '90 days');
    
    Logger.log("=== Processing 30-day demographic data ===");
    processDemographicData(ss, GENDER_30_QUERY, AGE_RANGE_30_QUERY, INCOME_RANGE_30_QUERY, TAB_30_DAYS, qualifyingDemographics, '30 days');
    
    Logger.log("=== Processing 14-day demographic data ===");
    processDemographicData(ss, GENDER_14_QUERY, AGE_RANGE_14_QUERY, INCOME_RANGE_14_QUERY, TAB_14_DAYS, qualifyingDemographics, '14 days');
    
    Logger.log(`✓ Account analyzed: ${account.getName()} (${ACCOUNT_ID})`);
    Logger.log(`✓ Data exported to four tabs: ${TAB_180_DAYS}, ${TAB_90_DAYS}, ${TAB_30_DAYS}, ${TAB_14_DAYS}`);
    
  } catch (error) {
    Logger.log(`❌ Error in main function: ${error.message}`);
    Logger.log(`Error details: ${error.stack || 'No stack trace available'}`);
  }
}

function getQualifyingDemographics() {
  try {
    Logger.log("Getting demographics with >10 impressions in last 180 days...");
    const qualifyingDemographics = new Set();
    
    // Process gender demographics
    Logger.log("Processing gender demographics...");
    const genderRows = AdsApp.search(GENDER_180_QUERY);
    
    // Log sample row structure for debugging
    const sampleGenderQuery = GENDER_180_QUERY + ' LIMIT 1';
    const sampleGenderRows = AdsApp.search(sampleGenderQuery);
    
    if (sampleGenderRows.hasNext()) {
      const sampleRow = sampleGenderRows.next();
      Logger.log("Sample gender row structure: " + JSON.stringify(sampleRow));
    } else {
      Logger.log("No sample gender rows found");
    }
    
    while (genderRows.hasNext()) {
      try {
        const row = genderRows.next();
        const adGroupCriterion = row.adGroupCriterion || {};
        const campaign = row.campaign || {};
        const metrics = row.metrics || {};
        
        const gender = adGroupCriterion.gender ? adGroupCriterion.gender.type : 'Unknown';
        const campaignId = campaign.id || 'Unknown';
        const impressions = Number(metrics.impressions) || 0;
        
        // Create unique key combining demographic type, value, and campaign
        const key = `gender_${gender}_${campaignId}`;
        
        if (impressions > 10) {
          qualifyingDemographics.add(key);
        }
        
      } catch (error) {
        Logger.log(`Error processing qualifying gender row: ${error.message}`);
      }
    }
    
    // Process age range demographics
    Logger.log("Processing age range demographics...");
    const ageRangeRows = AdsApp.search(AGE_RANGE_180_QUERY);
    
    // Log sample row structure for debugging
    const sampleAgeQuery = AGE_RANGE_180_QUERY + ' LIMIT 1';
    const sampleAgeRows = AdsApp.search(sampleAgeQuery);
    
    if (sampleAgeRows.hasNext()) {
      const sampleRow = sampleAgeRows.next();
      Logger.log("Sample age range row structure: " + JSON.stringify(sampleRow));
    } else {
      Logger.log("No sample age range rows found");
    }
    
    while (ageRangeRows.hasNext()) {
      try {
        const row = ageRangeRows.next();
        const adGroupCriterion = row.adGroupCriterion || {};
        const campaign = row.campaign || {};
        const metrics = row.metrics || {};
        
        const ageRange = adGroupCriterion.ageRange ? adGroupCriterion.ageRange.type : 'Unknown';
        const campaignId = campaign.id || 'Unknown';
        const impressions = Number(metrics.impressions) || 0;
        
        // Create unique key combining demographic type, value, and campaign
        const key = `age_${ageRange}_${campaignId}`;
        
        if (impressions > 10) {
          qualifyingDemographics.add(key);
        }
        
      } catch (error) {
        Logger.log(`Error processing qualifying age range row: ${error.message}`);
      }
    }
    
    // Process income range demographics
    Logger.log("Processing income range demographics...");
    const incomeRangeRows = AdsApp.search(INCOME_RANGE_180_QUERY);
    
    // Log sample row structure for debugging
    const sampleIncomeQuery = INCOME_RANGE_180_QUERY + ' LIMIT 1';
    const sampleIncomeRows = AdsApp.search(sampleIncomeQuery);
    
    if (sampleIncomeRows.hasNext()) {
      const sampleRow = sampleIncomeRows.next();
      Logger.log("Sample income range row structure: " + JSON.stringify(sampleRow));
    } else {
      Logger.log("No sample income range rows found");
    }
    
    while (incomeRangeRows.hasNext()) {
      try {
        const row = incomeRangeRows.next();
        const adGroupCriterion = row.adGroupCriterion || {};
        const campaign = row.campaign || {};
        const metrics = row.metrics || {};
        
        const incomeRange = adGroupCriterion.incomeRange ? adGroupCriterion.incomeRange.type : 'Unknown';
        const campaignId = campaign.id || 'Unknown';
        const impressions = Number(metrics.impressions) || 0;
        
        // Create unique key combining demographic type, value, and campaign
        const key = `income_${incomeRange}_${campaignId}`;
        
        if (impressions > 10) {
          qualifyingDemographics.add(key);
        }
        
      } catch (error) {
        Logger.log(`Error processing qualifying income range row: ${error.message}`);
      }
    }
    
    Logger.log(`Found ${qualifyingDemographics.size} demographic-campaign combinations with >10 impressions`);
    return qualifyingDemographics;
    
  } catch (error) {
    Logger.log(`Error getting qualifying demographics: ${error.message}`);
    return new Set();
  }
}

function processDemographicData(ss, genderQuery, ageRangeQuery, incomeRangeQuery, tabName, qualifyingDemographics, periodLabel) {
  try {
    // Get or create the tab
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      Logger.log(`Created new tab: ${tabName}`);
    } else {
      sheet.clear();
      Logger.log(`Cleared existing data in tab: ${tabName}`);
    }
    
    // Process gender demographics
    Logger.log(`Executing ${periodLabel} gender demographic query...`);
    const genderRows = AdsApp.search(genderQuery);
    const genderData = calculateGenderMetrics(genderRows, qualifyingDemographics, periodLabel);
    
    // Process age range demographics
    Logger.log(`Executing ${periodLabel} age range demographic query...`);
    const ageRangeRows = AdsApp.search(ageRangeQuery);
    const ageRangeData = calculateAgeRangeMetrics(ageRangeRows, qualifyingDemographics, periodLabel);
    
    // Process income range demographics
    Logger.log(`Executing ${periodLabel} income range demographic query...`);
    const incomeRangeRows = AdsApp.search(incomeRangeQuery);
    const incomeRangeData = calculateIncomeRangeMetrics(incomeRangeRows, qualifyingDemographics, periodLabel);
    
    // Combine all datasets
    const data = [...genderData, ...ageRangeData, ...incomeRangeData];
    
    if (data.length === 0) {
      Logger.log(`No qualifying demographic data found for ${periodLabel}`);
      return;
    }
    
    // Create headers
    const headers = [
      'Campaign ID',
      'Campaign Name', 
      'Channel Type',
      'Campaign Status',
      'Demographic Type',
      'Demographic Value',
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
    
    Logger.log(`✓ Successfully exported ${data.length} demographic records to ${tabName} tab`);
    
  } catch (error) {
    Logger.log(`❌ Error processing ${periodLabel} demographic data: ${error.message}`);
  }
}

function calculateGenderMetrics(rows, qualifyingDemographics, periodLabel) {
  let data = [];
  let totalRows = 0;
  let filteredRows = 0;
  
  Logger.log(`Processing ${periodLabel} gender demographic data...`);
  
  while (rows.hasNext()) {
    try {
      const row = rows.next();
      totalRows++;
      
      const adGroupCriterion = row.adGroupCriterion || {};
      const campaign = row.campaign || {};
      const metrics = row.metrics || {};
      
      const gender = adGroupCriterion.gender ? adGroupCriterion.gender.type : 'Unknown';
      const campaignId = campaign.id || 'Unknown';
      const campaignName = campaign.name || 'Unknown Campaign';
      const channelType = campaign.advertisingChannelType || 'Unknown';
      const campaignStatus = campaign.status || 'Unknown';
      
      // Create unique key to check against qualifying demographics
      const key = `gender_${gender}_${campaignId}`;
      
      // Only process if this demographic-campaign combination qualified in 180-day period
      if (qualifyingDemographics.has(key)) {
        filteredRows++;
        
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
        
        let newRow = [
          campaignId,
          campaignName,
          channelType,
          campaignStatus,
          'Gender',
          gender,
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
      }
      
    } catch (error) {
      Logger.log(`Error processing ${periodLabel} gender row: ${error.message}`);
    }
  }
  
  Logger.log(`✓ Processed ${totalRows} total gender rows, ${filteredRows} qualifying rows for ${periodLabel}`);
  
  return data;
}

function calculateAgeRangeMetrics(rows, qualifyingDemographics, periodLabel) {
  let data = [];
  let totalRows = 0;
  let filteredRows = 0;
  
  Logger.log(`Processing ${periodLabel} age range demographic data...`);
  
  while (rows.hasNext()) {
    try {
      const row = rows.next();
      totalRows++;
      
      const adGroupCriterion = row.adGroupCriterion || {};
      const campaign = row.campaign || {};
      const metrics = row.metrics || {};
      
      const ageRange = adGroupCriterion.ageRange ? adGroupCriterion.ageRange.type : 'Unknown';
      const campaignId = campaign.id || 'Unknown';
      const campaignName = campaign.name || 'Unknown Campaign';
      const channelType = campaign.advertisingChannelType || 'Unknown';
      const campaignStatus = campaign.status || 'Unknown';
      
      // Create unique key to check against qualifying demographics
      const key = `age_${ageRange}_${campaignId}`;
      
      // Only process if this demographic-campaign combination qualified in 180-day period
      if (qualifyingDemographics.has(key)) {
        filteredRows++;
        
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
        
        let newRow = [
          campaignId,
          campaignName,
          channelType,
          campaignStatus,
          'Age Range',
          ageRange,
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
      }
      
    } catch (error) {
      Logger.log(`Error processing ${periodLabel} age range row: ${error.message}`);
    }
  }
  
  Logger.log(`✓ Processed ${totalRows} total age range rows, ${filteredRows} qualifying rows for ${periodLabel}`);
  
  return data;
}

function calculateIncomeRangeMetrics(rows, qualifyingDemographics, periodLabel) {
  let data = [];
  let totalRows = 0;
  let filteredRows = 0;
  
  Logger.log(`Processing ${periodLabel} income range demographic data...`);
  
  while (rows.hasNext()) {
    try {
      const row = rows.next();
      totalRows++;
      
      const adGroupCriterion = row.adGroupCriterion || {};
      const campaign = row.campaign || {};
      const metrics = row.metrics || {};
      
      const incomeRange = adGroupCriterion.incomeRange ? adGroupCriterion.incomeRange.type : 'Unknown';
      const campaignId = campaign.id || 'Unknown';
      const campaignName = campaign.name || 'Unknown Campaign';
      const channelType = campaign.advertisingChannelType || 'Unknown';
      const campaignStatus = campaign.status || 'Unknown';
      
      // Create unique key to check against qualifying demographics
      const key = `income_${incomeRange}_${campaignId}`;
      
      // Only process if this demographic-campaign combination qualified in 180-day period
      if (qualifyingDemographics.has(key)) {
        filteredRows++;
        
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
        
        let newRow = [
          campaignId,
          campaignName,
          channelType,
          campaignStatus,
          'Income Range',
          incomeRange,
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
      }
      
    } catch (error) {
      Logger.log(`Error processing ${periodLabel} income range row: ${error.message}`);
    }
  }
  
  Logger.log(`✓ Processed ${totalRows} total income range rows, ${filteredRows} qualifying rows for ${periodLabel}`);
  
  return data;
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
  Logger.log("=== Testing Demographic Audit Report ===");
  main();
}
