// --- KEYWORD AUDIT REPORT FOR 80/20 ANALYSIS ---
// This script outputs keyword data by campaign for pivot table analysis
// Shows keyword performance metrics segmented by campaign across multiple date ranges
// Configure the account ID below to analyze a specific account

const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1ylEHOD31fvz4LAKxCucVD-zFPWxDww0LXBFF6eTRFuM/edit?usp=sharing';
const KEYWORDS_180_TAB = 'Keywords_180Days';
const KEYWORDS_90_TAB = 'Keywords_90Days';
const KEYWORDS_30_TAB = 'Keywords_30Days';
const KEYWORDS_14_TAB = 'Keywords_14Days';
const ACCOUNT_ID = '972-837-2864'; // Target account ID for testing

// Date ranges for analysis
const DATE_RANGES = [
  { days: 180, tab: KEYWORDS_180_TAB },
  { days: 90, tab: KEYWORDS_90_TAB },
  { days: 30, tab: KEYWORDS_30_TAB },
  { days: 14, tab: KEYWORDS_14_TAB }
];

// Minimum impressions threshold to include keywords
const MIN_IMPRESSIONS = 10;

// Base query template for keyword performance data - Search campaigns only, enabled campaigns only
const KEYWORDS_QUERY_TEMPLATE = `
  SELECT 
    keyword_view.resource_name,
    ad_group_criterion.keyword.text,
    ad_group_criterion.keyword.match_type,
    ad_group_criterion.status,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    campaign.status,
    ad_group.id,
    ad_group.name,
    ad_group.status,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
  FROM keyword_view
  WHERE segments.date BETWEEN "{START_DATE}" AND "{END_DATE}"
    AND campaign.status IN ('ENABLED', 'PAUSED', 'REMOVED')
    AND ad_group.status IN ('ENABLED', 'PAUSED', 'REMOVED')
    AND ad_group_criterion.status IN ('ENABLED', 'PAUSED')
    AND campaign.advertising_channel_type = 'SEARCH'
    AND ad_group_criterion.keyword.text IS NOT NULL
  ORDER BY metrics.cost_micros DESC
`;

// Date range utility function
function getDateRange(numDays) {
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(endDate.getDate() - numDays);
  
  const format = date => Utilities.formatDate(date, AdsApp.currentAccount().getTimeZone(), 'yyyyMMdd');
  
  return {
    startDate: format(startDate),
    endDate: format(endDate)
  };
}

function main() {
  try {
    Logger.log("=== Starting Keyword Audit Report for 80/20 Analysis ===");
    
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
      ss = SpreadsheetApp.create("Keyword Audit Report");
      let url = ss.getUrl();
      Logger.log("No SHEET_URL found, so this sheet was created: " + url);
    } else {
      ss = SpreadsheetApp.openByUrl(SHEET_URL);
    }
    
    // First, get keywords that qualify based on 180-day impressions
    Logger.log("=== Getting keywords that qualify based on 180-day impressions ===");
    const keywordsWithImpressions = getKeywordsWithSufficientImpressions();
    Logger.log(`Found ${keywordsWithImpressions.size} keywords with >= ${MIN_IMPRESSIONS} impressions over 180 days`);
    
    // Process Keywords for each date range
    Logger.log("=== Processing Keywords for Multiple Date Ranges ===");
    for (const dateRange of DATE_RANGES) {
      Logger.log(`--- Processing ${dateRange.days} days data ---`);
      processKeywordsForDateRange(ss, account, dateRange.days, dateRange.tab, keywordsWithImpressions);
    }
    
    Logger.log(`✓ Account analyzed: ${account.getName()} (${ACCOUNT_ID})`);
    Logger.log(`✓ Data exported to tabs: ${DATE_RANGES.map(r => r.tab).join(', ')}`);
    
  } catch (error) {
    Logger.log(`❌ Error in main function: ${error.message}`);
    Logger.log(`Error details: ${error.stack || 'No stack trace available'}`);
  }
}

function processKeywordsForDateRange(ss, account, numDays, tabName, keywordsWithImpressions) {
  try {
    // Get or create the keywords tab
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      Logger.log(`Created new tab: ${tabName}`);
    } else {
      sheet.clear();
      Logger.log(`Cleared existing data in tab: ${tabName}`);
    }
    
    // Generate date range
    const dateRange = getDateRange(numDays);
    const query = KEYWORDS_QUERY_TEMPLATE
      .replace('{START_DATE}', dateRange.startDate)
      .replace('{END_DATE}', dateRange.endDate);
    
    // Execute keywords query
    Logger.log(`Executing keywords query for ${numDays} days (${dateRange.startDate} to ${dateRange.endDate})...`);
    const rows = AdsApp.search(query);
    
    // Log sample row structure for debugging (only for first date range to avoid spam)
    if (numDays === 180) {
      const sampleQuery = query + ' LIMIT 1';
      const sampleRows = AdsApp.search(sampleQuery);
      
      if (sampleRows.hasNext()) {
        const sampleRow = sampleRows.next();
        Logger.log("Sample keyword row structure: " + JSON.stringify(sampleRow));
      } else {
        Logger.log("No sample rows found for keywords");
      }
    }
    
    // Process the data
    const data = calculateKeywordMetrics(rows, keywordsWithImpressions);
    
    if (data.length === 0) {
      Logger.log(`No keyword data found for ${numDays} days to export`);
      return;
    }
    
    // Create headers for keywords (removed Ad Group Status)
    const headers = [
      'Campaign ID',
      'Campaign Name', 
      'Channel Type',
      'Campaign Status',
      'Ad Group ID',
      'Ad Group Name',
      'Keyword Text',
      'Match Type',
      'Keyword Status',
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
    
    Logger.log(`✓ Successfully exported ${data.length} keyword records to ${tabName} tab`);
    
  } catch (error) {
    Logger.log(`❌ Error processing keywords for ${numDays} days: ${error.message}`);
  }
}

function calculateKeywordMetrics(rows, keywordsWithImpressions) {
  let data = [];
  let totalRows = 0;
  
  Logger.log("Processing keyword data...");
  
  while (rows.hasNext()) {
    try {
      const row = rows.next();
      
      // Access nested objects correctly
      const campaign = row.campaign || {};
      const adGroup = row.adGroup || {};
      const adGroupCriterion = row.adGroupCriterion || {};
      const keyword = adGroupCriterion.keyword || {};
      const metrics = row.metrics || {};
      
      // Extract campaign data
      const campaignId = campaign.id || 'Unknown';
      const campaignName = campaign.name || 'Unknown Campaign';
      const channelType = campaign.advertisingChannelType || 'Unknown';
      const campaignStatus = campaign.status || 'Unknown';
      
      // Extract ad group data
      const adGroupId = adGroup.id || 'Unknown';
      const adGroupName = adGroup.name || 'Unknown Ad Group';
      
      // Extract keyword data
      const keywordText = keyword.text || 'Unknown Keyword';
      const matchType = keyword.matchType || 'Unknown';
      const keywordStatus = adGroupCriterion.status || 'Unknown';
      
      // Create unique keyword identifier for filtering
      const keywordId = `${campaignId}_${adGroupId}_${keywordText}_${matchType}`;
      
      // Only process keywords that qualified based on 180-day impressions
      if (!keywordsWithImpressions.has(keywordId)) {
        continue;
      }
      
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
      
      // Include all qualified keywords regardless of impressions in this specific period
      // (they qualified based on 180-day performance)
      let newRow = [
        campaignId,
        campaignName,
        channelType,
        campaignStatus,
        adGroupId,
        adGroupName,
        keywordText,
        matchType,
        keywordStatus,
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
      
    } catch (error) {
      Logger.log(`Error processing keyword row: ${error.message}`);
    }
  }
  
  Logger.log(`✓ Processed ${totalRows} qualified keyword rows (qualified based on 180-day impressions)`);
  
  return data;
}

function getKeywordsWithSufficientImpressions() {
  try {
    Logger.log("Getting keywords with sufficient impressions over 180 days (qualification period)...");
    const keywordImpressions = new Map();
    
    // Only check 180 days for qualification
    const dateRangeObj = getDateRange(180);
    const query = KEYWORDS_QUERY_TEMPLATE
      .replace('{START_DATE}', dateRangeObj.startDate)
      .replace('{END_DATE}', dateRangeObj.endDate);
    
    Logger.log(`Checking 180 days (${dateRangeObj.startDate} to ${dateRangeObj.endDate}) for keyword impressions...`);
    
    const rows = AdsApp.search(query);
    
    while (rows.hasNext()) {
      try {
        const row = rows.next();
        const campaign = row.campaign || {};
        const adGroup = row.adGroup || {};
        const adGroupCriterion = row.adGroupCriterion || {};
        const keyword = adGroupCriterion.keyword || {};
        const metrics = row.metrics || {};
        
        const campaignId = campaign.id || 'Unknown';
        const adGroupId = adGroup.id || 'Unknown';
        const keywordText = keyword.text || 'Unknown Keyword';
        const matchType = keyword.matchType || 'Unknown';
        const impressions = Number(metrics.impressions) || 0;
        
        // Create unique keyword identifier
        const keywordId = `${campaignId}_${adGroupId}_${keywordText}_${matchType}`;
        
        // Add impressions to the total for this keyword
        const currentImpressions = keywordImpressions.get(keywordId) || 0;
        keywordImpressions.set(keywordId, currentImpressions + impressions);
        
      } catch (error) {
        Logger.log(`Error processing row in impression check: ${error.message}`);
      }
    }
    
    // Filter keywords that have at least MIN_IMPRESSIONS in 180 days
    const keywordsWithSufficientImpressions = new Set();
    for (const [keywordId, totalImpressions] of keywordImpressions) {
      if (totalImpressions >= MIN_IMPRESSIONS) {
        keywordsWithSufficientImpressions.add(keywordId);
      }
    }
    
    Logger.log(`Found ${keywordsWithSufficientImpressions.size} keywords with >= ${MIN_IMPRESSIONS} impressions over 180 days`);
    return keywordsWithSufficientImpressions;
    
  } catch (error) {
    Logger.log(`Error getting keywords with sufficient impressions: ${error.message}`);
    return new Set();
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
  Logger.log("=== Testing Keyword Audit Report ===");
  main();
}
