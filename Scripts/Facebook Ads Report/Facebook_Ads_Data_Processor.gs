/**
 * Facebook Ads Data Processor
 * Automatically processes Facebook ads data and generates comprehensive reports
 * 
 * Features:
 * - Data validation and cleaning
 * - Automatic metric calculations
 * - Report generation
 * - Dashboard updates
 * - Error handling and notifications
 */

// Configuration
const CONFIG = {
  // Sheet names
  RAW_DATA_SHEET: 'Raw Data',
  PROCESSED_DATA_SHEET: 'Processed Data',
  REPORT_SHEET: 'Report',
  DASHBOARD_SHEET: 'Dashboard',
  
  // Column mappings for Facebook ads data
  COLUMN_MAPPINGS: {
    'Campaign Name': 'campaign_name',
    'Ad Set Name': 'ad_set_name',
    'Ad Name': 'ad_name',
    'Date': 'date',
    'Impressions': 'impressions',
    'Clicks': 'clicks',
    'Spend': 'spend',
    'Results': 'results',
    'Cost per Result': 'cost_per_result',
    'Reach': 'reach',
    'Frequency': 'frequency',
    'CPM': 'cpm',
    'CPC': 'cpc',
    'CTR': 'ctr',
    'Relevance Score': 'relevance_score',
    'Quality Ranking': 'quality_ranking',
    'Engagement Rate Ranking': 'engagement_rate_ranking',
    'Conversion Rate Ranking': 'conversion_rate_ranking',
    'Revenue': 'revenue',
    'ROAS': 'roas'
  },
  
  // Performance thresholds
  THRESHOLDS: {
    CTR_MIN: 0.01, // 1%
    CPC_MAX: 5.00, // $5
    CPM_MAX: 50.00, // $50
    ROAS_MIN: 2.00 // 2:1
  },
  
  // Report settings
  REPORT_SETTINGS: {
    DATE_RANGE_DAYS: 30,
    TOP_CAMPAIGNS_COUNT: 10,
    UPDATE_FREQUENCY_HOURS: 6
  }
};

/**
 * Main function to process Facebook ads data
 */
function processFacebookAdsData() {
  try {
    console.log('Starting Facebook Ads data processing...');
    
    // Get the active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Process raw data
    const processedData = processRawData(spreadsheet);
    
    // Generate reports
    generateReports(spreadsheet, processedData);
    
    // Update dashboard
    updateDashboard(spreadsheet, processedData);
    
    // Create client summary report
    createClientSummaryReport(spreadsheet, processedData);
    
    // Send notifications
    sendProcessingNotification('Data processing completed successfully');
    
    console.log('Facebook Ads data processing completed successfully');
    
  } catch (error) {
    console.error('Error processing Facebook Ads data:', error);
    sendProcessingNotification('Error processing data: ' + error.message, true);
  }
}

/**
 * Process raw Facebook ads data
 */
function processRawData(spreadsheet) {
  let rawDataSheet = spreadsheet.getSheetByName(CONFIG.RAW_DATA_SHEET);
  
  // If not found, try to find it with different names
  if (!rawDataSheet) {
    const sheets = spreadsheet.getSheets();
    const possibleNames = ['Raw Data', 'RawData', 'raw data', 'rawdata', 'Data', 'data'];
    
    for (const sheet of sheets) {
      if (possibleNames.includes(sheet.getName())) {
        console.log(`Found sheet with name: "${sheet.getName()}" - renaming to "Raw Data"`);
        sheet.setName('Raw Data');
        rawDataSheet = sheet;
        break;
      }
    }
  }
  
  if (!rawDataSheet) {
    throw new Error('Raw Data sheet not found. Please create a sheet named "Raw Data" with the required headers.');
  }
  
  const rawData = rawDataSheet.getDataRange().getValues();
  const headers = rawData[0];
  const dataRows = rawData.slice(1);
  
  console.log('Processing ' + dataRows.length + ' rows of data');
  
  // Validate and clean data
  const cleanedData = validateAndCleanData(dataRows, headers);
  
  // Calculate additional metrics
  const processedData = calculateMetrics(cleanedData);
  
  // Write processed data to sheet
  writeProcessedData(spreadsheet, processedData);
  
  return processedData;
}

/**
 * Validate and clean raw data
 */
function validateAndCleanData(dataRows, headers) {
  const cleanedData = [];
  
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    const cleanedRow = {};
    
    // Map columns to standardized names
    for (let j = 0; j < headers.length; j++) {
      const header = headers[j];
      const value = row[j];
      
      if (CONFIG.COLUMN_MAPPINGS[header]) {
        cleanedRow[CONFIG.COLUMN_MAPPINGS[header]] = value;
      }
    }
    
    // Validate required fields
    if (isValidRow(cleanedRow)) {
      cleanedData.push(cleanedRow);
    } else {
      console.warn('Skipping invalid row ' + (i + 2) + ': missing required fields');
    }
  }
  
  console.log('Cleaned data: ' + cleanedData.length + ' valid rows');
  return cleanedData;
}

/**
 * Check if a row has required fields
 */
function isValidRow(row) {
  const requiredFields = ['campaign_name', 'date', 'impressions', 'clicks', 'spend'];
  
  for (const field of requiredFields) {
    if (!row[field] || row[field] === '') {
      return false;
    }
  }
  
  return true;
}

/**
 * Calculate additional metrics
 */
function calculateMetrics(data) {
  return data.map(row => {
    const metrics = { ...row };
    
    // Convert string values to numbers
    const impressions = Number(row.impressions) || 0;
    const clicks = Number(row.clicks) || 0;
    const spend = Number(row.spend) || 0;
    const results = Number(row.results) || 0;
    const revenue = Number(row.revenue) || 0;
    
    // Calculate derived metrics
    metrics.ctr = impressions > 0 ? (clicks / impressions) * 100 : 0;
    metrics.cpc = clicks > 0 ? spend / clicks : 0;
    metrics.cpm = impressions > 0 ? (spend / impressions) * 1000 : 0;
    metrics.cost_per_result = results > 0 ? spend / results : 0;
    metrics.conversion_rate = clicks > 0 ? (results / clicks) * 100 : 0; // This is correct: conversions/clicks
    metrics.roas = spend > 0 ? revenue / spend : 0;
    
    // Add performance indicators
    metrics.performance_score = calculatePerformanceScore(metrics);
    metrics.status = getPerformanceStatus(metrics);
    
    return metrics;
  });
}

/**
 * Calculate performance score (0-100)
 */
function calculatePerformanceScore(metrics) {
  let score = 0;
  
  // CTR score (0-25 points)
  if (metrics.ctr >= CONFIG.THRESHOLDS.CTR_MIN) {
    score += 25;
  } else {
    score += (metrics.ctr / CONFIG.THRESHOLDS.CTR_MIN) * 25;
  }
  
  // CPC score (0-25 points)
  if (metrics.cpc <= CONFIG.THRESHOLDS.CPC_MAX) {
    score += 25;
  } else {
    score += (CONFIG.THRESHOLDS.CPC_MAX / metrics.cpc) * 25;
  }
  
  // CPM score (0-25 points)
  if (metrics.cpm <= CONFIG.THRESHOLDS.CPM_MAX) {
    score += 25;
  } else {
    score += (CONFIG.THRESHOLDS.CPM_MAX / metrics.cpm) * 25;
  }
  
  // ROAS score (0-25 points)
  if (metrics.roas >= CONFIG.THRESHOLDS.ROAS_MIN) {
    score += 25;
  } else {
    score += (metrics.roas / CONFIG.THRESHOLDS.ROAS_MIN) * 25;
  }
  
  return Math.round(score);
}

/**
 * Get performance status based on metrics
 */
function getPerformanceStatus(metrics) {
  if (metrics.performance_score >= 80) return 'Excellent';
  if (metrics.performance_score >= 60) return 'Good';
  if (metrics.performance_score >= 40) return 'Fair';
  return 'Poor';
}

/**
 * Write processed data to sheet
 */
function writeProcessedData(spreadsheet, data) {
  let processedSheet = spreadsheet.getSheetByName(CONFIG.PROCESSED_DATA_SHEET);
  
  // Create sheet if it doesn't exist
  if (!processedSheet) {
    processedSheet = spreadsheet.insertSheet(CONFIG.PROCESSED_DATA_SHEET);
  } else {
    processedSheet.clear();
  }
  
  // Define headers
  const headers = [
    'Campaign Name', 'Ad Set Name', 'Ad Name', 'Date', 'Impressions', 'Clicks',
    'Spend', 'Results', 'Revenue', 'CTR (%)', 'CPC ($)', 'CPM ($)', 'Cost per Result ($)',
    'Conversion Rate (%)', 'ROAS', 'Performance Score', 'Status'
  ];
  
  // Prepare data for writing
  const sheetData = [headers];
  
  data.forEach(row => {
    sheetData.push([
      row.campaign_name || '',
      row.ad_set_name || '',
      row.ad_name || '',
      row.date || '',
      row.impressions || 0,
      row.clicks || 0,
      row.spend || 0,
      row.results || 0,
      row.revenue || 0,
      row.ctr || 0,
      row.cpc || 0,
      row.cpm || 0,
      row.cost_per_result || 0,
      row.conversion_rate || 0,
      row.roas || 0,
      row.performance_score || 0,
      row.status || ''
    ]);
  });
  
  // Write data to sheet
  processedSheet.getRange(1, 1, sheetData.length, headers.length).setValues(sheetData);
  
  // Format the sheet
  formatProcessedDataSheet(processedSheet);
  
  console.log('Processed data written to sheet');
}

/**
 * Format the processed data sheet
 */
function formatProcessedDataSheet(sheet) {
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  
  // Format numeric columns
  const numericColumns = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]; // Impressions through Performance Score
  numericColumns.forEach(col => {
    const range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
    range.setNumberFormat('#,##0.00');
  });
  
  // Format percentage columns
  const percentageColumns = [10, 14]; // CTR and Conversion Rate
  percentageColumns.forEach(col => {
    const range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
    range.setNumberFormat('0.00%');
  });
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  
  // Add filters
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).createFilter();
}

/**
 * Generate comprehensive reports
 */
function generateReports(spreadsheet, data) {
  console.log('Generating reports...');
  
  // Campaign performance summary
  const campaignSummary = generateCampaignSummary(data);
  
  // Ad set analysis
  const adSetAnalysis = generateAdSetAnalysis(data);
  
  // Creative performance
  const creativeAnalysis = generateCreativeAnalysis(data);
  
  // Financial analysis
  const financialAnalysis = generateFinancialAnalysis(data);
  
  // Write reports to sheet
  writeReportsToSheet(spreadsheet, {
    campaignSummary,
    adSetAnalysis,
    creativeAnalysis,
    financialAnalysis
  });
}

/**
 * Generate campaign performance summary
 */
function generateCampaignSummary(data) {
  const campaignMap = new Map();
  
  data.forEach(row => {
    const campaignName = row.campaign_name;
    
    if (!campaignMap.has(campaignName)) {
      campaignMap.set(campaignName, {
        campaign_name: campaignName,
        impressions: 0,
        clicks: 0,
        spend: 0,
        results: 0,
        revenue: 0
      });
    }
    
    const campaign = campaignMap.get(campaignName);
    campaign.impressions += Number(row.impressions) || 0;
    campaign.clicks += Number(row.clicks) || 0;
    campaign.spend += Number(row.spend) || 0;
    campaign.results += Number(row.results) || 0;
    campaign.revenue += Number(row.revenue) || 0;
  });
  
  // Calculate metrics for each campaign
  const summary = Array.from(campaignMap.values()).map(campaign => {
    const impressions = campaign.impressions;
    const clicks = campaign.clicks;
    const spend = campaign.spend;
    const results = campaign.results;
    const revenue = campaign.revenue;
    
    return {
      ...campaign,
      ctr: impressions > 0 ? (clicks / impressions) * 100 : 0,
      cpc: clicks > 0 ? spend / clicks : 0,
      cpm: impressions > 0 ? (spend / impressions) * 1000 : 0,
      cost_per_result: results > 0 ? spend / results : 0,
      conversion_rate: clicks > 0 ? (results / clicks) * 100 : 0,
      roas: spend > 0 ? revenue / spend : 0
    };
  });
  
  // Sort by spend descending
  return summary.sort((a, b) => b.spend - a.spend);
}

/**
 * Generate ad set analysis
 */
function generateAdSetAnalysis(data) {
  const adSetMap = new Map();
  
  data.forEach(row => {
    const adSetName = row.ad_set_name;
    
    if (!adSetMap.has(adSetName)) {
      adSetMap.set(adSetName, {
        ad_set_name: adSetName,
        campaign_name: row.campaign_name,
        impressions: 0,
        clicks: 0,
        spend: 0,
        results: 0,
        revenue: 0
      });
    }
    
    const adSet = adSetMap.get(adSetName);
    adSet.impressions += Number(row.impressions) || 0;
    adSet.clicks += Number(row.clicks) || 0;
    adSet.spend += Number(row.spend) || 0;
    adSet.results += Number(row.results) || 0;
    adSet.revenue += Number(row.revenue) || 0;
  });
  
  // Calculate metrics for each ad set
  const analysis = Array.from(adSetMap.values()).map(adSet => {
    const impressions = adSet.impressions;
    const clicks = adSet.clicks;
    const spend = adSet.spend;
    const results = adSet.results;
    const revenue = adSet.revenue;
    
    return {
      ...adSet,
      ctr: impressions > 0 ? (clicks / impressions) * 100 : 0,
      cpc: clicks > 0 ? spend / clicks : 0,
      cpm: impressions > 0 ? (spend / impressions) * 1000 : 0,
      cost_per_result: results > 0 ? spend / results : 0,
      conversion_rate: clicks > 0 ? (results / clicks) * 100 : 0,
      roas: spend > 0 ? revenue / spend : 0
    };
  });
  
  // Sort by spend descending
  return analysis.sort((a, b) => b.spend - a.spend);
}

/**
 * Generate creative performance analysis
 */
function generateCreativeAnalysis(data) {
  const creativeMap = new Map();
  
  data.forEach(row => {
    const adName = row.ad_name;
    
    if (!creativeMap.has(adName)) {
      creativeMap.set(adName, {
        ad_name: adName,
        campaign_name: row.campaign_name,
        ad_set_name: row.ad_set_name,
        impressions: 0,
        clicks: 0,
        spend: 0,
        results: 0,
        revenue: 0
      });
    }
    
    const creative = creativeMap.get(adName);
    creative.impressions += Number(row.impressions) || 0;
    creative.clicks += Number(row.clicks) || 0;
    creative.spend += Number(row.spend) || 0;
    creative.results += Number(row.results) || 0;
    creative.revenue += Number(row.revenue) || 0;
  });
  
  // Calculate metrics for each creative
  const analysis = Array.from(creativeMap.values()).map(creative => {
    const impressions = creative.impressions;
    const clicks = creative.clicks;
    const spend = creative.spend;
    const results = creative.results;
    const revenue = creative.revenue;
    
    return {
      ...creative,
      ctr: impressions > 0 ? (clicks / impressions) * 100 : 0,
      cpc: clicks > 0 ? spend / clicks : 0,
      cpm: impressions > 0 ? (spend / impressions) * 1000 : 0,
      cost_per_result: results > 0 ? spend / results : 0,
      conversion_rate: clicks > 0 ? (results / clicks) * 100 : 0,
      roas: spend > 0 ? revenue / spend : 0
    };
  });
  
  // Sort by spend descending
  return analysis.sort((a, b) => b.spend - a.spend);
}

/**
 * Generate financial analysis
 */
function generateFinancialAnalysis(data) {
  const totalSpend = data.reduce((sum, row) => sum + (Number(row.spend) || 0), 0);
  const totalRevenue = data.reduce((sum, row) => sum + (Number(row.revenue) || 0), 0);
  const totalImpressions = data.reduce((sum, row) => sum + (Number(row.impressions) || 0), 0);
  const totalClicks = data.reduce((sum, row) => sum + (Number(row.clicks) || 0), 0);
  const totalResults = data.reduce((sum, row) => sum + (Number(row.results) || 0), 0);
  
  return {
    total_spend: totalSpend,
    total_revenue: totalRevenue,
    total_impressions: totalImpressions,
    total_clicks: totalClicks,
    total_results: totalResults,
    overall_ctr: totalImpressions > 0 ? (totalClicks / totalImpressions) * 100 : 0,
    overall_cpc: totalClicks > 0 ? totalSpend / totalClicks : 0,
    overall_cpm: totalImpressions > 0 ? (totalSpend / totalImpressions) * 1000 : 0,
    overall_cost_per_result: totalResults > 0 ? totalSpend / totalResults : 0,
    overall_conversion_rate: totalClicks > 0 ? (totalResults / totalClicks) * 100 : 0,
    overall_roas: totalSpend > 0 ? totalRevenue / totalSpend : 0,
    profit: totalRevenue - totalSpend,
    profit_margin: totalRevenue > 0 ? ((totalRevenue - totalSpend) / totalRevenue) * 100 : 0
  };
}

/**
 * Write reports to sheet
 */
function writeReportsToSheet(spreadsheet, reports) {
  let reportSheet = spreadsheet.getSheetByName(CONFIG.REPORT_SHEET);
  
  // Create sheet if it doesn't exist
  if (!reportSheet) {
    reportSheet = spreadsheet.insertSheet(CONFIG.REPORT_SHEET);
  } else {
    reportSheet.clear();
  }
  
  let currentRow = 1;
  
  // Write financial summary
  currentRow = writeFinancialSummary(reportSheet, reports.financialAnalysis, currentRow);
  currentRow += 2;
  
  // Write campaign summary
  currentRow = writeCampaignSummary(reportSheet, reports.campaignSummary, currentRow);
  currentRow += 2;
  
  // Write ad set analysis
  currentRow = writeAdSetAnalysis(reportSheet, reports.adSetAnalysis, currentRow);
  currentRow += 2;
  
  // Write creative analysis
  currentRow = writeCreativeAnalysis(reportSheet, reports.creativeAnalysis, currentRow);
  
  // Format the report sheet
  formatReportSheet(reportSheet);
}

/**
 * Write financial summary section
 */
function writeFinancialSummary(sheet, financial, startRow) {
  sheet.getRange(startRow, 1).setValue('FINANCIAL SUMMARY').setFontWeight('bold').setFontSize(14);
  startRow++;
  
  const summaryData = [
    ['Metric', 'Value'],
    ['Total Spend', financial.total_spend],
    ['Total Revenue', financial.total_revenue],
    ['Profit', financial.profit],
    ['Profit Margin', financial.profit_margin + '%'],
    ['Overall ROAS', financial.overall_roas],
    ['Overall CTR', financial.overall_ctr + '%'],
    ['Overall CPC', financial.overall_cpc],
    ['Overall CPM', financial.overall_cpm],
    ['Overall Cost per Result', financial.overall_cost_per_result],
    ['Overall Conversion Rate', financial.overall_conversion_rate + '%']
  ];
  
  sheet.getRange(startRow, 1, summaryData.length, 2).setValues(summaryData);
  
  return startRow + summaryData.length;
}

/**
 * Write campaign summary section
 */
function writeCampaignSummary(sheet, campaigns, startRow) {
  sheet.getRange(startRow, 1).setValue('CAMPAIGN PERFORMANCE').setFontWeight('bold').setFontSize(14);
  startRow++;
  
  const headers = [
    'Campaign Name', 'Spend', 'Impressions', 'Clicks', 'Results', 'Revenue',
    'CTR (%)', 'CPC ($)', 'CPM ($)', 'Cost per Result ($)', 'ROAS'
  ];
  
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  startRow++;
  
  // Write top campaigns (limit to 10)
  const topCampaigns = campaigns.slice(0, 10);
  const campaignData = topCampaigns.map(campaign => [
    campaign.campaign_name,
    campaign.spend,
    campaign.impressions,
    campaign.clicks,
    campaign.results,
    campaign.revenue,
    campaign.ctr,
    campaign.cpc,
    campaign.cpm,
    campaign.cost_per_result,
    campaign.roas
  ]);
  
  if (campaignData.length > 0) {
    sheet.getRange(startRow, 1, campaignData.length, headers.length).setValues(campaignData);
  }
  
  return startRow + campaignData.length;
}

/**
 * Write ad set analysis section
 */
function writeAdSetAnalysis(sheet, adSets, startRow) {
  sheet.getRange(startRow, 1).setValue('AD SET PERFORMANCE').setFontWeight('bold').setFontSize(14);
  startRow++;
  
  const headers = [
    'Ad Set Name', 'Campaign', 'Spend', 'Impressions', 'Clicks', 'Results', 'Revenue',
    'CTR (%)', 'CPC ($)', 'CPM ($)', 'Cost per Result ($)', 'ROAS'
  ];
  
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  startRow++;
  
  // Write top ad sets (limit to 10)
  const topAdSets = adSets.slice(0, 10);
  const adSetData = topAdSets.map(adSet => [
    adSet.ad_set_name,
    adSet.campaign_name,
    adSet.spend,
    adSet.impressions,
    adSet.clicks,
    adSet.results,
    adSet.revenue,
    adSet.ctr,
    adSet.cpc,
    adSet.cpm,
    adSet.cost_per_result,
    adSet.roas
  ]);
  
  if (adSetData.length > 0) {
    sheet.getRange(startRow, 1, adSetData.length, headers.length).setValues(adSetData);
  }
  
  return startRow + adSetData.length;
}

/**
 * Write creative analysis section
 */
function writeCreativeAnalysis(sheet, creatives, startRow) {
  sheet.getRange(startRow, 1).setValue('CREATIVE PERFORMANCE').setFontWeight('bold').setFontSize(14);
  startRow++;
  
  const headers = [
    'Ad Name', 'Campaign', 'Ad Set', 'Spend', 'Impressions', 'Clicks', 'Results', 'Revenue',
    'CTR (%)', 'CPC ($)', 'CPM ($)', 'Cost per Result ($)', 'ROAS'
  ];
  
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  startRow++;
  
  // Write top creatives (limit to 10)
  const topCreatives = creatives.slice(0, 10);
  const creativeData = topCreatives.map(creative => [
    creative.ad_name,
    creative.campaign_name,
    creative.ad_set_name,
    creative.spend,
    creative.impressions,
    creative.clicks,
    creative.results,
    creative.revenue,
    creative.ctr,
    creative.cpc,
    creative.cpm,
    creative.cost_per_result,
    creative.roas
  ]);
  
  if (creativeData.length > 0) {
    sheet.getRange(startRow, 1, creativeData.length, headers.length).setValues(creativeData);
  }
  
  return startRow + creativeData.length;
}

/**
 * Format the report sheet
 */
function formatReportSheet(sheet) {
  // Format section headers
  const sectionHeaders = sheet.getRangeList(['A1', 'A15', 'A30', 'A45']);
  sectionHeaders.setFontWeight('bold');
  sectionHeaders.setFontSize(14);
  sectionHeaders.setBackground('#f8f9fa');
  
  // Format data headers
  const dataHeaders = sheet.getRangeList(['A2:A12', 'A16:A26', 'A31:A41', 'A46:A56']);
  dataHeaders.setFontWeight('bold');
  dataHeaders.setBackground('#e8eaed');
  
  // Format numeric columns
  const numericRanges = sheet.getRangeList(['B3:B13', 'B17:B27', 'B32:B42', 'B47:B57']);
  numericRanges.setNumberFormat('#,##0.00');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

/**
 * Update dashboard with key metrics and visualizations
 */
function updateDashboard(spreadsheet, data) {
  console.log('Updating dashboard...');
  
  let dashboardSheet = spreadsheet.getSheetByName(CONFIG.DASHBOARD_SHEET);
  
  // Create sheet if it doesn't exist
  if (!dashboardSheet) {
    dashboardSheet = spreadsheet.insertSheet(CONFIG.DASHBOARD_SHEET);
  } else {
    dashboardSheet.clear();
  }
  
  // Calculate key metrics
  const totalSpend = data.reduce((sum, row) => sum + (Number(row.spend) || 0), 0);
  const totalRevenue = data.reduce((sum, row) => sum + (Number(row.revenue) || 0), 0);
  const totalImpressions = data.reduce((sum, row) => sum + (Number(row.impressions) || 0), 0);
  const totalClicks = data.reduce((sum, row) => sum + (Number(row.clicks) || 0), 0);
  const totalResults = data.reduce((sum, row) => sum + (Number(row.results) || 0), 0);
  
  const overallCtr = totalImpressions > 0 ? (totalClicks / totalImpressions) * 100 : 0;
  const overallCpc = totalClicks > 0 ? totalSpend / totalClicks : 0;
  const overallRoas = totalSpend > 0 ? totalRevenue / totalSpend : 0;
  const profit = totalRevenue - totalSpend;
  const conversionRate = totalClicks > 0 ? (totalResults / totalClicks) * 100 : 0;
  
  // Create dashboard layout
  const dashboardData = [
    ['FACEBOOK ADS PERFORMANCE DASHBOARD', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['KEY PERFORMANCE INDICATORS', '', '', '', '', ''],
    ['Metric', 'Value', 'Target', 'Status', 'Trend', ''],
    ['Total Spend', totalSpend, '$10,000', getStatusIndicator(totalSpend, 10000, 'below'), '📊', ''],
    ['Total Revenue', totalRevenue, '$20,000', getStatusIndicator(totalRevenue, 20000, 'above'), '📈', ''],
    ['Profit', profit, '$10,000', getStatusIndicator(profit, 10000, 'above'), profit >= 0 ? '📈' : '📉', ''],
    ['ROAS', overallRoas, '2.0', getStatusIndicator(overallRoas, 2.0, 'above'), overallRoas >= 2.0 ? '📈' : '📉', ''],
    ['CTR', overallCtr + '%', '1%', getStatusIndicator(overallCtr, 1, 'above'), overallCtr >= 1 ? '📈' : '📉', ''],
    ['CPC', overallCpc, '$5.00', getStatusIndicator(overallCpc, 5, 'below'), overallCpc <= 5 ? '📈' : '📉', ''],
    ['Conversion Rate', conversionRate + '%', '5%', getStatusIndicator(conversionRate, 5, 'above'), conversionRate >= 5 ? '📈' : '📉', ''],
    ['Total Impressions', totalImpressions, '100,000', getStatusIndicator(totalImpressions, 100000, 'above'), '📊', ''],
    ['Total Clicks', totalClicks, '1,000', getStatusIndicator(totalClicks, 1000, 'above'), '📊', ''],
    ['Total Results', totalResults, '50', getStatusIndicator(totalResults, 50, 'above'), '📊', ''],
    ['', '', '', '', '', ''],
    ['CAMPAIGN PERFORMANCE', '', '', '', '', ''],
    ['Campaign Name', 'Spend', 'Revenue', 'ROAS', 'Status', 'Performance']
  ];
  
  // Get top campaigns
  const campaignMap = new Map();
  data.forEach(row => {
    const campaignName = row.campaign_name;
    if (!campaignMap.has(campaignName)) {
      campaignMap.set(campaignName, {
        spend: 0,
        revenue: 0,
        impressions: 0,
        clicks: 0,
        results: 0
      });
    }
    const campaign = campaignMap.get(campaignName);
    campaign.spend += Number(row.spend) || 0;
    campaign.revenue += Number(row.revenue) || 0;
    campaign.impressions += Number(row.impressions) || 0;
    campaign.clicks += Number(row.clicks) || 0;
    campaign.results += Number(row.results) || 0;
  });
  
  const topCampaigns = Array.from(campaignMap.entries())
    .map(([name, data]) => ({
      name,
      spend: data.spend,
      revenue: data.revenue,
      roas: data.spend > 0 ? data.revenue / data.spend : 0,
      ctr: data.impressions > 0 ? (data.clicks / data.impressions) * 100 : 0,
      conversionRate: data.clicks > 0 ? (data.results / data.clicks) * 100 : 0
    }))
    .sort((a, b) => b.spend - a.spend)
    .slice(0, 5);
  
  // Add top campaigns to dashboard
  topCampaigns.forEach(campaign => {
    const performanceScore = calculateCampaignPerformanceScore(campaign);
    dashboardData.push([
      campaign.name,
      campaign.spend,
      campaign.revenue,
      campaign.roas,
      getStatusIndicator(campaign.roas, 2.0, 'above'),
      getPerformanceEmoji(performanceScore)
    ]);
  });
  
  // Write dashboard data
  dashboardSheet.getRange(1, 1, dashboardData.length, 6).setValues(dashboardData);
  
  // Format dashboard
  formatDashboardSheet(dashboardSheet);
  
  // Create visualizations
  createDashboardCharts(dashboardSheet, data, topCampaigns);
  
  console.log('Dashboard updated successfully with visualizations');
}

/**
 * Get status indicator based on value and target
 */
function getStatusIndicator(value, target, direction) {
  if (direction === 'above') {
    return value >= target ? '✅' : '❌';
  } else if (direction === 'below') {
    return value <= target ? '✅' : '❌';
  }
  return '📊';
}

/**
 * Calculate campaign performance score
 */
function calculateCampaignPerformanceScore(campaign) {
  let score = 0;
  
  // ROAS score (0-40 points)
  if (campaign.roas >= 2.0) score += 40;
  else if (campaign.roas >= 1.5) score += 30;
  else if (campaign.roas >= 1.0) score += 20;
  else score += 10;
  
  // CTR score (0-30 points)
  if (campaign.ctr >= 2.0) score += 30;
  else if (campaign.ctr >= 1.0) score += 20;
  else if (campaign.ctr >= 0.5) score += 10;
  
  // Conversion rate score (0-30 points)
  if (campaign.conversionRate >= 5.0) score += 30;
  else if (campaign.conversionRate >= 2.0) score += 20;
  else if (campaign.conversionRate >= 1.0) score += 10;
  
  return score;
}

/**
 * Get performance emoji based on score
 */
function getPerformanceEmoji(score) {
  if (score >= 80) return '🏆';
  if (score >= 60) return '🥇';
  if (score >= 40) return '🥈';
  if (score >= 20) return '🥉';
  return '📊';
}

/**
 * Create dashboard charts and visualizations
 */
function createDashboardCharts(sheet, data, topCampaigns) {
  try {
    // Create campaign performance chart
    createCampaignPerformanceChart(sheet, topCampaigns);
    
    // Create metrics comparison chart
    createMetricsComparisonChart(sheet, data);
    
    // Create performance summary chart
    createPerformanceSummaryChart(sheet, data);
    
    console.log('Dashboard charts created successfully');
  } catch (error) {
    console.error('Error creating charts:', error);
  }
}

/**
 * Create campaign performance chart
 */
function createCampaignPerformanceChart(sheet, campaigns) {
  if (campaigns.length === 0) return;
  
  // Prepare data for chart
  const chartData = [['Campaign', 'Spend', 'Revenue', 'ROAS']];
  campaigns.forEach(campaign => {
    chartData.push([
      campaign.name,
      campaign.spend,
      campaign.revenue,
      campaign.roas
    ]);
  });
  
  // Write chart data to sheet (starting from column H)
  const chartRange = sheet.getRange(1, 8, chartData.length, chartData[0].length);
  chartRange.setValues(chartData);
  
  // Create chart
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(chartRange)
    .setPosition(5, 8, 0, 0)
    .setOption('title', 'Campaign Performance Overview')
    .setOption('width', 600)
    .setOption('height', 400)
    .setOption('legend', { position: 'bottom' })
    .setOption('colors', ['#4285f4', '#34a853', '#fbbc04'])
    .build();
  
  sheet.insertChart(chart);
}

/**
 * Create metrics comparison chart
 */
function createMetricsComparisonChart(sheet, data) {
  // Calculate overall metrics
  const totalSpend = data.reduce((sum, row) => sum + (Number(row.spend) || 0), 0);
  const totalRevenue = data.reduce((sum, row) => sum + (Number(row.revenue) || 0), 0);
  const totalImpressions = data.reduce((sum, row) => sum + (Number(row.impressions) || 0), 0);
  const totalClicks = data.reduce((sum, row) => sum + (Number(row.clicks) || 0), 0);
  const totalResults = data.reduce((sum, row) => sum + (Number(row.results) || 0), 0);
  
  const overallCtr = totalImpressions > 0 ? (totalClicks / totalImpressions) * 100 : 0;
  const overallCpc = totalClicks > 0 ? totalSpend / totalClicks : 0;
  const overallRoas = totalSpend > 0 ? totalRevenue / totalSpend : 0;
  const conversionRate = totalClicks > 0 ? (totalResults / totalClicks) * 100 : 0;
  
  // Prepare data for chart
  const chartData = [
    ['Metric', 'Value', 'Target'],
    ['CTR (%)', overallCtr, 1.0],
    ['Conversion Rate (%)', conversionRate, 5.0],
    ['ROAS', overallRoas, 2.0],
    ['CPC ($)', overallCpc, 5.0]
  ];
  
  // Write chart data to sheet (starting from column H, row 25)
  const chartRange = sheet.getRange(25, 8, chartData.length, chartData[0].length);
  chartRange.setValues(chartData);
  
  // Create chart
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(chartRange)
    .setPosition(25, 8, 0, 0)
    .setOption('title', 'Performance vs Targets')
    .setOption('width', 500)
    .setOption('height', 300)
    .setOption('legend', { position: 'bottom' })
    .setOption('colors', ['#4285f4', '#ea4335'])
    .build();
  
  sheet.insertChart(chart);
}

/**
 * Create performance summary chart
 */
function createPerformanceSummaryChart(sheet, data) {
  // Group data by campaign
  const campaignMap = new Map();
  data.forEach(row => {
    const campaignName = row.campaign_name;
    if (!campaignMap.has(campaignName)) {
      campaignMap.set(campaignName, {
        spend: 0,
        revenue: 0,
        impressions: 0,
        clicks: 0,
        results: 0
      });
    }
    const campaign = campaignMap.get(campaignName);
    campaign.spend += Number(row.spend) || 0;
    campaign.revenue += Number(row.revenue) || 0;
    campaign.impressions += Number(row.impressions) || 0;
    campaign.clicks += Number(row.clicks) || 0;
    campaign.results += Number(row.results) || 0;
  });
  
  // Prepare data for pie chart (spend by campaign)
  const chartData = [['Campaign', 'Spend']];
  Array.from(campaignMap.entries()).forEach(([name, data]) => {
    chartData.push([name, data.spend]);
  });
  
  // Write chart data to sheet (starting from column H, row 50)
  const chartRange = sheet.getRange(50, 8, chartData.length, chartData[0].length);
  chartRange.setValues(chartData);
  
  // Create chart
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(chartRange)
    .setPosition(50, 8, 0, 0)
    .setOption('title', 'Spend Distribution by Campaign')
    .setOption('width', 400)
    .setOption('height', 300)
    .setOption('legend', { position: 'right' })
    .build();
  
  sheet.insertChart(chart);
}

/**
 * Create client-friendly summary report
 */
function createClientSummaryReport(spreadsheet, data) {
  console.log('Creating client summary report...');
  
  let summarySheet = spreadsheet.getSheetByName('Client Summary');
  
  // Create sheet if it doesn't exist
  if (!summarySheet) {
    summarySheet = spreadsheet.insertSheet('Client Summary');
  } else {
    summarySheet.clear();
  }
  
  // Calculate key metrics
  const totalSpend = data.reduce((sum, row) => sum + (Number(row.spend) || 0), 0);
  const totalRevenue = data.reduce((sum, row) => sum + (Number(row.revenue) || 0), 0);
  const totalImpressions = data.reduce((sum, row) => sum + (Number(row.impressions) || 0), 0);
  const totalClicks = data.reduce((sum, row) => sum + (Number(row.clicks) || 0), 0);
  const totalResults = data.reduce((sum, row) => sum + (Number(row.results) || 0), 0);
  
  const overallCtr = totalImpressions > 0 ? (totalClicks / totalImpressions) * 100 : 0;
  const overallCpc = totalClicks > 0 ? totalSpend / totalClicks : 0;
  const overallRoas = totalSpend > 0 ? totalRevenue / totalSpend : 0;
  const conversionRate = totalClicks > 0 ? (totalResults / totalClicks) * 100 : 0;
  const profit = totalRevenue - totalSpend;
  
  // Create client-friendly summary
  const summaryData = [
    ['FACEBOOK ADS PERFORMANCE SUMMARY', '', '', ''],
    ['', '', '', ''],
    ['📊 EXECUTIVE SUMMARY', '', '', ''],
    ['', '', '', ''],
    ['💰 Financial Performance', '', '', ''],
    ['Total Investment', totalSpend, 'Total amount spent on ads', ''],
    ['Total Revenue Generated', totalRevenue, 'Revenue from ad campaigns', ''],
    ['Net Profit', profit, 'Revenue minus ad spend', profit >= 0 ? '✅ Profitable' : '❌ Loss'],
    ['Return on Ad Spend (ROAS)', overallRoas, 'Revenue per dollar spent', overallRoas >= 2.0 ? '✅ Excellent' : overallRoas >= 1.5 ? '🟡 Good' : '❌ Needs Improvement'],
    ['', '', '', ''],
    ['📈 Campaign Performance', '', '', ''],
    ['Total Impressions', totalImpressions, 'Number of times ads were shown', ''],
    ['Total Clicks', totalClicks, 'Number of clicks on ads', ''],
    ['Click-Through Rate (CTR)', overallCtr + '%', 'Percentage of impressions that resulted in clicks', overallCtr >= 1.0 ? '✅ Above Average' : '❌ Below Average'],
    ['Cost Per Click (CPC)', overallCpc, 'Average cost per click', overallCpc <= 5.0 ? '✅ Good' : '❌ High'],
    ['Conversion Rate', conversionRate + '%', 'Percentage of clicks that resulted in conversions', conversionRate >= 5.0 ? '✅ Excellent' : conversionRate >= 2.0 ? '🟡 Good' : '❌ Needs Improvement'],
    ['', '', '', ''],
    ['🏆 TOP PERFORMING CAMPAIGNS', '', '', ''],
    ['Campaign Name', 'Spend', 'Revenue', 'ROAS', 'Performance Rating']
  ];
  
  // Get top campaigns
  const campaignMap = new Map();
  data.forEach(row => {
    const campaignName = row.campaign_name;
    if (!campaignMap.has(campaignName)) {
      campaignMap.set(campaignName, {
        spend: 0,
        revenue: 0,
        impressions: 0,
        clicks: 0,
        results: 0
      });
    }
    const campaign = campaignMap.get(campaignName);
    campaign.spend += Number(row.spend) || 0;
    campaign.revenue += Number(row.revenue) || 0;
    campaign.impressions += Number(row.impressions) || 0;
    campaign.clicks += Number(row.clicks) || 0;
    campaign.results += Number(row.results) || 0;
  });
  
  const topCampaigns = Array.from(campaignMap.entries())
    .map(([name, data]) => ({
      name,
      spend: data.spend,
      revenue: data.revenue,
      roas: data.spend > 0 ? data.revenue / data.spend : 0,
      ctr: data.impressions > 0 ? (data.clicks / data.impressions) * 100 : 0,
      conversionRate: data.clicks > 0 ? (data.results / data.clicks) * 100 : 0
    }))
    .sort((a, b) => b.roas - a.roas) // Sort by ROAS
    .slice(0, 5);
  
  // Add top campaigns to summary
  topCampaigns.forEach(campaign => {
    const performanceScore = calculateCampaignPerformanceScore(campaign);
    const rating = getPerformanceRating(performanceScore);
    summaryData.push([
      campaign.name,
      campaign.spend,
      campaign.revenue,
      campaign.roas,
      rating
    ]);
  });
  
  // Add recommendations
  summaryData.push(['', '', '', '']);
  summaryData.push(['💡 KEY INSIGHTS & RECOMMENDATIONS', '', '', '']);
  
  // Generate insights based on performance
  const insights = generateInsights(data, overallRoas, overallCtr, conversionRate, profit);
  insights.forEach(insight => {
    summaryData.push([insight, '', '', '']);
  });
  
  // Write summary data
  summarySheet.getRange(1, 1, summaryData.length, 4).setValues(summaryData);
  
  // Format the summary sheet
  formatClientSummarySheet(summarySheet);
  
  console.log('Client summary report created successfully');
}

/**
 * Get performance rating
 */
function getPerformanceRating(score) {
  if (score >= 80) return '🏆 Excellent';
  if (score >= 60) return '🥇 Good';
  if (score >= 40) return '🥈 Fair';
  if (score >= 20) return '🥉 Poor';
  return '📊 Needs Review';
}

/**
 * Generate insights based on performance data
 */
function generateInsights(data, roas, ctr, conversionRate, profit) {
  const insights = [];
  
  if (roas >= 2.0) {
    insights.push('✅ Strong ROAS indicates effective ad targeting and messaging');
  } else if (roas >= 1.5) {
    insights.push('🟡 ROAS is acceptable but could be improved with better targeting');
  } else {
    insights.push('❌ Low ROAS suggests need for audience optimization or creative improvements');
  }
  
  if (ctr >= 1.0) {
    insights.push('✅ Good click-through rate shows ads are resonating with audience');
  } else {
    insights.push('❌ Low CTR indicates need for more compelling ad creatives');
  }
  
  if (conversionRate >= 5.0) {
    insights.push('✅ High conversion rate shows effective landing pages and offers');
  } else if (conversionRate >= 2.0) {
    insights.push('🟡 Conversion rate is decent but could be improved');
  } else {
    insights.push('❌ Low conversion rate suggests need for landing page optimization');
  }
  
  if (profit > 0) {
    insights.push('💰 Campaigns are generating positive ROI');
  } else {
    insights.push('⚠️ Campaigns are currently operating at a loss - immediate optimization needed');
  }
  
  // Add campaign-specific insights
  const campaignCount = new Set(data.map(row => row.campaign_name)).size;
  if (campaignCount > 3) {
    insights.push('📊 Multiple campaigns running - consider consolidating for better budget allocation');
  }
  
  return insights;
}

/**
 * Format the client summary sheet
 */
function formatClientSummarySheet(sheet) {
  // Format title
  sheet.getRange('A1:D1').merge();
  sheet.getRange('A1').setFontWeight('bold').setFontSize(18).setHorizontalAlignment('center');
  sheet.getRange('A1').setBackground('#1a73e8').setFontColor('white');
  
  // Format section headers
  const sectionHeaders = ['A3:D3', 'A5:D5', 'A11:D11', 'A19:D19', 'A26:D26'];
  sectionHeaders.forEach(range => {
    sheet.getRange(range).merge();
    sheet.getRange(range.split(':')[0]).setFontWeight('bold').setFontSize(14).setBackground('#f8f9fa');
  });
  
  // Format currency columns
  sheet.getRange('B6:B8').setNumberFormat('$#,##0.00');
  sheet.getRange('B9').setNumberFormat('0.00');
  
  // Format percentage columns
  sheet.getRange('B14').setNumberFormat('0.00%');
  sheet.getRange('B16').setNumberFormat('0.00%');
  
  // Format campaign data
  const campaignDataRange = sheet.getRange(21, 1, sheet.getLastRow() - 20, 4);
  campaignDataRange.setBorder(true, true, true, true, true, true);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 4);
}

/**
 * Format the dashboard sheet
 */
function formatDashboardSheet(sheet) {
  // Format title
  sheet.getRange('A1:F1').merge();
  sheet.getRange('A1').setFontWeight('bold').setFontSize(18).setHorizontalAlignment('center');
  sheet.getRange('A1').setBackground('#1a73e8').setFontColor('white');
  
  // Format section headers
  sheet.getRange('A3:F3').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.getRange('A16:F16').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
  
  // Format data headers
  sheet.getRange('A4:F4').setFontWeight('bold').setBackground('#e8eaed');
  sheet.getRange('A17:F17').setFontWeight('bold').setBackground('#e8eaed');
  
  // Format numeric columns
  const numericColumns = [2, 3, 4]; // B, C, D columns
  numericColumns.forEach(col => {
    const range = sheet.getRange(5, col, 10, 1);
    range.setNumberFormat('#,##0.00');
  });
  
  // Format percentage columns
  sheet.getRange('B8').setNumberFormat('0.00%'); // ROAS
  sheet.getRange('B9').setNumberFormat('0.00%'); // CTR
  sheet.getRange('B10').setNumberFormat('0.00%'); // Conversion Rate
  
  // Format currency columns
  sheet.getRange('B5').setNumberFormat('$#,##0.00'); // Total Spend
  sheet.getRange('B6').setNumberFormat('$#,##0.00'); // Total Revenue
  sheet.getRange('B7').setNumberFormat('$#,##0.00'); // Profit
  sheet.getRange('B11').setNumberFormat('$#,##0.00'); // CPC
  
  // Add conditional formatting for status indicators
  const statusRange = sheet.getRange('D5:D14');
  const rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('✅')
    .setBackground('#d4edda')
    .setFontColor('#155724')
    .setRanges([statusRange])
    .build();
  
  const rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('❌')
    .setBackground('#f8d7da')
    .setFontColor('#721c24')
    .setRanges([statusRange])
    .build();
  
  sheet.setConditionalFormatRules([rule1, rule2]);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 6);
  
  // Add borders
  sheet.getRange('A1:F' + sheet.getLastRow()).setBorder(true, true, true, true, true, true);
}

/**
 * Send processing notification
 */
function sendProcessingNotification(message, isError = false) {
  try {
    // You can customize this to send emails, Slack notifications, etc.
    if (isError) {
      console.error('NOTIFICATION: ' + message);
    } else {
      console.log('NOTIFICATION: ' + message);
    }
    
    // Example: Send email notification
    // MailApp.sendEmail({
    //   to: 'your-email@example.com',
    //   subject: isError ? 'Facebook Ads Processing Error' : 'Facebook Ads Processing Complete',
    //   body: message
    // });
    
  } catch (error) {
    console.error('Error sending notification:', error);
  }
}

/**
 * Set up triggers for automatic processing
 */
function setupTriggers() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processFacebookAdsData') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new triggers
  // Trigger on edit to the Raw Data sheet
  ScriptApp.newTrigger('processFacebookAdsData')
    .onEdit()
    .create();
  
  // Time-based trigger (every 6 hours)
  ScriptApp.newTrigger('processFacebookAdsData')
    .timeBased()
    .everyHours(CONFIG.REPORT_SETTINGS.UPDATE_FREQUENCY_HOURS)
    .create();
  
  console.log('Triggers set up successfully');
}

/**
 * Manual trigger function for testing
 */
function manualProcess() {
  processFacebookAdsData();
}

/**
 * Initialize the reporting system
 */
function initializeReportingSystem() {
  try {
    console.log('Initializing Facebook Ads reporting system...');
    
    // Set up triggers
    setupTriggers();
    
    // Process existing data
    processFacebookAdsData();
    
    console.log('Facebook Ads reporting system initialized successfully');
    
  } catch (error) {
    console.error('Error initializing reporting system:', error);
  }
} 