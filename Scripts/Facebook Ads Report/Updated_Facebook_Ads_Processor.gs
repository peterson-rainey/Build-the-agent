/**
 * UPDATED FACEBOOK ADS PROCESSOR
 * Designed to handle the new CSV format with proper column mapping and calculations
 */

// Configuration for the new CSV format
const CONFIG = {
  // Sheet names
  RAW_DATA_SHEET: 'Raw Data',
  PROCESSED_DATA_SHEET: 'Processed Data',
  REPORT_SHEET: 'Report',
  DASHBOARD_SHEET: 'Dashboard',
  
  // Column mappings for the new CSV format
  COLUMN_MAPPINGS: {
    'Campaign name': 'campaign_name',
    'Ad Set Name': 'ad_set_name',
    'Day': 'date',
    'Impressions': 'impressions',
    'Link clicks': 'clicks',
    'Amount spent (USD)': 'spend',
    'Results': 'results',
    'Reach': 'reach',
    'Frequency': 'frequency',
    'Result type': 'result_type'
  },
  
  // Performance thresholds
  THRESHOLDS: {
    CTR_MIN: 0.01, // 1%
    CPC_MAX: 5.00, // $5
    CPM_MAX: 50.00, // $50
    ROAS_MIN: 2.00 // 2:1 (estimated)
  },
  
  // Estimated revenue per conversion (you can adjust this)
  ESTIMATED_REVENUE_PER_CONVERSION: 100, // $100 per conversion
  
  // Revenue configuration sheet name
  REVENUE_CONFIG_SHEET: 'Revenue Config'
};

/**
 * Main function to process Facebook ads data with retry logic
 */
function processFacebookAdsData() {
  const maxRetries = 3;
  let attempt = 0;
  
  while (attempt < maxRetries) {
    try {
      console.log(`Starting Facebook Ads data processing... (Attempt ${attempt + 1}/${maxRetries})`);
      
      // Get the active spreadsheet
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // Process raw data
      const processedData = processRawData(spreadsheet);
      
      // Generate reports
      generateReports(spreadsheet, processedData);
      
      // Update dashboard
      updateDashboard(spreadsheet, processedData);
      
      // Send notifications
      sendProcessingNotification('Data processing completed successfully');
      
      console.log('Facebook Ads data processing completed successfully');
      return; // Success, exit the retry loop
      
    } catch (error) {
      attempt++;
      console.error(`Error processing Facebook Ads data (Attempt ${attempt}/${maxRetries}):`, error);
      
      if (attempt >= maxRetries) {
        sendProcessingNotification('Error processing data after ' + maxRetries + ' attempts: ' + error.message, true);
        throw error; // Re-throw the error after all retries exhausted
      }
      
      // Wait before retrying (exponential backoff)
      const waitTime = Math.pow(2, attempt) * 1000; // 2s, 4s, 8s
      console.log(`Waiting ${waitTime}ms before retry...`);
      Utilities.sleep(waitTime);
    }
  }
}

/**
 * Process raw Facebook ads data with error handling
 */
function processRawData(spreadsheet) {
  try {
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
    const processedData = calculateMetrics(cleanedData, spreadsheet);
    
    // Write processed data to sheet
    writeProcessedData(spreadsheet, processedData);
    
    return processedData;
    
  } catch (error) {
    console.error('Error in processRawData:', error);
    throw error;
  }
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
 * Check if a row has required fields and clean empty values
 */
function isValidRow(row) {
  const requiredFields = ['campaign_name', 'date'];
  
  // Check that campaign name and date exist (these are truly required)
  for (const field of requiredFields) {
    if (!row[field] || row[field] === '') {
      return false;
    }
  }
  
  // Convert empty numeric fields to 0
  const numericFields = ['impressions', 'clicks', 'spend', 'results', 'reach', 'frequency'];
  numericFields.forEach(field => {
    if (row[field] === '' || row[field] === null || row[field] === undefined) {
      row[field] = 0;
    }
  });
  
  return true;
}

/**
 * Calculate additional metrics for the new format
 */
function calculateMetrics(data, spreadsheet) {
  return data.map(row => {
    const metrics = { ...row };
    
    // Convert string values to numbers
    const impressions = Number(row.impressions) || 0;
    const clicks = Number(row.clicks) || 0;
    const spend = Number(row.spend) || 0;
    const results = Number(row.results) || 0;
    const reach = Number(row.reach) || 0;
    
    // Determine if this is a conversion result (not reach or link clicks)
    const isConversion = determineIfConversion(results, clicks, reach, row.result_type);
    metrics.is_conversion = isConversion;
    metrics.conversions = isConversion ? results : 0;
    
    // Calculate derived metrics
    metrics.ctr = impressions > 0 ? (clicks / impressions) * 100 : 0;
    metrics.cpc = clicks > 0 ? spend / clicks : 0;
    metrics.cpm = impressions > 0 ? (spend / impressions) * 1000 : 0;
    metrics.cost_per_result = metrics.conversions > 0 ? spend / metrics.conversions : 0;
    metrics.conversion_rate = clicks > 0 ? (metrics.conversions / clicks) * 100 : 0;
    
    // Calculate estimated revenue and ROAS (only for conversions)
    const revenuePerResult = getRevenuePerResult(spreadsheet);
    const estimatedRevenue = metrics.conversions * revenuePerResult;
    metrics.revenue = estimatedRevenue;
    metrics.roas = spend > 0 ? estimatedRevenue / spend : 0;
    
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
 * Write processed data to sheet with error handling
 */
function writeProcessedData(spreadsheet, data) {
  try {
    let processedSheet = spreadsheet.getSheetByName(CONFIG.PROCESSED_DATA_SHEET);
    
    // Create sheet if it doesn't exist
    if (!processedSheet) {
      processedSheet = spreadsheet.insertSheet(CONFIG.PROCESSED_DATA_SHEET);
    } else {
      processedSheet.clear();
    }
    
    // Define headers
    const headers = [
      'Campaign Name', 'Ad Set Name', 'Date', 'Impressions', 'Clicks', 'Spend',
      'Results', 'Reach', 'Frequency', 'Result Type', 'CTR (%)', 'CPC ($)', 'CPM ($)', 
      'Cost per Result ($)', 'Conversion Rate (%)', 'Estimated Revenue', 'ROAS', 
      'Performance Score', 'Status'
    ];
    
    // Prepare data for writing
    const sheetData = [headers];
    
    data.forEach(row => {
      sheetData.push([
        row.campaign_name || '',
        row.ad_set_name || '',
        row.date || '',
        row.impressions || 0,
        row.clicks || 0,
        row.spend || 0,
        row.results || 0,
        row.reach || 0,
        row.frequency || 0,
        row.result_type || '',
        (row.ctr || 0) / 100, // Convert percentage to decimal for proper formatting
        row.cpc || 0,
        row.cpm || 0,
        row.cost_per_result || 0,
        (row.conversion_rate || 0) / 100, // Convert percentage to decimal for proper formatting
        row.revenue || 0,
        row.roas || 0,
        row.performance_score || 0,
        row.status || ''
      ]);
    });
    
    // Write data to sheet in smaller chunks to avoid API limits
    const chunkSize = 100;
    for (let i = 0; i < sheetData.length; i += chunkSize) {
      const chunk = sheetData.slice(i, i + chunkSize);
      const range = processedSheet.getRange(i + 1, 1, chunk.length, headers.length);
      range.setValues(chunk);
      
      // Small delay between chunks
      if (i + chunkSize < sheetData.length) {
        Utilities.sleep(100);
      }
    }
    
    // Format the sheet
    formatProcessedDataSheet(processedSheet);
    
    console.log('Processed data written to sheet');
    
  } catch (error) {
    console.error('Error writing processed data:', error);
    throw error;
  }
}

/**
 * Format the processed data sheet with error handling
 */
function formatProcessedDataSheet(sheet) {
  try {
    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    
    // Format numeric columns
    const numericColumns = [4, 5, 6, 7, 8, 9, 12, 13, 14, 16, 17, 18]; // Impressions through Performance Score
    numericColumns.forEach(col => {
      const range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
      range.setNumberFormat('#,##0.00');
    });
    
    // Format percentage columns
    const percentageColumns = [11, 15]; // CTR and Conversion Rate
    percentageColumns.forEach(col => {
      const range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
      range.setNumberFormat('0.00%');
    });
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, sheet.getLastColumn());
    
    // Add filters (only if they don't already exist)
    try {
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).createFilter();
    } catch (error) {
      // Filter already exists, which is fine
      console.log('Filter already exists on sheet, skipping filter creation');
    }
    
  } catch (error) {
    console.error('Error formatting processed data sheet:', error);
    // Don't throw error for formatting issues, continue processing
  }
}

/**
 * Generate comprehensive reports with error handling
 */
function generateReports(spreadsheet, data) {
  try {
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
    
  } catch (error) {
    console.error('Error generating reports:', error);
    throw error;
  }
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
 * Generate creative performance analysis (using date as creative identifier)
 */
function generateCreativeAnalysis(data) {
  const creativeMap = new Map();
  
  data.forEach(row => {
    const creativeKey = row.date + ' - ' + row.campaign_name;
    
    if (!creativeMap.has(creativeKey)) {
      creativeMap.set(creativeKey, {
        creative_name: creativeKey,
        campaign_name: row.campaign_name,
        ad_set_name: row.ad_set_name,
        date: row.date,
        impressions: 0,
        clicks: 0,
        spend: 0,
        results: 0,
        revenue: 0
      });
    }
    
    const creative = creativeMap.get(creativeKey);
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
 * Write reports to sheet with error handling
 */
function writeReportsToSheet(spreadsheet, reports) {
  try {
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
    
  } catch (error) {
    console.error('Error writing reports to sheet:', error);
    throw error;
  }
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
    ['Total Estimated Revenue', financial.total_revenue],
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
    'Campaign Name', 'Spend', 'Impressions', 'Clicks', 'Results', 'Estimated Revenue',
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
    'Ad Set Name', 'Campaign', 'Spend', 'Impressions', 'Clicks', 'Results', 'Estimated Revenue',
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
  sheet.getRange(startRow, 1).setValue('DAILY PERFORMANCE').setFontWeight('bold').setFontSize(14);
  startRow++;
  
  const headers = [
    'Date - Campaign', 'Campaign', 'Ad Set', 'Spend', 'Impressions', 'Clicks', 'Results', 'Estimated Revenue',
    'CTR (%)', 'CPC ($)', 'CPM ($)', 'Cost per Result ($)', 'ROAS'
  ];
  
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  startRow++;
  
  // Write top creatives (limit to 10)
  const topCreatives = creatives.slice(0, 10);
  const creativeData = topCreatives.map(creative => [
    creative.creative_name,
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
  try {
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
    
  } catch (error) {
    console.error('Error formatting report sheet:', error);
    // Don't throw error for formatting issues
  }
}

/**
 * Update dashboard with key metrics and charts
 */
function updateDashboard(spreadsheet, data) {
  try {
    console.log('Updating dashboard...');
    
    let dashboardSheet = spreadsheet.getSheetByName(CONFIG.DASHBOARD_SHEET);
    
    // Create sheet if it doesn't exist
    if (!dashboardSheet) {
      dashboardSheet = spreadsheet.insertSheet(CONFIG.DASHBOARD_SHEET);
    } else {
      dashboardSheet.clear();
    }
    
    // Create dashboard sections
    createTopCampaignsSection(dashboardSheet, data);
    createPivotTableTab(spreadsheet, data);
    createAllChartsSection(dashboardSheet, data);
    
    console.log('Dashboard updated successfully');
    
  } catch (error) {
    console.error('Error updating dashboard:', error);
    throw error;
  }
}

/**
 * Format the dashboard sheet
 */
function formatDashboardSheet(sheet) {
  try {
    // Format title
    sheet.getRange('A1:D1').merge();
    sheet.getRange('A1').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
    
    // Format section headers
    sheet.getRange('A3:D3').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    sheet.getRange('A14:D14').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
    
    // Format numeric columns
    const numericRange = sheet.getRange('B4:B12');
    numericRange.setNumberFormat('#,##0.00');
    
    // Format percentage and currency columns
    sheet.getRange('B7').setNumberFormat('0.00%');
    sheet.getRange('B8').setNumberFormat('$#,##0.00');
    sheet.getRange('B9').setNumberFormat('$#,##0.00');
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, 4);
    
  } catch (error) {
    console.error('Error formatting dashboard sheet:', error);
    // Don't throw error for formatting issues
  }
}

/**
 * Send processing notification
 */
function sendProcessingNotification(message, isError = false) {
  try {
    if (isError) {
      console.error('NOTIFICATION: ' + message);
    } else {
      console.log('NOTIFICATION: ' + message);
    }
  } catch (error) {
    console.error('Error sending notification:', error);
  }
}

/**
 * Manual trigger function for testing
 */
function manualProcess() {
  processFacebookAdsData();
}

/**
 * Determine if a result is a conversion (not reach or link clicks)
 */
function determineIfConversion(results, clicks, reach, resultType) {
  // If result type is explicitly "Reach" or "Link clicks", it's not a conversion
  if (resultType && (resultType.toLowerCase() === 'reach' || resultType.toLowerCase() === 'link clicks')) {
    return false;
  }
  
  // If we have results and result type is not reach/link clicks, it's a conversion
  if (results > 0 && resultType && resultType.toLowerCase() !== 'reach' && resultType.toLowerCase() !== 'link clicks') {
    return true;
  }
  
  return false;
}

/**
 * Get revenue per result from config sheet
 */
function getRevenuePerResult(spreadsheet) {
  try {
    let configSheet = spreadsheet.getSheetByName(CONFIG.REVENUE_CONFIG_SHEET);
    
    if (!configSheet) {
      // Create config sheet if it doesn't exist
      configSheet = spreadsheet.insertSheet(CONFIG.REVENUE_CONFIG_SHEET);
      configSheet.getRange('A1').setValue('Revenue per Result ($)');
      configSheet.getRange('B1').setValue(CONFIG.ESTIMATED_REVENUE_PER_CONVERSION);
      configSheet.getRange('A1:B1').setFontWeight('bold');
      configSheet.getRange('B1').setNumberFormat('$#,##0.00');
      configSheet.autoResizeColumns(1, 2);
    }
    
    const revenuePerResult = configSheet.getRange('B1').getValue();
    return Number(revenuePerResult) || CONFIG.ESTIMATED_REVENUE_PER_CONVERSION;
    
  } catch (error) {
    console.error('Error getting revenue per result:', error);
    return CONFIG.ESTIMATED_REVENUE_PER_CONVERSION;
  }
}

/**
 * Generate enhanced reports with new sections
 */
function generateEnhancedReports(spreadsheet, data) {
  try {
    console.log('Generating enhanced reports...');
    
    // Campaign performance summary (aggregated by campaign)
    const campaignSummary = generateCampaignSummaryAggregated(data);
    
    // Day of week analysis for entire account
    const dayOfWeekAnalysis = generateDayOfWeekAnalysis(data);
    
    // Daily conversion trends
    const dailyTrends = generateDailyTrends(data);
    
    // Write enhanced reports to sheet
    writeEnhancedReportsToSheet(spreadsheet, {
      campaignSummary,
      dayOfWeekAnalysis,
      dailyTrends
    });
    
  } catch (error) {
    console.error('Error generating enhanced reports:', error);
    throw error;
  }
}

/**
 * Generate campaign summary with aggregated results
 */
function generateCampaignSummaryAggregated(data) {
  const campaignMap = new Map();
  
  data.forEach(row => {
    const campaignName = row.campaign_name;
    
    if (!campaignMap.has(campaignName)) {
      campaignMap.set(campaignName, {
              campaign_name: campaignName,
      total_days: 0,
      impressions: 0,
      clicks: 0,
      spend: 0,
      conversions: 0,
      revenue: 0
      });
    }
    
    const campaign = campaignMap.get(campaignName);
    campaign.total_days++;
    campaign.impressions += Number(row.impressions) || 0;
    campaign.clicks += Number(row.clicks) || 0;
    campaign.spend += Number(row.spend) || 0;
    campaign.conversions += Number(row.conversions) || 0;
    campaign.revenue += Number(row.revenue) || 0;
  });
  
  // Calculate metrics for each campaign
  const summary = Array.from(campaignMap.values()).map(campaign => {
    const impressions = campaign.impressions;
    const clicks = campaign.clicks;
    const spend = campaign.spend;
    const conversions = campaign.conversions;
    const revenue = campaign.revenue;
    
    return {
      ...campaign,
      avg_daily_conversions: campaign.total_days > 0 ? conversions / campaign.total_days : 0,
      ctr: impressions > 0 ? (clicks / impressions) * 100 : 0,
      cpc: clicks > 0 ? spend / clicks : 0,
      cpm: impressions > 0 ? (spend / impressions) * 1000 : 0,
      cost_per_conversion: conversions > 0 ? spend / conversions : 0,
      conversion_rate: clicks > 0 ? (conversions / clicks) * 100 : 0,
      roas: spend > 0 ? revenue / spend : 0
    };
  });
  
  // Sort by total conversions descending
  return summary.sort((a, b) => b.conversions - a.conversions);
}

/**
 * Generate day of week analysis for entire account
 */
function generateDayOfWeekAnalysis(data) {
  const dayMap = new Map();
  const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  
  data.forEach(row => {
    const date = new Date(row.date);
    const dayOfWeek = date.getDay();
    const dayName = dayNames[dayOfWeek];
    
    if (!dayMap.has(dayName)) {
      dayMap.set(dayName, {
        day_name: dayName,
        day_number: dayOfWeek,
        total_days: 0,
        impressions: 0,
        clicks: 0,
        spend: 0,
        conversions: 0,
        revenue: 0
      });
    }
    
    const day = dayMap.get(dayName);
    day.total_days++;
    day.impressions += Number(row.impressions) || 0;
    day.clicks += Number(row.clicks) || 0;
    day.spend += Number(row.spend) || 0;
    day.conversions += Number(row.conversions) || 0;
    day.revenue += Number(row.revenue) || 0;
  });
  
  // Calculate metrics for each day
  const analysis = Array.from(dayMap.values()).map(day => {
    const impressions = day.impressions;
    const clicks = day.clicks;
    const spend = day.spend;
    const conversions = day.conversions;
    const revenue = day.revenue;
    
    return {
      ...day,
      avg_daily_conversions: day.total_days > 0 ? conversions / day.total_days : 0,
      avg_daily_spend: day.total_days > 0 ? spend / day.total_days : 0,
      ctr: impressions > 0 ? (clicks / impressions) * 100 : 0,
      cpc: clicks > 0 ? spend / clicks : 0,
      cpm: impressions > 0 ? (spend / impressions) * 1000 : 0,
      cost_per_conversion: conversions > 0 ? spend / conversions : 0,
      conversion_rate: clicks > 0 ? (conversions / clicks) * 100 : 0,
      roas: spend > 0 ? revenue / spend : 0
    };
  });
  
  // Sort by day of week
  return analysis.sort((a, b) => a.day_number - b.day_number);
}

/**
 * Generate daily conversion trends
 */
function generateDailyTrends(data) {
  const dailyMap = new Map();
  
  data.forEach(row => {
    const date = row.date;
    const conversions = Number(row.conversions) || 0;
    
    // Account-wide daily trends (only conversions)
    if (!dailyMap.has(date)) {
      dailyMap.set(date, {
        date: date,
        conversions: 0
      });
    }
    dailyMap.get(date).conversions += conversions;
  });
  
  return {
    account: Array.from(dailyMap.values()).sort((a, b) => new Date(a.date) - new Date(b.date))
  };
}

/**
 * Write enhanced reports to sheet
 */
function writeEnhancedReportsToSheet(spreadsheet, reports) {
  try {
    let reportSheet = spreadsheet.getSheetByName(CONFIG.REPORT_SHEET);
    
    // Create sheet if it doesn't exist
    if (!reportSheet) {
      reportSheet = spreadsheet.insertSheet(CONFIG.REPORT_SHEET);
    } else {
      reportSheet.clear();
    }
    
    let currentRow = 1;
    
    // Write revenue configuration section
    currentRow = writeRevenueConfigSection(reportSheet, currentRow);
    currentRow += 2;
    
    // Write campaign summary (aggregated)
    currentRow = writeCampaignSummaryAggregated(reportSheet, reports.campaignSummary, currentRow);
    currentRow += 2;
    
    // Write day of week analysis
    currentRow = writeDayOfWeekAnalysis(reportSheet, reports.dayOfWeekAnalysis, currentRow);
    currentRow += 2;
    
    // Write daily trends summary
    currentRow = writeDailyTrendsSummary(reportSheet, reports.dailyTrends, currentRow);
    
    // Format the report sheet
    formatEnhancedReportSheet(reportSheet);
    
  } catch (error) {
    console.error('Error writing enhanced reports to sheet:', error);
    throw error;
  }
}

/**
 * Write revenue configuration section
 */
function writeRevenueConfigSection(sheet, startRow) {
  sheet.getRange(startRow, 1).setValue('REVENUE CONFIGURATION').setFontWeight('bold').setFontSize(14);
  startRow++;
  
  // Write headers
  sheet.getRange(startRow, 1).setValue('Setting');
  sheet.getRange(startRow, 2).setValue('Value');
  sheet.getRange(startRow, 3).setValue('Description');
  sheet.getRange(startRow, 1, 1, 3).setFontWeight('bold');
  startRow++;
  
  // Write revenue configuration
  sheet.getRange(startRow, 1).setValue('Revenue per Result');
  sheet.getRange(startRow, 2).setValue('$100.00');
  sheet.getRange(startRow, 3).setValue('Amount earned per conversion (edit this value)');
  startRow++;
  
  // Write note
  sheet.getRange(startRow, 1).setValue('Note: Change the value above and re-run the script to update all calculations');
  sheet.getRange(startRow, 1).setFontStyle('italic');
  
  return startRow + 2;
}

/**
 * Write campaign summary aggregated section
 */
function writeCampaignSummaryAggregated(sheet, campaigns, startRow) {
  sheet.getRange(startRow, 1).setValue('CAMPAIGN SUMMARY (AGGREGATED)').setFontWeight('bold').setFontSize(14);
  startRow++;
  
  const headers = [
    'Campaign Name', 'Total Days', 'Total Conversions', 'Avg Daily Conversions', 'Total Spend', 'Total Revenue',
    'CTR (%)', 'CPC ($)', 'CPM ($)', 'Cost per Conversion ($)', 'ROAS'
  ];
  
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  startRow++;
  
  // Write all campaigns
  const campaignData = campaigns.map(campaign => [
    campaign.campaign_name,
    campaign.total_days,
    campaign.conversions,
    campaign.avg_daily_conversions,
    campaign.spend,
    campaign.revenue,
    campaign.ctr,
    campaign.cpc,
    campaign.cpm,
    campaign.cost_per_conversion,
    campaign.roas
  ]);
  
  if (campaignData.length > 0) {
    sheet.getRange(startRow, 1, campaignData.length, headers.length).setValues(campaignData);
  }
  
  return startRow + campaignData.length;
}

/**
 * Write day of week analysis section
 */
function writeDayOfWeekAnalysis(sheet, dayAnalysis, startRow) {
  sheet.getRange(startRow, 1).setValue('DAY OF WEEK ANALYSIS (ACCOUNT-WIDE)').setFontWeight('bold').setFontSize(14);
  startRow++;
  
  const headers = [
    'Day of Week', 'Total Days', 'Total Conversions', 'Avg Daily Conversions', 'Total Spend', 'Avg Daily Spend',
    'CTR (%)', 'CPC ($)', 'CPM ($)', 'Cost per Conversion ($)', 'ROAS'
  ];
  
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  startRow++;
  
  // Write day analysis
  const dayData = dayAnalysis.map(day => [
    day.day_name,
    day.total_days,
    day.conversions,
    day.avg_daily_conversions,
    day.spend,
    day.avg_daily_spend,
    day.ctr,
    day.cpc,
    day.cpm,
    day.cost_per_conversion,
    day.roas
  ]);
  
  if (dayData.length > 0) {
    sheet.getRange(startRow, 1, dayData.length, headers.length).setValues(dayData);
  }
  
  return startRow + dayData.length;
}

/**
 * Write daily trends summary section
 */
function writeDailyTrendsSummary(sheet, dailyTrends, startRow) {
  sheet.getRange(startRow, 1).setValue('DAILY CONVERSION TRENDS').setFontWeight('bold').setFontSize(14);
  startRow++;
  
  // Check if dailyTrends exists and has the expected structure
  if (!dailyTrends || !dailyTrends.account) {
    sheet.getRange(startRow, 1).setValue('No daily trends data available');
    return startRow + 2;
  }
  
  // Account-wide daily trends
  sheet.getRange(startRow, 1).setValue('Account-Wide Daily Conversions:').setFontWeight('bold');
  startRow++;
  
  const accountHeaders = ['Date', 'Total Conversions'];
  sheet.getRange(startRow, 1, 1, 2).setValues([accountHeaders]);
  startRow++;
  
  const accountData = dailyTrends.account.map(day => [day.date, day.conversions]);
  if (accountData.length > 0) {
    sheet.getRange(startRow, 1, accountData.length, 2).setValues(accountData);
  }
  
  startRow += accountData.length + 2;
  
  // Campaign-specific trends (only if campaigns exist)
  if (dailyTrends.campaigns && dailyTrends.campaigns.length > 0) {
    sheet.getRange(startRow, 1).setValue('Campaign-Specific Daily Conversions:').setFontWeight('bold');
    startRow++;
    
    dailyTrends.campaigns.forEach(campaign => {
      sheet.getRange(startRow, 1).setValue(campaign.campaign_name + ':').setFontWeight('bold');
      startRow++;
      
      const campaignHeaders = ['Date', 'Conversions'];
      sheet.getRange(startRow, 1, 1, 2).setValues([campaignHeaders]);
      startRow++;
      
      const campaignData = campaign.daily_data.map(day => [day.date, day.results]);
      if (campaignData.length > 0) {
        sheet.getRange(startRow, 1, campaignData.length, 2).setValues(campaignData);
      }
      
      startRow += campaignData.length + 2;
    });
  }
  
  return startRow;
}

/**
 * Create conversion charts - DISABLED (using createAllChartsSection instead)
 */
function createConversionCharts(spreadsheet, dailyTrends) {
  // This function is disabled - using createAllChartsSection instead
  console.log('createConversionCharts disabled - using createAllChartsSection');
  return;
  try {
    // Create charts sheet
    let chartsSheet = spreadsheet.getSheetByName('Charts');
    if (!chartsSheet) {
      chartsSheet = spreadsheet.insertSheet('Charts');
    } else {
      chartsSheet.clear();
    }
    
    // Account-wide daily conversions chart
    createAccountChart(chartsSheet, dailyTrends.account);
    
    // Campaign-specific charts
    createCampaignCharts(chartsSheet, dailyTrends.campaigns);
    
  } catch (error) {
    console.error('Error creating charts:', error);
  }
}

/**
 * Create account-wide conversion chart
 */
function createAccountChart(sheet, accountData) {
  try {
    // Write data for chart
    sheet.getRange('A1').setValue('Account-Wide Daily Conversions');
    sheet.getRange('A2:B2').setValues([['Date', 'Conversions']]);
    
    const chartData = accountData.map(day => [day.date, day.results]);
    if (chartData.length > 0) {
      sheet.getRange(3, 1, chartData.length, 2).setValues(chartData);
    }
    
    // Create chart
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange(2, 1, chartData.length + 1, 2))
      .setPosition(1, 3, 0, 0)
      .setOption('title', 'Account-Wide Daily Conversions')
      .setOption('width', 600)
      .setOption('height', 400);
    
    sheet.insertChart(chart);
    
  } catch (error) {
    console.error('Error creating account chart:', error);
  }
}

/**
 * Create campaign-specific charts
 */
function createCampaignCharts(sheet, campaigns) {
  try {
    let currentRow = 1;
    
    campaigns.forEach((campaign, index) => {
      const startRow = currentRow + (index * 15);
      
      // Write campaign data
      sheet.getRange(startRow, 1).setValue(campaign.campaign_name + ' - Daily Conversions');
      sheet.getRange(startRow + 1, 1, 1, 2).setValues([['Date', 'Conversions']]);
      
      const chartData = campaign.daily_data.map(day => [day.date, day.results]);
      if (chartData.length > 0) {
        sheet.getRange(startRow + 2, 1, chartData.length, 2).setValues(chartData);
      }
      
      // Create chart
      const chart = sheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(sheet.getRange(startRow + 1, 1, chartData.length + 1, 2))
        .setPosition(startRow, 3, 0, 0)
        .setOption('title', campaign.campaign_name + ' - Daily Conversions')
        .setOption('width', 600)
        .setOption('height', 300);
      
      sheet.insertChart(chart);
      
      currentRow = startRow + 15;
    });
    
  } catch (error) {
    console.error('Error creating campaign charts:', error);
  }
}

/**
 * Format enhanced report sheet
 */
function formatEnhancedReportSheet(sheet) {
  try {
    // Format section headers
    const sectionHeaders = sheet.getRangeList([sheet.getRange('A1'), sheet.getRange('A8'), sheet.getRange('A25'), sheet.getRange('A45')]);
    sectionHeaders.setFontWeight('bold');
    sectionHeaders.setFontSize(14);
    sectionHeaders.setBackground('#f8f9fa');
    
    // Format data headers
    const dataHeaders = sheet.getRangeList([sheet.getRange('A2:A6'), sheet.getRange('A9:A19'), sheet.getRange('A26:A36'), sheet.getRange('A46:A56')]);
    dataHeaders.setFontWeight('bold');
    dataHeaders.setBackground('#e8eaed');
    
    // Format numeric columns
    const numericRanges = sheet.getRangeList(['B3:B6', 'B10:B20', 'B27:B37', 'B47:B57']);
    numericRanges.setNumberFormat('#,##0.00');
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, sheet.getLastColumn());
    
  } catch (error) {
    console.error('Error formatting enhanced report sheet:', error);
  }
}

/**
 * Get conversion types summary
 */
function getConversionTypesSummary(data) {
  const conversionTypeMap = new Map();
  
  data.forEach(row => {
    const resultType = row.result_type;
    const results = Number(row.results) || 0;
    
    // Only count if it's not reach or link clicks
    if (resultType && 
        resultType.toLowerCase() !== 'reach' && 
        resultType.toLowerCase() !== 'link clicks' && 
        results > 0) {
      
      if (!conversionTypeMap.has(resultType)) {
        conversionTypeMap.set(resultType, 0);
      }
      conversionTypeMap.set(resultType, conversionTypeMap.get(resultType) + results);
    }
  });
  
  return Array.from(conversionTypeMap.entries())
    .map(([type, count]) => ({ type, count }))
    .sort((a, b) => b.count - a.count); // Sort by count descending
}

/**
 * Create top campaigns section
 */
function createTopCampaignsSection(sheet, data) {
  try {
    // Aggregate campaign data
    const campaignMap = new Map();
    data.forEach(row => {
      const campaignName = row.campaign_name;
      if (!campaignMap.has(campaignName)) {
        campaignMap.set(campaignName, {
          impressions: 0,
          clicks: 0,
          conversions: 0,
          spend: 0
        });
      }
      const campaign = campaignMap.get(campaignName);
      campaign.impressions += Number(row.impressions) || 0;
      campaign.clicks += Number(row.clicks) || 0;
      campaign.conversions += Number(row.conversions) || 0;
      campaign.spend += Number(row.spend) || 0;
    });
    
    // Calculate metrics for each campaign
    const campaigns = Array.from(campaignMap.entries()).map(([name, data]) => ({
      name: name,
      impressions: data.impressions,
      clicks: data.clicks,
      conversions: data.conversions,
      spend: data.spend,
      cost_per_conversion: data.conversions > 0 ? data.spend / data.conversions : 0
    }));
    
    // Get conversion types summary
    const conversionTypes = getConversionTypesSummary(data);
    
    // Write section header
    sheet.getRange('A1').setValue('TOP CAMPAIGNS');
    sheet.getRange('A1').setFontWeight('bold');
    sheet.getRange('A1').setFontSize(16);
    
    // Write conversion types summary
    sheet.getRange('A2').setValue('CONVERSION TYPES SUMMARY');
    sheet.getRange('A2').setFontWeight('bold');
    sheet.getRange('A2').setFontSize(14);
    
    if (conversionTypes.length > 0) {
      const conversionHeaders = ['Conversion Type', 'Total Count'];
      sheet.getRange(4, 1, 1, conversionHeaders.length).setValues([conversionHeaders]);
      sheet.getRange(4, 1, 1, 2).setFontWeight('bold');
      sheet.getRange(4, 1, 1, 2).setBackground('#e8eaed');
      
      const conversionData = conversionTypes.map(conv => [conv.type, conv.count]);
      sheet.getRange(5, 1, conversionData.length, conversionHeaders.length).setValues(conversionData);
      sheet.getRange(5, 2, conversionData.length, 1).setNumberFormat('#,##0');
      
      const conversionSectionHeight = 6 + conversionData.length;
      
      // Write campaign headers
      const campaignHeaders = ['Campaign Name', 'Cost per Conversion', 'Cost', 'Conversions', 'Clicks', 'Impressions'];
      const campaignHeaderRow = conversionSectionHeight + 2;
      sheet.getRange(campaignHeaderRow, 1, 1, campaignHeaders.length).setValues([campaignHeaders]);
      sheet.getRange(campaignHeaderRow, 1, 1, 6).setFontWeight('bold');
      sheet.getRange(campaignHeaderRow, 1, 1, 6).setBackground('#e8eaed');
      
      // Write campaign data
      const campaignData = campaigns.map(campaign => [
        campaign.name,
        campaign.cost_per_conversion,
        campaign.spend,
        campaign.conversions,
        campaign.clicks,
        campaign.impressions
      ]);
      
      if (campaignData.length > 0) {
        sheet.getRange(conversionSectionHeight + 3, 1, campaignData.length, campaignHeaders.length).setValues(campaignData);
      }
      
      // Format numeric columns
      const startRow = conversionSectionHeight + 3;
      const endRow = conversionSectionHeight + 2 + campaignData.length;
      sheet.getRange(startRow, 2, campaignData.length, 1).setNumberFormat('$#,##0.00'); // Cost per Conversion
      sheet.getRange(startRow, 3, campaignData.length, 1).setNumberFormat('$#,##0.00'); // Cost
      sheet.getRange(startRow, 4, campaignData.length, 3).setNumberFormat('#,##0'); // Conversions, Clicks, Impressions
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, Math.max(conversionHeaders.length, campaignHeaders.length));
      
      return conversionSectionHeight + 4 + campaignData.length; // Return next row position
    } else {
      // No conversion types found, just show campaigns
      const campaignHeaders = ['Campaign Name', 'Cost per Conversion', 'Cost', 'Conversions', 'Clicks', 'Impressions'];
      sheet.getRange(4, 1, 1, campaignHeaders.length).setValues([campaignHeaders]);
      sheet.getRange(4, 1, 1, 6).setFontWeight('bold');
      sheet.getRange(4, 1, 1, 6).setBackground('#e8eaed');
      
      const campaignData = campaigns.map(campaign => [
        campaign.name,
        campaign.cost_per_conversion,
        campaign.spend,
        campaign.conversions,
        campaign.clicks,
        campaign.impressions
      ]);
      
      if (campaignData.length > 0) {
        sheet.getRange(5, 1, campaignData.length, campaignHeaders.length).setValues(campaignData);
      }
      
      // Format numeric columns
      sheet.getRange(5, 2, campaignData.length, 1).setNumberFormat('$#,##0.00'); // Cost per Conversion
      sheet.getRange(5, 3, campaignData.length, 1).setNumberFormat('$#,##0.00'); // Cost
      sheet.getRange(5, 4, campaignData.length, 3).setNumberFormat('#,##0'); // Conversions, Clicks, Impressions
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, campaignHeaders.length);
      
      return 6 + campaignData.length; // Return next row position
    }
    
  } catch (error) {
    console.error('Error creating top campaigns section:', error);
    return 10;
  }
}

/**
 * Create pivot table tab with aggregated daily data
 */
function createPivotTableTab(spreadsheet, data) {
  try {
    console.log('Creating pivot table tab...');
    
    // Create or get the pivot table sheet
    let pivotSheet = spreadsheet.getSheetByName('Pivot Table');
    if (!pivotSheet) {
      pivotSheet = spreadsheet.insertSheet('Pivot Table');
    } else {
      pivotSheet.clear();
    }
    
    // Aggregate daily data across ALL campaigns
    const dailyMap = new Map();
    
    console.log('Starting aggregation with', data.length, 'rows');
    console.log('Sample data rows:', data.slice(0, 3));
    
    data.forEach((row, index) => {
      const date = row.date;
      // Convert date to consistent string format (YYYY-MM-DD)
      const dateString = date instanceof Date ? 
        date.getFullYear() + '-' + 
        String(date.getMonth() + 1).padStart(2, '0') + '-' + 
        String(date.getDate()).padStart(2, '0') : 
        String(date);
      
      console.log(`Row ${index}: Date = "${dateString}", Original = "${date}", Type = ${typeof date}`);
      
      if (!dailyMap.has(dateString)) {
        dailyMap.set(dateString, {
          impressions: 0,
          clicks: 0,
          conversions: 0,
          spend: 0,
          reach: 0
        });
        console.log(`Created new entry for date: ${dateString}`);
      }
      
      const day = dailyMap.get(dateString);
      const impressions = Number(row.impressions) || 0;
      const clicks = Number(row.clicks) || 0;
      const conversions = Number(row.conversions) || 0;
      const spend = Number(row.spend) || 0;
      const reach = Number(row.reach) || 0;
      
      day.impressions += impressions;
      day.clicks += clicks;
      day.conversions += conversions;
      day.spend += spend;
      day.reach += reach;
      
      if (index < 5) {
        console.log(`Row ${index}: Added ${impressions} impressions, ${clicks} clicks, ${conversions} conversions, $${spend} spend to date ${dateString}`);
      }
    });
    
    console.log('Daily aggregation complete. Unique dates:', dailyMap.size);
    console.log('All dates:', Array.from(dailyMap.keys()));
    console.log('Sample daily data:', Array.from(dailyMap.entries()).slice(0, 3));
    
    // Calculate daily metrics and sort by date
    const pivotData = Array.from(dailyMap.entries()).map(([dateString, data]) => ({
      date: dateString,
      impressions: data.impressions,
      clicks: data.clicks,
      conversions: data.conversions,
      spend: data.spend,
      reach: data.reach,
      ctr: data.impressions > 0 ? (data.clicks / data.impressions) * 100 : 0,
      conversion_rate: data.clicks > 0 ? (data.conversions / data.clicks) * 100 : 0,
      cpc: data.clicks > 0 ? data.spend / data.clicks : 0,
      cpm: data.impressions > 0 ? (data.spend / data.impressions) * 1000 : 0,
      cost_per_conversion: data.conversions > 0 ? data.spend / data.conversions : 0
    })).sort((a, b) => new Date(a.date) - new Date(b.date));
    
    console.log('Pivot data created:', pivotData.length, 'days');
    console.log('Sample pivot data:', pivotData.slice(0, 3));
    
    // Write headers
    const headers = [
      'Date', 'Impressions', 'Clicks', 'Conversions', 'Spend', 'Reach',
      'CTR (%)', 'Conversion Rate (%)', 'CPC ($)', 'CPM ($)', 'Cost per Conversion ($)'
    ];
    pivotSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    pivotSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    pivotSheet.getRange(1, 1, 1, headers.length).setBackground('#e8eaed');
    
    // Write data
    const dataRows = pivotData.map(day => [
      day.date,
      day.impressions,
      day.clicks,
      day.conversions,
      day.spend,
      day.reach,
      day.ctr / 100, // Convert to decimal for percentage formatting
      day.conversion_rate / 100, // Convert to decimal for percentage formatting
      day.cpc,
      day.cpm,
      day.cost_per_conversion
    ]);
    
    if (dataRows.length > 0) {
      pivotSheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    }
    
    // Format columns
    pivotSheet.getRange(2, 1, dataRows.length, 1).setNumberFormat('yyyy-mm-dd'); // Date
    pivotSheet.getRange(2, 2, dataRows.length, 4).setNumberFormat('#,##0'); // Impressions, Clicks, Conversions, Spend
    pivotSheet.getRange(2, 6, dataRows.length, 1).setNumberFormat('#,##0'); // Reach
    pivotSheet.getRange(2, 7, dataRows.length, 2).setNumberFormat('0.00%'); // CTR, Conversion Rate
    pivotSheet.getRange(2, 9, dataRows.length, 3).setNumberFormat('$#,##0.00'); // CPC, CPM, Cost per Conversion
    
    // Auto-resize columns
    pivotSheet.autoResizeColumns(1, headers.length);
    
    // Add totals row
    const totalRow = dataRows.length + 2;
    pivotSheet.getRange(totalRow, 1).setValue('TOTALS');
    pivotSheet.getRange(totalRow, 1).setFontWeight('bold');
    pivotSheet.getRange(totalRow, 1).setBackground('#f8f9fa');
    
    // Calculate totals
    const totals = pivotData.reduce((acc, day) => ({
      impressions: acc.impressions + day.impressions,
      clicks: acc.clicks + day.clicks,
      conversions: acc.conversions + day.conversions,
      spend: acc.spend + day.spend,
      reach: acc.reach + day.reach
    }), { impressions: 0, clicks: 0, conversions: 0, spend: 0, reach: 0 });
    
    // Overall metrics
    const overallCTR = totals.impressions > 0 ? (totals.clicks / totals.impressions) * 100 : 0;
    const overallConversionRate = totals.clicks > 0 ? (totals.conversions / totals.clicks) * 100 : 0;
    const overallCPC = totals.clicks > 0 ? totals.spend / totals.clicks : 0;
    const overallCPM = totals.impressions > 0 ? (totals.spend / totals.impressions) * 1000 : 0;
    const overallCostPerConversion = totals.conversions > 0 ? totals.spend / totals.conversions : 0;
    
    // Write totals
    const totalsRow = [
      '', // Empty date cell
      totals.impressions,
      totals.clicks,
      totals.conversions,
      totals.spend,
      totals.reach,
      overallCTR / 100,
      overallConversionRate / 100,
      overallCPC,
      overallCPM,
      overallCostPerConversion
    ];
    
    pivotSheet.getRange(totalRow, 2, 1, totalsRow.length - 1).setValues([totalsRow.slice(1)]);
    pivotSheet.getRange(totalRow, 2, 1, 4).setNumberFormat('#,##0'); // Totals
    pivotSheet.getRange(totalRow, 6, 1, 1).setNumberFormat('#,##0'); // Reach total
    pivotSheet.getRange(totalRow, 7, 1, 2).setNumberFormat('0.00%'); // Overall rates
    pivotSheet.getRange(totalRow, 9, 1, 3).setNumberFormat('$#,##0.00'); // Overall costs
    
    console.log('Pivot table created successfully');
    return pivotData; // Return the data for use in charts
    
  } catch (error) {
    console.error('Error creating pivot table:', error);
    return [];
  }
}

/**
 * Create all charts section with three horizontal charts
 */
function createAllChartsSection(sheet, data) {
  try {
    console.log('Starting chart creation...');
    
    // Get pivot table data
    const pivotSheet = sheet.getParent().getSheetByName('Pivot Table');
    if (!pivotSheet) {
      console.error('Pivot table not found');
      return;
    }
    
    // Read data from pivot table
    const pivotRange = pivotSheet.getDataRange();
    const pivotValues = pivotRange.getValues();
    const headers = pivotValues[0];
    const dataRows = pivotValues.slice(1, -1); // Exclude header and totals row
    
    console.log('Pivot table headers:', headers);
    console.log('Pivot table data rows:', dataRows.length);
    
    if (dataRows.length === 0) {
      console.error('No data rows found in pivot table');
      return;
    }
    
    // Find column indices
    const dateIndex = headers.indexOf('Date');
    const clicksIndex = headers.indexOf('Clicks');
    const ctrIndex = headers.indexOf('CTR (%)');
    const conversionsIndex = headers.indexOf('Conversions');
    const conversionRateIndex = headers.indexOf('Conversion Rate (%)');
    const costIndex = headers.indexOf('Spend');
    const cpcIndex = headers.indexOf('CPC ($)');
    
    console.log('Column indices:', {
      date: dateIndex,
      clicks: clicksIndex,
      ctr: ctrIndex,
      conversions: conversionsIndex,
      conversionRate: conversionRateIndex,
      cost: costIndex,
      cpc: cpcIndex
    });
    
    // Calculate overall metrics from pivot data
    const totalClicks = dataRows.reduce((sum, row) => sum + (Number(row[clicksIndex]) || 0), 0);
    const totalConversions = dataRows.reduce((sum, row) => sum + (Number(row[conversionsIndex]) || 0), 0);
    const totalCost = dataRows.reduce((sum, row) => sum + (Number(row[costIndex]) || 0), 0);
    const overallCTR = dataRows.reduce((sum, row) => sum + (Number(row[ctrIndex]) || 0), 0) / dataRows.length;
    const overallConversionRate = dataRows.reduce((sum, row) => sum + (Number(row[conversionRateIndex]) || 0), 0) / dataRows.length;
    const overallCPC = dataRows.reduce((sum, row) => sum + (Number(row[cpcIndex]) || 0), 0) / dataRows.length;
    
    // Start position (after top campaigns section)
    const startRow = 30;
    
    // Write section header
    sheet.getRange(startRow, 1).setValue('ACCOUNT PERFORMANCE CHARTS');
    sheet.getRange(startRow, 1).setFontWeight('bold');
    sheet.getRange(startRow, 1).setFontSize(16);
    
    // Write metrics summary
    sheet.getRange(startRow + 2, 1).setValue('Summary Metrics:');
    sheet.getRange(startRow + 2, 1).setFontWeight('bold');
    sheet.getRange(startRow + 3, 1).setValue('Total Clicks:');
    sheet.getRange(startRow + 3, 2).setValue(totalClicks);
    sheet.getRange(startRow + 4, 1).setValue('Overall CTR:');
    sheet.getRange(startRow + 4, 2).setValue(overallCTR / 100);
    sheet.getRange(startRow + 5, 1).setValue('Total Conversions:');
    sheet.getRange(startRow + 5, 2).setValue(totalConversions);
    sheet.getRange(startRow + 6, 1).setValue('Conversion Rate:');
    sheet.getRange(startRow + 6, 2).setValue(overallConversionRate / 100);
    sheet.getRange(startRow + 7, 1).setValue('Total Cost:');
    sheet.getRange(startRow + 7, 2).setValue(totalCost);
    sheet.getRange(startRow + 8, 1).setValue('Average CPC:');
    sheet.getRange(startRow + 8, 2).setValue(overallCPC);
    
    // Format metrics
    sheet.getRange(startRow + 3, 2).setNumberFormat('#,##0');
    sheet.getRange(startRow + 4, 2).setNumberFormat('0.00%');
    sheet.getRange(startRow + 5, 2).setNumberFormat('#,##0');
    sheet.getRange(startRow + 6, 2).setNumberFormat('0.00%');
    sheet.getRange(startRow + 7, 2).setNumberFormat('$#,##0.00');
    sheet.getRange(startRow + 8, 2).setNumberFormat('$#,##0.00');
    
    // Clear any existing charts first
    const existingCharts = sheet.getCharts();
    existingCharts.forEach(chart => sheet.removeChart(chart));
    
    // Create charts using pivot table data
    console.log('Creating exactly 3 charts with', dataRows.length, 'data points');
    
    // Chart 1: Clicks & CTR (Date=1, Clicks=3, CTR=7)
    const chart1 = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(pivotSheet.getRange(2, 1, dataRows.length, 1)) // Date column
      .addRange(pivotSheet.getRange(2, 3, dataRows.length, 1)) // Clicks column
      .addRange(pivotSheet.getRange(2, 7, dataRows.length, 1)) // CTR column
      .setPosition(startRow + 2, 6, 0, 0)
      .setOption('title', 'Daily Clicks & CTR')
      .setOption('width', 400)
      .setOption('height', 250)
      .setOption('series', {
        0: {targetAxisIndex: 0, color: '#4285f4', labelInLegend: 'Clicks'},
        1: {targetAxisIndex: 1, color: '#ea4335', labelInLegend: 'CTR (%)'}
      })
      .setOption('vAxes', {
        0: {title: 'Clicks', minValue: 0},
        1: {title: 'CTR (%)', format: 'percent', minValue: 0}
      })
      .setOption('hAxis', {title: 'Date'})
      .setOption('legend', {position: 'top'})
      .build();
    
    // Chart 2: Conversions & Conversion Rate (Date=1, Conversions=4, Conversion Rate=8)
    const chart2 = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(pivotSheet.getRange(2, 1, dataRows.length, 1)) // Date column
      .addRange(pivotSheet.getRange(2, 4, dataRows.length, 1)) // Conversions column
      .addRange(pivotSheet.getRange(2, 8, dataRows.length, 1)) // Conversion Rate column
      .setPosition(startRow + 2, 12, 0, 0)
      .setOption('title', 'Daily Conversions & Conversion Rate')
      .setOption('width', 400)
      .setOption('height', 250)
      .setOption('series', {
        0: {targetAxisIndex: 0, color: '#34a853', labelInLegend: 'Conversions'},
        1: {targetAxisIndex: 1, color: '#fbbc04', labelInLegend: 'Conversion Rate (%)'}
      })
      .setOption('vAxes', {
        0: {title: 'Conversions', minValue: 0},
        1: {title: 'Conversion Rate (%)', format: 'percent', minValue: 0}
      })
      .setOption('hAxis', {title: 'Date'})
      .setOption('legend', {position: 'top'})
      .build();
    
    // Chart 3: Cost & CPC (Date=1, Spend=5, CPC=9)
    const chart3 = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(pivotSheet.getRange(2, 1, dataRows.length, 1)) // Date column
      .addRange(pivotSheet.getRange(2, 5, dataRows.length, 1)) // Spend column
      .addRange(pivotSheet.getRange(2, 9, dataRows.length, 1)) // CPC column
      .setPosition(startRow + 2, 18, 0, 0)
      .setOption('title', 'Daily Cost & CPC')
      .setOption('width', 400)
      .setOption('height', 250)
      .setOption('series', {
        0: {targetAxisIndex: 0, color: '#9c27b0', labelInLegend: 'Cost ($)'},
        1: {targetAxisIndex: 1, color: '#ff9800', labelInLegend: 'CPC ($)'}
      })
      .setOption('vAxes', {
        0: {title: 'Cost ($)', minValue: 0, format: 'currency'},
        1: {title: 'CPC ($)', minValue: 0, format: 'currency'}
      })
      .setOption('hAxis', {title: 'Date'})
      .setOption('legend', {position: 'top'})
      .build();
    
    // Insert exactly 3 charts
    sheet.insertChart(chart1);
    sheet.insertChart(chart2);
    sheet.insertChart(chart3);
    
    console.log('Exactly 3 charts created successfully');
    
    return startRow + 35;
    
  } catch (error) {
    console.error('Error creating all charts section:', error);
    return 80;
  }
}

/**
 * Create clicks and CTR chart section - DISABLED (using createAllChartsSection instead)
 */
function createClicksAndCTRChart(sheet, data) {
  // This function is disabled - using createAllChartsSection instead
  console.log('createClicksAndCTRChart disabled - using createAllChartsSection');
  return;
  try {
    // Aggregate daily data across ALL campaigns
    const dailyMap = new Map();
    data.forEach(row => {
      const date = row.date;
      if (!dailyMap.has(date)) {
        dailyMap.set(date, {
          impressions: 0,
          clicks: 0
        });
      }
      const day = dailyMap.get(date);
      day.impressions += Number(row.impressions) || 0;
      day.clicks += Number(row.clicks) || 0;
    });
    
    // Calculate daily CTR and overall metrics
    const dailyData = Array.from(dailyMap.entries()).map(([date, data]) => ({
      date: date,
      clicks: data.clicks,
      ctr: data.impressions > 0 ? (data.clicks / data.impressions) * 100 : 0
    })).sort((a, b) => new Date(a.date) - new Date(b.date));
    
    console.log('Daily data for clicks chart:', dailyData.slice(0, 5)); // Debug first 5 days
    
    const totalClicks = dailyData.reduce((sum, day) => sum + day.clicks, 0);
    const totalImpressions = data.reduce((sum, row) => sum + (Number(row.impressions) || 0), 0);
    const overallCTR = totalImpressions > 0 ? (totalClicks / totalImpressions) * 100 : 0;
    
    // Start position (after top campaigns section)
    const startRow = 30;
    
    // Write section header
    sheet.getRange(startRow, 1).setValue('CLICKS & CTR PERFORMANCE');
    sheet.getRange(startRow, 1).setFontWeight('bold');
    sheet.getRange(startRow, 1).setFontSize(16);
    
    // Write metrics next to chart
    sheet.getRange(startRow + 2, 4).setValue('Total Clicks:');
    sheet.getRange(startRow + 2, 5).setValue(totalClicks);
    sheet.getRange(startRow + 3, 4).setValue('Overall CTR:');
    sheet.getRange(startRow + 3, 5).setValue(overallCTR / 100);
    sheet.getRange(startRow + 4, 4).setValue('Total Impressions:');
    sheet.getRange(startRow + 4, 5).setValue(totalImpressions);
    
    // Format metrics
    sheet.getRange(startRow + 2, 5).setNumberFormat('#,##0');
    sheet.getRange(startRow + 3, 5).setNumberFormat('0.00%');
    sheet.getRange(startRow + 4, 5).setNumberFormat('#,##0');
    
    // Create chart data (visible for debugging)
    const chartData = dailyData.map(day => [day.date, day.clicks, day.ctr / 100]); // Convert percentage to decimal
    console.log('Chart data for clicks:', chartData.slice(0, 5)); // Debug first 5 rows
    
    if (chartData.length > 0) {
      // Write data in a visible area for the chart
      const dataRow = startRow + 20;
      sheet.getRange(dataRow, 1, 1, 3).setValues([['Date', 'Clicks', 'CTR']]);
      sheet.getRange(dataRow + 1, 1, chartData.length, 3).setValues(chartData);
      console.log('Wrote chart data to rows', dataRow, 'to', dataRow + chartData.length);
    }
    
    // Create line chart - positioned horizontally
    const chartRange = sheet.getRange(startRow + 20, 1, chartData.length + 1, 3);
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartRange)
      .setPosition(startRow + 2, 6, 0, 0)
      .setOption('title', 'Account-Wide Daily Clicks & Click-Through Rate')
      .setOption('width', 400)
      .setOption('height', 250)
      .setOption('series', {
        0: {targetAxisIndex: 0, color: '#4285f4', labelInLegend: 'Clicks'}, // Clicks on primary axis
        1: {targetAxisIndex: 1, color: '#ea4335', labelInLegend: 'CTR (%)'}  // CTR on secondary axis
      })
      .setOption('vAxes', {
        0: {title: 'Clicks', minValue: 0},
        1: {title: 'CTR (%)', format: 'percent', minValue: 0}
      })
      .setOption('hAxis', {title: 'Date'})
      .setOption('legend', {position: 'top'})
      .build();
    
    sheet.insertChart(chart);
    
    return startRow + 35; // Return next row position with more spacing
    
  } catch (error) {
    console.error('Error creating clicks and CTR chart:', error);
    return 80;
  }
}

/**
 * Create conversions and rate chart section - DISABLED (using createAllChartsSection instead)
 */
function createConversionsAndRateChart(sheet, data) {
  // This function is disabled - using createAllChartsSection instead
  console.log('createConversionsAndRateChart disabled - using createAllChartsSection');
  return;
  try {
    // Aggregate daily data across ALL campaigns
    const dailyMap = new Map();
    data.forEach(row => {
      const date = row.date;
      if (!dailyMap.has(date)) {
        dailyMap.set(date, {
          conversions: 0,
          clicks: 0
        });
      }
      const day = dailyMap.get(date);
      day.conversions += Number(row.conversions) || 0;
      day.clicks += Number(row.clicks) || 0;
    });
    
    // Calculate daily conversion rate and overall metrics
    const dailyData = Array.from(dailyMap.entries()).map(([date, data]) => ({
      date: date,
      conversions: data.conversions,
      conversion_rate: data.clicks > 0 ? (data.conversions / data.clicks) * 100 : 0
    })).sort((a, b) => new Date(a.date) - new Date(b.date));
    
    const totalConversions = dailyData.reduce((sum, day) => sum + day.conversions, 0);
    const totalClicks = data.reduce((sum, row) => sum + (Number(row.clicks) || 0), 0);
    const overallConversionRate = totalClicks > 0 ? (totalConversions / totalClicks) * 100 : 0;
    const totalSpend = data.reduce((sum, row) => sum + (Number(row.spend) || 0), 0);
    const costPerConversion = totalConversions > 0 ? totalSpend / totalConversions : 0;
    
    // Start position (after clicks section) - positioned horizontally
    const startRow = 80;
    
    // Write section header
    sheet.getRange(startRow, 1).setValue('CONVERSIONS & CONVERSION RATE');
    sheet.getRange(startRow, 1).setFontWeight('bold');
    sheet.getRange(startRow, 1).setFontSize(16);
    
    // Write metrics next to chart
    sheet.getRange(startRow + 2, 4).setValue('Total Conversions:');
    sheet.getRange(startRow + 2, 5).setValue(totalConversions);
    sheet.getRange(startRow + 3, 4).setValue('Conversion Rate:');
    sheet.getRange(startRow + 3, 5).setValue(overallConversionRate / 100);
    sheet.getRange(startRow + 4, 4).setValue('Cost per Conversion:');
    sheet.getRange(startRow + 4, 5).setValue(costPerConversion);
    
    // Format metrics
    sheet.getRange(startRow + 2, 5).setNumberFormat('#,##0');
    sheet.getRange(startRow + 3, 5).setNumberFormat('0.00%');
    sheet.getRange(startRow + 4, 5).setNumberFormat('$#,##0.00');
    
    // Create chart data (visible for debugging)
    const chartData = dailyData.map(day => [day.date, day.conversions, day.conversion_rate / 100]); // Convert percentage to decimal
    if (chartData.length > 0) {
      // Write data in a visible area for the chart
      const dataRow = startRow + 20;
      sheet.getRange(dataRow, 1, 1, 3).setValues([['Date', 'Conversions', 'Conversion Rate']]);
      sheet.getRange(dataRow + 1, 1, chartData.length, 3).setValues(chartData);
    }
    
    // Create line chart - positioned horizontally
    const chartRange = sheet.getRange(startRow + 20, 1, chartData.length + 1, 3);
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartRange)
      .setPosition(startRow + 2, 6, 0, 0) // Position to the right of first chart
      .setOption('title', 'Account-Wide Daily Conversions & Conversion Rate')
      .setOption('width', 400)
      .setOption('height', 250)
      .setOption('series', {
        0: {targetAxisIndex: 0, color: '#34a853', labelInLegend: 'Conversions'}, // Conversions on primary axis
        1: {targetAxisIndex: 1, color: '#fbbc04', labelInLegend: 'Conversion Rate (%)'}  // Conversion rate on secondary axis
      })
      .setOption('vAxes', {
        0: {title: 'Conversions', minValue: 0},
        1: {title: 'Conversion Rate (%)', format: 'percent', minValue: 0}
      })
      .setOption('hAxis', {title: 'Date'})
      .setOption('legend', {position: 'top'})
      .build();
    
    sheet.insertChart(chart);
    
    return startRow + 35; // Return next row position with more spacing
    
  } catch (error) {
    console.error('Error creating conversions and rate chart:', error);
    return 80;
  }
}

/**
 * Create cost and CPC chart section - DISABLED (using createAllChartsSection instead)
 */
function createCostAndCPCChart(sheet, data) {
  // This function is disabled - using createAllChartsSection instead
  console.log('createCostAndCPCChart disabled - using createAllChartsSection');
  return;
  try {
    // Aggregate daily data across ALL campaigns
    const dailyMap = new Map();
    data.forEach(row => {
      const date = row.date;
      if (!dailyMap.has(date)) {
        dailyMap.set(date, {
          cost: 0,
          clicks: 0
        });
      }
      const day = dailyMap.get(date);
      day.cost += Number(row.spend) || 0;
      day.clicks += Number(row.clicks) || 0;
    });
    
    // Calculate daily CPC and overall metrics
    const dailyData = Array.from(dailyMap.entries()).map(([date, data]) => ({
      date: date,
      cost: data.cost,
      cpc: data.clicks > 0 ? data.cost / data.clicks : 0
    })).sort((a, b) => new Date(a.date) - new Date(b.date));
    
    const totalCost = dailyData.reduce((sum, day) => sum + day.cost, 0);
    const totalClicks = data.reduce((sum, row) => sum + (Number(row.clicks) || 0), 0);
    const overallCPC = totalClicks > 0 ? totalCost / totalClicks : 0;
    const totalImpressions = data.reduce((sum, row) => sum + (Number(row.impressions) || 0), 0);
    const overallCPM = totalImpressions > 0 ? (totalCost / totalImpressions) * 1000 : 0;
    
    // Start position (after conversions section) - positioned horizontally
    const startRow = 130;
    
    // Write section header
    sheet.getRange(startRow, 1).setValue('COST & CPC PERFORMANCE');
    sheet.getRange(startRow, 1).setFontWeight('bold');
    sheet.getRange(startRow, 1).setFontSize(16);
    
    // Write metrics next to chart
    sheet.getRange(startRow + 2, 4).setValue('Total Cost:');
    sheet.getRange(startRow + 2, 5).setValue(totalCost);
    sheet.getRange(startRow + 3, 4).setValue('Average CPC:');
    sheet.getRange(startRow + 3, 5).setValue(overallCPC);
    sheet.getRange(startRow + 4, 4).setValue('Average CPM:');
    sheet.getRange(startRow + 4, 5).setValue(overallCPM);
    
    // Format metrics
    sheet.getRange(startRow + 2, 5).setNumberFormat('$#,##0.00');
    sheet.getRange(startRow + 3, 5).setNumberFormat('$#,##0.00');
    sheet.getRange(startRow + 4, 5).setNumberFormat('$#,##0.00');
    
    // Create chart data (visible for debugging)
    const chartData = dailyData.map(day => [day.date, day.cost, day.cpc]);
    if (chartData.length > 0) {
      // Write data in a visible area for the chart
      const dataRow = startRow + 20;
      sheet.getRange(dataRow, 1, 1, 3).setValues([['Date', 'Cost', 'CPC']]);
      sheet.getRange(dataRow + 1, 1, chartData.length, 3).setValues(chartData);
    }
    
    // Create line chart - positioned horizontally
    const chartRange = sheet.getRange(startRow + 20, 1, chartData.length + 1, 3);
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartRange)
      .setPosition(startRow + 2, 6, 0, 0) // Position to the right of second chart
      .setOption('title', 'Account-Wide Daily Cost & Cost Per Click')
      .setOption('width', 400)
      .setOption('height', 250)
      .setOption('series', {
        0: {targetAxisIndex: 0, color: '#9c27b0', labelInLegend: 'Cost ($)'}, // Cost on primary axis
        1: {targetAxisIndex: 1, color: '#ff9800', labelInLegend: 'CPC ($)'}  // CPC on secondary axis
      })
      .setOption('vAxes', {
        0: {title: 'Cost ($)', minValue: 0, format: 'currency'},
        1: {title: 'CPC ($)', minValue: 0, format: 'currency'}
      })
      .setOption('hAxis', {title: 'Date'})
      .setOption('legend', {position: 'top'})
      .build();
    
    sheet.insertChart(chart);
    
    return startRow + 35; // Return next row position with more spacing
    
  } catch (error) {
    console.error('Error creating cost and CPC chart:', error);
    return 80;
  }
}

/**
 * Debug calculations for your data
 */
function debugCalculations() {
  try {
    console.log('=== CALCULATION DEBUG START ===');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const rawDataSheet = spreadsheet.getSheetByName('Raw Data');
    
    if (!rawDataSheet) {
      console.log('ERROR: Raw Data sheet not found');
      return;
    }
    
    const rawData = rawDataSheet.getDataRange().getValues();
    const headers = rawData[0];
    const dataRows = rawData.slice(1);
    
    console.log('Headers found:', headers);
    console.log('Number of data rows:', dataRows.length);
    
    // Process first few rows to debug
    for (let i = 0; i < Math.min(3, dataRows.length); i++) {
      const row = dataRows[i];
      console.log(`\n--- ROW ${i + 1} DEBUG ---`);
      
      // Find column indices
      const impressionsIndex = headers.findIndex(h => h === 'Impressions');
      const clicksIndex = headers.findIndex(h => h === 'Link clicks');
      const resultsIndex = headers.findIndex(h => h === 'Results');
      const spendIndex = headers.findIndex(h => h === 'Amount spent (USD)');
      
      console.log('Column indices:', {
        impressions: impressionsIndex,
        clicks: clicksIndex,
        results: resultsIndex,
        spend: spendIndex
      });
      
      // Get raw values
      const impressions = Number(row[impressionsIndex]) || 0;
      const clicks = Number(row[clicksIndex]) || 0;
      const results = Number(row[resultsIndex]) || 0;
      const spend = Number(row[spendIndex]) || 0;
      
      console.log('Raw values:', {
        impressions: impressions,
        clicks: clicks,
        results: results,
        spend: spend
      });
      
      // Calculate metrics
      const ctr = impressions > 0 ? (clicks / impressions) * 100 : 0;
      const conversionRate = clicks > 0 ? (results / clicks) * 100 : 0;
      const cpc = clicks > 0 ? spend / clicks : 0;
      const cpm = impressions > 0 ? (spend / impressions) * 1000 : 0;
      const revenuePerResult = getRevenuePerResult(spreadsheet);
      const estimatedRevenue = results * revenuePerResult;
      const roas = spend > 0 ? estimatedRevenue / spend : 0;
      
      console.log('Calculated metrics:', {
        ctr: ctr + '%',
        conversionRate: conversionRate + '%',
        cpc: '$' + cpc,
        cpm: '$' + cpm,
        revenuePerResult: '$' + revenuePerResult,
        estimatedRevenue: '$' + estimatedRevenue,
        roas: roas
      });
      
      // Show the math
      console.log('Math breakdown:');
      console.log(`CTR = (${clicks} / ${impressions}) * 100 = ${ctr}%`);
      console.log(`Conversion Rate = (${results} / ${clicks}) * 100 = ${conversionRate}%`);
      console.log(`CPC = ${spend} / ${clicks} = $${cpc}`);
      console.log(`CPM = (${spend} / ${impressions}) * 1000 = $${cpm}`);
      console.log(`Revenue per Result = $${revenuePerResult}`);
      console.log(`Estimated Revenue = ${results} * $${revenuePerResult} = $${estimatedRevenue}`);
      console.log(`ROAS = ${estimatedRevenue} / ${spend} = ${roas}`);
    }
    
    console.log('\n=== CALCULATION DEBUG END ===');
    
  } catch (error) {
    console.error('Error in debug calculations:', error);
  }
}

/**
 * Debug data aggregation
 */
function debugDataAggregation() {
  try {
    console.log('=== DEBUGGING DATA AGGREGATION ===');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const data = processRawData(spreadsheet);
    
    console.log('Total data rows:', data.length);
    
    // Group by date to see what we have
    const dateGroups = {};
    data.forEach((row, index) => {
      const date = row.date;
      if (!dateGroups[date]) {
        dateGroups[date] = [];
      }
      dateGroups[date].push(row);
    });
    
    console.log('Unique dates found:', Object.keys(dateGroups).length);
    console.log('Dates:', Object.keys(dateGroups).sort());
    
    // Show first few dates with their row counts
    const sortedDates = Object.keys(dateGroups).sort();
    sortedDates.slice(0, 5).forEach(date => {
      console.log(`Date ${date}: ${dateGroups[date].length} rows`);
      console.log('Sample rows:', dateGroups[date].slice(0, 2));
    });
    
    console.log('=== END DEBUGGING ===');
  } catch (error) {
    console.error('Error debugging data aggregation:', error);
  }
}

/**
 * Test pivot table creation only
 */
function testPivotTable() {
  try {
    console.log('Testing pivot table creation...');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Delete existing pivot table if it exists
    const existingPivot = spreadsheet.getSheetByName('Pivot Table');
    if (existingPivot) {
      spreadsheet.deleteSheet(existingPivot);
      console.log('Deleted existing pivot table');
    }
    
    const data = processRawData(spreadsheet);
    console.log('Processed data length:', data.length);
    console.log('Sample processed data:', data.slice(0, 3));
    
    createPivotTableTab(spreadsheet, data);
    console.log('Pivot table test completed');
  } catch (error) {
    console.error('Error testing pivot table:', error);
  }
}

/**
 * Enhanced manual process function
 */
function enhancedManualProcess() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear existing dashboard
  const dashboardSheet = spreadsheet.getSheetByName('Dashboard');
  if (dashboardSheet) {
    dashboardSheet.clear();
    console.log('Cleared existing dashboard');
  }
  
  // Process raw data
  const processedData = processRawData(spreadsheet);
  
  // Generate enhanced reports
  generateEnhancedReports(spreadsheet, processedData);
  
  // Update dashboard
  updateDashboard(spreadsheet, processedData);
  
  console.log('Enhanced processing completed successfully');
}

/**
 * Debug invalid rows to see what's missing
 */
function debugInvalidRows() {
  try {
    console.log('=== DEBUGGING INVALID ROWS ===');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const rawDataSheet = spreadsheet.getSheetByName('Raw Data');
    
    if (!rawDataSheet) {
      console.log('ERROR: Raw Data sheet not found');
      return;
    }
    
    const rawData = rawDataSheet.getDataRange().getValues();
    const headers = rawData[0];
    const dataRows = rawData.slice(1);
    
    console.log('Headers found:', headers);
    console.log('Total rows:', dataRows.length);
    
    // Check first few invalid rows
    let invalidCount = 0;
    for (let i = 0; i < Math.min(10, dataRows.length); i++) {
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
      
      // Check if row is valid
      const requiredFields = ['campaign_name', 'date'];
      const missingFields = requiredFields.filter(field => !cleanedRow[field] || cleanedRow[field] === '');
      
      if (missingFields.length > 0) {
        invalidCount++;
        console.log(`\n--- INVALID ROW ${i + 2} ---`);
        console.log('Row data:', row);
        console.log('Cleaned row:', cleanedRow);
        console.log('Missing fields:', missingFields);
        console.log('Available fields:', Object.keys(cleanedRow));
        
        if (invalidCount >= 5) break; // Only show first 5 invalid rows
      }
    }
    
    console.log(`\nFound ${invalidCount} invalid rows in first 10 rows`);
    console.log('=== END DEBUGGING ===');
    
  } catch (error) {
    console.error('Error debugging invalid rows:', error);
  }
}


