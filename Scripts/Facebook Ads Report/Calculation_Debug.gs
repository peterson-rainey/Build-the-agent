/**
 * CALCULATION DEBUG - Diagnostic function to check calculations
 * Run this to see exactly what's happening with your data
 */

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
      const clicksIndex = headers.findIndex(h => h === 'Clicks');
      const resultsIndex = headers.findIndex(h => h === 'Results');
      const spendIndex = headers.findIndex(h => h === 'Spend');
      const revenueIndex = headers.findIndex(h => h === 'Revenue');
      
      console.log('Column indices:', {
        impressions: impressionsIndex,
        clicks: clicksIndex,
        results: resultsIndex,
        spend: spendIndex,
        revenue: revenueIndex
      });
      
      // Get raw values
      const impressions = Number(row[impressionsIndex]) || 0;
      const clicks = Number(row[clicksIndex]) || 0;
      const results = Number(row[resultsIndex]) || 0;
      const spend = Number(row[spendIndex]) || 0;
      const revenue = Number(row[revenueIndex]) || 0;
      
      console.log('Raw values:', {
        impressions: impressions,
        clicks: clicks,
        results: results,
        spend: spend,
        revenue: revenue
      });
      
      // Calculate metrics
      const ctr = impressions > 0 ? (clicks / impressions) * 100 : 0;
      const conversionRate = clicks > 0 ? (results / clicks) * 100 : 0;
      const cpc = clicks > 0 ? spend / clicks : 0;
      const cpm = impressions > 0 ? (spend / impressions) * 1000 : 0;
      const roas = spend > 0 ? revenue / spend : 0;
      
      console.log('Calculated metrics:', {
        ctr: ctr + '%',
        conversionRate: conversionRate + '%',
        cpc: '$' + cpc,
        cpm: '$' + cpm,
        roas: roas
      });
      
      // Show the math
      console.log('Math breakdown:');
      console.log(`CTR = (${clicks} / ${impressions}) * 100 = ${ctr}%`);
      console.log(`Conversion Rate = (${results} / ${clicks}) * 100 = ${conversionRate}%`);
      console.log(`CPC = ${spend} / ${clicks} = $${cpc}`);
      console.log(`CPM = (${spend} / ${impressions}) * 1000 = $${cpm}`);
      console.log(`ROAS = ${revenue} / ${spend} = ${roas}`);
    }
    
    console.log('\n=== CALCULATION DEBUG END ===');
    
  } catch (error) {
    console.error('Error in debug calculations:', error);
  }
}

/**
 * Test with sample data to verify calculations
 */
function testCalculations() {
  console.log('=== TESTING CALCULATIONS WITH SAMPLE DATA ===');
  
  // Sample data
  const testData = [
    {
      impressions: 1000,
      clicks: 50,
      results: 5,
      spend: 100,
      revenue: 200
    },
    {
      impressions: 2000,
      clicks: 100,
      results: 10,
      spend: 200,
      revenue: 400
    }
  ];
  
  testData.forEach((data, index) => {
    console.log(`\n--- TEST DATA ${index + 1} ---`);
    console.log('Input:', data);
    
    const ctr = (data.clicks / data.impressions) * 100;
    const conversionRate = (data.results / data.clicks) * 100;
    const cpc = data.spend / data.clicks;
    const cpm = (data.spend / data.impressions) * 1000;
    const roas = data.revenue / data.spend;
    
    console.log('Expected results:');
    console.log(`CTR = (${data.clicks} / ${data.impressions}) * 100 = ${ctr}%`);
    console.log(`Conversion Rate = (${data.results} / ${data.clicks}) * 100 = ${conversionRate}%`);
    console.log(`CPC = ${data.spend} / ${data.clicks} = $${cpc}`);
    console.log(`CPM = (${data.spend} / ${data.impressions}) * 1000 = $${cpm}`);
    console.log(`ROAS = ${data.revenue} / ${data.spend} = ${roas}`);
  });
  
  console.log('\n=== TEST COMPLETE ===');
}

/**
 * Compare processed data with raw data
 */
function compareData() {
  try {
    console.log('=== COMPARING RAW VS PROCESSED DATA ===');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get raw data
    const rawDataSheet = spreadsheet.getSheetByName('Raw Data');
    const rawData = rawDataSheet.getDataRange().getValues();
    const rawHeaders = rawData[0];
    const rawRows = rawData.slice(1);
    
    // Get processed data
    const processedSheet = spreadsheet.getSheetByName('Processed Data');
    const processedData = processedSheet.getDataRange().getValues();
    const processedHeaders = processedData[0];
    const processedRows = processedData.slice(1);
    
    console.log('Raw headers:', rawHeaders);
    console.log('Processed headers:', processedHeaders);
    
    // Compare first few rows
    for (let i = 0; i < Math.min(3, rawRows.length); i++) {
      console.log(`\n--- ROW ${i + 1} COMPARISON ---`);
      
      const rawRow = rawRows[i];
      const processedRow = processedRows[i];
      
      // Find relevant columns
      const rawImpressionsIndex = rawHeaders.findIndex(h => h === 'Impressions');
      const rawClicksIndex = rawHeaders.findIndex(h => h === 'Clicks');
      const rawResultsIndex = rawHeaders.findIndex(h => h === 'Results');
      
      const processedCtrIndex = processedHeaders.findIndex(h => h === 'CTR (%)');
      const processedConvIndex = processedHeaders.findIndex(h => h === 'Conversion Rate (%)');
      
      console.log('Raw values:', {
        impressions: rawRow[rawImpressionsIndex],
        clicks: rawRow[rawClicksIndex],
        results: rawRow[rawResultsIndex]
      });
      
      console.log('Processed values:', {
        ctr: processedRow[processedCtrIndex],
        conversionRate: processedRow[processedConvIndex]
      });
      
      // Calculate what it should be
      const impressions = Number(rawRow[rawImpressionsIndex]) || 0;
      const clicks = Number(rawRow[rawClicksIndex]) || 0;
      const results = Number(rawRow[rawResultsIndex]) || 0;
      
      const expectedCtr = impressions > 0 ? (clicks / impressions) * 100 : 0;
      const expectedConv = clicks > 0 ? (results / clicks) * 100 : 0;
      
      console.log('Expected values:', {
        ctr: expectedCtr + '%',
        conversionRate: expectedConv + '%'
      });
      
      console.log('Match?', {
        ctr: Math.abs(expectedCtr - Number(processedRow[processedCtrIndex])) < 0.01,
        conversionRate: Math.abs(expectedConv - Number(processedRow[processedConvIndex])) < 0.01
      });
    }
    
    console.log('\n=== COMPARISON COMPLETE ===');
    
  } catch (error) {
    console.error('Error comparing data:', error);
  }
}
