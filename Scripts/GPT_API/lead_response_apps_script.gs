/**
 * Lead Response Processing Google Apps Script
 * Handles lead conversation processing and response generation
 */

// Configuration
const SPREADSHEET_ID = '1ZhrUxuevNYM2pD_TJ6IH_WRePKKo_QCjCPW-64hKiB4';
const DATA_LOG_SHEET = 'ChatGPT_Data_Log';
const COMPANY_RULES_SHEET = 'Company_Rules';
const API_KEY = 'YOUR_OPENAI_API_KEY_HERE'; // Replace with your actual API key

/**
 * Main doGet function (for ChatGPT actions)
 */
function doGet(e) {
  try {
    // Handle case where e or e.parameter is undefined
    if (!e || !e.parameter) {
      console.log('No parameters provided, returning default response');
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: "No parameters provided. Use ?action=get_company_rules to test.",
        timestamp: new Date().toISOString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Parse request parameters
    const action = e.parameter.action;
    const dataParam = e.parameter.data;
    
    // Parse data if provided
    let data = {};
    if (dataParam) {
      try {
        data = JSON.parse(dataParam);
      } catch (parseError) {
        console.error('Error parsing data parameter:', parseError);
        data = {};
      }
    }
    
    const request = { action, data };
    
    // Log the request
    console.log('Lead Response GET Request:', request);
    
    // Handle different actions
    let result;
    switch (request.action) {
      case 'get_response_examples':
        result = handleGetResponseExamples(request.data);
        break;
      case 'get_company_rules':
        result = handleGetCompanyRules(request.data);
        break;
      case 'get_lead_history':
        result = handleGetLeadHistory(request.data);
        break;
      default:
        result = createErrorResponse('Unknown action: ' + request.action);
    }
    
    // Check if result is already a ContentService object
    if (result && typeof result.getContent === 'function') {
      // It's already a ContentService object, return it directly
      return result;
    } else {
      // It's a plain object, convert to ContentService
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch (error) {
    console.error('Lead Response GET Error:', error);
    const errorResult = createErrorResponse('Server error: ' + error.message);
    // createErrorResponse already returns a ContentService object
    return errorResult;
  }
}

/**
 * Main doPost function (for testing/debugging)
 */
function doPost(e) {
  try {
    // Parse the request
    const request = JSON.parse(e.postData.contents);
    
    // Verify API key
    if (!verifyApiKey(e)) {
      return createErrorResponse('Unauthorized: Invalid API key');
    }
    
    // Log the request
    console.log('Lead Response Request:', request);
    
    // Handle different actions
    switch (request.action) {
      case 'get_response_examples':
        return handleGetResponseExamples(request.data);
      case 'get_company_rules':
        return handleGetCompanyRules(request.data);
      case 'get_lead_history':
        return handleGetLeadHistory(request.data);
      default:
        return createErrorResponse('Unknown action: ' + request.action);
    }
    
  } catch (error) {
    console.error('Lead Response Error:', error);
    return createErrorResponse('Server error: ' + error.message);
  }
}

/**
 * Get company rules and best practices
 */
function handleGetCompanyRules(data) {
  try {
    const sheet = getCompanyRulesSheet();
    if (!sheet) {
      console.error('Company rules sheet not found');
      return createErrorResponse('Company rules sheet not found');
    }
    
    const allData = sheet.getDataRange().getValues();
    
    // Check if sheet is empty (only header row or no data)
    if (allData.length <= 1) {
      console.log('Company rules sheet is empty or has no data');
      return createSuccessResponse({
        rules: [],
        total_rules: 0,
        search_criteria: data,
        message: 'No company rules found'
      });
    }
    
    // Skip header row
    const rules = allData.slice(1).map((row, index) => ({
      row_number: index + 2,
      date_updated: row[0],
      category: row[1],
      rule_title: row[2],
      rule_description: row[3]
    }));
    
    // Filter by category if provided
    let filteredRules = rules;
    if (data.category) {
      filteredRules = rules.filter(rule => 
        rule.category === data.category
      );
    }
    
    // Filter by keywords if provided
    if (data.keywords) {
      const keywords = data.keywords.toLowerCase().split(' ');
      filteredRules = filteredRules.filter(rule => {
        const title = (rule.rule_title || '').toLowerCase();
        const description = (rule.rule_description || '').toLowerCase();
        return keywords.some(keyword => 
          title.includes(keyword) || description.includes(keyword)
        );
      });
    }
    
    return createSuccessResponse('Company rules retrieved successfully', {
      rules: filteredRules,
      total_rules: filteredRules.length,
      search_criteria: {
        category: data.category,
        keywords: data.keywords
      }
    });
    
  } catch (error) {
    console.error('Error getting company rules:', error);
    return createErrorResponse('Failed to get company rules: ' + error.message);
  }
}

/**
 * Get response examples for similar conversations
 */
function handleGetResponseExamples(data) {
  try {
    const sheet = getSheet();
    const allData = sheet.getDataRange().getValues();
    
    // Skip header row
    const responses = allData.slice(1).map((row, index) => ({
      row_number: index + 2,
      date: row[0],
      conversation_id: row[1],
      turn_index: row[2],
      conversation_context: row[3],
      industry: row[4],
      response_approach: row[5],
      key_components: row[6],
      original_question: row[7],
      immediate_context: row[8],
      context_summary: row[9],
      full_response: row[10]
    }));
    
    // Debug: Log basic info only
    console.log('Total rows found:', responses.length);
    console.log('Search keywords:', data.context_keywords);
    
    // Intelligent search logic - prioritize by relevance and recency
    let filteredResponses = responses;
    
    // Smart keyword matching with scoring
    if (data.context_keywords) {
      const keywords = data.context_keywords.toLowerCase().split(' ').filter(k => k.length > 2); // Filter out short words
      
      // Score each response based on keyword matches
      const scoredResponses = responses.map(response => {
        const searchText = [
          response.conversation_context || '',
          response.industry || '',
          response.response_approach || '',
          response.key_components || '',
          response.original_question || '',
          response.full_response || ''
        ].join(' ').toLowerCase();
        
        // Calculate relevance score
        let score = 0;
        keywords.forEach(keyword => {
          if (searchText.includes(keyword)) {
            score += 1;
            // Bonus points for matches in full_response (most important)
            if ((response.full_response || '').toLowerCase().includes(keyword)) {
              score += 2;
            }
            // Bonus points for matches in response_approach (shows strategy)
            if ((response.response_approach || '').toLowerCase().includes(keyword)) {
              score += 1;
            }
          }
        });
        
        return { ...response, relevance_score: score };
      });
      
      // Filter to responses with at least one keyword match, then sort by score
      filteredResponses = scoredResponses
        .filter(response => response.relevance_score > 0)
        .sort((a, b) => b.relevance_score - a.relevance_score);
      
      // If we don't have enough relevant examples, include more recent ones
      if (filteredResponses.length < 20) {
        console.log('Not enough relevant examples, including more recent responses');
        const recentResponses = responses
          .filter(response => {
            const responseDate = new Date(response.date);
            const threeMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate());
            return responseDate >= threeMonthsAgo;
          })
          .sort((a, b) => new Date(b.date) - new Date(a.date))
          .slice(0, 30);
        
        // Combine relevant and recent examples
        const combined = [...filteredResponses, ...recentResponses];
        // Remove duplicates and limit
        const uniqueResponses = combined.filter((response, index, self) => 
          index === self.findIndex(r => r.row_number === response.row_number)
        );
        filteredResponses = uniqueResponses.slice(0, 50);
      }
    }
    
    // Debug: Log filtering results
    console.log('After intelligent filtering:', filteredResponses.length, 'responses remain');
    
    // Industry filtering is now optional - don't restrict if no exact match
    if (data.industry && data.industry !== 'Unknown' && data.industry !== 'General') {
      const industryFiltered = filteredResponses.filter(response => 
        response.industry === data.industry
      );
      // Only apply industry filter if it doesn't make results too restrictive
      if (industryFiltered.length >= 10) {
        filteredResponses = industryFiltered;
      }
    }
    
    // Apply weighted prioritization based on recency (8, 4, 2, 1)
    const now = new Date();
    const oneMonthAgo = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());
    const threeMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate());
    const july10th2025 = new Date('2025-07-10');
    
    // Assign priority weights and sort by weighted score
    const prioritizedResponses = filteredResponses.map(response => {
      const responseDate = new Date(response.date);
      let priorityWeight = 1; // Default: deprecated
      
      if (responseDate >= oneMonthAgo) {
        priorityWeight = 8; // High priority: Last 1 month
      } else if (responseDate >= threeMonthsAgo) {
        priorityWeight = 4; // Medium priority: Last 3 months (excluding 1 month)
      } else if (responseDate > july10th2025) {
        priorityWeight = 2; // Low priority: After July 10, 2025 but older than 3 months
      } else {
        priorityWeight = 1; // Deprecated: July 10, 2025 and earlier
      }
      
      return {
        ...response,
        priority_weight: priorityWeight
      };
    }).sort((a, b) => {
      // Primary sort: by priority weight (higher first)
      if (a.priority_weight !== b.priority_weight) {
        return b.priority_weight - a.priority_weight;
      }
      // Secondary sort: by date (more recent first) within same priority
      return new Date(b.date) - new Date(a.date);
    });
    
    // Limit results - significantly increased for better examples
    const limit = data.limit || 100; // Increased from 50 to 100 for much more data
    filteredResponses = prioritizedResponses.slice(0, limit);
    
    // Count examples by priority weight
    const priorityCounts = {
      weight_8: filteredResponses.filter(r => r.priority_weight === 8).length,
      weight_4: filteredResponses.filter(r => r.priority_weight === 4).length,
      weight_2: filteredResponses.filter(r => r.priority_weight === 2).length,
      weight_1: filteredResponses.filter(r => r.priority_weight === 1).length
    };
    
    return createSuccessResponse('Response examples retrieved successfully', {
      examples: filteredResponses,
      total_found: filteredResponses.length,
      prioritization_info: {
        weight_8_count: priorityCounts.weight_8, // Last 1 month (highest priority)
        weight_4_count: priorityCounts.weight_4, // Last 3 months (high priority)
        weight_2_count: priorityCounts.weight_2, // Older than 3 months (low priority)
        weight_1_count: priorityCounts.weight_1, // July 2025 and earlier (deprecated)
        note: "Priority weights: 8 (last 1 month) > 4 (last 3 months) > 2 (older than 3 months) > 1 (July 2025 and earlier)"
      },
      search_criteria: {
        context_keywords: data.context_keywords,
        industry: data.industry,
        limit: limit
      }
    });
    
  } catch (error) {
    console.error('Error getting response examples:', error);
    return createErrorResponse('Failed to get response examples: ' + error.message);
  }
}


/**
 * Get conversation history for a specific lead
 */
function handleGetLeadHistory(data) {
  try {
    const sheet = getSheet();
    const allData = sheet.getDataRange().getValues();
    
    // Skip header row
    const responses = allData.slice(1).map((row, index) => ({
      row_number: index + 2,
      date: row[0],
      conversation_id: row[1],
      turn_index: row[2],
      conversation_context: row[3],
      industry: row[4],
      response_approach: row[5],
      key_components: row[6],
      original_question: row[7],
      immediate_context: row[8],
      context_summary: row[9],
      full_response: row[10]
    }));
    
    // Filter by conversation ID if provided
    let leadHistory = [];
    if (data.conversation_id) {
      leadHistory = responses.filter(response => 
        response.conversation_id === data.conversation_id
      );
    } else if (data.context_keywords) {
      // Search by context keywords if no specific conversation ID
      const keywords = data.context_keywords.toLowerCase().split(' ');
      leadHistory = responses.filter(response => {
        const context = (response.conversation_context || '').toLowerCase();
        const industry = (response.industry || '').toLowerCase();
        return keywords.some(keyword => 
          context.includes(keyword) || industry.includes(keyword)
        );
      });
    }
    
    // Sort by conversation ID and turn index
    leadHistory.sort((a, b) => {
      if (a.conversation_id === b.conversation_id) {
        return a.turn_index - b.turn_index;
      }
      return a.conversation_id - b.conversation_id;
    });
    
    return createSuccessResponse('Lead history retrieved successfully', {
      lead_history: leadHistory,
      total_interactions: leadHistory.length,
      search_criteria: {
        conversation_id: data.conversation_id,
        context_keywords: data.context_keywords
      }
    });
    
  } catch (error) {
    console.error('Error getting lead history:', error);
    return createErrorResponse('Failed to get lead history: ' + error.message);
  }
}

/**
 * Get the data log sheet
 */
function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Debug: List all sheet names
  const sheets = spreadsheet.getSheets();
  console.log('Available sheets:', sheets.map(sheet => sheet.getName()));
  console.log('Looking for sheet:', DATA_LOG_SHEET);
  
  const targetSheet = spreadsheet.getSheetByName(DATA_LOG_SHEET);
  if (!targetSheet) {
    console.error('Sheet not found:', DATA_LOG_SHEET);
    console.log('Available sheets:', sheets.map(sheet => sheet.getName()));
  }
  
  return targetSheet;
}

/**
 * Get the company rules sheet
 */
function getCompanyRulesSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Debug: List all sheet names
  const sheets = spreadsheet.getSheets();
  console.log('Available sheets for company rules:', sheets.map(sheet => sheet.getName()));
  console.log('Looking for company rules sheet:', COMPANY_RULES_SHEET);
  
  const targetSheet = spreadsheet.getSheetByName(COMPANY_RULES_SHEET);
  if (!targetSheet) {
    console.error('Company rules sheet not found:', COMPANY_RULES_SHEET);
    console.log('Available sheets:', sheets.map(sheet => sheet.getName()));
  }
  
  return targetSheet;
}

/**
 * Verify API key
 */
function verifyApiKey(e) {
  const authHeader = e.headers.Authorization || e.headers.authorization;
  return authHeader === `Bearer ${API_KEY}` || authHeader === API_KEY;
}

/**
 * Create success response
 */
function createSuccessResponse(message, data = null) {
  const response = {
    success: true,
    message: message,
    timestamp: new Date().toISOString()
  };
  
  if (data) {
    response.data = data;
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Create error response
 */
function createErrorResponse(message) {
  const response = {
    success: false,
    error: message,
    timestamp: new Date().toISOString()
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Simple test function for doGet
 */
function testDoGetSimple() {
  const mockEvent = {
    parameter: {
      action: 'get_company_rules'
    }
  };
  
  console.log('Testing doGet with mock event...');
  const result = doGet(mockEvent);
  console.log('doGet result:', result.getContent());
  return result;
}

/**
 * Test function specifically for company rules debugging
 */
function testCompanyRules() {
  console.log('Testing company rules function...');
  
  try {
    // Test sheet access
    const sheet = getCompanyRulesSheet();
    console.log('Sheet found:', sheet ? 'Yes' : 'No');
    
    if (sheet) {
      const allData = sheet.getDataRange().getValues();
      console.log('Data rows:', allData.length);
      console.log('First few rows:', allData.slice(0, 3));
    }
    
    // Test the full function
    const result = handleGetCompanyRules({});
    console.log('Company rules result:', result);
    
    return result;
  } catch (error) {
    console.error('Error in testCompanyRules:', error);
    return { error: error.message };
  }
}

/**
 * Test function to check available sheets
 */
function testSheets() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    console.log('=== SHEET DEBUG INFO ===');
    console.log('Spreadsheet ID:', SPREADSHEET_ID);
    console.log('Looking for Data Log sheet:', DATA_LOG_SHEET);
    console.log('Looking for Company Rules sheet:', COMPANY_RULES_SHEET);
    console.log('Available sheets:');
    
    sheets.forEach((sheet, index) => {
      console.log(`${index + 1}. "${sheet.getName()}"`);
    });
    
    return {
      success: true,
      spreadsheet_id: SPREADSHEET_ID,
      data_log_sheet: DATA_LOG_SHEET,
      company_rules_sheet: COMPANY_RULES_SHEET,
      available_sheets: sheets.map(sheet => sheet.getName())
    };
    
  } catch (error) {
    console.error('Error checking sheets:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test function to check raw data without filtering
 */
function testRawData() {
  try {
    const sheet = getSheet();
    const allData = sheet.getDataRange().getValues();
    
    console.log('=== RAW DATA DEBUG ===');
    console.log('Total rows in sheet:', allData.length);
    console.log('Has header:', allData.length > 0);
    console.log('Has data rows:', allData.length > 1);
    
    if (allData.length > 1) {
      const sampleRow = allData[1];
      console.log('Sample row (first 5 columns):', sampleRow.slice(0, 5));
    }
    
    return {
      success: true,
      total_rows: allData.length,
      has_data: allData.length > 1
    };
    
  } catch (error) {
    console.error('Error checking raw data:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test function
 */
function testLeadResponseAPI() {
  const testData = {
    action: 'get_response_examples',
    data: {
      context_keywords: 'social media management',
      status: 'Active',
      limit: 25
    }
  };
  
  const mockEvent = {
    postData: {
      contents: JSON.stringify(testData)
    },
    headers: {
      Authorization: `Bearer ${API_KEY}`
    }
  };
  
  const result = doPost(mockEvent);
  console.log('Test Result:', result.getContent());
}
