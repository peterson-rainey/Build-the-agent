// ============================================
// LEAD RESPONSE GPT - MAIN BACKEND (Apps Script)
// ============================================

function doGet(e) {
  // If accessed with ?test=auth parameter, trigger authorization test
  if (e && e.parameter && e.parameter.test === 'auth') {
    try {
      // This will trigger authorization if not already granted
      var response = UrlFetchApp.fetch('https://httpbin.org/get', {
        muteHttpExceptions: true
      });
      return HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head><title>Auth Test</title></head><body style="font-family: Arial; padding: 40px; background: #f5f5f5;">' +
        '<h1>✅ Authorization Test</h1>' +
        '<p><strong>Status Code:</strong> ' + response.getResponseCode() + '</p>' +
        '<p>If you see 200, authorization is working!</p>' +
        '<p><a href="?">Go to main app</a></p>' +
        '</body></html>'
      );
    } catch (err) {
      return HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head><title>Auth Error</title></head><body style="font-family: Arial; padding: 40px; background: #f5f5f5;">' +
        '<h1>❌ Authorization Required</h1>' +
        '<p><strong>Error:</strong> ' + err.toString() + '</p>' +
        '<p>Please check the execution log in Apps Script editor for authorization prompt.</p>' +
        '<p><a href="?">Try main app</a></p>' +
        '</body></html>'
      );
    }
  }
  
  // Normal web app - with error handling
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Lead Response GPT')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    // If Index.html is missing or there's an error, show diagnostic
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><title>Error</title></head><body style="font-family: Arial; padding: 40px;">' +
      '<h1>Error Loading App</h1>' +
      '<p><strong>Error:</strong> ' + err.toString() + '</p>' +
      '<p>Please check that Index.html exists in your Apps Script project.</p>' +
      '</body></html>'
    );
  }
}

// ============================================
// PUBLIC SERVER FUNCTIONS (called from Index.html)
// ============================================

function getCompanyRules(options) {
  const ss = SpreadsheetApp.openById(LEADS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(COMPANY_RULES_SHEET);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return [];

  const rows = values.slice(1).map(function(row, idx){
    return {
      row_number: idx + 2,
      date_updated: row[0],
      category: row[1],
      rule_title: row[2],
      rule_description: row[3]
    };
  });

  var filtered = rows;
  if (options && options.category) {
    filtered = filtered.filter(function(r){ return r.category === options.category; });
  }
  if (options && options.keywords) {
    var kws = String(options.keywords).toLowerCase().split(' ').filter(function(k){ return k; });
    filtered = filtered.filter(function(r){
      var hay = ((r.rule_title || '') + ' ' + (r.rule_description || '')).toLowerCase();
      return kws.some(function(k){ return hay.indexOf(k) !== -1; });
    });
  }

  var limit = (options && options.limit) ? Math.max(1, parseInt(options.limit, 10)) : 25;
  var offset = (options && options.offset) ? Math.max(0, parseInt(options.offset, 10)) : 0;
  return filtered.slice(offset, offset + limit);
}

function getResponseExamples(options) {
  const ss = SpreadsheetApp.openById(LEADS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(DATA_LOG_SHEET);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return [];

  const rows = values.slice(1).map(function(row, idx){
    return {
      row_number: idx + 2,
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
    };
  });

  var filtered = rows;
  if (options && options.context_keywords) {
    // Enhanced keyword matching with synonyms and stemming approximation
    filtered = matchExamplesWithKeywords_(rows, options.context_keywords);
  }

  // Recency weighting (8,4,2,1)
  var now = new Date();
  var oneMonthAgo = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());
  var threeMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate());
  var july10th2025 = new Date('2025-07-10');

  filtered = (filtered === rows ? rows.slice() : filtered).map(function(r){
    var d = new Date(r.date);
    var weight = 1;
    if (d >= oneMonthAgo) weight = 8;
    else if (d >= threeMonthsAgo) weight = 4;
    else if (d > july10th2025) weight = 2;
    else weight = 1;
    return Object.assign({}, r, { priority_weight: weight });
  }).sort(function(a,b){
    if (b.priority_weight !== a.priority_weight) return b.priority_weight - a.priority_weight;
    return new Date(b.date) - new Date(a.date);
  });

  var limit = (options && options.limit) ? Math.max(1, parseInt(options.limit, 10)) : 25;
  var offset = (options && options.offset) ? Math.max(0, parseInt(options.offset, 10)) : 0;
  return filtered.slice(offset, offset + limit);
}

/**
 * Get lightweight examples (only metadata, not full responses) for AI selection
 * Returns: { row_number, response_approach, key_components, context_summary, date, priority_weight }
 */
function getLightweightExamples(options) {
  var examples = getResponseExamples(options);
  return examples.map(function(ex){
    return {
      row_number: ex.row_number,
      response_approach: ex.response_approach || '',
      key_components: ex.key_components || '',
      context_summary: ex.context_summary || '',
      date: ex.date || '',
      priority_weight: ex.priority_weight || 1
    };
  });
}

/**
 * Get lightweight rules (only titles, not descriptions) for AI selection
 * Returns: { row_number, category, rule_title }
 */
function getLightweightRules(options) {
  var rules = getCompanyRules(options);
  return rules.map(function(r){
    return {
      row_number: r.row_number,
      category: r.category || '',
      rule_title: r.rule_title || ''
    };
  });
}

/**
 * Get full example details by row numbers
 */
function getFullExamplesByRowNumbers(rowNumbers) {
  if (!rowNumbers || !Array.isArray(rowNumbers) || rowNumbers.length === 0) return [];
  const ss = SpreadsheetApp.openById(LEADS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(DATA_LOG_SHEET);
  if (!sheet) return [];
  
  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return [];
  
  var result = [];
  rowNumbers.forEach(function(rowNum){
    if (rowNum >= 2 && rowNum <= values.length) {
      var row = values[rowNum - 1]; // rowNum is 1-indexed, array is 0-indexed
      result.push({
        row_number: rowNum,
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
      });
    }
  });
  return result;
}

/**
 * Get full rule details by row numbers
 */
function getFullRulesByRowNumbers(rowNumbers) {
  if (!rowNumbers || !Array.isArray(rowNumbers) || rowNumbers.length === 0) return [];
  const ss = SpreadsheetApp.openById(LEADS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(COMPANY_RULES_SHEET);
  if (!sheet) return [];
  
  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return [];
  
  var result = [];
  rowNumbers.forEach(function(rowNum){
    if (rowNum >= 2 && rowNum <= values.length) {
      var row = values[rowNum - 1]; // rowNum is 1-indexed, array is 0-indexed
      result.push({
        row_number: rowNum,
        date_updated: row[0],
        category: row[1],
        rule_title: row[2],
        rule_description: row[3]
      });
    }
  });
  return result;
}

function getLeadHistory(options) {
  const ss = SpreadsheetApp.openById(LEADS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(DATA_LOG_SHEET);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return [];
  var rows = values.slice(1).map(function(row, idx){
    return {
      row_number: idx + 2,
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
    };
  });
  if (options && options.conversation_id) {
    rows = rows.filter(function(r){ return r.conversation_id === options.conversation_id; });
  }
  if (options && options.context_keywords) {
    var kws = String(options.context_keywords).toLowerCase().split(' ');
    rows = rows.filter(function(r){
      var hay = ((r.conversation_context || '') + ' ' + (r.industry || '')).toLowerCase();
      return kws.some(function(k){ return hay.indexOf(k) !== -1; });
    });
  }
  rows.sort(function(a,b){
    if (a.conversation_id === b.conversation_id) return a.turn_index - b.turn_index;
    return (a.conversation_id || 0) - (b.conversation_id || 0);
  });
  return rows;
}

// ============================================
// OPENAI CALL AND RESPONSE GENERATION
// ============================================

/**
 * Helper: Check if a rule is a formatting/writing style rule
 */
function isFormattingRule_(rule) {
  var category = ((rule.category || '') + ' ' + (rule.rule_title || '')).toLowerCase();
  return category.indexOf('formatting') !== -1 || 
         category.indexOf('writing') !== -1 || 
         category.indexOf('style') !== -1 ||
         category.indexOf('em-dash') !== -1 ||
         category.indexOf('dash') !== -1 ||
         category.indexOf('markdown') !== -1 ||
         category.indexOf('citation') !== -1 ||
         category.indexOf('bullet') !== -1;
}

/**
 * Helper: Prioritize formatting rules when selecting rules
 * Returns array of rule row numbers with formatting rules first
 */
function prioritizeFormattingRules_(lightweightRules, maxCount) {
  var formattingRules = [];
  var otherRules = [];
  lightweightRules.forEach(function(r){
    if (isFormattingRule_(r)) {
      formattingRules.push(r);
    } else {
      otherRules.push(r);
    }
  });
  // Sort both by date (most recent first)
  formattingRules.sort(function(a, b){
    var dateA = a.date_updated ? new Date(a.date_updated) : new Date(0);
    var dateB = b.date_updated ? new Date(b.date_updated) : new Date(0);
    return dateB - dateA;
  });
  otherRules.sort(function(a, b){
    var dateA = a.date_updated ? new Date(a.date_updated) : new Date(0);
    var dateB = b.date_updated ? new Date(b.date_updated) : new Date(0);
    return dateB - dateA;
  });
  // Combine: all formatting rules first, then fill remaining slots
  return formattingRules.concat(otherRules).slice(0, maxCount).map(function(r){ return r.row_number; });
}

/**
 * Use AI to select the most relevant examples and rules based on conversation context
 * Returns: { exampleRowNumbers: [], ruleRowNumbers: [] }
 */
function selectRelevantItemsWithAI(conversation, lightweightExamples, lightweightRules) {
  try {
    var apiKey = getOpenAIKey();
    if (!apiKey) {
      // Fallback: return top items by priority_weight, sorted by date (most recent first)
      var topExamples = lightweightExamples.slice(0, (typeof EXAMPLES_IN_PROMPT !== 'undefined' ? EXAMPLES_IN_PROMPT : 10));
      var topRules = prioritizeFormattingRules_(lightweightRules, 25);
      return {
        exampleRowNumbers: topExamples.map(function(e){ return e.row_number; }),
        ruleRowNumbers: topRules,
        usage: null
      };
    }

    // Build lightweight data for AI
    // Limit to prevent token overflow - estimate ~100 tokens per example, ~50 per rule
    var maxExamplesForSelection = Math.min(lightweightExamples.length, 300); // Cap at 300 examples for selection
    var maxRulesForSelection = Math.min(lightweightRules.length, 150); // Cap at 150 rules for selection
    
    var examplesText = 'Response Examples (select up to ' + (typeof EXAMPLES_IN_PROMPT !== 'undefined' ? EXAMPLES_IN_PROMPT : 10) + '):\n';
    lightweightExamples.slice(0, maxExamplesForSelection).forEach(function(ex, idx){
      examplesText += (idx + 1) + '. Row ' + ex.row_number + ':\n';
      examplesText += '   Response Approach: ' + (ex.response_approach || 'N/A') + '\n';
      examplesText += '   Key Components: ' + (ex.key_components || 'N/A') + '\n';
      examplesText += '   Context Summary: ' + (ex.context_summary || 'N/A') + '\n';
      examplesText += '   Date: ' + (ex.date || 'N/A') + '\n';
      examplesText += '   Priority Weight: ' + (ex.priority_weight || 1) + '\n\n';
    });
    if (lightweightExamples.length > maxExamplesForSelection) {
      examplesText += '\n[Note: Showing first ' + maxExamplesForSelection + ' of ' + lightweightExamples.length + ' examples for selection. Prioritize from these.]\n';
    }

    // Separate formatting/writing style rules from other rules
    var formattingRules = [];
    var otherRules = [];
    lightweightRules.slice(0, maxRulesForSelection).forEach(function(r){
      if (isFormattingRule_(r)) {
        formattingRules.push(r);
      } else {
        otherRules.push(r);
      }
    });
    
    var rulesText = 'Company Rules (select up to 25, prioritizing most recent):\n';
    rulesText += 'Rules are sorted by date (most recent first). When multiple rules are equally relevant, prefer the more recent ones.\n\n';
    rulesText += 'IMPORTANT: Formatting and Writing Style rules (listed first) should ALWAYS be included if present. These are universal guidelines that apply to all responses.\n\n';
    
    // List formatting rules first
    if (formattingRules.length > 0) {
      rulesText += '=== FORMATTING/WRITING STYLE RULES (ALWAYS INCLUDE THESE) ===\n';
      formattingRules.forEach(function(r, idx){
        rulesText += (idx + 1) + '. Row ' + r.row_number + ':\n';
        rulesText += '   Category: ' + (r.category || 'N/A') + '\n';
        rulesText += '   Rule Title: ' + (r.rule_title || 'N/A') + '\n';
        if (r.date_updated) {
          rulesText += '   Date Updated: ' + (r.date_updated || 'N/A') + '\n';
        }
        rulesText += '\n';
      });
      rulesText += '\n';
    }
    
    // Then list other rules
    if (otherRules.length > 0) {
      rulesText += '=== OTHER RULES (Select relevant ones) ===\n';
      otherRules.forEach(function(r, idx){
        rulesText += (formattingRules.length + idx + 1) + '. Row ' + r.row_number + ':\n';
        rulesText += '   Category: ' + (r.category || 'N/A') + '\n';
        rulesText += '   Rule Title: ' + (r.rule_title || 'N/A') + '\n';
        if (r.date_updated) {
          rulesText += '   Date Updated: ' + (r.date_updated || 'N/A') + '\n';
        }
        rulesText += '\n';
      });
    }
    if (lightweightRules.length > maxRulesForSelection) {
      rulesText += '\n[Note: Showing first ' + maxRulesForSelection + ' of ' + lightweightRules.length + ' rules for selection. Prioritize from these.]\n';
    }

    var systemPrompt = 'You are an expert at selecting the most relevant examples and rules from a pre-filtered set to generate an exceptional response.\n\n' +
      'You will receive:\n' +
      '1. A conversation/query that needs a response\n' +
      '2. A pre-filtered list of response examples (already keyword-matched) with metadata\n' +
      '3. A list of company rules with titles\n\n' +
      'Your task: Intelligently select the items that will help generate the BEST, most contextually appropriate response.\n\n' +
      'Selection Strategy:\n' +
      '- Prioritize examples that match the conversation\'s intent, tone, and context\n' +
      '- Consider the conversation_context, response_approach, and key_components when evaluating relevance\n' +
      '- Prefer examples with higher priority_weight when relevance is similar\n' +
      '- Select diverse examples that cover different aspects if the conversation is complex\n' +
      '- For rules: Always include formatting rules, then select rules that directly apply to the conversation context\n\n' +
      'Output format (JSON only, no other text):\n' +
      '{\n' +
      '  "exampleRowNumbers": [row1, row2, ...],\n' +
      '  "ruleRowNumbers": [row1, row2, ...]\n' +
      '}\n\n' +
      'For examples: Select up to ' + (typeof EXAMPLES_IN_PROMPT !== 'undefined' ? EXAMPLES_IN_PROMPT : 10) + ' examples. Prioritize items that are most relevant to the conversation context. Consider priority_weight when examples are equally relevant.\n\n' +
      'For rules: Select up to ' + (typeof MAX_RULES_SELECTION !== 'undefined' ? MAX_RULES_SELECTION : 50) + ' rules total. Rules fall into different categories:\n' +
      '1. FORMATTING/WRITING STYLE RULES: These are universal guidelines (e.g., "never use em-dashes", "no markdown", "plain text only"). These should ALWAYS be included if present - they apply to every response.\n' +
      '2. CADENCE/COMMUNICATION RULES: Guidelines on how to communicate, tone, approach (e.g., "diagnostic-first approach", "casual professionalism"). Include these if they relate to the conversation style.\n' +
      '3. PRICING RULES: Pricing policies, fee structures, billing guidelines. Include if the conversation mentions pricing, budget, or costs.\n' +
      '4. RESPONSE GUIDELINES: How to respond to leads, qualification processes, objection handling. Include if relevant to the conversation context.\n' +
      '5. GENERAL POLICIES: Other company rules and procedures. Include if they would help craft an appropriate response.\n\n' +
      'Selection criteria:\n' +
      '- ALWAYS include ALL formatting/writing style rules (they are universal)\n' +
      '- Include cadence/communication rules if they relate to the conversation style\n' +
      '- Include pricing rules if pricing/budget is mentioned\n' +
      '- Include response guidelines if relevant to the conversation context\n' +
      '- Include general policies that would help craft the response\n' +
      '- Prioritize most recent rules when multiple are equally relevant\n' +
      '- Be inclusive rather than exclusive - when in doubt, include the rule\n\n' +
      'IMPORTANT: You MUST select at least 20-30 rules (including all formatting rules). Formatting and writing style rules are mandatory. Do not return an empty ruleRowNumbers array.';

    var userPrompt = '# Conversation/Query\n' + conversation + '\n\n' +
      examplesText + '\n' +
      rulesText + '\n\n' +
      '# Task\n' +
      'Select the most relevant examples and rules for generating a response to the conversation above.\n' +
      'Return ONLY a JSON object with exampleRowNumbers and ruleRowNumbers arrays.\n\n' +
      'CRITICAL: You MUST include rules in your selection.\n' +
      '- ALWAYS include ALL formatting/writing style rules (they are listed first and marked as mandatory)\n' +
      '- Select at least 20-30 total rules (including formatting rules)\n' +
      '- Include cadence, pricing, response guidelines, and general policies that are relevant\n' +
      '- Do not return an empty ruleRowNumbers array\n' +
      '- When in doubt, include the rule - it\'s better to have more context';

    var body = {
      model: OPENAI_MODEL,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ],
      response_format: { type: 'json_object' }
    };

    // GPT-5/5.1 uses max_completion_tokens instead of max_tokens, and only supports temperature=1
    var isGPT5 = (typeof OPENAI_MODEL !== 'undefined' && (String(OPENAI_MODEL).toLowerCase().includes('gpt-5') || String(OPENAI_MODEL).toLowerCase().includes('gpt-5.1')));
    if (isGPT5) {
      body.max_completion_tokens = 500;
      body.temperature = 1;
    } else {
      body.max_tokens = 500;
      body.temperature = 0.3; // Lower temperature for selection task
    }

    var resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    var responseCode = resp.getResponseCode();
    if (responseCode !== 200) {
      Logger.log('AI selection failed: ' + responseCode + ' - ' + resp.getContentText());
      // Fallback to priority-based selection (prioritize formatting rules)
      var topExamples = lightweightExamples.slice(0, (typeof EXAMPLES_IN_PROMPT !== 'undefined' ? EXAMPLES_IN_PROMPT : 10));
      var topRules = prioritizeFormattingRules_(lightweightRules, 25);
      return {
        exampleRowNumbers: topExamples.map(function(e){ return e.row_number; }),
        ruleRowNumbers: topRules,
        usage: null
      };
    }

    var data = JSON.parse(resp.getContentText());
    var usage = data && data.usage;
    if (!data.choices || !data.choices[0] || !data.choices[0].message) {
      // Fallback (prioritize formatting rules)
      var topExamples = lightweightExamples.slice(0, (typeof EXAMPLES_IN_PROMPT !== 'undefined' ? EXAMPLES_IN_PROMPT : 10));
      var topRules = prioritizeFormattingRules_(lightweightRules, 25);
      return {
        exampleRowNumbers: topExamples.map(function(e){ return e.row_number; }),
        ruleRowNumbers: topRules,
        usage: usage
      };
    }

    var answer = data.choices[0].message.content;
    if (!answer) {
      // Fallback (prioritize formatting rules)
      var topExamples = lightweightExamples.slice(0, (typeof EXAMPLES_IN_PROMPT !== 'undefined' ? EXAMPLES_IN_PROMPT : 10));
      var topRules = prioritizeFormattingRules_(lightweightRules, 25);
      return {
        exampleRowNumbers: topExamples.map(function(e){ return e.row_number; }),
        ruleRowNumbers: topRules,
        usage: usage
      };
    }

    // Parse JSON response
    var selection;
    try {
      selection = JSON.parse(answer);
    } catch (e) {
      Logger.log('Failed to parse AI selection JSON: ' + answer);
      // Fallback (rules already sorted by date)
      var topExamples = lightweightExamples.slice(0, (typeof EXAMPLES_IN_PROMPT !== 'undefined' ? EXAMPLES_IN_PROMPT : 10));
      var maxRules = (typeof MAX_RULES_SELECTION !== 'undefined' ? MAX_RULES_SELECTION : 50);
      var topRules = lightweightRules.slice(0, maxRules); // Limit to max rules, already sorted by date
      Logger.log('Fallback: Using ' + topRules.length + ' rules from sorted list');
      return {
        exampleRowNumbers: topExamples.map(function(e){ return e.row_number; }),
        ruleRowNumbers: topRules.map(function(r){ return r.row_number; }),
        usage: usage
      };
    }

    var maxRules = (typeof MAX_RULES_SELECTION !== 'undefined' ? MAX_RULES_SELECTION : 50);
    var selectedRuleNumbers = (selection.ruleRowNumbers || []).slice(0, maxRules);
    
    // Log what AI selected for debugging
    if (selectedRuleNumbers.length === 0) {
      Logger.log('WARNING: AI selected zero rules. This may indicate the selection prompt needs adjustment.');
      Logger.log('AI response was: ' + answer.substring(0, 500));
    } else {
      Logger.log('AI selected ' + selectedRuleNumbers.length + ' rules: ' + selectedRuleNumbers.join(', '));
    }
    
    return {
      exampleRowNumbers: (selection.exampleRowNumbers || []).slice(0, (typeof EXAMPLES_IN_PROMPT !== 'undefined' ? EXAMPLES_IN_PROMPT : 10)),
      ruleRowNumbers: selectedRuleNumbers, // Limit to max rules
      usage: usage
    };

  } catch (e) {
    Logger.log('Error in selectRelevantItemsWithAI: ' + String(e));
    // Fallback (rules already sorted by date)
    var topExamples = lightweightExamples.slice(0, (typeof EXAMPLES_IN_PROMPT !== 'undefined' ? EXAMPLES_IN_PROMPT : 10));
    var topRules = lightweightRules.slice(0, 25); // Limit to 25 rules max, already sorted by date
    return {
      exampleRowNumbers: topExamples.map(function(e){ return e.row_number; }),
      ruleRowNumbers: topRules.map(function(r){ return r.row_number; }),
      usage: null
    };
  }
}

function generateLeadResponse(payload) {
  try {
    if (!payload || !payload.conversation) {
      Logger.log('ERROR: Missing conversation text in generateLeadResponse');
      return { success: false, error: 'Missing conversation text' };
    }

    var startTime = new Date();
    Logger.log('=== Starting Lead Response Generation ===');
    Logger.log('Type: ' + (payload.type || 'lead'));
    Logger.log('Conversation length: ' + (payload.conversation || '').length + ' characters');

    // PHASE 1: Extract semantic keywords using AI
    Logger.log('[Phase 1] Extracting keywords with AI...');
    var aiKeywords = extractKeywordsWithAI(payload.conversation);
    Logger.log('[Phase 1] AI extracted keywords: ' + aiKeywords);

    // PHASE 2: Use AI keywords to filter examples in Apps Script (get 50-300 rows)
    Logger.log('[Phase 2] Filtering examples with AI keywords...');
    var filterStartTime = new Date();
    var filteredExamples = getResponseExamples({ 
      context_keywords: aiKeywords, 
      limit: 300, // Get up to 300 examples that match AI keywords
      offset: 0 
    });
    Logger.log('[Phase 2] Filtering took ' + (new Date() - filterStartTime) + 'ms');
    
    // Get lightweight version of filtered examples
    var lightweightExamples = filteredExamples.map(function(ex){
      return {
        row_number: ex.row_number,
        response_approach: ex.response_approach || '',
        key_components: ex.key_components || '',
        context_summary: ex.context_summary || '',
        date: ex.date || '',
        priority_weight: ex.priority_weight || 1
      };
    });
    
    Logger.log('[Phase 2] Filtered to ' + lightweightExamples.length + ' examples using AI keywords');
    if (lightweightExamples.length > 0) {
      Logger.log('[Phase 2] Top example scores: ' + filteredExamples.slice(0, 5).map(function(e){ return 'Row ' + e.row_number + ' (' + (e.relevance_score || 0) + ')'; }).join(', '));
    }
    
    // Get all rules (both regular and PDF) for lightweight selection - no limit
    var allRules = getCompanyRules({ limit: 1000, offset: 0 }); // Get all rules, no artificial limit
    // Sort by date_updated (most recent first) before creating lightweight version
    var regularRules = allRules
      .filter(function(r){ return !(r.category || '').startsWith('PDF -'); })
      .sort(function(a, b){
        // Sort by date_updated descending (most recent first)
        var dateA = a.date_updated ? new Date(a.date_updated) : new Date(0);
        var dateB = b.date_updated ? new Date(b.date_updated) : new Date(0);
        return dateB - dateA; // Descending order
      });
    var regularRulesLight = regularRules
      .map(function(r){ return { row_number: r.row_number, category: r.category, rule_title: r.rule_title, date_updated: r.date_updated }; });
    
    // PHASE 3: Use AI to select most relevant items from filtered set
    Logger.log('[Phase 3] AI selecting top examples and rules from filtered set...');
    var selectionStartTime = new Date();
    var selection = selectRelevantItemsWithAI(payload.conversation, lightweightExamples, regularRulesLight);
    Logger.log('[Phase 3] Selection took ' + (new Date() - selectionStartTime) + 'ms');
    Logger.log('[Phase 3] AI selected: ' + (selection.exampleRowNumbers ? selection.exampleRowNumbers.length : 0) + ' examples, ' + (selection.ruleRowNumbers ? selection.ruleRowNumbers.length : 0) + ' rules');
    
    // PHASE 4: Fetch full details for selected items only
    Logger.log('[Phase 4] Fetching full details for selected items...');
    var fetchStartTime = new Date();
    var selectedExamples = getFullExamplesByRowNumbers(selection.exampleRowNumbers);
    var selectedRules = getFullRulesByRowNumbers(selection.ruleRowNumbers);
    Logger.log('[Phase 4] Fetching took ' + (new Date() - fetchStartTime) + 'ms');
    Logger.log('[Phase 4] Fetched: ' + selectedExamples.length + ' examples, ' + selectedRules.length + ' rules');
    
    // Ensure formatting/writing style rules are always included
    var formattingRuleNumbers = [];
    regularRules.forEach(function(r){
      if (isFormattingRule_(r)) {
        formattingRuleNumbers.push(r.row_number);
      }
    });
    
    // Add formatting rules that weren't selected by AI
    var selectedRuleNumbers = (selection.ruleRowNumbers || []).slice();
    formattingRuleNumbers.forEach(function(ruleNum){
      if (selectedRuleNumbers.indexOf(ruleNum) === -1) {
        selectedRuleNumbers.push(ruleNum);
        Logger.log('Added formatting rule ' + ruleNum + ' that was missed by AI selection');
      }
    });
    
    // Fetch full details for all selected rules (including added formatting rules)
    selectedRules = getFullRulesByRowNumbers(selectedRuleNumbers);
    
    // Safety check: If AI selected too many rules, cap it and log a warning
    var maxRules = (typeof MAX_RULES_SELECTION !== 'undefined' ? MAX_RULES_SELECTION : 50);
    if (selectedRules.length > maxRules) {
      Logger.log('WARNING: AI selected ' + selectedRules.length + ' rules. Capping at ' + maxRules + '.');
      // Keep formatting rules, then most recent others
      var formattingRulesList = selectedRules.filter(function(r){ return isFormattingRule_(r); });
      var otherRulesList = selectedRules.filter(function(r){ return !isFormattingRule_(r); });
      // Keep all formatting rules, then fill remaining slots with other rules
      var remainingSlots = maxRules - formattingRulesList.length;
      selectedRules = formattingRulesList.concat(otherRulesList.slice(0, Math.max(0, remainingSlots)));
    }
    
    // Fallback: If AI selected very few rules (<20), include more recent rules as a safety net
    if (selectedRules.length < 20 && regularRules.length > 0) {
      Logger.log('WARNING: AI selected only ' + selectedRules.length + ' rules. Adding more recent rules as fallback.');
      var selectedRuleNumbersSet = {};
      selectedRules.forEach(function(r){ selectedRuleNumbersSet[r.row_number] = true; });
      var additionalRules = regularRules.filter(function(r){
        return !selectedRuleNumbersSet[r.row_number];
      }).slice(0, 20 - selectedRules.length);
      selectedRules = selectedRules.concat(additionalRules.map(function(r){
        return {
          row_number: r.row_number,
          date_updated: r.date_updated,
          category: r.category,
          rule_title: r.rule_title,
          rule_description: r.rule_description
        };
      }));
    }
    
    // Get PDF chunks separately (they're handled differently)
    var pdfChunks = allRules.filter(function(r){ return (r.category || '').startsWith('PDF -'); });
    var selectedPDFChunks = selectRelevantPDFChunks_(pdfChunks, aiKeywords, 6);
    
    // Lead history is no longer needed since the full conversation is pasted each time
    // The conversation text itself contains all previous messages and responses
    var history = [];
    
    // For follow-ups, use a different prompt structure
    var systemPrompt, userPrompt;
    if (payload.type === 'followup') {
      systemPrompt = buildFollowupSystemPrompt();
      // Get lead magnets and check usage
      var leadMagnets = [];
      var usedMagnets = [];
      var availableMagnets = [];
      try {
        leadMagnets = getLeadMagnets();
        usedMagnets = getUsedLeadMagnets_(payload.conversationId || '', history);
        availableMagnets = filterAvailableLeadMagnets_(leadMagnets, usedMagnets, extractKeywords_(payload.conversation));
      } catch (e) {
        console.error('Error getting lead magnets:', e);
        // Continue without lead magnets if there's an error
      }
      
      // Determine follow-up type
      var followupType = determineFollowupType_(payload.conversation, payload.callTranscript || '');
      
      // For follow-ups, still need to check for used lead magnets from data log
      var usedMagnets = [];
      if (payload.conversationId) {
        var followupHistory = getLeadHistory({ conversation_id: payload.conversationId });
        usedMagnets = getUsedLeadMagnets_(payload.conversationId || '', followupHistory);
      }
      // For follow-ups, extract keywords if not already done
      if (!aiKeywords) {
        aiKeywords = extractKeywordsWithAI(payload.conversation);
      }
      availableMagnets = filterAvailableLeadMagnets_(leadMagnets, usedMagnets, aiKeywords);
      
      userPrompt = buildFollowupUserPrompt(payload, selectedRules, selectedExamples, [], selectedPDFChunks, payload.callTranscript || '', availableMagnets, followupType);
    } else {
      systemPrompt = buildLeadSystemPrompt();
      userPrompt = buildLeadUserPrompt(payload, selectedRules, selectedExamples, [], selectedPDFChunks);
    }

    var apiKey = getOpenAIKey();
    if (!apiKey) return { success: false, error: 'OpenAI API key not configured' };

    // GPT-5/5.1 uses max_completion_tokens instead of max_tokens, and only supports temperature=1
    // Note: GPT-5.1-Codex is for Codex only, not Chat Completions API
    var isGPT5 = (typeof OPENAI_MODEL !== 'undefined' && (String(OPENAI_MODEL).toLowerCase().includes('gpt-5') || String(OPENAI_MODEL).toLowerCase().includes('gpt-5.1')));
    var maxTokensValue = (typeof OPENAI_MAX_TOKENS !== 'undefined' ? OPENAI_MAX_TOKENS : 700);
    
    var body = {
      model: OPENAI_MODEL,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ]
    };
    
    // Use max_completion_tokens for GPT-5, max_tokens for older models
    if (isGPT5) {
      body.max_completion_tokens = maxTokensValue;
      body.temperature = 1; // GPT-5 only supports default temperature (1)
    } else {
      body.max_tokens = maxTokensValue;
      body.temperature = (payload.type === 'followup') ? 0.7 : 0.5; // Custom temperature for older models
    }

    var resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    
    // Check HTTP response code
    var responseCode = resp.getResponseCode();
    if (responseCode !== 200) {
      return { success: false, error: 'HTTP error ' + responseCode + ': ' + resp.getContentText() };
    }
    
    // Parse JSON with error handling
    var data;
    try {
      data = JSON.parse(resp.getContentText());
    } catch (e) {
      return { success: false, error: 'Invalid JSON response from OpenAI: ' + String(e) };
    }
    
    if (data.error) {
      // If model not found, suggest alternatives
      var errorMsg = data.error.message || '';
      if (errorMsg.indexOf('model') !== -1 && errorMsg.indexOf('not found') !== -1) {
        return { success: false, error: 'Model "' + OPENAI_MODEL + '" not found. Try "gpt-5-turbo" or "gpt-5o" in Config.gs, or check OpenAI API documentation for the correct model name.' };
      }
      return { success: false, error: 'OpenAI error: ' + errorMsg };
    }
    
    // Validate response structure
    if (!data.choices || !data.choices[0] || !data.choices[0].message) {
      // Log the actual response for debugging
      var responsePreview = JSON.stringify(data).substring(0, 1000);
      Logger.log('Unexpected response structure. Full response: ' + responsePreview);
      return { 
        success: false, 
        error: 'Unexpected response structure from OpenAI API. The model "' + OPENAI_MODEL + '" may not exist or may have a different API structure. Try changing OPENAI_MODEL to "gpt-4o" in Config.gs. Response preview: ' + responsePreview.substring(0, 200)
      };
    }
    
    var rawAnswer = data.choices[0].message.content;
    
    // Validate that we got an answer
    if (!rawAnswer || typeof rawAnswer !== 'string') {
      // Log the actual response for debugging
      Logger.log('Empty or invalid response. Raw answer type: ' + typeof rawAnswer);
      var responsePreview = JSON.stringify(data).substring(0, 1000);
      Logger.log('Full response structure: ' + responsePreview);
      Logger.log('Choices[0]: ' + JSON.stringify(data.choices[0]).substring(0, 300));
      
      // Try alternative response structures (GPT-5 might differ)
      if (data.choices[0].content) {
        rawAnswer = data.choices[0].content;
      } else if (data.choices[0].text) {
        rawAnswer = data.choices[0].text;
      } else if (data.content) {
        rawAnswer = data.content;
      }
      
      // Final check
      if (!rawAnswer || typeof rawAnswer !== 'string') {
        return { 
          success: false, 
          error: 'OpenAI returned empty or invalid response. The model "' + OPENAI_MODEL + '" may not be available or may have a different response format. Try changing OPENAI_MODEL to "gpt-4o" in Config.gs. Response preview: ' + responsePreview.substring(0, 200)
        };
      }
    }

    // Check for contradiction warning BEFORE cleaning (so we can detect it)
    var contradictionWarning = /\[RULE_CONTRADICTION\s*:\s*YES\]/i.test(rawAnswer || '');

    // Clean the answer: remove any "Data Log:" text, "VA Next Steps", or other metadata that shouldn't be in the lead response
    var answer = String(rawAnswer);
    // Remove "Data Log:" and everything after it (case insensitive)
    answer = answer.replace(/\s*[Dd]ata\s+[Ll]og\s*:.*$/s, '');
    // Remove "VA Next Steps" sections
    answer = answer.replace(/\s*###?\s*VA\s+Next\s+Steps.*$/is, '');
    answer = answer.replace(/\s*VA\s+Next\s+Steps.*$/is, '');
    // Remove any trailing metadata patterns (including the contradiction marker)
    answer = answer.replace(/\s*\[RULE_CONTRADICTION\s*:\s*YES\].*$/i, '');
    // Trim whitespace
    answer = answer.trim();
    
    // Final validation - ensure we still have content after cleaning
    if (!answer || answer.length === 0) {
      return { success: false, error: 'Response was empty after cleaning metadata' };
    }

    var metaForRow = inferRowMetadata_(payload.conversation, answer);
    // Override context if this is a follow-up
    if (payload.type === 'followup') {
      metaForRow.context = 'Follow-up';
    }
    
    // Ensure metaForRow is an object with all required fields
    if (!metaForRow || typeof metaForRow !== 'object') {
      metaForRow = { 
        context: 'General Inquiry', 
        responseApproach: 'Unknown', 
        keyComponents: 'Unknown', 
        originalQuestion: 'Unknown',
        immediateContext: 'Unknown',
        contextSummary: 'Unknown'
      };
    }
    // Ensure immediateContext and contextSummary are set (fallback if missing)
    if (!metaForRow.immediateContext) {
      metaForRow.immediateContext = 'Unknown';
    }
    if (!metaForRow.contextSummary) {
      metaForRow.contextSummary = 'Unknown';
    }
    
    var row = buildCopyPasteRow(Object.assign({}, payload, metaForRow), answer);

    // track approximate daily cost and warn
    // Handle potential differences in usage structure for GPT-5
    var usage = data && data.usage;
    if (!usage && data && data.usage_info) {
      usage = data.usage_info; // Some models might use usage_info instead
    }
    // Track cost for generation query
    var costWarning = trackCostAndCheckLimit_(usage);
    
    // Calculate cost for keyword extraction, selection query, and generation query
    // Note: Keyword extraction cost is already tracked in extractKeywordsWithAI()
    var selectionUsage = selection.usage || {};
    var estimatedCost = computeCostUsd_(usage);
    var selectionCost = computeCostUsd_(selectionUsage);
    // Keyword extraction cost is minimal (~100 tokens) and already tracked, so we don't need to add it here
    var totalCost = estimatedCost + selectionCost;
    
    // Also track selection cost in daily limit (if selectionUsage exists)
    if (selectionUsage && selectionUsage.total_tokens) {
      trackCostAndCheckLimit_(selectionUsage);
    }
    
    var generationTime = new Date() - generationStartTime;
    var totalTime = new Date() - startTime;
    
    Logger.log('[Phase 5] Generation took ' + generationTime + 'ms');
    Logger.log('[Complete] Total time: ' + totalTime + 'ms');
    Logger.log('[Complete] Total cost: $' + totalCost.toFixed(4));
    Logger.log('[Complete] Examples used: ' + selectedExamples.length + ' (rows: ' + selectedExamples.map(function(e){ return e.row_number; }).join(', ') + ')');
    Logger.log('[Complete] Rules used: ' + selectedRules.length + ' (rows: ' + selectedRules.map(function(r){ return r.row_number; }).join(', ') + ')');
    Logger.log('=== Lead Response Generation Complete ===');

    var refs = {
      examples_rows: selectedExamples.map(function(e){ return e.row_number; }),
      rules_rows: selectedRules.map(function(r){ return r.row_number; }).filter(function(x){ return x != null; }),
      pdf_chunks: selectedPDFChunks.map(function(p){ 
        var pdfName = (p.category || '').replace(/^PDF - /, '');
        return { pdf: pdfName, section: p.rule_title, row: p.row_number }; 
      }),
      selection_cost: selectionCost,
      ai_keywords: aiKeywords
    };

    return { success: true, answer: answer, row: row, contradictionWarning: contradictionWarning, costWarning: costWarning, meta: { model: OPENAI_MODEL, usage: usage, estimated_cost_usd: totalCost, selection_cost_usd: selectionCost, generation_cost_usd: estimatedCost, refs: refs, generation_time_ms: generationTime, total_time_ms: totalTime } };
  } catch (e) {
    Logger.log('ERROR in generateLeadResponse: ' + String(e));
    Logger.log('Stack trace: ' + (e.stack || 'No stack trace'));
    return { success: false, error: String(e) };
  }
}

function buildLeadSystemPrompt() {
  // Enhanced system prompt for Lead Response Assistant
  return (
    'You are an expert Lead Response Assistant for Creekside Marketing, trained exclusively on "Creekside Marketing: A Strategic Compendium for AI-Powered Lead Qualification" as the single source of truth.\n\n' +
    
    'CORE IDENTITY:\n' +
    'You write as Samuel Rainey in Upwork message threads. Your voice is:\n' +
    '- Clear, confident, and diagnostic (not salesy or pushy)\n' +
    '- Zero fluff, zero filler - every word serves a purpose\n' +
    '- Casual professionalism: sharp, helpful, and authentically human\n' +
    '- Direct and to-the-point without being abrupt\n\n' +
    
    'CRITICAL FORMATTING RULES (NON-NEGOTIABLE):\n' +
    '- ABSOLUTELY NO MARKDOWN: Plain text only - no **bold**, no *italics*, no # headers, no - bullets, no numbered lists with markdown\n' +
    '- NEVER use asterisks (*), double asterisks (**), hash symbols (#), or dashes (-) for formatting\n' +
    '- Use ONLY plain text paragraphs separated by line breaks\n' +
    '- If organizing information, use natural language flow, not lists or headers\n' +
    '- Examples show plain text - use them for cadence, tone, language, and vocabulary, NOT for exact formatting replication\n' +
    '- Structured for direct copy/paste into Upwork\n' +
    '- Never use em-dashes (use regular hyphens or commas instead)\n' +
    '- Never restate the user\'s question\n' +
    '- Never include calendar links or suggest specific times\n\n' +
    
    'RESPONSE STRATEGY:\n' +
    '- Answer ALL client messages/questions that came after Samuel\'s last response\n' +
    '- Use past conversation context to inform your response, but focus on the most recent messages\n' +
    '- Be comprehensive but concise - answer thoroughly without word vomiting\n' +
    '- If brief is sufficient, be brief. If detail is needed, provide detail. Don\'t force either.\n' +
    '- Match the complexity and depth of the client\'s question\n' +
    '- Use diagnostic-first approach: understand their situation before proposing solutions\n\n' +
    
    'PRICING HANDLING:\n' +
    '- Default: Ask them to hop on a call to discuss pricing\n' +
    '- EXCEPTION: If they\'ve already had a call, you can provide pricing\n' +
    '- When providing pricing: Note that pricing varies business-to-business\n' +
    '- Important: I don\'t bill hourly for campaign management. Fees are based on complexity and ad spend.\n' +
    '- Budget questions: Respond with platform-specific ad spend recommendations (testing/scaling), not management fees\n\n' +
    
    'VOICE AND TONE:\n' +
    '- Match Samuel\'s phrasing and structure from real Upwork chats\n' +
    '- Pivots, qualification, and objection handling follow the diagnostic-first approach\n' +
    '- Be helpful and insightful, not just responsive\n' +
    '- Show expertise through clarity, not jargon\n\n' +
    
    'MANDATORY WORKFLOW:\n' +
    '1. Review ALL company rules - they are critical and must be followed\n' +
    '2. Review response examples for cadence, tone, language, and approach\n' +
    '3. Generate response that follows rules and matches example style\n' +
    '4. If rules contradict, the most recent rule takes priority\n' +
    '5. Spreadsheet rules override any conflicting general instructions\n' +
    '6. If you detect a rule contradiction, include [RULE_CONTRADICTION:YES] in your response\n\n' +
    
    'OUTPUT REQUIREMENTS:\n' +
    '- Output ONLY the response text - no metadata, no "Data Log:", no "VA Next Steps"\n' +
    '- Ready for direct copy/paste to the lead\n' +
    '- Never reference documents, citations, PDFs, or line numbers\n' +
    '- All output must match the Strategic Compendium in tone, content, and structure'
  );
}

function buildLeadUserPrompt(payload, rules, examples, history, pdfChunks) {
  var parts = [];
  parts.push('# Conversation');
  parts.push('CRITICAL: The conversation above contains the ENTIRE conversation thread, including all previous messages and responses.');
  parts.push('');
  parts.push('RESPONSE LOGIC (Follow these steps in order):');
  parts.push('1. Identify which messages are from you (Samuel Rainey) vs from the client');
  parts.push('2. Find the LAST response you (Samuel) sent - this is your cutoff point');
  parts.push('3. Identify ALL client messages (questions, statements, concerns) that came AFTER your last response');
  parts.push('4. Respond to ALL of those recent client messages comprehensively in a single response');
  parts.push('5. Do NOT re-answer questions already addressed in your previous responses');
  parts.push('6. Prioritize the most recent messages, but address all client messages since your last response');
  parts.push('');
  parts.push('CONVERSATION:');
  parts.push(String(payload.conversation || ''));
  parts.push('\n# Company Rules (CRITICAL - all rules must be followed)');
  parts.push('All company rules are critical. If rules contradict, the most recent rule takes priority.');
  parts.push('Pricing and service rules will typically be found here or in conversation examples.');
  if (rules && Array.isArray(rules) && rules.length > 0) {
    rules.forEach(function(r){
      if (r) {
        parts.push('- [' + (r.category || 'General') + '] ' + (r.rule_title || '') + ': ' + (r.rule_description || ''));
      }
    });
  } else {
    parts.push('(No company rules found)');
  }
  parts.push('\n# Relevant Response Examples (weighted by recency)');
  parts.push('These examples are selected because they are relevant to the question being asked.');
  parts.push('Use these examples as a guide for:');
  parts.push('- Cadence: The rhythm and flow of how messages are structured');
  parts.push('- Tone of voice: The level of formality, confidence, and professionalism');
  parts.push('- Language and vocabulary: The words and phrases used');
  parts.push('- Approach: How questions are answered and information is presented');
  parts.push('Note: These examples show plain text formatting (no markdown) - follow that formatting style, but focus on matching the tone and cadence rather than exact structure.');
  if (examples && Array.isArray(examples) && examples.length > 0) {
    examples.forEach(function(ex){
      if (ex && ex.full_response) {
        parts.push('\n--- Example Response ---');
        parts.push(ex.full_response || '');
        parts.push('--- End Example ---\n');
      }
    });
  } else {
    parts.push('(No relevant examples found - use company rules and general guidelines)');
  }
  if (pdfChunks && Array.isArray(pdfChunks) && pdfChunks.length > 0) {
    parts.push('\n# PDF Knowledge Base (de-prioritized - use only if relevant)');
    parts.push('These PDF sections are less critical than company rules above. Only reference if directly relevant.');
    pdfChunks.forEach(function(chunk){
      if (chunk) {
        var pdfName = (chunk.category || '').replace(/^PDF - /, '');
        parts.push('## ' + pdfName + ' - ' + (chunk.rule_title || 'Section'));
        parts.push(chunk.rule_description || '');
        parts.push('');
      }
    });
  }
  parts.push('\n# Output Requirements');
  parts.push('Output ONLY the response to the lead. Plain text only, no formatting, no bullets, no markdown, structured for Upwork copy/paste.');
  parts.push('CRITICAL FORMATTING RULES:');
  parts.push('- NEVER use markdown: no **bold**, no *italics*, no # headers, no - bullets, no numbered lists with markdown');
  parts.push('- NEVER use asterisks (*) or double asterisks (**) for formatting');
  parts.push('- NEVER use hash symbols (#) for headers');
  parts.push('- NEVER use dashes (-) or asterisks (*) for bullet points');
  parts.push('- Use ONLY plain text paragraphs separated by line breaks');
  parts.push('- If you need to organize information, use plain text with line breaks');
  parts.push('- The response examples show plain text formatting - follow that style, but focus on matching their cadence, tone, and language rather than exact structure');
  parts.push('DO NOT include "Data Log:", "Data Log row:", or any other metadata in your response.');
  parts.push('DO NOT include "VA Next Steps" or any internal instructions in your response.');
  parts.push('The response should be ready to copy and paste directly to the lead - nothing else.');
  parts.push('\n# Response Guidelines');
  parts.push('CRITICAL: Follow the response logic above. Additional guidance:');
  parts.push('- Use the conversation as your primary source - it contains everything');
  parts.push('- Address all recent client messages comprehensively in one response');
  parts.push('- Use past context to inform your response, but focus on what\'s new');
  parts.push('- Match the complexity of the client\'s question - brief for simple, detailed for complex');
  parts.push('- Reference examples for cadence, tone, language, and vocabulary - not exact formatting');
  parts.push('- Be helpful and insightful, showing expertise through clarity');
  parts.push('\n# Non-Negotiable Policies');
  parts.push('- All company rules are critical - follow them strictly');
  parts.push('- If rules contradict, the most recent rule takes priority');
  parts.push('- Response must be plain text only - no formatting, bullets, markdown, or citations');
  parts.push('- Never restate the user\'s question in your response');
  parts.push('- Never use em-dashes (use regular hyphens instead)');
  return parts.join('\n');
}

function buildFollowupSystemPrompt() {
  return (
    'You are an expert Follow-up Message Assistant for Creekside Marketing. Your job is to craft context-driven follow-up messages that are natural, helpful, and maintain engagement with leads.\n\n' +
    
    'CORE IDENTITY:\n' +
    'You write as Samuel Rainey in Upwork message threads. Your voice is:\n' +
    '- Clear, confident, and diagnostic (not salesy or pushy)\n' +
    '- Zero fluff, zero filler - every word serves a purpose\n' +
    '- Casual professionalism: sharp, helpful, and authentically human\n' +
    '- Direct and to-the-point without being abrupt\n\n' +
    
    'CRITICAL FORMATTING RULES (NON-NEGOTIABLE):\n' +
    '- ABSOLUTELY NO MARKDOWN: Plain text only - no **bold**, no *italics*, no # headers, no - bullets, no numbered lists with markdown\n' +
    '- NEVER use asterisks (*), double asterisks (**), hash symbols (#), or dashes (-) for formatting\n' +
    '- Use ONLY plain text paragraphs separated by line breaks\n' +
    '- If organizing information, use natural language flow, not lists or headers\n' +
    '- Examples show plain text - use them for cadence, tone, language, and vocabulary, NOT for exact formatting replication\n' +
    '- Structured for direct copy/paste into Upwork\n' +
    '- Never use em-dashes (use regular hyphens or commas instead)\n' +
    '- Never include a calendar link or suggest times someone is available\n' +
    '- Informal and direct - a few sentences is good\n' +
    '- Context-driven - show you care about the client and project\n' +
    '- Reference specific examples or pain points from the conversation\n\n' +
    'Follow-up types:\n' +
    '- Pre-call: Encourage the person to continue talking about their business and potentially hop on a call to discuss it in more detail.\n' +
    '- Post-call: Encourage them to explain any objections they might have. Encourage them to at least reengage with us about the project, to ask questions that may be causing them to have doubts about working with us.\n' +
    '- Lost: Encourage them to give us updates on how the project went since we last talked to them.\n\n' +
    'Lead magnets:\n' +
    '- A lead magnet is contextually appropriate if it specifically solves one of their pain points or objections, or helps them realize the value of our services\n' +
    '- Example: If they\'re interested in Google Ads, only send them lead magnets relevant to Google Ads. Same for Facebook.\n' +
    '- Never force a lead magnet - only include if it naturally fits and adds value\n\n' +
    'Goal: Re-engage, not sell. The follow-up should not try to sell something; it\'s just to get them to re-engage.\n\n' +
    'You can analyze sentiment and engagement levels to determine if they\'re ready to hop on a call to discuss further vs. if they need more questions answered before being ready to hop on a call.\n\n' +
    'Spreadsheet rules override any conflicting general instructions. Never contradict rules; if conflict is detected, regenerate to comply and include [RULE_CONTRADICTION:YES].'
  );
}

function buildFollowupUserPrompt(payload, rules, examples, history, pdfChunks, callTranscript, availableMagnets, followupType) {
  var parts = [];
  parts.push('# Follow-up Type: ' + followupType);
  parts.push('# Full Conversation');
  parts.push('Review the entire conversation to understand context and ensure variety. Reference the entire conversation to see what follow-ups have already been used.');
  parts.push(String(payload.conversation));
  
  if (callTranscript && callTranscript.trim()) {
    parts.push('\n# Call Transcript');
    parts.push('The following is a transcript from a sales call that occurred:');
    parts.push(callTranscript);
    parts.push('\nCRITICAL: The biggest things to address from the call transcript are OBJECTIONS and PAIN POINTS.');
    parts.push('Do NOT summarize the call. Instead, use specific objections and pain points from the transcript to inform your follow-up.');
    parts.push('Reference these specific points to show you were listening and care about their concerns.');
  } else if (followupType === 'post-call') {
    parts.push('\n# Note: This is a post-call follow-up, but no transcript was provided.');
    parts.push('Reference the call conversation based on context clues in the conversation above, focusing on objections and pain points mentioned.');
  }
  
  if (history && Array.isArray(history) && history.length > 0) {
    parts.push('\n# Previous Follow-ups (most recent first)');
    parts.push('CRITICAL: Review ALL previous follow-ups to ensure variety. Avoid sounding repetitive or having a similar cadence to any follow-up sent in the past.');
    history.forEach(function(h){
      if (h && h.full_response) {
        parts.push('- [' + (h.date || '') + '] ' + (h.full_response || ''));
      }
    });
    parts.push('\nIMPORTANT: Do not repeat the same approach, lead magnet, or messaging style used in previous follow-ups.');
    parts.push('Ensure your follow-up has a different cadence and approach than what has been sent before.');
  } else {
    parts.push('\n# Previous Follow-ups');
    parts.push('No previous follow-ups found for this conversation. This may be the first follow-up.');
  }
  
  if (availableMagnets && Array.isArray(availableMagnets) && availableMagnets.length > 0) {
    parts.push('\n# Available Lead Magnets (use only if contextually appropriate)');
    parts.push('A lead magnet is appropriate if it specifically solves one of their pain points or objections, or helps them realize the value of our services.');
    parts.push('Example: If they\'re interested in Google Ads, only use Google Ads-related lead magnets. Same for Facebook.');
    availableMagnets.forEach(function(m){
      if (m && m.title) {
        // Display as: Title: URL (Description if available)
        var display = m.title;
        if (m.url) {
          display += ': ' + m.url;
        }
        if (m.description && m.description.trim()) {
          display += ' - ' + m.description;
        }
        parts.push('- ' + display);
      }
    });
    parts.push('\nNote: Only include a lead magnet if it naturally fits the context and directly addresses their pain points or objections. Do not force it.');
  }
  
  if (pdfChunks && Array.isArray(pdfChunks) && pdfChunks.length > 0) {
    parts.push('\n# PDF Knowledge Base (relevant sections)');
    pdfChunks.forEach(function(chunk){
      if (chunk) {
        var pdfName = (chunk.category || '').replace(/^PDF - /, '');
        parts.push('## ' + pdfName + ' - ' + (chunk.rule_title || 'Section'));
        parts.push(chunk.rule_description || '');
        parts.push('');
      }
    });
  }
  
  parts.push('\n# Company Rules');
  if (rules && Array.isArray(rules) && rules.length > 0) {
    rules.forEach(function(r){
      if (r) {
        parts.push('- [' + (r.category || 'General') + '] ' + (r.rule_title || '') + ': ' + (r.rule_description || ''));
      }
    });
  } else {
    parts.push('(No company rules found)');
  }
  
  var maxExamples = (typeof EXAMPLES_IN_PROMPT !== 'undefined' ? EXAMPLES_IN_PROMPT : 20);
  parts.push('\n# Relevant Response Examples (up to ' + maxExamples + ', weighted by recency)');
  parts.push('These examples are selected because they are relevant to the conversation context. Use them to inform your follow-up style and approach.');
  if (examples && Array.isArray(examples) && examples.length > 0) {
    // Limit examples for follow-ups (use EXAMPLES_IN_PROMPT constant)
    var followupExamples = examples.slice(0, maxExamples);
    followupExamples.forEach(function(ex){
      if (ex) {
        parts.push('\n## Example ' + (ex.row_number || '') + ':');
        if (ex.original_question) parts.push('Question: ' + ex.original_question);
        if (ex.full_response) parts.push('Response: ' + ex.full_response);
        if (ex.response_approach) parts.push('Approach: ' + ex.response_approach);
      }
    });
  } else {
    parts.push('(No relevant examples found - use company rules and general guidelines)');
  }
  
  parts.push('\n# Output Requirements');
  parts.push('Generate a context-driven follow-up message that:');
  parts.push('1. Is appropriate for the follow-up type (' + followupType + ')');
  if (followupType === 'pre-call') {
    parts.push('   - Encourages them to continue talking about their business');
    parts.push('   - Potentially suggests hopping on a call to discuss in more detail');
  } else if (followupType === 'post-call') {
    parts.push('   - Encourages them to explain any objections they might have');
    parts.push('   - Encourages them to reengage about the project');
    parts.push('   - Asks questions that may be causing them to have doubts');
  } else if (followupType === 'lost') {
    parts.push('   - Encourages them to give updates on how the project went since we last talked');
  }
  if (callTranscript && callTranscript.trim()) {
    parts.push('2. References specific OBJECTIONS and PAIN POINTS from the call transcript (do not summarize the call)');
  } else {
    parts.push('2. References specific examples or pain points from the conversation');
  }
  parts.push('3. Shows you care about the client and project by referencing specific details');
  parts.push('4. Is informal and direct - a few sentences is good');
  parts.push('5. Only includes a lead magnet if it specifically solves their pain points or objections');
  parts.push('6. Does not repeat approaches, cadence, or messaging style from previous follow-ups');
  parts.push('7. Goal is to re-engage, not to sell something');
  parts.push('8. Is plain text only - ABSOLUTELY NO FORMATTING: no markdown, no asterisks, no bold, no bullets, no headers');
  parts.push('   - NEVER use **text** or *text* or # headers or - bullets or numbered lists with markdown');
  parts.push('   - Use ONLY plain text paragraphs separated by line breaks');
  parts.push('   - Use the response examples as a guide for cadence, tone, language, and vocabulary - not for exact formatting replication');
  parts.push('9. Is ready to copy and paste directly to the lead');
  
  parts.push('\n# Variety and Uniqueness');
  parts.push('- Reference the entire conversation to see what has been discussed');
  parts.push('- Review all previous follow-ups to ensure your message has a different cadence and approach');
  parts.push('- Avoid sounding repetitive or using similar phrasing to past follow-ups');
  parts.push('- Make it feel personal and context-driven, not templated');
  
  parts.push('\n# Non-Negotiable Policies');
  parts.push('- Spreadsheet rules override any conflicting guidance');
  parts.push('- Response must be plain text only - no formatting, bullets, markdown, or citations');
  parts.push('- Never use em-dashes (use regular hyphens instead)');
  parts.push('- Never include "Data Log:", "VA Next Steps", or any metadata in your response');
  
  return parts.join('\n');
}

function buildCopyPasteRow(payload, fullResponse) {
  // IMPORTANT: This function must create exactly 11 fields matching the sheet columns:
  // 1. Date, 2. Conversation_ID, 3. Turn_Index, 4. Conversation_Context, 5. Industry,
  // 6. Response_Approach, 7. Key_Components, 8. Original_Question, 9. Immediate_Context,
  // 10. Context_Summary, 11. Full_Response
  // DO NOT add or remove fields - the sheet has exactly these 11 columns.
  
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var conversationId = payload.conversationId || '';
  var turnIndex = payload.turnIndex != null ? payload.turnIndex : 1;
  var context = payload.context || 'General Inquiry';
  var industry = payload.industry || 'Unknown';
  var responseApproach = payload.responseApproach || 'Unknown';
  var keyComponents = payload.keyComponents || 'Unknown';
  var originalQuestion = payload.originalQuestion || 'Unknown';
  var immediateContext = payload.immediateContext || '';
  var contextSummary = payload.contextSummary || '';

  // Replace newlines with spaces in the response, but keep pipes as-is for display
  // The saveGeneratedRow function will handle pipes intelligently
  var fields = [
    dateStr,                    // 1. Date
    conversationId,              // 2. Conversation_ID
    turnIndex,                  // 3. Turn_Index
    context,                    // 4. Conversation_Context
    industry,                   // 5. Industry
    responseApproach,           // 6. Response_Approach
    keyComponents,              // 7. Key_Components
    originalQuestion,           // 8. Original_Question
    immediateContext,           // 9. Immediate_Context
    contextSummary,             // 10. Context_Summary
    (fullResponse || '').replace(/\n/g, ' ')  // 11. Full_Response
  ];
  
  // Ensure exactly 11 fields - no more, no less
  if (fields.length !== 11) {
    throw new Error('buildCopyPasteRow must return exactly 11 fields');
  }
  
  return fields.join('|');
}

function getOpenAIKey() {
  try {
    var props = PropertiesService.getScriptProperties();
    var key = props && props.getProperty('OPENAI_API_KEY');
    if (key) return key.trim();
  } catch (e) {}
  // Load from spreadsheet OpenAI tab B1 if available
  try {
    var ss = SpreadsheetApp.openById(LEADS_SPREADSHEET_ID);
    var tab = ss.getSheetByName(OPENAI_TAB_NAME);
    if (tab) {
      var cell = (tab.getRange(1, 2).getValue() || '').toString().trim();
      if (cell) return cell;
    }
  } catch (e2) {}
  return OPENAI_API_KEY || null;
}

// ================================
// Cost tracking helpers
// ================================

function trackCostAndCheckLimit_(usage) {
  try {
    if (!usage) return false;
    // Use the same calculation as computeCostUsd_ for consistency
    var inputTokens = (usage.prompt_tokens || 0);
    var outputTokens = (usage.completion_tokens || 0);
    var inputRate = (typeof COST_PER_1K_INPUT_TOKENS_USD !== 'undefined') ? COST_PER_1K_INPUT_TOKENS_USD : 0.0025;
    var outputRate = (typeof COST_PER_1K_OUTPUT_TOKENS_USD !== 'undefined') ? COST_PER_1K_OUTPUT_TOKENS_USD : 0.01;
    var cost = ((inputTokens / 1000.0) * inputRate) + ((outputTokens / 1000.0) * outputRate);
    var tz = Session.getScriptTimeZone();
    var ymd = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var props = PropertiesService.getScriptProperties();
    var currentYmd = props.getProperty('COST_YMD');
    var acc = parseFloat(props.getProperty('COST_USD') || '0');
    if (currentYmd !== ymd) {
      acc = 0; currentYmd = ymd;
    }
    acc += cost;
    props.setProperty('COST_YMD', currentYmd);
    props.setProperty('COST_USD', String(acc));
    var limit = (typeof DAILY_COST_LIMIT_USD !== 'undefined') ? DAILY_COST_LIMIT_USD : 5;
    return acc > limit;
  } catch (e) {
    return false;
  }
}

function computeCostUsd_(usage) {
  try {
    if (!usage) return 0;
    
    // GPT-4o has different pricing for input vs output tokens
    // Input: $2.50 per million tokens = $0.0025 per 1K tokens
    // Output: $10.00 per million tokens = $0.01 per 1K tokens
    var inputTokens = usage.prompt_tokens || 0;
    var outputTokens = usage.completion_tokens || 0;
    
    // Use separate input/output rates if available, otherwise fall back to flat rate
    var inputRate = (typeof COST_PER_1K_INPUT_TOKENS_USD !== 'undefined') ? COST_PER_1K_INPUT_TOKENS_USD : 0.0025;
    var outputRate = (typeof COST_PER_1K_OUTPUT_TOKENS_USD !== 'undefined') ? COST_PER_1K_OUTPUT_TOKENS_USD : 0.01;
    
    var inputCost = (inputTokens / 1000.0) * inputRate;
    var outputCost = (outputTokens / 1000.0) * outputRate;
    var totalCost = inputCost + outputCost;
    
    return Math.round(totalCost * 1000) / 1000; // round to mills
  } catch (e) {
    return 0;
  }
}

// ================================
// Enhanced Keyword Matching
// ================================

/**
 * Enhanced keyword matching with synonyms, stemming approximation, and intelligent scoring
 * Returns: Filtered and scored rows sorted by relevance
 */
function matchExamplesWithKeywords_(rows, keywords) {
  try {
    // Expand keywords with synonyms and variations
    var expandedKeywords = expandKeywordsWithSynonyms_(keywords);
    
    return rows.map(function(r){
      // Build searchable text from all relevant fields
      var searchFields = {
        conversation_context: (r.conversation_context || '').toLowerCase(),
        industry: (r.industry || '').toLowerCase(),
        response_approach: (r.response_approach || '').toLowerCase(),
        key_components: (r.key_components || '').toLowerCase(),
        original_question: (r.original_question || '').toLowerCase(),
        context_summary: (r.context_summary || '').toLowerCase(),
        full_response: (r.full_response || '').toLowerCase()
      };
      
      var combinedText = Object.values(searchFields).join(' ');
      var score = 0;
      var matchedKeywords = [];
      
      // Score each keyword and its variations
      expandedKeywords.forEach(function(keywordGroup){
        var keyword = keywordGroup.base;
        var variations = keywordGroup.variations;
        var found = false;
        var matchScore = 0;
        
        // Check base keyword
        if (combinedText.indexOf(keyword) !== -1) {
          found = true;
          matchScore = 1;
          
          // Bonus points for matches in important fields
          if (searchFields.full_response.indexOf(keyword) !== -1) matchScore += 3; // Full response is most important
          if (searchFields.response_approach.indexOf(keyword) !== -1) matchScore += 2;
          if (searchFields.conversation_context.indexOf(keyword) !== -1) matchScore += 2;
          if (searchFields.key_components.indexOf(keyword) !== -1) matchScore += 1;
          if (searchFields.original_question.indexOf(keyword) !== -1) matchScore += 1;
        }
        
        // Check variations/synonyms
        if (!found) {
          for (var i = 0; i < variations.length; i++) {
            var variation = variations[i];
            if (combinedText.indexOf(variation) !== -1) {
              found = true;
              matchScore = 0.8; // Slightly lower score for variations
              
              // Bonus points for variations in important fields
              if (searchFields.full_response.indexOf(variation) !== -1) matchScore += 2.5;
              if (searchFields.response_approach.indexOf(variation) !== -1) matchScore += 1.5;
              if (searchFields.conversation_context.indexOf(variation) !== -1) matchScore += 1.5;
              break;
            }
          }
        }
        
        if (found) {
          score += matchScore;
          matchedKeywords.push(keyword);
        }
      });
      
      // Additional scoring factors
      var uniqueMatches = matchedKeywords.length;
      var matchDensity = uniqueMatches / Math.max(expandedKeywords.length, 1); // Percentage of keywords matched
      
      // Boost score for matching multiple keywords (shows stronger relevance)
      if (uniqueMatches >= 3) score *= 1.3;
      else if (uniqueMatches >= 2) score *= 1.15;
      
      // Boost score for high match density
      if (matchDensity >= 0.7) score *= 1.2;
      else if (matchDensity >= 0.5) score *= 1.1;
      
      return Object.assign({}, r, { 
        relevance_score: Math.round(score * 100) / 100, // Round to 2 decimals
        matched_keywords: matchedKeywords.join(', ')
      });
    }).filter(function(r){ return r.relevance_score > 0; })
      .sort(function(a,b){ 
        // Sort by relevance score (highest first)
        if (b.relevance_score !== a.relevance_score) {
          return b.relevance_score - a.relevance_score;
        }
        // If scores are equal, prefer more recent
        return new Date(b.date) - new Date(a.date);
      });
  } catch (e) {
    Logger.log('Error in matchExamplesWithKeywords_: ' + String(e));
    // Fallback to simple matching
    return rows;
  }
}

/**
 * Expand keywords with synonyms, variations, and stemming approximations
 * Returns: Array of {base: keyword, variations: [synonyms, stems, etc.]}
 */
function expandKeywordsWithSynonyms_(keywords) {
  try {
    var keywordMap = {
      // Pricing related
      'pricing': {base: 'pricing', variations: ['price', 'cost', 'fee', 'rate', 'retainer', 'budget', 'billing', 'hourly', 'monthly']},
      'price': {base: 'pricing', variations: ['pricing', 'cost', 'fee', 'rate', 'retainer', 'budget', 'billing']},
      'cost': {base: 'pricing', variations: ['pricing', 'price', 'fee', 'rate', 'budget']},
      'budget': {base: 'pricing', variations: ['pricing', 'price', 'cost', 'spend', 'investment']},
      
      // Services
      'google': {base: 'google', variations: ['google ads', 'adwords', 'ppc', 'search ads']},
      'ads': {base: 'ads', variations: ['advertising', 'advertisements', 'campaigns', 'adwords']},
      'meta': {base: 'meta', variations: ['facebook', 'instagram', 'fb', 'social media']},
      'facebook': {base: 'meta', variations: ['meta', 'fb', 'instagram', 'social media']},
      'tracking': {base: 'tracking', variations: ['analytics', 'attribution', 'conversion tracking', 'pixel', 'ga4']},
      'analytics': {base: 'tracking', variations: ['tracking', 'attribution', 'ga4', 'google analytics']},
      
      // Business types
      'ecommerce': {base: 'ecommerce', variations: ['e-commerce', 'online store', 'shopify', 'retail', 'online retail']},
      'saas': {base: 'saas', variations: ['software', 'subscription', 'b2b software']},
      'healthcare': {base: 'healthcare', variations: ['medical', 'health', 'clinic', 'hospital']},
      
      // Request types
      'audit': {base: 'audit', variations: ['review', 'analysis', 'assessment', 'evaluation', 'check']},
      'onboarding': {base: 'onboarding', variations: ['setup', 'start', 'begin', 'initial', 'kickoff', 'getting started']},
      'setup': {base: 'onboarding', variations: ['onboarding', 'start', 'begin', 'initial', 'configuration']},
      'quote': {base: 'quote', variations: ['proposal', 'estimate', 'pricing', 'bid']},
      'consultation': {base: 'consultation', variations: ['consult', 'advice', 'guidance', 'strategy', 'discussion']},
      
      // Problems/pain points
      'roas': {base: 'roas', variations: ['return on ad spend', 'return on investment', 'roi', 'performance']},
      'scaling': {base: 'scaling', variations: ['scale', 'growth', 'expand', 'increase']},
      'attribution': {base: 'attribution', variations: ['tracking', 'conversion tracking', 'analytics']},
      'conversion': {base: 'conversion', variations: ['conversions', 'leads', 'sales', 'purchases']},
      
      // General
      'management': {base: 'management', variations: ['manage', 'managing', 'managed', 'oversight']},
      'optimization': {base: 'optimization', variations: ['optimize', 'optimizing', 'improve', 'improvement']},
      'campaign': {base: 'campaign', variations: ['campaigns', 'ad campaign', 'advertising campaign']}
    };
    
    var kws = String(keywords).toLowerCase()
      .replace(/[^a-z0-9\s]/g, ' ')
      .split(/\s+/)
      .filter(function(k){ return k.length > 2; });
    
    var expanded = [];
    var seen = {};
    
    kws.forEach(function(kw){
      // Check if we have a mapping for this keyword
      var mapped = keywordMap[kw];
      if (mapped && !seen[mapped.base]) {
        expanded.push(mapped);
        seen[mapped.base] = true;
      } else if (!seen[kw]) {
        // No mapping found, use keyword as-is with basic variations
        var variations = generateStemVariations_(kw);
        expanded.push({base: kw, variations: variations});
        seen[kw] = true;
      }
    });
    
    return expanded;
  } catch (e) {
    Logger.log('Error in expandKeywordsWithSynonyms_: ' + String(e));
    // Fallback: return keywords as-is
    var kws = String(keywords).toLowerCase().split(' ').filter(function(k){ return k.length > 2; });
    return kws.map(function(k){ return {base: k, variations: []}; });
  }
}

/**
 * Generate basic stem variations (approximate stemming)
 */
function generateStemVariations_(word) {
  var variations = [];
  
  // Common suffixes to try removing/adding
  var suffixes = ['ing', 'ed', 'er', 'est', 'ly', 's', 'es', 'tion', 'sion'];
  
  // Try removing suffixes
  suffixes.forEach(function(suffix){
    if (word.endsWith(suffix) && word.length > suffix.length + 2) {
      var stem = word.substring(0, word.length - suffix.length);
      if (stem.length > 2) variations.push(stem);
    }
  });
  
  // Try adding common suffixes
  if (word.length > 3) {
    variations.push(word + 'ing');
    variations.push(word + 'ed');
    variations.push(word + 's');
  }
  
  return variations.slice(0, 3); // Limit to 3 variations
}

// ================================
// Utility
// ================================

/**
 * Extract keywords using simple regex (fallback method)
 * @deprecated Use extractKeywordsWithAI() for better semantic keyword extraction
 */
function extractKeywords_(text) {
  try {
    var s = (text || '').toLowerCase();
    var words = s.replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(function(w){ return w.length > 3; });
    var seen = {};
    var uniq = [];
    words.forEach(function(w){ if (!seen[w]) { seen[w] = true; uniq.push(w); } });
    return uniq.slice(0, 8).join(' ');
  } catch (e) { return ''; }
}

/**
 * Extract semantic keywords from conversation using AI
 * Returns: String of keywords separated by spaces
 */
function extractKeywordsWithAI(conversation) {
  try {
    var apiKey = getOpenAIKey();
    if (!apiKey) {
      // Fallback to regex if no API key
      return extractKeywords_(conversation);
    }

    var systemPrompt = 'You are an expert keyword extraction assistant. Your job is to extract the most relevant, searchable keywords from a conversation that will help find similar response examples in a database.\n\n' +
      'CRITICAL: Extract keywords that are:\n' +
      '1. Searchable - Use terms that would appear in example responses (not just conversation-specific phrases)\n' +
      '2. Semantic - Include synonyms and related terms (e.g., if "pricing" is mentioned, also include "cost", "rate", "budget")\n' +
      '3. Specific - Include industry, service type, and business model when mentioned\n' +
      '4. Problem-focused - Include pain points, challenges, and goals\n\n' +
      'Extract 12-18 keywords covering:\n' +
      '- Main topics/themes (e.g., "pricing", "audit", "onboarding", "strategy")\n' +
      '- Industry/business type (e.g., "ecommerce", "SaaS", "healthcare", "B2B")\n' +
      '- Services/platforms (e.g., "Google Ads", "Meta", "tracking", "analytics")\n' +
      '- Pain points/problems (e.g., "low ROAS", "attribution", "scaling", "conversion")\n' +
      '- Intent/request type (e.g., "quote", "consultation", "setup", "optimization")\n' +
      '- Business metrics (e.g., "ROAS", "CPA", "conversion rate")\n\n' +
      'Return ONLY a space-separated list of keywords. No explanations, no formatting, no punctuation. Just keywords.';

    var userPrompt = 'Extract keywords from this conversation:\n\n' + conversation;

    var body = {
      model: OPENAI_MODEL,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ],
      max_tokens: 100,
      temperature: 0.3
    };

    // GPT-5/5.1 uses max_completion_tokens instead of max_tokens
    var isGPT5 = (typeof OPENAI_MODEL !== 'undefined' && (String(OPENAI_MODEL).toLowerCase().includes('gpt-5') || String(OPENAI_MODEL).toLowerCase().includes('gpt-5.1')));
    if (isGPT5) {
      body.max_completion_tokens = 100;
      body.temperature = 1;
      delete body.max_tokens;
    }

    var resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    var responseCode = resp.getResponseCode();
    if (responseCode !== 200) {
      Logger.log('AI keyword extraction failed: ' + responseCode + ' - ' + resp.getContentText());
      // Fallback to regex
      return extractKeywords_(conversation);
    }

    var data = JSON.parse(resp.getContentText());
    if (data.error || !data.choices || !data.choices[0] || !data.choices[0].message) {
      Logger.log('AI keyword extraction error: ' + JSON.stringify(data));
      // Fallback to regex
      return extractKeywords_(conversation);
    }

    var keywords = data.choices[0].message.content.trim();
    // Clean up: remove any formatting, newlines, extra spaces
    keywords = keywords.replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim();
    
    // Track cost for keyword extraction
    if (data.usage) {
      trackCostAndCheckLimit_(data.usage);
    }

    return keywords || extractKeywords_(conversation); // Fallback if empty
  } catch (e) {
    Logger.log('Error in extractKeywordsWithAI: ' + String(e));
    // Fallback to regex
    return extractKeywords_(conversation);
  }
}

// ================================
// Heuristics to fill Data Log fields when VA provides only conversation
// ================================

function inferRowMetadata_(conversation, answer) {
  try {
    var text = String(conversation || '');
    var originalQ = inferOriginalQuestion_(text) || '';

    var lc = text.toLowerCase() + ' ' + String(answer || '').toLowerCase();
    
    // More flexible context inference - check for various conversation types
    var context = 'General Inquiry'; // Default
    if (/(price|pricing|rate|cost|budget|fee|retainer|hourly|monthly)/.test(lc)) {
      context = 'Pricing inquiry';
    } else if (/(audit|review|analysis|assessment|evaluation)/.test(lc)) {
      context = 'Audit request';
    } else if (/(onboard|setup|start|begin|initial|kickoff)/.test(lc)) {
      context = 'Onboarding inquiry';
    } else if (/(quote|proposal|estimate|scope|project)/.test(lc)) {
      context = 'Project inquiry';
    } else if (/(strategy|consult|advice|recommend|guidance)/.test(lc)) {
      context = 'Strategy consultation';
    } else if (/(problem|issue|challenge|struggling|help)/.test(lc)) {
      context = 'Problem-solving inquiry';
    } else if (/(follow.?up|check.?in|update|status)/.test(lc)) {
      context = 'Follow-up';
    } else if (/(objection|concern|hesitation|doubt|worry)/.test(lc)) {
      context = 'Objection handling';
    } else if (/(qualification|fit|suitable|right|match)/.test(lc)) {
      context = 'Qualification inquiry';
    }
    
    // Infer response approach and key components based on context
    var responseApproach = 'General inquiry response';
    var keyComponents = 'Answer + Value proposition + Next steps';
    
    if (context === 'Pricing inquiry') {
      responseApproach = 'Pricing response with options';
      keyComponents = 'Pricing + Retainer option + Next steps';
    } else if (context === 'Audit request') {
      responseApproach = 'Audit offer with process explanation';
      keyComponents = 'Audit scope + Value proposition + Next steps';
    } else if (context === 'Onboarding inquiry') {
      responseApproach = 'Onboarding process explanation';
      keyComponents = 'Process overview + Timeline + Next steps';
    } else if (context === 'Project inquiry') {
      responseApproach = 'Project scope and proposal';
      keyComponents = 'Scope + Approach + Next steps';
    } else if (context === 'Strategy consultation') {
      responseApproach = 'Strategic guidance and recommendations';
      keyComponents = 'Insights + Recommendations + Next steps';
    } else if (context === 'Problem-solving inquiry') {
      responseApproach = 'Diagnostic problem-solving';
      keyComponents = 'Problem analysis + Solution approach + Next steps';
    } else if (context === 'Objection handling') {
      responseApproach = 'Objection addressing and reassurance';
      keyComponents = 'Address concerns + Provide reassurance + Next steps';
    } else if (context === 'Qualification inquiry') {
      responseApproach = 'Qualification and fit assessment';
      keyComponents = 'Fit assessment + Value alignment + Next steps';
    }

    // Infer immediate_context: Extract ONLY the most recent message (not truncated)
    var immediateContext = '';
    if (text) {
      // Find the last name/time pattern to get the most recent message
      var nameTimePattern = /([A-Z][a-zA-Z\s]{2,})\n(\d{1,2}:\d{2}\s*(?:AM|PM)?)\n/g;
      var lastMatch = null;
      var match;
      nameTimePattern.lastIndex = 0;
      while ((match = nameTimePattern.exec(text)) !== null) {
        lastMatch = match;
      }
      
      if (lastMatch) {
        // Get ALL text after the last name/time pattern (the complete most recent message)
        var afterLastPattern = text.substring(lastMatch.index + lastMatch[0].length).trim();
        // Clean up: remove "View details" and file attachments, but keep the full message
        afterLastPattern = afterLastPattern.replace(/View details.*$/i, '').replace(/\d+\s+files?.*$/i, '').trim();
        immediateContext = afterLastPattern; // Use the full message, not truncated
      } else {
        // No name/time pattern found, use the entire text as immediate context
        immediateContext = text.trim();
      }
    }
    
    // Infer context_summary: Summarize the first couple of messages in the conversation (up to 300 chars, shorter if possible)
    var contextSummary = '';
    if (text) {
      // Find all name/time patterns to identify message boundaries
      var nameTimePattern2 = /([A-Z][a-zA-Z\s]{2,})\n(\d{1,2}:\d{2}\s*(?:AM|PM)?)\n/g;
      var matches = [];
      var match;
      nameTimePattern2.lastIndex = 0;
      while ((match = nameTimePattern2.exec(text)) !== null) {
        matches.push(match);
      }
      
      if (matches.length > 0) {
        // Get the first 2-3 messages to summarize
        var messagesToSummarize = [];
        var maxMessages = Math.min(3, matches.length);
        
        for (var i = 0; i < maxMessages; i++) {
          var startIdx = matches[i].index + matches[i][0].length;
          var endIdx = (i < matches.length - 1) ? matches[i + 1].index : text.length;
          var messageText = text.substring(startIdx, endIdx).trim();
          // Clean up
          messageText = messageText.replace(/View details.*$/i, '').replace(/\d+\s+files?.*$/i, '').trim();
          if (messageText) {
            messagesToSummarize.push(messageText);
          }
        }
        
        // Combine first couple of messages and create a concise summary
        var combinedMessages = messagesToSummarize.join(' ').trim();
        if (combinedMessages.length <= 300) {
          // If already short enough, use as-is
          contextSummary = combinedMessages;
        } else {
          // If too long, summarize by taking key parts
          // Try to get first sentence or first meaningful chunk
          var firstSentence = combinedMessages.split(/[.!?]\s+/)[0];
          if (firstSentence && firstSentence.length <= 300) {
            contextSummary = firstSentence.trim();
          } else {
            // Take first 300 characters but try to end at a word boundary
            var truncated = combinedMessages.substring(0, 300);
            var lastSpace = truncated.lastIndexOf(' ');
            if (lastSpace > 200) {
              contextSummary = truncated.substring(0, lastSpace).trim();
            } else {
              contextSummary = truncated.trim();
            }
          }
        }
      } else {
        // No name/time patterns found, summarize first 300 characters intelligently
        var firstPart = text.substring(0, Math.min(500, text.length)).trim();
        var firstSentence = firstPart.split(/[.!?]\s+/)[0];
        if (firstSentence && firstSentence.length <= 300) {
          contextSummary = firstSentence.trim();
        } else {
          var truncated = firstPart.substring(0, 300);
          var lastSpace = truncated.lastIndexOf(' ');
          contextSummary = (lastSpace > 200) ? truncated.substring(0, lastSpace).trim() : truncated.trim();
        }
      }
      
      // Fallback if still empty
      if (!contextSummary || contextSummary.length < 10) {
        contextSummary = text.substring(0, Math.min(300, text.length)).trim() || 'Unknown';
      }
    }

    return {
      context: context,
      responseApproach: responseApproach,
      keyComponents: keyComponents,
      originalQuestion: originalQ,
      immediateContext: immediateContext || 'Unknown',
      contextSummary: contextSummary || 'Unknown'
    };
  } catch (e) {
    return { 
      context: 'General Inquiry', 
      responseApproach: 'Unknown', 
      keyComponents: 'Unknown', 
      originalQuestion: 'Unknown',
      immediateContext: 'Unknown',
      contextSummary: 'Unknown'
    };
  }
}

function inferOriginalQuestion_(conversation) {
  try {
    var text = String(conversation || '');
    if (!text) return '';
    
    // Pattern: Look for name followed by time (e.g., "Name\nTime\n" or "Name\nTime PM\n")
    // Examples: "Matt Sexton\n5:57 PM\n" or "Samuel Rainey\n5:57 PM\n"
    // Find the last occurrence of this pattern and extract everything after it
    // The pattern matches: Capital letter(s) + name (letters/spaces) + newline + time (HH:MM AM/PM) + newline
    var nameTimePattern = /([A-Z][a-zA-Z\s]{2,})\n(\d{1,2}:\d{2}\s*(?:AM|PM)?)\n/g;
    var matches = [];
    var match;
    var lastIndex = 0;
    
    // Find all name/time patterns (reset regex lastIndex to avoid issues)
    nameTimePattern.lastIndex = 0;
    while ((match = nameTimePattern.exec(text)) !== null) {
      matches.push({ index: match.index, endIndex: match.index + match[0].length });
      lastIndex = nameTimePattern.lastIndex;
    }
    
    // If we found name/time patterns, get text after the last one
    if (matches.length > 0) {
      var lastMatch = matches[matches.length - 1];
      var questionText = text.substring(lastMatch.endIndex).trim();
      
      // Clean up: remove common trailing elements like "View details", file attachments, etc.
      // Look for patterns like "View details" or file names and remove everything after
      var viewDetailsMatch = questionText.match(/^(.*?)(?:\nView details.*|$)/s);
      if (viewDetailsMatch && viewDetailsMatch[1]) {
        questionText = viewDetailsMatch[1].trim();
      }
      
      // Remove file attachment lines (lines that look like file paths or just file names)
      questionText = questionText.replace(/\n[^\n]*\.(pdf|docx?|xlsx?|png|jpg|jpeg|gif)(\s|$)/gi, '');
      
      if (questionText) return questionText.trim();
    }
    
    // Fallback: if no name/time pattern, look for the last question mark and get text from there
    var lastQMark = text.lastIndexOf('?');
    if (lastQMark !== -1) {
      // Find the start of that question (look backwards for newline or start)
      var start = text.lastIndexOf('\n', lastQMark);
      if (start === -1) start = 0;
      else start += 1; // Include the newline
      var question = text.substring(start).trim();
      if (question) return question;
    }
    
    // Final fallback: return the last 500 characters if nothing else works
    return text.length > 500 ? text.substring(text.length - 500).trim() : text.trim();
  } catch (e) {
    return '';
  }
}

// ================================
// Company Rules Management
// ================================

/**
 * Generate a company rule from input description
 * payload: { input: string }
 */
function generateCompanyRule(payload) {
  try {
    if (!payload || !payload.input) {
      return { success: false, error: 'Missing input text' };
    }

    // Get existing rules for context
    var allRules = getCompanyRules({ limit: 20, offset: 0 });
    var regularRules = allRules.filter(function(r){ return !(r.category || '').startsWith('PDF -'); });

    var systemPrompt = 'You are a company policy assistant. Your job is to create clear, structured company rules from descriptions or context provided by users.\n\n' +
      'Output format:\n' +
      '1. A clear, concise rule title (one line)\n' +
      '2. A detailed rule description that explains the policy, procedure, or guideline\n' +
      '3. A suggested category (e.g., Pricing, Communication, Compliance, Operations, etc.)\n\n' +
      'The rule should be professional, actionable, and suitable for a company knowledge base.';

    var userPrompt = '# Input\n' + String(payload.input) + '\n\n';
    userPrompt += '# Existing Company Rules (for reference)\n';
    regularRules.slice(0, 5).forEach(function(r){
      userPrompt += '- [' + (r.category || '') + '] ' + (r.rule_title || '') + ': ' + (r.rule_description || '').substring(0, 200) + '\n';
    });
    userPrompt += '\n# Task\n';
    userPrompt += 'Generate a company rule with:\n';
    userPrompt += '1. Rule Title: [one line title]\n';
    userPrompt += '2. Rule Description: [detailed description]\n';
    userPrompt += '3. Category: [suggested category]\n\n';
    userPrompt += 'Output the rule in a clear, structured format.';

    var apiKey = getOpenAIKey();
    if (!apiKey) return { success: false, error: 'OpenAI API key not configured' };

    // GPT-5/5.1 uses max_completion_tokens instead of max_tokens, and only supports temperature=1
    // Note: GPT-5.1-Codex is for Codex only, not Chat Completions API
    var isGPT5 = (typeof OPENAI_MODEL !== 'undefined' && (String(OPENAI_MODEL).toLowerCase().includes('gpt-5') || String(OPENAI_MODEL).toLowerCase().includes('gpt-5.1')));
    var maxTokensValue = (typeof OPENAI_MAX_TOKENS !== 'undefined' ? OPENAI_MAX_TOKENS : 700);
    
    var body = {
      model: OPENAI_MODEL,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ]
    };
    
    // Use max_completion_tokens for GPT-5, max_tokens for older models
    if (isGPT5) {
      body.max_completion_tokens = maxTokensValue;
      body.temperature = 1; // GPT-5 only supports default temperature (1)
    } else {
      body.max_tokens = maxTokensValue;
      body.temperature = 0.6; // Custom temperature for older models
    }

    var resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    
    // Check HTTP response code
    var responseCode = resp.getResponseCode();
    if (responseCode !== 200) {
      return { success: false, error: 'HTTP error ' + responseCode + ': ' + resp.getContentText() };
    }
    
    // Parse JSON with error handling
    var data;
    try {
      data = JSON.parse(resp.getContentText());
    } catch (e) {
      return { success: false, error: 'Invalid JSON response from OpenAI: ' + String(e) };
    }
    
    if (data.error) {
      return { success: false, error: 'OpenAI error: ' + (data.error.message || 'Unknown error') };
    }
    
    // Validate response structure
    if (!data.choices || !data.choices[0] || !data.choices[0].message) {
      // Log the actual response for debugging
      Logger.log('Unexpected response structure. Full response: ' + JSON.stringify(data).substring(0, 500));
      return { success: false, error: 'Unexpected response structure from OpenAI API. Check execution log for details.' };
    }
    
    var answer = data.choices[0].message.content;
    
    // Try alternative response structures if content is missing (GPT-5 might differ)
    if (!answer || typeof answer !== 'string') {
      Logger.log('Empty or invalid response. Trying alternative structures...');
      if (data.choices[0].content) {
        answer = data.choices[0].content;
      } else if (data.choices[0].text) {
        answer = data.choices[0].text;
      } else if (data.content) {
        answer = data.content;
      }
    }
    
    if (!answer || typeof answer !== 'string') {
      Logger.log('Still no valid answer found. Full response: ' + JSON.stringify(data).substring(0, 500));
      return { success: false, error: 'OpenAI returned empty or invalid response. Check execution log for response details.' };
    }

    // Extract category, title, and description from the response
    var category = 'General';
    var ruleTitle = 'Generated Rule';
    var ruleDescription = ''; // Start empty - must be extracted, don't default to full answer
    
    if (answer) {
      // Handle numbered format: "1. Rule Title: ... 2. Rule Description: ... 3. Category: ..."
      var numberedFormat = answer.match(/\d+\.\s*[Rr]ule\s+[Tt]itle:\s*([^\n]+)/i);
      if (numberedFormat) {
        ruleTitle = numberedFormat[1].trim();
      }
      
      // Try numbered description format - match everything after "Rule Description:" until next numbered item or end
      // Pattern: "2. Rule Description: [content] 3. Category:" or "2. Rule Description: [content] END"
      // CRITICAL: The capture group [1] should contain ONLY the description text, not the label
      var numberedDesc = answer.match(/\d+\.\s*[Rr]ule\s+[Dd]escription:\s*([\s\S]+?)(?=\n\s*\d+\.|$)/i);
      if (numberedDesc && numberedDesc[1] && numberedDesc[1].trim().length > 0) {
        ruleDescription = numberedDesc[1].trim();
        
        // CRITICAL: Remove any leading "Rule Description:" label that might have been captured
        ruleDescription = ruleDescription.replace(/^\d+\.\s*[Rr]ule\s+[Dd]escription:\s*/i, '');
        ruleDescription = ruleDescription.replace(/^[Rr]ule\s+[Dd]escription:\s*/i, '');
        
        // Remove any trailing metadata lines (Category, Title, etc.) that might have been captured
        ruleDescription = ruleDescription.replace(/\n\s*\d+\.\s*[Cc]ategory:.*$/i, '');
        ruleDescription = ruleDescription.replace(/\n\s*[Cc]ategory:.*$/i, '');
        ruleDescription = ruleDescription.replace(/\n\s*\d+\.\s*[Rr]ule\s+[Tt]itle:.*$/i, '');
        ruleDescription = ruleDescription.replace(/\n\s*[Rr]ule\s+[Tt]itle:.*$/i, '');
        ruleDescription = ruleDescription.replace(/\n\s*[Tt]itle:.*$/i, '');
        // Remove any remaining numbered items at the end
        ruleDescription = ruleDescription.replace(/\n\s*\d+\.\s*.*$/m, '');
        ruleDescription = ruleDescription.trim();
        
        // Log for debugging
        Logger.log('Extracted description (numbered format): ' + ruleDescription.substring(0, 100));
      } else {
        Logger.log('WARNING: Numbered description format did not match. Answer: ' + answer.substring(0, 300));
      }
      
      // Try numbered category format
      var numberedCat = answer.match(/\d+\.\s*[Cc]ategory:\s*([^\n]+)/i);
      if (numberedCat) {
        category = numberedCat[1].trim();
      }
      
      // If numbered format didn't work, try non-numbered format
      if (!numberedDesc) {
        // Try to extract category (non-numbered)
        var categoryMatch = answer.match(/[Cc]ategory:\s*([^\n]+)/i);
        if (categoryMatch && !numberedCat) category = categoryMatch[1].trim();
        
        // Try to extract title (non-numbered)
        var titleMatch = answer.match(/[Rr]ule\s+[Tt]itle:\s*([^\n]+)/i) || answer.match(/[Tt]itle:\s*([^\n]+)/i);
        if (titleMatch && !numberedFormat) ruleTitle = titleMatch[1].trim();
        
        // Extract description - get text after "Rule Description:" or "Description:" and before next numbered item or category
        var descMatch = answer.match(/[Rr]ule\s+[Dd]escription:\s*([\s\S]+?)(?=\n\s*\d+\.|$)/i) || 
                        answer.match(/[Dd]escription:\s*([\s\S]+?)(?=\n\s*[Cc]ategory:|$)/i) ||
                        answer.match(/[Dd]escription:\s*([\s\S]+?)(?=\n\s*\d+\.|$)/i);
        if (descMatch && descMatch[1] && !numberedDesc && descMatch[1].trim().length > 0) {
          ruleDescription = descMatch[1].trim();
          // Remove any trailing "Category:", "Title:", or numbered items
          ruleDescription = ruleDescription.replace(/\n\s*[Cc]ategory:.*$/i, '');
          ruleDescription = ruleDescription.replace(/\n\s*[Rr]ule\s+[Tt]itle:.*$/i, '');
          ruleDescription = ruleDescription.replace(/\n\s*[Tt]itle:.*$/i, '');
          ruleDescription = ruleDescription.replace(/\n\s*\d+\.\s*.*$/m, '');
          ruleDescription = ruleDescription.trim();
        } else if (!numberedDesc && !ruleDescription) {
          // Fallback: if we found title and category, try to extract description by removing them
          if ((titleMatch || numberedFormat) && (categoryMatch || numberedCat)) {
            // Try to find content between title and category
            var tempDesc = answer
              .replace(/\d+\.\s*[Rr]ule\s+[Tt]itle:.*?\n/gi, '')
              .replace(/[Rr]ule\s+[Tt]itle:.*?\n/gi, '')
              .replace(/[Tt]itle:.*?\n/gi, '')
              .replace(/\d+\.\s*[Rr]ule\s+[Dd]escription:\s*/gi, '')
              .replace(/[Rr]ule\s+[Dd]escription:\s*/gi, '')
              .replace(/[Dd]escription:\s*/gi, '')
              .replace(/\d+\.\s*[Cc]ategory:.*$/gi, '')
              .replace(/[Cc]ategory:.*$/gi, '')
              .trim();
            // Only use this if it looks like actual description content (not just whitespace or labels)
            if (tempDesc && tempDesc.length > 10 && !tempDesc.match(/^\d+\./)) {
              ruleDescription = tempDesc.replace(/\n\s*\d+\.\s*.*$/m, '').trim();
            }
          }
        }
      }
    }

    // Handle potential differences in usage structure for GPT-5
    var usage = data && data.usage;
    if (!usage && data && data.usage_info) {
      usage = data.usage_info; // Some models might use usage_info instead
    }
    var estimatedCost = computeCostUsd_(usage);

    // Final cleanup: ensure description doesn't contain title or category
    // CRITICAL: Never use the full answer as description - it must be extracted
    var originalDesc = ruleDescription;
    
    // ALWAYS remove any leading "Rule Description:" or "2. Rule Description:" labels (shouldn't be there, but safety check)
    ruleDescription = ruleDescription.replace(/^\d+\.\s*[Rr]ule\s+[Dd]escription:\s*/i, '');
    ruleDescription = ruleDescription.replace(/^[Rr]ule\s+[Dd]escription:\s*/i, '');
    
    // Remove trailing metadata patterns
    ruleDescription = ruleDescription
      .replace(/\n\s*\d+\.\s*[Cc]ategory:.*$/i, '')
      .replace(/\n\s*[Cc]ategory:.*$/i, '')
      .replace(/\n\s*\d+\.\s*[Rr]ule\s+[Tt]itle:.*$/i, '')
      .replace(/\n\s*[Rr]ule\s+[Tt]itle:.*$/i, '')
      .replace(/\n\s*\d+\.\s*.*$/m, '')
      .trim();
    
    // If description is empty or still contains metadata, try aggressive extraction from original answer
    if (!ruleDescription || 
        ruleDescription.length < 10 ||
        ruleDescription.indexOf('Rule Title:') !== -1 || 
        ruleDescription.indexOf('Category:') !== -1 ||
        ruleDescription.match(/^\d+\.\s*[Rr]ule\s+[Tt]itle:/i) ||
        ruleDescription.match(/^\d+\.\s*[Cc]ategory:/i) ||
        ruleDescription.indexOf('2. Rule Description:') !== -1 ||
        ruleDescription.indexOf('1. Rule Title:') !== -1) {
      Logger.log('WARNING: Description extraction failed or contains metadata. Attempting aggressive extraction.');
      Logger.log('Current description: ' + (ruleDescription || '(empty)').substring(0, 200));
      Logger.log('Original answer: ' + answer.substring(0, 500));
      
      // Try multiple extraction patterns - be very specific about what we want
      // Pattern 1: "2. Rule Description: [TEXT] 3. Category:"
      var aggressiveMatch = answer.match(/\d+\.\s*[Rr]ule\s+[Dd]escription:\s*([\s\S]+?)(?=\n\s*\d+\.\s*[Cc]ategory:|$)/i);
      
      // Pattern 2: "Rule Description: [TEXT] Category:" (non-numbered)
      if (!aggressiveMatch || !aggressiveMatch[1]) {
        aggressiveMatch = answer.match(/[Rr]ule\s+[Dd]escription:\s*([\s\S]+?)(?=\n\s*\d+\.\s*[Cc]ategory:|\n\s*[Cc]ategory:|$)/i);
      }
      
      // Pattern 3: "2. Rule Description: [TEXT] 3." (any numbered item)
      if (!aggressiveMatch || !aggressiveMatch[1]) {
        aggressiveMatch = answer.match(/\d+\.\s*[Rr]ule\s+[Dd]escription:\s*([\s\S]+?)(?=\n\s*\d+\.|$)/i);
      }
      
      if (aggressiveMatch && aggressiveMatch[1] && aggressiveMatch[1].trim().length > 10) {
        ruleDescription = aggressiveMatch[1].trim();
        // CRITICAL: Remove any labels that might have been captured (shouldn't happen with proper regex, but safety)
        ruleDescription = ruleDescription.replace(/^\d+\.\s*[Rr]ule\s+[Dd]escription:\s*/i, '');
        ruleDescription = ruleDescription.replace(/^[Rr]ule\s+[Dd]escription:\s*/i, '');
        ruleDescription = ruleDescription.replace(/\n\s*\d+\.\s*[Cc]ategory:.*$/i, '');
        ruleDescription = ruleDescription.replace(/\n\s*[Cc]ategory:.*$/i, '');
        ruleDescription = ruleDescription.replace(/\n\s*\d+\.\s*.*$/m, '');
        ruleDescription = ruleDescription.trim();
        Logger.log('Extracted via aggressive match: ' + ruleDescription.substring(0, 200));
      } else {
        Logger.log('ERROR: All extraction attempts failed. Description will be empty.');
        ruleDescription = ''; // Don't use the full answer - return empty instead
      }
    }
    
    // Final validation - ensure we have actual content and it's clean
    if (!ruleDescription || ruleDescription.length < 10) {
      Logger.log('ERROR: Description is empty or too short after all extraction attempts');
      Logger.log('Full answer was: ' + answer);
      // Return empty rather than the full answer to prevent saving metadata
      ruleDescription = '';
    }
    
    // Final check: if description still contains numbered labels or metadata, it's invalid - reject it
    if (ruleDescription.match(/^\d+\.\s*[Rr]ule\s+[Tt]itle:/i) || 
        ruleDescription.match(/^\d+\.\s*[Cc]ategory:/i) ||
        ruleDescription.match(/^\d+\.\s*[Rr]ule\s+[Dd]escription:/i) ||
        ruleDescription.indexOf('2. Rule Description:') !== -1 ||
        ruleDescription.indexOf('1. Rule Title:') !== -1 ||
        ruleDescription.indexOf('3. Category:') !== -1) {
      Logger.log('ERROR: Description still contains numbered labels or metadata. Rejecting.');
      Logger.log('Rejected description: ' + ruleDescription.substring(0, 200));
      ruleDescription = '';
    }
    
    return {
      success: true,
      answer: answer || '',
      category: category,
      ruleTitle: ruleTitle,
      ruleDescription: ruleDescription, // Clean description without metadata
      ruleDetails: 'Category: ' + category + '\nTitle: ' + ruleTitle + '\n\nDescription:\n' + ruleDescription,
      meta: {
        model: OPENAI_MODEL,
        usage: usage,
        estimated_cost_usd: estimatedCost
      }
    };
  } catch (e) {
    return { success: false, error: String(e) };
  }
}

/**
 * Save multiple PDF rules to the Company_Rules sheet
 * payload: { pdfName, rules: [{ category, title, description }, ...] }
 */
function savePDFRules(payload) {
  try {
    if (!payload || !payload.pdfName || !payload.rules || !Array.isArray(payload.rules) || payload.rules.length === 0) {
      return { success: false, error: 'Missing required fields: pdfName and rules array' };
    }

    var ss = SpreadsheetApp.openById(LEADS_SPREADSHEET_ID);
    var sheet = ss.getSheetByName(COMPANY_RULES_SHEET);
    
    if (!sheet) {
      sheet = ss.insertSheet(COMPANY_RULES_SHEET);
      sheet.appendRow(['Date_Updated', 'Category', 'Rule_Title', 'Rule_Description']);
      var headerRange = sheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
    }

    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var rowsAdded = 0;

    payload.rules.forEach(function(rule) {
      sheet.appendRow([
        dateStr,
        'PDF - ' + payload.pdfName,  // Category: PDF - [Name]
        rule.title || rule.ruleTitle || '',
        rule.description || rule.ruleDescription || ''
      ]);
      rowsAdded++;
    });

    return {
      success: true,
      message: 'Successfully saved ' + rowsAdded + ' rule(s) to Company Rules',
      rulesSaved: rowsAdded
    };
  } catch (e) {
    return { success: false, error: String(e) };
  }
}

/**
 * Save a company rule to the Company_Rules sheet
 * payload: { category, ruleTitle, ruleDescription }
 */
function saveCompanyRule(payload) {
  try {
    if (!payload || !payload.category || !payload.ruleTitle || !payload.ruleDescription) {
      return { success: false, error: 'Missing required fields: category, ruleTitle, and ruleDescription' };
    }

    // Clean the description - should already be clean from generateCompanyRule, but do light cleanup
    var cleanDescription = String(payload.ruleDescription || '').trim();
    
    // Only do minimal cleanup here - the extraction in generateCompanyRule should have handled it
    // Remove any leading "Rule Description:" label
    cleanDescription = cleanDescription.replace(/^[Rr]ule\s+[Dd]escription:\s*/i, '');
    
    // Remove trailing metadata (shouldn't be there, but just in case)
    cleanDescription = cleanDescription
      .replace(/\n\s*\d+\.\s*[Cc]ategory:.*$/i, '')
      .replace(/\n\s*[Cc]ategory:.*$/i, '')
      .trim();
    
    // Log if we detect issues
    if (cleanDescription.indexOf('Rule Title:') !== -1 || 
        cleanDescription.indexOf('Category:') !== -1) {
      Logger.log('WARNING: Description may contain metadata in saveCompanyRule. Description: ' + cleanDescription.substring(0, 200));
    }
    
    // Ensure we have content
    if (!cleanDescription || cleanDescription.length < 5) {
      Logger.log('ERROR: Description is empty in saveCompanyRule. Payload was: ' + JSON.stringify(payload));
      return { success: false, error: 'Description is empty or invalid' };
    }

    var ss = SpreadsheetApp.openById(LEADS_SPREADSHEET_ID);
    var sheet = ss.getSheetByName(COMPANY_RULES_SHEET);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(COMPANY_RULES_SHEET);
      // Add headers
      sheet.appendRow(['Date_Updated', 'Category', 'Rule_Title', 'Rule_Description']);
      // Format header row
      var headerRange = sheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
    }

    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    sheet.appendRow([
      dateStr,
      String(payload.category || '').trim(),
      String(payload.ruleTitle || '').trim(),
      cleanDescription // Use cleaned description
    ]);

    return {
      success: true,
      message: 'Company rule saved successfully',
      category: payload.category,
      ruleTitle: payload.ruleTitle
    };

  } catch (e) {
    return { success: false, error: String(e) };
  }
}

// ================================
// Lead Magnets Management
// ================================

/**
 * Get all lead magnets from the Lead_Magnets sheet
 * Expected columns: Title (required), URL (required), Description (optional)
 * Simple structure for URL-based lead magnets
 */
function getLeadMagnets() {
  try {
    var ss = SpreadsheetApp.openById(LEADS_SPREADSHEET_ID);
    var sheet = ss.getSheetByName(LEAD_MAGNETS_SHEET);
    if (!sheet) return []; // Sheet doesn't exist, return empty array
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return []; // No data rows (only header)
    
    var headers = data[0];
    var titleIdx = headers.indexOf('Title');
    var urlIdx = headers.indexOf('URL');
    var descIdx = headers.indexOf('Description');
    
    // Title and URL are required
    if (titleIdx === -1 || urlIdx === -1) return []; // Required columns missing
    
    var magnets = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[titleIdx] || !row[urlIdx]) continue; // Skip empty rows
      var url = String(row[urlIdx] || '').trim();
      var title = String(row[titleIdx] || '').trim();
      var description = descIdx !== -1 ? String(row[descIdx] || '').trim() : '';
      
      magnets.push({
        title: title,
        url: url,
        description: description
      });
    }
    return magnets;
  } catch (e) {
    console.error('Error getting lead magnets:', e);
    return [];
  }
}

/**
 * Get list of lead magnets that have been used in previous follow-ups
 */
function getUsedLeadMagnets_(conversationId, history) {
  try {
    var used = [];
    if (!history || history.length === 0) return used;
    
    // Check history for mentions of lead magnets
    // This is a simple heuristic - could be improved
    var allMagnets = [];
    try {
      allMagnets = getLeadMagnets();
    } catch (e) {
      return []; // If we can't get magnets, return empty
    }
    
    if (allMagnets.length === 0) return used;
    
    // Filter history to only follow-ups if we have conversation context
    var followupHistory = history;
    if (history.length > 0 && history[0].conversation_context) {
      followupHistory = history.filter(function(h){
        var ctx = (h.conversation_context || '').toLowerCase();
        return ctx.indexOf('follow-up') !== -1 || ctx.indexOf('followup') !== -1;
      });
    }
    
    followupHistory.forEach(function(h){
      var responseText = (h.full_response || '').toLowerCase();
      allMagnets.forEach(function(m){
        var title = (m.title || '').toLowerCase();
        if (title && title.length > 3 && responseText.indexOf(title) !== -1) {
          if (used.indexOf(m.title) === -1) {
            used.push(m.title);
          }
        }
      });
    });
    return used;
  } catch (e) {
    console.error('Error getting used lead magnets:', e);
    return [];
  }
}

/**
 * Filter and score available lead magnets based on context and usage
 */
function filterAvailableLeadMagnets_(allMagnets, usedMagnets, contextKeywords) {
  try {
    if (!allMagnets || allMagnets.length === 0) return [];
    
    // Filter out used magnets
    var available = allMagnets.filter(function(m){
      return usedMagnets.indexOf(m.title) === -1;
    });
    
    if (available.length === 0) return [];
    
    // Score magnets based on keyword relevance
    var scored = available.map(function(m){
      var score = 0;
      // Search in title, URL, and description
      var searchText = ((m.title || '') + ' ' + (m.url || '') + ' ' + (m.description || '')).toLowerCase();
      var kwList = (contextKeywords || '').toLowerCase().split(/\s+/).filter(function(k){ return k.length > 2; });
      
      kwList.forEach(function(kw){
        if (searchText.indexOf(kw) !== -1) {
          score += 1;
          if ((m.title || '').toLowerCase().indexOf(kw) !== -1) score += 2; // Title matches are more important
          if ((m.description || '').toLowerCase().indexOf(kw) !== -1) score += 1; // Description matches
        }
      });
      
      return Object.assign({}, m, { relevance_score: score });
    });
    
    // Sort by relevance and return top 3
    scored.sort(function(a, b){ return b.relevance_score - a.relevance_score; });
    return scored.slice(0, 3);
  } catch (e) {
    return allMagnets.slice(0, 3); // Fallback: return first 3
  }
}

/**
 * Determine follow-up type based on conversation and transcript
 * Returns: 'pre-call', 'post-call', or 'lost'
 */
function determineFollowupType_(conversation, callTranscript) {
  try {
    var text = (conversation || '').toLowerCase();
    var transcript = (callTranscript || '').toLowerCase();
    
    // If transcript is provided, it's definitely post-call
    if (transcript && transcript.trim().length > 50) {
      return 'post-call';
    }
    
    // Check for call indicators in conversation
    var callIndicators = [
      'we spoke', 'we talked', 'on the call', 'during our call', 'after our call',
      'call went well', 'great call', 'thanks for the call', 'appreciate the call',
      'scheduled a call', 'set up a call', 'had a call', 'call yesterday', 'call today'
    ];
    
    var hasCallIndicators = callIndicators.some(function(indicator){
      return text.indexOf(indicator) !== -1;
    });
    
    if (hasCallIndicators) {
      return 'post-call';
    }
    
    // Check for "lost" indicators - long time since last contact, project abandoned, etc.
    var lostIndicators = [
      'a while back', 'a while ago', 'long time', 'been a while',
      'given up', 'moved on', 'decided to go with', 'went with another',
      'no longer interested', 'not moving forward'
    ];
    
    var hasLostIndicators = lostIndicators.some(function(indicator){
      return text.indexOf(indicator) !== -1;
    });
    
    if (hasLostIndicators) {
      return 'lost';
    }
    
    // Default to pre-call if no indicators found
    return 'pre-call';
  } catch (e) {
    return 'pre-call'; // Default fallback
  }
}

// ================================
// PDF Knowledge Base Management
// ================================

/**
 * Process pasted PDF text: chunk it and save to Company_Rules sheet
 * payload: { pdfName, sourceFile, text, chunkSize (optional, default 2000) }
 */
function processPDFText(payload) {
  try {
    if (!payload || !payload.text || !payload.pdfName) {
      return { success: false, error: 'Missing required fields: pdfName and text' };
    }

    var text = String(payload.text || '');
    var pdfName = String(payload.pdfName || 'Unknown');
    var chunkSize = payload.chunkSize || 2000; // characters per chunk

    // Chunk the text intelligently
    var chunks = chunkTextIntelligently_(text, chunkSize);
    
    if (chunks.length === 0) {
      return { success: false, error: 'No content to chunk' };
    }

    // Save chunks to Company_Rules sheet
    var ss = SpreadsheetApp.openById(LEADS_SPREADSHEET_ID);
    var sheet = ss.getSheetByName(COMPANY_RULES_SHEET);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(COMPANY_RULES_SHEET);
      // Add headers
      sheet.appendRow(['Date_Updated', 'Category', 'Rule_Title', 'Rule_Description']);
      // Format header row
      var headerRange = sheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
    }

    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var rowsAdded = 0;

    // Save each chunk as a row in Company_Rules
    // Category format: "PDF - [PDF Name]" to distinguish from regular rules
    chunks.forEach(function(chunk, idx) {
      var section = chunk.section || ('Section ' + (idx + 1));
      var content = chunk.text;

      sheet.appendRow([
        dateStr,
        'PDF - ' + pdfName,  // Category: PDF - [Name]
        section,              // Rule_Title: section name
        content               // Rule_Description: chunk content
      ]);
      rowsAdded++;
    });

    return {
      success: true,
      message: 'PDF processed successfully',
      chunksCreated: rowsAdded,
      pdfName: pdfName
    };

  } catch (e) {
    return { success: false, error: String(e) };
  }
}

/**
 * Process PDF text using GPT to extract multiple valuable company rules
 * payload: { pdfName, text }
 */
function processPDFToRules(payload) {
  try {
    if (!payload || !payload.text) {
      return { success: false, error: 'Missing required field: text' };
    }

    var text = String(payload.text || '');
    
    if (text.length === 0) {
      return { success: false, error: 'No text provided' };
    }

    // PDF name is optional - try to extract from text or use default
    var pdfName = payload.pdfName || '';
    if (!pdfName || pdfName.trim() === '') {
      // Try to extract PDF name from text (look for title patterns)
      var titleMatch = text.match(/(?:^|\n)([A-Z][^.\n]{10,80})(?:\n|$)/);
      if (titleMatch && titleMatch[1]) {
        pdfName = titleMatch[1].trim().substring(0, 50);
      } else {
        pdfName = 'PDF Document';
      }
    }
    pdfName = String(pdfName).trim();

    // If text is very long, chunk it first to avoid token limits
    var maxTextLength = 15000; // characters
    var textToAnalyze = text;
    if (text.length > maxTextLength) {
      // Take first part for analysis, but note that we may need multiple passes
      textToAnalyze = text.substring(0, maxTextLength) + '\n\n[Content continues...]';
    }

    var systemPrompt = 'You are a company knowledge base assistant. Your job is to extract valuable, actionable company rules, policies, and guidelines from business documents.\n\n' +
      'Extract ONLY business rules, policies, procedures, and guidelines that would guide how the company operates or how employees should behave. Focus on:\n' +
      '- Business policies and procedures\n' +
      '- Customer service guidelines and standards\n' +
      '- Communication rules and tone guidelines\n' +
      '- Pricing policies and fee structures\n' +
      '- Operational processes and workflows\n' +
      '- Best practices for client interactions\n' +
      '- Strategic principles and approaches\n' +
      '- Quality standards and expectations\n\n' +
      'IGNORE and DO NOT extract:\n' +
      '- Technical implementation details (code, APIs, system architecture)\n' +
      '- Error handling or debugging procedures\n' +
      '- Software configuration or setup instructions\n' +
      '- Filler text, introductions, or fluff\n' +
      '- Redundant information\n' +
      '- Examples that don\'t contain actual rules\n' +
      '- Marketing or promotional content\n' +
      '- Any technical documentation about how systems work\n\n' +
      'Output format: A numbered list where each item is a rule in this format:\n' +
      '1. [Category] | [Rule Title] | [Rule Description]\n' +
      '2. [Category] | [Rule Title] | [Rule Description]\n' +
      '...\n\n' +
      'Categories should be business-focused (e.g., Pricing, Communication, Client Service, Operations, Strategy).\n' +
      'Keep each rule brief but complete. Aim for 5-15 rules depending on the content quality.';

    var userPrompt = '# PDF Document: ' + pdfName + '\n\n';
    userPrompt += '# Content\n' + textToAnalyze + '\n\n';
    userPrompt += '# Task\n';
    userPrompt += 'Extract the most valuable company rules and guidelines from this PDF. Output them as a numbered list in the format:\n';
    userPrompt += '1. [Category] | [Rule Title] | [Rule Description]\n\n';
    userPrompt += 'Be selective - only extract truly valuable rules. Keep descriptions concise but complete.';

    var apiKey = getOpenAIKey();
    if (!apiKey) return { success: false, error: 'OpenAI API key not configured' };

    // GPT-5/5.1 uses max_completion_tokens instead of max_tokens, and only supports temperature=1
    var isGPT5 = (typeof OPENAI_MODEL !== 'undefined' && (String(OPENAI_MODEL).toLowerCase().includes('gpt-5') || String(OPENAI_MODEL).toLowerCase().includes('gpt-5.1')));
    var maxTokensValue = 2000; // Need more tokens for multiple rules
    
    var body = {
      model: OPENAI_MODEL,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ]
    };
    
    if (isGPT5) {
      body.max_completion_tokens = maxTokensValue;
      body.temperature = 1;
    } else {
      body.max_tokens = maxTokensValue;
      body.temperature = 0.7;
    }

    var resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    
    var responseCode = resp.getResponseCode();
    if (responseCode !== 200) {
      return { success: false, error: 'HTTP error ' + responseCode + ': ' + resp.getContentText() };
    }
    
    var data;
    try {
      data = JSON.parse(resp.getContentText());
    } catch (e) {
      return { success: false, error: 'Invalid JSON response from OpenAI: ' + String(e) };
    }
    
    if (data.error) {
      return { success: false, error: 'OpenAI error: ' + (data.error.message || 'Unknown error') };
    }
    
    if (!data.choices || !data.choices[0] || !data.choices[0].message) {
      return { success: false, error: 'Unexpected response structure from OpenAI API' };
    }
    
    var answer = data.choices[0].message.content;
    if (!answer || typeof answer !== 'string') {
      return { success: false, error: 'OpenAI returned empty response' };
    }

    // Parse the rules from the response
    var rules = [];
    var lines = answer.split('\n');
    
    lines.forEach(function(line) {
      line = line.trim();
      if (!line || line.length === 0) return;
      
      // Match pattern: "1. Category | Title | Description" or "1. Category|Title|Description"
      var match = line.match(/^\d+\.\s*(.+?)\s*\|\s*(.+?)\s*\|\s*(.+)$/);
      if (match) {
        rules.push({
          category: match[1].trim(),
          title: match[2].trim(),
          description: match[3].trim()
        });
        return;
      }
      
      // Try pattern with brackets: "1. [Category] | [Title] | Description"
      match = line.match(/^\d+\.\s*\[([^\]]+)\]\s*\|\s*\[([^\]]+)\]\s*\|\s*(.+)$/);
      if (match) {
        rules.push({
          category: match[1].trim(),
          title: match[2].trim(),
          description: match[3].trim()
        });
        return;
      }
      
      // Try pattern: "1. [Category] | Title | Description" (mixed)
      match = line.match(/^\d+\.\s*\[([^\]]+)\]\s*\|\s*(.+?)\s*\|\s*(.+)$/);
      if (match) {
        rules.push({
          category: match[1].trim(),
          title: match[2].trim(),
          description: match[3].trim()
        });
        return;
      }
    });

    // If still no rules, try multi-line parsing
    if (rules.length === 0) {
      // Try to find numbered items that span multiple lines
      var numberedBlocks = answer.split(/(?=^\d+\.)/m);
      numberedBlocks.forEach(function(block) {
        block = block.trim();
        if (!block) return;
        
        // Try to extract category, title, description from the block
        var lines = block.split('\n').map(function(l) { return l.trim(); }).filter(function(l) { return l; });
        if (lines.length >= 2) {
          var firstLine = lines[0].replace(/^\d+\.\s*/, '');
          var parts = firstLine.split(/\s*\|\s*/);
          if (parts.length >= 3) {
            rules.push({
              category: parts[0].replace(/^\[|\]$/g, '').trim(),
              title: parts[1].replace(/^\[|\]$/g, '').trim(),
              description: parts.slice(2).join(' | ').trim() + (lines.length > 1 ? ' ' + lines.slice(1).join(' ') : '')
            });
          }
        }
      });
    }

    if (rules.length === 0) {
      return { success: false, error: 'Could not extract rules from response. GPT may need better formatting instructions.' };
    }

    // Return rules without saving - user will click button to save
    return {
      success: true,
      message: 'PDF processed successfully. Review the rules below and click "Add to Spreadsheet" to save them.',
      rulesCreated: rules.length,
      pdfName: pdfName,
      rules: rules
    };

  } catch (e) {
    return { success: false, error: String(e) };
  }
}

/**
 * Intelligently chunk text by looking for natural breaks (paragraphs, headings, etc.)
 */
function chunkTextIntelligently_(text, maxChunkSize) {
  if (!text || text.length === 0) return [];
  
  var chunks = [];
  var lines = text.split(/\n+/);
  var currentChunk = '';
  var currentSection = '';
  var chunkIndex = 0;

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;

    // Detect section headers (lines that are short, all caps, or end with colon)
    var isHeader = (line.length < 80 && (
      line === line.toUpperCase() ||
      /^[A-Z][^.!?]*:$/.test(line) ||
      /^#{1,3}\s/.test(line) // Markdown headers
    ));

    if (isHeader && currentChunk.length > 100) {
      // Save current chunk before starting new section
      if (currentChunk.length > 0) {
        chunks.push({
          text: currentChunk.trim(),
          section: currentSection || ('Section ' + (chunkIndex + 1)),
          pageRange: ''
        });
        chunkIndex++;
        currentChunk = '';
      }
      currentSection = line.replace(/^#{1,3}\s*/, '').replace(/:$/, '');
    }

    // Add line to current chunk
    var lineWithNewline = line + '\n';
    
    if (currentChunk.length + lineWithNewline.length > maxChunkSize && currentChunk.length > 0) {
      // Current chunk is getting too large, save it
      chunks.push({
        text: currentChunk.trim(),
        section: currentSection || ('Section ' + (chunkIndex + 1)),
        pageRange: ''
      });
      chunkIndex++;
      currentChunk = lineWithNewline;
    } else {
      currentChunk += lineWithNewline;
    }
  }

  // Save final chunk
  if (currentChunk.trim().length > 0) {
    chunks.push({
      text: currentChunk.trim(),
      section: currentSection || ('Section ' + (chunkIndex + 1)),
      pageRange: ''
    });
  }

  return chunks;
}

/**
 * Select relevant PDF chunks from Company_Rules based on keywords
 */
function selectRelevantPDFChunks_(pdfChunks, keywords, limit) {
  if (!pdfChunks || pdfChunks.length === 0) return [];
  if (!keywords) return pdfChunks.slice(0, limit || 6);

  try {
    var kwList = String(keywords).toLowerCase().split(/\s+/).filter(function(k) { return k.length > 2; });
    
    var scored = pdfChunks.map(function(chunk) {
      var searchText = [
        chunk.rule_description || '',
        chunk.rule_title || '',
        chunk.category || ''
      ].join(' ').toLowerCase();
      
      var score = 0;
      kwList.forEach(function(kw) {
        if (searchText.indexOf(kw) !== -1) {
          score += 1;
          if ((chunk.rule_description || '').toLowerCase().indexOf(kw) !== -1) score += 2;
          if ((chunk.rule_title || '').toLowerCase().indexOf(kw) !== -1) score += 1;
        }
      });
      
      return Object.assign({}, chunk, { relevance_score: score });
    }).filter(function(c) { return c.relevance_score > 0; })
      .sort(function(a, b) { return b.relevance_score - a.relevance_score; });

    var maxResults = limit || 6;
    return scored.slice(0, maxResults);

  } catch (e) {
    console.error('Error selecting PDF chunks:', e);
    return pdfChunks.slice(0, limit || 6);
  }
}

// ================================
// Logging generated/adjusted responses to the sheet
// ================================

function saveGeneratedRow(rowString) {
  if (!rowString) return { success: false, error: 'Missing row' };
  
  // CRITICAL: The sheet has exactly 11 columns:
  // Date, Conversation_ID, Turn_Index, Conversation_Context, Industry, Response_Approach,
  // Key_Components, Original_Question, Immediate_Context, Context_Summary, Full_Response
  // We must ensure exactly 11 values are written - no more, no less.
  
  // Split by pipe, but handle cases where pipes appear in the content
  // We know the structure: 11 fields total, with Full_Response being the last field
  var parts = String(rowString).split('|');
  
  // If we have more than 11 parts, it means pipes were in one of the fields
  // The most likely culprit is Full_Response (field 11), but could be any field
  // Strategy: Take first 10 fields, then join everything else as the 11th field
  if (parts.length > 11) {
    var firstTen = parts.slice(0, 10);
    var fullResponse = parts.slice(10).join('|'); // Rejoin with pipes since they were part of the content
    parts = firstTen.concat([fullResponse]);
  }
  
  // Ensure we have exactly 11 columns (pad with empty strings if needed)
  while (parts.length < 11) {
    parts.push('');
  }
  
  // CRITICAL: Trim to exactly 11 columns - never write more than 11 columns
  parts = parts.slice(0, 11);
  
  // Final validation: must have exactly 11 elements
  if (parts.length !== 11) {
    return { success: false, error: 'Row must have exactly 11 columns, got ' + parts.length };
  }
  
  var ss = SpreadsheetApp.openById(LEADS_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(DATA_LOG_SHEET);
  if (!sheet) return { success: false, error: 'Data Log sheet not found' };
  
  // Write exactly 11 columns to the sheet
  sheet.appendRow(parts);
  return { success: true };
}

function saveAdjustedResponse(payload) {
  // payload: { baseRow: string, adjustedResponse: string }
  if (!payload || !payload.baseRow || !payload.adjustedResponse) return { success: false, error: 'Missing data' };
  
  // CRITICAL: The sheet has exactly 11 columns - we must maintain this structure
  // Replace only the Full_Response field (column 11, index 10) while keeping all other fields intact
  
  // Split by pipe, handling pipes in content
  var parts = String(payload.baseRow).split('|');
  
  // If we have more than 11 parts, reconstruct properly
  if (parts.length > 11) {
    var firstTen = parts.slice(0, 10);
    var fullResponse = parts.slice(10).join('|');
    parts = firstTen.concat([fullResponse]);
  }
  
  // Ensure we have exactly 11 columns
  while (parts.length < 11) {
    parts.push('');
  }
  
  // CRITICAL: Trim to exactly 11 columns - never write more than 11 columns
  parts = parts.slice(0, 11);
  
  // Replace full response (last field, index 10) - keep pipes as-is if they exist
  parts[10] = String(payload.adjustedResponse).replace(/\n/g, ' ');
  
  // Final validation: must have exactly 11 elements
  if (parts.length !== 11) {
    return { success: false, error: 'Row must have exactly 11 columns, got ' + parts.length };
  }
  
  var ss = SpreadsheetApp.openById(LEADS_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(DATA_LOG_SHEET);
  if (!sheet) return { success: false, error: 'Data Log sheet not found' };
  
  // Write exactly 11 columns to the sheet
  sheet.appendRow(parts);
  return { success: true };
}

// ============================================
// TEST FUNCTION - Run this to trigger authorization
// ============================================
function testAuthorization() {
  // This function will trigger the authorization prompt
  // Run it from: Run → Run function → testAuthorization
  try {
    var apiKey = getOpenAIKey();
    if (!apiKey || apiKey === 'your-openai-api-key-here') {
      Logger.log('ERROR: Please set your OpenAI API key in the OpenAI tab (cell B1) or in Script Properties');
      return 'ERROR: Please set your OpenAI API key first';
    }
    
    // Make a minimal API call to trigger authorization
    var resp = UrlFetchApp.fetch('https://api.openai.com/v1/models', {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      muteHttpExceptions: true
    });
    
    var code = resp.getResponseCode();
    if (code === 200) {
      Logger.log('SUCCESS: Authorization working!');
      return 'SUCCESS: Authorization is working! You can now use the web app.';
    } else {
      Logger.log('API call failed with code: ' + code);
      Logger.log('Response: ' + resp.getContentText());
      return 'API call failed. Check logs for details.';
    }
  } catch (e) {
    Logger.log('Error: ' + e.toString());
    return 'Error: ' + e.toString();
  }
}

// ============================================
// SIMPLE AUTHORIZATION TRIGGER
// ============================================
function triggerAuthorization() {
  // This is the SIMPLEST possible external request to trigger authorization
  // Run this function: Run → Run function → triggerAuthorization
  try {
    // Make a simple GET request to a public API (no auth needed)
    var response = UrlFetchApp.fetch('https://httpbin.org/get', {
      muteHttpExceptions: true
    });
    Logger.log('Authorization test successful! Status: ' + response.getResponseCode());
    return 'SUCCESS: External request permission granted!';
  } catch (e) {
    Logger.log('Error: ' + e.toString());
    // If this fails, it means we need authorization
    return 'ERROR: ' + e.toString() + '\n\nPlease authorize the script by:\n1. Deploying as web app\n2. Accessing the web app URL in a browser\n3. Or manually reviewing permissions in Project Settings';
  }
}


