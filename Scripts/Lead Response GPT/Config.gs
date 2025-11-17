// ============================================
// LEAD RESPONSE GPT - CONFIGURATION
// ============================================

// OpenAI
const OPENAI_API_KEY = 'your-openai-api-key-here';
const OPENAI_MODEL = 'gpt-4o'; // Options: 'gpt-4o' (stable, uses Chat Completions API). Note: GPT-5.1 requires Responses API (/v1/responses) - would need code changes to support
const OPENAI_MAX_TOKENS = 2000; // Doubled from 1000 to allow more detailed and comprehensive responses
const DAILY_COST_LIMIT_USD = 5; // warn when exceeding this per calendar day
// GPT-4o pricing (as of 2025): Input $2.50/million, Output $10.00/million
const COST_PER_1K_INPUT_TOKENS_USD = 0.0025;  // $2.50 per million = $0.0025 per 1K
const COST_PER_1K_OUTPUT_TOKENS_USD = 0.01;    // $10.00 per million = $0.01 per 1K
// Legacy flat rate (deprecated - use input/output rates above)
const COST_PER_1K_TOKENS_USD = 0.01; // Fallback only

// Sheets
const LEADS_SPREADSHEET_ID = '1ZhrUxuevNYM2pD_TJ6IH_WRePKKo_QCjCPW-64hKiB4';
const DATA_LOG_SHEET = 'ChatGPT_Data_Log';
const COMPANY_RULES_SHEET = 'Company_Rules';
const LEAD_MAGNETS_SHEET = 'Lead_Magnets';
const OPENAI_TAB_NAME = 'OpenAI'; // API key in B1

// Optional defaults for UI
const DEFAULT_KEYWORDS_HINT = 'pricing audit onboarding quote schedule call';

// Prompt context sizing
const RULES_IN_PROMPT = 24;          // number of rules to inject (doubled from 12)
const EXAMPLES_IN_PROMPT = 20;       // number of examples to inject (doubled from 10)
const MAX_EXAMPLES_FETCH = 240;      // upper bound on examples fetched before filtering (doubled from 120)
const MAX_RULES_SELECTION = 50;      // maximum rules that can be selected (doubled from 25)


