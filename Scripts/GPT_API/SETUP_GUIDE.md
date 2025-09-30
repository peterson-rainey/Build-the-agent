# Lead Response System - Complete Setup Guide

## What This System Does
Helps VAs process lead conversations and generate appropriate responses using ChatGPT and Google Sheets.

## Workflow
1. Lead messages → VA copies conversation to ChatGPT
2. ChatGPT finds similar examples → Generates response + spreadsheet row
3. VA sends to boss → Boss approves/revises → VA sends final response
4. VA updates spreadsheet with final response

## Step 1: Google Sheets Setup ✅

Your Google Sheet is already set up correctly: [https://docs.google.com/spreadsheets/d/1ZhrUxuevNYM2pD_TJ6IH_WRePKKo_QCjCPW-64hKiB4/edit](https://docs.google.com/spreadsheets/d/1ZhrUxuevNYM2pD_TJ6IH_WRePKKo_QCjCPW-64hKiB4/edit)

### Tab 1: "Data Log" ✅
Your existing structure is perfect:
- **A**: Date (2025-07-10)
- **B**: Conversation_ID (1)
- **C**: Turn_Index (1)
- **D**: Conversation_Context (Direct pricing inquiry)
- **E**: Industry (E-commerce)
- **F**: Response_Approach (Budget inquiry with value proposition)
- **G**: Key_Components (Audit pricing + Call option + Value proposition)
- **H**: Original_Question (Glad I found a fellow player. We have some ads running now...)
- **I**: Immediate_Context (Usually empty or additional context)
- **J**: Context_Summary (Initial outreach from lead, establishing contact and interest)
- **K**: Full_Response (I've managed over $20 million in ad spend across Google and Facebook Ads...)

### Tab 2: "Company_Rules" ✅
Your existing structure is perfect:
- **A**: Date_Updated (2025-09-20)
- **B**: Category (Pricing, Tone, Process, Services)
- **C**: Rule_Title (Standard Pricing Structure)
- **D**: Rule_Description (Google Ads Management: $500-2000/month, Facebook Ads: $300-1500/month, Landing Pages: $2000-5000 one-time)

### Spreadsheet ID
Your Spreadsheet ID is: `1ZhrUxuevNYM2pD_TJ6IH_WRePKKo_QCjCPW-64hKiB4`

## Step 2: Deploy Google Apps Script

### Create New Project
1. Go to [script.google.com](https://script.google.com)
2. Create new project
3. Copy code from `lead_response_apps_script.gs`
4. Update these variables:
   ```javascript
   const SPREADSHEET_ID = '1ZhrUxuevNYM2pD_TJ6IH_WRePKKo_QCjCPW-64hKiB4'; // ✅ Already set
   const API_KEY = 'YOUR_OPENAI_API_KEY_HERE'; // Replace with your actual API key // ✅ Already set
   ```

### Deploy as Web App
1. Click "Deploy" → "New deployment"
2. Type: "Web app"
3. Execute as: "Me"
4. Who has access: "Anyone"
5. Click "Deploy"
6. Copy the web app URL

## Step 3: Configure ChatGPT

### Upload Schema
1. Copy `lead_response_openapi_schema.json`
2. Update the web app URL in the schema
3. Upload to ChatGPT as custom action

### Recommended Model
**Use GPT-5 (latest model)** for optimal performance with:
- Most advanced function calling capabilities
- Superior business communication quality
- Best rule compliance and validation
- Most reliable data formatting
- Latest features and improvements

**Alternative**: GPT-4 if GPT-5 is not available in your region

### Set Instructions
Use the instructions from `GPT_CONCISE_INSTRUCTIONS.md` (copy/paste into Instructions field)
Upload `CHATGPT_INSTRUCTIONS.txt` to Knowledge section for detailed reference

## Step 4: Test the System

### Test Functions
1. **Get Company Rules**: Retrieve Creekside Marketing guidelines from Company_Rules tab
2. **Get Examples**: Search for similar conversations in Data Log tab
3. **Get History**: View lead conversation history by ID or keywords

### Test Workflow
1. Paste sample conversation into ChatGPT
2. Verify it generates response + copy-pasteable Data Log row
3. Test copying the row into your Data Log tab
4. Verify it checks Company_Rules tab first

## Troubleshooting

### Common Issues
- **API Key**: Make sure it matches in both script and schema
- **Spreadsheet ID**: ✅ Already set to `1ZhrUxuevNYM2pD_TJ6IH_WRePKKo_QCjCPW-64hKiB4`
- **Permissions**: Ensure web app is set to "Anyone"
- **Sheet Names**: ✅ "Data Log" and "Company_Rules" tabs already exist

### Testing
- Use the test function in the Google Apps Script
- Check execution logs for errors
- Verify spreadsheet data is being added correctly

## What Each Function Does

### `get_company_rules`
- **Purpose**: Get Creekside Marketing rules and best practices
- **When to Use**: ALWAYS before generating responses
- **Input**: Category, keywords
- **Output**: List of company rules and guidelines

### `get_response_examples`
- **Purpose**: Find similar conversations for ChatGPT to learn from
- **When to Use**: MANDATORY - for every conversation after checking rules
- **Input**: Context keywords, industry, limit
- **Output**: List of similar successful responses from Data Log with weighted prioritization:
  - **🟢 Weight 8**: Last 1 month (highest priority - 8x emphasis)
  - **🟡 Weight 4**: Last 3 months (high priority - 4x emphasis)
  - **🟠 Weight 2**: Older than 3 months (low priority - 2x emphasis)
  - **🔴 Weight 1**: July 2025 and earlier (deprecated - 1x emphasis)

### `get_lead_history`
- **Purpose**: Get conversation history for specific leads
- **When to Use**: Only when VA mentions this is a repeat conversation or provides conversation ID
- **Input**: Conversation ID, context keywords
- **Output**: All previous interactions from Data Log

## Function Call Decision Tree

**For ChatGPT to follow:**

1. **MANDATORY**: Always call `get_company_rules` first (no exceptions)
2. **MANDATORY**: Always call `get_response_examples` for every conversation (no exceptions)
   - Prioritizes recent examples over old ones (communication style evolves)
3. **OPTIONAL**: Call `get_lead_history` only if VA mentions repeat conversation
4. **Generate response** following all company rules + recent examples
5. **Output** response + copy-pasteable Data Log row for VA

**Never call**: Functions for adding/updating data (ChatGPT only outputs data)

## Success Criteria
- ✅ ChatGPT can retrieve company rules from Company_Rules tab
- ✅ ChatGPT can find similar examples from Data Log tab
- ✅ ChatGPT generates responses that follow company rules
- ✅ ChatGPT outputs copy-pasteable Data Log row for VA
- ✅ System can retrieve lead history by conversation ID
- ✅ VA can manually copy/paste data into Google Sheets
- ✅ VAs can use the system efficiently
