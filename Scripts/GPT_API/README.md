# Lead Response Processing System

## What This System Does
Helps VAs process lead conversations and generate appropriate responses using ChatGPT and Google Sheets.

## Workflow
1. **Lead messages** → VA copies conversation to ChatGPT
2. **ChatGPT finds examples** → Generates response + spreadsheet row
3. **VA sends to boss** → Boss approves/revises → VA sends final response
4. **VA updates spreadsheet** with final response

## Files (5 total)

### Core System Files
- `lead_response_apps_script.gs` - Google Apps Script backend
- `lead_response_openapi_schema.json` - OpenAPI schema for ChatGPT

### Documentation
- `SETUP_GUIDE.md` - Complete setup instructions
- `CHATGPT_INSTRUCTIONS.md` - Instructions for ChatGPT
- `README.md` - This overview

## Quick Start

1. **Set up Google Sheets** (see SETUP_GUIDE.md)
2. **Deploy Google Apps Script** (see SETUP_GUIDE.md)
3. **Configure ChatGPT** (see SETUP_GUIDE.md)
4. **Use ChatGPT Instructions** (see CHATGPT_INSTRUCTIONS.md)

## What Each Function Does

- **`get_response_examples`** - Find similar conversations to learn from
- **`add_lead_response`** - Save new lead interactions
- **`update_lead_response`** - Update responses when revised
- **`get_lead_history`** - View lead conversation history

## Benefits

- **Faster responses** - VAs generate responses quickly
- **Better quality** - Based on successful examples
- **Consistent branding** - Maintains company voice
- **Scalable process** - Handles high volume
- **Continuous improvement** - Learns from past successes