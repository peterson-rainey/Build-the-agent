# Email Processing System

A complete system for extracting emails from MBOX files and filtering them by specific users for tone training and analysis.

## 📁 Folder Structure

```
Email_Processing/
├── raw_data/           # Original MBOX files (place your .mbox files here)
├── processed_data/     # Extracted email datasets (output appears here)
├── extract_user_emails.py    # Script for processing MBOX files
├── extract_from_json.py      # Script for processing JSON files
├── setup.py           # Setup script (run this first!)
└── README.md          # This documentation
```

## 🚀 Quick Start

### Step 1: Setup (Run Once)
```bash
python3 setup.py
```

### Step 2: Process Your Emails
1. **Place your MBOX file** in the `raw_data/` folder
2. **Extract emails from any user:**
   ```bash
   python3 extract_user_emails.py "raw_data/your_emails.mbox" "user@example.com"
   ```

### Examples

**Extract Peterson's emails:**
```bash
python3 extract_user_emails.py "raw_data/All mail Including Spam and Trash.mbox" "peterson@fullcirclemedia.co"
```

**Extract John's emails with custom output:**
```bash
python3 extract_user_emails.py "raw_data/emails.mbox" "john@company.com" "processed_data/john_emails.json"
```

**Extract from already-converted JSON:**
```bash
python3 extract_from_json.py "processed_data/all_emails_complete.json" "user@example.com"
```

## 📋 What the Scripts Do

### extract_user_emails.py (MBOX → User Emails)
1. **Converts MBOX to JSON** using the mbox-to-json tool
2. **Filters emails** sent by the specified user
3. **Cleans message content** (removes signatures, excessive whitespace)
4. **Saves structured data** with essential fields

### extract_from_json.py (JSON → User Emails)
1. **Reads already-converted JSON** file
2. **Filters emails** sent by the specified user
3. **Cleans message content** (removes signatures, excessive whitespace)
4. **Saves structured data** with essential fields

## 📊 Output Format

The extracted emails are saved as JSON with this structure:

```json
[
  {
    "email_number": 1,
    "from": "user@example.com",
    "to": "recipient@example.com",
    "subject": "Email Subject",
    "date": "Mon, 6 Oct 2025 17:50:56 -0500",
    "message_body": "Clean email content...",
    "labels": "Sent",
    "message_id": "<message-id>"
  }
]
```

## 🎯 Use Cases

- **ChatGPT Tone Training** - Extract your emails to train AI on your writing style
- **Email Analysis** - Analyze communication patterns
- **Data Export** - Convert email archives to structured data
- **Backup Processing** - Process email backups for specific users

## 🔧 Advanced Usage

### Custom Output Location
```bash
python3 extract_user_emails.py input.mbox user@email.com custom_output.json
```

### Processing Multiple Users
```bash
# Extract emails for different users
python3 extract_user_emails.py emails.mbox user1@company.com user1_emails.json
python3 extract_user_emails.py emails.mbox user2@company.com user2_emails.json
```

### Two-Step Process (Recommended for Large Files)
```bash
# Step 1: Convert MBOX to JSON (once)
mbox-to-json raw_data/emails.mbox -o processed_data/all_emails.json

# Step 2: Extract specific users (multiple times)
python3 extract_from_json.py processed_data/all_emails.json user1@company.com
python3 extract_from_json.py processed_data/all_emails.json user2@company.com
```

## 📁 File Management

- **Raw MBOX files** go in `raw_data/`
- **Processed JSON files** go in `processed_data/`
- **Scripts** are in the main folder

## 🛠️ Troubleshooting

### Common Issues

1. **"mbox-to-json not found"**
   - Run: `python3 setup.py`
   - Or install manually: `pip install mbox-to-json`

2. **"No emails found"**
   - Check the email address spelling
   - Verify the user actually sent emails in the MBOX file

3. **"Input file not found"**
   - Make sure the MBOX file path is correct
   - Use relative paths from the Email_Processing directory

### Getting Help

Run any script without arguments to see usage:
```bash
python3 extract_user_emails.py
python3 extract_from_json.py
```

## 📈 Performance

- **Processing time**: ~3 seconds for 382 emails
- **Output size**: Typically 1-10% of original MBOX size
- **Memory usage**: Minimal (processes one email at a time)

## 🔒 Privacy & Security

- Scripts only process local files
- No data is sent to external services
- Temporary files are automatically cleaned up
- Original MBOX files are never modified

---

**Created for**: Email tone training and analysis  
**Last Updated**: October 2025  
**Version**: 2.0