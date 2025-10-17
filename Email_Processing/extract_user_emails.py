#!/usr/bin/env python3
"""
Email Extraction Tool for Specific Users
Converts MBOX files to JSON and extracts emails from a specific user.

Usage:
    python3 extract_user_emails.py <input.mbox> <target_email> [output.json]

Example:
    python3 extract_user_emails.py "All mail Including Spam and Trash.mbox" "peterson@fullcirclemedia.co" "peterson_emails.json"
"""

import json
import re
import sys
import os
from datetime import datetime

def extract_user_emails(mbox_file, target_email, output_file=None):
    """
    Extract emails sent by a specific user from an MBOX file.
    
    Args:
        mbox_file (str): Path to the input MBOX file
        target_email (str): Email address of the user to extract emails from
        output_file (str): Output JSON file path (optional)
    
    Returns:
        int: Number of emails found
    """
    
    # Set default output file if not provided
    if not output_file:
        base_name = os.path.splitext(os.path.basename(mbox_file))[0]
        output_file = f"{base_name}_{target_email.split('@')[0]}_emails.json"
    
    print(f"🔍 Extracting emails from: {mbox_file}")
    print(f"👤 Target user: {target_email}")
    print(f"📁 Output file: {output_file}")
    print("-" * 50)
    
    # Step 1: Convert MBOX to JSON using mbox-to-json
    print("📧 Converting MBOX to JSON...")
    temp_json_file = "temp_emails.json"
    
    # Run mbox-to-json command
    import subprocess
    try:
        result = subprocess.run([
            "mbox-to-json", mbox_file, "-o", temp_json_file
        ], capture_output=True, text=True, check=True)
        print("✅ MBOX conversion completed")
    except subprocess.CalledProcessError as e:
        print(f"❌ Error converting MBOX: {e}")
        return 0
    except FileNotFoundError:
        print("❌ mbox-to-json not found. Please install it with: pip install mbox-to-json")
        return 0
    
    # Step 2: Load and process JSON
    print("📖 Reading converted email data...")
    try:
        with open(temp_json_file, 'r', encoding='utf-8') as f:
            emails = json.load(f)
        print(f"📊 Found {len(emails)} total emails")
    except Exception as e:
        print(f"❌ Error reading JSON file: {e}")
        return 0
    
    # Step 3: Extract emails from target user
    print(f"🔍 Searching for emails from {target_email}...")
    user_emails = []
    
    for i, email in enumerate(emails):
        try:
            from_field = email.get("From", "")
            
            # Check if this email was sent by the target user
            is_target_user = (
                target_email.lower() in from_field.lower() or
                target_email.split('@')[0].lower() in from_field.lower()
            )
            
            if is_target_user:
                cleaned_email = {
                    "email_number": len(user_emails) + 1,
                    "from": from_field,
                    "to": email.get("To", ""),
                    "subject": email.get("Subject", ""),
                    "date": email.get("Date", ""),
                    "message_body": clean_message_body(email.get("Body", "")),
                    "labels": email.get("X-Gmail-Labels", ""),
                    "message_id": email.get("Message-ID", "")
                }
                
                # Only include emails with meaningful content
                if (cleaned_email["message_body"] and 
                    len(cleaned_email["message_body"].strip()) > 10):
                    user_emails.append(cleaned_email)
                    print(f"✅ Found email #{len(user_emails)}: {cleaned_email['subject'][:50]}...")
                    
        except Exception as e:
            print(f"⚠️  Error processing email {i+1}: {e}")
            continue
    
    # Step 4: Save results
    print(f"\n💾 Saving {len(user_emails)} emails to {output_file}...")
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(user_emails, f, indent=2, ensure_ascii=False)
        print(f"✅ Successfully saved to {output_file}")
    except Exception as e:
        print(f"❌ Error saving file: {e}")
        return 0
    
    # Step 5: Clean up temporary file
    try:
        os.remove(temp_json_file)
        print("🧹 Cleaned up temporary files")
    except:
        pass
    
    # Step 6: Show summary
    print("\n" + "="*50)
    print("📊 EXTRACTION SUMMARY")
    print("="*50)
    print(f"📧 Total emails processed: {len(emails)}")
    print(f"👤 Emails from {target_email}: {len(user_emails)}")
    print(f"📁 Output file: {output_file}")
    print(f"📏 File size: {os.path.getsize(output_file) / 1024:.1f} KB")
    
    if user_emails:
        print(f"\n📝 Sample emails found:")
        for i, email in enumerate(user_emails[:3]):
            print(f"  {i+1}. {email['subject'][:60]}...")
    
    return len(user_emails)

def clean_message_body(body):
    """
    Clean and format the email body text.
    """
    if not body:
        return ""
    
    # Remove excessive whitespace and newlines
    cleaned = re.sub(r'\s+', ' ', body)
    
    # Remove common email signatures and footers
    cleaned = re.sub(r'--\s*\n.*', '', cleaned, flags=re.DOTALL)
    cleaned = re.sub(r'This email was sent.*', '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'Unsubscribe.*', '', cleaned, flags=re.IGNORECASE)
    
    # Remove email addresses in brackets
    cleaned = re.sub(r'<[^>]+>', '', cleaned)
    
    # Clean up and return
    return cleaned.strip()

def main():
    """
    Main function to handle command line arguments.
    """
    if len(sys.argv) < 3:
        print("❌ Usage: python3 extract_user_emails.py <input.mbox> <target_email> [output.json]")
        print("\nExample:")
        print('  python3 extract_user_emails.py "All mail Including Spam and Trash.mbox" "peterson@fullcirclemedia.co"')
        print('  python3 extract_user_emails.py "emails.mbox" "john@company.com" "john_emails.json"')
        sys.exit(1)
    
    mbox_file = sys.argv[1]
    target_email = sys.argv[2]
    output_file = sys.argv[3] if len(sys.argv) > 3 else None
    
    # Check if input file exists
    if not os.path.exists(mbox_file):
        print(f"❌ Input file not found: {mbox_file}")
        sys.exit(1)
    
    # Extract emails
    count = extract_user_emails(mbox_file, target_email, output_file)
    
    if count > 0:
        print(f"\n🎉 Successfully extracted {count} emails!")
    else:
        print(f"\n😞 No emails found from {target_email}")
        sys.exit(1)

if __name__ == "__main__":
    main()
