# Weekly Email Reminder Setup Guide

## Overview
This guide will help you set up an automated weekly email reminder for your Sweet Hands Q4 GTM Schedule using Google Ads Scripts.

## Step 1: Copy the Script
1. Open the `Weekly_Email_Reminder.gs` file
2. Copy all the code from the file
3. Go to [Google Ads Scripts](https://ads.google.com/um/Welcome/Home?subid=US-en-AW-G-ACQ-Scripts)
4. Click "Create new script"
5. Paste the code into the editor

## Step 2: Update Configuration
The script is already configured to send emails to both team members:

```javascript
const EMAIL_ADDRESSES = [
  'peterson@creeksidemarketingpros.com',
  'sophiansnow2@gmail.com'
];
```

If you need to add or remove email addresses, simply edit this array.

## Step 3: Test the Script
Before setting up the automatic trigger, test the script manually:

1. In the Google Ads Scripts editor, click the "Run" button
2. Select the `testEmail` function from the dropdown
3. Click "Run" to send a test email
4. Check your email to confirm it was received
5. Check the logs to ensure no errors occurred

## Step 4: Set Up the Trigger
To make the script run automatically every Monday:

1. In the Google Ads Scripts editor, click on "Triggers" (clock icon)
2. Click "Create new trigger"
3. Configure the trigger:
   - **Function to run**: `main`
   - **Event source**: `Time-driven`
   - **Type of time-based trigger**: `Week timer`
   - **Day of the week**: `Monday`
   - **Time of day**: `9:00 AM` (or your preferred time)
4. Click "Save"

## Step 5: Verify Setup
To verify everything is working:

1. Run the `checkScriptStatus` function to see current settings
2. Wait for the next Monday to see if the email is sent automatically
3. Check the logs after the trigger runs to confirm success

## Troubleshooting

### Email Not Received
- Check your spam folder
- Verify the email address is correct
- Check the script logs for errors
- Ensure you have permission to send emails from Google Ads

### Script Not Running
- Verify the trigger is set up correctly
- Check that the `main` function is selected in the trigger
- Ensure the script is saved and enabled

### Permission Issues
- Make sure you're logged into the correct Google Ads account
- Check that you have admin access to the account
- Verify the script has the necessary permissions

## Customization Options

### Change Email Content
You can modify the email subject and body by editing these constants:
```javascript
const EMAIL_SUBJECT = 'Check Sweet Hands Q4 GTM Schedule';
const EMAIL_BODY = `...`;
```

### Change Timing
To change when the email is sent:
1. Edit the trigger settings
2. Choose a different day or time
3. Or modify the script to send on different days

### Add or Remove Recipients
To modify the email recipients, edit the EMAIL_ADDRESSES array:
```javascript
const EMAIL_ADDRESSES = [
  'peterson@creeksidemarketingpros.com',
  'sophiansnow2@gmail.com',
  'another-team-member@example.com'  // Add more emails here
];
```

## Security Notes
- The script uses `noReply: true` to prevent reply emails
- Your email address is stored in the script - keep it secure
- The script only runs on Mondays to avoid spam
- All actions are logged for audit purposes

## Support
If you encounter issues:
1. Check the script logs for error messages
2. Verify all configuration settings
3. Test with the `testEmail` function
4. Ensure the Google Ads account has proper permissions
