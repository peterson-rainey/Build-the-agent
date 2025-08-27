# Email Testing Troubleshooting Guide

## Why You Might Not Receive the Test Email

There are several common reasons why the test email function might not work. Let's troubleshoot step by step.

## Step 1: Check the Script Logs First

**This is the most important step!** Always check the logs to see what happened:

1. In Google Ads Scripts, after running the test function
2. Click on "Execution log" or "Logs" tab
3. Look for any error messages or success messages

## Step 2: Run the Debug Functions

I've added several debug functions to help identify the issue. Run these in order:

### 1. Check Permissions
```javascript
checkPermissions()
```
This will verify that the script has access to MailApp and other services.

### 2. Check Script Status
```javascript
checkScriptStatus()
```
This will show your current configuration and settings.

### 3. Debug Email Test
```javascript
debugEmail()
```
This sends a simple test email to just the first email address.

### 4. Full Test Email
```javascript
testEmail()
```
This runs the complete test with all the logging.

## Step 3: Common Issues and Solutions

### Issue 1: "MailApp is not available"
**Solution**: This means you're not in the right environment. Make sure you're:
- In Google Ads Scripts (not Google Apps Script)
- Logged into the correct Google Ads account
- Have admin access to the account

### Issue 2: "Permission denied" or "Access denied"
**Solution**: 
1. Make sure you're an admin on the Google Ads account
2. Try running the script from a different browser
3. Clear browser cache and cookies
4. Log out and log back into Google Ads

### Issue 3: Email sent but not received
**Solutions**:
1. **Check spam folder** - Most common issue!
2. **Check email filters** - Your email might be filtering it
3. **Wait 5-10 minutes** - Sometimes there's a delay
4. **Check the "To" field** - Make sure the email address is correct

### Issue 4: Script runs but no email sent
**Check the logs for**:
- Error messages about MailApp
- Permission errors
- Network errors

## Step 4: Step-by-Step Testing Process

### First Test: Basic Permissions
1. Run `checkPermissions()`
2. Check logs for any "✗" marks
3. If all show "✓", proceed to next test

### Second Test: Simple Email
1. Run `debugEmail()`
2. Check logs for success message
3. Check your email (and spam folder)
4. If this works, proceed to full test

### Third Test: Full Email
1. Run `testEmail()`
2. Check logs for detailed information
3. Check both email addresses

## Step 5: Alternative Testing Methods

### Method 1: Test in Google Apps Script
If Google Ads Scripts isn't working, you can test the email functionality in Google Apps Script:

1. Go to [script.google.com](https://script.google.com)
2. Create a new project
3. Copy just the email sending code
4. Test there first

### Method 2: Manual Email Test
Create a simple test script:
```javascript
function simpleTest() {
  try {
    MailApp.sendEmail({
      to: 'your-email@example.com',
      subject: 'Simple Test',
      body: 'This is a test'
    });
    Logger.log('Email sent');
  } catch (e) {
    Logger.log('Error: ' + e.toString());
  }
}
```

## Step 6: What to Look for in Logs

### Successful Execution:
```
=== Starting test email function ===
Current time: [timestamp]
Attempting to send email to: peterson@creeksidemarketingpros.com, sophiansnow2@gmail.com
Email subject: Check Sweet Hands Q4 GTM Schedule
Email body length: 456 characters
Weekly reminder email sent successfully to: peterson@creeksidemarketingpros.com, sophiansnow2@gmail.com
Test email sent successfully
=== Test email function completed ===
```

### Common Error Messages:
- `MailApp is not available` → Wrong environment
- `Permission denied` → Account access issue
- `Invalid email address` → Check email format
- `Quota exceeded` → Too many emails sent recently

## Step 7: Still Not Working?

If none of the above works:

1. **Try a different browser** (Chrome, Firefox, Safari)
2. **Use incognito/private mode**
3. **Check if you have multiple Google accounts** - make sure you're logged into the right one
4. **Contact Google Ads support** if it's a platform issue

## Step 8: Success Checklist

✅ Script runs without errors  
✅ Logs show "email sent successfully"  
✅ Email received in inbox (not spam)  
✅ Both email addresses receive the email  
✅ Email contains the correct content and link  

Once you've completed this checklist, your script is ready for the automatic trigger!
