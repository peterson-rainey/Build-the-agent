# Facebook Ads Automation Setup Guide

## 🚀 **Option 1: Facebook Ads API + Google Apps Script (Recommended)**

### **Step 1: Get Facebook Ads API Access**

1. **Create a Facebook App**:
   - Go to [Facebook Developers](https://developers.facebook.com/)
   - Click "Create App" → "Business" → "Facebook Login"
   - Name your app (e.g., "Facebook Ads Exporter")

2. **Get App Credentials**:
   - Copy your **App ID** and **App Secret**
   - Go to "Tools" → "Graph API Explorer"
   1760703891989950
   adf35c9de2a7c3796353c5abd1af3526

3. **Generate Access Token**:
   - Select your app from dropdown
   - Add permissions: `ads_read`, `ads_management`
   - Click "Generate Access Token"
   - Copy the **Access Token**

4. **Get Ad Account ID**:
   - Go to [Facebook Ads Manager](https://www.facebook.com/adsmanager)
   - Your Ad Account ID is in the URL: `act_123456789`

### **Step 2: Configure Google Apps Script**

1. **Open Google Apps Script**:
   - Go to [script.google.com](https://script.google.com)
   - Create new project

2. **Add the Automation Script**:
   - Copy the code from `Scripts/Facebook_Ads_Automation.gs`
   - Update the CONFIG section with your credentials:

```javascript
const CONFIG = {
  FACEBOOK_APP_ID: '123456789012345',           // Your Facebook App ID
  FACEBOOK_APP_SECRET: 'your_app_secret_here',  // Your Facebook App Secret
  ACCESS_TOKEN: 'your_access_token_here',       // Your Access Token
  AD_ACCOUNT_ID: 'act_123456789',              // Your Ad Account ID
  SPREADSHEET_ID: '1ABC...XYZ',                // Your Google Sheet ID
  RAW_DATA_SHEET: 'Raw Data'
};
```

3. **Get Your Google Sheet ID**:
   - Open your Google Sheet
   - Copy the ID from the URL: `https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit`

### **Step 3: Set Up Automation**

1. **Run Setup Function**:
   - In Google Apps Script, run `setupAutomation()`
   - This will create a daily trigger at 2 AM

2. **Test the Automation**:
   - Run `testAutomation()` to test manually
   - Check your Google Sheet for new data

### **Step 4: Monitor & Troubleshoot**

1. **View Execution Logs**:
   - In Google Apps Script, go to "Executions"
   - Check for any errors

2. **Set Up Notifications**:
   - Uncomment the email notification code in `sendErrorNotification()`
   - Add your email address

---

## 🔄 **Option 2: Facebook Ads Manager Export + Google Apps Script**

### **If you prefer manual exports:**

1. **Set up Google Apps Script**:
   - Create a script that watches for file uploads
   - Automatically processes CSV files

2. **Use Google Drive API**:
   - Monitor a specific folder
   - Process new CSV files automatically

---

## 🔄 **Option 3: Third-Party Tools**

### **Zapier Integration**:
1. Create a Zapier account
2. Set up Facebook Ads → Google Sheets trigger
3. Configure daily export

### **Supermetrics**:
1. Use Supermetrics for Google Sheets
2. Set up Facebook Ads data source
3. Schedule daily refresh

---

## ⚙️ **Configuration Options**

### **Customize Export Schedule**:
```javascript
// Change trigger time (currently 2 AM)
ScriptApp.newTrigger('dailyFacebookAdsExport')
  .timeBased()
  .everyDays(1)
  .atHour(9)  // Change to 9 AM
  .create();
```

### **Export Different Date Ranges**:
```javascript
// Change from 'yesterday' to 'last_7d', 'last_30d', etc.
date_preset: 'last_7d'
```

### **Add More Fields**:
```javascript
// Add more fields to export
fields: 'campaign_name,adset_name,date_start,date_stop,impressions,clicks,spend,reach,frequency,results,result_type,cpm,cpc,ctr'
```

---

## 🛠️ **Troubleshooting**

### **Common Issues**:

1. **"Invalid Access Token"**:
   - Regenerate your access token
   - Check token permissions

2. **"Ad Account Not Found"**:
   - Verify your Ad Account ID
   - Ensure you have access to the account

3. **"Rate Limit Exceeded"**:
   - Add delays between API calls
   - Reduce export frequency

4. **"Script Timeout"**:
   - Break large exports into chunks
   - Use batch processing

### **Testing Commands**:
```javascript
// Test API connection
testFacebookAPIConnection()

// Test data export
testDataExport()

// Check trigger status
checkTriggers()
```

---

## 📊 **Monitoring & Alerts**

### **Set Up Email Notifications**:
```javascript
function sendSuccessNotification() {
  MailApp.sendEmail({
    to: 'your-email@example.com',
    subject: 'Facebook Ads Export Success',
    body: 'Daily Facebook Ads export completed successfully'
  });
}
```

### **Slack Integration**:
```javascript
function sendSlackNotification(message) {
  const webhookUrl = 'YOUR_SLACK_WEBHOOK_URL';
  const payload = { text: message };
  
  UrlFetchApp.fetch(webhookUrl, {
    method: 'POST',
    payload: JSON.stringify(payload),
    contentType: 'application/json'
  });
}
```

---

## 🔒 **Security Best Practices**

1. **Store Credentials Securely**:
   - Use Google Apps Script Properties
   - Don't hardcode in script

2. **Limit API Permissions**:
   - Only request necessary permissions
   - Use read-only tokens when possible

3. **Monitor Usage**:
   - Check API call limits
   - Monitor for unusual activity

---

## 📈 **Advanced Features**

### **Incremental Updates**:
- Only export new data
- Append to existing data
- Avoid duplicates

### **Multiple Ad Accounts**:
- Export from multiple accounts
- Combine data in one sheet

### **Custom Metrics**:
- Calculate custom KPIs
- Add business-specific metrics

---

## 🎯 **Next Steps**

1. **Choose your preferred method** (API recommended)
2. **Set up Facebook App and get credentials**
3. **Configure Google Apps Script**
4. **Test the automation**
5. **Monitor and optimize**

The API method is the most reliable and gives you the most control over the data export process.
