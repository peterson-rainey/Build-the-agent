# Facebook Ads Reporting System - Complete Setup Guide

## Overview

This guide will walk you through setting up a complete Facebook Ads reporting system in Google Sheets that automatically processes data and generates comprehensive reports.

## Prerequisites

- Google account with access to Google Sheets
- Facebook Ads account with data to export
- Basic understanding of Google Sheets

## Step-by-Step Setup

### Step 1: Create the Google Sheet

1. **Open Google Sheets**
   - Go to [sheets.google.com](https://sheets.google.com)
   - Sign in with your Google account

2. **Create New Spreadsheet**
   - Click the "+" button to create a new spreadsheet
   - Rename it to "Facebook Ads Reporting System"

3. **Set Up Raw Data Sheet**
   - Rename the first sheet to "Raw Data"
   - Add the following headers in row 1:

```
A1: Campaign Name
B1: Ad Set Name
C1: Ad Name
D1: Date
E1: Impressions
F1: Clicks
G1: Spend
H1: Results
I1: Cost per Result
J1: Reach
K1: Frequency
L1: CPM
M1: CPC
N1: CTR
O1: Relevance Score
P1: Quality Ranking
Q1: Engagement Rate Ranking
R1: Conversion Rate Ranking
S1: Revenue
T1: ROAS
```

4. **Format the Header Row**
   - Select row 1 (A1:T1)
   - Make text bold (Ctrl+B or Cmd+B)
   - Set background color to blue (#4285f4)
   - Set text color to white
   - Center align the text

5. **Add Filters**
   - Select row 1 (A1:T1)
   - Click Data → Create a filter
   - You should see filter icons appear in each header cell

6. **Set Up Data Validation**
   - Select column D (Date column)
   - Click Data → Data validation
   - Set criteria to "Date"
   - Add custom validation message: "Please enter a valid date"

### Step 2: Install Google Apps Script

1. **Open Apps Script**
   - In your Google Sheet, click Extensions → Apps Script
   - This will open a new tab with the Apps Script editor

2. **Clear Default Code**
   - Delete any existing code in the editor
   - You should have a blank editor

3. **Copy the Script**
   - Open the `Facebook_Ads_Data_Processor.gs` file
   - Copy all the code
   - Paste it into the Apps Script editor

4. **Save the Project**
   - Click File → Save
   - Name the project "Facebook Ads Data Processor"
   - Click "Save"

5. **Authorize the Script**
   - Click "Run" → "initializeReportingSystem"
   - Google will ask for permissions
   - Click "Review Permissions"
   - Choose your Google account
   - Click "Advanced" → "Go to Facebook Ads Data Processor (unsafe)"
   - Click "Allow"

### Step 3: Set Up Automatic Triggers

1. **Open Triggers**
   - In Apps Script, click on the clock icon (Triggers)
   - This will open the triggers page

2. **Create Edit Trigger**
   - Click "Add Trigger"
   - Configure as follows:
     - Choose function: `processFacebookAdsData`
     - Choose event source: "From spreadsheet"
     - Choose event type: "On edit"
     - Click "Save"

3. **Create Time-based Trigger**
   - Click "Add Trigger" again
   - Configure as follows:
     - Choose function: `processFacebookAdsData`
     - Choose event source: "Time-driven"
     - Choose type of time-based trigger: "Hours timer"
     - Select time interval: "Every 6 hours"
     - Click "Save"

### Step 4: Test the System

1. **Add Sample Data**
   - Go back to your Google Sheet
   - In the Raw Data sheet, add some sample data below the headers
   - Example data:

```
Campaign A | Ad Set 1 | Ad 1 | 2024-01-01 | 1000 | 50 | 100 | 5 | 20 | 800 | 1.25 | 100 | 2 | 5 | 8 | Good | Above Average | Above Average | 200 | 2
Campaign A | Ad Set 1 | Ad 2 | 2024-01-01 | 800 | 40 | 80 | 4 | 20 | 600 | 1.33 | 100 | 2 | 5 | 7 | Good | Above Average | Average | 160 | 2
Campaign B | Ad Set 2 | Ad 3 | 2024-01-01 | 1200 | 60 | 120 | 6 | 20 | 1000 | 1.2 | 100 | 2 | 5 | 9 | Excellent | Above Average | Above Average | 240 | 2
```

2. **Trigger Processing**
   - The script should automatically run when you add data
   - If not, go to Apps Script and run `manualProcess()` function

3. **Check Results**
   - Look for new sheets created:
     - "Processed Data"
     - "Report"
     - "Dashboard"
   - Verify that calculations are correct

### Step 5: Import Real Facebook Ads Data

1. **Export from Facebook Ads Manager**
   - Go to Facebook Ads Manager
   - Select the campaigns you want to analyze
   - Click "Export" → "Export to CSV"
   - Download the CSV file

2. **Prepare the Data**
   - Open the CSV file in Excel or Google Sheets
   - Make sure the column headers match the expected format
   - Remove any unnecessary columns
   - Clean up any formatting issues

3. **Import to Google Sheet**
   - Copy the data from your prepared file
   - Paste it into the Raw Data sheet below the headers
   - The script will automatically process the new data

## Configuration Options

### Customizing Performance Thresholds

You can modify the performance thresholds in the script:

1. **Open the Script**
   - Go to Apps Script editor
   - Find the `CONFIG.THRESHOLDS` section

2. **Adjust Values**
   ```javascript
   THRESHOLDS: {
     CTR_MIN: 0.01, // 1% - minimum acceptable CTR
     CPC_MAX: 5.00, // $5 - maximum acceptable CPC
     CPM_MAX: 50.00, // $50 - maximum acceptable CPM
     ROAS_MIN: 2.00 // 2:1 - minimum acceptable ROAS
   }
   ```

3. **Save and Test**
   - Save the script
   - Run `manualProcess()` to test with new thresholds

### Adding Custom Metrics

1. **Modify Column Mappings**
   - Find `CONFIG.COLUMN_MAPPINGS`
   - Add new Facebook Ads columns to the mapping

2. **Add Calculation Functions**
   - Create new functions to calculate custom metrics
   - Update the `calculateMetrics()` function

3. **Update Report Templates**
   - Modify the report generation functions
   - Add new sections to the Report sheet

### Setting Up Notifications

1. **Email Notifications**
   - Uncomment the email code in `sendProcessingNotification()`
   - Replace `'your-email@example.com'` with your email
   - Save the script

2. **Slack Notifications**
   - Add Slack webhook integration
   - Modify the notification function to send to Slack

## Troubleshooting

### Common Issues and Solutions

#### Issue: Script Not Running
**Symptoms**: No new sheets created, no processing happening
**Solutions**:
1. Check Apps Script logs for errors
2. Verify triggers are set up correctly
3. Ensure script has proper permissions
4. Try running `manualProcess()` function

#### Issue: Data Not Processing
**Symptoms**: Raw data exists but processed data is empty
**Solutions**:
1. Check column headers match expected format
2. Verify data types (numbers vs text)
3. Look for missing required fields
4. Check Apps Script logs for validation errors

#### Issue: Calculations Incorrect
**Symptoms**: Metrics don't match expected values
**Solutions**:
1. Verify data format in Raw Data sheet
2. Check for empty or invalid values
3. Review calculation formulas in the script
4. Test with known good data

#### Issue: Sheets Not Created
**Symptoms**: Script runs but no new sheets appear
**Solutions**:
1. Check script permissions
2. Verify sheet names in CONFIG
3. Look for naming conflicts
4. Check Apps Script logs for errors

### Debugging Steps

1. **Check Execution Logs**
   - In Apps Script, click "Executions"
   - Look for recent executions
   - Check for error messages

2. **Test Individual Functions**
   - Run `processRawData()` to test data processing
   - Run `generateReports()` to test report generation
   - Run `updateDashboard()` to test dashboard creation

3. **Verify Data Format**
   - Check that Raw Data sheet has correct headers
   - Ensure data is in the right format
   - Look for missing or invalid values

4. **Check Permissions**
   - Ensure script has access to the spreadsheet
   - Verify Google account permissions
   - Check for any security restrictions

## Advanced Features

### Automated Data Import

1. **Set Up Facebook Ads API**
   - Get API access token
   - Configure API permissions
   - Set up automated data extraction

2. **Create Import Function**
   - Add function to fetch data from Facebook API
   - Schedule regular imports
   - Handle API rate limits

### Custom Report Templates

1. **Modify Report Layout**
   - Change section headers
   - Add new metrics
   - Customize formatting

2. **Add Visual Elements**
   - Create charts and graphs
   - Add conditional formatting
   - Include performance indicators

### Integration with Other Tools

1. **Google Data Studio**
   - Connect processed data to Data Studio
   - Create interactive dashboards
   - Share reports with stakeholders

2. **Slack Integration**
   - Send daily/weekly reports to Slack
   - Alert on performance issues
   - Share insights automatically

## Maintenance

### Regular Tasks

1. **Weekly**
   - Review performance thresholds
   - Check for new Facebook Ads features
   - Validate report accuracy

2. **Monthly**
   - Update script with improvements
   - Review and optimize performance
   - Backup important data

3. **Quarterly**
   - Major system updates
   - Performance optimization
   - User feedback integration

### Performance Optimization

1. **Script Efficiency**
   - Minimize API calls
   - Optimize data processing
   - Use efficient data structures

2. **Sheet Performance**
   - Limit data size
   - Use efficient formulas
   - Regular cleanup of old data

## Support and Resources

### Documentation
- Google Apps Script documentation
- Facebook Ads API documentation
- Google Sheets API documentation

### Community Resources
- Google Apps Script community
- Facebook Ads community
- Stack Overflow for technical questions

### Getting Help
1. Check the troubleshooting section above
2. Review Apps Script execution logs
3. Search community forums
4. Contact support if needed

## Next Steps

Once your system is set up and running:

1. **Optimize Performance**
   - Fine-tune thresholds based on your data
   - Add custom metrics relevant to your business
   - Set up automated alerts

2. **Scale the System**
   - Add more data sources
   - Create additional report types
   - Integrate with other tools

3. **Share and Collaborate**
   - Share reports with team members
   - Set up automated distribution
   - Create presentation templates

4. **Continuous Improvement**
   - Gather user feedback
   - Monitor system performance
   - Plan future enhancements 