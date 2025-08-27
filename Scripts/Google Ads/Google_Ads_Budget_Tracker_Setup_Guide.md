# Google Ads Budget Tracker Setup Guide

## Overview
This setup guide will help you implement a Google Ads script that automatically tracks your monthly budget spending and daily averages, outputting the data to your Google Sheets tracking template.

## What the Script Does
- **Total Monthly Spend**: Calculates total spend across all campaigns (active and paused) for the current month
- **Average Daily Spend**: Calculates average daily spend based on current day of the month
- **Total Daily Budget**: Sums all individual campaign daily budgets
- **Automatic Tab Detection**: Automatically finds and updates the most recent tab in your spreadsheet
- **Currency Formatting**: Outputs all values formatted as USD currency

## Setup Instructions

### Step 1: Google Ads Script Setup

1. **Access Google Ads Scripts**:
   - Go to your Google Ads account
   - Navigate to Tools & Settings > Bulk Actions > Scripts
   - Click the "+" button to create a new script

2. **Copy the Google Ads Script**:
   - Copy the entire contents of `Google_Ads_Budget_Tracker.gs`
   - Paste it into the Google Ads script editor
   - Save the script with a name like "Budget Tracker"

3. **Test the Script**:
   - Click "Preview" to test the script
   - Check the logs to ensure it's working correctly
   - Verify the data is being written to your spreadsheet

### Step 2: Google Apps Script Setup (Optional but Recommended)

1. **Access Google Apps Script**:
   - Go to [script.google.com](https://script.google.com)
   - Create a new project
   - Name it "Spreadsheet Tab Manager"

2. **Copy the Apps Script**:
   - Copy the entire contents of `Spreadsheet_Tab_Manager.gs`
   - Paste it into the Apps Script editor
   - Save the project

3. **Test the Apps Script**:
   - Run the `testTabManager()` function
   - Check the logs to ensure it can access your spreadsheet

### Step 3: Set Up Automated Execution

1. **In Google Ads Scripts**:
   - Go back to your Google Ads script
   - Click "Schedule" to set up automated execution
   - Set it to run daily at your preferred time (e.g., 9:00 AM)
   - Choose your timezone

2. **Optional: Set up weekly tab creation**:
   - In Google Apps Script, you can set up a trigger to run `createWeeklyTab()` weekly
   - This will automatically create new tabs for each week

## Cell Locations
The script will update these cells in your most recent spreadsheet tab:
- **F8**: Total spend for the current month
- **G8**: Average daily spend for the current month  
- **G6**: Sum of all campaign daily budgets

## Troubleshooting

### Common Issues

1. **Permission Errors**:
   - Ensure your Google Ads account has access to the spreadsheet
   - Check that the spreadsheet URL is correct
   - Verify the spreadsheet is not restricted

2. **No Data Returned**:
   - Check that you have campaigns in your account
   - Verify campaigns have spending data for the current month
   - Check the script logs for specific error messages

3. **Wrong Tab Updated**:
   - The script automatically finds the last tab in the spreadsheet
   - Ensure your most recent tab is the last one in the sheet order
   - You can manually reorder tabs if needed

### Debugging Tips

1. **Check Script Logs**:
   - Always review the execution logs after running the script
   - Look for any error messages or warnings
   - Verify the calculated values make sense

2. **Test with Sample Data**:
   - Run the script in preview mode first
   - Check that the values are being calculated correctly
   - Verify the spreadsheet is being updated

3. **Verify GAQL Queries**:
   - The script uses modern GAQL syntax
   - Queries are optimized for performance
   - All metrics are properly converted from micros to dollars

## Customization Options

### Modify Date Range
To change from current month to a different period, modify the `SPEND_QUERY`:
```sql
-- For last 30 days
WHERE segments.date DURING LAST_30_DAYS

-- For specific date range
WHERE segments.date BETWEEN "20250101" AND "20250131"
```

### Add Additional Metrics
You can easily add more metrics by:
1. Modifying the GAQL queries
2. Adding new calculation functions
3. Updating the `updateSpreadsheet()` function

### Change Cell Locations
To update different cells, modify the `updateSpreadsheet()` function:
```javascript
mostRecentSheet.getRange("YOUR_CELL").setValue(formattedValue);
```

## Support
If you encounter any issues:
1. Check the script logs for error messages
2. Verify your Google Ads account has the necessary permissions
3. Ensure your spreadsheet is accessible and not restricted
4. Test the script in preview mode before scheduling

## Notes
- The script runs efficiently with minimal API calls
- All monetary values are converted from micros to dollars automatically
- The script includes comprehensive error handling
- Daily execution ensures your data stays current
