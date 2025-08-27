# Google Ads Error Monitor

A comprehensive Google Ads script that monitors multiple accounts for critical errors and sends email notifications when issues are detected.

## 🚀 Features

### Error Monitoring
- **Account-Level Issues**: Suspended accounts, billing issues, script failures
- **Ad-Level Issues**: Disapproved ads
- **Keyword-Level Issues**: Disapproved keywords, low quality scores, low search volume
- **Performance Issues**: Campaigns with no conversions
- **Shopping Campaign Issues**: Missing merchant IDs
- **Spending Issues**: No spend detection with ad schedule awareness
- **Landing Page Issues**: Invalid URLs
- **Targeting Issues**: Overly broad geographic targeting
- **Conversion Tracking Issues**: Missing or inactive conversion actions
- **Feed Issues**: Missing Merchant Center connections

### Key Features
- **MCC Script Monitoring**: Checks all scripts in the MCC for failures
- **Ad Schedule Awareness**: Only flags spending issues when campaigns should be running
- **Campaign-Level Monitoring**: Individual campaign performance tracking
- **Enhanced Conversion Tracking**: Multiple fallback methods for access issues
- **Email Notifications**: Detailed error reports with recommendations

## 📁 Project Structure

```
├── Scripts/
│   ├── Google Ads/
│   │   ├── google ads error monitor          # Main error monitoring script
│   │   ├── Google_Ads_Budget_Tracker.gs      # Budget tracking script
│   │   ├── MCC_Budget_Tracker.gs             # MCC-level budget tracking
│   │   └── Spreadsheet_Tab_Manager.gs        # Spreadsheet management
│   ├── Facebook Ads Report/                  # Facebook Ads automation scripts
│   └── sweethands email script/              # Email automation scripts
└── Prompts/                                  # AI prompts and documentation
```

## 🛠️ Setup Instructions

### 1. Google Ads Script Setup
1. Open Google Ads Scripts
2. Create a new script
3. Copy the contents of `Scripts/Google Ads/google ads error monitor`
4. Configure the following variables:
   - `MASTER_SPREADSHEET_URL`: Your account mappings spreadsheet
   - `EMAIL_RECIPIENTS`: Array of email addresses to notify
   - `SINGLE_CID_FOR_TESTING`: Account ID for testing (optional)

### 2. Account Mappings Spreadsheet
Create a Google Sheet with the following structure:
- **Sheet Name**: "AccountMappings"
- **Columns**: 
  - Account ID
  - Account Name
  - Spreadsheet URL (optional)

### 3. Script Scheduling
- **Daily Monitoring**: Run `setupDailyMonitoring()`
- **Hourly Monitoring**: Run `setupHourlyMonitoring()`

## 🔧 Configuration

### Error Thresholds
```javascript
const ERROR_THRESHOLDS = {
  MIN_QUALITY_SCORE: 3,
  LOW_SPEND_VS_BUDGET_THRESHOLD: 0.25, // 25%
  BUDGET_UTILIZATION_WARNING: 0.8, // 80%
  BUDGET_UTILIZATION_CRITICAL: 0.95 // 95%
};
```

### Email Configuration
```javascript
const EMAIL_RECIPIENTS = ["your-email@domain.com"];
const EMAIL_SUBJECT_PREFIX = "[Google Ads Critical Alert]";
```

## 🧪 Testing

### Test Functions
- `testScript()`: Run complete validation
- `testEmailConfiguration()`: Test email delivery
- `checkScriptPermissions()`: Verify account access
- `createTestErrors()`: Create test error scenarios

### Validation
- `runCompleteValidation()`: Comprehensive script validation
- `monitorExecutionTime()`: Performance monitoring
- `checkSpecificErrorType()`: Test specific error detection

## 📊 Error Types Monitored

| Error Type | Severity | Description |
|------------|----------|-------------|
| ACCOUNT_SUSPENDED | CRITICAL | Account is suspended |
| BILLING_SUSPENDED | CRITICAL | Billing is suspended |
| SCRIPT_ERROR | CRITICAL | Script failures in MCC |
| AD_DISAPPROVED | HIGH | Ads are disapproved |
| KEYWORD_DISAPPROVED | MEDIUM | Keywords are disapproved |
| LOW_QUALITY_SCORE | MEDIUM | Keywords with low quality scores |
| CAMPAIGN_NO_CONVERSIONS | HIGH | Campaigns with no conversions |
| NO_SPEND_24H | HIGH | No spending in 24 hours |
| LOW_SPEND_VS_BUDGET | LOW | Low budget utilization |
| INVALID_LANDING_PAGE | HIGH | Invalid landing page URLs |
| BROAD_GEO_TARGETING | LOW | Overly broad geographic targeting |
| NO_CONVERSION_TRACKING | MEDIUM | Missing conversion tracking |
| MISSING_MERCHANT_CENTER | HIGH | Shopping campaigns not linked |

## 🔄 Auto-Save Workflow

This repository includes GitHub Actions for automatic version control:

### Automatic Commits
- **Daily Backups**: Automatic commits every day at 2 AM
- **Change Detection**: Commits only when changes are detected
- **Version History**: Maintains complete version history

### GitHub Actions
- **Auto-Save**: `/.github/workflows/auto-save.yml`
- **Backup Schedule**: Daily automated backups
- **Change Tracking**: Monitors file modifications

## 📝 Usage

### Manual Execution
```javascript
// Run the main error monitor
main();

// Test the script
testScript();

// Setup scheduling
setupDailyMonitoring();
```

### Automated Execution
The script can be scheduled to run:
- **Daily**: For comprehensive monitoring
- **Hourly**: For critical error detection
- **Custom**: Based on your needs

## 🚨 Troubleshooting

### Common Issues
1. **Permission Errors**: Check account access and API permissions
2. **Email Not Sending**: Verify email configuration and quotas
3. **Conversion Data Access**: Check account-level restrictions
4. **Script Failures**: Review execution logs and error messages

### Debug Mode
Enable debug mode for detailed logging:
```javascript
const DEBUG_MODE = true;
```

## 📈 Performance

### Optimization
- **Efficient Queries**: Uses optimized GAQL queries
- **Batch Processing**: Processes accounts in batches
- **Error Handling**: Comprehensive error handling and recovery
- **Execution Time Monitoring**: Tracks script performance

### Monitoring
- **Execution Logs**: Detailed logging for troubleshooting
- **Performance Metrics**: Execution time tracking
- **Error Reporting**: Comprehensive error categorization

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

For issues and questions:
1. Check the troubleshooting section
2. Review the test functions
3. Enable debug mode for detailed logs
4. Create an issue in this repository

## 🔄 Version History

- **v1.0**: Initial release with comprehensive error monitoring
- **v1.1**: Added MCC script monitoring and ad schedule awareness
- **v1.2**: Enhanced conversion tracking with fallback methods
- **v1.3**: Added auto-save functionality and GitHub Actions

---

**Last Updated**: January 2025
**Maintainer**: Peterson Rainey
