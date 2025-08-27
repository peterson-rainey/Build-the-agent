/**
 * Weekly Email Reminder Script for Sweet Hands Q4 GTM Schedule
 * 
 * This script automatically sends an email every Monday morning
 * with a link to the Q4 GTM schedule spreadsheet.
 * 
 * Setup Instructions:
 * 1. Copy this script into Google Ads Scripts
 * 2. Set up a trigger to run every Monday at 9:00 AM
 * 3. Update the EMAIL_ADDRESS constant with your email
 * 4. Test the script manually first
 */

// Configuration
const EMAIL_ADDRESSES = [
  'peterson@creeksidemarketingpros.com',
  'sophiansnow2@gmail.com'
];
const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1EwOfEeNkw54aBE_B7G0ycjRl85CfHbkPT328CJO6WiE/edit?usp=sharing';
const EMAIL_SUBJECT = 'Check Sweet Hands Q4 GTM Schedule';
const EMAIL_BODY = `
Hi there,

This is your weekly reminder to check the Sweet Hands Q4 GTM Schedule.

Please review the schedule and update your tasks as needed:
${SPREADSHEET_URL}

Key things to check:
- Review upcoming deadlines
- Update task status
- Plan for the week ahead
- Coordinate with team members

Have a great week!

Best regards,
Your Google Ads Script
`;

/**
 * Main function that sends the weekly email reminder
 */
function main() {
  try {
    // Check if today is Monday (0 = Sunday, 1 = Monday, etc.)
    const today = new Date();
    const dayOfWeek = today.getDay();
    
    // Only send email on Monday (day 1)
    if (dayOfWeek === 1) {
      sendWeeklyReminder();
      Logger.log('Weekly reminder email sent successfully on ' + today.toDateString());
    } else {
      Logger.log('Not Monday - skipping email. Today is day ' + dayOfWeek);
    }
    
  } catch (error) {
    Logger.log('Error in main function: ' + error.toString());
  }
}

/**
 * Sends the weekly reminder email
 */
function sendWeeklyReminder() {
  try {
    Logger.log('Attempting to send email to: ' + EMAIL_ADDRESSES.join(', '));
    Logger.log('Email subject: ' + EMAIL_SUBJECT);
    Logger.log('Email body length: ' + EMAIL_BODY.length + ' characters');
    
    // Send the email to all recipients
    MailApp.sendEmail({
      to: EMAIL_ADDRESSES.join(', '),
      subject: EMAIL_SUBJECT,
      body: EMAIL_BODY,
      noReply: true
    });
    
    Logger.log('Weekly reminder email sent successfully to: ' + EMAIL_ADDRESSES.join(', '));
    
  } catch (error) {
    Logger.log('Error sending email: ' + error.toString());
    Logger.log('Error details: ' + JSON.stringify(error));
    throw error;
  }
}

/**
 * Test function to manually trigger the email
 * Use this to test the script before setting up the trigger
 */
function testEmail() {
  try {
    Logger.log('=== Starting test email function ===');
    Logger.log('Current time: ' + new Date().toString());
    
    // Check if MailApp is available
    if (typeof MailApp === 'undefined') {
      Logger.log('ERROR: MailApp is not available');
      return;
    }
    
    sendWeeklyReminder();
    Logger.log('Test email sent successfully');
    Logger.log('=== Test email function completed ===');
    
  } catch (error) {
    Logger.log('Error in test function: ' + error.toString());
    Logger.log('Full error details: ' + JSON.stringify(error));
  }
}

/**
 * Function to check if the script is working correctly
 * This will log the current date and day of week
 */
function checkScriptStatus() {
  const today = new Date();
  const dayOfWeek = today.getDay();
  const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  
  Logger.log('Current date: ' + today.toDateString());
  Logger.log('Day of week: ' + dayOfWeek + ' (' + dayNames[dayOfWeek] + ')');
  Logger.log('Email addresses: ' + EMAIL_ADDRESSES.join(', '));
  Logger.log('Spreadsheet URL: ' + SPREADSHEET_URL);
  
  if (dayOfWeek === 1) {
    Logger.log('Today is Monday - email would be sent');
  } else {
    Logger.log('Today is not Monday - email would be skipped');
  }
}

/**
 * Debug function to test email sending with minimal content
 */
function debugEmail() {
  try {
    Logger.log('=== Debug Email Test ===');
    
    // Test with a simple email first
    const testSubject = 'Test Email from Google Ads Script';
    const testBody = 'This is a test email to verify the script is working.';
    
    Logger.log('Sending test email to: ' + EMAIL_ADDRESSES[0]);
    
    MailApp.sendEmail({
      to: EMAIL_ADDRESSES[0], // Send to first email only for testing
      subject: testSubject,
      body: testBody
    });
    
    Logger.log('Debug test email sent successfully');
    
  } catch (error) {
    Logger.log('Debug email error: ' + error.toString());
  }
}

/**
 * Function to verify Google Apps Script permissions
 */
function checkPermissions() {
  try {
    Logger.log('=== Checking Permissions ===');
    
    // Test if we can access basic Google Apps Script services
    Logger.log('Testing MailApp availability...');
    if (typeof MailApp !== 'undefined') {
      Logger.log('✓ MailApp is available');
    } else {
      Logger.log('✗ MailApp is not available');
    }
    
    Logger.log('Testing Logger availability...');
    if (typeof Logger !== 'undefined') {
      Logger.log('✓ Logger is available');
    } else {
      Logger.log('✗ Logger is not available');
    }
    
    Logger.log('Testing Date object...');
    const testDate = new Date();
    Logger.log('✓ Date object works: ' + testDate.toString());
    
    Logger.log('=== Permission check completed ===');
    
  } catch (error) {
    Logger.log('Permission check error: ' + error.toString());
  }
}
