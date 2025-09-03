function sendEmail() {
  // Email configuration
  const recipientEmail = 'sophiansnow2@gmail.com';
  const subject = 'Pay for Creatives';
  const body = 'Please email me the number of creatives that you made this last month.';
  
  try {
    // Send the email using MailApp service
    MailApp.sendEmail(recipientEmail, subject, body);
    
    // Log success
    Logger.log('Email sent successfully to: ' + recipientEmail);
    Logger.log('Subject: ' + subject);
    Logger.log('Body: ' + body);
    
  } catch (error) {
    // Log any errors
    Logger.log('Error sending email: ' + error.toString());
  }
}

function main() {
  // Call the email function
  sendEmail();
}
