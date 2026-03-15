function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = e.parameter;
    
    // Save to Sheet
    sheet.appendRow([new Date(), data.name, data.email]);

    // Send Email Notification using GmailApp
    var myEmail = "hello.sumithaldar@gmail.com"; 
    // Adding a timestamp to the subject prevents Gmail from blocking "duplicate" subjects
    var subject = "New Subscriber: " + data.name + " (" + new Date().getTime() + ")";
    var body = "You have a new subscriber!\n\nName: " + data.name + "\nEmail: " + data.email;

    // Use GmailApp instead of MailApp
    GmailApp.sendEmail(myEmail, subject, body);

    // 3. Send Thank You Email to the Subscriber
    var subjectToUser = "Welcome to our Newsletter!";
    var bodyToUser = "Hi " + data.name + ",\n\nThank you for subscribing to our newsletter! We're excited to have you on board.\n\nBest regards,\nThe Team";
    GmailApp.sendEmail(data.email, subjectToUser, bodyToUser);
    
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "success" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "error", "error": error }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
function sendNotificationOnNewRow(e) {
  // This runs whenever the sheet changes
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var rowData = sheet.getRange(lastRow, 1, 1, 3).getValues()[0];

  GmailApp.sendEmail("hello.sumithaldar@gmail.com", "Sheet Updated", 
    "New row added: " + rowData[1] + " (" + rowData[2] + ")");
}
