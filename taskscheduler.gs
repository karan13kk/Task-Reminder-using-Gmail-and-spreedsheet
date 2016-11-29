var EMAIL_SENT = "EMAIL_SENT";
var REMINDER_DONE = "REMINDER_DONE";
function sendEmails2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = sheet.getLastRow()-1;   // Number of rows to process
  // Fetch the range of cells A2:B7
  var dataRange = sheet.getRange(startRow, 1, numRows, 7)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var currentDate;     //current date;
  var sendDate;        //send date
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0];  // First column
    var message = row[1];       // Second column
    var emailSent = row[2];     // Third column
    var date1 = row[3];
    var reminderDone = row[6];
    var remainingTime;
    var subject = "Sending emails from a Spreadsheet";
    var date = new Date();
    // check schedule for sending mail
    var sendingDate = sheet.getRange(startRow + i, 4).getValue();
    var sendingDateValue = new Date(sendingDate);
    // track of remaining days
    var finalDate = sheet.getRange(startRow + i, 5).getValue();
    var date1value = new Date(finalDate);
    var date2 = new Date();
     // sending mails 
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
      if(date2.getTime()>sendingDateValue.getTime())
      {
      Sdate=Utilities.formatDate(date,'GMT+0530',"MM/dd/yyyy");
      sendDate = Sdate;
      sheet.getRange(startRow + i, 4).setValue(sendDate);
      MailApp.sendEmail(emailAddress, subject, message);
      sheet.getRange(startRow + i, 3).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
      }
      }
    
    //check for the reminder
    if(date2.getTime()<date1value.getTime()){
    var timeDiff = Math.abs(date1value.getTime() - date2.getTime());
    var diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24)); 
    // display remaining days
    sheet.getRange(startRow + i, 6).setValue(diffDays);
    sheet.getRange(startRow + i, 7).setValue(" ");
    }
    
    
    //check for the reminder
    else{ 
      if(reminderDone != REMINDER_DONE)  //will not send multiple reminder
      {
       var subject = "REMINDER!!!";
       sheet.getRange(startRow + i, 6).setValue("-1")
       MailApp.sendEmail(emailAddress, subject, message);
       sheet.getRange(startRow + i, 7).setValue(REMINDER_DONE);
       // Make sure the cell is updated right away in case the script is interrupted
       SpreadsheetApp.flush();
      }
    }
   }
} 