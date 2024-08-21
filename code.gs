
const settingsSheetName = 'Settings';

function processEmails() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settingsSheetName);

  // Get time window
  var startHour = settingsSheet.getRange('B14').getValue();
  var startMinute = settingsSheet.getRange('C14').getValue();
  var endHour = settingsSheet.getRange('B15').getValue();
  var endMinute = settingsSheet.getRange('C15').getValue();

  var maxEmailsPerDay = settingsSheet.getRange('B17').getValue();
  var emailsSentToday = settingsSheet.getRange('B18').getValue();

  // Get spreadsheet time zone
  var timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var currentTime = new Date();
  var currentHour = currentTime.getHours();
  var currentMinute = currentTime.getMinutes();
  var currentDate = Utilities.formatDate(currentTime, timezone, "yyyy-MM-dd hh:mm:ss");

  // Reset email count if outside of time window
  if (currentHour < startHour || (currentHour === startHour && currentMinute < startMinute) ||
      currentHour > endHour || (currentHour === endHour && currentMinute > endMinute)) {
    settingsSheet.getRange('B18').setValue(0);
    Logger.log('Email count reset because current time is outside the allowed time window.');
    return;
  } else {
    Logger.log('Current time is within the allowed time window.');
  }

  // Check if the email count exceeds the daily limit
  if (emailsSentToday >= maxEmailsPerDay) {
    Logger.log('Email sending stopped because the daily limit has been reached.');
    return;
  } else {
    Logger.log('Daily email limit has not been reached.');
  }

  // Get data sheet and column settings
  var dataSheetName = settingsSheet.getRange('B8').getValue();
  var statusColumnName = settingsSheet.getRange('B11').getValue();
  var emailColumnName = settingsSheet.getRange('B10').getValue();
  var timestampcolName = settingsSheet.getRange('B20').getValue();

  if (!trj()) {
    Logger.log('Function trj() returned false. Stopping the email process.');
    return;
  } else {
    Logger.log('Function trj() returned true. Continuing the email process.');
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName);
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var emailColumn = -1;
  var emailStatusColumn = -1;
  var timeStampColumn = -1;

  var headers = values[0];

  for (var i = 0; i < headers.length; i++) {
    if (headers[i] == emailColumnName) {
      emailColumn = i;
    } else if (headers[i] == statusColumnName) {
      emailStatusColumn = i;
    }
    
    if (timestampcolName) {
      if(headers[i] == timestampcolName){
        timeStampColumn = i;
      }
    }
  }

  if (emailColumn === -1 || emailStatusColumn === -1) {
    Logger.log('One or more required columns not found: emailColumn = ' + emailColumn + ', emailStatusColumn = ' + emailStatusColumn);
    return;
  } else {
    Logger.log('Required columns found: emailColumn = ' + emailColumn + ', emailStatusColumn = ' + emailStatusColumn);
  }

  var message = settingsSheet.getRange('C2').getValue();
  var subject = settingsSheet.getRange('A2').getValue();
  var fromEmail = settingsSheet.getRange('B9').getValue();

  for (var j = 1; j < values.length; j++) {
    var row = values[j];
    var email = row[emailColumn];
    var emailStatus = row[emailStatusColumn];

    if (emailStatus !== 'SENT' && emailStatus !== 'FAILED') {
      if (email != "") {
        var replacedSubject = replaceVariables(subject, headers, row);
        var replacedMessage = replaceVariables(message, headers, row);

        if (replacedMessage == null || replacedSubject == null) {
          Logger.log('Email or subject replacement resulted in null. Stopping the process.');
          return;
        }

        var result = sendMessage(replacedSubject, replacedMessage, fromEmail, email);
        if (result) {
          sheet.getRange(j + 1, emailStatusColumn + 1).setValue('SENT');
          settingsSheet.getRange('B18').setValue(emailsSentToday + 1); // Increment email count
          Logger.log('Email sent successfully to ' + email + '. Email count updated.');
          try{
            if(timestampcolName){
            sheet.getRange(j + 1, timeStampColumn + 1).setValue(currentDate);
            }
          }catch{}
          break;
        } else {
          sheet.getRange(j + 1, emailStatusColumn + 1).setValue('FAILED');
          Logger.log('Failed to send email to ' + email + '.');
          try{
            if(timestampcolName){
            sheet.getRange(j + 1, timeStampColumn + 1).setValue(currentDate);
            }
          }catch{}
        }
      } else {
        Logger.log('Empty email address found. Skipping row.');
      }
    } else {
      Logger.log('Email already processed (status: ' + emailStatus + '). Skipping row.');
    }
  }
}
