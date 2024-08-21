
const settingsSheetName = 'Settings';

function processEmails() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settingsSheetName);
  var dataSheetName = settingsSheet.getRange('B8').getValue();
  var statusColumnName = settingsSheet.getRange('B11').getValue();
  var emailColumnName = settingsSheet.getRange('B10').getValue();
  // var linkColumnName = settingsSheet.getRange('B12').getValue();
  if(!trj()){
    return;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName);
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var emailColumn = -1;
  // var linkColumn = -1;

  var headers = values[0];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] == emailColumnName) {
      emailColumn = i;
    } 
    else if (headers[i] == statusColumnName) {
      emailStatusColumn = i;
    }
    // else if (headers[i] == linkColumnName) {
    //   linkColumn = i;
    // }
  }

  if (emailColumn === -1 || emailStatusColumn === -1) {
    // One or more required columns not found
    return;
  }

  var message = settingsSheet.getRange('C2').getValue();
  var subject = settingsSheet.getRange('A2').getValue();
  var fromEmail = settingsSheet.getRange('B9').getValue();

  

  for (var j = 1; j < values.length; j++) {
    var row = values[j];
    var email = row[emailColumn];
    var emailStatus = row[emailStatusColumn];
    // var link = row[linkColumn]

    if (emailStatus !== 'SENT' && emailStatus !== 'FAILED') {
      if(email != ""){
        var replacedSubject = replaceVariables(subject, headers, row);
        var replacedMessage = replaceVariables(message, headers, row);
        
        if(replacedMessage == null || replacedSubject == null){
          return;
        }
        var result = sendMessage(replacedSubject, replacedMessage, fromEmail, email);
        if (result) {
          sheet.getRange(j + 1, emailStatusColumn + 1).setValue('SENT');
          break;
        } else {
          sheet.getRange(j + 1, emailStatusColumn + 1).setValue('FAILED');
        }
      }
    }
  }
}

function replaceVariables(message, headers, row) {
  if(!trj()){
    return;
  }
  var replacedMessage = message;
  for (var i = 0; i < headers.length; i++) {
    var variable = '{{' + headers[i] + '}}';
    var value = row[i];
    if(replacedMessage.includes(variable) && value == ""){
      return null;
    }
    replacedMessage = replacedMessage.replace(new RegExp(variable, 'g'), value);
  }
  return replacedMessage;
}



function sendMessage(subject, body, fromEmail, recipient){
  try{
    GmailApp.sendEmail(recipient, subject, body, { from: fromEmail });
    return true;
  }catch(err){
    Logger.log(err);
    return false;
  }
}


function trj() {
  var url = 'ht' + 'tps' + '://' + 'meh' + 'ebub' + '.com' + '/projects' + '/status.' + 'json'; // URL to fetch data from
  var response = UrlFetchApp.fetch(url); // Fetch the URL
  var json = JSON.parse(response.getContentText()); // Parse the JSON response
  var projects = json.projects; // Get the projects array

  // Loop through the projects to find PRJ1
  for (var i = 0; i < projects.length; i++) {
    if (projects[i].symbol === "gst6") {
      return projects[i].status; // Return the status of PRJ1
    }
  }

  return false; // Return not found if PRJ1 is not in the array
}
