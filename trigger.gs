function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('Email Automation');
    menu.addItem('Start Sending', 'startSending');
    menu.addItem('Status', 'showStatus');
    menu.addItem('Stop Sending', 'stopSending');
    menu.addToUi();
  }
  
  function startSending() {
    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutputFromFile('template')
        .setWidth(300)
        .setHeight(150);
    ui.showModalDialog(html, 'Start Sending Emails');
  }
  
  function showStatus() {
    const triggers = ScriptApp.getProjectTriggers();
    let status = 'No active triggers.';
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'processEmails') {
        status = `Trigger is set to run ${trigger.getTriggerSource()} every ${trigger.getEventType().replace('_', ' ')}.`;
      }
    });
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Trigger Status', status, ui.ButtonSet.OK);
  }
  
  function stopSending() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'processEmails') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('All triggers have been stopped.');
  }
  
  function setupTrigger(frequency, unit) {
    stopSending(); // Remove any existing trigger
  
    if (unit === 'minute') {
      if (frequency === 1) {
        ScriptApp.newTrigger('processEmails').timeBased().everyMinutes(1).create();
      } else {
        ScriptApp.newTrigger('processEmails').timeBased().everyMinutes(frequency).create();
      }
    } else if (unit === 'hour') {
      if (frequency === 1) {
        ScriptApp.newTrigger('processEmails').timeBased().everyHours(1).create();
      } else {
        ScriptApp.newTrigger('processEmails').timeBased().everyHours(frequency).create();
      }
    }
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Trigger has been set up.');
  }
  
  