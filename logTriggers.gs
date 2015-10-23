/**
 * Function to check user triggers and log them using
 * Logger
 */
function logTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(ss);
  Logger.log('Number of User Triggers: ' + triggers.length);
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    Logger.log('Trigger EventType= ' + trigger.getEventType());
    Logger.log('Trigger Handler= ' + trigger.getHandlerFunction());    
  }
  
  triggers = ScriptApp.getProjectTriggers();
  Logger.log('Number of Project Triggers: ' + triggers.length);
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    Logger.log('Trigger EventType= ' + trigger.getEventType());
    Logger.log('Trigger Handler= ' + trigger.getHandlerFunction());    
  }
} 
