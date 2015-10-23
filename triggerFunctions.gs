/**
 * *********************** FunctionsToTrigger.gs
 * ***************************************
 */
function clearOurTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getEventType() == ScriptApp.EventType.CLOCK)
      ScriptApp.deleteTrigger(triggers[i]);
  }
  TRIGGER_MODIFIED = false;
  updateTriggerStatusDisplay_();
}

function checkTrigger_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var triggers = ScriptApp.getUserTriggers(ss);

    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getEventType() == ScriptApp.EventType.CLOCK) {
        Logger.log("We have a Clock based trigger");
        return true;
      }
    }
    return false;
  } catch (f) {
    setStatusInfoForUser_('ERROR', f.toString(), 'ERROR');
    return false;
   }
}

function setOurTrigger() {
  clearOurTrigger();
  var settings = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName(SettingsSheet);
  var interval = settings.getRange(TWEET_INTERVAL_CELL_INDEX)
                         .getValue();
  switch (interval) {
    case "12 hours":
      ScriptApp.newTrigger(functionNameForTrigger)
        .timeBased().everyHours(12)
        .inTimezone(localTimeZone)
        .create();
      break;
    case "8 hours":
      ScriptApp.newTrigger(functionNameForTrigger)
        .timeBased().everyHours(8)
        .inTimezone(localTimeZone)
        .create();
      break;
    case "6 hours":
      ScriptApp.newTrigger(functionNameForTrigger)
        .timeBased().everyHours(6)
        .inTimezone(localTimeZone)
        .create();
      break;
    case "4 hours":
      ScriptApp.newTrigger(functionNameForTrigger)
        .timeBased().everyHours(4)
        .inTimezone(localTimeZone)
        .create();
      break;
    case "2 hours":
      ScriptApp.newTrigger(functionNameForTrigger)
        .timeBased().everyHours(2)
        .inTimezone(localTimeZone)
        .create();
      break;
    case "1 hour":
      ScriptApp.newTrigger(functionNameForTrigger)
        .timeBased().everyHours(1)
        .inTimezone(localTimeZone)
        .create();
      break;
    case "30 minutes":
      ScriptApp.newTrigger(functionNameForTrigger)
        .timeBased().everyMinutes(30)
        .inTimezone(localTimeZone)
        .create();
      break;
    case "15 minutes":
      ScriptApp.newTrigger(functionNameForTrigger)
        .timeBased().everyMinutes(15)
        .inTimezone(localTimeZone)
        .create();
      break;
    case "10 minutes":
      ScriptApp.newTrigger(functionNameForTrigger)
        .timeBased().everyMinutes(10)
        .inTimezone(localTimeZone)
        .create();
      break;
    case "5 minutes":
      ScriptApp.newTrigger(functionNameForTrigger)
        .timeBased()
        .inTimezone(localTimeZone)
        .everyMinutes(5).create();
      break;
    case "2 minutes":
      ScriptApp.newTrigger(functionNameForTrigger)
        .timeBased()
        .inTimezone(localTimeZone)
        .everyMinutes(2).create();
      break;
    case "1 minutes":
      ScriptApp.newTrigger(functionNameForTrigger)
        .timeBased()
        .inTimezone(localTimeZone)
        .everyMinutes(1).create();
      break;
      
    default:
      Logger.log("Error: Invalid interval.");
  }
  // we have set the trigger so we don't want to set it again
  TRIGGER_MODIFIED = false;
  
  updateTriggerStatusDisplay_();
  SpreadsheetApp.flush();
  Logger.log("Trigger set: " + interval);
}
