/*********************************************************************************
 *  S c r i p T w e e t 2.5
 *  - - - - - - - - - - -------
 *  Written by Vivekananda Baindoor Rao http://www.wake-technology.com
 *
 *  This code is inspired by multiple sources but mainly by:
 *      "Send Personalized Tweets & DMs in Bulk from a Google Spreadsheet"
 *    Which is a work by Amit Agarval
 *          http://www.labnol.org/internet/send-personalized-tweets/28880/
 *    and
 *  An Example written by Kyle Finley, 2015 Twitter: @KFinley
 *      http://kylefinley.net and details: http://goo.gl/xrPziQ
 *
 *********************************************************************************
 */
/******************** NOTE ***************************************************
 *
 * All Global Constants and Variables are defined/declared in
 * the GlobalConstantsVariables.gs file
 *  
 * The Google URL Shortener API is to be enabled by going into Resources and
 * Advanced Google Services, scroll down, find Google URL Shortener API and
 * Enable it. You may also need to go into Google Developers Console and enable
 * them.
 * 
 * Under Resouces -> Libraries -> Find a Library copy paste the following string
 * to search Mb2Vpd5nfD3Pz-_a-39Q4VfxhMjh3Sh48
 * 
 * It should find OAuth1 library - this needs to be added to resources used by
 * the script. You may have to select a specific version of the libarary after
 * you you done the previous step.
 *
 * ****************** NOTE*****************************************************
 */

/**
 * Function to clear Status column of spreadsheet so that tweets
 * can be tweeted again. The tweeting function would check
 * the Status column and skip all those rows which have "SENT"
 * for status value. This function simply clears that value.
 *
 * @param  null
 *
 * @returns null
 */
function clearStatus_() {
  try {
    var start = Browser.msgBox("CONFIRMATION",
      'Are you sure you want to clear all Status values? Select YES to confirm.',
                      Browser.Buttons.YES_NO);
    if (start === "no")
      return;
    
    clearStatusWithoutUserConfirmation_();
    
  } catch (f) {
      Browser.msg("ERROR: " + f.toString());
      setStatusInfoForUser_('ERROR', f.toString(), 'ERROR');
      return;
  }
  setStatusInfoForUser_('READY',
          'You can send all the tweets again if you want', 'OK');
}

/**
 * Function to clear formatted Tweet messages from Tweet column of spreadsheet.
 * Once this function is run, the user will have to re-prepare the Tweets
 *
 * @param  null
 *
 * @returns null
 */
function clearTweets_() {
  try {
    var start = Browser.msgBox("CONFIRMATION",
      'Are you sure you want to clear all Tweet values? Select YES to confirm.',
                    Browser.Buttons.YES_NO);
    if (start === "no")
      return;
    setStatusInfoForUser_('WORKING...', 'Busy working. Please be patient',
                          'WARN');
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(MessagesForTweetingSheet);
    var data = sheet.getDataRange().getValues();
    var headers = data[1];
    var colTweet = getColumnIndex_(sheet, headers, "Tweet");
    for (var i = 2; i < data.length; i++) {
      ss.toast("Clearing Tweet in row #" + (i + 1));
      sheet.getRange(i + 1, colTweet + 1).clear();
      sheet.getRange(i + 1, colTweet + 1).setWrap(true);
    }
  } catch (f) {
    Browser.msg("ERROR: " + f.toString());
    setStatusInfoForUser_('ERROR', f.toString(), 'ERROR');
    return;
  }
  setStatusInfoForUser_('CAREFUL', 
      'You cleared all prepared Tweets. You need to re-prepare Tweets',
                        'WARN');
}

/**
 * This is the function which is going to be called by the periodic Trigger
 * This function simply goes through the spreadsheet, row by row and send
 * Tweets again. It would first clear the Status column of the spreadsheet
 * so that the previous status is cleared away
 *
 * @param  null
 *
 * @returns null
 */
function tweetAgain() {
	clearStatusWithoutUserConfirmation_();
	sendTweets();
}

/**
 * This function simply clears the Status column without asking for
 * confirmation from the user. The assumption is that user confirmation
 * is asked prior to calling this function or that this is called from
 * the trigger driven function in which case the user may not be present.
 *
 * @param  null
 *
 * @returns null
 */
function clearStatusWithoutUserConfirmation_() {
  try {
    setStatusInfoForUser_('WORKING...', 
                          'Busy working. Please be patient', 
                          'WARN');
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(MessagesForTweetingSheet);
    var data = sheet.getDataRange().getValues();
    var headers = data[1];
    var colStatus = getColumnIndex_(sheet, headers, "Status");
    for (var i = 2; i < data.length; i++) {
      ss.toast("Clearing Status in row #" + (i + 1));
      sheet.getRange(i + 1, colStatus + 1).clear();
      sheet.getRange(i + 1, colStatus + 1).setWrap(true);
    }
  } catch (f) {
      Browser.msg("ERROR: " + f.toString());
      setStatusInfoForUser_('ERROR', f.toString(), 'ERROR');
      return;
  }
  setStatusInfoForUser_('OK',
    'You cleared all old status. You can send the tweets again',
                        'OK');
}

/**
 * This function checks the length of content
 *
 * @param {String} content  Message to be tweeted.
 *
 * @returns a string or a null string.
 *          But all returns are constants defined in GlobalContantsVariables.gs
 */
function checkPostLength(content) {
  var contentLength = content.length;
  if (contentLength == 0)
    return '';
  if (contentLength > 140)
    return POST_TOO_LONG;
  if (contentLength > 120)
    return POST_HARD_TO_RETWEET;
  if (contentLength < 121)
    return POST_GREAT_TO_RETWEET;
}

/**
 * This function simply calls the ScriptApp getProjectKey function 
 * and returns it
 * This function is called mainly by user interface HTML files within
 * their <script> sections and also used to set the value in the Settings
 * Sheet
 *
 * @param  null
 *
 * @returns {String} value of Project Key of Script associated with sheet
 */
function getScriptProjectKey() {
  return ScriptApp.getProjectKey();
}

/**
 * This function is used to prompt the user to Authorise this Script
 * by checking for a Script property. This happens when you either
 * make a copy of the ScripTweet and open it for the first time.
 *
 * @param  null
 *
 * @returns {Boolean} false when user refuses to authorise the script
 *                    true when the scriptAuthorised property is set to 'yes'
 */
function checkScriptAuthorisation_() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var ans = scriptProperties.getProperty('scriptAuthorised');
  if ((!ans) || (ans !== 'yes')) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var authoriseMenu = [
          { name : "Authorise this app",
            functionName : "authorise_"},
          null,
          { name : "About",
            functionName : "about_" },
          { name : "Support & Customization",
            functionName : "help_" }  
                        ];
    ss.addMenu("Authorise", authoriseMenu);
    var html = HtmlService.createTemplateFromFile('PromptForAuthorisation').evaluate()
			       .setWidth(1000).setHeight(540)
                   .setTitle("Please Authorise")
                   .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    ss.show(html);
    return false;
  }
  return true;
}


/**
 * This function is used to remove the Script Authorisation property.
 * This is useful to force re-authorisation which is required 
 * if you copy this script or you are about to share script with someone
 * by making a copy of the ScripTweet
 *
 * @param  null
 *
 * @returns null
 */
function removeScriptAuthorisationProperty() {

  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('scriptAuthorised');
}

/**
 * This function simply forces the Google's own 'authorise' prompt 
 * to appear and it stores a dummy property
 */
function authorise_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('scriptAuthorised', 'yes');
  
  var html = HtmlService.createTemplateFromFile('ScriptAuthorisationSuccess').evaluate()
			       .setWidth(1000).setHeight(540)
                   .setTitle("Confirmation")
                   .setSandboxMode(HtmlService.SandboxMode.IFRAME); 
  ss.show(html);
}

/*
 * The onOpen function is executed automatically every time a Spreadsheet is
 * loaded function onOpen() { var ss = SpreadsheetApp.getActiveSpreadsheet();
 * var menuEntries = []; // When the user clicks on "addMenuExample" then "Menu
 * Entry 1", the function function1 is // executed. menuEntries.push({name:
 * "Menu Entry 1", functionName: "function1"}); menuEntries.push(null); // line
 * separator menuEntries.push({name: "Menu Entry 2", functionName:
 * "function2"});
 * 
 * ss.addMenu("addMenuExample", menuEntries); }
 */
function onEdit(e) {
  setTweetNeedRePrepare_();
  var range = e.range;
  var cellNotation = range.getA1Notation();
  if (cellNotation == TWEET_INTERVAL_CELL_INDEX) {
    // We need to reset the trigger and force user to set it again
    TRIGGER_MODIFIED = true;
    var html = HtmlService.createTemplateFromFile('WarnUserToResetTrigger').evaluate()
                .setTitle("CAUTION")
                .setWidth(1000).setHeight(540)
                .setSandboxMode(HtmlService.SandboxMode.IFRAME); 
    var ss = SpreadsheetApp.getActive();
     ss.show(html);
  }
  
  // Set a comment on the edited cell to indicate when it was changed.
  // var range = e.range;
  // range.setNote('Last modified: ' + new Date());
}

// The Support & Customisation option in the menu calls this function
// We create a HTML Output from the Help.html file which is part of this
// Project and we show the html - using HtmlService function from AppScript
function help_() {
  	var html = HtmlService.createTemplateFromFile('Help').evaluate()
                 .setWidth(1000).setHeight(540).setTitle("Help for ScripTweet")
                 .setSandboxMode(HtmlService.SandboxMode.IFRAME);
	var ss = SpreadsheetApp.getActive();
	ss.show(html);
}

/**
 * Function to show warning message to the user since there has been some cell editing.
 */
function setTweetNeedRePrepare_() {
  TWEET_NEED_RE_PREPARED = true;
  clearRangeFormat();
  setStatusInfoForUser_("WARNING",
          "Need to Re-Prepare Tweets",
                        userAlertWarn);
  if (TRIGGER_MODIFIED == true) {
      setStatusInfoForUser_("WARNING",
          "Need to Modify Triggers",
                        userAlertWarn);
  }
}

/**
 * Function to clear the spreadsheet of values to remove existing data
 */
function clearRangeFormat() {
  var msgsSheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName(MessagesForTweetingSheet);
  var range = msgsSheet.getDataRange();
  var data = range.getValues();
  var lastColumn = msgsSheet.getLastColumn();
  
  for (var i = 2; i < data.length; i++) {
    range = msgsSheet.getRange(i + 1, 1, 1, lastColumn);
    range.setBackground('white').setFontColor('black');
  }
}

/**
 * Function to display the status of the Auto Tweet trigger status
 */
function updateTriggerStatusDisplay_() {
  var settings = SpreadsheetApp.getActiveSpreadsheet()
                        .getSheetByName(SettingsSheet);
  var triggersSet = checkTrigger_();
  var statusCell = settings.getRange(TRIGGER_STATUS_CELL_INDEX);
  if (triggersSet) {
    statusCell.setValue("SET")
              .setBackground('green')
              .setFontColor('white')      
              .setVerticalAlignment("middle") 
              .setWrap(true);
  } else {
    statusCell.setValue("OFF")
              .setBackground('red')
              .setFontColor('white')
              .setVerticalAlignment("middle")
              .setWrap(true);       
  }
  
  if (TRIGGER_MODIFIED == true) {
    statusCell.setValue("NEED RESET")
              .setBackground('red')
              .setFontColor('white')  
              .setVerticalAlignment("middle")
              .setWrap(true);       
  }
  
  var msgsSheet = SpreadsheetApp.getActiveSpreadsheet()
                       .getSheetByName(MessagesForTweetingSheet);
  var statusInfoCell = msgsSheet.getRange(TRIGGER_STATUS_INFO_CELL); 
  if (triggersSet) {
    statusInfoCell.setValue("Tweeting Periodically")
                  .setBackground('green')
                  .setFontColor('white')  
                  .setVerticalAlignment("middle")
                  .setWrap(true);
  } else {
    statusInfoCell.setValue("NOT Tweeting")
                  .setBackground('red')
                  .setFontColor('white')  
                  .setVerticalAlignment("middle")
                  .setWrap(true);       
  }
  if (TRIGGER_MODIFIED == true) {
    statusInfoCell.setValue("NEED RESET")
                  .setBackground('red')
                  .setFontColor('white')  
                  .setVerticalAlignment("middle")
                  .setWrap(true);       
  }
  
}

/**
 * *********************** TwitterFunctions.gs
 * ***************************************
 */

/**
 * This function creates a Service and sets up 'twitter' as the Callback
 * function.
 * @returns {Object} created service using the OAuth1 library
 */
function getTwitterService_() {
  getSettingsFromSheet_();
  return OAuth1.createService('twitter')
    .setAccessTokenUrl('https://api.twitter.com/oauth/access_token')
    .setRequestTokenUrl('https://api.twitter.com/oauth/request_token')
    .setAuthorizationUrl('https://api.twitter.com/oauth/authorize')
    .setConsumerKey(CONSUMER_KEY).setConsumerSecret(CONSUMER_SECRET)
    .setCallbackFunction(functionNameForUserCallback)
    .setPropertyStore(PropertiesService.getUserProperties());
}

/**
 * This function is set up as a call back function for a service
 * set up by getTwitterService_ function. The name of this function
 * is given indirectly via a Global variable functionNameForUserCallback
 * @param   {Object} request from Twitter.com
 * @returns {Object} HTML Service is used to create a Template from a HTML file
 *                   and depending on success or failure of handleCallback request
 *                   appropriate HTML template is chosen
 */
function userCallback(request) {
  var twitterService = getTwitterService_();
  var isAuthorized = twitterService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createTemplateFromFile('ShowTwtrAuthSuccess')
                 .evaluate().setWidth(1000).setHeight(540)
                 .setTitle("Twitter Authorization Status")
                 .setSandboxMode(HtmlService.SandboxMode.IFRAME); 
  } else {
    return HtmlService.createTemplateFromFile('ShowFailureTwtrAuth')
                 .evaluate().setWidth(1000).setHeight(540)
                 .setTitle("Twitter Authorization Failed")
                 .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }
}


/**
 * This function would clear the service set up by the getTwitterService_ function
 * Calling/Running (from debugger) this function would force Twitter re-authorisation requirement
 */
function clearService() {
  var userResponse = Browser.msgBox("CONFIRMATION",
                             'Are you sure you want to clear Twitter Authorisation? You would be required to re-authorise. Select YES to confirm.',
                             Browser.Buttons.YES_NO);
  if (userResponse === "no")
    return;
  
  OAuth1.createService('twitter')
        .setPropertyStore(PropertiesService.getUserProperties())
        .reset();
}

/**
 * This function deletes the stored Script Property oauth.twitter after re-confirming
 * from the user that it is what they want to do
 */
function rewokeTwitterService() {
  var userResponse = Browser.msgBox("CONFIRMATION",
                             'Are you sure you want to rewoke Twitter Authorisation? You would be required to re-authorise. Select YES to confirm.',
                             Browser.Buttons.YES_NO);
  if (userResponse === "no")
    return;

  var twitterService = getTwitterService_();
  twitterService.reset();
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('oauth1.twitter');
}

/**
 * Function to create template from HTML file Config.html and show it to user
 */
function authTwitter_() {
  var html = HtmlService.createTemplateFromFile('Config')
              .evaluate().setWidth(1000).setHeight(540)
              .setTitle("Twitter Authorization")
              .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  var ss = SpreadsheetApp.getActive();
  ss.show(html);
}

/**
 * This function creates a Template from file using the Html Service which loads the
 * HTML content from the 'text.html' file which also executes functions within that
 * html file. Then the validate function within that text.html is called which ends up
 * calling saveTweetTemplate function
 */
function createTweetsUsingTemplate() {
  var html = HtmlService.createTemplateFromFile('PrepareTemplate')
                .evaluate().setWidth(1000).setHeight(540)
                .setTitle("Prepare/Re-Prepare Tweets")
                .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  var ss = SpreadsheetApp.getActive();
  ss.show(html);
}

/**
 * This function is called from the HTML file Config.html
 * It checks the Twitter authorisation status and 
 * @returns {String} to inform the user.
 */
function getTwitterAuthStatus() {
  try {
    var twitterService = getTwitterService_();
    if (!twitterService.hasAccess()) {
      return twitterService.authorize();
    } else {
      setStatusInfoForUser_("OK",
                "Your Account is already authorized to use Twitter.",
                            'NORMAL');
      return 
      "Your Account is already authorized to use Twitter. Please close this window.";
    }
  } catch (f) {
    return "ERROR: " + f.toString();
  }
}



/**
 * This function is called from the HTML file text.html.
 * This calls mergeTemplateWithColumnData_ to actually merge the template and text
 * from columns referred in the template from MessagesForTweetingSheet sheet and
 * actually writes the formed Tweet into the column named Tweet.
 * @param   {Object} params the template
 * @returns {String} for informing the user
 */
function saveTweetTemplate(params) {
  try {
    doProperty_("templateForTweet", params.templateForTweet);
    mergeTemplateWithColumnData_();
    
    setStatusInfoForUser_("Template saved.",
               "Proceed to send tweets if all Post Length Status are Green.",
                          'WARN');
    return "Template saved. You can now proceed to send tweets.";
  } catch (f) {
      setStatusInfoForUser_('ERROR', f.toString(), 'ERROR');
      return "ERROR: " + f.toString();
  }
}

/**
 * This function does almost same thing as the saveTweetTemplate except
 * that this function uses a default template since there
 * is none passed in as parameter
 * @returns {String} for informing the user
 */
function useDefaultTweetTemplate() {
  try {
    doProperty_("templateForTweet", defaultTemplate);
    mergeTemplateWithColumnData_();
    
    setStatusInfoForUser_("Template saved.",
               "Proceed to send tweets if all Post Length Status are Green.",
                          'WARN');
    return "Template saved. You can now proceed to send tweets.";
  } catch (f) {
      setStatusInfoForUser_('ERROR', f.toString(), 'ERROR');
      return "ERROR: " + f.toString();
  }
}

/**
 * This function is called from the HTML file text.html
 * @returns {Object} the teplateForTweet from properties
 */
function getTweetTemplate() {
  return {templateForTweet : doProperty_("templateForTweet")};
}

/**
 * Function to Tweet away using default template
 */
function tweetAway() {
  var status;
  var thereWasError = false;
  var colourForStatus = '';
  var colourForText = 'black';
  try {
    setStatusInfoForUser_('WORKING...', 'Busy working. Please be patient',
                          'WARN');
    useDefaultTweetTemplate();
    var twitterService = getTwitterService_();
    if (!twitterService.hasAccess()) {
      setStatusInfoForUser_('ERROR', "Please authorize your Twitter account",
                            'ERROR');
      Browser.msgBox("Please authorize your Twitter account");
      return;
    }
    
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(MessagesForTweetingSheet);
    var data = sheet.getDataRange().getValues();
    var lastColumn = sheet.getLastColumn();
    var headers = data[1];
    
    var colMessage = getColumnIndex_(sheet, headers, "Message Text");
    var colTweet = getColumnIndex_(sheet, headers, "Tweet");
    var colStatus = getColumnIndex_(sheet, headers, "Status");
    
    for (var i = 2; i < data.length; i++) {
      if (data[i][colTweet].length < 10) {
        setStatusInfoForUser_('ERROR', 'You did not Prepare/Re-Prepare Tweets',
                              'ERROR');
        return;
      }
      // We shall skip the row if the Status in ColStatus says 'SENT'
      if (data[i][colStatus].toString().trim().toUpperCase() !== "SENT") {
        var api;
        var method;
        // var twitterUser = data[i][colUser].trim().replace(/^\@/, "");
        var tweet = data[i][colTweet].trim();
        // ss.toast("Sending tweet to @" + twitterUser);
        Logger.log("Sending tweet: " + tweet);
        api = "https://api.twitter.com/1.1/statuses/update.json?status="
                  + encodeString_(tweet);
        try {
          Logger.log("POSTing Tweet");
          var response = twitterService.fetch(api, {
                              method : "POST",
                              muteHttpExceptions : true
                            });
          Logger.log("Checking ResponseCode");
          if (response.getResponseCode() === 200) {
            status = "SENT";
            colourForStatus = 'green';
            colourForText = 'white';
          } else {
            status = "ERROR: "
                      + JSON.parse(response.getContentText())
                              .errors[0].message;
            setStatusInfoForUser_('ERROR', status, 'ERROR');
            colourForStatus = 'red';
            colourForText = 'white';
            thereWasError = true;
          }
        } catch (t) {
          status = "ERROR: " + t.toString();
          setStatusInfoForUser_('ERROR: ', status, 'ERROR');
          thereWasError = true;
        }
        Logger.log("Status is:" + status);
        sheet.getRange(i + 1, colStatus + 1).setValue(status);
        // Let us set the colour of the entire row depending on status
        sheet.getRange(i + 1, 1, 1, lastColumn).setBackground(colourForStatus);
        sheet.getRange(i + 1, 1, 1, lastColumn).setFontColor(colourForText);
        Utilities.sleep(2000);
      } else {
        ss.toast("Skipping row #" + (i + 1));
        Logger.log("Skipping row #" + (i + 1));
      }
      
    }
  } catch (f) {
    setStatusInfoForUser_('ERROR', f.toString(), 'ERROR');
    thereWasError = true;
    return;
  }
  // Let us update the Settings page to record the date and time
  if (thereWasError == false)
    updateLastRunUserInfo_();
}
/**
 * Function to send tweet messages. This function goes through the spreadsheet
 * and checks each row to see if it is to be tweeted or not and then tweets it.
 * After tweeting it checks for response from Twitter and records that response
 * in the Status column of the spreadsheet
 */
function sendTweets() {
  var status;
  var thereWasError = false;
  var colourForStatus = '';
  var colourForText = 'black';
  try {
    setStatusInfoForUser_('WORKING...', 'Busy working. Please be patient',
                          'WARN');
    var twitterService = getTwitterService_();
    if (!twitterService.hasAccess()) {
      setStatusInfoForUser_('ERROR', "Please authorize your Twitter account",
                            'ERROR');
      Browser.msgBox("Please authorize your Twitter account");
      return;
    }
    
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(MessagesForTweetingSheet);
    var data = sheet.getDataRange().getValues();
    var lastColumn = sheet.getLastColumn();
    var headers = data[1];
    
    var colMessage = getColumnIndex_(sheet, headers, "Message Text");
    var colTweet = getColumnIndex_(sheet, headers, "Tweet");
    var colStatus = getColumnIndex_(sheet, headers, "Status");
    
    var api;
    var method;
    var params;
    
    for (var i = 2; i < data.length; i++) {
      if (data[i][colTweet].length < 10) {
        prepareOneRowToTweet_(defaultTemplate, sheet, i);
        SpreadsheetApp.flush();
        data = sheet.getDataRange().getValues();
      }
      // We shall skip the row if we don't need to send Tweet
      if (toSendOrNotToSend_(data, i, colStatus)) {
        var tweet = data[i][colTweet].trim();
        Logger.log("Sending tweet: " + tweet);

        // params = handleMedia_(sheet, i, data);
        
        api = "https://api.twitter.com/1.1/statuses/update.json?status="
                  + encodeString_(tweet);
        try {
          Logger.log("POSTing Tweet");
          var response = twitterService.fetch(api, {
                              method : "POST",
                              muteHttpExceptions : true
                            });
          Logger.log("Checking ResponseCode");
          if (response.getResponseCode() === 200) {
            status = "SENT";
            colourForStatus = 'green';
            colourForText = 'white';
          } else {
            status = "ERROR: "
                      + JSON.parse(response.getContentText())
                              .errors[0].message;
            setStatusInfoForUser_('ERROR', status, 'ERROR');
            colourForStatus = 'red';
            colourForText = 'white';
            thereWasError = true;
          }
        } catch (t) {
          status = "ERROR: " + t.toString();
          setStatusInfoForUser_('ERROR: ', status, 'ERROR');
          thereWasError = true;
        }
        Logger.log("Status is:" + status);
        sheet.getRange(i + 1, colStatus + 1).setValue(status);
        // Let us set the colour of the entire row depending on status
        sheet.getRange(i + 1, 1, 1, lastColumn).setBackground(colourForStatus);
        sheet.getRange(i + 1, 1, 1, lastColumn).setFontColor(colourForText);
        Utilities.sleep(2000);
      } else {
        ss.toast("Skipping row #" + (i + 1));
        Logger.log("Skipping row #" + (i + 1));
      }
      
    }
  } catch (f) {
    setStatusInfoForUser_('ERROR', f.toString(), 'ERROR');
    thereWasError = true;
    return;
  }
  // Let us update the Settings page to record the date and time
  if (thereWasError == false)
    updateLastRunUserInfo_();
}

/**
 * Function to check whether a specific row has tweets that need to sent or not
 * Return true if we need to Tweet otherwise return false Incorporate more conditions
 * as we develop ScripTweet further At present there is only one condition
 * @param   {Object} data      spreadsheet range of cells containing possible tweet messages etc
 * @param   {Number} index     row number of the spreadsheet which is being considered
 * @param   {Number} colStatus column index for checking - whether to tweet or not to tweet
 * @returns {Boolean}  false if row is not to be used to send tweet, true if row has message to be tweeted
 */
function toSendOrNotToSend_(data, index, colStatus) {
  // We shall skip the row if the Status in ColStatus says 'SENT'
  if (data[index][colStatus].toString().trim().toUpperCase() === "SENT") {
    return false;
  }
  
  // more conditions to be checked and if not satisfied we simply return false
  // if all conditions are satisfied then we finally return true
  
  return true;
}


/**
 * This function is called to actually merge the template and text from columns referred
 * in the template from MessagesForTweetingSheet sheet and actually writes the formed Tweet
 * into the column named Tweet. The replaceVariables_ is the function which does 
 * the replacement of the template elements with actual data.
 * We also call checkPostLength function to check the length of the resulting string.
 * We also change the colour of the Post Length Status column of the spreadsheet to
 * indicate  green, orange or otherwise status of the tweet text that may get Tweeted.
 */
function mergeTemplateWithColumnData_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(MessagesForTweetingSheet);
  var data = sheet.getDataRange().getValues();
  var headers = data[1];
  var rows = rowsAsObjects_(data, normalizeHeaders_(headers));

  var colMessage = getColumnIndex_(sheet, headers, "Message Text");
  var colTweet = getColumnIndex_(sheet, headers, "Tweet");
  var colStatus = getColumnIndex_(sheet, headers, "Status");
  var colLength = getColumnIndex_(sheet, headers, "Length");

  var template = doProperty_("templateForTweet");

  for (var i = 2; i < rows.length; i++) {
    prepareOneRowToTweet_(template, sheet, i);
  }
  TWEET_NEED_RE_PREPARED = false;
}

/**
 * Function to prepare the tweet message from one row. This function is
 * mainly called from mergeTemplateWithColumnData function but can be called
 * from elsewhere
 * @param {Object} template this should either be a user defined template
 *                            or a default template
 * @param {Object} sheet    spreadsheet
 * @param {Number} i        Row index to be processed
 */
function prepareOneRowToTweet_(template, sheet, i) {
  var data = sheet.getDataRange().getValues();
  var headers = data[1];
  var rows = rowsAsObjects_(data, normalizeHeaders_(headers));
  var colTweet = getColumnIndex_(sheet, headers, "Tweet");
  var colLength = getColumnIndex_(sheet, headers, "Length");
  var colPostLengthStatus = getColumnIndex_(sheet, headers,
                                            "Post Length Status");
  
  var tweet = replaceVariables_(template, rows[i]);
  var lengthCheck = checkPostLength(tweet);
  var postLengthColour = getPostLengthColour(lengthCheck);
  sheet.getRange(i + 1, colTweet + 1).setValue(tweet);
  sheet.getRange(i + 1, colLength + 1).setValue(tweet.length);
  
  // Let us validate what we just merged as Tweet text
  sheet.getRange(i + 1, colPostLengthStatus + 1).setValue(lengthCheck);
  
  // and set the colour of the column appropriately to alert the user
  sheet.getRange(i + 1, colPostLengthStatus + 1)
       .setBackground(postLengthColour);
}

/**
 * *********************** SettingsFunctions.gs
 * ***************************************
 */

/**
 * Function to get settings from the spreadsheet and assign them 
 * to Global Constants
 */
function getSettingsFromSheet_() {
  var settings = SpreadsheetApp.getActiveSpreadsheet()
                    .getSheetByName(SettingsSheet);
  TWITTER_APP_NAME = settings.getRange(TWITTER_APP_NAME_CELL_INDEX).getValue();
  CONSUMER_KEY = settings.getRange(CONSUMER_KEY_CELL_INDEX).getValue();
  CONSUMER_SECRET = settings.getRange(CONSUMER_SECRET_CELL_INDEX).getValue();
  ACCESS_TOKEN = settings.getRange(ACCESS_TOKEN_CELL_INDEX).getValue();
  ACCESS_TOKEN_SECRET = settings.getRange(ACCESS_TOKEN_SECRET_CELL_INDEX)
                        .getValue();
}

/**
 * Function to set the settings into the spreadsheet
 * @param {String} twitterAppName    this is for user's own information - the name they gave to 
 *                                   Twitter App that they created on Twitter.com servers
 * @param {String} ConsumerKey       Consumer key that is generated for the Twitter App
 * @param {String} ConsumerSecret    Consumer Secret that is generated for the Twitter App
 * @param {String} AccessToken       Access Token that is generated for the Twitter App
 * @param {String} AccessTokenSecret Access Token Secret that is generated for the Twitter App
 */
function setSettingsToSheet_(twitterAppName,
                             ConsumerKey, ConsumerSecret,
                             AccessToken, AccessTokenSecret) {
	var settings = SpreadsheetApp.getActiveSpreadsheet()
                                 .getSheetByName(SettingsSheet);
    settings.getRange(TWITTER_APP_NAME_CELL_INDEX).setValue(twitterAppName);
	settings.getRange(CONSUMER_KEY_CELL_INDEX).setValue(ConsumerKey);
	settings.getRange(CONSUMER_SECRET_CELL_INDEX).setValue(ConsumerSecret);
	settings.getRange(ACCESS_TOKEN_CELL_INDEX).setValue(AccessToken);
	settings.getRange(ACCESS_TOKEN_SECRET_CELL_INDEX).setValue(
			AccessTokenSecret);
}

/**
 * Function to get the Project Key from the built in getScriptProjectKey function
 * and set it as a value for user's information
 */
function setProjectKey_() {
	var settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
			SettingsSheet);
	PROJECT_KEY = settings.getRange(SCRIPT_PROJECT_KEY_CELL_INDEX).setValue(
			getScriptProjectKey());
}

/**
 * if the 'value' parameter is null (not given), it retrieves
 * the present value and returns it
 * @param   {String} key   string name for the property
 * @param   {Object} value value to be assigned to the property
 * @returns {Object} If the value parameter is not given then this function
 *                   returns the current value of that property
 */
function doProperty_(key, value) {
	var properties = PropertiesService.getUserProperties();
	if (value) {
		properties.setProperty(key, value);
	} else {
		return properties.getProperty(key) || "";
	}
}

/** *********************** AddRow.gs *************************************** */

/**
 * Function to insert a row of data - this inserts the row as the first row
 * as against the usual append to the end
 * 
 * @param {Object} sheet    Spreadsheet to insert row into
 * @param {Object}   rowData  an array of parameterised elements that go into specified columns
 *                            the parameter name would indicate the column name
 * @param {[[Type]]} optIndex this is an offset to protect the header rows so that 
 *                            the insertion happens just below the header rows
 *                            The optIndex needs to be 2 to skip the first row in the
 *                            spreadsheet and insert data. In our case it needs to be 3
 *                            as first two rows are header and status info
 */
function insertRow_(sheet, rowData, optIndex) {
  var messageToTweet = rowData.messageToTweet;
  var longURL        = rowData.longURL;
  var shortURL       = rowData.shortURL;
  var imageFileName  = rowData.imageFileName;
  
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var index = optIndex || 1;
    sheet.insertRowBefore(index).getRange(index, 1);

    var data = sheet.getDataRange().getValues();
    var headers = data[1];
    var colMessage       = getColumnIndex_(sheet, headers, "Message Text");
    var colLongURL       = getColumnIndex_(sheet, headers, "Long URL");
    var colShortURL      = getColumnIndex_(sheet, headers, "URL");
    var colImageFileName = getColumnIndex_(sheet, headers, "Image File Name");

    sheet.getRange(index, colMessage + 1).setValue(messageToTweet);
    sheet.getRange(index, colLongURL + 1).setValue(longURL);
    sheet.getRange(index, colShortURL + 1).setValue(shortURL);
    sheet.getRange(index, colImageFileName + 1).setValue(imageFileName);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

/** *********************** shortenURL.gs *************************************** */

/**
 * Function to get short URL given the long URL. This uses the Google URL Shortener
 * Service and hence that API needs to be enabled for this to work.
 * @param   {[[Type]]} longURL the normal URL given by the user
 * @returns {[[Type]]} the shortened URL as provided by the Google URL Shortener service
 */
function shortenURL_(longURL) {
  Logger.log("Going to call Google Shortener for URL: " + longURL);
  var shortUrl = UrlShortener.Url.insert({longUrl : longURL});
  Logger.log("Shortned URL: " + shortUrl.id);
  return shortUrl.id;
}

/**
 * *********************** ScripTweetStatusFunctions.gs
 * ***************************************
 */
/** 

/**
 * Function to get a colour to be used for notifying the user - depending on
 * the length of a post that is meant to be tweeted
 * @param   {String} contentCheckMsg In this package the parameter
 *                                     passed is what is returned by the function checkPostLength
 * @returns {String}   returns name of the colour to be used
 *                     - red, orange, green or white as a string
 */
function getPostLengthColour(contentCheckMsg) {
  switch (contentCheckMsg) {
    case POST_TOO_LONG:
      return 'red';
      break;
    case POST_HARD_TO_RETWEET:
      return 'orange';
      break;
    case POST_GREAT_TO_RETWEET:
      return 'green';
      break;
    default:
      return 'white';
      break;
  }
}

/**
 * Function to update the last run information - this records the date and time
 * as per the timezone and locale settings into two different locations - one in
 * the Message for Tweeting sheet and the other in Settings sheet.
 * Note that the one on Messages for Tweeting sheet gets over written by 
 * other user informational messages
 */
function updateLastRunUserInfo_() {
  var msgsSheet = SpreadsheetApp.getActiveSpreadsheet()
                    .getSheetByName(MessagesForTweetingSheet);
  
  var rangeForMsg = msgsSheet.getRange(USER_INFO_MSG_CELL_INDEX);
  var rangeForStatus = msgsSheet.getRange(USER_INFO_STATUS_CELL_INDEX);
  
  rangeForMsg.setValue("Last Run:")
             .setBackground('green')
             .setFontColor('white')
             .setFontSize(14)
             .setVerticalAlignment("middle")
             .setHorizontalAlignment("left")
             .setWrap(true);
  
  rangeForStatus.setValue(new Date())
                .setNumberFormat(dateFormatString)
                .setBackground('green')
                .setFontColor('white')
                .setFontSize(14)
                .setVerticalAlignment("middle")
                .setHorizontalAlignment("left")
                .setWrap(true);
  
  updateTriggerStatusDisplay_();
  
  var settings = SpreadsheetApp.getActiveSpreadsheet()
                               .getSheetByName(SettingsSheet);
  var rangeForHistoryDate = settings.getRange(HISTORY_DATE_CELL_INDEX);
  
  rangeForHistoryDate.setValue(new Date())
                     .setNumberFormat(dateFormatString)
                     .setBackground('green')
                     .setFontColor('white')
                     .setFontSize(16)
                     .setVerticalAlignment("middle")
                     .setHorizontalAlignment("left")
                     .setWrap(true);
}

/**
 * Function to display information to the user.
 * @param {String} msg1       This is a very short message one word or two
 * @param {String} msg2       This is the message
 * @param {String} statusType This determines the colour of the displayed message
 */
function setStatusInfoForUser_(msg1, msg2, statusType) {
  switch (statusType) {
	case 'ERROR':
		setUserStatusInfo_(msg1, msg2, 'red', 'yellow');
		break;
	case 'WARN':
		setUserStatusInfo_(msg1, msg2, 'orange', 'black');
		break;
	case 'OK':
		setUserStatusInfo_(msg1, msg2, 'green', 'yellow');
		break;
	case 'NORMAL':
		setUserStatusInfo_(msg1, msg2, 'white', 'black');
		break;
	default:
		setUserStatusInfo_(msg1, msg2, 'white', 'black');
		break;
  }
}

/**
 * This is the actual function which sets the content of the spread sheet
 * so that the user can read the status messages
 * @param {String} msg1             This is a very short message one word or two
 * @param {String} msg2             This is the message
 * @param {String} backgroundColour background colour - the cell is set to this 
 *                                    so that the user can notice the change
 * @param {String} fontColour       colour of the text
 */
function setUserStatusInfo_(msg1, msg2, backgroundColour, fontColour) {
	var msgsSheet = SpreadsheetApp.getActiveSpreadsheet()
                    .getSheetByName(MessagesForTweetingSheet);
    var rangeForUserInfo = msgsSheet.getRange(USER_INFO_MSG_CELL_INDEX);
    var rangeForUserStatus = msgsSheet.getRange(USER_INFO_STATUS_CELL_INDEX);

    rangeForUserInfo.clear()
	                .setValue(msg1)
	                .setBackground(backgroundColour)
	                .setFontColor(fontColour)
	                .setFontSize(14)
                    .setVerticalAlignment("middle")
                    .setHorizontalAlignment("left")
                    .setWrap(true);

	rangeForUserStatus.clear()
                      .setValue(msg2)
                      .setNumberFormat(dateFormatString)
	                  .setBackground(backgroundColour)
	                  .setFontColor(fontColour)
	                  .setFontSize(14)
                      .setVerticalAlignment("middle")
                      .setHorizontalAlignment("left")
                      .setWrap(true);
}
/**
 * **************************** EncodeString.gs
 * **************************************
 */

/**
 * Twitter prohibits usage of certain characters in messages and hence this function
 * replaces those characters with what is acceptable to Twitter
 * @param   {String} q the string representing the message required to be tweeted
 *                     but may have characters prohibited by Twitter
 * @returns {String} The string which is encoded and formed as a URI Component. 
 *                   Note that this string may be difficult to read on its own
 *                   but would render correctly when used by the POST function.
 */
function encodeString_(q) {
  var str = q;
  str = str.replace(/!/g, '?');
  str = str.replace(/\*/g, '×');
  str = str.replace(/\(/g, '[');
  str = str.replace(/\)/g, ']');
  str = str.replace(/'/g, '’');
  return encodeURIComponent(str);
}

/**
 * ******************************************* ColumnsRowsUtils.gs
 * ********************************
 */

/**
 * Function to get the column index given a sheet, column headers and
 * for the specified name of the column header.
 * For example "Message To Tweet" or "Long URL" - this function returns the column
 * index for the column with that name
 * Note that if there is no column with that name,
 * then a new column gets added and given that name
 * @param   {Object} sheet   Sheet object to be operated on
 * @param   {Object} headers Header columns as object
 * @param   {String} name    Name of the column to be found / created
 * @returns {Number} Column index of the column found or created
 */
function getColumnIndex_(sheet, headers, name) {
  var col = headers.indexOf(name);
  if (col === -1) {
    sheet.insertColumnAfter(sheet.getLastColumn());
    col = sheet.getLastColumn();
    sheet.getRange(2, col + 1).setValue(name);
  }
  return col;
}

/**
 * Function to convert specified rows into objects and retrn the object
 * @param   {Object} data [[Description]]
 * @param   {Object} keys [[Description]]
 * @returns {Object} [[Description]]
 */
function rowsAsObjects_(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty_(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

/**
 * For comparison purposes it is necessary to normalise the headers
 * so that the case of the text would not throw off the comparison
 * This function goes through a bunch of headers as name indicates
 * @param   {Object} headers [[Description]]
 * @returns {Object} [[Description]]
 */
function normalizeHeaders_(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader_(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

/**
 * For comparison purposes it is necessary to normalise the headers
 * so that the case of the text would not throw off the comparison
 * This function goes through one string
 * @param   {String} headers [[Description]]
 * @returns {String} [[Description]]
 */
function normalizeHeader_(str) {
  return str.replace(/[^\w]+/g, "").toLowerCase();
}

/**
 * [[Description]]
 * @param   {[[Type]]} cellData [[Description]]
 * @returns {[[Type]]} [[Description]]
 */
function isCellEmpty_(cellData) {
  return typeof (cellData) == "string" && cellData == "";
}

/**
 * Function to create HTML Template from About.html file and show it on user's screen
 */
function about_() {
	var html = HtmlService.createTemplateFromFile('About').evaluate()
                 .setWidth(1000).setHeight(540).setTitle("About ScripTweet")
                 .setSandboxMode(HtmlService.SandboxMode.IFRAME);
	var ss = SpreadsheetApp.getActive();
	ss.show(html);
}

/**
 * Function to create a HTML Template from AddNewTweet.html file and show it on user's screen.
 * This function would then go on to add rows into the spreadsheet with messages that need to be
 * Tweeted.
 */
function addNewTweetMessage() {
	var html = HtmlService.createTemplateFromFile('AddNewTweet').evaluate()
			       .setWidth(1000).setHeight(540)
                   .setTitle("Add a new message to Tweet")
                   .setSandboxMode(HtmlService.SandboxMode.IFRAME);
	var ss = SpreadsheetApp.getActive();
	ss.show(html);
}

/**
 * Function to handle user's data entry this gets called from the HTML display when the user
 * wants to add a new Tweet message. This function calls insertRow_ function to actually
 * insert a row with the data that the user keyed in.
 * @param {Object} params From the HTML interface entries by the user are passed down
 *                        to this function
 */
function handleUserDataEntry_(params) {
	try {
		var msgsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
				MessagesForTweetingSheet);
		// enableUserDataEntry_();
		var messageToTweet = params.textMessageToTweet;
		var longURL = params.longURL;
		var shortURL;
        var imageFileName = params.imageFileName;

		Logger.log("Got messageToTweet as: " + messageToTweet);
		if (!longURL)
			longURL = shortURL;

		Logger.log("Got LongURL as: " + longURL);
		if (longURL.length > 20) {
			shortURL = shortenURL_(longURL);
			Logger.log("Got shortURL as: " + shortURL);
		} else {
			shortURL = longURL;
		}
        var rowData = {
            messageToTweet : messageToTweet,
            longURL        : longURL,
            shortURL       : shortURL,
            imageFileName  : imageFileName
          };
		// OFFSET_TO_PROTECT_HEADERS to skip the two header rows
		insertRow_(msgsSheet, rowData, OFFSET_TO_PROTECT_HEADERS);

		setTweetNeedRePrepare_();
	} catch (f) {
		setStatusInfoForUser_('ERROR', f.toString(), 'ERROR');
		return;
	}
}

/**
 * This function is called from the HTML file AddNewTweet.html. The user keyed in
 * data and this function attempts to memorise it.
 * @param   {Object} params Code within HTML file would put together user keyed in data
 *                          and calls this function with that as the parameter
 * @returns {String} a string representing success or failure of the user's action.
 */
function memoriseUserKeyedInData(params) {
	try {
		doProperty_("textMessageToTweet", params.textMessageToTweet);
		doProperty_("longURL",            params.longURL);
        doProperty_("imageFileName",      params.imageFileName);
		handleUserDataEntry_(params);
		return "Data added to Spreadsheet. You can close this window.";
	} catch (f) {
		setStatusInfoForUser_('ERROR', f.toString(), 'ERROR');
		return "ERROR: " + f.toString();
	}
}

//This function is called from the HTML file AddNewTweet.html

/**
 * This function is called from the HTML file AddNewTweet.html. 
 * The stored data that the user keyed in previously is accessed and 
 * displayed for the benefit of user.
 * @returns {Object} Function puts together values as an object it got from storage.
 */
function recollectUserKeyedInData() {
	return {
		textMessageToTweet : doProperty_("textMessageToTweet"),
		longURL            : doProperty_("longURL"),
        imageFileName      : doProperty_("imageFileName")
	};
}

/**
 * Function to remove the Keys and Key Secrets from the spreadsheet
 */
function removeKeysSecrets() {
  var settings = SpreadsheetApp.getActiveSpreadsheet()
                               .getSheetByName(SettingsSheet);
  settings.getRange(TWITTER_APP_NAME_CELL_INDEX).clearContent();
  settings.getRange(CONSUMER_KEY_CELL_INDEX).clearContent();
  settings.getRange(CONSUMER_SECRET_CELL_INDEX).clearContent();
  settings.getRange(ACCESS_TOKEN_CELL_INDEX).clearContent();
  settings.getRange(ACCESS_TOKEN_SECRET_CELL_INDEX).clearContent();
}

/**
 * Log information about the data-validation rule for cell TWEET_INTERVAL_CELL_INDEX
 * useful during testing
 */
function getTweetIntervalCellValidationRules() {
	var settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
			SettingsSheet);
	var cell = settings.getRange(TWEET_INTERVAL_CELL_INDEX);
	var rule = cell.getDataValidation();
	if (rule != null) {
		var criteria = rule.getCriteriaType();
		var args = rule.getCriteriaValues();
		Logger.log('The data-validation rule is %s %s', criteria, args);
	} else {
		Logger.log('The cell does not have a data-validation rule.')
	}
}

/**
 * Function to set data-validation rule for cell TWEET_INTERVAL_CELL_INDEX
 */
function setTweetIntervalCellValidationRules() {
	var settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
			SettingsSheet);
	var cell = settings.getRange(TWEET_INTERVAL_CELL_INDEX);
	var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false)
                   .setHelpText("Select one from the list")
                   .requireValueInList(
					  [ '12 hours', '8 hours', '4 hours', '2 hours',
                        '1 hour', '30 minutes', '15 minutes', '10 minutes',
					    '2 minutes', '1 minute' ],
                      true)
                   .build();
	cell.setDataValidation(rule);
}

/**
 * Function to add data validation rule and set default value for cell TWEET_INTERVAL_CELL_INDEX
 */
function setTweetIntervalCellDefaultValue() {
	setTweetIntervalCellValidationRules();
	var settings = SpreadsheetApp.getActiveSpreadsheet()
                    .getSheetByName(SettingsSheet);
	var cell = settings.getRange(TWEET_INTERVAL_CELL_INDEX)
                       .setValue('30 minutes');
}

/**
 * This function is called to clear out rows from the Spreadsheet to start afresh
 */
function clearOutRows() {
	var sheet = SpreadsheetApp.getActive().getSheetByName(
			MessagesForTweetingSheet);
	var range = sheet.getDataRange();
	var numRows = range.getNumRows();
	Logger.log('Checking and deleting all rows except one');
	if (numRows - OFFSET_TO_PROTECT_HEADERS) {
		Logger.log("Deleting rows between " + OFFSET_TO_PROTECT_HEADERS
				   + " and " + (numRows - OFFSET_TO_PROTECT_HEADERS));
		sheet.deleteRows(OFFSET_TO_PROTECT_HEADERS,
                         numRows - OFFSET_TO_PROTECT_HEADERS);
	}
}

/**
 * This function is to clear out the last run information
 * - typically called from factoryDefault function
 */
function clearOutLastRunUserInfo_() {
	var msgsSheet = SpreadsheetApp.getActiveSpreadsheet()
                      .getSheetByName(MessagesForTweetingSheet);
	msgsSheet.getRange(USER_INFO_MSG_CELL_INDEX)
             .clearContent().clearFormat();

    msgsSheet.getRange(USER_INFO_STATUS_CELL_INDEX)
             .clearContent().clearFormat();

	var settings = SpreadsheetApp.getActiveSpreadsheet()
                      .getSheetByName(SettingsSheet);

	settings.getRange(HISTORY_DATE_CELL_INDEX)
            .clearContent().clearFormat();

}

/**
 * Function to 'Factory Default' everything
 */
function factoryDefault() {
  var userResponse = Browser.msgBox("CONFIRMATION",
                                    'Are you sure you want to RESET to Factory Defaults?. Select YES to confirm.',
                                    Browser.Buttons.YES_NO);
  if (userResponse === "no")
    return;
  
  // Make sure the user is really sure of what they are doing
  var userResponse2 = Browser.msgBox("FINAL CONFIRMATION",
                             'Are you sure REALLY really sure you want to RESET to Factory Defaults? This is your last chance to say NO! Select YES if you are sure.',
                             Browser.Buttons.YES_NO);
  if (userResponse2 !== "yes")
    return;
  
  removeKeysSecrets();
  clearOutLastRunUserInfo_();
  clearOutRows();
  setTweetIntervalCellDefaultValue();
  clearOurTrigger();
  removeScriptAuthorisationProperty();
}


