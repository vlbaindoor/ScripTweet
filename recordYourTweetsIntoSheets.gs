// READ THIS Before TESTING https://developers.google.com/apps-script/migration/oauth-config
// Function to pull in your tweets from Twitter and puts them in a spreadsheet

var fields = {'in_reply_to_screen_name':true,'created_at':true,'text':true};

function saveYourTweets_() {
  // Setup OAuthServiceConfig
  var oAuthConfig = UrlFetchApp.addOAuthService("twitter");
  oAuthConfig.setAccessTokenUrl("https://api.twitter.com/oauth/access_token");
  oAuthConfig.setRequestTokenUrl("https://api.twitter.com/oauth/request_token");
  oAuthConfig.setAuthorizationUrl("https://api.twitter.com/oauth/authorize");
  
  oAuthConfig.setConsumerKey(ScriptProperties.getProperty(CONSUMER_KEY));
  oAuthConfig.setConsumerSecret(ScriptProperties.getProperty(CONSUMER_SECRET));
  
  // Setup optional parameters to point request at OAuthConfigService.  The "twitter"
  // value matches the argument to "addOAuthService" above.
  var options =
      {
        "oAuthServiceName" : "twitter",
        "oAuthUseToken" : "always"
      };
  
  var result = UrlFetchApp.fetch("https://api.twitter.com/1.1/statuses/user_timeline.json",
                                 options);
  var o  = Utilities.jsonParse(result.getContentText());
  
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(TweetedMessagesSheet);
  
  // var doc = SpreadsheetApp.getActiveSpreadsheet();
  
  var cell = sheet.getRange('a1');
  var index = 0;
  for (var i in o) {
    var row = o[i];
    var col = 0;
    for (var j in row) {
      if (fields[j]) {
        cell.offset(index, col).setValue(row[j]);
        col++;
      }
    }
    index++;
  }
}

function pullYourTweets() {
  var service = getTwitterService_();
  if (service.hasAccess()) {
    var url = 'https://api.twitter.com/1.1/statuses/user_timeline.json';
    var response = service.fetch(url);
    var tweets = JSON.parse(response.getContentText());
    recordInSheet_(tweets);
    for (var i = 0; i < tweets.length; i++) {
      Logger.log(tweets[i].text);
    }
  } else {
    var authorizationUrl = service.authorize();
    Logger.log('Please visit the following URL and then re-run the script: ' + authorizationUrl);
  }
}

function TOTESTgetTwitterService_() {
  var projectKey = getScriptProjectKey();
  var service = OAuth1.createService('twitter');
  service.setAccessTokenUrl('https://api.twitter.com/oauth/access_token')
  service.setRequestTokenUrl('https://api.twitter.com/oauth/request_token')
  service.setAuthorizationUrl('https://api.twitter.com/oauth/authorize')
  service.setConsumerKey(CONSUMER_KEY);
  service.setConsumerSecret(CONSUMER_SECRET);
  service.setProjectKey(projectKey);
  service.setCallbackFunction('TOTESTauthCallback_');
  service.setPropertyStore(PropertiesService.getScriptProperties());
  return service;
}

function TOTESTauthCallback_(request) {
  var service = getTwitterService();
  var isAuthorized = service.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this page.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this page');
  }
}

function recordInSheet_(tweetMsg) {  
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(TweetedMessagesSheet); 
  var lastRow = sheet.getLastRow();
  var cell = sheet.getRange(lastRow + 1, 1);
  var index = 0;
  for (var i = 0; i < tweetMsg.length; i++) {
    var row = tweetMsg[i];
    var col = 0;
    for (var j in row) {
      if (fields[j]) {
        cell.offset(index, col).setValue(row[j]);
        col++;
      }
    }
    index++;
  }
}
