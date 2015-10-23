/**
 * When spreadsheet is opened this function adds menu items
 */
function onOpen() {
  getSettingsFromSheet_();
  var menuTitle;
  var menu;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var msgsSheet = ss.getSheetByName(MessagesForTweetingSheet);

  // We have a TimeZoneLocale sheet, let us hide it as it is not required
  // for the viewer to see it - it would only confuse them.
  var timeZoneLocaleSheet = ss.getSheetByName(TimeZoneLocaleSheet);
  timeZoneLocaleSheet.hideSheet();
  
  ss.setActiveSheet(msgsSheet);
  
  if (checkScriptAuthorisation_() == false) {
    // yet to be authorised so simply return
    return;
  }
  
  menu = [{ name : "1. Add a new Tweet message",
            functionName : "addNewTweetMessage"},
          { name : "2. Prepare/Re-Prepare Tweets",
            functionName : "createTweetsUsingTemplate"},
          { name : "3. Send Tweets",
            functionName : "sendTweets" },
          { name : "4. Tweet Away using default template",
            functionName : "tweetAway" },
          { name : "5. Tweet periodically",
            functionName : "setOurTrigger" },
          null,
          { name : "6. Stop periodic Tweeting",
            functionName : "clearOurTrigger"},
          null,
          { name : "Revoke Twitter Authority",
            functionName : "rewokeTwitterService"	},
          { name : "Erase Stored Keys and Secrets",
            functionName : "removeKeysSecrets" },
          { name : "Reset to Factory Default",
            functionName : "factoryDefault" },
          null,
          { name : "Enable Engineering Menu",
            functionName : "enableEngineering_" },
          { name : "Enable Pro Menu",
            functionName : "proMenu_"	},
          null,
          { name : "About",
            functionName : "about_" },
          { name : "Support & Customization",
            functionName : "help_" }        
         ];
  
  menuTitle = getMenuTitle_();
  ss.addMenu(menuTitle, menu);
  
  // hideOneColumn_("Tweet");
  
  setStatusInfoForUser_('Welcome to ScripTweet',
  'ScripTweet recommends you use Blue buttons at Top Left corner or ScriptTweet Menu above.',
                        'NORMAL');
  setProjectKey_();
  
  var params = getTimeZoneLocale();
  var useLocale = params.presentLocale;
  var useTimeZone = params.presentTimeZone;

  localTimeZone = useTimeZone;
  
  SpreadsheetApp.getActiveSpreadsheet()
                .getSheetByName(SettingsSheet)
                .getRange(SCRIPT_TIMEZONE_CELL_INDEX)
                .setValue(localTimeZone);
  
  if (CONSUMER_KEY === '') {
    enableEngineering_();
  }
  
}

/**
 * Function to get the menu title based on the version number,
 * release number etc
 * @returns {[[Type]]} [[Description]]
 */
function getMenuTitle_() {
  var versionNumber = getVersionNumber_();
  var versionRelease = getReleaseNumber_();
  var menuTitle;
  // Let us convert the numbers into strings
  versionNumber = Utilities.formatString( '%.1f', versionNumber);
  versionRelease = Utilities.formatString( '%.1f', versionRelease);
  menuTitle = "ScripTweet " + versionNumber + "." + versionRelease;
  return menuTitle;
}

/**
 * Function to set up the Engineering menu consisting
 * of set up options
 */
function enableEngineering_() {
  var engineering;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  engineering = [
    { name : "[Engr] Help me set up",
      functionName : "getKeySecretsFromUser_" },
    { name : "[Engr] Authorize Twitter",
      functionName : "authTwitter_"	},
    null,
    { name : "[Engr] Add TimeZone and Locale Menu",
     functionName : "localeTimeZoneMenu_" },
    null,
    { name : "[Engr] Create New Sheet",
      functionName : "createNewSpreadSheet_"},
    { name : "[Engr] Copy OldTweetMessages to new Sheet",
     functionName : "makeSpreadSheetCopy_"},
    { name : "[Engr] Move OldTweetMessages to new Sheet",
      functionName : "moveOldTweetsToNewSpreadSheet_" },   
    null,
    { name : "[Engr] Add Version, Release Menu",
     functionName : "versionReleaseDistroMenu_"}
  ];
  ss.addMenu("Engineering", engineering);
  
  setStatusInfoForUser_('You have been Warned.', 
         'Engineering Menu Enabled Follow instructions from Engineer',
                        'WARN');
}

/**
 * Function to set up the Pro menu for use by user who
 * knows what he/she is doing and hence considered a Pro
 */
function proMenu_() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var proMenu = [
      {	name : "[Pro User] Initial Setup",
		functionName : "getKeySecretsFromUser_"	},
      {	name : "[Pro User] Finalising Setup",
	    functionName : "setupKeysSecretsFromPropertiesToSheet_"	},
      {	name : "[Pro User] Clear Tweets to Prepare/Re-prepare",
		functionName : "clearTweets_" },
      {	name : "[Pro User] Clear Status to send again",
		functionName : "clearStatus_" },
      {	name : "[Pro User] Tweet again",
		functionName : "tweetAgain"	}
     ];

	ss.addMenu("Pro Menu", proMenu);

	setStatusInfoForUser_('You have been Warned.',
		'Pro Menu Enabled. We hope you know what you are doing!',
			 'WARN');
}

/**
 * Function to change Version numbers, release numbers and dates etc
 */
function versionReleaseDistroMenu_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var versionReleaseDistroMenu = [
    { name : "[DeepEngr] Bump Release and Date",
      functionName : "bumpReleaseDate_" },
    { name : "[DeepEngr] Make New Version (Reset Release and set date)",
      functionName : "makeNewVersion_" },
    null,  
    { name : "[DeepEngr] Create New Release",
      functionName : "setDistroReady" }
    ];
  	ss.addMenu("VersionReleaseDistro", versionReleaseDistroMenu);

	setStatusInfoForUser_('Second WARNING!',
		'Version Release Distro Menu Enabled. You could get into deeper trouble!',
			 'WARN');
}
