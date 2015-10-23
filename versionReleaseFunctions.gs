/******************* Version, Release, Release Date related functions ******************/

//Function to bump version number and get ScripTweet ready for distribution
function setDistroReady() {
	factoryDefault();
	bumpReleaseNumber_();
	setReleaseDate_();
}

function bumpReleaseDate_()
{
  bumpReleaseNumber_();
  setReleaseDate_();
}

// Make New Version, Reset Release and set date
function makeNewVersion_() {
  bumpVersionNumber_();
  setReleaseNumber_(0.1);
  setReleaseDate_();
}

function getVersionNameNumber() {
  var versionNumber = getVersionNumber_();
  var versionRelease = getReleaseNumber_();
  var releaseDate = getReleaseDate_();
  
  // Let us convert the numbers into strings
  versionNumber = Utilities.formatString( '%.1f', versionNumber);
  versionRelease = Utilities.formatString( '%.1f', versionRelease);
  
  var versionNameNumber = versionNumber + '.'
                           + versionRelease
                           + ' Release Date: '
                           + releaseDate;
  Logger.log("Returning VersionNumberDate as: " + versionNameNumber);
  return versionNameNumber;
}

//The following function to increase the version number by 0.1
function bumpVersionNumber_() {
	var versionNumber = getVersionNumber_();
	setVersionNumber_(versionNumber + 0.1);
}

//This function is called to set version number to specified in parameter
function setVersionNumber_(versionNumber) {
	var settings = SpreadsheetApp.getActiveSpreadsheet()
                      .getSheetByName(SettingsSheet);
	settings.getRange(SCRIPT_VERSION_INDEX).setValue(versionNumber);
    settings.getRange(SCRIPT_VERSION_INDEX).setNumberFormat("0.0");
}

//This function is called to get version number - it returns it as a string
function getVersionNumber_() {
  var settings = SpreadsheetApp.getActiveSpreadsheet()
                               .getSheetByName(SettingsSheet);
  var versionNumber = settings.getRange(SCRIPT_VERSION_INDEX).getValue();
  var versionNumberStr = Utilities.formatString( '%.1f', versionNumber);
  if (versionNumberStr != BIG_VERSION) {
      Browser.msgBox("WARNING",
                     'Please NOTE: Your GlobalConstantsVariable says Version is: '
                         + BIG_VERSION
                         + ' But your Settings Sheet says it is : ' 
                         + versionNumberStr
                         + ' Please correct it. Select OK to continue.',
                      Browser.Buttons.OK);
  }
  return versionNumber;
}

//The following function to increase the Release number by 0.1
function bumpReleaseNumber_() {
	var releaseNumber = getReleaseNumber_();
	setReleaseNumber_(releaseNumber + 0.1);
}

//This function is called to set Release number to specified in parameter
function setReleaseNumber_(releaseNumber) {
	var settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
			SettingsSheet);
	settings.getRange(SCRIPT_RELEASE_NUMBER_INDEX).setValue(releaseNumber);
    settings.getRange(SCRIPT_RELEASE_NUMBER_INDEX).setNumberFormat("0.0");
}

//This function is called to get version number - it returns it as a string
function getReleaseNumber_() {
	var settings = SpreadsheetApp.getActiveSpreadsheet()
                    .getSheetByName(SettingsSheet);
	var releaseNumber = settings.getRange(SCRIPT_RELEASE_NUMBER_INDEX).getValue();   
    return releaseNumber;
}

//This function is called to set release date to NOW
function setReleaseDate_() {
	var settings = SpreadsheetApp.getActiveSpreadsheet()
                      .getSheetByName(SettingsSheet);
	settings.getRange(SCRIPT_RELEASE_DATE_INDEX).setValue(new Date());
	settings.getRange(SCRIPT_RELEASE_DATE_INDEX).setNumberFormat(dateFormatString);
}

//This function is called to get release date - it returns it as a string
function getReleaseDate_() {
	var settings = SpreadsheetApp.getActiveSpreadsheet()
                      .getSheetByName(SettingsSheet);
	var releaseNumber = settings.getRange(SCRIPT_RELEASE_DATE_INDEX).getValue();
	return releaseNumber;
}

function getVersionNameNumberURL() {
  var versionNameNumber = getVersionNameNumber();
  var spreadSheetUrl = getSpreadSheetUrl();
  var sharingUrl = spreadSheetUrl.replace('\/edit', '\/copy?usp=sharing');
  return {versionName : versionNameNumber,
          spreadSheetURL : spreadSheetUrl,
          sharingURL : sharingUrl};
}

function getSpreadSheetUrl() {
  var ss = SpreadsheetApp.getActive();
  var spreadSheetUrl = ss.getUrl();
  return spreadSheetUrl;
}
