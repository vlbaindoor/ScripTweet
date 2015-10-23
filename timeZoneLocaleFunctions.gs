/* TimeZone related functions for getting and setting */
/* For testing these functions comment out all other onOpen functions
 * But enable the one found here.
 * Create a Sheet called as "Locale and TimeZone"
 * The cells A2 and B2 are used to over write with present settings for the Spreadsheet
 * The cells  C2 and D2 are used to set the setting for the spreadsheet
 *
 * IF YOU ARE CONFUSED BY ALL OF THE ABOVE - IGORE THIS WHOLE CODE AND SHEET AND DON'T CHANGE
 * ANYTHING HERE.
 */

// Cell locations for GET data onto sheet called as "Locale and TimeZone"
var GET_LOCALE_CELL_INDEX = "A2";
var GET_TIMEZONE_CELL_INDEX = "B2";

// Cell locations for SETTING from sheet to settings on sheet called as "Locale and TimeZone"
var SET_LOCALE_CELL_INDEX = "C2";
var SET_TIMEZONE_CELL_INDEX = "D2";

/**
 * When spreadsheet is opened this function adds menu items during Testing
 */
function localeTimeZoneMenu_() {
  var menuTitle;
  var menu;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZoneLocaleSheet = ss.getSheetByName(TimeZoneLocaleSheet);
  timeZoneLocaleSheet.showSheet();
  ss.setActiveSheet(timeZoneLocaleSheet);
 
  menu = [{ name : "[DeepEngr] Get present Locale and TimeZone",
            functionName : "getShowTimeZoneLocale_"},
          { name : "[DeepEngr] Set Locale and TimeZone",
            functionName : "setShowTimeZoneLocale_"} ];
  
  menuTitle = "LocaleAndTimeZoneMenu";
  ss.addMenu(menuTitle, menu);

}

/**
 * Function to set the timezone and locale settins for the spreadsheet
 * @param {String} newLocale 
 * @param {String} newTimeZone
 */
function setTimeZoneLocale(newLocale, newTimeZone) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetLocale(newLocale);
  ss.setSpreadsheetTimeZone(newTimeZone); 
}

/**
 * Function to get the present locale and timezone settings
 * for the spreadsheet
 * @returns {Object} present Locale and timezones are
 *                   returned as an object
 */
function getTimeZoneLocale() {
  var presentLocale;
  var presentTimeZone;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  presentLocale = ss.getSpreadsheetLocale();
  presentTimeZone = ss.getSpreadsheetTimeZone();
  return { presentLocale: presentLocale, presentTimeZone: presentTimeZone};
}

/**
 * Function to test the functionality
 */
function getShowTimeZoneLocale_() {
  // Get the present values
  
  var params = getTimeZoneLocale();
  var presentLocale = params.presentLocale;
  var presentTimeZone = params.presentTimeZone;
  Logger.log("Came back with Present Locale is: " + presentLocale);
  Logger.log("Came back with Present TimeZone is: " + presentTimeZone);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZoneLocaleSheet = ss.getSheetByName(TimeZoneLocaleSheet);

  // Set the locale value into spreadsheet cell
  var showLocaleCell = timeZoneLocaleSheet.getRange(GET_LOCALE_CELL_INDEX);
  showLocaleCell.setValue(presentLocale);
  
  // Set the timezone value into spreadsheet cell
  var showTimeZoneCell = timeZoneLocaleSheet.getRange(GET_TIMEZONE_CELL_INDEX);
  showTimeZoneCell.setValue(presentTimeZone);
}

/**
 * Function to test the functionality
 */
function setShowTimeZoneLocale_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZoneLocaleSheet = ss.getSheetByName(TimeZoneLocaleSheet);
  
  // Get Locale value from spreadsheet
  var setLocaleCell = timeZoneLocaleSheet.getRange(SET_LOCALE_CELL_INDEX);
  var newLocale = setLocaleCell.getValue();
  
  // Get TimeZone value from spreadsheet
  var setTimeZoneCell = timeZoneLocaleSheet.getRange(SET_TIMEZONE_CELL_INDEX);
  var newTimeZone = setTimeZoneCell.getValue();

  setTimeZoneLocale(newLocale, newTimeZone);
}

