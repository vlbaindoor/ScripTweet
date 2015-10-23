/**
 * Function to Hide a column
 * @param {String} columnName string which contains 
 *                            the text used in the column
 */
function hideOneColumn_(columnName) {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(MessagesForTweetingSheet);
    var data = sheet.getDataRange().getValues();
    var headers = data[1];
    var columnToHide = getColumnIndex_(sheet, headers, columnName);
    var range = sheet.getRange(1, columnToHide + 1);
  
    sheet.hideColumn(range);
}

/**
 * Function to unHide or rather show a column
 * @param {String} columnName string which contains 
 *                            the text used in the column
 */
function unHideOneColumn_(columnName) {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(MessagesForTweetingSheet);
    var data = sheet.getDataRange().getValues();
    var headers = data[1];
    var columnToHide = getColumnIndex_(sheet, headers, columnName);
    var range = sheet.getRange(1, columnToHide + 1);
  
    sheet.unhideColumn(range);
}
