function hideOneColumn_(columnName) {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(MessagesForTweetingSheet);
    var data = sheet.getDataRange().getValues();
    var headers = data[1];
    var columnToHide = getColumnIndex_(sheet, headers, columnName);
    var range = sheet.getRange(1, columnToHide + 1);
  
    sheet.hideColumn(range);
}

function unHideOneColumn_(columnName) {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(MessagesForTweetingSheet);
    var data = sheet.getDataRange().getValues();
    var headers = data[1];
    var columnToHide = getColumnIndex_(sheet, headers, columnName);
    var range = sheet.getRange(1, columnToHide + 1);
  
    sheet.unhideColumn(range);
}
