/**
 * Function to copy rows from present Messages for Tweeting sheet
 * into a new sheet.
 */
function makeSpreadSheetCopy_() {
  //get the data from current Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var msgsSheet = ss.getSheetByName(MessagesForTweetingSheet);
  var lastRow = msgsSheet.getLastRow();
  var lastColumn = msgsSheet.getLastColumn();
  var dataRange = msgsSheet.getRange(1, 1, lastRow, lastColumn);
  var myData = dataRange.getValues();


  var params = createNewSpreadSheet_();
  var newSS =  params.fileHandle;
  var newMsgsSheet = newSS.getActiveSheet();
  newMsgsSheet.setName('Old '+ MessagesForTweetingSheet);
  newMsgsSheet.getRange(1, 1, lastRow, lastColumn).setValues(myData);
  setStatusInfoForUser_('OK', 'Your old Tweet messages copied to file: ' + params.fileName, 'OK');
}

/**
 * Function to move old tweet messages from the Messages for Tweeting sheet
 * to a new spreadsheet
 */
function moveOldTweetsToNewSpreadSheet_() {
  //get the data from current Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var msgsSheet = ss.getSheetByName(MessagesForTweetingSheet);
  var lastRow = msgsSheet.getLastRow();
  var lastColumn = msgsSheet.getLastColumn();
  var dataRange = msgsSheet.getRange(1, 1, lastRow, lastColumn);
  var myData = dataRange.getValues();


  var params = createNewSpreadSheet_();
  var newSS =  params.fileHandle;
  var newMsgsSheet = newSS.getActiveSheet();
  newMsgsSheet.setName('Old '+ MessagesForTweetingSheet);
  newMsgsSheet.getRange(1, 1, lastRow, lastColumn).setValues(myData);
  setStatusInfoForUser_('OK', 'Your old Tweet messages moved to file: ' + params.fileName, 'OK');

  clearOutRows(); 
}

/**
 * Function to create new spreadsheet. The file name given will have 
 * ScripTweet-OldTweetsOn followed by string representing date and time.
 * @returns {Object} new file name and file handle itself are returned as
 *                   a object containing fileName and fileHandle
 */
function createNewSpreadSheet_() {
  // Create new Spreadsheet
  var newFileName = 'ScripTweet-OldTweetsOn'+ getDate_();
  var newSS =  SpreadsheetApp.create(newFileName);
  setStatusInfoForUser_('OK', 'New File ' + newFileName +' Created.', 'OK');
  return { fileName : newFileName,
           fileHandle : newSS
         };
}

/**
 * Function to get Date & Time
 * @returns {String} which represents date and time - useful for keeping
 *                   track of when the file was created etc.
 */
function getDate_() {
  var d = new Date();
  var dateofDay = new Date(d.getTime());
  return Utilities.formatDate(dateofDay, localTimeZone, dateFormatForFileNameString);
}
