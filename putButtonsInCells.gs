/**
* From http://googlesheetshelp.blogspot.fr/2014/11/method-for-displaying-images-stored-on.html
* 1. Make Google Drive Folder Public
* 2. Change the Sharing options so that it is viewable by anyone with the link on the internet without logging in
* 3. Right click on the required file and 'Get Link for Sharing'
* 4. Copy the link to clipboard
* 5. Here's an example url
*    https://drive.google.com/open?id=0B9biAW_mGiSpQ1ZjWTNxLTE2RDQ&authuser=0
* 6. Modify it to look like:
*    http://drive.google.com/uc?export=view&id=0B9biAW_mGiSpQ1ZjWTNxLTE2RDQ
* 7. What follows the '?id=' is what is required as suffix
* 8. The prefix is "http://drive.google.com/uc?export=view&id="
* 9. Combine the prefix and suffix to get the URL
* 10. Use the URL
*
* But "Serge insas"  who is a Google Apps Script Top Contributor says:
* (http://stackoverflow.com/questions/26801041/what-is-the-right-way-to-put-a-drive-image-into-a-sheets-cell-programmatically)
* var img = DriveApp.getFileById('image ID'); // or any other way to get the image object
* // in this example the image is in column E
* var imageInsert = sheet.getRange(lastRow+1, 5)
*            .setFormula('=image("https://drive.google.com/uc?export=view&id='+img.getId()+'")');
*            
*  // define a row height to determine the size of the image in the cell
*  sheet.setRowHeight(lastRow+1, 80);
*/
  // https://drive.google.com/open?id=0B7YVNdei_3c-SG9pZmx1dy1lMTA
  // http://drive.google.com/uc?export=view&id=

var UrlPrefix = "http://drive.google.com/uc?export=view&id="
var button1UrlSuffix = "0B7YVNdei_3c-SG9pZmx1dy1lMTA";
var button2UrlSuffix = "0B7YVNdei_3c-Mms3UkhXaUU2dzA";

function putButtonsInCells() {
  
  
}
