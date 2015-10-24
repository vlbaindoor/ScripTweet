// Get file by name in the user's Drive
function getMyFileByName_(fileName) {
  var files = DriveApp.getFilesByName(fileName);
  if (!files)
    return null;
  Logger.log("Looking for files by name %s", fileName);
  while (files.hasNext()) {
    var file = files.next();
    Logger.log("We got file name %s", file.getName());
    if (file.getName() == fileName) {
      // Logger.log("Found the file and the ID is " + file.getId());
      return file;
    }
  }
}

/**
* Upload a single image to Twitter and retrieve the media ID for later use in 
* sendTweet() (using the media_id_string params)
*
* @param {Blob} imageblob the Blob object representing the image data to upload
* @return {object} the Twitter response as an object if successful, null otherwise
*/
function uploadMedia_(imageblob) {
  var twitterService = getTwitterService_();
  if (!twitterService.hasAccess()) {
    return twitterService.authorize();
  }
  var url = "https://upload.twitter.com/1.1/media/upload.json";
  var old_location = twitterService.paramLocation_;
  
  var options = {
    method: "POST",
    payload: { "media" : imageblob }
  };
  
  twitterService.paramLocation_ = "uri-query";
  try {
    var result = twitterService.fetch(url, options);
    Logger.log("Upload media success. Response was:\n" + result);
    return JSON.parse(result.getContentText("UTF-8"));
  }  
  catch (e) {
    options.payload = options.payload && options.payload.length > 100 ? "<truncated>" : options.payload;
    Logger.log("Upload media failed. Error was:\n" + JSON.stringify(e) 
                     + "\n\noptions were:\n" + JSON.stringify(options) + "\n\n");
    return e;  // Changed from null to e so we can report the error messages returned.
  } finally {
    twitterService.paramLocation_ = old_location;
  }
}

/**
* Function to handle media if specified cell contains file name and it accessible
* Twitter API allows for upto 4 images.
* @param {Blob} imageblob the Blob object representing the image data to upload
* @return {object} the Twitter Media IDs as an object if successful, null otherwise
*/
function handleMedia_(sheet, rowIndex, dataValues) {
  var data = sheet.getDataRange().getValues();
  var headers = data[1];
  var colImageFileName = getColumnIndex_(sheet, headers, "Image File Name");
  var TwitterMediaIndex = getColumnIndex_(sheet, headers, "Twitter Media IDs");
  var params = "";
  // Handle any media to add to the post.
  if (sheet.getRange(rowIndex+1, colImageFileName+1).getValue()) {
    Logger.log("Got Media ID to handle");
    var mediaIds = dataValues[rowIndex][colImageFileName].split('\n');
    if (mediaIds.length > 4) {
      throw "Error: Up to 4 images can be uploaded to one tweet.";
    }
    var twitterMediaIds = [];
    mediaIds.forEach(function(fileName) {
      var media = getMyFileByName_(fileName);
      if (media) {
        if (media.getSize() > 3145728) {
          throw "Error: Image size over 3MB. Size: " + media.getSize();
        }
        var mediaResponse = uploadMedia_(media.getBlob());
        // If an error occurred throw an exception with the message returned from the upload.
        if (mediaResponse.hasOwnProperty("message")) {
          throw "Error uploading Image: " + mediaResponse.message;
        }
        twitterMediaIds.push(mediaResponse.media_id_string);
      } else {
        throw "Error: Invalid Media ID {" + fileName + "}";
      }
    });
    // Setup media_ids to be used in Status Update call
    params = { media_ids: twitterMediaIds.join(",") };
    Logger.log("Saving Media IDs in sheet");
    // Save Twitter Media IDs in sheet for reference
    sheet.getRange(rowIndex+1, TwitterMediaIndex+1).setValue(twitterMediaIds.join("\n"));
  }
  return params;
}

function tweetTweet_(tweet, params) {
  Logger.log("Going to tweet: <%s>", tweet);

  var twitterService = getTwitterService_();
  if (!twitterService.hasAccess()) {
    return twitterService.authorize();
  }
  var payload = {
    "status" : tweet
  };

  if (params) {
    for(var i in params) {
      if(params.hasOwnProperty(i)) {
        payload[i.toString()] = params[i];   
      }
    }
  }

  var options = {
    method: "POST",
    payload: payload,
    muteHttpExceptions : true
  };
  
  var statusUrl = "https://api.twitter.com/1.1/statuses/update.json";
  
  try {    
    var result = twitterService.fetch(statusUrl, options);
    Logger.log("Send tweet success. Response was: " + result.getContentText("UTF-8")); 
    return JSON.parse(result.getContentText("UTF-8"));
  } catch (e) {
    Logger.log("Send tweet failure. Error was:\n" + JSON.stringify(e) + "options were:\n" + JSON.stringify(options));
    throw e;  // Changed from null to e so we can check the error messages returned.
  }
}

/*  For testing.
* Use Vivek'sJustAMinuteThumbnail.gif as file name
function getMyFile() {
 var fileName;
 var msg;
 // Display a dialog box with a title, message, input field, and "Yes" and "No" buttons. The
 // user can also close the dialog by clicking the close button in its title bar.
 var ui = SpreadsheetApp.getUi();
 var response = ui.prompt('File Name Please', 'File Name Please:', ui.ButtonSet.YES_NO);

 // Process the user's response.
 if (response.getSelectedButton() == ui.Button.YES) {
   Logger.log('File Name entered is %s.', response.getResponseText());
   fileName = response.getResponseText();
 } else if (response.getSelectedButton() == ui.Button.NO) {
   Logger.log('The user didn\'t want to provide a file name.');
 } else {
   Logger.log('The user clicked the close button in the dialog\'s title bar.');
 }
 
 Logger.log("Going to look for the file %s", fileName);
 var file = getMyFileByName_(fileName);
 msg = "File Name: " + fileName + " and ID is :<" + file.getId() + ">";
 Logger.log(msg);
 ui.alert( msg);
}
*/
