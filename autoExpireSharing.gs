var EXPIRY_TIME  = "2016-01-01 23:42"; 
/**
 * Function to auto expire sharing of the file at the EXPIRY_TIME
 */
function autoExpire_() {
  var id;
  var asset;
  var i;
  var email;
  var users;
 
  // The URL of the Google Drive file or folder 
  var URL = "https://drive.google.com/folderview?id=0B4fk8L6brI_ednJaa052";
  
  try {
    // Extract the File or Folder ID from the Drive URL
    var id = URL.match(/[-\w]{25,}/);
    
    if (id) {
      asset = DriveApp.getFileById(id) ? DriveApp.getFileById(id) : DriveApp.getFolderById(id);
      if (asset) {
        // Make the folder / file Private 
        asset.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.NONE);  
        asset.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.NONE); 
        
        // Remove all users who have edit permissions
        users = asset.getEditors();
        for (i in users) {
          email = users[i].getEmail();
          if (email != "") {
            asset.removeEditor(email);
          }
        }
        
        // Remove all users who have view permssions
        users = asset.getViewers();
        for (i in users) {
          email = users[i].getEmail();
          if (email != "") {
            asset.removeViewer(email);
          }
        }  
      }
    }
  
  } catch (e) {
    Logger.log(e.toString());
  }
}

/**
 * Gets the triggers and adds the trigger at EXPIRY TIME
 */
function start_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  var time = EXPIRY_TIME;
  
  // Run the auto-expiry script at this date and time
  var expireAt = new Date(time.substr(0,4),
                          time.substr(5,2)-1,
                          time.substr(8,2),
                          time.substr(11,2),
                          time.substr(14,2));
  
  if ( !isNaN ( expireAt.getTime() ) ) {
    ScriptApp.newTrigger("autoExpire").timeBased().at(expireAt).create();
  }
}
