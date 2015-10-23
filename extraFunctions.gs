/**
 * Functions that can be used to protect and unprotect cells and prevent user
 * entry without using menu etc - for future development
 */
function disableUserDataEntry_() {
	var ss = SpreadsheetApp.getActive();
	var sheet = ss.getSheetByName(MessagesForTweetingSheet);
	var data = sheet.getDataRange().getValues();
	var headers = data[1];

	var colMessage = getColumnIndex_(sheet, headers, "Message Text");
	var colLongURL = getColumnIndex_(sheet, headers, "Long URL");
	var colURL = getColumnIndex_(sheet, headers, "URL");
	var colTweet = getColumnIndex_(sheet, headers, "Tweet");
	var colStatus = getColumnIndex_(sheet, headers, "Status");

	for (var i = 3; i < data.length; i++) {
		var range1 = sheet.getRange(i + 1, colMessage + 1);
		protectRange_(range1);

		var range2 = sheet.getRange(i + 1, colLongURL + 1);
		protectRange_(range2);

		var range3 = sheet.getRange(i + 1, colURL + 1);
		protectRange_(range3);

		var range4 = sheet.getRange(i + 1, colURL + 1);
		protectRange_(range4);

		var range5 = sheet.getRange(i + 1, colTweet + 1);
		protectRange_(range5);
	}
}

function enableUserDataEntry_() {
	var ss = SpreadsheetApp.getActive();
	var sheet = ss.getSheetByName(MessagesForTweetingSheet);
	var data = sheet.getDataRange().getValues();
	var headers = data[1];

	var colMessage = getColumnIndex_(sheet, headers, "Message Text");
	var colLongURL = getColumnIndex_(sheet, headers, "Long URL");
	var colURL = getColumnIndex_(sheet, headers, "URL");
	var colTweet = getColumnIndex_(sheet, headers, "Tweet");
	var colStatus = getColumnIndex_(sheet, headers, "Status");
	for (var i = 3; i < data.length; i++) {
		var range1 = sheet.getRange(i + 1, colMessage + 1);
		unProtectRange_(range1);

		var range2 = sheet.getRange(i + 1, colLongURL + 1);
		unProtectRange_(range2);

		var range3 = sheet.getRange(i + 1, colURL + 1);
		unProtectRange_(range3);

		var range4 = sheet.getRange(i + 1, colURL + 1);
		unProtectRange_(range4);

		var range5 = sheet.getRange(i + 1, colTweet + 1);
		unProtectRange_(range5);
	}
}

function protectRange_(range) {
	var protection = range.protect().setDescription(
			'To edit choose Start Adding new Tweet Messages option in menu');
}

function unProtectRange_(range) {
	var protection = range.protect().remove();
}

function unProtectAll_() {
	// Remove all range protections in the spreadsheet that the user has
	// permission to edit.
	var ss = SpreadsheetApp.getActive();
	var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
	for (var i = 0; i < protections.length; i++) {
		var protection = protections[i];
		if (protection.canEdit()) {
			protection.remove();
		}
	}
}

