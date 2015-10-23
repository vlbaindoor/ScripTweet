/***************************************************************************
 * 1. Go to apps.twitter.com and create a new app. Name it what ever you
 *    want - but keep it closer to the Brand identity you are trying to
 *    establish using this tool. 2. Use your website for the Website setting.
 * 3. For the Callback URL use the following URL:
 *    https://script.google.com/macros/d/<YOUR PROJECT KEY HERE>/usercallback
 * 4. Create Access Token by going to the Keys and Access Tokens tab in the
 *    App you just created and click the Generate an Access Token link/button
 *    to connect from the Google Sheet with. Make sure you key in the right
 *    values otherwise this tool would fail to Authenticate you to Twitter and
 *    this tool would not work.
 * 5. Once you've created the Access Token you need to get the following
 *    values to the Settings tab of the Google Sheet:
 *    Consumer Key (API Key), Consumer Secret (API Secret) and Access Token,
 *    Access Token Secret You can do that by pasting them in the text fields in
 *    the HTML form that pops up to help you.
 * 6. After you have done all of the above, save the Spread Sheet as it is
 *    and Reload the spread sheet by using the browser button to refresh
 *    the browser window.
 * 7. Authorize Script with Twitter App using the Menu Option on the
 *    Spread Sheet. The Apps Script code that runs you'll need to grant
 *    Authorization. You'll get a message saying that Authorization
 *    is Required. Accept and Continue. 
 * 8. It would bring up a list of things this Script will need to authorise.
 *    You would need to accept that as well.
 * 
 */

/**
 * This records the Keys and Secrets that are stored as Script Properties
 * back into the spreadsheet itself so that user can see them.
 */
function setupKeysSecretsFromPropertiesToSheet_() {
	// Let us get the settings from the Property store
	var params = getKeySecretsFromProperties();
	setSettingsToSheet_(params.twitterAppName,
                        params.consumerKey, params.consumerSecret,
			            params.accessToken, params.acessTokenSecret);
}


/**
 * This function creates a Template from file using the Html Service
 * which loads the HTML content from the 'KeySecrets.html' file which also
 * executes functions within that html file.
 * Then the validate function within that KeySecrets.html is called
 * which ends up calling setKeySecrets function
 */
function getKeySecretsFromUser_() {
  var html = HtmlService.createTemplateFromFile('KeysSecrets')
               .evaluate().setWidth(1000).setHeight(540)
               .setTitle("Connecting up ScripTweet to Twitter")
               .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  var ss = SpreadsheetApp.getActive();
  ss.show(html);
}

/**
 * The function validate calls this function from KeySecrets.html
 * @param   {Object} params an array containing multiple objects
 * @returns {String} for showing to the user the result
 *                   of setting Key Secrets
 */
function setKeySecrets(params) {
	try {
		doProperty_("ProjectId", params.projectId);
        doProperty_("twitter_app_name", params.twitterAppName.trim());
		doProperty_("consumer_key", params.consumerKey.trim());
		doProperty_("consumer_secret", params.consumerSecret.trim());
		doProperty_("access_token", params.accessToken.trim());
		doProperty_("access_token_secret", params.acessTokenSecret.trim());
        setSettingsToSheet_(params.twitterAppName.trim(),
                        params.consumerKey.trim(), params.consumerSecret.trim(),
			            params.accessToken.trim(), params.acessTokenSecret.trim());
		return "Keys and Secrets saved. You can close this window and proceed to Authorise Twitter.";
	} catch (f) {
		setStatusInfoForUser_('ERROR', f.toString(), 'ERROR');
		return "ERROR: " + f.toString();
	}
}

/**
 * Temporary function to use existing data to put onto project properties
 */
function getAndSetKeySecrets() {
	var params = getPresentKeySecrets_();
	setKeySecrets(params);
}

/**
 * Function to get the present Keys and Secrets from the sheet
 * and return them as on object
 * @returns {Object} the object consists of Keys and Secrets
 */
function getPresentKeySecrets_() {
	getSettingsFromSheet_();
	var projectId = getScriptProjectKey();
	return {
		projectId : projectId,
        twitterAppName : TWITTER_APP_NAME,
		consumerKey : CONSUMER_KEY,
		consumerSecret : CONSUMER_SECRET,
		accessToken : ACCESS_TOKEN,
		acessTokenSecret : ACCESS_TOKEN_SECRET
	};
}

/**
 * This function is called from the HTML file KeySecrets.html.
 * This function is used to get the values stored for Keys and
 * Secrets as Script Properties and returns them as an objet
 * @returns {Object} Keys and Secrets as Script Properties
 */
function getKeySecretsFromProperties() {
	var projectId = getScriptProjectKey();
	return {
		projectId : projectId,
        twitterAppName: doProperty_("twitter_app_name"),
		consumerKey : doProperty_("consumer_key"),
		consumerSecret : doProperty_("consumer_secret"),
		accessToken : doProperty_("access_token"),
		acessTokenSecret : doProperty_("access_token_secret")
	};
}
