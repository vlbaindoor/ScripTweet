/**
 * ******************************** Global Variables Used
 * *****************************************
 */

var BIG_VERSION = "2.7";

var TWITTER_APP_NAME = "";
var CONSUMER_KEY = "";
var CONSUMER_SECRET = "";
var ACCESS_TOKEN = "";
var ACCESS_TOKEN_SECRET = "";

// Set to true when Tweets need to be re-configured and set to false
// when tweets are just re-configured
var TWEET_NEED_RE_PREPARED = true;

// Set to be true when Tweet interval is modified and set to false
// when trigger is set
var TRIGGER_MODIFIED = false;

// To change the local timezone used for triggers, refer to
//             http://joda-time.sourceforge.net/timezones.html
// This gets re-assigned in onOpen
var localTimeZone = "Europe/London";


/**
 * *********************** Constants Used
 * ***************************************
 */
//userStatusMessages
var userAlertError = 'ERROR';
var userAlertOK = 'OK';
var userAlertWarn = 'WARN';
var userAlertNormal = 'NORMAL';

//Strings returned by checkPostLength
var POST_TOO_LONG = 'Too Long';
var POST_HARD_TO_RETWEET = 'Hard To ReTweet';
var POST_GREAT_TO_RETWEET = 'Great To ReTweet';

// To set date format
var	dateFormatString = "dd MMM yyyy at HH:mm am/pm";
// For fileName with date time as part of name
var dateFormatForFileNameString = "yyyy-MM-dd'at'HH:mm:ss";

//To set Triggers - the functon to be called by the trigger
var functionNameForTrigger = 'tweetAgain';

var functionNameForUserCallback = 'userCallback';


// Default Template
var defaultTemplate = "{{Message Text}} URL:{{URL}}";

//Sheet names
var MessagesForTweetingSheet = "Messages for Tweeting";
var SettingsSheet = "Settings";
var TimeZoneLocaleSheet = "Locale and TimeZone";
var TweetedMessagesSheet = "Tweeted Messages Record";

//Cell indexes for Settings Sheet
var TWEET_INTERVAL_CELL_INDEX = "B2";
var TWITTER_APP_NAME_CELL_INDEX = "B3";
var CONSUMER_KEY_CELL_INDEX = "B4";
var CONSUMER_SECRET_CELL_INDEX = "B5";
var ACCESS_TOKEN_CELL_INDEX = "B6";
var ACCESS_TOKEN_SECRET_CELL_INDEX = "B7";
var SCRIPT_PROJECT_KEY_CELL_INDEX = "B8";
var SCRIPT_ABOUT_INDEX = "B9";
var SCRIPT_VERSION_INDEX = "B10";
var SCRIPT_RELEASE_NUMBER_INDEX = "B11";
var SCRIPT_RELEASE_DATE_INDEX = "B12";
var TRIGGER_STATUS_CELL_INDEX = "B13";
var HISTORY_DATE_CELL_INDEX = "B14";
var SCRIPT_TIMEZONE_CELL_INDEX = "B15";

//Cell indexes for 'Messages for Tweeting' Sheet
var USER_INFO_MSG_CELL_INDEX = "F1";
var USER_INFO_STATUS_CELL_INDEX = "G1";
var TRIGGER_STATUS_INFO_CELL = "C1";

var OFFSET_TO_PROTECT_HEADERS = 3;
