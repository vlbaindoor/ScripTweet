# ScripTweet
Twitter tool using Google Apps Script.

This is a tool in development and already being Beta tested.

The ScripTweet app lets you send Tweets from a Google Spreadsheet.

This is a Google Spreadsheet Apps Script which allows you to add Tweet messages in its rows and send out those messages periodically depending on your choice of time interval. When there are more than one message to be tweeted, there would be a delay between each Tweet.

This adds a custom menu to the top of the Spreadsheet and encourages users to use options from the menu. Although it is possible to edit the Spreadsheet cells directly it is not recommended. Some cells should not be actually manually edited as it may affect the performance of this tool.

Please note that this is still a Work In Progress

WARNING

Do not put symbols in the text message
This includes the commonly used symbol for 'and' which is "&"
Do NOT use it

If you do then you may have unexplainable behaviour!
Use the "@" symbol only to refer to someone's Twitter Handle and not for anything else.

Here are general Oprating Instructions - but these are valid only after the initial set up is succesfully completed.

ALLWAYS use the Blue "1. Add Message To TWEET" button to add new messages for periodic tweeting
You can use this button repeatedly to add more messages.
After you have added message or messages for Tweeting, you need to click on the "2. Prepare TWEET" button to assign a template, save the template so that the tool can create the actual messages to be Tweeted. These messages get saved on the same SpreadSheet under a column called "Tweet". Please do not directly edit the entries in this column
Once you are happy with the entries you see under the "Tweet" column, you can go ahead and click on the "3. TWEET Away" button.
Every humanly possible effort is made to guide you in using this tool throughout but these are guidelines
The "TWEET Regularly" button is used to set up a timer based trigger which runs on Google's Servers. This trigger would go through the rows in this SpreadSheet and Tweet them at regular intervals.
You can use the button named "STOP Regular Tweets" to remove the timer based trigger. You may have to use this button to stop the existing timer trigger and start a new one if in case you changed the "Tweet Interval in Minutes" from the "Settings" Sheet - which is a sub sheet of this Spread Sheet.
NOTE

The Google URL Shortener API is to be enabled by going into Resources and Advanced Google Services, scroll down, find Google URL Shortener API and Enable it.
You may also need to go into Google Developers Console and enable them.

Enabling Twitter Connectivity, Authentication etc

Go to Twitter’s developer dashboard by browsing to
https://apps.twitter.com/
create a new app and copy the Consumer Secret and Key.
These will have to be copied from Twitter Website and copied onto the corresponding place in the "Settings" Sheet of this Spreadsheet.

To restart call rewokeTwitterService, then call the authTwitter_

Libraries etc

Under Resouces -> Libraries -> Find a Library copy paste the following string to search

Mb2Vpd5nfD3Pz-_a-39Q4VfxhMjh3Sh48
It should find OAuth1 library - this needs to be added to resources used by the script.
You may have to select a specific version of the libarary after you you done the previous step.
