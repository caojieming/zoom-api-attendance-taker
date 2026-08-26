/*
THIS IS IMPORTANT, READ ME:

This file is used to contain all sensitive keys, passwords, IDs, etc. that this Apps Script uses.

If you plan on copy/pasting this Apps Script into somewhere public or into an AI, DO NOT INCLUDE THIS FILE DIRECTLY.

For backup purposes or forking into a new project, just create a new file with constants of the same names as below.

For plugging into an AI, you shouldn't need to include this file at all, the AI should be smart enough to recognize whenever these constants are called that it's sensitive info that is stored elsewhere.


Technically, Google Apps Script has functionality explicitly for storing important keys, but for the sake of simplicity, I'm not using it.

If you ever want to change to using that functionality, go to Project Settings -> Script Properties -> Add script property, and add the key name and value there.

To then access that key, refer to it in the script like this:
const key = PropertiesService.getScriptProperties().getProperty('SAMPLE_API_KEY_NAME');
*/


/* for getting zoom access token */
// set these according to your Zoom App
const ACCOUNT_ID = '';
const CLIENT_ID = '';
const CLIENT_SECRET = '';


/* for getting meetings + participants */
// used for filtering only meetings with a certain meeting ID
const MEETING_ID = "";
