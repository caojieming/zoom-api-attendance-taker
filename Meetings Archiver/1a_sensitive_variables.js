/**
THIS IS IMPORTANT, READ ALL OF ME:

This file is used to contain the global constants for secrets (sensitive keys, passwords, IDs, etc.) that this Apps Script uses.

If you plan on copy/pasting this Apps Script into somewhere public or into an AI, DOUBLE CHECK AND MAKE SURE YOU DON'T INCLUDE RAW SECRETS HERE.


Google Apps Script has functionality explicitly for storing secrets: go to Project Settings -> Script Properties -> Add script property, and add the key name and value there.

Then, set the values below like such:
const KEY = PropertiesService.getScriptProperties().getProperty('SAMPLE_API_KEY_NAME');


You can also just copy/paste the raw secrets into the constants below, just be ABSOLUTELY SURE that if you share this particular file in any way, shape, or form to somthing/someone you don't trust, REMOVE ALL RAW SECRETS. If you don't, congrats, you're a security hazard.
*/


/* for getting zoom OAuth access tokens, set these according to your Zoom App */
// ACCOUNT_ID is the same for all Zoom apps on the same account
const ACCOUNT_ID = "";

// Zoom App ID + "password" for attendance taker
const ATTENDANCE_CLIENT_ID = "";
const ATTENDANCE_CLIENT_SECRET = "";

// Zoom App ID + "password" for recordings/transcripts getter
const RECORDINGS_CLIENT_ID = "";
const RECORDINGS_CLIENT_SECRET = "";


// for getting meetings + participants, used for filtering only meetings with a certain meeting ID
const MEETING_ID = "";

// the last part of the desired folder link (https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID})
const DRIVE_FOLDER_ID = "";

// ID to spreadsheet that contains all members info, found in the link of the spreadsheet between "/d/" and "/edit"
const MEMBERS_SPREADSHEET_ID = "";

// self explanatory, API key to gemini via Google AI Studio.
const GEMINI_API_KEY = "";
