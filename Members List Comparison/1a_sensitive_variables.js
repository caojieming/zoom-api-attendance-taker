/**
THIS IS IMPORTANT, READ ALL OF ME:

This file is used to contain the global constants for secrets (sensitive keys, passwords, IDs, etc.) that this Apps Script uses.

If you plan on copy/pasting this Apps Script into somewhere public or into an AI, DOUBLE CHECK AND MAKE SURE YOU DON'T INCLUDE RAW SECRETS HERE.


Google Apps Script has functionality explicitly for storing secrets: go to Project Settings -> Script Properties -> Add script property, and add the key name and value there.

Then, set the values below like such:
const KEY = PropertiesService.getScriptProperties().getProperty('SAMPLE_API_KEY_NAME');


You can also just copy/paste the raw secrets into the constants below, just be ABSOLUTELY SURE that if you share this particular file in any way, shape, or form to somthing/someone you don't trust, REMOVE ALL RAW SECRETS. If you don't, congrats, you're a security hazard.
*/

// ID to spreadsheet that contains all members info, found in the link of the spreadsheet between "/d/" and "/edit"
const FULL_MEMBER_LIST_SPREADSHEET_ID = "";

const ATTENDANCE_RANKER_SPREADSHEET_ID = "";
