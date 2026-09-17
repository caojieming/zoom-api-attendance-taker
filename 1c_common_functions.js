/**
This file contains all functions that are used in multiple files (or are at least generic enough that they could be used in multiple files)
*/


/**
 * Gets a Zoom OAuth access token using account credentials.
 * @param {string} accountId
 * @param {string} clientId
 * @param {string} clientSecret
 * @returns {string}
 */
function getZoomAccessToken(accountId, clientId, clientSecret) {
  if (!accountId || !clientId || !clientSecret) {
    throw new Error('Missing Zoom OAuth credentials.');
  }

  const tokenUrl = 'https://zoom.us/oauth/token';
  const basicAuth = Utilities.base64Encode(`${clientId}:${clientSecret}`);

  const response = UrlFetchApp.fetch(tokenUrl, {
    method: 'post',
    headers: {
      Authorization: `Basic ${basicAuth}`,
    },
    payload: {
      grant_type: 'account_credentials',
      account_id: accountId,
    },
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const body = response.getContentText();

  if (status !== 200) {
    throw new Error(`Token request failed (${status}): ${body}`);
  }

  const data = JSON.parse(body);
  if (!data.access_token) {
    throw new Error('Token response did not include access_token.');
  }

  // the below log()s are purely for making sure there is a token, keep it commented out otherwise
  // Logger.log(resp.getResponseCode());
  // Logger.log(resp.getContentText());

  return data.access_token;
}


/**
 * Returns true if the given date falls on the fourth Thursday of its month.
 * @param {string|Date} isoDate
 * @returns {boolean}
 */
function isFourthThursday(isoDate) {
  const date = new Date(isoDate);
  if (isNaN(date.getTime())) return false;

  // Thursday = 4 in JavaScript's getDay()
  if (date.getDay() !== 4) return false;

  const dayOfMonth = date.getDate();
  return dayOfMonth >= 22 && dayOfMonth <= 28;
}


/**
 * Pads a number to 2 digits.
 * @param {number} value
 * @returns {string}
 */
function pad2(value) {
  return String(value).padStart(2, '0');
}


/**
 * Converts minutes into "00h 00m".
 * @param {number|string} minutes
 * @returns {string}
 */
function minutesToHM(minutes) {
  const n = parseFloat(minutes);
  if (!isFinite(n)) return '00h 00m';

  const total = Math.floor(Math.abs(n));
  const hours = Math.floor(total / 60);
  const mins = total % 60;

  return `${pad2(hours)}h ${pad2(mins)}m`;
}


/**
 * Converts seconds into "00h 00m 00s".
 * @param {number|string} seconds
 * @returns {string}
 */
function secondsToHMS(seconds) {
  const n = parseFloat(seconds);
  if (!isFinite(n)) return '00h 00m 00s';

  const total = Math.floor(Math.abs(n));
  const hours = Math.floor(total / 3600);
  const mins = Math.floor((total % 3600) / 60);
  const secs = total % 60;

  return `${pad2(hours)}h ${pad2(mins)}m ${pad2(secs)}s`;
}


/**
 * Resizes all populated columns in a sheet to fit their content.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function resizeColumnsToFit(sheet) {
  const dataRange = sheet.getDataRange();
  const numColumns = dataRange.getNumColumns();

  if (numColumns > 0) {
    sheet.autoResizeColumns(1, numColumns);
  }
}


/**
 * Converts a plain datetime string like "2026-07-23 17:51:22"
 * into an ISO-like UTC string like "2026-07-23T17:51:22Z".
 *
 * If the input already appears to be ISO UTC, it is returned unchanged.
 *
 * @param {string} plainDT
 * @returns {string}
 */
function convertPlainToISO(plainDT) {
  if (!plainDT) return '';

  const value = String(plainDT).trim();
  if (value.includes('T') && value.endsWith('Z')) return value;

  return `${value.replace(' ', 'T')}Z`;
}


/**
 * Converts an ISO UTC datetime string into a formatted datetime in another time zone.
 * @param {string} iso
 * @param {string} newTimeZone
 * @returns {string}
 */
function convertISOTimeZone(iso, newTimeZone = 'America/Los_Angeles') {
  const date = new Date(iso);
  if (isNaN(date.getTime())) return '';

  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: newTimeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: true,
  }).formatToParts(date);

  const getPart = (type) => parts.find((p) => p.type === type)?.value || '';

  const year = getPart('year');
  const month = getPart('month');
  const day = getPart('day');
  const hour = getPart('hour');
  const minute = getPart('minute');
  const second = getPart('second');
  const dayPeriod = getPart('dayPeriod');

  return `${year}-${month}-${day}, ${hour}:${minute}:${second} ${dayPeriod}`;
}


/**
 * Extracts only the time portion from a datetime string produced by convertISOTimeZone().
 * @param {string} datetime
 * @returns {string}
 */
function timeOnly(datetime) {
  if (!datetime) return '';

  const index = datetime.indexOf(' ');
  return index >= 0 ? datetime.slice(index + 1) : datetime;
}


/**
 * Prepares a Zoom meeting UUID for use in a URL path.
 * Some UUIDs need double-encoding if they contain forward slashes.
 * @param {string} uuid
 * @returns {string}
 */
function prepareUuid(uuid) {
  if (!uuid) return '';

  const value = String(uuid);
  if (value.includes('/')) {
    return encodeURIComponent(encodeURIComponent(value));
  }

  return encodeURIComponent(value);
}


/**
 * Returns the UTC date string for a number of days ago, in "YYYY-MM-DD" format.
 * @param {number|string} days
 * @returns {string}
 */
function daysAgo(days) {
  const n = parseInt(days, 10);
  if (!isFinite(n)) return '';

  const date = new Date(Date.now() - n * 24 * 60 * 60 * 1000);
  return date.toISOString().split('T')[0];
}


/**
 * Switches the active spreadsheet tab to the sheet with the given name.
 * @param {string} sheetName
 */
function goToSheet(sheetName) {
  if (!sheetName) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}


/**
 * Creates a new spreadsheet and moves it into the specified Drive folder.
 * @param {string} sheetName
 * @param {string} folderId
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function createSpreadsheetInFolder(sheetName, folderId) {
  if (!sheetName || !folderId) {
    throw new Error('Missing sheetName or folderId.');
  }

  const ss = SpreadsheetApp.create(sheetName);
  const file = DriveApp.getFileById(ss.getId());
  const folder = DriveApp.getFolderById(folderId);

  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  return ss;
}
