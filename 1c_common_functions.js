/**
This file contains all functions that are used in multiple files (or are at least generic enough that they could be used in different files)
*/


/**
 * gets Zoom OAuth token for the Zoom App specified by ACCOUNT_ID, CLIENT_ID, and CLIENT_SECRET
 * basically the keycard to accessing Zoom info
 */
function getZoomAccessToken(accountId, clientId, clientSecret) {
  const tokenUrl = 'https://zoom.us/oauth/token';
  const basic = Utilities.base64Encode(`${clientId}:${clientSecret}`);
  const options = {
    method: 'post',
    headers: {
      Authorization: 'Basic ' + basic,
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: {
      grant_type: 'account_credentials',
      account_id: accountId
    },
    muteHttpExceptions: true
  };
  const resp = UrlFetchApp.fetch(tokenUrl, options);
  if (resp.getResponseCode() !== 200) {
    throw new Error('Token request failed: ' + resp.getContentText());
  }
  const data = JSON.parse(resp.getContentText());
  // the below log()s are purely for making sure there is a token, keep it commented out otherwise
  // Logger.log(resp.getResponseCode());
  // Logger.log(resp.getContentText());
  return data.access_token;
}


/**
 * check if the input ISO is the 4th thursday in the month
 */
function isFourthThursday(isoDate) {
  const date = new Date(isoDate);
  // Check if the day is Thursday (4)
  if (date.getDay() !== 4) {
    return false;
  }
  const dayOfMonth = date.getDate();
  // Check if the date is between 22 and 28
  return dayOfMonth >= 22 && dayOfMonth <= 28;
}


/**
 * converts a string representing minutes into hours, minutes, format: 00h 00m
 */
function minutesToHM(minutes) {
  const n = parseFloat(minutes);
  if (!isFinite(n)) return '00h 00m';
  const total = Math.floor(Math.abs(n));
  const h = Math.floor(total / 60);
  const m = total % 60;

  const pad = (v) => String(v).padStart(2, '0');
  return `${pad(h)}h ${pad(m)}m`;
}

/**
 * converts a string representing seconds into hours, minutes, seconds, format: 00h 00m 00s
 */
function secondsToHMS(seconds) {
  const n = parseFloat(seconds);
  if (!isFinite(n)) return '00h 00m 00s';
  const total = Math.floor(Math.abs(n));
  const h = Math.floor(total / 3600);
  const m = Math.floor((total % 3600) / 60);
  const s = total % 60;

  const pad = (v) => String(v).padStart(2, '0');
  return `${pad(h)}h ${pad(m)}m ${pad(s)}s`;
}


/**
 * simple func that takes in a sheet and auto resizes all columns that contain values
 */
function resizeColumnsToFit(sheet) {
  const dataRange = sheet.getDataRange();
  if (dataRange.getNumColumns() > 0) {
    sheet.autoResizeColumns(1, dataRange.getNumColumns());
  }
}


/**
 * a function that sounds stupid, but is needed: converts plain datetime format into ISO 8601
 * example:
 * input: 2026-07-23 17:51:22
 * output: 2026-07-23T17:51:22Z
 */
function convertPlainToISO(plainDT) {
  if (!plainDT) return "";
  if (plainDT.includes('T') && plainDT.includes('Z')) return plainDT;
  return plainDT.trim().replace(' ', 'T') + 'Z';
}

/**
 * Converts input ISO 8601 (UTC) string into a specified locale string (defaulting to PT)
 * iso format: '2023-06-08T18:30:00Z'
 * newTimeZone format: 'America/Los_Angeles'
 */
function convertISOTimeZone(iso, newTimeZone = 'America/Los_Angeles') {
  const dt = new Date(iso);

  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: newTimeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: true
  }).formatToParts(dt);

  const get = (type) => parts.find(p => p.type === type)?.value;

  const year = get('year');
  const month = get('month');
  const day = get('day');
  const hour = get('hour');
  const minute = get('minute');
  const second = get('second');
  const dayPeriod = get('dayPeriod'); // AM/PM

  return `${year}-${month}-${day}, ${hour}:${minute}:${second} ${dayPeriod}`;
}

/**
 * intended to be used after convertISOTimeZone(), returns only the time
 */
function timeOnly(datetime) {
  if (!datetime) return "";
  const i = datetime.indexOf(' ');
  return i >= 0 ? datetime.slice(i + 1) : datetime;
}


/**
 * Prepares the Zoom Meeting UUID for use in URL paths, applying double-encoding
 * if the UUID contains a forward slash or begins with one.
 * @param {string} uuid - The raw Zoom UUID.
 * @returns {string} The URL encoded UUID.
 */
function prepareUuid(uuid) {
  if (!uuid) return "";
  if (uuid.indexOf('/') !== -1 || uuid.startsWith('/')) {
    return encodeURIComponent(encodeURIComponent(uuid));
  }
  return encodeURIComponent(uuid);
}


/**
 * a simple function to get the datetime "days" ago
 */
function daysAgo(days) {
  return new Date(Date.now() - days * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
}


/**
 * sets the active sheet tab to the sheet tab with the name "sheetName"
 */
function goToSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.setActiveSheet(sheet);
}
