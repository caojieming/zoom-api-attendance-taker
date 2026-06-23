// helper constants used for setting FROM and TO times
const NOW = new Date().toISOString().split('T')[0]; // Today (YYYY-MM-DD)
const ONE_DAY_AGO = new Date(Date.now() - 1 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
const THREE_DAYS_AGO = new Date(Date.now() - 1 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
const ONE_WEEK_AGO = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]; // One week ago (YYYY-MM-DD)
const ONE_MONTH_AGO = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
const TWO_MONTHS_AGO = new Date(Date.now() - 60 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
const THREE_MONTHS_AGO = new Date(Date.now() - 90 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
const FOUR_MONTHS_AGO = new Date(Date.now() - 120 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
const FIVE_MONTHS_AGO = new Date(Date.now() - 150 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
const SIX_MONTHS_AGO = new Date(Date.now() - 180 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];


// request constants, these are sent to Zoom API as part of the request
// time period of past meetings to GET
const FROM = SIX_MONTHS_AGO;
const TO = NOW;
// Type of meeting (meeting or webinar, can also send "" for both)
const MEETING_TYPE = "meeting";
// Optional search query key if you only want meetings with specific word(s) in the topic name
const SEARCH_KEY = "";
// Max meetings per request page (up to 300), used to lower request rate to prevent hitting API rate limits
const PAGE_SIZE = 200;


// extra constants, used for misc filtering
const MEETING_ID = "";

// toggle to include only 4th thursdays of the month
const ONLY_FOURTH_THURS = true;

// if participant name has any of these words, exclude them from the sheet
const PARTICIPANT_BLACKLIST = ['notetaker', 'read.ai'];



// can't get more than 1 month worth of records at a time, need to call multiple times
function getParticipantsHalfYear() {
  getParticipants(ONE_MONTH_AGO, NOW);
  getParticipants(TWO_MONTHS_AGO, ONE_MONTH_AGO);
  getParticipants(THREE_MONTHS_AGO, TWO_MONTHS_AGO);
  getParticipants(FOUR_MONTHS_AGO, THREE_MONTHS_AGO);
  getParticipants(FIVE_MONTHS_AGO, FOUR_MONTHS_AGO);
  getParticipants(SIX_MONTHS_AGO, FIVE_MONTHS_AGO);
}


/**
 * Main function to get historical Zoom meetings within the past week,
 * extract details and participant lists, and write them into Google Sheets.
 * inFrom: start date of time period observed, defaulting to const FROM
 * inTo: end date of time period observed, defaulting to const TO
 */
function getParticipants(inFrom = FROM, inTo = TO) {
  // Fetch access token using existing client function (assumed to be defined globally)
  var accessToken = getZoomAccessToken();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var meetings = [];
  var nextMeetingPageToken = "";

  // 1. Paginated fetch of past meetings from the Zoom report API
  do {
    var url = "https://api.zoom.us/v2" + "/report/history_meetings" +
      "?from=" + inFrom +
      "&to=" + inTo +
      "&page_size=" + PAGE_SIZE +
      "&meeting_type=" + MEETING_TYPE;

    if (SEARCH_KEY) {
      url += "&search_key=" + encodeURIComponent(SEARCH_KEY);
    }
    if (nextMeetingPageToken) {
      url += "&next_page_token=" + encodeURIComponent(nextMeetingPageToken);
    }

    var options = {
      method: "get",
      headers: {
        "Authorization": "Bearer " + accessToken
      },
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();

    if (responseCode === 200) {
      var data = JSON.parse(response.getContentText());
      // console.log(data);
      if (data.history_meetings && data.history_meetings.length > 0) {
        meetings = meetings.concat(data.history_meetings);
      }
      nextMeetingPageToken = data.next_page_token;
    } else {
      console.error("Error fetching historical meetings. Code: " + responseCode + ", Response: " + response.getContentText());
      nextMeetingPageToken = ""; // Stop pagination on error
    }
  } while (nextMeetingPageToken);

  // console.log(meetings);

  // Filter for unique meetings based on UUID
  var uniqueMeetings = [];
  var seenUuids = new Set();
  meetings.forEach(function (meeting) {
    if (meeting.meeting_uuid && !seenUuids.has(meeting.meeting_uuid)) {
      seenUuids.add(meeting.meeting_uuid);
      uniqueMeetings.push(meeting);
    }
  });

  // console.log(uniqueMeetings);
  // console.log(seenUuids);

  // 2. Iterate through each unique meeting and build the spreadsheets
  uniqueMeetings.forEach(function (meeting) {
    var meetingId = meeting.meeting_id;

    // optional check/filter for meeting ID
    // if const MEETING_ID is filled/is not empty
    if (MEETING_ID !== "") {
      // if current meeting ID does not equal const MEETING_ID, skip (note, use soft inequality check: meetingId is apparently not a string)
      if (meetingId != MEETING_ID) {
        return;
      }
    }

    var startTime = meeting.start_time || "";
    // optional check/filter if only looking for meetings on the fourth thursday of the month
    if (ONLY_FOURTH_THURS && !isFourthThursday(startTime)) {
      return;
    }

    var endTime = meeting.end_time || "";
    var rawUuid = meeting.meeting_uuid;
    var topic = meeting.topic || "Untitled Meeting";
    var duration = meeting.duration || 0;
    var hostDisplayName = meeting.host_display_name || "";
    var hostEmail = meeting.host_email || "";
    // var totalParticipantsCount = meeting.participants;
    // will manually count totalParticipantsCount later, this is important if we enable EXCLUDE_NOTETAKERS

    // Format sheet name: MM/DD/YYYY, Topic
    // var dateString = formatMeetingDate(startTime);
    // var sanitizedTopic = topic.replace(/[\\?\*:\[\]]/g, "-"); // Replace forbidden sheet characters
    // var sheetName = convertISOTimeZone(startTime) + ", " + sanitizedTopic;
    var sheetName = convertISOTimeZone(startTime);
    if (sheetName.length > 50) {
      sheetName = sheetName.substring(0, 50) + "...";
    }

    // Skip if a sheet with this name already exists
    if (ss.getSheetByName(sheetName)) {
      console.log("Sheet '" + sheetName + "' already exists. UUID: " + rawUuid + ". Skipping.");
      return;
    }

    // Fetch all participants for this meeting UUID
    var participants = [];
    var participantNextPageToken = "";
    var encodedUuid = prepareUuid(rawUuid);

    do {
      var partUrl = "https://api.zoom.us/v2/past_meetings/" + encodedUuid + "/participants?page_size=300";
      if (participantNextPageToken) {
        partUrl += "&next_page_token=" + encodeURIComponent(participantNextPageToken);
      }

      var partOptions = {
        method: "get",
        headers: {
          "Authorization": "Bearer " + accessToken
        },
        muteHttpExceptions: true
      };

      var partResponse = UrlFetchApp.fetch(partUrl, partOptions);
      var partResponseCode = partResponse.getResponseCode();

      if (partResponseCode === 200) {
        var partData = JSON.parse(partResponse.getContentText());
        if (partData.participants && partData.participants.length > 0) {
          participants = participants.concat(partData.participants);
        }
        participantNextPageToken = partData.next_page_token;
      } else {
        console.error("Error fetching participants for UUID: " + rawUuid + ". Code: " + partResponseCode + ", Response: " + partResponse.getContentText());
        participantNextPageToken = ""; // Stop pagination on error
      }
    } while (participantNextPageToken);

    // sanitize participants list (optionally remove Notetakers, merge dupe names)
    var sanitizedParticipants = [];
    participants.forEach(function (curParticipant) {
      // skip cases
      if (PARTICIPANT_BLACKLIST.length > 0 && PARTICIPANT_BLACKLIST.some(keyword => curParticipant.name.toLowerCase().includes(keyword.toLowerCase()))) {
        return;
      }

      // check for dupes (participant has has left and is rejoining)
      let dupeMerged = false;
      sanitizedParticipants.forEach(function (pastParticipant) {
        // participant is a dupe: merge participant into past participant
        if (curParticipant.name === pastParticipant.name) {
          pastParticipant.leave_time = curParticipant.leave_time;
          pastParticipant.duration = pastParticipant.duration + curParticipant.duration;
          pastParticipant.timesRejoined += 1;
          dupeMerged = true;
          return;
        }
      });

      // participant is not a dupe: add as a new entry
      if (!dupeMerged) {
        curParticipant.timesRejoined = 0;
        sanitizedParticipants.push(curParticipant);
        totalParticipantsCount++;
      }

    });
    var totalParticipantsCount = sanitizedParticipants.length;

    // if totalParticipantsCount = 0, then all participants were notetakers: skip this meeting
    // if totalParticipantsCount = 1, then it can hardly by called a meeting: skip this meeting
    if (totalParticipantsCount <= 1) {
      return;
    }


    // Insert new sheet for the meeting
    var newSheet = ss.insertSheet(sheetName);

    // 3. Generate and place the meeting details table first (starts at column H / 8)
    var detailsHeaders = [
      "meeting_uuid",
      "meeting_id",
      "topic",
      "host_display_name",
      "host_email",
      "participants",
      "duration",
      "start_time",
      "end_time"
    ];

    var detailsRow = [
      rawUuid,
      meetingId,
      topic,
      hostDisplayName,
      hostEmail,
      totalParticipantsCount,
      minutesToHM(duration),
      timeOnly(convertISOTimeZone(startTime)),
      timeOnly(convertISOTimeZone(endTime))
    ];


    // 4. Generate and place the participants table next (starts at column A / 1)
    var participantHeaders = [
      // "id",
      "name",
      // "user_email",
      "join_time",
      "leave_time",
      "duration_sec",
      "duration",
      "rejoined"
    ];

    var participantRows = sanitizedParticipants.map(function (p) {
      return [
        // p.id || "",
        p.name || "",
        // p.user_email || "",
        timeOnly(convertISOTimeZone(p.join_time)) || "",
        timeOnly(convertISOTimeZone(p.leave_time)) || "",
        p.duration || 0,
        secondsToHMS(p.duration) || 0,
        p.timesRejoined || 0
      ];
    });

    // getRange(row, col, num rows, num cols)
    // set meeting details
    newSheet.getRange(1, participantHeaders.length + 2, 1, detailsHeaders.length).setValues([detailsHeaders]);
    newSheet.getRange(2, participantHeaders.length + 2, 1, detailsRow.length).setValues([detailsRow]);

    // set participant details
    // (row, col, num rows, num cols)
    newSheet.getRange(1, 1, 1, participantHeaders.length).setValues([participantHeaders]);
    newSheet.getRange(2, 1, participantRows.length, participantHeaders.length).setValues(participantRows);

    // auto resize columns
    resizeColumnsToFit(newSheet);

    // add a filter to columns A to F
    newSheet.getRange(1, 1, participantRows.length + 1, participantHeaders.length).createFilter();

  });
}


// check if the input ISO is the 4th thursday in the month
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

// converts a string representing minutes into hours, minutes
function minutesToHM(minutes) {
  const n = parseFloat(minutes);
  if (!isFinite(n)) return '0h 0m 0s';
  const total = Math.floor(Math.abs(n));
  const h = Math.floor(total / 60);
  const m = total % 60;
  return `${h}h ${m}m`;
}

// converts a string representing seconds into hours, minutes, seconds, format: "0h 0m 0s"
function secondsToHMS(seconds) {
  const n = parseFloat(seconds);
  if (!isFinite(n)) return '0h 0m 0s';
  const total = Math.floor(Math.abs(n));
  const h = Math.floor(total / 3600);
  const m = Math.floor((total % 3600) / 60);
  const s = total % 60;
  return `${h}h ${m}m ${s}s`;
}

// simple func that takes in a sheet and auto resizes all columns that contain values
function resizeColumnsToFit(sheet) {
  const dataRange = sheet.getDataRange();
  if (dataRange.getNumColumns() > 0) {
    sheet.autoResizeColumns(1, dataRange.getNumColumns());
  }
}

/**
 * Formats an ISO 8601 date string to MM/DD/YYYY format.
 * @param {string} dateStr - The ISO date string.
 * @returns {string} The formatted date string.
 */
function formatMeetingDate(dateStr) {
  if (!dateStr) return "";
  var date = new Date(dateStr);
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
  } catch (e) {
    var month = String(date.getUTCMonth() + 1).padStart(2, '0');
    var day = String(date.getUTCDate()).padStart(2, '0');
    var year = date.getUTCFullYear();
    return month + "/" + day + "/" + year;
  }
}

/**
 * Converts input ISO 8601 (UTC) string into a specified locale string (default PT)
 * iso format: '2023-06-08T18:30:00Z'
 * newTimeZone format: 'America/Los_Angeles'
 */
function convertISOTimeZone(iso, newTimeZone = 'America/Los_Angeles') {
  const dt = new Date(iso);
  const converted = dt.toLocaleString('en-US', {
    timeZone: newTimeZone,
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true
  });
  return converted; // e.g. "06/08/2023, 11:30:00 AM"
}

// intended to be used after convertISOTimeZone(), returns only the date
function dateOnly(datetime) {
  const i = datetime.indexOf(' ');
  const date = datetime.slice(0, i);
  return date;
}

// intended to be used after convertISOTimeZone(), returns only the time
function timeOnly(datetime) {
  const i = datetime.indexOf(' ');
  const time = datetime.slice(i + 1);
  return time;
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
