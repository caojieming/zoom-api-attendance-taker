// The most up to date version of this code
// This is all untested code. Need to fill in credentials in globals.gs with the Zoom App credentials

const NOW = new Date().toISOString().split('T')[0]; // Today (YYYY-MM-DD)
const ONE_WEEK_AGO = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]; // One week ago (YYYY-MM-DD)
const ONE_MONTH_AGO = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

const FROM = ONE_WEEK_AGO;
const TO = NOW;
const MEETING_TYPE = "meeting"; // Type of meeting (meeting or webinar)
const SEARCH_KEY = ""; // Optional search query key if you only want meetings with specific word(s) in the topic name
const PAGE_SIZE = 50; // Max meetings per page (up to 300)

/**
 * Main function to get historical Zoom meetings within the past week,
 * extract details and participant lists, and write them into Google Sheets.
 */
function getAndWritePastMeetingParticipants() {
  // Fetch access token using existing client function (assumed to be defined globally)
  var accessToken = getZoomAccessToken();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var meetings = [];
  var nextMeetingPageToken = "";

  // 1. Paginated fetch of past meetings from the Zoom report API
  do {
    var url = "https://api.zoom.us/v2" + "/report/history_meetings" +
      "?from=" + FROM +
      "&to=" + TO +
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
    var rawUuid = meeting.meeting_uuid;
    var meetingId = meeting.meeting_id;
    var topic = meeting.topic || "Untitled Meeting";
    var startTime = meeting.start_time || "";
    var endTime = meeting.end_time || "";
    var duration = meeting.duration || 0;
    var hostDisplayName = meeting.user_name || "";
    var hostEmail = meeting.user_email || "";

    // Format sheet name: MM/DD/YYYY, Topic
    var dateString = formatMeetingDate(startTime);
    var sanitizedTopic = topic.replace(/[\\?\*:\[\]]/g, "-"); // Replace forbidden sheet characters
    var sheetName = dateString + ", " + sanitizedTopic + ", " + rawUuid;
    if (sheetName.length > 100) {
      sheetName = sheetName.substring(0, 100);
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

    // Count unique viewers based on email/ID/name
    var uniqueKeys = new Set();
    participants.forEach(function (p) {
      var key = p.user_email || p.id || p.name;
      if (key) {
        uniqueKeys.add(key.toString().toLowerCase().trim());
      }
    });
    var uniqueViewersCount = uniqueKeys.size;
    var totalParticipantsCount = participants.length;

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
      "unique_viewers",
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
      uniqueViewersCount,
      duration,
      startTime,
      endTime
    ];

    newSheet.getRange(1, 8, 1, detailsHeaders.length).setValues([detailsHeaders]);
    newSheet.getRange(2, 8, 1, detailsRow.length).setValues([detailsRow]);

    // 4. Generate and place the participants table next (starts at column A / 1)
    var participantHeaders = [
      "id",
      "name",
      "user_email",
      "join_time",
      "leave_time",
      "duration"
    ];

    newSheet.getRange(1, 1, 1, participantHeaders.length).setValues([participantHeaders]);

    if (participants.length > 0) {
      var participantRows = participants.map(function (p) {
        return [
          p.id || "",
          p.name || "",
          p.user_email || "",
          p.join_time || "",
          p.leave_time || "",
          p.duration || 0
        ];
      });
      newSheet.getRange(2, 1, participantRows.length, participantHeaders.length).setValues(participantRows);
    }

    // auto resize columns
    const dataRange = newSheet.getDataRange();
    if (dataRange.getNumColumns() > 0) {
      newSheet.autoResizeColumns(1, dataRange.getNumColumns());
    }

  });
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
