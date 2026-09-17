/* Extra filters */

// Delimiters that indicate a participant name contains extra text
// (e.g. device names, locations, punctuation, etc.)
// Everything from the delimiter onward is removed
const PARTICIPANT_DELIMITERS = [" - ", " (", "iPhone", " | ", " SoCal", ", ", ": ", " SaaS ", "’s iPad"];

// Minimum allowed length for a participant name after cleanup.
const PARTICIPANT_MIN_NAME_LENGTH = 2;

// Names containing any of these substrings are excluded entirely.
const PARTICIPANT_BLACKLIST = ["notetaker", "read.ai"];

// If true, exact duplicate participant names are merged into one row.
const MERGE_DUPES = true;

// If true, similar names are merged using string similarity.
const MERGE_SIMILAR = true;

// Similarity threshold for merging near-duplicate names.
const MERGE_SIMILAR_PERCENTAGE = 0.8;


/**
 * Zoom only returns about 1 month of records per request, so this helper calls the main function multiple times to cover 6 months.
 */
function getParticipantsHalfYear() {
  [
    [ONE_MONTH_AGO, NOW],
    [TWO_MONTHS_AGO, ONE_MONTH_AGO],
    [THREE_MONTHS_AGO, TWO_MONTHS_AGO],
    [FOUR_MONTHS_AGO, THREE_MONTHS_AGO],
    [FIVE_MONTHS_AGO, FOUR_MONTHS_AGO],
    [SIX_MONTHS_AGO, FIVE_MONTHS_AGO],
  ].forEach(([from, to]) => getParticipants(from, to));
}


/**
 * Main function:
 * 1. Fetch historical Zoom meetings in the requested date range
 * 2. Filter out unwanted meetings
 * 3. Fetch participants for each meeting
 * 4. Clean and merge participant names
 * 5. Write the data into a new spreadsheet tab
 */
function getParticipants(inFrom = FROM, inTo = TO) {
  // Get OAuth token for Zoom API access.
  const accessToken = getZoomAccessToken(ACCOUNT_ID, ATTENDANCE_CLIENT_ID, ATTENDANCE_CLIENT_SECRET);

  // Current spreadsheet where new meeting tabs will be created.
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Fetch all meetings from Zoom for the requested date range.
  const meetings = fetchAllMeetings(accessToken, inFrom, inTo);

  // Keep only meetings that match our filter rules.
  const filteredMeetings = meetings.filter(filterMeeting);
  console.log("bleh");

  // Process each meeting one by one.
  filteredMeetings.forEach((meeting) => {
    const rawUuid = meeting.meeting_uuid;

    // use meeting start time as the sheet nameu
    const sheetName = convertISOTimeZone(convertPlainToISO(meeting.start_time) || "");

    // skip if sheet already exists
    if (!sheetName || ss.getSheetByName(sheetName)) return;

    // get all participants for this meeting
    const participants = fetchAllParticipants(accessToken, rawUuid);

    // clean and merge participant records
    const sanitizedParticipants = sanitizeParticipants(participants);

    // skip meetings with 0 or 1 remaining participants after cleanup
    if (sanitizedParticipants.length <= 1) return;

    // insert the new sheet in the correct position
    insertSheetBeforeAttendanceSheets(ss, sheetName);

    // grab the newly created sheet
    const newSheet = ss.getSheetByName(sheetName);

    // write meeting + participant data into the sheet
    writeMeetingSheet(newSheet, meeting, rawUuid, sanitizedParticipants);

    // auto-fit columns for readability
    resizeColumnsToFit(newSheet);
  });
}


/**
 * Fetches all historical meetings from Zoom using pagination.
 */
function fetchAllMeetings(accessToken, from, to) {
  let meetings = [];
  let nextPageToken = "";

  do {
    // Build the request URL for the current page.
    let url =
      "https://api.zoom.us/v2/report/history_meetings" +
      "?from=" + encodeURIComponent(from) +
      "&to=" + encodeURIComponent(to) +
      "&page_size=" + PAGE_SIZE +
      "&meeting_type=" + encodeURIComponent(MEETING_TYPE);

    // Optional search key filter.
    if (SEARCH_KEY) url += "&search_key=" + encodeURIComponent(SEARCH_KEY);

    // Pagination token if this is not the first page.
    if (nextPageToken) url += "&next_page_token=" + encodeURIComponent(nextPageToken);

    // Call Zoom API.
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { Authorization: "Bearer " + accessToken },
      muteHttpExceptions: true,
    });

    // Stop if Zoom returns an error.
    if (response.getResponseCode() !== 200) {
      console.error("Error fetching meetings: " + response.getContentText());
      break;
    }

    // Parse the response and append the current page of meetings.
    const data = JSON.parse(response.getContentText());
    meetings = meetings.concat(data.history_meetings || []);
    nextPageToken = data.next_page_token || "";
  } while (nextPageToken);

  return meetings;
}


/**
 * Fetches all participants for one meeting UUID using pagination.
 */
function fetchAllParticipants(accessToken, rawUuid) {
  let participants = [];
  let nextPageToken = "";
  const encodedUuid = prepareUuid(rawUuid);

  do {
    // Build request URL for participant list.
    let url = `https://api.zoom.us/v2/past_meetings/${encodedUuid}/participants?page_size=300`;

    // Add pagination token if needed.
    if (nextPageToken) url += "&next_page_token=" + encodeURIComponent(nextPageToken);

    // Call Zoom API.
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { Authorization: "Bearer " + accessToken },
      muteHttpExceptions: true,
    });

    // Stop if Zoom returns an error.
    if (response.getResponseCode() !== 200) {
      console.error("Error fetching participants for UUID " + rawUuid + ": " + response.getContentText());
      break;
    }

    // Parse and append this page of participants.
    const data = JSON.parse(response.getContentText());
    participants = participants.concat(data.participants || []);
    nextPageToken = data.next_page_token || "";
  } while (nextPageToken);

  return participants;
}


/**
 * Returns true only if the meeting passes the configured filters.
 */
function filterMeeting(meeting) {
  const meetingId = meeting.meeting_id;

  // If a specific meeting ID is set, only keep that meeting.
  if (MEETING_ID !== "" && meetingId.toString() !== MEETING_ID) return false;

  // Ignore very small meetings.
  if (Number(meeting.participants) <= 5) return false;

  // Optional filter: only include meetings held on the fourth Thursday.
  if (ONLY_FOURTH_THURS && !isFourthThursday(meeting.start_time || "")) return false;

  return true;
}


/**
 * Cleans participant names and merges duplicates/similar entries.
 */
function sanitizeParticipants(participants) {
  const sanitized = [];

  participants.forEach((p) => {
    // Start from the original participant name.
    const originalName = (p.name || "").trim();
    if (!originalName) return;

    // Skip blacklisted names.
    if (
      PARTICIPANT_BLACKLIST.length > 0 &&
      PARTICIPANT_BLACKLIST.some((keyword) =>
        originalName.toLowerCase().includes(keyword.toLowerCase())
      )
    ) {
      return;
    }

    // Remove extra text like device names, locations, etc.
    const name = normalizeParticipantName(originalName);

    // Skip empty or too-short results after cleanup.
    if (!name || name.length < PARTICIPANT_MIN_NAME_LENGTH) return;

    // Create a copy so we don't mutate the original Zoom API object.
    const cur = Object.assign({}, p, {
      name,
      timesRejoined: 0,
    });

    // If merging is disabled, keep every row as-is.
    if (!MERGE_DUPES && !MERGE_SIMILAR) {
      cur.timesRejoined = "disabled";
      sanitized.push(cur);
      return;
    }

    // Try to merge this participant with a previous one.
    let merged = false;
    for (const past of sanitized) {
      // Merge exact duplicates.
      if (MERGE_DUPES && cur.name === past.name) {
        past.leave_time = cur.leave_time;
        past.duration += Number(cur.duration || 0);
        past.timesRejoined += 1;
        merged = true;
        break;
      }

      // Merge near-duplicates using similarity score.
      if (MERGE_SIMILAR && stringSimilarity(cur.name, past.name) >= MERGE_SIMILAR_PERCENTAGE) {
        past.name = cur.name;
        past.leave_time = cur.leave_time;
        past.duration += Number(cur.duration || 0);
        past.timesRejoined += 1;
        merged = true;
        break;
      }
    }

    // If no match was found, keep this participant as a new entry.
    if (!merged) sanitized.push(cur);
  });

  return sanitized;
}


/**
 * Removes trailing extra text from participant names using the configured delimiters.
 */
function normalizeParticipantName(name) {
  let result = name;

  // Apply each delimiter in order.
  for (const delimiter of PARTICIPANT_DELIMITERS) {
    result = result.split(delimiter)[0].trim();
  }

  return result.trim();
}


/**
 * Inserts the new sheet before the first date-named sheet,
 * so attendance sheets stay grouped together.
 */
function insertSheetBeforeAttendanceSheets(ss, sheetName) {
  const sheets = ss.getSheets();

  // Default: insert at the end if no date-named sheet is found.
  let insertAtIndex = sheets.length + 1;

  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName();
    if (DATE_PREFIX.test(name)) {
      insertAtIndex = i;
      break;
    }
  }

  ss.insertSheet(sheetName, insertAtIndex);
}


/**
 * Writes the meeting details and participant table to the sheet.
 */
function writeMeetingSheet(sheet, meeting, rawUuid, participants) {
  const meetingId = meeting.meeting_id;
  const topic = meeting.topic || "Untitled Meeting";
  const duration = Number(meeting.duration || 0);
  const hostDisplayName = meeting.host_display_name || "";
  const hostEmail = meeting.host_email || "";

  // Convert timestamps into the format used in the sheet.
  const startTime = convertPlainToISO(meeting.start_time) || "";
  const endTime = convertPlainToISO(meeting.end_time) || "";
  const convertedStartTime = convertISOTimeZone(startTime);
  const convertedEndTime = convertISOTimeZone(endTime);

  // Column headers for the participant table.
  const participantHeaders = ["name", "join_time", "leave_time", "duration", "rejoined"];

  // Column headers for the meeting metadata block.
  const detailsHeaders = [
    "meeting_uuid",
    "meeting_id",
    "topic",
    "host_display_name",
    "host_email",
    "participants",
    "duration",
    "start_time",
    "end_time",
  ];

  // Convert participant objects into sheet rows.
  const participantRows = participants.map((p) => [
    p.name || "",
    timeOnly(convertISOTimeZone(p.join_time)) || "",
    timeOnly(convertISOTimeZone(p.leave_time)) || "",
    secondsToHMS(Number(p.duration || 0)) || "",
    p.timesRejoined,
  ]);

  // Meeting metadata row.
  const detailsRow = [
    rawUuid,
    meetingId,
    topic,
    hostDisplayName,
    hostEmail,
    participants.length,
    minutesToHM(duration),
    timeOnly(convertedStartTime),
    timeOnly(convertedEndTime),
  ];

  // Write participant headers and rows starting in column A.
  sheet.getRange(1, 1, 1, participantHeaders.length).setValues([participantHeaders]);
  if (participantRows.length > 0) {
    sheet.getRange(2, 1, participantRows.length, participantHeaders.length).setValues(participantRows);
    sheet.getRange(1, 1, participantRows.length + 1, participantHeaders.length).createFilter();
  }

  // Write meeting metadata block starting two columns after participant table.
  sheet.getRange(1, participantHeaders.length + 2, 1, detailsHeaders.length).setValues([detailsHeaders]);
  sheet.getRange(2, participantHeaders.length + 2, 1, detailsRow.length).setValues([detailsRow]);
}


/**
 * Returns decimal similarity between two strings.
 * Example: 0.9 means 90% similar.
 */
function stringSimilarity(s1, s2) {
  const longer = s1.length >= s2.length ? s1 : s2;
  const shorter = s1.length < s2.length ? s1 : s2;
  const longerLength = longer.length;

  if (longerLength === 0) return 1.0;

  return (longerLength - editDistance(longer, shorter)) / longerLength;
}
/**
 * Computes Levenshtein distance between two strings.
 * Lower value = more similar strings.
 */
function editDistance(s1, s2) {
  s1 = s1.toLowerCase();
  s2 = s2.toLowerCase();

  const costs = new Array(s2.length + 1);

  for (let i = 0; i <= s1.length; i++) {
    let lastValue = i;

    for (let j = 0; j <= s2.length; j++) {
      if (i === 0) {
        costs[j] = j;
      } else if (j > 0) {
        let newValue = costs[j - 1];
        if (s1.charAt(i - 1) !== s2.charAt(j - 1)) {
          newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
        }
        costs[j - 1] = lastValue;
        lastValue = newValue;
      }
    }

    if (i > 0) costs[s2.length] = lastValue;
  }

  return costs[s2.length];
}
