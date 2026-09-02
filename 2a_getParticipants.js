/* Extra filters */

// if participant name has any of these phrases, cut off everything from this point onwards (including the phrase)
const PARTICIPANT_DELIMITERS = [" - ", " (", "iPhone", " | ", " SoCal", ", ", ": ", " SaaS ", "’s iPad"];

// if participant name is less than this length, exclude them
const PARTICIPANT_MIN_NAME_LENGTH = 2;

// if participant name has any of these words, exclude them from the sheet
const PARTICIPANT_BLACKLIST = ['notetaker', 'read.ai'];


// merge participant entries with exact same name
const MERGE_DUPES = true;
// merge similar participant entries within a certain margin of error
const MERGE_SIMILAR = true;
const MERGE_SIMILAR_PERCENTAGE = 0.8;



/**
 * can't get more than 1 month worth of records at a time, need to call multiple times
 */
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
  const accessToken = getZoomAccessToken(ACCOUNT_ID, ATTENDANCE_CLIENT_ID, ATTENDANCE_CLIENT_SECRET);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let meetings = [];
  let nextMeetingPageToken = "";

  // 1. Paginated fetch of past meetings from the Zoom report API
  do {
    let meetingsUrl = "https://api.zoom.us/v2/report/history_meetings" +
      "?from=" + inFrom +
      "&to=" + inTo +
      "&page_size=" + PAGE_SIZE +
      "&meeting_type=" + MEETING_TYPE;

    if (SEARCH_KEY) {
      meetingsUrl += "&search_key=" + encodeURIComponent(SEARCH_KEY);
    }
    if (nextMeetingPageToken) {
      meetingsUrl += "&next_page_token=" + encodeURIComponent(nextMeetingPageToken);
    }

    const options = {
      method: "get",
      headers: {
        "Authorization": "Bearer " + accessToken
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(meetingsUrl, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
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


  // Filter for wanted meetings
  const filteredMeetings = [];
  meetings.forEach(function (meeting) {
    const meetingId = meeting.meeting_id;

    // optional check/filter for meeting ID
    // if const MEETING_ID is filled/is not empty AND current meeting ID does not equal const MEETING_ID (note, use soft inequality check: meetingId is apparently not a string)
    if(MEETING_ID !== "" && meetingId != MEETING_ID) {
      return;
    }

    // skip false/empty meetings (I assume meetings comprised of 5 or less total participants aren't meetings we're interested in)
    if(Number(meeting.participants) <= 5) {
      return;
    }

    const startTime = meeting.start_time || "";
    // optional check/filter if only looking for meetings on the fourth thursday of the month
    if(ONLY_FOURTH_THURS && !isFourthThursday(startTime)) {
      return;
    }

    // passed all filters, add meeting to filteredMeetings list
    filteredMeetings.push(meeting);
  });


  // 2. Iterate through each meeting and build the spreadsheets
  filteredMeetings.forEach(function (meeting) {
    const meetingId = meeting.meeting_id;
    const startTime = convertPlainToISO(meeting.start_time) || "";
    const endTime = convertPlainToISO(meeting.end_time) || "";
    const rawUuid = meeting.meeting_uuid;
    const topic = meeting.topic || "Untitled Meeting";
    const duration = meeting.duration || 0;
    const hostDisplayName = meeting.host_display_name || "";
    const hostEmail = meeting.host_email || "";

    const convertedStartTime = convertISOTimeZone(startTime);
    const convertedEndTime = convertISOTimeZone(endTime);

    // Format sheet name: YYYY-MM-DD, HH:mm:ss am/pm
    const sheetName = convertedStartTime;

    // Skip if a sheet with this name already exists
    if (ss.getSheetByName(sheetName)) {
      console.log("Sheet '" + sheetName + "' already exists. UUID: " + rawUuid + ". Skipping.");
      return;
    }

    // Fetch all participants for this meeting UUID
    let participants = [];
    let participantNextPageToken = "";
    const encodedUuid = prepareUuid(rawUuid);

    do {
      // don't bother with much with page_size here as it's much less likely to hit rate limits than meetings
      let participantsUrl = `https://api.zoom.us/v2/past_meetings/${encodedUuid}/participants` +
        "?page_size=" + 300;

      if (participantNextPageToken) {
        participantsUrl += "&next_page_token=" + encodeURIComponent(participantNextPageToken);
      }

      const partOptions = {
        method: "get",
        headers: {
          "Authorization": "Bearer " + accessToken
        },
        muteHttpExceptions: true
      };

      const partResponse = UrlFetchApp.fetch(participantsUrl, partOptions);
      const partResponseCode = partResponse.getResponseCode();

      if (partResponseCode === 200) {
        const partData = JSON.parse(partResponse.getContentText());
        if (partData.participants && partData.participants.length > 0) {
          participants = participants.concat(partData.participants);
        }
        participantNextPageToken = partData.next_page_token;
      } else {
        console.error("Error fetching participants for UUID: " + rawUuid + ". Code: " + partResponseCode + ", Response: " + partResponse.getContentText());
        participantNextPageToken = ""; // Stop pagination on error
      }
    } while (participantNextPageToken);

    // sanitize participants list (optionally remove Notetakers, remove chapter names, merge dupe names)
    const sanitizedParticipants = [];
    participants.forEach(function (curParticipant) {
      // if any of the blacklist substrings are in the curParticipant name, skip
      if(PARTICIPANT_BLACKLIST.length > 0 && PARTICIPANT_BLACKLIST.some(keyword => curParticipant.name.toLowerCase().includes(keyword.toLowerCase()))) {
        return;
      }

      // remove chapter names or other similar extra non-useful info in name
      PARTICIPANT_DELIMITERS.forEach(function (delimiter) {
        curParticipant.name = curParticipant.name.split(delimiter)[0].trim();
      });
      if(curParticipant.name === '') {
        // after trimming, empty string -> no useful info, so skip
        return;
      }

      // remove names that are of a certain length or shorter
      if(curParticipant.name.length < PARTICIPANT_MIN_NAME_LENGTH) {
        return;
      }

      // dupe checks
      if(MERGE_DUPES || MERGE_SIMILAR){

        let merged = false;
        for (const pastParticipant of sanitizedParticipants) {
          if (MERGE_DUPES && curParticipant.name === pastParticipant.name) {
            pastParticipant.leave_time = curParticipant.leave_time;
            pastParticipant.duration += curParticipant.duration;
            pastParticipant.timesRejoined += 1;
            merged = true;
            break;
          }

          if (MERGE_SIMILAR && stringSimilarity(curParticipant.name, pastParticipant.name) >= MERGE_SIMILAR_PERCENTAGE) {
            pastParticipant.name = curParticipant.name;
            pastParticipant.leave_time = curParticipant.leave_time;
            pastParticipant.duration += curParticipant.duration;
            pastParticipant.timesRejoined += 1;
            merged = true;
            break;
          }
        }

        if (!merged) {
          curParticipant.timesRejoined = 0;
          sanitizedParticipants.push(curParticipant);
        }
      }
      // allow any dupes, so just add participant
      else {
        // neither merge toggle enabled, so just set timesRejoined = "disabled"
        curParticipant.timesRejoined = "disabled";
        sanitizedParticipants.push(curParticipant);
      }

    });
    const totalParticipantsCount = sanitizedParticipants.length;

    // if totalParticipantsCount = 0, then all participants were blacklisted: skip this meeting
    // if totalParticipantsCount = 1, then it can hardly by called a meeting: skip this meeting
    if(totalParticipantsCount <= 1) {
      return;
    }


    // make sure to only insert sheets after [Base] and Ranked Attendance and before other attendance sheets
    const datePrefix = /^(\d{4})-(\d{2})-(\d{2})/;
    while (true) {
      const sheets = ss.getSheets();
      const activeSheet = ss.getActiveSheet();
      const activeIndex = activeSheet.getIndex(); // 1-based

      if (activeIndex >= sheets.length) break; // already last sheet

      const rightSheet = sheets[activeIndex]; // directly to the right
      const rightName = rightSheet.getName();

      if (datePrefix.test(rightName)) break;

      rightSheet.activate();
    }

    // Insert new sheet for the meeting
    const newSheet = ss.insertSheet(sheetName);
    

    // 3. Generate and place the meeting details table first (starts at column H / 8)
    const detailsHeaders = [
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

    const detailsRow = [
      rawUuid,
      meetingId,
      topic,
      hostDisplayName,
      hostEmail,
      totalParticipantsCount,
      minutesToHM(duration),
      timeOnly(convertedStartTime),
      timeOnly(convertedEndTime)
    ];


    // 4. Generate and place the participants table next (starts at column A / 1)
    const participantHeaders = [
      // "id",
      "name",
      // "user_email",
      "join_time",
      "leave_time",
      // "duration_sec",
      "duration",
      "rejoined"
    ];

    const participantRows = sanitizedParticipants.map(function (p) {
      return [
        // p.id || "",
        p.name || "",
        // p.user_email || "",
        timeOnly(convertISOTimeZone(p.join_time)) || "",
        timeOnly(convertISOTimeZone(p.leave_time)) || "",
        // p.duration || 0,
        secondsToHMS(p.duration) || 0,
        p.timesRejoined
      ];
    });


    // getRange(row, col, num rows, num cols)
    // set meeting details
    newSheet.getRange(1, participantHeaders.length + 2, 1, detailsHeaders.length).setValues([detailsHeaders]);
    newSheet.getRange(2, participantHeaders.length + 2, 1, detailsRow.length).setValues([detailsRow]);

    // set participant details
    newSheet.getRange(1, 1, 1, participantHeaders.length).setValues([participantHeaders]);
    newSheet.getRange(2, 1, participantRows.length, participantHeaders.length).setValues(participantRows);

    // auto resize columns
    resizeColumnsToFit(newSheet);

    // add a filter to columns A to F
    newSheet.getRange(1, 1, participantRows.length + 1, participantHeaders.length).createFilter();

  });
}



/**
 * function that returns the decimal/percentage similarity between 2 strings (example: 0.9 = 90% similarity between 2 strings)
 */
function stringSimilarity(s1, s2) {
  let longer = s1;
  let shorter = s2;
  if (s1.length < s2.length) {
    longer = s2;
    shorter = s1;
  }
  const longerLength = longer.length;
  if (longerLength === 0) {
    return 1.0;
  }
  return (longerLength - editDistance(longer, shorter)) / parseFloat(longerLength);
}
/**
 * function that returns the Levenshtein distance of 2 strings (aka the # of edits needed to make string1 into string2)
 */
function editDistance(s1, s2) {
  s1 = s1.toLowerCase();
  s2 = s2.toLowerCase();

  const costs = new Array();
  for (let i = 0; i <= s1.length; i++) {
    let lastValue = i;
    for (let j = 0; j <= s2.length; j++) {
      if (i === 0)
        costs[j] = j;
      else {
        if (j > 0) {
          let newValue = costs[j - 1];
          if (s1.charAt(i - 1) !== s2.charAt(j - 1))
            newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
          costs[j - 1] = lastValue;
          lastValue = newValue;
        }
      }
    }
    if (i > 0)
      costs[s2.length] = lastValue;
  }
  return costs[s2.length];
}
