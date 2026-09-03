// can't get more than 1 month worth of records at a time, need to call multiple times
function getRecordingsHalfYear() {
  getRecordings(ONE_MONTH_AGO, NOW);
  getRecordings(TWO_MONTHS_AGO, ONE_MONTH_AGO);
  getRecordings(THREE_MONTHS_AGO, TWO_MONTHS_AGO);
  getRecordings(FOUR_MONTHS_AGO, THREE_MONTHS_AGO);
  getRecordings(FIVE_MONTHS_AGO, FOUR_MONTHS_AGO);
  getRecordings(SIX_MONTHS_AGO, FIVE_MONTHS_AGO);
}


/**
 * Main function to get historical Zoom meetings within the past week,
 * extract transcripts and recordings, and write them into Google drive folder.
 * dateFrom: start date of time period observed, defaulting to const FROM
 * dateTo: end date of time period observed, defaulting to const TO
 */
function getRecordings(dateFrom = FROM, dateTo = TO) {
  const accessToken = getZoomAccessToken(ACCOUNT_ID, RECORDINGS_CLIENT_ID, RECORDINGS_CLIENT_SECRET);
  const existingFilenames = getExistingFilenameSet(DRIVE_FOLDER_ID);

  const meetings = fetchHistoricalMeetings(accessToken, dateFrom, dateTo);

  const filteredMeetings = meetings.filter(function (meeting) {

    // optional check/filter for meeting ID
    // if const MEETING_ID is filled/is not empty AND current meeting ID does not equal const MEETING_ID
    const meetingId = meeting.meeting_id.toString();
    if(MEETING_ID !== "" && meetingId !== MEETING_ID) {
      return false;
    }

    // skip false/empty meetings (I assume meetings comprised of 5 or less total participants aren't meetings we're interested in)
    if(Number(meeting.participants) <= 5) {
      return false;
    }

    // optional check/filter if only looking for meetings on the fourth thursday of the month
    const startTime = meeting.start_time || "";
    if(ONLY_FOURTH_THURS && !isFourthThursday(startTime)) {
      return false;
    }

    return true;
  });

  // 2. Iterate through filtered list of meetings for recordings
  filteredMeetings.forEach(function (meeting) {
    const startTime = convertPlainToISO(meeting.start_time) || "";
    const datetime = convertISOTimeZone(startTime);

    const meetingUuid = meeting.meeting_uuid;
    const meetingUuidEncoded = prepareUuid(meetingUuid);

    // GET /meetings/{meetingId}/recordings
    const recordingsUrl = `https://api.zoom.us/v2/meetings/${meetingUuidEncoded}/recordings`;
    const summaryUrl = `https://api.zoom.us/v2/meetings/${meetingUuidEncoded}/meeting_summary`;

    const rawRecordingsData = httpGetData(recordingsUrl, accessToken, datetime);

    // check code from httpGetData
    if(rawRecordingsData.code === 200) {
      const recordingsData = JSON.parse(rawRecordingsData.data);

      // loop through all recording_files, searching for file_type = TRANSCRIPT, MP4, CHAT -> place them in drive (see if they can be placed raw or as original files)
      (recordingsData.recording_files || []).forEach(function(file) {
        if(file.file_type === "TRANSCRIPT") {
          const fileName = datetime + " [Transcript]";
          // Skip if a file with this name already exists in the folder
          if (existingFilenames.has(fileName)) {
            console.log("File already imported: " + fileName);
          }
          else {
            const downloadUrl = file.download_url;
            console.log("Downloading/Importing: " + fileName);
            const rawTranscript = httpGetData(downloadUrl, accessToken);
            let transcript = rawTranscript.data;
            // just removes the excessive number of extra newlines in the transcript
            transcript = transcript.replace(/\n/g, '');
            createGoogleDocInFolder(DRIVE_FOLDER_ID, fileName, transcript);
            existingFilenames.add(fileName);

            // create key topics sheet from transcript if it doesn't already exist
            const sheetName = `${datetime} [Key Topics]`;
            if(!existingFilenames.has(sheetName)) {
              console.log("Generating: " + sheetName);
              createKeyTopicsSheet(DRIVE_FOLDER_ID, sheetName, transcript);
              existingFilenames.add(sheetName);
            }
            else {
              console.log("File already generated: " + sheetName);
            }
          }
        }
        else if(file.file_type === "CHAT") {
          const fileName = datetime + " [Chat Log]";
          // Skip if a file with this name already exists in the folder
          if (existingFilenames.has(fileName)) {
            console.log("File already imported: " + fileName);
          }
          else {
            const downloadUrl = file.download_url;
            console.log("Downloading/Importing: " + fileName);
            const rawChatLog = httpGetData(downloadUrl, accessToken);
            const chatLog = rawChatLog.data;
            createGoogleDocInFolder(DRIVE_FOLDER_ID, fileName, chatLog);
            existingFilenames.add(fileName);

            // create chat participants sheet from chat logs if it doesn't already exist
            const sheetName = `${datetime} [Chat Participants]`;
            if(!existingFilenames.has(sheetName)) {
              console.log("Generating: " + sheetName);
              createChatParticipantsSheet(DRIVE_FOLDER_ID, sheetName, chatLog);
              existingFilenames.add(sheetName);
            }
            else {
              console.log("File already generated: " + sheetName);
            }
          }
        }
      });
    }

    const rawSummaryData = httpGetData(summaryUrl, accessToken, datetime);

    // check code from httpGetData
    if(rawSummaryData.code === 200) {
      const fileName = datetime + " [Summary]";
      // Skip if a file with this name already exists in the folder
      if (existingFilenames.has(fileName)) {
        console.log("File already imported: " + fileName);
      }
      else {
        const summaryData = JSON.parse(rawSummaryData.data);
        const summaryText = summaryData.summary_content;
        console.log("Downloading/Importing: " + fileName);
        createGoogleDocInFolder(DRIVE_FOLDER_ID, fileName, summaryText);
        existingFilenames.add(fileName);
      }
    }

    // small timeout to prevent very specific errors
    // Utilities.sleep(200);
  });

  // moves all loose files into date folders
  organizeFilesByDate(DRIVE_FOLDER_ID);
}


/**
 * Loops through all files in rootFolderId, and if their name starts with a date in the format "YYYY-MM-DD", then the function moves that file into the folder with the same date name.
 * If that date folder doesn't exist, the function creates it before moving the file into it
 */
function organizeFilesByDate(rootFolderId) {
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  const files = rootFolder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();

    // Match files that start with YYYY-MM-DD
    const match = name.match(/^(\d{4}-\d{2}-\d{2})/);
    if (!match) continue;

    const dateFolderName = match[1];

    // Find or create the date folder inside the root folder
    const folderIterator = rootFolder.getFoldersByName(dateFolderName);
    const dateFolder = folderIterator.hasNext()
      ? folderIterator.next()
      : rootFolder.createFolder(dateFolderName);

    // Skip if file is already in the correct folder
    const parents = file.getParents();
    let alreadyInFolder = false;
    while (parents.hasNext()) {
      if (parents.next().getId() === dateFolder.getId()) {
        alreadyInFolder = true;
        break;
      }
    }
    if (alreadyInFolder) continue;

    // Move file into the date folder
    dateFolder.addFile(file);
    rootFolder.removeFile(file);
  }
}


/**
 * Fetch historical meetings with pagination.
 */
function fetchHistoricalMeetings(accessToken, inFrom, inTo) {
  const meetings = [];
  let nextMeetingPageToken = "";

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

    const response = UrlFetchApp.fetch(meetingsUrl, {
      method: "get",
      headers: {
        "Authorization": "Bearer " + accessToken
      },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.history_meetings && data.history_meetings.length > 0) {
        meetings.push.apply(meetings, data.history_meetings);
      }
      nextMeetingPageToken = data.next_page_token || "";
    } else {
      console.error("Error fetching historical meetings. Code: " + responseCode + ", Response: " + response.getContentText());
      nextMeetingPageToken = ""; // Stop pagination on error
    }
  } while (nextMeetingPageToken);

  return meetings;
}


/**
 * Returns a Set of all file names already in the target folder.
 * also recursively traverses all subfolders for file names
 */
function getExistingFilenameSet(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const names = new Set();

  function walkFolder(currentFolder) {
    const files = currentFolder.getFiles();
    while (files.hasNext()) {
      names.add(files.next().getName());
    }

    const subfolders = currentFolder.getFolders();
    while (subfolders.hasNext()) {
      walkFolder(subfolders.next());
    }
  }

  walkFolder(folder);
  return names;
}


/**
 * generic get data from a link/API endpoint (use if you don't care for custom error messages/actions)
 */
function httpGetData(url, accessToken, id = "") {
  const options = {
    method: "get",
    headers: {
      "Authorization": "Bearer " + accessToken
    },
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  const data = res.getContentText();
  if (code < 200 || code >= 300) {
    // throw new Error(`HTTP code ${code} for ${url}, data: ${data}`);
    console.error(`[${id}]  HTTP code ${code}, data: ${data}`);
  }
  return { code: code, data: data };
}


/**
 * creates a google doc with a specific name, specific content, and in a specific drive folder
 */
function createGoogleDocInFolder(driveFolderId, docName, docContent) {
  const folder = DriveApp.getFolderById(driveFolderId);

  const doc = DocumentApp.create(docName);
  doc.getBody().setText(docContent);

  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(folder); // puts the file in that folder

  return doc.getId();
}
