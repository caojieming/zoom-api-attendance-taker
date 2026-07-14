// creates a new sheet that shows attendees ranked by attendance rate (# of meetings attended / total # of meetings traversed) + which meetings they attended

function rankAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const activeSheet = ss.getActiveSheet();
  const activeIndex = sheets.findIndex(sh => sh.getSheetId() === activeSheet.getSheetId());

  // Match start-of-string: YYYY/MM/DD, then anything after (time, etc.)
  const startDateRe = /^(\d{4})\/(\d{2})\/(\d{2})/;

  // 1) Sheets to the right of active whose names start with a date
  const meetingSheets = [];
  for (let i = 0; i < sheets.length; i++) {
    if (activeIndex !== -1 && i <= activeIndex) continue;
    const nm = sheets[i].getName().trim();
    if (startDateRe.test(nm)) meetingSheets.push(sheets[i]);
  }

  // Date label is the matching prefix only (MM/DD/YYYY)
  const dates = meetingSheets.map(sh => {
    const m = sh.getName().trim().match(startDateRe);
    return m ? m[0] : sh.getName().trim();
  });
  const totalNumMeetings = dates.length;

  // 2) Count meetings attended + store (string) duration per meeting per participant
  // participantData = { name: { attended: #, durationsByDate: { [dateLabel]: durationString } } }
  const participantData = {};

  for (let i = 0; i < meetingSheets.length; i++) {
    const sh = meetingSheets[i];
    const dateLabel = dates[i];

    const lastRow = sh.getLastRow();
    if (lastRow < 2) continue;

    const lastCol = sh.getLastColumn();
    if (lastCol < 1) continue;

    // Find "duration" column in header row (row 1) via substring match
    const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(x => (x ?? "").toString().trim());

    const durationColIndex = headerRow.findIndex(h => h === "duration");
    if (durationColIndex === -1) continue;

    const numRows = lastRow - 1;

    const nameValues = sh.getRange(2, 1, numRows, 1).getValues(); // col A
    const durationValues = sh.getRange(2, durationColIndex + 1, numRows, 1).getValues(); // "duration" col

    // go through each participant in the current sheet/meeting
    for (let r = 0; r < numRows; r++) {
      let name = (nameValues[r][0] ?? "").toString().trim(); // don't use .toLowerCase()

      // should not ever happen, but just in case
      if (!name) continue;

      // sanitize names, removing extra info that is not name related
      name = name.split(" - ")[0].trim();
      name = name.split(" (")[0].trim();

      const durStr = (durationValues[r][0] ?? "").toString().trim();

      if (!participantData[name]) {
        participantData[name] = { attended: 0, durationsByDate: {} };
      }

      participantData[name].attended += 1;
      participantData[name].durationsByDate[dateLabel] = durStr; // store duration as-is because it's not a number
    }
  }

  // Convert keyed object to an array of participant objects (makes merging similar dupes easier)
  // participant objects: { name, attended, durationsByDate }
  let participants = Object.keys(participantData).map(name => ({
    name,
    attended: participantData[name].attended,
    durationsByDate: participantData[name].durationsByDate
  }));

  // merge similar names
  if(MERGE_SIMILAR) {
    participants = mergeSimilarParticipants(participants);
  }

  // 3) Rank by attendance rate (no tie-break beyond name)
  participants.sort((p, q) => {
    const ra = totalNumMeetings ? (p.attended / totalNumMeetings) : 0;
    const rb = totalNumMeetings ? (q.attended / totalNumMeetings) : 0;
    if (rb !== ra) return rb - ra;
    return p.name.localeCompare(q.name);
  });

  // 4) Overwrite output sheet
  const existing = ss.getSheetByName("Ranked Attendance");
  if (existing) ss.deleteSheet(existing);
  const rankedSheet = ss.insertSheet("Ranked Attendance");

  const header = ["name", "rate", ...dates];

  const rows = participants.map(p => {
    const attended = p.attended || 0;
    const rate = totalNumMeetings ? (attended / totalNumMeetings) : 0;
    const rateStr = totalNumMeetings ? (rate * 100).toFixed(2) + "%" : "";

    const row = [p.name, rateStr];
    for (const dateLabel of dates) {
      const val = p.durationsByDate?.[dateLabel];
      row.push(val === undefined || val === null || val === "" ? "" : val);
    }
    return row;
  });

  rankedSheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) rankedSheet.getRange(2, 1, rows.length, header.length).setValues(rows);

  rankedSheet.getRange(1, 1, rows.length + 1, header.length).createFilter();
  resizeColumnsToFit(rankedSheet);
}


// helper function to merge similar participants
function mergeSimilarParticipants(participants) {
  // Merge participants that satisfy stringSimilarity() >= MERGE_SIMILAR_PERCENTAGE
  // choose merged name from participant with highest attended value
  // merged attended = sum of all attended
  // merged durationsByDate includes all dateLabels (no collisions assumed), sorted in reverse alphabetical order.
  const mergedParticipants = [];
  const used = new Array(participants.length).fill(false);

  for (let i = 0; i < participants.length; i++) {
    if (used[i]) continue;

    // Build a merge group containing i + any other participant j where any member x in the current group satisfies stringSimilarity(x, j) >= MERGE_SIMILAR_PERCENTAGE
    const groupIdx = [i];
    used[i] = true;

    // get all participants within a certain % of similarity of participants[i]
    for (let j = i + 1; j < participants.length; j++) {
      if (used[j]) continue;

      // x == i
      const shouldMerge = groupIdx.some(x => (stringSimilarity(participants[x].name, participants[j].name) >= MERGE_SIMILAR_PERCENTAGE));
      if (!shouldMerge) continue;

      groupIdx.push(j);
      used[j] = true;
    }

    // Pick representative name (participant with higher attended; ties keep first encountered)
    let repIdx = groupIdx[0];
    for (let g = 1; g < groupIdx.length; g++) {
      const idx = groupIdx[g];
      if (participants[idx].attended > participants[repIdx].attended) {
        repIdx = idx;
      }
    }

    // Combine attended and durations
    let mergedAttended = 0;
    const combinedDurations = {};

    for (const idx of groupIdx) {
      mergedAttended += participants[idx].attended || 0;
      Object.assign(combinedDurations, participants[idx].durationsByDate || {});
    }

    // Sort dateLabel keys in reverse alphabetical order (sort dates from recent to oldest)
    const sortedDateLabels = Object.keys(combinedDurations).sort((u, v) => v.localeCompare(u));
    const sortedDurationsByDate = {};
    for (const dateLabel of sortedDateLabels) {
      sortedDurationsByDate[dateLabel] = combinedDurations[dateLabel];
    }

    mergedParticipants.push({
      name: participants[repIdx].name,
      attended: mergedAttended,
      durationsByDate: sortedDurationsByDate
    });
  }

  return mergedParticipants;
}
