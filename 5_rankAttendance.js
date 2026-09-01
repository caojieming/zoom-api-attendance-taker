// if merging similar is enabled, merge participants when one name is directly contained in the other, as long as the shorter name is at least PARTIAL_NAME_THRESHOLD characters long
const PARTIAL_NAME_THRESHOLD = 6;


/**
 * creates a new sheet that shows attendees ranked by attendance rate (# of meetings attended / total # of meetings traversed) + which meetings they attended
 */
function rankAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const activeSheet = ss.getActiveSheet();
  const activeIndex = activeSheet.getIndex();

  // Match start-of-string: YYYY-MM-DD, then anything after (time, etc.)
  const startDateRe = /^(\d{4})-(\d{2})-(\d{2})/;

  // 1) Sheets to the right of active whose names start with a date
  const meetingSheets = [];
  for (let i = 0; i < sheets.length; i++) {
    // getIndex() is 1-based, sheets[] is 0-based
    if (i + 1 <= activeIndex) continue;
    const nm = sheets[i].getName().trim();
    if (startDateRe.test(nm)) meetingSheets.push(sheets[i]);
  }

  // Date label is the matching prefix only (YYYY-MM-DD)
  const dates = meetingSheets.map(sh => {
    const m = sh.getName().trim().match(startDateRe);
    return m ? m[0] : sh.getName().trim();
  });

  // 2) Count meetings attended + store (string) duration per meeting per participant
  // participantData = { name: { attended: #, durationsByDate: { [dateLabel]: durationString } } }
  const participantData = {};
  let totalNumMeetings = 0;

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

    totalNumMeetings++;

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
  if (MERGE_SIMILAR) {
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
  // google sheets is 1 indexed, this ensures that rankedSheet will be the 2nd sheet
  ss.setActiveSheet(rankedSheet);
  ss.moveActiveSheet(2);

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

  if (rows.length) {
    rankedSheet.getRange(1, 1, rows.length + 1, header.length).createFilter();
  }
  resizeColumnsToFit(rankedSheet);
}


/**
 * helper function to merge similar participants
 */
function mergeSimilarParticipants(participants) {
  // Merge participants that satisfy stringSimilarity() >= MERGE_SIMILAR_PERCENTAGE
  // Also merge participants when one name is directly contained in the other, as long as the shorter name is at least PARTIAL_NAME_THRESHOLD characters long.
  // For direct containment, only merge when the shorter name is a whole word or a prefix of the longer name, not just any substring match. Choose the longest name as the main name.
  // merged attended = sum of all attended
  // merged durationsByDate includes all dateLabels (no collisions assumed), sorted in reverse alphabetical order.
  const mergedParticipants = [];
  const used = new Array(participants.length).fill(false);

  for (let i = 0; i < participants.length; i++) {
    if (used[i]) continue;

    const groupIdx = [i];
    used[i] = true;

    // Expand the group until no more similar participants are found
    for (let g = 0; g < groupIdx.length; g++) {
      const baseIdx = groupIdx[g];

      for (let j = i + 1; j < participants.length; j++) {
        if (used[j]) continue;

        const nameA = participants[baseIdx].name;
        const nameB = participants[j].name;

        // Direct containment merge:
        // only merge if the shorter name has at least PARTIAL_NAME_THRESHOLD characters
        // and is either a whole word or a prefix of the longer name
        const shorter = nameA.length <= nameB.length ? nameA : nameB;
        const longer = nameA.length > nameB.length ? nameA : nameB;

        const shorterLen = shorter.length;
        const wholeWord =
          new RegExp(`(^|\\s)${escapeRegExp(shorter)}($|\\s)`).test(longer);
        const prefix = longer.startsWith(shorter);

        const directlyContains = shorterLen >= PARTIAL_NAME_THRESHOLD && (wholeWord || prefix);

        const similar =
          stringSimilarity(nameA, nameB) >= MERGE_SIMILAR_PERCENTAGE;

        if (directlyContains || similar) {
          groupIdx.push(j);
          used[j] = true;
        }
      }
    }

    // Pick representative name (participant with the longest name; ties keep first encountered)
    let repIdx = groupIdx[0];
    for (let g = 1; g < groupIdx.length; g++) {
      const idx = groupIdx[g];
      if (participants[idx].name.length > participants[repIdx].name.length) {
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

/**
 * Escape special regex characters in a string.
 */
function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

