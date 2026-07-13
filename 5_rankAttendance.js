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
  // # of meetings attended by each participant. { name: #_meetings_attended }
  const countByParticipant = {};
  // duration value for each meeting for each participant. { name: { dateLabel: durationString } }
  const durationByParticipant = {};

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
      countByParticipant[name] = (countByParticipant[name] || 0) + 1;

      if (!durationByParticipant[name]) durationByParticipant[name] = {};
      durationByParticipant[name][dateLabel] = durStr; // store duration as-is because it's not a number
    }
  }

  // merge similar names
  const participants = Object.keys(countByParticipant);
  // TODO: CONTINUE WORKING HERE

  // 3) Rank by attendance rate (no tie-break)
  participants.sort((a, b) => {
    const ra = totalNumMeetings ? (countByParticipant[a] / totalNumMeetings) : 0;
    const rb = totalNumMeetings ? (countByParticipant[b] / totalNumMeetings) : 0;
    if (rb !== ra) return rb - ra;
    return a.localeCompare(b);
  });

  // 4) Overwrite output sheet
  const existing = ss.getSheetByName("Ranked Attendance");
  if (existing) ss.deleteSheet(existing);
  const rankedSheet = ss.insertSheet("Ranked Attendance");

  const header = ["name", "rate", ...dates];

  const rows = participants.map(name => {
    const attended = countByParticipant[name] || 0;
    const rate = totalNumMeetings ? (attended / totalNumMeetings) : 0;
    const rateStr = totalNumMeetings ? (rate * 100).toFixed(2) + "%" : "";

    const row = [name, rateStr];
    for (const dateLabel of dates) {
      const val = durationByParticipant[name]?.[dateLabel];
      row.push(val === undefined || val === null || val === "" ? "" : val);
    }
    return row;
  });

  rankedSheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) rankedSheet.getRange(2, 1, rows.length, header.length).setValues(rows);

  rankedSheet.getRange(1, 1, rows.length + 1, header.length).createFilter();
  resizeColumnsToFit(rankedSheet);
}
