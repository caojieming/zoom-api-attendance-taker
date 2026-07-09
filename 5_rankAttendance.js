// creates a new sheet that shows attendees ranked by attendance rate (# of meetings attended / total # of meetings traversed) + which meetings they attended
function rankAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const activeSheet = ss.getActiveSheet();
  const activeIndex = sheets.findIndex(sh => sh.getSheetId() === activeSheet.getSheetId());

  // Match start-of-string: YYYY/MM/DD, then anything after (time, etc.)
  const startDateRe = /^(\d{4})\/(\d{2})\/(\d{2})/;

  // 1) Sheets to the right of active whose names start with a date
  const eligible = [];
  for (let i = 0; i < sheets.length; i++) {
    if (activeIndex !== -1 && i <= activeIndex) continue;
    const nm = sheets[i].getName().trim();
    if (startDateRe.test(nm)) eligible.push(sheets[i]);
  }

  // Date label is the matching prefix only (MM/DD/YYYY)
  const dates = eligible.map(sh => {
    const m = sh.getName().trim().match(startDateRe);
    return m ? m[0] : sh.getName().trim();
  });
  const totalMeetings = dates.length;

  // 2) Count unique participants per sheet (based on presence anywhere in col A, rows 2+)
  const countByParticipant = {};        // name -> #meetings present
  const presenceByParticipant = {};   // name -> { dateLabel -> true }

  for (let i = 0; i < eligible.length; i++) {
    const sh = eligible[i];
    const dateLabel = dates[i];

    const lastRow = sh.getLastRow();
    if (lastRow < 2) continue;

    const values = sh.getRange(2, 1, lastRow - 1, 1).getValues(); // column A, rows 2+
    const seen = new Set();

    for (const [v] of values) {
      const name = (v ?? "").toString().trim();
      if (name) seen.add(name);
    }

    for (const name of seen) {
      countByParticipant[name] = (countByParticipant[name] || 0) + 1;
      if (!presenceByParticipant[name]) presenceByParticipant[name] = {};
      presenceByParticipant[name][dateLabel] = true;
    }
  }

  // 3) Rank by attendance rate
  const participants = Object.keys(countByParticipant);
  participants.sort((a, b) => {
    const ra = totalMeetings ? (countByParticipant[a] / totalMeetings) : 0;
    const rb = totalMeetings ? (countByParticipant[b] / totalMeetings) : 0;
    if (rb !== ra) return rb - ra;
    return a.localeCompare(b);
  });

  // Overwrite output sheet
  const existing = ss.getSheetByName("Ranked Attendance");
  if (existing) ss.deleteSheet(existing);
  const rankedSheet = ss.insertSheet("Ranked Attendance");

  const header = ["name", "rate", ...dates];
  const rows = participants.map(name => {
    const attended = countByParticipant[name] || 0;
    const rate = totalMeetings ? (attended / totalMeetings) : 0;
    const rateStr = totalMeetings ? (rate * 100).toFixed(2) + "%" : "";

    const row = [name, rateStr];
    for (const dateLabel of dates) {
      row.push(presenceByParticipant[name]?.[dateLabel] ? "present" : "");
    }
    return row;
  });

  rankedSheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) rankedSheet.getRange(2, 1, rows.length, header.length).setValues(rows);

  // add a filter to all columns
  rankedSheet.getRange(1, 1, rows.length + 1, header.length).createFilter();

  // auto resize columns
  resizeColumnsToFit(rankedSheet);
}
