const fullMemberListCols = ["FirstName", "LastName", "Email Primary", "Phone Work"];
function getFullMemberList() {
  copyColumnsFromSpreadsheet(fullMemberListCols, FULL_MEMBER_LIST_SPREADSHEET_ID, "Full Member List", "Full Member List");
}


function getAttendanceRanker() {
  copyColumnsFromSpreadsheet(null, ATTENDANCE_RANKER_SPREADSHEET_ID, "Ranked Attendance", "Ranked Attendance");

  const attendanceRankerCols = ["name", "matched_name", ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Ranked Attendance");
  if (!sheet) throw new Error('Sheet "Ranked Attendance" not found.');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const displayHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

  // Find the 2 leftmost date-header columns.
  const dateHeaderRegex = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
  const dateCols = [];
  for (let i = 0; i < displayHeaders.length; i++) {
    if (dateHeaderRegex.test(String(displayHeaders[i]).trim())) {
      dateCols.push(i);
      if (dateCols.length === 2) break;
    }
  }
  if (dateCols.length < 2) {
    throw new Error("Could not find 2 date columns in the header row.");
  }

  // Remove rows that don't have a filled cell in either of the 2 date columns.
  for (let r = data.length - 1; r >= 1; r--) {
    const v1 = data[r][dateCols[0]];
    const v2 = data[r][dateCols[1]];
    const hasValue =
      (v1 !== "" && v1 !== null) ||
      (v2 !== "" && v2 !== null);

    if (!hasValue) {
      sheet.deleteRow(r + 1);
    }
  }

  // Re-read headers after row deletions.
  const updatedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Remove columns whose headers aren't in attendanceRankerCols.
  for (let c = updatedHeaders.length - 1; c >= 0; c--) {
    if (!attendanceRankerCols.includes(String(updatedHeaders[c]).trim())) {
      sheet.deleteColumn(c + 1);
    }
  }


  // resize columns
  resizeColumnsToFit_(sheet);
}


function compareFullToEmail() {
  mergeTwoSheets("Full Member List", "Email List", "Email Primary", "email", "Compare Full to Email");
}


function compareAttendanceToEmail() {
  mergeTwoSheets("Attendance Ranker", "Email List", "Email Primary", "email", "Compare Full to Email");
}
