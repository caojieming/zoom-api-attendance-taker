const fullMemberListCols = ["FirstName", "LastName", "Email Primary", "Phone Work"];
function getFullMemberList() {
  copyColumnsFromSpreadsheet(fullMemberListCols, FULL_MEMBER_LIST_SPREADSHEET_ID, "Full Member List", "Full Member List");
}


const attendanceRankerCols = ["name", "matched_name", "attended"];
function getAttendanceRanker() {
  copyColumnsFromSpreadsheet(attendanceRankerCols, ATTENDANCE_RANKER_SPREADSHEET_ID, "Ranked Attendance", "Ranked Attendance");
}


function compareFullToEmail() {
  mergeTwoSheets("Full Member List", "Email List", "Email Primary", "email", "Compare Full to Email");
}


function compareAttendanceToEmail() {
  mergeTwoSheets("Attendance Ranker", "Email List", "Email Primary", "email", "Compare Full to Email");
}
