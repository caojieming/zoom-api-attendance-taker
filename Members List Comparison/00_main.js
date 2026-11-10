const fullMemberListCols = ["FirstName", "LastName"];
function getFullMemberList() {
  copyColumnsFromSpreadsheet(fullMemberListCols, FULL_MEMBER_LIST_SPREADSHEET_ID, "Full Member List", "Full Member List");
}


const attendanceRankerCols = ["name", "matched_name"];
function getAttendanceRanker() {
  copyColumnsFromSpreadsheet(attendanceRankerCols, ATTENDANCE_RANKER_SPREADSHEET_ID, "Ranked Attendance", "Ranked Attendance");
}
