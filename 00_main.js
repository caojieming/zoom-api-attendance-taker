/**
This is the main file for functions related to updating both attendance records and meeting recordings/transcripts at once.

All files starting with 1 are necessary for anything to function.

If you only want to update attendance records, please see file 2a.
If you only want to change attendance getter code, please see all files starting with 2.

If you only want to update recordings/transcripts, please see file 3a.
If you only want to change recordings/transcripts getter code, please see all files starting with 3.
*/


function archivePastMonth() {
  getParticipants(ONE_MONTH_AGO, NOW);
  goToSheet(BASE_SHEET_NAME);
  sortRecords();
  goToSheet(BASE_SHEET_NAME);
  rankAttendance();

  getRecordings(ONE_MONTH_AGO, NOW);
}


/**
 * generally don't want to run this unless you are creating BOTH attendance sheets and recordings/transcripts from scratch
 */
function archivePastHalfYear() {
  getParticipantsHalfYear();
  goToSheet(BASE_SHEET_NAME);
  sortRecords();
  goToSheet(BASE_SHEET_NAME);
  rankAttendance();

  getRecordingsHalfYear();
}
