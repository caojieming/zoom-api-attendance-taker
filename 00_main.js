/**
This is the main file for functions related to updating both attendance records and meeting recordings/transcripts at once.

If you only want to update attendance records, please see file 2a.
If you only want to change attendance getter code, please see all files starting with 2.

If you only want to update recordings/transcripts, please see file 3a.
If you only want to change recordings/transcripts getter code, please see all files starting with 3.
*/


function archivePastMonth() {
  getParticipants(FROM, TO);
  goToSheet_(BASE_SHEET_NAME);
  sortRecords();
  goToSheet_(BASE_SHEET_NAME);
  rankAttendance();

  getRecordings(FROM, TO);
  buildChatterLinkedInSheet();
}


/**
 * generally don't want to run this unless you are creating BOTH attendance sheets and recordings/transcripts from scratch
 */
function archivePastHalfYear() {
  getParticipantsHalfYear();
  goToSheet_(BASE_SHEET_NAME);
  sortRecords();
  goToSheet_(BASE_SHEET_NAME);
  rankAttendance();

  getRecordingsHalfYear();
  buildChatterLinkedInSheet();
}
