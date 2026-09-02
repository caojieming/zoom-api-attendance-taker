/**
 * Delete all sheets in the active spreadsheet except the sheet with the given name in BASE_SHEET_NAME.
 * Usage: set BASE_SHEET_NAME (in 1b) then run clearRecords()
 */
function clearRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  if (sheets.length <= 1) return; // nothing to delete

  const keep = ss.getSheetByName(BASE_SHEET_NAME);
  if (!keep) {
    throw new Error(`Sheet named "${BASE_SHEET_NAME}" not found.`);
  }

  // Delete every sheet whose name is not the keep name
  sheets.forEach(sheet => {
    if (sheet.getName() !== BASE_SHEET_NAME) {
      ss.deleteSheet(sheet);
    }
  });
}


/**
 * sorts subsheets with names that start with dates from most recent to least recent
 */
function sortRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsInOrder = ss.getSheets(); // current left-to-right order

  const dateRegex = /^(\d{4})-(\d{2})-(\d{2})/; // starts with YYYY-MM-DD

  // Keep everything before the first matching (non-date) sheet in place.
  const firstValidIdx = sheetsInOrder.findIndex(s => dateRegex.test(s.getName()));
  const prefix = firstValidIdx === -1 ? sheetsInOrder : sheetsInOrder.slice(0, firstValidIdx);

  const dateSheets = firstValidIdx === -1 ? [] : sheetsInOrder.slice(firstValidIdx).filter(s => dateRegex.test(s.getName()));

  // Sort only the date sheets; leave any non-date sheets after that untouched/ignored per your request.
  dateSheets.sort((a, b) => b.getName().localeCompare(a.getName()));

  // Desired final order = prefix (unchanged order) + sorted date sheets
  const desired = prefix.concat(dateSheets);

  desired.forEach((sheet, i) => {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(i + 1); // positions are 1-based
  });
}