/**
 * Delete all sheets except the one named in BASE_SHEET_NAME.
 */
function clearRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const keepName = BASE_SHEET_NAME;
  const keepSheet = ss.getSheetByName(keepName);

  if (!keepSheet) {
    throw new Error(`Sheet named "${keepName}" not found.`);
  }

  ss.getSheets().forEach(sheet => {
    if (sheet.getName() !== keepName) {
      ss.deleteSheet(sheet);
    }
  });
}


/**
 * Sort sheets whose names start with YYYY-MM-DD from newest to oldest.
 * Non-date sheets stay in their current relative order and remain first.
 */
function sortRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const nonDateSheets = [];
  const dateSheets = [];

  sheets.forEach(sheet => {
    if (DATE_PREFIX.test(sheet.getName())) {
      dateSheets.push(sheet);
    } else {
      nonDateSheets.push(sheet);
    }
  });

  dateSheets.sort((a, b) => b.getName().localeCompare(a.getName()));

  [...nonDateSheets, ...dateSheets].forEach((sheet, index) => {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(index + 1);
  });
}
