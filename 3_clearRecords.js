const KEEP_SHEET_NAME = 'Base';

/**
 * Delete all sheets in the active spreadsheet except the sheet with the given name in KEEP_SHEET_NAME.
 * Usage: set KEEP_SHEET_NAME then run clearRecords()
 */
function clearRecords() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    if (sheets.length <= 1) return; // nothing to delete

    const keep = ss.getSheetByName(KEEP_SHEET_NAME);
    if (!keep) {
        throw new Error(`Sheet named "${KEEP_SHEET_NAME}" not found.`);
    }

    // Delete every sheet whose name is not the keep name
    sheets.forEach(sheet => {
        if (sheet.getName() !== KEEP_SHEET_NAME) {
            ss.deleteSheet(sheet);
        }
    });
}
