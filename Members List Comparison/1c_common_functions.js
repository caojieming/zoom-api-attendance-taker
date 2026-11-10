/**
 * Resizes all populated columns in a sheet to fit their content.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function resizeColumnsToFit_(sheet) {
  const dataRange = sheet.getDataRange();
  const numColumns = dataRange.getNumColumns();

  if (numColumns > 0) {
    sheet.autoResizeColumns(1, numColumns);
  }
}
