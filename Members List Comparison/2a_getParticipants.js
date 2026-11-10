/**
 * Copies selected columns from a source spreadsheet into a sheet
 * in the spreadsheet this script is bound to.
 *
 * If a sheet with the same name already exists, it is deleted and recreated.
 * The recreated sheet is inserted at the same tab position as the old one.
 *
 * @param {string[]} columnNames - List of header names to copy.
 * @param {string} sourceSpreadsheetId - ID of the input Google Spreadsheet.
 * @param {string} sourceSheetName - Sheet name in the source spreadsheet.
 * @param {string} newSheetName - Sheet name for the created sheet.
 */
function copyColumnsFromSpreadsheet(columnNames, sourceSpreadsheetId, sourceSheetName, newSheetName) {
  // Get the spreadsheet this script is bound to.
  const targetSS = SpreadsheetApp.getActiveSpreadsheet();

  // Open the source spreadsheet by ID.
  const sourceSS = SpreadsheetApp.openById(sourceSpreadsheetId);

  // Find the source sheet by name.
  const sourceSheet = sourceSS.getSheetByName(sourceSheetName);
  if (!sourceSheet) throw new Error("Source sheet not found.");

  // Read all data from the source sheet.
  const data = sourceSheet.getDataRange().getValues();
  if (data.length === 0) throw new Error("Source sheet is empty.");

  // The first row is assumed to be the header row.
  const headers = data[0];

  // Put requested column names into a Set for fast lookup.
  const wanted = new Set(columnNames);

  // Store the column indexes we want to keep.
  const matchedIndexes = [];

  // Walk through the headers and remember which columns match.
  headers.forEach((header, i) => {
    if (wanted.has(String(header))) {
      matchedIndexes.push(i);
    }
  });

  // Stop if none of the requested columns were found.
  if (matchedIndexes.length === 0) {
    throw new Error("No matching columns found.");
  }

  // Build the output array by keeping only the matched columns.
  const output = data.map(row => matchedIndexes.map(colIndex => row[colIndex]));

  // Determine the name of the output sheet.
  const sheetName = newSheetName;

  // Check whether a sheet with that name already exists.
  const existing = targetSS.getSheetByName(sheetName);

  // If it exists, remember its tab position before deleting it.
  // Google Sheets tab positions are 1-based for insertSheet(name, index).
  let tabIndex = null;
  if (existing) {
    tabIndex = targetSS.getSheets().indexOf(existing) + 1;
    targetSS.deleteSheet(existing);
  }

  // Create a new sheet.
  // If we saved a tab position, insert it back into that same position.
  const newSheet = tabIndex
    ? targetSS.insertSheet(sheetName, tabIndex)
    : targetSS.insertSheet(sheetName);

  // Write the copied data into the new sheet.
  newSheet
    .getRange(1, 1, output.length, output[0].length)
    .setValues(output);


  // Freeze the header row.
  newSheet.setFrozenRows(1);

  // Add a filter over the copied data range.
  newSheet.getRange(1, 1, output.length, output[0].length).createFilter();

  // Auto-fit columns for readability.
  resizeColumnsToFit_(newSheet);
}
