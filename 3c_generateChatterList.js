// How far back to include spreadsheets, in months.
const MONTHS_BACK = 6;


function buildComprehensiveChatParticipantsSheet() {
  // Get the base folder, will be where comprehensive sheet is generated.
  const baseFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  // If an output file already exists in this folder, delete it.
  const existingFiles = baseFolder.getFilesByName('Chat Participants');
  while (existingFiles.hasNext()) {
    const file = existingFiles.next();
    file.setTrashed(true);
  }

  // Create the comprehensive spreadsheet file in the base folder.
  const outputSpreadsheet = SpreadsheetApp.create(`Past ${MONTHS_BACK} Months of Chat Participants`);
  const outputFile = DriveApp.getFileById(outputSpreadsheet.getId());
  baseFolder.addFile(outputFile);

  // Remove it from the root "My Drive" location so it lives only in the target folder.
  DriveApp.getRootFolder().removeFile(outputFile);

  // Get the first sheet in the newly created spreadsheet.
  const outputSheet = outputSpreadsheet.getSheets()[0];

  let isFirstFile = true;
  let outputRow = 1;

  // Find all matching spreadsheet files inside the folder tree.
  const matchingFiles = [];
  collectMatchingFilesRecursive_(baseFolder, matchingFiles);
  matchingFiles.sort((a, b) => a.name.localeCompare(b.name));

  // Loop through each matching spreadsheet and append its data vertically.
  matchingFiles.forEach(function(fileInfo) {
    const ss = SpreadsheetApp.openById(fileInfo.id);
    console.log(`Looking at spreadsheet '${ss.getName()}'`);
    const sheet = ss.getSheets()[0]; // Assumes the data is on the first sheet.
    const data = sheet.getDataRange().getValues();

    // Skip empty sheets.
    if (!data || data.length === 0) return;

    if (isFirstFile) {
      // Copy the headers from the first eligible sheet into row 1.
      outputSheet.getRange(outputRow, 1, 1, data[0].length).setValues([data[0]]);
      outputRow++;

      // Copy all rows except the header row.
      if (data.length > 1) {
        outputSheet.getRange(outputRow, 1, data.length - 1, data[0].length)
          .setValues(data.slice(1));
        outputRow += data.length - 1;
      }

      isFirstFile = false;
    } else {
      // For subsequent sheets, append only the data rows (skip headers).
      if (data.length > 1) {
        outputSheet.getRange(outputRow, 1, data.length - 1, data[0].length)
          .setValues(data.slice(1));
        outputRow += data.length - 1;
      }
    }
  });

  // freeze the header row for easier viewing.
  outputSheet.setFrozenRows(1);

  // Add a filter so the table can be sorted/filtered easily.
  outputSheet.getRange(1, 1, outputSheet.getLastRow(), outputSheet.getLastColumn()).createFilter();

  // auto resize columns
  resizeColumnsToFit(outputSheet);

  // Logger.log('Comprehensive sheet created: ' + outputSpreadsheet.getUrl());
  Logger.log('Comprehensive sheet created');
}


/**
 * Recursively walks through a folder and all of its subfolders,
 * collecting spreadsheet files whose titles contain "Chat Participants"
 * and whose leading YYYY-MM-DD date is within the last 6 months.
 */
function collectMatchingFilesRecursive_(folder, results) {
  // Check files directly inside this folder.
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();

    // Only consider Google Sheets files.
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    const title = file.getName();

    // Title must contain the phrase.
    if (title.indexOf('Chat Participants') === -1) continue;

    // Title must begin with a date like YYYY-MM-DD.
    if (!isWithinLastMonthsFromTitle_(title, MONTHS_BACK)) continue;

    // Save matching file metadata.
    results.push({
      id: file.getId(),
      name: title
    });
  }

  // Recurse into subfolders.
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    collectMatchingFilesRecursive_(subfolder, results);
  }
}


/**
 * Extracts the date from the start of the title and checks whether
 * it is within the last N months.
 */
function isWithinLastMonthsFromTitle_(title, monthsBack) {
  // Match a leading YYYY-MM-DD date.
  const match = title.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return false;

  const year = parseInt(match[1], 10);
  const month = parseInt(match[2], 10) - 1; // JS months are 0-based
  const day = parseInt(match[3], 10);

  const fileDate = new Date(year, month, day);
  if (isNaN(fileDate.getTime())) return false;

  // Compute the cutoff date.
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - monthsBack);

  // Only include files dated on or after the cutoff.
  return fileDate >= cutoff;
}
