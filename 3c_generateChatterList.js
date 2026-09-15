// How far back to include spreadsheets, in months.
const MONTHS_BACK = 6;

// Output columns.
// linkedin_url is the unique ID column.
const REQUIRED_COLUMNS = ['source_file_date', 'name', 'chapter', 'linkedin_url'];


/**
 * Builds a consolidated sheet from all matching "Chat Participants" spreadsheets
 * found in DRIVE_FOLDER_ID and its subfolders.
 *
 * Rules:
 * - linkedin_url is the unique key.
 * - If the same linkedin_url appears multiple times, keep the most recent row.
 * - If the most recent row is missing a value that an older row has, backfill it.
 */
function buildChatterLinkedInSheet() {
  // Get the base folder, which will also be the destination folder for the output file.
  const baseFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const outputName = `Past ${MONTHS_BACK} Months of Chat Participants`;

  // If an output file already exists in this folder, trash it first.
  const existingFiles = baseFolder.getFilesByName(outputName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  // Create the new output spreadsheet.
  const outputSpreadsheet = SpreadsheetApp.create(outputName);
  const outputFile = DriveApp.getFileById(outputSpreadsheet.getId());

  // Move the file into the target folder and remove it from root.
  baseFolder.addFile(outputFile);
  DriveApp.getRootFolder().removeFile(outputFile);

  // Use the first sheet in the newly created spreadsheet as the output sheet.
  const outputSheet = outputSpreadsheet.getSheets()[0];
  outputSheet.clearContents();

  // Collect matching files from the folder tree.
  const matchingFiles = [];
  collectMatchingFilesRecursive_(baseFolder, matchingFiles);

  // Sort oldest to newest so newer rows can overwrite older rows cleanly.
  matchingFiles.sort((a, b) => {
    const dateDiff = a.fileDate.getTime() - b.fileDate.getTime();
    if (dateDiff !== 0) return dateDiff;
    return a.name.localeCompare(b.name);
  });

  // Build a map keyed by linkedin_url.
  // Each value stores the most recent row, with older rows used to backfill missing values.
  const rowByLinkedin = new Map();
  let sourceHeaders = null;

  matchingFiles.forEach(function(fileInfo) {
    try {
      const ss = SpreadsheetApp.openById(fileInfo.id);
      console.log(`Looking at spreadsheet '${ss.getName()}'`);

      // Assumes data is on the first sheet.
      const sheet = ss.getSheets()[0];
      if (!sheet) return;

      const data = sheet.getDataRange().getValues();

      // Skip empty sheets.
      if (!data || data.length === 0) return;

      // Use the first non-empty sheet's header row.
      if (!sourceHeaders) {
        sourceHeaders = data[0].map(h => String(h).trim().toLowerCase());
      }

      const nameIdx = sourceHeaders.indexOf('name');
      const chapterIdx = sourceHeaders.indexOf('chapter');
      const linkedinIdx = sourceHeaders.indexOf('linkedin_url');

      if (nameIdx === -1 || chapterIdx === -1 || linkedinIdx === -1) {
        throw new Error('Missing one or more required columns: name, chapter, linkedin_url');
      }

      // Process each data row.
      for (let i = 1; i < data.length; i++) {
        const row = data[i];

        // Skip malformed rows.
        if (!row || row.length <= Math.max(nameIdx, chapterIdx, linkedinIdx)) continue;

        const linkedinValue = String(row[linkedinIdx] ?? '').trim();
        if (!linkedinValue) continue;

        const normalizedRow = {
          sourceFileDate: fileInfo.fileDate,
          name: row[nameIdx],
          chapter: row[chapterIdx],
          linkedin_url: row[linkedinIdx],
        };

        mergeRowByUniqueLinkedin_(rowByLinkedin, linkedinValue, normalizedRow);
      }
    } catch (err) {
      // Skip unreadable or broken files instead of failing the whole run.
      console.warn(`Skipping file '${fileInfo.name}': ${err}`);
    }
  });

  // Convert the map into output rows.
  const outputRows = [REQUIRED_COLUMNS];
  rowByLinkedin.forEach(function(item) {
    outputRows.push([
      item.sourceFileDate,
      item.name,
      item.chapter,
      item.linkedin_url,
    ]);
  });

  // Write the final output.
  outputSheet.clearContents();
  outputSheet.getRange(1, 1, outputRows.length, REQUIRED_COLUMNS.length).setValues(outputRows);

  // Format the source date column nicely.
  if (outputRows.length > 1) {
    outputSheet.getRange(2, 1, outputRows.length - 1, 1).setNumberFormat('yyyy-mm-dd');
  }

  // Freeze the header row for easier viewing.
  outputSheet.setFrozenRows(1);

  // Remove any existing filter before creating a new one.
  const existingFilter = outputSheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }

  // Add a filter for the full table.
  outputSheet.getRange(1, 1, outputSheet.getLastRow(), outputSheet.getLastColumn()).createFilter();

  // Auto-resize columns.
  resizeColumnsToFit(outputSheet);

  Logger.log('Comprehensive sheet created');
}


/**
 * Merges a row into the unique-linkedin map.
 *
 * Rules:
 * - If the linkedin_url is new, store it.
 * - If it already exists and the incoming row is newer, replace the stored row,
 *   but backfill any missing fields from the older row.
 * - If the incoming row is older, only backfill missing fields in the stored newer row.
 */
function mergeRowByUniqueLinkedin_(rowByLinkedin, linkedinKey, incomingRow) {
  const existingRow = rowByLinkedin.get(linkedinKey);

  // First time we've seen this linkedin_url: store it.
  if (!existingRow) {
    rowByLinkedin.set(linkedinKey, incomingRow);
    return;
  }

  const incomingTime = incomingRow.sourceFileDate.getTime();
  const existingTime = existingRow.sourceFileDate.getTime();

  // Incoming row is newer: make it the primary row,
  // but backfill any missing fields from the older row.
  if (incomingTime > existingTime) {
    rowByLinkedin.set(linkedinKey, {
      sourceFileDate: incomingRow.sourceFileDate,
      name: chooseLongerName_(incomingRow.name, existingRow.name),
      chapter: isBlank_(incomingRow.chapter) ? existingRow.chapter : incomingRow.chapter,
      linkedin_url: incomingRow.linkedin_url || existingRow.linkedin_url,
    });
    return;
  }

  // Existing row is newer: keep it, but backfill any missing fields from the older row.
  if (incomingTime < existingTime) {
    existingRow.name = chooseLongerName_(existingRow.name, incomingRow.name);
    existingRow.chapter = isBlank_(existingRow.chapter) ? incomingRow.chapter : existingRow.chapter;
    if (isBlank_(existingRow.linkedin_url)) {
      existingRow.linkedin_url = incomingRow.linkedin_url;
    }
    return;
  }

  // Same date: merge missing values, and prefer the longer name.
  existingRow.name = chooseLongerName_(existingRow.name, incomingRow.name);
  existingRow.chapter = isBlank_(existingRow.chapter) ? incomingRow.chapter : existingRow.chapter;
  if (isBlank_(existingRow.linkedin_url)) {
    existingRow.linkedin_url = incomingRow.linkedin_url;
  }
}

/**
 * Chooses the longer non-blank name.
 * If one name is blank, returns the other.
 * If both are present, returns the longer one.
 */
function chooseLongerName_(a, b) {
  const aStr = String(a ?? '').trim();
  const bStr = String(b ?? '').trim();

  if (!aStr) return bStr;
  if (!bStr) return aStr;
  return bStr.length > aStr.length ? bStr : aStr;
}


/**
 * Recursively walks through a folder and all of its subfolders,
 * collecting spreadsheet files whose titles contain "Chat Participants"
 * and whose leading YYYY-MM-DD date is within the last N months.
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
    if (!title.includes('Chat Participants')) continue;

    // Title must begin with a valid date like YYYY-MM-DD.
    const fileDate = extractDateFromTitle_(title);
    if (!fileDate) continue;

    // Only include files within the configured month range.
    if (!isWithinLastMonths_(fileDate, MONTHS_BACK)) continue;

    results.push({
      id: file.getId(),
      name: title,
      fileDate: fileDate,
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
 * Extracts the leading YYYY-MM-DD date from a file title.
 * Returns a Date object, or null if the title does not start with a valid date.
 */
function extractDateFromTitle_(title) {
  const match = title.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return null;

  const year = parseInt(match[1], 10);
  const month = parseInt(match[2], 10) - 1; // JS months are 0-based.
  const day = parseInt(match[3], 10);

  const fileDate = new Date(year, month, day);

  // Validate the date components to avoid rollover dates like 2024-02-31.
  if (
    isNaN(fileDate.getTime()) ||
    fileDate.getFullYear() !== year ||
    fileDate.getMonth() !== month ||
    fileDate.getDate() !== day
  ) {
    return null;
  }

  return fileDate;
}


/**
 * Checks whether a date is within the last N months from now.
 * Uses a calendar-month cutoff rather than a fixed number of days.
 */
function isWithinLastMonths_(date, monthsBack) {
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - monthsBack);

  return date >= cutoff;
}


/**
 * Returns true when a value is blank or only whitespace.
 */
function isBlank_(value) {
  return String(value ?? '').trim() === '';
}
