// How far back to include spreadsheets, in months.
const MONTHS_BACK = 6;

// Output columns.
// linkedin_url is the unique ID column.
const REQUIRED_COLUMNS = ['latest_date', 'name', 'chapter', 'linkedin_url'];


/**
 * Builds a consolidated sheet from all matching "Chat Participants" spreadsheets
 * found in DRIVE_FOLDER_ID and its subfolders.
 *
 * Rules:
 * - full name is the unique key.
 * - If the same full name appears multiple times, keep the most recent row.
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

  // Create the new output spreadsheet in DRIVE_FOLDER_ID.
  const outputSpreadsheet = createSpreadsheetInFolder_(outputName, DRIVE_FOLDER_ID);

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

  // Build a map keyed by full name.
  // Each value stores the most recent row, with older rows used to backfill missing values.
  const rowByName = new Map();

  // Cache source column indexes once instead of looking them up repeatedly.
  let sourceHeaders = null;
  let nameIdx = -1;
  let chapterIdx = -1;
  let linkedinIdx = -1;

  // Get contents of full members list for cross-referencing if certain info is missing.
  const columnNamesFilter = ["FirstName", "LastName", "Chapter", "LinkedIn"];
  const fullMembersList = getTableFromSheet_(
    MEMBERS_SPREADSHEET_ID,
    MEMBERS_SPREADSHEET_SHEETNAME,
    columnNamesFilter
  ) || [];

  // Precompute normalized member records once instead of redoing string cleanup for every source row (better performance when there are many chat participant rows)
  const normalizedMembers = fullMembersList
    .map(member => {
      const firstName = String(member.FirstName ?? '').trim();
      const lastName = String(member.LastName ?? '').trim();
      const fullName = `${firstName} ${lastName}`.trim();

      return {
        name: fullName,
        chapter: String(member.Chapter ?? '').trim(),
        linkedin_url: normalizeLinkedInUrl_(member.LinkedIn),
      };
    })
    .filter(member => member.name);

  // loop through all spreadsheets that contain "Chat Participants" in their title
  for (let fileIndex = 0; fileIndex < matchingFiles.length; fileIndex++) {
    const fileInfo = matchingFiles[fileIndex];

    try {
      // One SpreadsheetApp open call per source file is unavoidable, so keep the rest of the work local.
      const ss = SpreadsheetApp.openById(fileInfo.id);
      console.log(`Looking at spreadsheet '${ss.getName()}'`);

      // Assumes data is on the first sheet.
      const sheet = ss.getSheets()[0];
      if (!sheet) continue;

      // One data-range read per file.
      const data = sheet.getDataRange().getValues();

      // Skip empty sheets.
      if (!data || data.length === 0) continue;

      // Use the first non-empty sheet's header row.
      if (!sourceHeaders) {
        sourceHeaders = data[0].map(h => String(h).trim().toLowerCase());

        nameIdx = sourceHeaders.indexOf('name');
        chapterIdx = sourceHeaders.indexOf('chapter');
        linkedinIdx = sourceHeaders.indexOf('linkedin_url');

        if (nameIdx === -1 || chapterIdx === -1 || linkedinIdx === -1) {
          throw new Error('Missing one or more required columns: name, chapter, linkedin_url');
        }
      }

      // Process each data row (ignoring the first because it's the header row)
      for (let i = 1; i < data.length; i++) {
        const row = data[i];

        // Skip malformed rows.
        if (!row || row.length <= Math.max(nameIdx, chapterIdx, linkedinIdx)) continue;

        const sourceName = normalizeParticipantName_(String(row[nameIdx] ?? ''));

        // Skip rows with empty names or names that are too short
        if (!sourceName || sourceName.length < PARTICIPANT_MIN_NAME_LENGTH) continue;

        // skip rows with names that are part of the blacklist
        if (
          PARTICIPANT_BLACKLIST.length > 0 &&
          PARTICIPANT_BLACKLIST.some((keyword) =>
            sourceName.toLowerCase().includes(keyword.toLowerCase())
          )
        ) continue;

        const sourceChapter = String(row[chapterIdx] ?? '').trim();
        const sourceLinkedinUrl = normalizeLinkedInUrl_(row[linkedinIdx]);

        // Track only the best matching member instead of storing every above-threshold match.
        let bestMatch = null;

        // Try to find relevant existing member info in fullMembersList.
        for (let m = 0; m < normalizedMembers.length; m++) {
          const member = normalizedMembers[m];
          const similarityRatio = stringSimilarity_(sourceName, member.name);
          // name from full members list first, name from chat list second
          const isSubsequence = isSubsequence_(member.name, sourceName);

          if (similarityRatio >= SIMILARITY_THRESHOLD_RATIO || isSubsequence) {
            if (!bestMatch || similarityRatio > bestMatch.similarity_ratio) {
              bestMatch = {
                name: member.name,
                chapter: member.chapter,
                linkedin_url: member.linkedin_url,
                similarity_ratio: similarityRatio,
              };
            }
          }
        }

        // Prioritize info obtained from fullMembersList to keep values as consistent as possible for later duplicate merging.
        const normalizedRow = {
          sourceFileDate: fileInfo.fileDate,
          name: bestMatch?.name || sourceName,
          chapter: bestMatch?.chapter || sourceChapter,
          linkedin_url: bestMatch?.linkedin_url || sourceLinkedinUrl,
        };

        // Skip rows that still have no key after normalization.
        // Might remove this later to allow rows without names, just remember to also alter mergeRowByUniqueName_() to ignore rows with empty names
        if (!normalizedRow.name) continue;

        mergeRowByUniqueName_(rowByName, normalizedRow.name, normalizedRow);
      }
    } catch (err) {
      // Skip unreadable or broken files instead of failing the whole run.
      console.warn(`Skipping file '${fileInfo.name}': ${err}`);
    }
  }

  // Convert the map into output rows.
  const outputRows = [REQUIRED_COLUMNS];
  rowByName.forEach(function(item) {
    outputRows.push([
      item.sourceFileDate,
      item.name,
      item.chapter,
      item.linkedin_url,
    ]);
  });

  // Write the final output
  outputSheet.clearContents();
  const outputRange = outputSheet.getRange(1, 1, outputRows.length, REQUIRED_COLUMNS.length);
  outputRange.setValues(outputRows);

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

  // Add a filter only when there is actual data beyond the header row.
  // This avoids an unnecessary filter call on empty outputs.
  if (outputRows.length > 1) {
    outputSheet.getRange(1, 1, outputSheet.getLastRow(), outputSheet.getLastColumn()).createFilter();
  }

  // Auto-resize columns.
  resizeColumnsToFit_(outputSheet);

  Logger.log('Comprehensive sheet created');
}


/**
 * Normalizes LinkedIn URLs into a consistent key format.
 */
function normalizeLinkedInUrl_(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\//, '') // remove "https://"
    .replace(/^www\./, '') // remove "www."
    .replace(/\/$/, ''); // remove trailing '/'
}


/**
 * Merges a row into the unique-name map.
 *
 * Rules:
 * - If the full name is new, store it.
 * - If it already exists and the incoming row is newer, replace the stored row,
 *   but backfill any missing fields from the older row.
 * - If the incoming row is older, only backfill missing fields in the stored newer row.
 */
function mergeRowByUniqueName_(rowByName, nameKey, incomingRow) {
  const existingRow = rowByName.get(nameKey);

  // First time we've seen this full name: store it.
  if (!existingRow) {
    rowByName.set(nameKey, incomingRow);
    return;
  }

  const incomingTime = incomingRow.sourceFileDate.getTime();
  const existingTime = existingRow.sourceFileDate.getTime();

  // Incoming row is newer: make it the primary row,
  // but backfill any missing fields from the older row.
  if (incomingTime > existingTime) {
    rowByName.set(nameKey, {
      sourceFileDate: incomingRow.sourceFileDate,
      name: chooseLongerName_(incomingRow.name, existingRow.name),
      // name: existingRow.name,
      chapter: isBlank_(incomingRow.chapter) ? existingRow.chapter : incomingRow.chapter,
      linkedin_url: incomingRow.linkedin_url || existingRow.linkedin_url,
    });
    return;
  }

  // Existing row is newer: keep it, but backfill any missing fields from the older row.
  if (incomingTime < existingTime) {
    existingRow.name = chooseLongerName_(existingRow.name, incomingRow.name);
    // existingRow.name = existingRow.name;
    existingRow.chapter = isBlank_(existingRow.chapter) ? incomingRow.chapter : existingRow.chapter;
    if (isBlank_(existingRow.linkedin_url)) {
      existingRow.linkedin_url = incomingRow.linkedin_url;
    }
    return;
  }

  // Same date: merge missing values, and prefer the longer name.
  existingRow.name = chooseLongerName_(existingRow.name, incomingRow.name);
  // existingRow.name = existingRow.name;
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
  const match = title.match(DATE_PREFIX);
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


/**
 * returns a google sheet as a list of objects, optionally taking in a list of column names to filter in
 */
function getTableFromSheet_(spreadsheetId, sheetName, columnNames) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(String);
  const useColumns = Array.isArray(columnNames) && columnNames.length > 0
    ? columnNames.map(String)
    : headers;

  const indexes = useColumns.map(name => {
    const idx = headers.indexOf(name);
    if (idx === -1) throw new Error(`Column not found: ${name}`);
    return idx;
  });

  return data.slice(1).map(row => {
    const obj = {};
    indexes.forEach((idx, i) => {
      obj[useColumns[i]] = row[idx];
    });
    return obj;
  });
}

