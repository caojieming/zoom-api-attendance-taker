/**
 * Full outer join of two sheets by matching two columns case-insensitively.
 * - Each table starts at column A
 * - Each table ends at the first blank header cell
 * - Keeps both join columns
 * - Preserves all original columns from both sheets
 * - Includes unmatched rows from both sheets
 *
 * @param {string} sheetName1 First sheet/tab name
 * @param {string} sheetName2 Second sheet/tab name
 * @param {string} joinColumn1 Header name to join on in first sheet
 * @param {string} joinColumn2 Header name to join on in second sheet
 * @param {string} outputSheetName Optional output sheet name (default: "MergedSheet")
 */
function mergeTwoSheets(sheetName1, sheetName2, joinColumn1, joinColumn2, outputSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  outputSheetName = outputSheetName || "MergedSheet";

  const sheet1 = ss.getSheetByName(sheetName1);
  const sheet2 = ss.getSheetByName(sheetName2);

  if (!sheet1) throw new Error(`Sheet not found: ${sheetName1}`);
  if (!sheet2) throw new Error(`Sheet not found: ${sheetName2}`);

  const data1 = sheet1.getDataRange().getValues();
  const data2 = sheet2.getDataRange().getValues();

  if (data1.length < 1) throw new Error(`Sheet "${sheetName1}" has no data.`);
  if (data2.length < 1) throw new Error(`Sheet "${sheetName2}" has no data.`);

  const headers1 = getHeadersUntilBlank_(data1[0]);
  const headers2 = getHeadersUntilBlank_(data2[0]);

  const joinIndex1 = headers1.findIndex(h => String(h).trim().toLowerCase() === String(joinColumn1).trim().toLowerCase());
  const joinIndex2 = headers2.findIndex(h => String(h).trim().toLowerCase() === String(joinColumn2).trim().toLowerCase());

  if (joinIndex1 === -1) throw new Error(`Join column "${joinColumn1}" not found in "${sheetName1}".`);
  if (joinIndex2 === -1) throw new Error(`Join column "${joinColumn2}" not found in "${sheetName2}".`);

  // Normalize join values for case-insensitive matching
  const normalize = (value) => String(value === null || value === undefined ? "" : value).trim().toLowerCase();

  // Build lookup map for sheet2
  const map2 = new Map();
  for (let r = 1; r < data2.length; r++) {
    const row2 = data2[r].slice(0, headers2.length);
    const key = normalize(row2[joinIndex2]);
    if (!key) continue;

    if (!map2.has(key)) map2.set(key, []);
    map2.get(key).push(row2);
  }

  // Track which sheet2 rows were matched so we can output unmatched ones later
  const matchedSheet2Keys = new Set();

  // Output headers: keep all columns from both sheets, including both join columns
  const outputHeaders = [
    ...headers1.map(h => `${sheetName1}.${h}`),
    ...headers2.map(h => `${sheetName2}.${h}`)
  ];

  const outputRows = [outputHeaders];

  // Full outer join: process all rows from sheet1 first
  for (let r1 = 1; r1 < data1.length; r1++) {
    const row1 = data1[r1].slice(0, headers1.length);
    const key = normalize(row1[joinIndex1]);
    const matches = map2.get(key);

    if (matches && matches.length) {
      matches.forEach(row2 => {
        matchedSheet2Keys.add(key);
        const fullRow = [
          ...padRow_(row1, headers1.length),
          ...padRow_(row2, headers2.length)
        ];
        outputRows.push(fullRow);
      });
    }
    else {
      // No match in sheet2: keep row1 and blank out sheet2 columns
      outputRows.push([
        ...padRow_(row1, headers1.length),
        ...Array(headers2.length).fill("")
      ]);
    }
  }

  // Add unmatched rows from sheet2
  for (let r2 = 1; r2 < data2.length; r2++) {
    const row2 = data2[r2].slice(0, headers2.length);
    const key = normalize(row2[joinIndex2]);

    // If key is empty or never matched, include it
    // If there are duplicate keys in sheet2 and some matched, this still outputs only the unmatched keys
    if (!key || !Array.from(matchedSheet2Keys).includes(key)) {
      const row1Blank = Array(headers1.length).fill("");
      outputRows.push([
        ...row1Blank,
        ...padRow_(row2, headers2.length)
      ]);
    }
  }

  // Create or clear output sheet
  let outSheet = ss.getSheetByName(outputSheetName);
  if (!outSheet) {
    outSheet = ss.insertSheet(outputSheetName);
  }
  else {
    outSheet.clearContents();
  }

  // Write result
  outSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
  outSheet.autoResizeColumns(1, outputRows[0].length);


  // Freeze the header row.
  outSheet.setFrozenRows(1);

  // Add a left border to each source table block:
  // - column A for sheet1 block
  // - first column of sheet2 block
  const blockStarts = [1, headers1.length + 1];
  blockStarts.forEach(col => {
    outSheet.getRange(1, col, outputRows.length, 1)
      .setBorder(null, true, null, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  });

  // (re)Add a filter over the copied data range.
  const filter = outSheet.getFilter();
  if (filter) filter.remove();
  outSheet.getRange(1, 1, outputRows.length, outputRows[0].length).createFilter();
}

/**
 * Returns headers from a row up to (but not including) the first blank header cell.
 * Assumes tables start at column A.
 */
function getHeadersUntilBlank_(headerRow) {
  const headers = [];
  for (let i = 0; i < headerRow.length; i++) {
    if (headerRow[i] === "" || headerRow[i] === null) break;
    headers.push(headerRow[i]);
  }
  return headers;
}

/**
 * Pads or trims a row to a specific length.
 */
function padRow_(row, length) {
  const out = row.slice(0, length);
  while (out.length < length) out.push("");
  return out;
}
