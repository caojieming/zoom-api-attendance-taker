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


/**
 * return true if partString is a subsequence of mainString
 * subsequence: a string that can be derived from another string by deleting some chars without changing the order of the remaining chars
 */
/**
 * return true if partString is a subsequence of mainString
 * subsequence: a string that can be derived from another string by deleting some chars without changing the order of the remaining chars
 */
function isSubsequence_(mainString, partString, minStrLen = 5) {
  mainString = mainString.replace(/[^a-zA-Z]/g, '').toLowerCase();
  partString = partString.replace(/[^a-zA-Z]/g, '').toLowerCase();
  let i = 0, j = 0;
  while (i < mainString.length && j < partString.length) {
    if (mainString[i] === partString[j]) j++;
    i++;
  }
  // If j reaches end of partString, that means all characters of partString were found in mainString
  // minStrLen is a safeguard to prevent coincidental subsequence recognition due to being extremely generic
  return j === partString.length && mainString.length >= minStrLen && partString.length >= minStrLen;
}
