/**
 * (re)generates the sheet with the name "Email List" by taking
 * the contents in cell J1 and converting it into a list, then
 * going through the list and adding them to the sheet.
 */
function updateEmailListSheet() {
  // Get the active spreadsheet and the target sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Email List");
  if (!sheet) throw new Error('Sheet "Email List" not found.');

  // Read the raw member list from J1; if it's empty, do nothing
  const j1 = String(sheet.getRange("J1").getValue() || "").trim();
  if (!j1) return;

  // Read the full table so we can inspect the headers and rebuild the rows
  const data = sheet.getDataRange().getValues();
  if (data.length < 1) return;

  // Find the "name" and "email" columns from the header row
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const nameCol = headers.indexOf("name") + 1;
  const emailCol = headers.indexOf("email") + 1;

  if (nameCol === 0 || emailCol === 0) {
    throw new Error('Table must contain headers "name" and "email".');
  }

  // Parse the member list from J1 and prepare rows for output
  const members = parseMembers_(j1);
  const rowsToWrite = [];

  // Track duplicates so the same person isn't added twice
  const seenEmails = new Set();
  const seenNames = new Set();

  for (const rawMember of members) {
    const member = rawMember.trim();
    if (!member) continue;
    if (member.toLowerCase() === "you") continue;

    // Split each entry into name/email parts if possible
    const parsed = parseMember_(member);
    if (!parsed) continue;

    const normalizedEmail = normalizeEmail_(parsed.email);
    const normalizedName = parsed.name ? normalizeName_(parsed.name) : "";

    // Deduplicate by email first, otherwise by name
    if (normalizedEmail) {
      if (seenEmails.has(normalizedEmail)) continue;
      seenEmails.add(normalizedEmail);
      if (normalizedName) seenNames.add(normalizedName);
    } else if (normalizedName) {
      if (seenNames.has(normalizedName)) continue;
      seenNames.add(normalizedName);
    } else {
      continue;
    }

    // Build a row matching the sheet width, filling in only the relevant columns
    const row = new Array(sheet.getLastColumn()).fill("");
    if (parsed.name) row[nameCol - 1] = parsed.name;
    if (parsed.email) row[emailCol - 1] = parsed.email;
    rowsToWrite.push(row);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Clear existing data below the header row, but keep formatting intact
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }

  // Write the rebuilt list back starting on row 2
  if (rowsToWrite.length > 0) {
    sheet.getRange(2, 1, rowsToWrite.length, lastCol).setValues(rowsToWrite);
  }

  resizeColumnsToFit_(sheet);
}


// turns the contents of cell J1 into an list (Array)
function parseMembers_(text) {
  // Split on commas only when not inside quotes or angle brackets
  const members = [];
  let current = "";
  let inQuotes = false;
  let inAngles = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];

    if (ch === '"') {
      inQuotes = !inQuotes;
      current += ch;
      continue;
    }

    if (ch === "<" && !inQuotes) {
      inAngles = true;
      current += ch;
      continue;
    }

    if (ch === ">" && !inQuotes) {
      inAngles = false;
      current += ch;
      continue;
    }

    if (ch === "," && !inQuotes && !inAngles) {
      if (current.trim()) members.push(current.trim());
      current = "";
      continue;
    }

    current += ch;
  }

  if (current.trim()) members.push(current.trim());

  return members;
}


function parseMember_(text) {
  const value = String(text).trim();

  // Match: Name <email@domain.com>
  const match = value.match(/^(.*?)\s*<\s*([^<>@\s]+@[^<>@\s]+\.[^<>@\s]+)\s*>$/);
  if (match) {
    return {
      name: cleanupName_(match[1]),
      email: match[2].trim(),
    };
  }

  // If the whole value is an email, return email only
  if (isEmail_(value)) {
    return {
      name: "",
      email: value,
    };
  }

  // Otherwise treat it as a name only
  return {
    name: cleanupName_(value),
    email: "",
  };
}


function cleanupName_(value) {
  return String(value)
    .trim()
    .replace(/^"(.+)"$/, "$1")
    .replace(/\s+/g, " ")
    .trim();
}


function isEmail_(value) {
  const text = String(value).trim();
  return /^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/.test(text);
}


function normalizeEmail_(value) {
  return String(value).trim().toLowerCase();
}


function normalizeName_(value) {
  return String(value)
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}


function buildRow_(sheet, nameCol, emailCol, name, email) {
  const row = new Array(sheet.getLastColumn()).fill("");
  if (nameCol > 0 && name) row[nameCol - 1] = name;
  if (emailCol > 0 && email) row[emailCol - 1] = email;
  return row;
}
