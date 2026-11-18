/**
 * updates the sheet with the name "WhatsApp Members" by taking
 * the contents in cell J1 and converting it into a list, then
 * going through the list and checking if the current item already
 * exists in column "name" or "phone". If it doesn't, it adds it.
 */
function updateWhatsAppSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("WhatsApp Members");
  if (!sheet) throw new Error('Sheet "WhatsApp Members" not found.');

  // currently using cell J1 in "WhatsApp Members" to input the raw list
  const j1 = String(sheet.getRange("J1").getValue() || "").trim();
  if (!j1) return;

  const data = sheet.getDataRange().getValues();
  if (data.length < 1) return;

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const nameCol = headers.indexOf("name") + 1;
  const phoneCol = headers.indexOf("phone") + 1;

  if (nameCol === 0 || phoneCol === 0) {
    throw new Error('Table must contain headers "name" and "phone".');
  }

  // Existing values for duplicate checking
  const existingNames = new Set();
  const existingPhones = new Set();

  // adding existing table names and phones to their respective sets
  for (let r = 1; r < data.length; r++) {
    // 1 of name or phone cols will be filled, and the other may be empty: grab both, prioritizing name
    const nameVal = String(data[r][nameCol - 1] || "").trim();
    const phoneVal = String(data[r][phoneCol - 1] || "").trim();
    if (nameVal) existingNames.add(normalizeName_(nameVal));
    if (phoneVal) existingPhones.add(normalizePhone_(phoneVal));
  }

  // Split by commas, but keep entries like "MBA,MS,PMP,SPC" attached to the name
  const members = parseMembers_(j1);

  const rowsToAdd = [];

  for (const rawMember of members) {
    const member = rawMember.trim();
    if (!member) continue;

    // ignore "You", that provides no useful info (manually add your own row if you want to)
    if (member.toLowerCase() === "you") continue;

    // is a phone
    if (isPhoneNumber_(member)) {
      const normalizedPhone = normalizePhone_(member);
      if (!existingPhones.has(normalizedPhone)) {
        rowsToAdd.push(buildRow_(sheet, nameCol, phoneCol, null, member));
        existingPhones.add(normalizedPhone);
      }
    }
    // is a name
    else {
      const normalizedName = normalizeName_(member);
      if (!existingNames.has(normalizedName)) {
        rowsToAdd.push(buildRow_(sheet, nameCol, phoneCol, member, null));
        existingNames.add(normalizedName);
      }
    }
  }

  if (rowsToAdd.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAdd.length, sheet.getLastColumn()).setValues(rowsToAdd);
  }

  // auto resize columns
  resizeColumnsToFit_(sheet);
}


// turns the contents of cell J1 into an list (Array)
function parseMembers_(text) {
  // Prefer commas as separators, but preserve comma-separated suffixes in names
  const parts = text.split(",");
  const members = [];

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i].trim();
    if (!part) continue;

    // If this part starts a name with "~", collect following short uppercase credential tokens
    if (part.startsWith("~")) {
      let current = part;
      while (i + 1 < parts.length) {
        const next = parts[i + 1].trim();
        // Attach credential-like fragments such as MBA, MS, PMP, SPC
        if (/^[A-Z]{2,10}$/.test(next)) {
          current += "," + next;
          i++;
        }
        else {
          break;
        }
      }
      // trims and removes leading ~
      members.push(current.replace(/^~\s*/, "").trim());
    }
    else {
      // trims and removes leading ~
      members.push(part.replace(/^~\s*/, "").trim());
    }
  }

  return members;
}


function isPhoneNumber_(value) {
  // removes all non number chars, leaving only numbers
  const digits = String(value).replace(/\D/g, "");

  // return true if value contains >= 7 numbers and only contains numbers or "+.()-"
  return digits.length >= 7 && /^[+\d][\d\s().-]+$/.test(String(value).trim());
}


function normalizePhone_(value) {
  // removes all non number chars
  return String(value).replace(/\D/g, "");
}


function normalizeName_(value) {
  return String(value)
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}


function buildRow_(sheet, nameCol, phoneCol, name, phone) {
  const row = new Array(sheet.getLastColumn()).fill("");
  if (nameCol > 0 && name) row[nameCol - 1] = name;
  if (phoneCol > 0 && phone) row[phoneCol - 1] = phone;
  return row;
}
