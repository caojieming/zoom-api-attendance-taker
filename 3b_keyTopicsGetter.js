// gemini model to use
const GEMINI_MODEL = "gemini-3.5-flash-lite";
/**
List of ideal models to use (these models are in high demand and may fail to run):
gemini-3.7-flash
gemini-3.6-flash
gemini-3.5-flash

Models below are the safe backup models (pretty much guarenteed to work without issue):
gemini-3.5-flash-lite

Generally best to stick with the lite model, I've consistently run into high traffic errors with the normal models.
*/


function createKeyTopicsSheet(driveFolderId, fileName, transcript) {
  const spreadsheet = createSpreadsheetInFolder(fileName, driveFolderId);
  const sheet = spreadsheet.getSheets()[0];
  sheet.setName("Key Topics");
  

  const prompt = `
You are analyzing a transcript and extracting its key topics.

Return ONLY valid JSON in this exact shape:
{
  "topics": [
    {
      "key_topic": "short concise topic name",
      "summary": "1-2 sentence summary",
      "significance": 1-5,
      "confidence": 1-5
    }
  ]
}

Rules:
- Keep topics concise and distinct.
- Merge duplicates and near-duplicates.
- Prioritize importance and recurrence.
- Prefer the most central themes, not minor details.
- Output between 5 and 20 topics unless the transcript clearly supports fewer or more.
- significance: 1 = minor, 5 = highly central.
- confidence: 1 = low confidence, 5 = very confident.
- No markdown, no code fences, no extra text.

Transcript:
${transcript}
`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;

  const payload = {
    contents: [{
      role: "user",
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: 0.2,
      topP: 0.95,
      maxOutputTokens: 8192,
      responseMimeType: "application/json"
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const status = response.getResponseCode();
  const body = response.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error(`Gemini API error (${status}): ${body}`);
  }

  const data = JSON.parse(body);
  const text = data.candidates?.[0]?.content?.parts?.[0]?.text;

  if (!text) {
    throw new Error("No model output returned.");
  }

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    throw new Error("Model did not return valid JSON: " + text);
  }

  const topics = Array.isArray(parsed.topics) ? parsed.topics : [];
  const rows = [
    ["key topic", "summary", "significance", "confidence"]
  ];

  topics.forEach(t => {
    rows.push([
      t.key_topic || "",
      t.summary || "",
      Number(t.significance) || "",
      Number(t.confidence) || ""
    ]);
  });

  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);

  // Basic formatting
  sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
  resizeColumnsToFit(sheet);
  // sheet.autoResizeColumns(1, 4);

  // setting up filter table
  sheet.getRange(1, 1, rows.length, rows[0].length).createFilter();

  return {
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    topicsCount: topics.length
  };
}


function createSpreadsheetInFolder(sheetName, folderId) {
  const ss = SpreadsheetApp.create(sheetName);
  const file = DriveApp.getFileById(ss.getId());
  const folder = DriveApp.getFolderById(folderId);

  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  return ss;
}
