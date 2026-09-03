/**
 * creates a sheet listing the key topics of a meeting by sending the transcript along with a prompt to Gemini
 */
function createKeyTopicsSheet(driveFolderId, fileName, transcript) {
  // Prompt tells the model exactly what JSON shape to return and how to behave.
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

  // Delegate all shared work to the generic helper.
  return createExtractionSheet({
    driveFolderId,
    fileName,
    prompt,
    rootKey: "topics",
    headers: ["key_topic", "summary", "significance", "confidence"],
    rowMapper: t => [
      t.key_topic || "",
      t.summary || "",
      Number(t.significance) || "",
      Number(t.confidence) || ""
    ],
    countKey: "topicsCount"
  });
}


/**
 * creates a sheet listing the chat participants of a meeting by sending the chat log along with a prompt to Gemini
 */
function createChatParticipantsSheet(driveFolderId, fileName, chatLog) {
  // Prompt tells the model to extract participants from a chat log.
  const prompt = `
You are analyzing a chat log and extracting its participants.

Return ONLY valid JSON in this exact shape:
{
  "participants": [
    {
      "name": "participant name",
      "linkedin_url": "full LinkedIn URL or empty string",
      "chapter": "city or state name or empty string",
      "self_introduction": "brief self-introduction extracted or inferred from the chat"
    }
  ]
}

Rules:
- Return one entry per unique participant.
- Include only people who appear to be participants in the chat.
- Use the person's stated name if available; otherwise use the best available identifier.
- linkedin_url: fill in only if explicitly found in the chat; otherwise use an empty string.
- chapter: generally a city or state name; fill in only if explicitly found in the chat; otherwise use an empty string.
- self_introduction should be short and based on how the participant introduces themselves in the chat.
- If a participant does not self-introduce, infer a concise one from the available context, or use an empty string if none is available.
- Merge duplicate references to the same person.
- Do not include people mentioned only in passing unless they clearly participate.
- Output between 1 and 50 participants unless the chat clearly supports fewer or more.
- No markdown, no code fences, no extra text.

Chat log:
${chatLog}
`;

  // Delegate all shared work to the generic helper.
  return createExtractionSheet({
    driveFolderId,
    fileName,
    prompt,
    rootKey: "participants",
    headers: ["name", "linkedin_url", "chapter", "self_introduction"],
    rowMapper: p => [
      p.name || "",
      p.linkedin_url || "",
      p.chapter || "",
      p.self_introduction || ""
    ],
    countKey: "participantsCount"
  });
}


// Generic extraction runner shared by both use cases.
function createExtractionSheet({
  driveFolderId,
  fileName,
  prompt,
  rootKey,
  headers,
  rowMapper,
  countKey
}) {
  // Create the spreadsheet and use its first sheet as the output destination.
  const spreadsheet = createSpreadsheetInFolder(fileName, driveFolderId);
  const sheet = spreadsheet.getSheets()[0];

  // Send the prompt to Gemini and parse the JSON response.
  const parsed = callGeminiJson(prompt);

  // Pull out the expected array from the returned JSON.
  const items = Array.isArray(parsed[rootKey]) ? parsed[rootKey] : [];

  // Convert each item into a row of sheet values.
  const rows = items.map(rowMapper);

  // Write the final table to the sheet.
  writeTableToSheet(sheet, headers, rows);

  // Return useful metadata about the generated spreadsheet.
  return {
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    [countKey]: items.length
  };
}


// Calls Gemini and returns parsed JSON, retrying only on transient "high traffic" style failures.
function callGeminiJson(prompt) {
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

  const maxAttempts = 3;
  let lastError = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const status = response.getResponseCode();
      const body = response.getContentText();

      if (status < 200 || status >= 300) {
        const errorText = (body || "").toLowerCase();

        // Retry only on transient capacity/traffic-style errors.
        const isTransient =
          status === 429 ||
          status === 503 ||
          errorText.includes("high traffic") ||
          errorText.includes("overloaded") ||
          errorText.includes("try again") ||
          errorText.includes("resource exhausted") ||
          errorText.includes("temporarily unavailable");

        if (isTransient) {
          throw new Error(`Transient Gemini error (${status}): ${body}`);
        }

        // Fail immediately for non-transient errors.
        throw new Error(`Gemini API error (${status}): ${body}`);
      }

      const data = JSON.parse(body);
      const text = data.candidates?.[0]?.content?.parts?.[0]?.text;

      if (!text) throw new Error("No model output returned.");

      // Bad JSON is not a traffic issue, so fail immediately.
      return JSON.parse(text);
    } catch (err) {
      lastError = err;

      const message = String(err && err.message ? err.message : err);
      const isTransient =
        message.includes("Transient Gemini error") ||
        message.includes("fetch failed") ||
        message.includes("timeout") ||
        message.includes("temporarily unavailable") ||
        message.includes("high traffic") ||
        message.includes("overloaded") ||
        message.includes("resource exhausted");

      // Retry only transient failures, and only if attempts remain.
      if (isTransient && attempt < maxAttempts) {
        Utilities.sleep(1000 * attempt); // 1s, then 2s backoff
        continue;
      }

      // Stop immediately for non-transient failures like invalid JSON or bad requests.
      throw err;
    }
  }

  throw new Error(`Gemini request failed after ${maxAttempts} attempts: ${lastError.message}`);
}



// Writes a simple tabular dataset into the first sheet of the spreadsheet.
function writeTableToSheet(sheet, headers, rows) {
  // Prepend the header row to the data rows.
  const values = [headers, ...rows];

  // Clear any existing content before writing fresh results.
  sheet.clearContents();

  // Write the full table in one shot for efficiency.
  sheet.getRange(1, 1, values.length, headers.length).setValues(values);

  // Freeze the header row so it stays visible while scrolling.
  sheet.setFrozenRows(1);

  // Make headers bold for readability.
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");

  // Auto-fit columns to the content.
  resizeColumnsToFit(sheet);

  // Add a filter so the table can be sorted/filtered easily.
  sheet.getRange(1, 1, values.length, headers.length).createFilter();
}
