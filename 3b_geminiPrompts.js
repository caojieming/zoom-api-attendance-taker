/**
 * creates a sheet listing the key topics of a meeting by sending the transcript along with a prompt to Gemini
 */
function createKeyTopicsSheet(driveFolderId, fileName, transcript) {
  // Prompt tells the model exactly what JSON shape to return and how to behave.
  const prompt = `
You are an expert thematic analyst specializing in transcription synthesis and key topic extraction.

Return ONLY valid JSON in this exact shape:
{
  "topics": [
    {
      "key_topic": "short concise topic name",
      "summary": "1-2 sentence objective, factual summary",
      "significance": 1-5,
      "confidence": 1-5
    }
  ]
}

Rules:
- Keep topics concise, distinct, and mutually exclusive.
- Merge duplicates and near-duplicates into a single representative topic.
- Determine significance (1-5) based on the duration of the discussion, the frequency of mention, and the topic's impact on the conversation's overall objective.
- Prioritize central themes over minor details or tangential remarks.
- Output between 5 and 20 topics unless the transcript clearly supports fewer or more.
- confidence (1-5) reflects your certainty that the topic was a primary focus of the speakers.
- Ensure all strings are properly escaped for JSON compliance to avoid syntax errors.
- No markdown, no code fences, no preamble, and no extra text.

Transcript:
${transcript}
`;

  // Delegate all shared work to the generic helper.
  return createExtractionSheet_({
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
 * creates a summary doc, following the format/logic of zoom summaries
 */
function createSummaryDoc(driveFolderId, fileName, transcript) {
  // Prompt tells the model to generate a Zoom-style meeting summary in a clean doc-friendly format.
  const prompt = `
You are an expert Executive Assistant specializing in corporate governance and project management.

Task:
Generate a concise, accurate, and highly scanable summary of the meeting transcript provided below.

Instructions:
- Use only information explicitly supported by the transcript; do not invent details or assume outcomes.
- Filter out all conversational filler, social pleasantries, and technical interruptions to focus on high-value information.
- Distinguish clearly between a proposal/suggestion and a final, confirmed decision.
- Attribute key points, decisions, and action items to the specific speaker(s) involved.
- Use professional, objective language and avoid adjectives that imply emotion unless the emotional tone of the meeting is a critical data point.
- Use bolding for names, deadlines, and key deliverables to improve scanability.
- If a specific detail (like a due date) is missing, write "Not specified."

Output format:

Title: [Meeting title if known, otherwise just use "Meeting Summary"]

Summary:
[2-5 sentence high-level overview of the meeting's purpose and primary outcome]

Key Discussion Points:
- [Point 1] ([Speaker Name])
- [Point 2] ([Speaker Name])
- [Point 3] ([Speaker Name])

Decisions Made:
- [Decision 1] - Confirmed by [Name/Group]
- [Decision 2] - Confirmed by [Name/Group]

Action Items:
- [Task] | Owner: **[Name]** | Due: **[Date/Not specified]**

Open Questions / Risks:
- [Question or risk item] | Raised by: [Name]

Notes:
- [Optional extra important details or context]

Transcript:
${transcript}
`;

  // Delegate the Gemini call and doc creation to a shared helper.
  return createTextDocFromPrompt_({
    driveFolderId,
    fileName,
    prompt
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
      "email": "participant email",
      "phone": "participant phone number",
      "linkedin_url": "full LinkedIn URL or empty string",
      "chapter": "city or state name or empty string",
      "self_introduction": "brief self-introduction extracted or inferred from the chat",
      "summary": "brief summary of what the participant contributed to the chat, extracted or inferred from the chat"
    }
  ]
}

Rules:
- Return one entry per unique participant.
- Include only people who appear to be participants in the chat.
- name: use the person's stated name if available; otherwise use the best available identifier.
- email: fill in only if explicitly found in the chat; otherwise use an empty string.
- phone: fill in only if explicitly found in the chat; otherwise use an empty string.
- linkedin_url: fill in only if explicitly found in the chat; otherwise use an empty string.
- chapter: generally a city or state name; fill in only if explicitly found in the chat; otherwise use an empty string.
- self_introduction: should be short and based on how the participant introduces themselves in the chat.
- If a participant does not self-introduce, infer a concise one from the available context, or use an empty string if none is available.
- summary: should be a no more than 3 sentences long, and should be about what the participant contributed to the chat log
- If the participant shared a URL link, be sure to include the full link in the summary (unless it's a LinkedIn URL, then don't include it to the summary).
- Merge duplicate references to the same person.
- Do not include people mentioned only in passing unless they clearly participate.
- Output between 1 and 50 participants unless the chat clearly supports fewer or more.
- No markdown, no code fences, no extra text.

Chat log:
${chatLog}
`;

  // Delegate all shared work to the generic helper.
  return createExtractionSheet_({
    driveFolderId,
    fileName,
    prompt,
    rootKey: "participants",
    headers: ["name", "email", "phone", "linkedin_url", "chapter", "self_introduction", "summary"],
    rowMapper: p => [
      p.name || "",
      p.email || "",
      p.phone || "",
      p.linkedin_url || "",
      p.chapter || "",
      p.self_introduction || "",
      p.summary || ""
    ],
    countKey: "participantsCount"
  });
}


// Generic extraction runner shared by both use cases.
function createExtractionSheet_({
  driveFolderId,
  fileName,
  prompt,
  rootKey,
  headers,
  rowMapper,
  countKey
}) {
  // Create the spreadsheet and use its first sheet as the output destination.
  const spreadsheet = createSpreadsheetInFolder_(driveFolderId, fileName);
  const sheet = spreadsheet.getSheets()[0];

  // Send the prompt to Gemini and parse the JSON response.
  const parsed = callGeminiJson_(prompt);

  // Pull out the expected array from the returned JSON.
  const items = Array.isArray(parsed[rootKey]) ? parsed[rootKey] : [];

  // Convert each item into a row of sheet values.
  const rows = items.map(rowMapper);

  // Write the final table to the sheet.
  writeTableToSheet_(sheet, headers, rows);

  // Return useful metadata about the generated spreadsheet.
  return {
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    [countKey]: items.length
  };
}


/**
 * Generic doc creator shared by doc-based Gemini outputs.
 * This keeps the Gemini call and file creation logic consistent with the sheet flows.
 */
function createTextDocFromPrompt_({ driveFolderId, fileName, prompt }) {
  // Send the prompt to Gemini and get plain text back.
  const text = callGeminiText_(prompt);

  // Create the Google Doc in the target folder with the generated content.
  const docId = createDocInFolder_(driveFolderId, fileName, text);

  // Return useful metadata about the generated document.
  return {
    docId: docId,
    docUrl: `https://docs.google.com/document/d/${docId}/edit`
  };
}


// Calls Gemini and returns parsed JSON, retrying only on transient "high traffic" style failures.
function callGeminiJson_(prompt) {
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


/**
 * Calls Gemini and returns plain text, retrying only on transient "high traffic" style failures.
 */
function callGeminiText_(prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;

  const payload = {
    contents: [{
      role: "user",
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: 0.2,
      topP: 0.95,
      maxOutputTokens: 8192
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

      return text;
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

      // Stop immediately for non-transient failures.
      throw err;
    }
  }

  throw new Error(`Gemini request failed after ${maxAttempts} attempts: ${lastError.message}`);
}


// Writes a simple tabular dataset into the first sheet of the spreadsheet.
function writeTableToSheet_(sheet, headers, rows) {
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

  // Add a filter so the table can be sorted/filtered easily.
  sheet.getRange(1, 1, values.length, headers.length).createFilter();

  // Auto-fit columns to the content.
  resizeColumnsToFit_(sheet);
}


/**
 * Resizes all populated columns in a sheet so content is easier to read.
 */
function resizeColumnsToFit_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn > 0) {
    sheet.autoResizeColumns(1, lastColumn);
  }
}
