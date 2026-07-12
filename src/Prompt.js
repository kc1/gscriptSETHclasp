function waterPromptStage1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();
  if (!sheetName == "WaterOPTIMIZER") {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Switch to the WaterOPTIMIZER sheet.",
      "promptStage1",
      5,
    );
    return;
  }
  const x = genericPromptStage1("Buildable", "WaterOPTIMIZER", "openRouterPromptWater");
}

function waterStatusMatches() {
  const out = statusCalculator(
    "Buildable",
    "NealsNotes",
    "Status",
    "WaterOPTIMIZER",
  );
  Logger.log(out);
}

function structurePresentUsingLLMTest() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();
  if (!sheetName == "ImprovementOPTIMIZER") {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Switch to an OPTIMIZER sheet.",
      "promptStage1",
      5,
    );
    return;
  }
  const x = genericPromptStage1("StructuresPresent", "ImprovementOPTIMIZER");
}

function roadAvailabilityPromptStage1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();
  if (!sheetName == "OPTIMIZER") {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Switch to an OPTIMIZER sheet.",
      "promptStage1",
      5,
    );
    return;
  }
  const x = genericPromptStage1("RoadAvailable", sheet);
}

function BuildingStatusMatches() {
  const out = statusCalculator(
    "StructuresPresent",
    "NealsNotes",
    "Status",
    "ImprovementOPTIMIZER",
  );
  Logger.log(out);
}

function stage1Status() {
  const out = statusCalculator("Generated", "NealsNotes", "Status");
  Logger.log(out);
}

function statusCalculator(column1, column2, outputColumn, sheetname) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  let currentSheetObjArr = sheet2Json(sheet);
  Logger.log("length: " + currentSheetObjArr.length);

  const toProcess = currentSheetObjArr.filter((row) => {
    return (
      [2, 3].includes(row[column1].length) &&
      [2, 3].includes(row[column2].length) &&
      row.PROMPT &&
      row.ScreenshotURL.includes("http")
    );
  });

  const totalRows = toProcess.length;
  Logger.log(`Rows to process: ${totalRows}`);

  let matchs = 0;
  let mismatchs = 0;
  let errors = 0;
  for (var i = 0; i < toProcess.length; i++) {
    const myRow = toProcess[i];
    if (
      ![2, 3].includes(myRow[column1].length) ||
      ![2, 3].includes(myRow[column2].length)
    ) {
      errors++;
    } else if (myRow[column1].length == myRow[column2].length) {
      var C = updateCell(sheet, myRow, outputColumn, "MATCH");
      matchs++;
    } else {
      var C = updateCell(sheet, myRow, outputColumn, "MISMATCH");
      mismatchs++;
    }
  }
  // const myRow = currentSheetObjArr[92];
  // var C = updateCell(sheet, myRow, outputColumn, matchNum + " MATCHS");

  Logger.log(`Rows processed: ${totalRows}`);
  Logger.log(`Matchs: ${matchs} (${((matchs / totalRows) * 100).toFixed(2)}%)`);
  Logger.log(
    `Mismatches: ${mismatchs} (${((mismatchs / totalRows) * 100).toFixed(2)}%)`,
  );
  Logger.log(`Errors: ${errors} (${((errors / totalRows) * 100).toFixed(2)}%)`);
  return "FIN";
}

// ====================== RUN THIS AFTER Status IS UPDATED ======================
function refineMismatchedPrompts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  let data = sheet2Json(sheet);

  // Find mismatches
  const mismatches = data.filter(
    (row) =>
      String(row.Status).trim().toUpperCase() === "MISMATCH" &&
      row.RoadURL &&
      row.PROMPT &&
      row.NealNotes,
  );

  Logger.log(`Found ${mismatches.length} mismatches to refine.`);

  if (mismatches.length === 0) {
    Logger.log("No mismatches found.");
    return;
  }

  // for (var i = 0; i < mismatches.length; i++) {
  for (var i = 0; i < 2; i++) {
    const myRow = mismatches[i];
    Logger.log(myRow.RoadAvailable);
    Logger.log(`Refining prompt for Row ID: ${myRow.ID}`);
    const improvedPrompt = callPromptOptimizer(myRow);

    if (improvedPrompt && improvedPrompt.trim() !== "") {
      const newVersion = incrementVersion(myRow.PromptVersion);

      var A = updateCell(sheet, myRow, "PROMPT", improvedPrompt);
      var B = updateCell(sheet, myRow, "RoadAvailable", "");
      var C = updateCell(sheet, myRow, "Status", "");
      var E = updateCell(sheet, myRow, "PromptVersion", newVersion);

      Logger.log(`✓ Row ${myRow.ID} prompt updated to ${newVersion}`);
    } else {
      Logger.log(`✗ Failed to get improved prompt for Row ${myRow.ID}`);
    }
  }

  Logger.log("Prompt refinement completed. Now run promptStage1");
}

function genericPromptStage1(responseColumnName, sheetName,netlifyFunctionName="openRouterPrompt2") {
  Logger.log(sheetName);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  const currentSheetObjArr = sheet2Json(sheet);
  Logger.log("first 2 ");
  Logger.log(currentSheetObjArr[0]);
  Logger.log(currentSheetObjArr[1]);

  // Filter rows that haven't been processed yet (no RoadAvailable prompt? Wait, adjust filter as needed)
  // Assuming you want rows where GeneratedAnswer is empty
  const toProcess = currentSheetObjArr.filter((row) => {
    return (
      row.PROMPT &&
      row.ScreenshotURL.includes("http") &&
      row[responseColumnName] == ""
    );
  });

  Logger.log(`Rows to process: ${toProcess.length}`);

  // Process first 10 (or all if you prefer)
  const firstXPushedRows = toProcess.slice(0, 10);

  if (firstXPushedRows.length === 0) {
    Logger.log("No new rows to process.");
    return;
  }

  Logger.log(JSON.stringify(firstXPushedRows));

  const url = APIURL + netlifyFunctionName;
  Logger.log(url);

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      myrows: firstXPushedRows,
      model: "set in netlify",
    }),
    muteHttpExceptions: true,
  };

  Logger.log("payload");
  Logger.log(
    JSON.stringify({
      myrows: firstXPushedRows,
      model: "set in netlify",
    }),
  );
  // model: "stepfun/step-3.7-flash",

  let responseObj;

  try {
    const response = UrlFetchApp.fetch(url, options);
    const httpCode = response.getResponseCode();
    const body = response.getContentText();
    Logger.log(httpCode);
    Logger.log(body);
    responseObj = parseSurveyorResponse(body);
  } catch (error) {
    Logger.log(error);
  }

  Logger.log("HERE:");

  Logger.log(responseObj);

  for (let i = 0; i < responseObj.length; i++) {
    try {
      const myRow = responseObj[i];
      const jsonString = myRow.GeneratedResponse;
      const parsed = parseSurveyorResponse(jsonString);

      let YN = parsed[responseColumnName] || "";
      // let YN = parsed["buildable"] || "";

      var C = updateCell(sheet, myRow, responseColumnName, YN);
    } catch (e) {
      Logger.log(`Error processing row ${i}: ${e}`);
      var C = updateCell(sheet, myRow, responseColumnName, "ERROR");
    }
  }
}

function incrementVersion(current) {
  if (!current || current === "v0") return "v1";
  const num = parseInt(current.replace("v", "")) || 0;
  return `v${num + 1}`;
}

function callPromptOptimizer(row) {
  const optimizerPrompt = `
You are an expert prompt engineer for vision-language models.

Your task is to improve the prompt so the model gives the correct Yes/No answer about road availability.

**Current Prompt** (the one that was just used):
"""
${row.PROMPT}
"""

**Image**: [attached]

**Result**:
- Model answered: "${row.RoadAvailable || "None"}"
- Correct answer (NealNotes): "${row.NealNotes}"

Analyze why it failed and write a **better prompt**.

Focus on:
- Clear definition of what counts as a road
- Strong instruction to answer with only "YES" or "NO"
- Better visual reasoning guidance
- Handling edge cases (far away roads, partial visibility, shadows, etc.)

Return **only** the new improved prompt. 
Do not add any explanations, markdown, or extra text.
`;

  const payload = {
    task: "refine_prompt",
    model: "anthropic/claude-opus-4.6", // ← Add this
    RoadURL: row.RoadURL,
    prompt: optimizerPrompt, // ← This is what gets sent to the model
    currentPrompt: row.PROMPT, // Optional: for reference
    wrongAnswer: row.RoadAvailable,
    correctAnswer: row.NealNotes,
    rowId: row.ID,
  };

  Logger.log("OPTIMIZER");
  Logger.log(optimizerPrompt);

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const res = UrlFetchApp.fetch(APIURL + "openRouterPrompt", options);
    const data = JSON.parse(res.getContentText());
    return data || null;
  } catch (e) {
    Logger.log(`Optimizer failed for row ${row.ID}: ${e}`);
    return null;
  }
}

// ====================== RUN THIS AFTER Status IS UPDATED ======================
function refineMismatchedPrompts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  let data = sheet2Json(sheet);

  // Find mismatches
  const mismatches = data.filter(
    (row) =>
      row.Status.includes("MISMATCH") &&
      row.RoadURL &&
      row.PROMPT &&
      row.NealsNotes,
  );

  Logger.log(`Found ${mismatches.length} mismatches to refine.`);

  if (mismatches.length === 0) {
    Logger.log("No mismatches found.");
    return;
  }

  // for (var i = 0; i < mismatches.length; i++) {
  for (var i = 0; i < 2; i++) {
    const myRow = mismatches[i];
    const improvedPrompt = callPromptOptimizer(myRow);
    Logger.log("improvedPrompt");
    Logger.log(improvedPrompt);

    if (improvedPrompt && improvedPrompt.trim() !== "") {
      const newVersion = incrementVersion(myRow.PromptVersion);
      var C = updateCell(sheet, myRow, "PROMPT", improvedPrompt);
      var C = updateCell(sheet, myRow, "RoadAvailable", "");
      var C = updateCell(sheet, myRow, "Status", "");
      var C = updateCell(sheet, myRow, "PromptVersion", newVersion);
    }
  }

  Logger.log("Prompt refinement completed.");
}

function incrementVersion(currentVersion) {
  if (!currentVersion || currentVersion === "") return "v1";
  const match = currentVersion.match(/v(\d+)/);
  const num = match ? parseInt(match[1]) : 0;
  return `v${num + 1}`;
}

/**
 * Builds a distilled road-availability prompt from MATCH rows on OPTIMIZER.
 * Sends a text-only request via openRouterPrompt (task create_unified_prompt).
 */
function createUnifiedPrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("OPTIMIZER");
  if (!sheet) {
    Logger.log("OPTIMIZER sheet missing.");
    try {
      SpreadsheetApp.getUi().alert(
        'Sheet "OPTIMIZER" not found. Rename or activate the optimizer sheet.',
      );
    } catch (e) {
      /* no UI context (trigger / headless) */
    }
    return;
  }

  const data = sheet2Json(sheet);

  const successfulRows = data.filter(function (row) {
    const status = String(row.STATUS || row.Status || "")
      .trim()
      .toUpperCase();
    return (
      status === "MATCH" &&
      row.PROMPT &&
      String(row.PROMPT).trim().length > 40 &&
      row.NealNotes
    );
  });

  Logger.log(`Found ${successfulRows.length} successful MATCH examples.`);

  if (successfulRows.length < 8) {
    Logger.log(
      "Not enough successful examples for unified prompt (need at least 8).",
    );
    try {
      SpreadsheetApp.getUi().alert(
        "Need at least 8 MATCH rows on OPTIMIZER (with PROMPT > 40 chars and NealNotes).",
      );
    } catch (e) {}
    return;
  }

  const sample = successfulRows.slice(0, 25);

  var distillationPrompt =
    "\nYou are an expert prompt engineer specializing in vision-language models.\n\n" +
    "I have run many experiments asking whether a road is available on parcel images.\n" +
    "Here are " +
    sample.length +
    " prompts that produced answers matching the ground truth (NealNotes):\n\n" +
    sample
      .map(function (row, i) {
        return (
          "\n=== Example " +
          (i + 1) +
          ' ===\nPrompt:\n"""\n' +
          row.PROMPT +
          '\n"""\nCorrect answer (NealNotes): ' +
          row.NealNotes +
          "\n"
        );
      })
      .join("\n") +
    "\n\nCreate the single best **unified prompt** that should work consistently across similar images.\n\n" +
    "Requirements:\n" +
    '- Clear definition of what counts as a "road" and what "available" means for the lot\n' +
    '- Instruction to follow the same answer format as in promptStage1 (reasoning, separator, then JSON with key RoadAvailable: "Yes" or "No" only — match how your working prompts are structured)\n' +
    "- Step-by-step visual reasoning guidance\n" +
    "- Edge cases: distant roads, partial visibility, shadows, unpaved paths, adjacency vs inside lot\n" +
    "- Concise but effective; suitable for fast vision models (e.g. Gemini Flash)\n\n" +
    "Return **ONLY** the new unified prompt text. No explanations, markdown fences, or preamble.\n";

  const payload = {
    task: "create_unified_prompt",
    textOnly: true,
    prompt: distillationPrompt,
    model: "anthropic/claude-opus-4.6",
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const url = APIURL + "openRouterPrompt";
  Logger.log(url);

  try {
    const response = UrlFetchApp.fetch(url, options);
    const httpCode = response.getResponseCode();
    const body = response.getContentText();

    if (httpCode < 200 || httpCode >= 300) {
      Logger.log("createUnifiedPrompt HTTP " + httpCode + ": " + body);
      try {
        SpreadsheetApp.getUi().alert(
          "openRouterPrompt failed (HTTP " + httpCode + "). See execution log.",
        );
      } catch (e) {}
      return;
    }

    var parsed;
    try {
      parsed = JSON.parse(body);
    } catch (eParse) {
      parsed = body;
    }

    var unifiedPrompt = "";
    if (typeof parsed === "string") {
      unifiedPrompt = parsed.trim();
    } else if (parsed && parsed.error) {
      Logger.log("openRouterPrompt error: " + parsed.error);
      try {
        SpreadsheetApp.getUi().alert("API error: " + parsed.error);
      } catch (e) {}
      return;
    } else {
      unifiedPrompt = String(
        parsed.unifiedPrompt ||
          parsed.prompt ||
          parsed.text ||
          parsed.content ||
          "",
      ).trim();
    }

    if (unifiedPrompt && unifiedPrompt.length > 100) {
      // sheet.getRange("I1").setValue("=== UNIFIED PROMPT ===");
      sheet.getRange("I2").setValue(unifiedPrompt);
      Logger.log(
        "Unified prompt written to I2, length: " + unifiedPrompt.length,
      );
      try {
        SpreadsheetApp.getUi().alert(
          "Unified prompt created. See cells I1–I2 on OPTIMIZER.",
        );
      } catch (e) {}
    } else {
      Logger.log(
        "Failed to extract unified prompt from response; body length: " +
          (body && body.length),
      );
      try {
        SpreadsheetApp.getUi().alert(
          "Could not parse a long enough unified prompt. Check execution log for raw body.",
        );
      } catch (e) {}
    }
  } catch (error) {
    Logger.log("Error in createUnifiedPrompt: " + error);
    try {
      SpreadsheetApp.getUi().alert("Error: " + error);
    } catch (e) {}
  }
}

/**
 * Builds a unified prompt from MATCH rows, prioritizing higher PromptVersion (refined prompts).
 * Same API path and parsing as createUnifiedPrompt; differs only in row ordering and example labels.
 */
function createUnifiedPrompt2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("OPTIMIZER");
  if (!sheet) {
    Logger.log("OPTIMIZER sheet missing.");
    try {
      SpreadsheetApp.getUi().alert(
        'Sheet "OPTIMIZER" not found. Rename or activate the optimizer sheet.',
      );
    } catch (e) {}
    return;
  }

  const data = sheet2Json(sheet);

  function promptVersionNum(row) {
    const m = String(row.PromptVersion || "v0")
      .toLowerCase()
      .match(/v(\d+)/);
    return m ? parseInt(m[1], 10) : 0;
  }

  const successfulRows = data
    .filter(function (row) {
      const status = String(row.STATUS || row.Status || row.Status1 || "")
        .trim()
        .toUpperCase();
      return (
        status === "MATCH" &&
        row.PROMPT &&
        String(row.PROMPT).trim().length > 40 &&
        row.NealsNotes
      );
    })
    .sort(function (a, b) {
      return promptVersionNum(b) - promptVersionNum(a);
    });

  Logger.log(
    `Found ${successfulRows.length} MATCH rows (sorted by PromptVersion desc).`,
  );

  if (successfulRows.length < 8) {
    Logger.log(
      "Not enough successful examples for unified prompt (need at least 8).",
    );
    try {
      SpreadsheetApp.getUi().alert(
        "Need at least 8 MATCH rows on OPTIMIZER (with PROMPT > 40 chars and NealNotes).",
      );
    } catch (e) {}
    return;
  }

  const sample = successfulRows.slice(0, 25);

  var distillationPrompt =
    "\nYou are an expert prompt engineer specializing in vision-language models.\n\n" +
    "I have run many experiments asking whether a road is available on parcel images.\n" +
    "Below are " +
    sample.length +
    " prompts that matched ground truth (NealNotes), ordered with **newer PromptVersion first** " +
    "(more refined prompts appear earlier):\n\n" +
    sample
      .map(function (row, i) {
        return (
          "\n=== Example " +
          (i + 1) +
          " [" +
          (row.PromptVersion || "v?") +
          '] ===\nPrompt:\n"""\n' +
          row.PROMPT +
          '\n"""\nCorrect answer (NealNotes): ' +
          row.NealNotes +
          "\n"
        );
      })
      .join("\n") +
    "\n\nCreate the single best **unified prompt** that should work consistently across similar images.\n\n" +
    "Requirements:\n" +
    '- Clear definition of what counts as a "road" and what "available" means for the lot\n' +
    '- Instruction to follow the same answer format as in promptStage1 (reasoning, separator, then JSON with key RoadAvailable: "Yes" or "No" only — match how your working prompts are structured)\n' +
    "- Step-by-step visual reasoning guidance\n" +
    "- Edge cases: distant roads, partial visibility, shadows, unpaved paths, adjacency vs inside lot\n" +
    "- Concise but effective; suitable for fast vision models (e.g. Gemini Flash)\n\n" +
    "Return **ONLY** the new unified prompt text. No explanations, markdown fences, or preamble.\n";

  const payload = {
    task: "create_unified_prompt",
    textOnly: true,
    prompt: distillationPrompt,
    model: "anthropic/claude-opus-4.6",
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const url = APIURL + "openRouterPrompt";
  Logger.log(url);

  try {
    const response = UrlFetchApp.fetch(url, options);
    const httpCode = response.getResponseCode();
    const body = response.getContentText();

    if (httpCode < 200 || httpCode >= 300) {
      Logger.log("createUnifiedPrompt2 HTTP " + httpCode + ": " + body);
      try {
        SpreadsheetApp.getUi().alert(
          "openRouterPrompt failed (HTTP " + httpCode + "). See execution log.",
        );
      } catch (e) {}
      return;
    }

    var parsed;
    try {
      parsed = JSON.parse(body);
    } catch (eParse) {
      parsed = body;
    }

    var unifiedPrompt = "";
    if (typeof parsed === "string") {
      unifiedPrompt = parsed.trim();
    } else if (parsed && parsed.error) {
      Logger.log("openRouterPrompt error: " + parsed.error);
      try {
        SpreadsheetApp.getUi().alert("API error: " + parsed.error);
      } catch (e) {}
      return;
    } else {
      unifiedPrompt = String(
        parsed.unifiedPrompt ||
          parsed.prompt ||
          parsed.text ||
          parsed.content ||
          "",
      ).trim();
    }

    if (unifiedPrompt && unifiedPrompt.length > 100) {
      sheet.getRange("I2").setValue(unifiedPrompt);
      Logger.log(
        "Unified prompt (v2, version-weighted) written to I2, length: " +
          unifiedPrompt.length,
      );
      try {
        SpreadsheetApp.getUi().alert(
          "Unified prompt created (refined examples first). See I2 on OPTIMIZER.",
        );
      } catch (e) {}
    } else {
      Logger.log(
        "Failed to extract unified prompt from response; body length: " +
          (body && body.length),
      );
      try {
        SpreadsheetApp.getUi().alert(
          "Could not parse a long enough unified prompt. Check execution log for raw body.",
        );
      } catch (e) {}
    }
  } catch (error) {
    Logger.log("Error in createUnifiedPrompt2: " + error);
    try {
      SpreadsheetApp.getUi().alert("Error: " + error);
    } catch (e) {}
  }
}

function resetPROMPTcolumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  let rows = sheet2Json(sheet);

  let filtered = rows.filter((row) => row.PROMPT.length > 20);
  Logger.log(filtered);
  const newPrompt = rows[0].UNIFIEDPROMPT;
  for (let i = 0; i < filtered.length; i++) {
    const myRow = filtered[i];
    let a = updateCell(sheet, myRow, "PROMPT", newPrompt);
  }
}
