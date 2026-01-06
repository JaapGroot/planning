/**
 * Planning Team Splitter - Performance/Robust v3
 * - Batched sync with triggers (avoids max execution time)
 * - Locking to prevent concurrent runs
 * - Drive cache (scan parent folder once)
 * - Split master reads:
 *    values: A..L
 *    info backgrounds: A..G
 *    week backgrounds: H..
 * - Fast write keeping your overwrite rule:
 *    info (A..G) always master
 *    weeks (H..) only overwrite if master != white AND (existing is white OR existing == master)
 * - Smart clearing: only clears up to last written row (stored per team sheet)
 * - One-time sheet setup (widths/merge/protect/hide/header) stored in properties
 * - Force reset helpers for setup
 */

const CONFIG = {
  TOTAL_COLS: 65,
  WEEK_START_COL: 8,              // H
  INFO_COLS_END: 7,               // A..G
  TEAM_NAME_CELL: "A1",
  INFO_HEADER_ROW: 3,
  DATA_START_ROW: 4,
  HEADER_ROWS: 3,
  PLANNING_SHEET: "Planning",
  LOG_SHEET_NAME: "Log",
  TEAM_SHEET_DEFAULTS: ["Blad1", "Sheet1"],

  // 0-based indices for VALUES array A..L (12 cols)
  COL_WERKNUMMER: 0,              // A
  COL_OPDRACHT: 5,                // F
  COL_TEAM: 11,                   // L

  EXCLUDED_TEAM_WORDS: ["team", "ja", "nee", "opdracht"],

  // Batch settings (as requested: not changing teams per run logic)
  TEAMS_PER_RUN: 5,
  BATCH_TRIGGER_MINUTES: 2,

  // Properties keys
  PROP_BATCH_INDEX: "PTS_BATCH_INDEX",
  PROP_BATCH_TEAMS: "PTS_BATCH_TEAMS",

  // Setup + last row keys are generated per file/sheet
  PROP_SETUP_PREFIX: "PTS_SETUP_",           // + fileId + "_" + sheetId
  PROP_LASTROW_PREFIX: "PTS_LASTROW_",       // + fileId + "_" + sheetId
};

/* =========================
 * PUBLIC ENTRYPOINTS
 * ========================= */

function showTeamDropdown() {
  withScriptLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName(CONFIG.PLANNING_SHEET);

    const teams = getTeamsFromPlanning_(planningSheet);
    if (!teams.length) {
      SpreadsheetApp.getUi().alert("Geen teamnamen gevonden.");
      return;
    }

    const html = `
    <html><body>
      <form onsubmit="google.script.run.startSingleTeamUpdate(this.team.value);google.script.host.close();return false;">
        <label>Selecteer een team:</label><br>
        <select name="team">
          ${teams.map(t => `<option value="${escapeHtml_(t)}">${escapeHtml_(t)}</option>`).join("")}
        </select>
        <br><br><input type="submit" value="Bijwerken">
      </form>
    </body></html>`;

    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(html).setWidth(320).setHeight(180),
      "Team bijwerken"
    );
  });
}

function startSingleTeamUpdate(teamName) {
  withScriptLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName(CONFIG.PLANNING_SHEET);

    const cache = buildTeamFileCache_(ss);
    const count = updateTeamFile_(ss, planningSheet, teamName, cache);

    SpreadsheetApp.getUi().alert(
      `✅ Update voor ${teamName} voltooid\nAantal regels: ${count}\nMode: Performance/Robust v3`
    );
  });
}

/**
 * Start batched sync for all teams.
 * Run manually once; it schedules its own next batches.
 */
function syncAllTeamsBatched() {
  withScriptLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName(CONFIG.PLANNING_SHEET);

    const teams = getTeamsFromPlanning_(planningSheet);
    if (!teams.length) {
      Logger.log("Geen teams gevonden.");
      return;
    }

    const props = PropertiesService.getDocumentProperties();
    props.setProperty(CONFIG.PROP_BATCH_TEAMS, JSON.stringify(teams));
    props.setProperty(CONFIG.PROP_BATCH_INDEX, "0");

    cleanupBatchTriggers_(); // avoid duplicates
    runNextBatch_();         // run first batch immediately
  });
}

/**
 * Handler for scheduled batches (time-based trigger).
 * IMPORTANT: must be global function name for trigger.
 */
function runNextBatch_() {
  withScriptLock_(() => {
    const props = PropertiesService.getDocumentProperties();
    const teamsJson = props.getProperty(CONFIG.PROP_BATCH_TEAMS);
    if (!teamsJson) {
      Logger.log("Geen batch data (teams) in properties.");
      cleanupBatchTriggers_();
      return;
    }

    const teams = JSON.parse(teamsJson);
    const index = parseInt(props.getProperty(CONFIG.PROP_BATCH_INDEX) || "0", 10);

    if (index >= teams.length) {
      Logger.log("✅ Alle teams klaar. Batch afgerond.");
      clearBatchState_();
      cleanupBatchTriggers_();
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName(CONFIG.PLANNING_SHEET);

    // Cache files once per batch run
    const cache = buildTeamFileCache_(ss);

    const end = Math.min(index + CONFIG.TEAMS_PER_RUN, teams.length);
    Logger.log(`Batch: teams ${index + 1}..${end} van ${teams.length}`);

    for (let i = index; i < end; i++) {
      const team = teams[i];
      try {
        const rows = updateTeamFile_(ss, planningSheet, team, cache);
        Logger.log(`✅ ${team} bijgewerkt (${rows} regels)`);
      } catch (err) {
        Logger.log(`❌ Fout bij team ${team}: ${err && err.message ? err.message : err}`);
      }
    }

    props.setProperty(CONFIG.PROP_BATCH_INDEX, String(end));

    if (end < teams.length) {
      scheduleNextBatchTrigger_();
    } else {
      Logger.log("✅ Laatste batch klaar.");
      clearBatchState_();
      cleanupBatchTriggers_();
    }
  });
}

/**
 * Force setup opnieuw voor 1 team (of huidige sheet), handig na header/kolombreedte changes.
 */
function forceResetTeamSetup(teamName) {
  withScriptLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cache = buildTeamFileCache_(ss);

    // Als teamName niet is opgegeven, vraag via prompt
    if (!teamName) {
      const ui = SpreadsheetApp.getUi();
      const res = ui.prompt("Team setup resetten", "Voer teamnaam in:", ui.ButtonSet.OK_CANCEL);
      if (res.getSelectedButton() !== ui.Button.OK) return;
      teamName = res.getResponseText().trim();
      if (!teamName) return;
    }

    const teamFile = openOrCreateTeamFileCached_(ss, teamName, cache);
    const teamSheet = getOrCreateTeamSheet_(teamFile, teamName);

    const props = PropertiesService.getDocumentProperties();
    props.deleteProperty(setupKey_(teamFile, teamSheet));

    SpreadsheetApp.getUi().alert(`✅ Setup reset flag verwijderd voor team: ${teamName}\nVolgende update doet volledige setup opnieuw.`);
  });
}

/**
 * Force setup opnieuw voor ALLE team sheets die in de parent folder staan met naam "Planning - ..."
 */
function forceResetAllTeamSetups() {
  withScriptLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cache = buildTeamFileCache_(ss);

    const props = PropertiesService.getDocumentProperties();
    const names = Object.keys(cache.byName).filter(n => n.startsWith("Planning - "));

    let count = 0;
    for (const name of names) {
      try {
        const file = SpreadsheetApp.openById(cache.byName[name]);
        // reset voor alle sheets in dat bestand
        file.getSheets().forEach(sh => {
          props.deleteProperty(setupKeyRaw_(file.getId(), sh.getSheetId()));
          props.deleteProperty(lastRowKeyRaw_(file.getId(), sh.getSheetId()));
        });
        count++;
      } catch (e) {
        Logger.log(`Reset fail for ${name}: ${e && e.message ? e.message : e}`);
      }
    }

    SpreadsheetApp.getUi().alert(`✅ Reset flags gezet voor ${count} team-bestanden.\nVolgende updates doen setup opnieuw.`);
  });
}

/**
 * Cleanup functie voor als je een batch “vast” hebt staan.
 */
function stopAndClearBatchState() {
  withScriptLock_(() => {
    clearBatchState_();
    cleanupBatchTriggers_();
    SpreadsheetApp.getUi().alert("✅ Batch state + triggers opgeschoond.");
  });
}

/* =========================
 * CORE PIPELINE
 * ========================= */

function updateTeamFile_(ss, planningSheet, teamName, fileCache) {
  const lastRow = planningSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;

  // ✅ Split reads:
  // Values A..L (12 cols)
  const valuesAL = planningSheet
    .getRange(CONFIG.DATA_START_ROW, 1, numRows, 12)
    .getValues();

  // Info bgs A..G (7 cols)
  const bgInfoAG = planningSheet
    .getRange(CONFIG.DATA_START_ROW, 1, numRows, CONFIG.INFO_COLS_END)
    .getBackgrounds();

  // Week bgs H..TOTAL_COLS
  const numWeekCols = CONFIG.TOTAL_COLS - CONFIG.WEEK_START_COL + 1;
  const bgWeeks = planningSheet
    .getRange(CONFIG.DATA_START_ROW, CONFIG.WEEK_START_COL, numRows, numWeekCols)
    .getBackgrounds();

  const werkMap = buildWerknummerMapSplit_(valuesAL, bgInfoAG, bgWeeks, teamName);

  const teamFile = openOrCreateTeamFileCached_(ss, teamName, fileCache);
  const teamSheet = getOrCreateTeamSheet_(teamFile, teamName);

  // One-time expensive setup
  ensureTeamSheetSetup_(teamFile, teamSheet, planningSheet, teamName);

  // Fast write, but keeps overwrite rule by reading ONLY week backgrounds
  writeTeamDataRobust_(teamFile, teamSheet, werkMap);

  // Lightweight log
  logSkippedFast_(werkMap.skipped, teamFile);

  return Object.keys(werkMap.data).length;
}

/* =========================
 * TEAM EXTRACTION (skip kopregels per werknummer)
 * kopregel = werknummer wijzigt t.o.v. vorige rij
 * ========================= */

function getTeamsFromPlanning_(planningSheet) {
  const lastRow = planningSheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = planningSheet.getRange(2, 1, lastRow - 1, 12).getValues(); // A..L
  const excluded = new Set(CONFIG.EXCLUDED_TEAM_WORDS.map(s => s.toLowerCase()));
  const teamsSet = new Set();

  let prevWn = null;

  for (let i = 0; i < rows.length; i++) {
    const wn = rows[i][CONFIG.COL_WERKNUMMER];
    const team = rows[i][CONFIG.COL_TEAM];

    const isHeaderRow = wn && wn !== prevWn;
    if (wn) prevWn = wn;

    if (isHeaderRow) continue;

    if (typeof team === "string") {
      const t = team.trim();
      if (t && !excluded.has(t.toLowerCase())) teamsSet.add(t);
    }
  }

  return [...teamsSet].sort();
}

/* =========================
 * BUILD WERKNUMMER MAP (split inputs)
 * - logs "No team match" max 1x per werknummer
 * ========================= */

function buildWerknummerMapSplit_(valuesAL, bgInfoAG, bgWeeks, teamName) {
  const data = {};
  const skipped = [];

  const teamNorm = normalize_(teamName);

  let currentWn = null;
  let include = false;
  let firstRowInfo = null;
  let hadMatch = false;

  for (let r = 0; r < valuesAL.length; r++) {
    const rowVals = valuesAL[r];
    const wn = rowVals[CONFIG.COL_WERKNUMMER];
    const opdracht = rowVals[CONFIG.COL_OPDRACHT];
    const team = rowVals[CONFIG.COL_TEAM];

    const isNewWerknummer = wn && wn !== currentWn;

    if (isNewWerknummer) {
      // close previous
      if (currentWn && include && !hadMatch) {
        skipped.push([new Date(), teamName, currentWn, "No team match"]);
      }

      currentWn = wn;
      hadMatch = false;

      include = typeof opdracht === "string" && normalize_(opdracht) === "ja";

      firstRowInfo = {
        data: rowVals.slice(0, CONFIG.INFO_COLS_END),            // A..G values
        colors: bgInfoAG[r].slice(0, CONFIG.INFO_COLS_END),      // A..G bgs
      };

      if (!include) {
        skipped.push([new Date(), teamName, wn, "Opdracht != Ja"]);
      }
      continue; // kopregel is never a detail row
    }

    if (!currentWn || !include) continue;

    if (typeof team === "string" && normalize_(team) === teamNorm) {
      hadMatch = true;

      if (!data[currentWn]) {
        data[currentWn] = {
          infoCols: firstRowInfo.data,
          colorCols: firstRowInfo.colors,
          weeks: new Array(bgWeeks[0].length).fill("#ffffff"),
        };
      }

      const weekRow = bgWeeks[r];
      for (let c = 0; c < weekRow.length; c++) {
        const col = weekRow[c];
        if (col !== "#ffffff") data[currentWn].weeks[c] = col;
      }
    }
  }

  // close last
  if (currentWn && include && !hadMatch) {
    skipped.push([new Date(), teamName, currentWn, "No team match"]);
  }

  return { data, skipped };
}

/* =========================
 * DRIVE / FILE CACHE
 * ========================= */

function buildTeamFileCache_(ss) {
  const parent = DriveApp.getFileById(ss.getId()).getParents().next();
  const files = parent.getFiles();

  const cache = {};
  while (files.hasNext()) {
    const f = files.next();
    cache[f.getName()] = f.getId();
  }
  return { parentId: parent.getId(), byName: cache };
}

function openOrCreateTeamFileCached_(ss, teamName, cache) {
  const name = "Planning - " + teamName;
  const existingId = cache.byName[name];
  if (existingId) return SpreadsheetApp.openById(existingId);

  const file = SpreadsheetApp.create(name);
  const created = DriveApp.getFileById(file.getId());

  const parent = DriveApp.getFolderById(cache.parentId);
  parent.addFile(created);
  DriveApp.getRootFolder().removeFile(created);

  cache.byName[name] = file.getId();
  return file;
}

/* =========================
 * TEAM SHEET SETUP (one-time)
 * ========================= */

function getOrCreateTeamSheet_(file, teamName) {
  let sheet = file.getSheetByName(teamName);
  if (!sheet) sheet = file.insertSheet(teamName);

  // Remove default tabs (cheap)
  CONFIG.TEAM_SHEET_DEFAULTS.forEach(n => {
    const def = file.getSheetByName(n);
    if (def) file.deleteSheet(def);
  });

  // Ensure columns exist (minimal)
  while (sheet.getMaxColumns() < CONFIG.TOTAL_COLS) {
    sheet.insertColumnAfter(sheet.getMaxColumns());
  }

  return sheet;
}

function ensureTeamSheetSetup_(file, sheet, planningSheet, teamName) {
  const props = PropertiesService.getDocumentProperties();
  const key = setupKey_(file, sheet);
  if (props.getProperty(key) === "1") return;

  sheet.getRange(CONFIG.TEAM_NAME_CELL).setValue(teamName);
  sheet.setFrozenRows(CONFIG.HEADER_ROWS);
  sheet.setFrozenColumns(CONFIG.INFO_COLS_END);

  // Column widths (expensive but once)
  for (let c = 1; c <= CONFIG.TOTAL_COLS; c++) {
    sheet.setColumnWidth(c, planningSheet.getColumnWidth(c));
  }

  // Header A..G values (as in original)
  sheet.getRange(CONFIG.INFO_HEADER_ROW, 1, 1, CONFIG.INFO_COLS_END)
    .setValues(planningSheet.getRange(CONFIG.DATA_START_ROW, 1, 1, CONFIG.INFO_COLS_END).getDisplayValues());

  // Hide cols (once)
  sheet.hideColumns(6);
  sheet.hideColumns(8, 6);

  // Week header rows 1..3, cols H..
  const numWeekCols = CONFIG.TOTAL_COLS - CONFIG.WEEK_START_COL + 1;
  const hdr = planningSheet.getRange(1, CONFIG.WEEK_START_COL, 3, numWeekCols);
  const dest = sheet.getRange(1, CONFIG.WEEK_START_COL, 3, numWeekCols);

  dest.setValues(hdr.getDisplayValues());
  dest.setBackgrounds(hdr.getBackgrounds());
  dest.setFontWeights(hdr.getFontWeights());
  dest.setFontSizes(hdr.getFontSizes());
  dest.setFontColors(hdr.getFontColors());
  dest.setHorizontalAlignments(hdr.getHorizontalAlignments());

  mergeRepeatingHeaderCells_(sheet, 1, CONFIG.WEEK_START_COL, numWeekCols);
  mergeRepeatingHeaderCells_(sheet, 2, CONFIG.WEEK_START_COL, numWeekCols);

  protectSheetOnce_(sheet);

  if (!sheet.getFilter()) {
    sheet.getRange(CONFIG.INFO_HEADER_ROW, 1, 1, CONFIG.INFO_COLS_END).createFilter();
  }

  props.setProperty(key, "1");
}

function mergeRepeatingHeaderCells_(sheet, row, startCol, numCols) {
  const rowVals = sheet.getRange(row, startCol, 1, numCols).getDisplayValues()[0];
  let mergeStart = null;

  for (let i = 0; i < rowVals.length; i++) {
    const val = rowVals[i];
    const next = rowVals[i + 1];

    if (val && mergeStart === null) mergeStart = startCol + i;

    const isEnd = mergeStart !== null && (next || i === rowVals.length - 1);
    if (isEnd) {
      const mergeEnd = startCol + i;
      const width = mergeEnd - mergeStart + 1;
      if (width > 1) sheet.getRange(row, mergeStart, 1, width).mergeAcross();
      mergeStart = null;
    }
  }
}

function protectSheetOnce_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  sheet.getRange("1:3").protect().addEditor(Session.getEffectiveUser());
  sheet.getRange("A:G").protect().addEditor(Session.getEffectiveUser());
}

/* =========================
 * FAST + ROBUST WRITE
 * - smart clear only up to last written row
 * - keep overwrite rule for week colors by reading only H.. backgrounds
 * ========================= */

function writeTeamDataRobust_(file, sheet, werkMap) {
  const entries = Object.keys(werkMap.data);
  const props = PropertiesService.getDocumentProperties();

  const lastKey = lastRowKey_(file, sheet);
  const prevLast = parseInt(props.getProperty(lastKey) || "0", 10);

  // Determine clear height: clear previous data rows only (not entire maxRows)
  // If never written before, clear a small default block just in case.
  const clearRows = Math.max(prevLast, 50); // robust default
  const clearStart = CONFIG.DATA_START_ROW;

  // Clear content + basic formatting only for the rows we previously touched
  const clearRange = sheet.getRange(clearStart, 1, clearRows, CONFIG.TOTAL_COLS);
  clearRange.clearContent();
  clearRange.setBackground("#ffffff");
  clearRange.setFontColor("#000000");

  if (!entries.length) {
    props.setProperty(lastKey, "0");
    return;
  }

  // Build arrays
  const outValues = [];
  const outInfoBgs = [];
  const outWeekBgs = [];

  const outWeekCols = CONFIG.TOTAL_COLS - CONFIG.WEEK_START_COL + 1;

  for (let i = 0; i < entries.length; i++) {
    const wn = entries[i];
    const item = werkMap.data[wn];

    const row = item.infoCols.slice();
    while (row.length < CONFIG.TOTAL_COLS) row.push("");
    outValues.push(row);

    outInfoBgs.push(item.colorCols.slice(0, CONFIG.INFO_COLS_END));

    const weekRow = item.weeks.slice();
    while (weekRow.length < outWeekCols) weekRow.push("#ffffff");
    outWeekBgs.push(weekRow);
  }

  // Write values (full width)
  const fullRange = sheet.getRange(CONFIG.DATA_START_ROW, 1, outValues.length, CONFIG.TOTAL_COLS);
  fullRange.setValues(outValues);

  // Info colors A..G exact
  const infoRange = sheet.getRange(CONFIG.DATA_START_ROW, 1, outValues.length, CONFIG.INFO_COLS_END);
  infoRange.setBackgrounds(outInfoBgs);

  // Week colors: read existing ONLY for H..
  const weekRange = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.WEEK_START_COL, outValues.length, outWeekCols);
  const existingWeekBgs = weekRange.getBackgrounds();

  for (let r = 0; r < outValues.length; r++) {
    for (let c = 0; c < outWeekCols; c++) {
      const master = outWeekBgs[r][c];
      const existing = existingWeekBgs[r][c];

      if (master !== "#ffffff" && (existing === "#ffffff" || existing === master)) {
        existingWeekBgs[r][c] = master;
      }
    }
  }
  weekRange.setBackgrounds(existingWeekBgs);

  // Font: all white (fast)
  fullRange.setFontColor("#ffffff");

  // Remember last written rows
  props.setProperty(lastKey, String(outValues.length));
}

/* =========================
 * LOGGING (lightweight, bounded)
 * ========================= */

function logSkippedFast_(skipped, file) {
  if (!skipped || !skipped.length) return;

  let logSheet = file.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = file.insertSheet(CONFIG.LOG_SHEET_NAME);
    logSheet.hideSheet();
  }

  logSheet.clearContents();
  logSheet.getRange(1, 1, 1, 4).setValues([["Timestamp", "Team", "Werknummer", "Reason"]]);
  logSheet.getRange(2, 1, skipped.length, 4).setValues(skipped);
}

/* =========================
 * BATCH TRIGGERS + STATE
 * ========================= */

function scheduleNextBatchTrigger_() {
  cleanupBatchTriggers_();
  ScriptApp.newTrigger("runNextBatch_")
    .timeBased()
    .after(CONFIG.BATCH_TRIGGER_MINUTES * 60 * 1000)
    .create();
}

function cleanupBatchTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === "runNextBatch_") {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function clearBatchState_() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty(CONFIG.PROP_BATCH_INDEX);
  props.deleteProperty(CONFIG.PROP_BATCH_TEAMS);
}

/* =========================
 * LOCKING WRAPPER
 * ========================= */

function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  // 30s wachten is meestal genoeg om dubbele runs te voorkomen
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/* =========================
 * PROPERTY KEY HELPERS
 * ========================= */

function setupKey_(file, sheet) {
  return setupKeyRaw_(file.getId(), sheet.getSheetId());
}
function setupKeyRaw_(fileId, sheetId) {
  return `${CONFIG.PROP_SETUP_PREFIX}${fileId}_${sheetId}`;
}

function lastRowKey_(file, sheet) {
  return lastRowKeyRaw_(file.getId(), sheet.getSheetId());
}
function lastRowKeyRaw_(fileId, sheetId) {
  return `${CONFIG.PROP_LASTROW_PREFIX}${fileId}_${sheetId}`;
}

/* =========================
 * HELPERS
 * ========================= */

function normalize_(v) {
  return String(v || "").trim().toLowerCase();
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
