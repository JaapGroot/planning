const CONFIG = {
  PLANNING_SHEET: 'Planning',
  TEAM_TEMPLATE_SHEET_NAME: 'Teamsheet',
  TEAM_FILE_PREFIX: 'Planning - ',

  DATA_START_ROW: 7,
  MASTER_TOTAL_COLS: 65,

  // Masterkolommen
  WORKORDER_COL: 1,     // A
  OPDRACHT_COL: 2,      // B (headerregel: Ja/Nee)
  TEAM_COL: 7,         // L

  // Teamsheet output
  TEAM_OUTPUT_START_ROW: 7,
  TEAM_OUTPUT_START_COL: 1,
  TEAM_OUTPUT_MASTER_COLS: Array.from({ length: 65 }, (_, i) => i + 1),

  // Visuele setup teamsheet
  COPY_BACKGROUNDS: true,
  HEADER_FONT_COLOR: '#ffffff',
  HEADER_FONT_WEIGHT: 'bold',

  // Teamwaarden die geen echt team zijn
  INVALID_TEAM_VALUES: ['team', 'ja', 'nee', 'opdracht'],

  // Template. Leeg laten = leeg spreadsheet maken.
  TEAM_TEMPLATE_SPREADSHEET_ID: '1u_UPDYRf4ccVr6XQeiC0RVuKgPabDIcJrn25tvE4aJc',
  TEAM_FILES_FOLDER_ID: '1Fcp0c1wPSoWSiGawA9JBeuMCer0rr2jv',

  TEAM_DEBUG: 'Danny Waltmann',

  DATA_START_ROW: 7,
  BLOCK_SORT_COL: 6,      // F
  BLOCK_SORT_COL_2: 8,    // H
  DETAIL_SORT_COL: 3,     // C
  TEMP_SHEET_NAME: "_tmp_sort_blocks_",

  DETAIL_ORDER: [
    'voorbereiding',
    'bomen',
    'heesters',
    'solitaire heesters',
    'heesters en vaste planten',
    'bosplantsoen',
    'rozen',
    'vaste planten',
    'vaste planten en siergras',
    'prairie garden',
    'siergrassen',
    'bolbloemen',
    'sedum',
    'plantenbakken',
    'plantvakken',
    'bodembedekkers',
    'klimop',
    'klimplanten',
    'hagen',
    'gazons',
    'ruig gras',
    'talud',
    'grind',
    'verharding',
    'riool',
    'watergangen',
    'algemeen',
  ]
};

/**
 * =========================
 * UI
 * =========================
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Team Planning')
    .addItem('Create/update alle teams', 'syncAllTeams')
    .addItem('Create/update 1 team', 'promptAndSyncSingleTeam')
    .addSeparator()
    .addItem('Debug timing alle teams', 'debugAllTeamsTiming')
    .addItem('Debug timing 1 team', 'debugSingleTeamTiming')
    .addItem('Debug teams', 'debugListTeams')
    .addToUi();

    SpreadsheetApp.getUi()
    .createMenu("Print")
    .addItem("Print 1 werknummer (A3)", "uiPrintSingle")
    .addItem("Batch: print alle locaties (PDF per locatie)", "uiPrintBatch")
    .addItem("Batch hervatten", "uiResumeBatch")
    .addItem("Sorteer planning (plaats)", "sortPlanningByLocation")
    .addItem("Sorteer op werknummer", "sortPlanningByWorkNumber")
    .addToUi();
}

function promptAndSyncSingleTeam() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Create/update 1 team', 'Vul exact de teamnaam in:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;

  const teamName = normalizeTeamNameKeepCase_(result.getResponseText());
  if (!teamName) {
    ui.alert('Geen geldige teamnaam ingevuld.');
    return;
  }

  syncSingleTeam(teamName);
}

/**
 * =========================
 * Public sync functies
 * =========================
 */
function syncAllTeams() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planningSheet = getPlanningSheetOrThrow_(ss);

  const snapshot = timeStep_('01 readPlanningSnapshot_', () => readPlanningSnapshot_(planningSheet));
  const blocks = timeStep_('02 buildWorkOrderBlocks_', () => buildWorkOrderBlocks_(snapshot));
  const teamIndex = timeStep_('03 indexBlocksByTeam_', () => indexBlocksByTeam_(blocks));
  const cache = timeStep_('04 buildTeamFileCache_', () => buildTeamFileCache_(ss));

  const teamNames = Object.keys(teamIndex).sort(localeCompareNl_);
  Logger.log('Teams gevonden: %s', JSON.stringify(teamNames));

  teamNames.forEach((teamName, idx) => {
    timeStep_(`05 sync team ${idx + 1}/${teamNames.length}: ${teamName}`, () => {
      syncTeamFromIndex_(ss, teamName, teamIndex[teamName], cache);
    });
  });
}

function syncSingleTeam(teamName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planningSheet = getPlanningSheetOrThrow_(ss);
  const cleanTeamName = normalizeTeamNameKeepCase_(teamName);
  if (!cleanTeamName) throw new Error('Ongeldige teamnaam.');

  const snapshot = timeStep_('01 readPlanningSnapshot_', () => readPlanningSnapshot_(planningSheet));
  const blocks = timeStep_('02 buildWorkOrderBlocks_', () => buildWorkOrderBlocks_(snapshot));
  const teamIndex = timeStep_('03 indexBlocksByTeam_', () => indexBlocksByTeam_(blocks));
  const cache = timeStep_('04 buildTeamFileCache_', () => buildTeamFileCache_(ss));

  const matchedTeamName = findCanonicalTeamName_(cleanTeamName, Object.keys(teamIndex));
  if (!matchedTeamName) {
    throw new Error(`Team niet gevonden in planning: ${cleanTeamName}`);
  }

  timeStep_(`05 sync team ${matchedTeamName}`, () => {
    syncTeamFromIndex_(ss, matchedTeamName, teamIndex[matchedTeamName], cache);
  });
}

/**
 * =========================
 * Debug
 * =========================
 */
function debugAllTeamsTiming() {
  syncAllTeams();
}

function debugSingleTeamTiming() {
  const TEAM_NAME = TEAM_DEBUG;
  syncSingleTeam(TEAM_NAME);
}

function debugListTeams() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planningSheet = getPlanningSheetOrThrow_(ss);
  const snapshot = readPlanningSnapshot_(planningSheet);
  const blocks = buildWorkOrderBlocks_(snapshot);
  const index = indexBlocksByTeam_(blocks);
  Logger.log('Teams: %s', JSON.stringify(Object.keys(index).sort(localeCompareNl_)));
}

/**
 * =========================
 * Read planning
 * =========================
 */
function getPlanningSheetOrThrow_(ss) {
  const sheet = ss.getSheetByName(CONFIG.PLANNING_SHEET);
  if (!sheet) throw new Error(`Tabblad niet gevonden: ${CONFIG.PLANNING_SHEET}`);
  return sheet;
}

function readPlanningSnapshot_(planningSheet) {
  const lastRow = planningSheet.getLastRow();
  const lastCol = Math.min(CONFIG.MASTER_TOTAL_COLS, planningSheet.getLastColumn());

  if (lastRow < CONFIG.DATA_START_ROW) {
    return {
      lastRow,
      lastCol,
      numRows: 0,
      values: [],
      backgrounds: [],
    };
  }

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const range = planningSheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, lastCol);

  return {
    lastRow,
    lastCol,
    numRows,
    values: range.getValues(),
    backgrounds: CONFIG.COPY_BACKGROUNDS ? range.getBackgrounds() : [],
  };
}

/**
 * =========================
 * Blocks (werknummerblokken)
 * =========================
 */
function buildWorkOrderBlocks_(snapshot) {
  const values = snapshot.values || [];
  const backgrounds = snapshot.backgrounds || [];
  const blocks = [];

  let currentBlock = null;
  let currentWorkOrder = null;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const bg = backgrounds[i] || null;
    const workOrder = normalizeWorkOrder_(row[CONFIG.WORKORDER_COL - 1]);
    if (!workOrder) continue;

    const isNewHeader = workOrder !== currentWorkOrder;

    if (isNewHeader) {
      if (currentBlock) {
        finalizeBlock_(currentBlock);
        if (currentBlock.isOpdracht) {
          blocks.push(currentBlock);
        }
      }

      currentWorkOrder = workOrder;
      currentBlock = {
        workOrder,
        headerRow: row,
        headerBg: bg,
        detailRows: [],
        detailBgs: [],
        teams: [],
        isOpdracht: false,
      };
      continue;
    }

    if (!currentBlock) continue;
    currentBlock.detailRows.push(row);
    currentBlock.detailBgs.push(bg);
  }

  if (currentBlock) {
    finalizeBlock_(currentBlock);
    if (currentBlock.isOpdracht) {
      blocks.push(currentBlock);
    }
  }

  Logger.log('Aantal werknummerblokken (alleen opdracht=ja): %s', blocks.length);
  return blocks;
}

function finalizeBlock_(block) {
  block.isOpdracht = isHeaderOpdracht_(block.headerRow);
  block.teams = block.isOpdracht ? extractTeamsFromBlock_(block) : [];
}

function isHeaderOpdracht_(headerRow) {
  const raw = String(headerRow[CONFIG.OPDRACHT_COL - 1] == null ? '' : headerRow[CONFIG.OPDRACHT_COL - 1])
    .trim()
    .toLowerCase();
  return raw === 'ja';
}

function extractTeamsFromBlock_(block) {
  const idx = CONFIG.TEAM_COL - 1;
  const teamSet = new Set();

  (block.detailRows || []).forEach(row => {
    splitTeamCell_(row[idx]).forEach(team => teamSet.add(team));
  });

  return Array.from(teamSet).sort(localeCompareNl_);
}

function indexBlocksByTeam_(blocks) {
  const index = {};

  (blocks || []).forEach(block => {
    (block.teams || []).forEach(teamName => {
      if (!index[teamName]) index[teamName] = [];
      index[teamName].push(block);
    });
  });

  return index;
}

/**
 * =========================
 * Team output bouwen
 * =========================
 */
function buildTeamOutput_(teamName, teamBlocks) {
  const rows = [];
  const backgrounds = [];
  const headerRows = [];
  const teamNorm = normalizeTeamName_(teamName);

  (teamBlocks || []).forEach(block => {
    rows.push(mapMasterRowToTeamRow_(block.headerRow));
    headerRows.push(rows.length); // 1-based binnen output
    if (CONFIG.COPY_BACKGROUNDS) {
      backgrounds.push(mapMasterBackgroundRowToTeamBackgroundRow_(block.headerBg));
    }

    for (let i = 0; i < block.detailRows.length; i++) {
      const row = block.detailRows[i];
      const rowBg = block.detailBgs[i] || null;
      if (!rowBelongsToTeam_(row, teamNorm)) continue;

      rows.push(mapMasterRowToTeamRow_(row));
      if (CONFIG.COPY_BACKGROUNDS) {
        backgrounds.push(mapMasterBackgroundRowToTeamBackgroundRow_(rowBg));
      }
    }
  });

  return { rows, backgrounds, headerRows };
}

function rowBelongsToTeam_(row, normalizedTeamName) {
  const teams = splitTeamCell_(row[CONFIG.TEAM_COL - 1]).map(normalizeTeamName_);
  return teams.indexOf(normalizedTeamName) >= 0;
}

function mapMasterRowToTeamRow_(masterRow) {
  return CONFIG.TEAM_OUTPUT_MASTER_COLS.map(col1 => masterRow[col1 - 1]);
}

function mapMasterBackgroundRowToTeamBackgroundRow_(bgRow) {
  if (!CONFIG.COPY_BACKGROUNDS) return [];
  if (!bgRow) return CONFIG.TEAM_OUTPUT_MASTER_COLS.map(() => '#ffffff');
  return CONFIG.TEAM_OUTPUT_MASTER_COLS.map(col1 => bgRow[col1 - 1] || '#ffffff');
}

/**
 * =========================
 * Team files / team sheets
 * =========================
 */
function buildTeamFileCache_(masterSs) {
  const parentFolder = DriveApp.getFolderById(CONFIG.TEAM_FILES_FOLDER_ID);
  const files = parentFolder.getFiles();
  const byName = {};

  while (files.hasNext()) {
    const file = files.next();
    byName[file.getName()] = file.getId();
  }

  return { parentFolder, byName };
}

function getFirstParentFolderOrThrow_(fileId) {
  const file = DriveApp.getFileById(fileId);
  const parents = file.getParents();
  if (!parents.hasNext()) {
    throw new Error('Master spreadsheet heeft geen parent folder in Drive.');
  }
  return parents.next();
}

function getOrCreateTeamSpreadsheet_(masterSs, teamName, cache) {
  const fileName = CONFIG.TEAM_FILE_PREFIX + teamName;
  const existingId = cache.byName[fileName];

  if (existingId) {
    return {
      spreadsheetId: existingId,
      spreadsheet: SpreadsheetApp.openById(existingId),
      isNew: false,
      fileName,
    };
  }

  let newFile;
  if (CONFIG.TEAM_TEMPLATE_SPREADSHEET_ID) {
    const templateFile = DriveApp.getFileById(CONFIG.TEAM_TEMPLATE_SPREADSHEET_ID);
    newFile = templateFile.makeCopy(fileName, cache.parentFolder);
  } else {
    const tempSs = SpreadsheetApp.create(fileName);
    newFile = DriveApp.getFileById(tempSs.getId());
    cache.parentFolder.addFile(newFile);
    try {
      DriveApp.getRootFolder().removeFile(newFile);
    } catch (e) {
      // Niet kritisch.
    }
  }

  const spreadsheet = SpreadsheetApp.openById(newFile.getId());
  cache.byName[fileName] = newFile.getId();

  return {
    spreadsheetId: newFile.getId(),
    spreadsheet,
    isNew: true,
    fileName,
  };
}

function getOrCreateTeamSheet_(teamSpreadsheet, teamName) {
  let sheet = teamSpreadsheet.getSheetByName(teamName);
  if (sheet) return sheet;

  sheet = teamSpreadsheet.getSheetByName(CONFIG.TEAM_TEMPLATE_SHEET_NAME);
  if (sheet) {
    sheet.setName(teamName);
    return sheet;
  }

  sheet = teamSpreadsheet.getSheets()[0] || teamSpreadsheet.insertSheet(teamName);
  if (sheet.getName() !== teamName) sheet.setName(teamName);
  return sheet;
}

function ensureTeamSheetSetup_(sheet, teamName) {
  sheet.getRange('A1').setValue(teamName);
}

/**
 * =========================
 * Create vs update
 * =========================
 */
function syncTeamFromIndex_(masterSs, teamName, teamBlocks, cache) {
  const teamFile = timeStep_(`getOrCreateTeamSpreadsheet_ ${teamName}`, () => getOrCreateTeamSpreadsheet_(masterSs, teamName, cache));
  const sheet = timeStep_(`getOrCreateTeamSheet_ ${teamName}`, () => getOrCreateTeamSheet_(teamFile.spreadsheet, teamName));

  if (teamFile.isNew) {
    timeStep_(`createTeamSheet_ ${teamName}`, () => createTeamSheet_(sheet, teamName, teamBlocks));
  } else {
    timeStep_(`updateTeamSheetAppendMissing_ ${teamName}`, () => updateTeamSheetAppendMissing_(sheet, teamName, teamBlocks));
  }
}

function createTeamSheet_(sheet, teamName, teamBlocks) {
  ensureTeamSheetSetup_(sheet, teamName);

  const output = timeStep_(`buildTeamOutput_ create ${teamName}`, () => buildTeamOutput_(teamName, teamBlocks));
  timeStep_(`clearTeamOutputArea_ create ${teamName}`, () => clearTeamOutputArea_(sheet));
  timeStep_(`writeTeamOutputFull_ create ${teamName}`, () => writeTeamOutputFull_(sheet, output));
}

function updateTeamSheetAppendMissing_(sheet, teamName, teamBlocks) {
  const existingWorkOrders = timeStep_(`readExistingWorkOrders_ ${teamName}`, () => readExistingWorkOrders_(sheet));

  const missingBlocks = (teamBlocks || []).filter(block => !existingWorkOrders.has(block.workOrder));
  Logger.log('%s -> ontbrekende blokken: %s', teamName, missingBlocks.length);

  if (!missingBlocks.length) return;

  const output = timeStep_(`buildTeamOutput_ update ${teamName}`, () => buildTeamOutput_(teamName, missingBlocks));
  timeStep_(`appendTeamOutput_ update ${teamName}`, () => appendTeamOutput_(sheet, output));
}

function readExistingWorkOrders_(sheet) {
  const lastRow = sheet.getLastRow();
  const set = new Set();
  if (lastRow < CONFIG.TEAM_OUTPUT_START_ROW) return set;

  const numRows = lastRow - CONFIG.TEAM_OUTPUT_START_ROW + 1;
  const values = sheet
    .getRange(CONFIG.TEAM_OUTPUT_START_ROW, CONFIG.TEAM_OUTPUT_START_COL, numRows, 1)
    .getValues();

  values.forEach(row => {
    const workOrder = normalizeWorkOrder_(row[0]);
    if (workOrder) set.add(workOrder);
  });

  return set;
}

/**
 * =========================
 * Write
 * =========================
 */
function clearTeamOutputArea_(sheet) {
  const maxRows = sheet.getMaxRows();
  const maxCols = Math.max(sheet.getMaxColumns(), CONFIG.TEAM_OUTPUT_MASTER_COLS.length);
  const numRows = Math.max(0, maxRows - CONFIG.TEAM_OUTPUT_START_ROW + 1);
  if (numRows <= 0) return;

  sheet
    .getRange(CONFIG.TEAM_OUTPUT_START_ROW, CONFIG.TEAM_OUTPUT_START_COL, numRows, maxCols)
    .clearContent()
    .clearFormat();

  clearRowGroups_(sheet, numRows);
}

function writeTeamOutputFull_(sheet, output) {
  const rows = output.rows || [];
  const backgrounds = output.backgrounds || [];
  if (!rows.length) return;

  const width = rows[0].length;
  ensureSheetHasEnoughSize_(sheet, CONFIG.TEAM_OUTPUT_START_ROW + rows.length - 1, CONFIG.TEAM_OUTPUT_START_COL + width - 1);

  const range = sheet.getRange(CONFIG.TEAM_OUTPUT_START_ROW, CONFIG.TEAM_OUTPUT_START_COL, rows.length, width);
  range.setValues(rows);

  if (CONFIG.COPY_BACKGROUNDS && backgrounds.length === rows.length) {
    range.setBackgrounds(backgrounds);
  }

  styleHeaderRows_(sheet, output.headerRows, CONFIG.TEAM_OUTPUT_START_ROW, width);
  createRowGroups_(sheet, output.headerRows, CONFIG.TEAM_OUTPUT_START_ROW, rows.length);
}

function appendTeamOutput_(sheet, output) {
  const rows = output.rows || [];
  const backgrounds = output.backgrounds || [];
  if (!rows.length) return;

  const width = rows[0].length;
  const startRow = Math.max(sheet.getLastRow() + 1, CONFIG.TEAM_OUTPUT_START_ROW);

  ensureSheetHasEnoughSize_(sheet, startRow + rows.length - 1, CONFIG.TEAM_OUTPUT_START_COL + width - 1);

  const range = sheet.getRange(startRow, CONFIG.TEAM_OUTPUT_START_COL, rows.length, width);
  range.setValues(rows);

  if (CONFIG.COPY_BACKGROUNDS && backgrounds.length === rows.length) {
    range.setBackgrounds(backgrounds);
  }

  styleHeaderRows_(sheet, output.headerRows, startRow, width);
  createRowGroups_(sheet, output.headerRows, startRow, rows.length);
}

function styleHeaderRows_(sheet, headerRows, startRow, width) {
  if (!headerRows || !headerRows.length) return;

  headerRows.forEach(relativeRow => {
    sheet
      .getRange(startRow + relativeRow - 1, CONFIG.TEAM_OUTPUT_START_COL, 1, width)
      .setFontWeight(CONFIG.HEADER_FONT_WEIGHT)
      .setFontColor(CONFIG.HEADER_FONT_COLOR);
  });
}

function createRowGroups_(sheet, headerRows, startRow, totalRows) {
  if (!headerRows || headerRows.length === 0) return;

  for (let i = 0; i < headerRows.length; i++) {
    const headerRow = headerRows[i];
    const detailStart = headerRow + 1;
    const detailEnd = (i < headerRows.length - 1) ? headerRows[i + 1] - 1 : totalRows;
    const numRows = detailEnd - detailStart + 1;
    if (numRows <= 0) continue;

    const sheetRowStart = startRow + detailStart - 1;
    sheet.getRange(sheetRowStart, 1, numRows, 1).shiftRowGroupDepth(1);
  }
}

function clearRowGroups_(sheet, totalRows) {
  if (totalRows <= 0) return;
  const startRow = CONFIG.TEAM_OUTPUT_START_ROW;
  for (let d = 0; d < 8; d++) {
    sheet.getRange(startRow, 1, totalRows, 1).shiftRowGroupDepth(-1);
  }
}

function ensureSheetHasEnoughSize_(sheet, requiredRows, requiredCols) {
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  if (requiredRows > maxRows) {
    sheet.insertRowsAfter(maxRows, requiredRows - maxRows);
  }
  if (requiredCols > maxCols) {
    sheet.insertColumnsAfter(maxCols, requiredCols - maxCols);
  }
}

/**
 * =========================
 * Helpers
 * =========================
 */
function timeStep_(label, fn) {
  const start = Date.now();
  const result = fn();
  Logger.log('%s: %sms', label, Date.now() - start);
  return result;
}

function normalizeWorkOrder_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function splitTeamCell_(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return [];

  return raw
    .split(/[;,/\n]+/)
    .map(s => normalizeTeamNameKeepCase_(s))
    .filter(Boolean)
    .filter(name => CONFIG.INVALID_TEAM_VALUES.indexOf(name.toLowerCase()) === -1);
}

function normalizeTeamName_(value) {
  return normalizeTeamNameKeepCase_(value).toLowerCase();
}

function normalizeTeamNameKeepCase_(value) {
  return String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
}

function findCanonicalTeamName_(searchTeam, availableTeamNames) {
  const wanted = normalizeTeamName_(searchTeam);
  for (let i = 0; i < availableTeamNames.length; i++) {
    if (normalizeTeamName_(availableTeamNames[i]) === wanted) {
      return availableTeamNames[i];
    }
  }
  return '';
}

function localeCompareNl_(a, b) {
  return String(a).localeCompare(String(b), 'nl');
}