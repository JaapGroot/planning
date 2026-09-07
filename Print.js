/************ PRINT CONFIG ************/
const PRINT_CONFIG = {
  CELL_OPDRACHTGEVER: 'E1',
  CELL_CONTACTPERSOON: 'E2',
  CELL_PLAATS: 'E3',
  CELL_ADRES: 'E4',
  SLEEP_BETWEEN_PDFS_MS: 2500,
  EXPORT_MAX_ATTEMPTS: 8,
  BATCH_CHUNK_SIZE: 20,
  BATCH_STATE_KEY: 'PRINT_BATCH_STATE_V1',
};

/************ UI ************/
function uiPrintSingle() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Werknummer printen', 'Vul werknummer in (bijv. G2700001-1)', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const wn = clean_(resp.getResponseText());
  if (!wn) return ui.alert('Geen werknummer ingevuld.');

  const file = printOne_(wn);
  showLinks_([{ label: wn, url: file.getUrl() }], 'Printlink');
}

function uiPrintBatch() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Batch printen', 'Vul BASIS werknummer in (bijv. G2700001)', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const base = clean_((resp.getResponseText() || '').split('-')[0]);
  if (!base) return ui.alert('Geen basis werknummer ingevuld.');

  const props = PropertiesService.getDocumentProperties();
  props.setProperty(PRINT_CONFIG.BATCH_STATE_KEY, JSON.stringify({
    base,
    index: 0,
    links: []
  }));

  runBatchChunk_();
}

function uiResumeBatch() {
  runBatchChunk_();
}

/************ CORE: 1 werknummer => 1 PDF ************/
function printOne_(werknummer) {
  return withPlanningDocumentLock_('PDF-print', () => printOneUnlocked_(werknummer));
}

function printOneUnlocked_(werknummer) {
  const ss = SpreadsheetApp.getActive();
  const src = getPlanningSheetOrThrow_(ss);
  assertPlanningLayout2027_(src);

  const headerRows = CONFIG.DATA_START_ROW - 1;
  const firstDataRow = CONFIG.DATA_START_ROW;
  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();

  if (lastRow < firstDataRow) throw new Error('Geen data onder de header.');

  const cols = getPrintPlanningColumns_();
  const numRows = lastRow - firstDataRow + 1;
  const data = src.getRange(firstDataRow, 1, numRows, lastCol).getDisplayValues();

  const block = findPrintBlock_(data, werknummer, cols);
  if (!block) throw new Error(`Werknummer niet gevonden: ${werknummer}`);

  const { startIdx, endIdx } = block;
  const firstRow = data[startIdx];

  const opdrachtgever = firstRow[cols.opdrachtgever - 1] || '';
  const contactpersoon = firstRow[cols.contactpersoon - 1] || '';
  const plaats = firstRow[cols.plaats - 1] || '';
  const adres = firstRow[cols.adres - 1] || '';

  // De kopregel zelf wordt niet als werkregel geprint.
  const printStartIdx = startIdx + 1;
  if (endIdx < printStartIdx) {
    throw new Error(`Niets om te printen voor ${werknummer} (geen werkregels).`);
  }

  const startAbsRow = firstDataRow + printStartIdx;
  const endAbsRow = firstDataRow + endIdx;
  const tmpName = buildTempPrintSheetName_(werknummer);
  const tmp = src.copyTo(ss).setName(tmpName);

  try {
    tmp.getRange(PRINT_CONFIG.CELL_OPDRACHTGEVER).setValue(opdrachtgever);
    tmp.getRange(PRINT_CONFIG.CELL_CONTACTPERSOON).setValue(contactpersoon);
    tmp.getRange(PRINT_CONFIG.CELL_PLAATS).setValue(plaats);
    tmp.getRange(PRINT_CONFIG.CELL_ADRES).setValue(adres);

    const tmpLastRow = tmp.getLastRow();
    if (endAbsRow < tmpLastRow) {
      tmp.deleteRows(endAbsRow + 1, tmpLastRow - endAbsRow);
    }
    if (startAbsRow > firstDataRow) {
      tmp.deleteRows(firstDataRow, startAbsRow - firstDataRow);
    }

    removeBlankPrintLines_(tmp, cols);

    const lastColToPrint = lastVisibleContentCol_(tmp, headerRows);
    const dataRowsNow = tmp.getLastRow() - headerRows;
    if (dataRowsNow <= 0) throw new Error(`Niets om te printen voor ${werknummer}.`);

    return exportPdfRetry_(
      ss,
      tmp,
      werknummer,
      headerRows + dataRowsNow,
      lastColToPrint
    );
  } finally {
    ss.deleteSheet(tmp);
  }
}

function getPrintPlanningColumns_() {
  return {
    werknummer: CONFIG.WORKORDER_COL,
    opdrachtgever: planningColumnByHeader_('werksoort'), // B op kopregel
    contactpersoon: planningColumnByHeader_('frequentie'), // C op kopregel
    plaats: planningColumnByHeader_('eenheid'), // E op kopregel
    adres: planningColumnByHeader_('werkzaamheden'), // G op kopregel
    detailFirst: planningColumnByHeader_('werksoort'), // B
    detailLast: planningColumnByHeader_('werkzaamheden'), // G
  };
}

function buildTempPrintSheetName_(werknummer) {
  const safeWorkNumber = clean_(werknummer).replace(/[\\/?*\[\]:]/g, '_').slice(0, 40);
  return `TMP_${safeWorkNumber}_${Date.now()}`.slice(0, 99);
}

/************ Block finding ************/
function clean_(v) {
  return String(v || '')
    .replace(/\u00A0/g, ' ')
    .replace(/[\u2000-\u200B]/g, '')
    .trim();
}

function hasPrintLineContent_(row, cols) {
  for (let c = cols.detailFirst; c <= Math.min(cols.detailLast, row.length); c++) {
    if (clean_(row[c - 1]) !== '') return true;
  }
  return false;
}

function findPrintBlock_(data, werknummer, cols) {
  const wanted = clean_(werknummer);
  const wnCol = cols.werknummer - 1;
  let startIdx = -1;
  let endIdx = -1;

  for (let i = 0; i < data.length; i++) {
    if (clean_(data[i][wnCol]) !== wanted) continue;

    if (startIdx === -1) {
      startIdx = i;
      continue; // eerste match is de headerregel
    }

    if (hasPrintLineContent_(data[i], cols)) endIdx = i;
  }

  if (startIdx === -1) return null;
  if (endIdx === -1) endIdx = startIdx;

  return { startIdx, endIdx };
}

/************ Remove blank work lines ************/
function removeBlankPrintLines_(sheet, cols) {
  const firstDataRow = CONFIG.DATA_START_ROW;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < firstDataRow) return;

  const numRows = lastRow - firstDataRow + 1;
  const display = sheet.getRange(firstDataRow, 1, numRows, lastCol).getDisplayValues();
  const rowsToDelete = [];

  for (let i = 0; i < display.length; i++) {
    if (!hasPrintLineContent_(display[i], cols)) {
      rowsToDelete.push(firstDataRow + i);
    }
  }

  // Verwijder aaneengesloten lege ranges van onder naar boven. Dat is sneller
  // dan elke lege rij afzonderlijk verwijderen.
  deleteRowRunsBottomUp_(sheet, rowsToDelete);
}

function deleteRowRunsBottomUp_(sheet, rowNumbers) {
  if (!rowNumbers.length) return;

  let runEnd = rowNumbers[rowNumbers.length - 1];
  let runStart = runEnd;

  for (let i = rowNumbers.length - 2; i >= -1; i--) {
    const row = i >= 0 ? rowNumbers[i] : null;
    if (row !== null && row === runStart - 1) {
      runStart = row;
      continue;
    }

    sheet.deleteRows(runStart, runEnd - runStart + 1);
    if (row !== null) {
      runStart = row;
      runEnd = row;
    }
  }
}

/************ Last visible content column in header rows ************/
function lastVisibleContentCol_(sheet, maxRow) {
  const lastCol = sheet.getLastColumn();
  const vals = sheet.getRange(1, 1, maxRow, lastCol).getDisplayValues();

  for (let c = lastCol; c >= 1; c--) {
    if (sheet.isColumnHiddenByUser(c)) continue;
    for (let r = 1; r <= maxRow; r++) {
      if (clean_(vals[r - 1][c - 1]) !== '') return c;
    }
  }
  return 1;
}

/************ Export PDF with retry/backoff ************/
function exportPdfRetry_(ss, sheet, fileName, lastRowToPrint, lastColToPrint) {
  const ssId = ss.getId();
  const gid = sheet.getSheetId();
  const token = ScriptApp.getOAuthToken();

  const params = {
    format: 'pdf',
    size: 'A3',
    portrait: 'false',
    fitw: 'true',
    sheetnames: 'false',
    printtitle: 'false',
    pagenumbers: 'false',
    gridlines: 'false',
    fzr: 'false',
    r1: '0',
    c1: '0',
    r2: String(lastRowToPrint),
    c2: String(lastColToPrint),
    top_margin: '0.50',
    bottom_margin: '0.50',
    left_margin: '0.50',
    right_margin: '0.50'
  };

  const query = Object.keys(params)
    .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
    .join('&');

  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?gid=${gid}&${query}`;

  for (let attempt = 1; attempt <= PRINT_CONFIG.EXPORT_MAX_ATTEMPTS; attempt++) {
    Utilities.sleep(1200 + attempt * 900);

    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    const code = resp.getResponseCode();
    const ct = String(resp.getHeaders()['Content-Type'] || '').toLowerCase();
    const blob = resp.getBlob();
    const bytes = blob.getBytes();

    const looksPdf = bytes.length >= 4 &&
      bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46;

    if (code === 200 && (ct.includes('pdf') || looksPdf)) {
      const folder = getPlanningenFolder_();
      return folder.createFile(blob.setName(`${fileName}.pdf`));
    }

    const text = resp.getContentText() || '';
    const isHtml = ct.includes('text/html') || text.startsWith('<!DOCTYPE html') || text.startsWith('<html');

    if (attempt < PRINT_CONFIG.EXPORT_MAX_ATTEMPTS && (code === 429 || code === 503 || isHtml)) {
      Utilities.sleep(2500 * attempt);
      continue;
    }

    throw new Error(`PDF export mislukt (HTTP ${code}). ${isHtml ? 'Rate limit/HTML terug.' : text.substring(0, 250)}`);
  }

  throw new Error('PDF export mislukt na meerdere pogingen.');
}

/************ Variants (base / base-<nummer>) ************/
function findVariants_(sheet, base) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return [];

  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const colA = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.WORKORDER_COL, numRows, 1).getValues();
  const re = new RegExp('^' + escapeRegex_(base) + '(?:-\\d+)?$');
  const set = new Set();

  for (const [v] of colA) {
    const s = clean_(v);
    if (s && re.test(s)) set.add(s);
  }

  const arr = Array.from(set);
  arr.sort((a, b) => variantNumber_(a) - variantNumber_(b));
  return arr;
}

function variantNumber_(workNumber) {
  const match = String(workNumber).match(/-(\d+)$/);
  return match ? Number(match[1]) : 0;
}

function escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/************ Links dialog ************/
function showLinks_(items, title) {
  const list = items
    .map(x => `<li><a href="${escapeHtml_(x.url)}" target="_blank">${escapeHtml_(x.label)}</a></li>`)
    .join('');

  const html = HtmlService.createHtmlOutput(`<p><b>Klaar ✅</b></p><ol>${list}</ol>`)
    .setWidth(420)
    .setHeight(Math.min(600, 140 + items.length * 22));

  SpreadsheetApp.getUi().showModelessDialog(html, title);
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function runBatchChunk_() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  const raw = props.getProperty(PRINT_CONFIG.BATCH_STATE_KEY);
  if (!raw) return ui.alert("Geen batch om te hervatten. Start eerst via 'Batch: print alle locaties'.");

  const state = JSON.parse(raw);
  const base = state.base;
  let index = state.index || 0;
  const links = state.links || [];

  const ss = SpreadsheetApp.getActive();
  const sheet = getPlanningSheetOrThrow_(ss);
  assertPlanningLayout2027_(sheet);

  const variants = findVariants_(sheet, base);
  if (variants.length === 0) {
    props.deleteProperty(PRINT_CONFIG.BATCH_STATE_KEY);
    return ui.alert(`Geen locaties gevonden voor ${base}.`);
  }

  const end = Math.min(index + PRINT_CONFIG.BATCH_CHUNK_SIZE, variants.length);
  const slice = variants.slice(index, end);

  for (const wn of slice) {
    const file = printOne_(wn);
    links.push({ label: wn, url: file.getUrl() });
    Utilities.sleep(PRINT_CONFIG.SLEEP_BETWEEN_PDFS_MS);
  }

  index = end;
  props.setProperty(PRINT_CONFIG.BATCH_STATE_KEY, JSON.stringify({ base, index, links }));

  if (index >= variants.length) {
    props.deleteProperty(PRINT_CONFIG.BATCH_STATE_KEY);
    showLinks_(links, `PDF-links (${base})`);
    ui.alert(`Batch klaar ✅ (${variants.length} locaties).`);
  } else {
    ui.alert(
      `Batch gedeeltelijk klaar: ${index}/${variants.length}.\n` +
      `Klik opnieuw op 'Print → Batch hervatten' om door te gaan.`
    );
  }
}

function getPlanningenFolder_() {
  const ss = SpreadsheetApp.getActive();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parents = ssFile.getParents();

  if (!parents.hasNext()) {
    const root = DriveApp.getRootFolder();
    const it = root.getFoldersByName('planningen');
    return it.hasNext() ? it.next() : root.createFolder('planningen');
  }

  const parentFolder = parents.next();
  const it = parentFolder.getFoldersByName('planningen');
  if (it.hasNext()) return it.next();

  return parentFolder.createFolder('planningen');
}
