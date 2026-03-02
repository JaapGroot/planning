/************ CONFIG ************/
const SHEET_NAME = "Planning";
const HEADER_ROWS = 6;                 // rijen 1..6
const FIRST_DATA_ROW = HEADER_ROWS + 1; // rij 7

// Werknummer en header-info kolommen (1-based)
const COL_WN = 1;            // A
const COL_OPDR = 2;          // B
const COL_PLAATS = 3;        // C
const COL_ADRES = 4;         // D
const COL_CONTACT = 5;       // E

// "Echte werkregel inhoud" check: B..M (M=13)
const LINE_FIRST_COL = 2;    // B
const LINE_LAST_COL = 13;   // M

// Header invulcellen
const CELL_OPDR = "J1";
const CELL_CONTACT = "J2";
const CELL_PLAATS = "J3";
const CELL_ADRES = "J4";

// Batch throttling
const SLEEP_BETWEEN_PDFS_MS = 2500;
const EXPORT_MAX_ATTEMPTS = 8;

const BATCH_CHUNK_SIZE = 20;              // 10–15 is meestal veilig
const BATCH_STATE_KEY = "PRINT_BATCH_STATE_V1";

/************ UI ************/
function uiPrintSingle() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Werknummer printen", "Vul werknummer in (bijv. G2600001-1)", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const wn = (resp.getResponseText() || "").trim();
  if (!wn) return ui.alert("Geen werknummer ingevuld.");

  const file = printOne_(wn);
  showLinks_([{ label: wn, url: file.getUrl() }], "Printlink");
}

function uiPrintBatch() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Batch printen", "Vul BASIS werknummer in (bijv. G2600001)", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const base = ((resp.getResponseText() || "").trim().split("-")[0] || "").trim();
  if (!base) return ui.alert("Geen basis werknummer ingevuld.");

  // reset + start nieuwe batch
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(BATCH_STATE_KEY, JSON.stringify({
    base,
    index: 0,
    links: []   // we bewaren links zodat je na meerdere runs 1 lijst krijgt
  }));

  runBatchChunk_();
}

/************ CORE: 1 werknummer => 1 PDF ************/
function printOne_(werknummer) {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SHEET_NAME);
  if (!src) throw new Error(`Tabblad "${SHEET_NAME}" niet gevonden.`);

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  if (lastRow < FIRST_DATA_ROW) throw new Error("Geen data onder de header.");

  // Lees data als DISPLAY (om "" goed als leeg te zien)
  const num = lastRow - HEADER_ROWS;
  const data = src.getRange(FIRST_DATA_ROW, 1, num, lastCol).getDisplayValues();

  // 1) Vind start/eind (eind = laatste rij met echte inhoud in B..M)
  const block = findBlock_(data, werknummer);
  if (!block) throw new Error(`Werknummer niet gevonden: ${werknummer}`);

  const { startIdx, endIdx } = block;

  // Header info uit eerste regel
  const firstRow = data[startIdx];
  const opdrachtgever = firstRow[COL_OPDR - 1] || "";
  const contactpersoon = firstRow[COL_CONTACT - 1] || "";
  const plaats = firstRow[COL_PLAATS - 1] || "";
  const adres = firstRow[COL_ADRES - 1] || "";

  // 2) Eerste regel niet printen => printStartIdx
  const printStartIdx = Math.min(startIdx + 1, endIdx);
  const keptCount = endIdx - printStartIdx + 1;
  if (keptCount <= 0) throw new Error(`Niets om te printen voor ${werknummer} (alleen 1 regel).`);

  // 3) Absolute rijen in sheet
  const startAbsRow = FIRST_DATA_ROW + printStartIdx;
  const endAbsRow = FIRST_DATA_ROW + endIdx;

  // 4) Temp tab in hetzelfde spreadsheet (formules blijven werken)
  const tmpName = `TMP_${werknummer}_${Date.now()}`;
  const tmp = src.copyTo(ss).setName(tmpName);

  try {
    // Header invullen
    tmp.getRange(CELL_OPDR).setValue(opdrachtgever);
    tmp.getRange(CELL_CONTACT).setValue(contactpersoon);
    tmp.getRange(CELL_PLAATS).setValue(plaats);
    tmp.getRange(CELL_ADRES).setValue(adres);


    // Rijen knippen (exact op start/eind)
    const tmpLastRow = tmp.getLastRow();
    if (endAbsRow < tmpLastRow) tmp.deleteRows(endAbsRow + 1, tmpLastRow - endAbsRow);
    if (startAbsRow > FIRST_DATA_ROW) tmp.deleteRows(FIRST_DATA_ROW, startAbsRow - FIRST_DATA_ROW);

    // Lege werkregels binnen het blok weghalen (B..M leeg)
    removeBlankLines_(tmp);

    // Laatste kolom bepalen (laatste zichtbare kolom met inhoud in rijen 1..6)
    const lastColToPrint = lastVisibleContentCol_(tmp, HEADER_ROWS);

    // ✅ Belangrijk: exporthoogte NIET op getLastRow, maar puur op start/eind logica
    // Na removeBlankLines_ kan het aantal data-rijen lager worden,
    // daarom gebruiken we de werkelijke data-hoogte:
    const dataRowsNow = tmp.getLastRow() - HEADER_ROWS;
    const lastRowToPrint = HEADER_ROWS + dataRowsNow;

    return exportPdfRetry_(ss, tmp, `${werknummer}`, lastRowToPrint, lastColToPrint);

  } finally {
    ss.deleteSheet(tmp);
  }
}

/************ Block finding ************/
function clean_(v) {
  return String(v || "")
    .replace(/\u00A0/g, " ")
    .replace(/[\u2000-\u200B]/g, "")
    .trim();
}

function hasLineContent_(row) {
  for (let c = LINE_FIRST_COL; c <= Math.min(LINE_LAST_COL, row.length); c++) {
    if (clean_(row[c - 1]) !== "") return true;
  }
  return false;
}

function findBlock_(data, werknummer) {
  const wnCol = COL_WN - 1;
  let startIdx = -1;
  let endIdx = -1;

  for (let i = 0; i < data.length; i++) {
    if (clean_(data[i][wnCol]) === werknummer) {
      if (startIdx === -1) startIdx = i;
      if (hasLineContent_(data[i])) endIdx = i;
    }
  }
  if (startIdx === -1) return null;
  if (endIdx === -1) {
    for (let i = data.length - 1; i >= startIdx; i--) {
      if (clean_(data[i][wnCol]) === werknummer) { endIdx = i; break; }
    }
  }
  return { startIdx, endIdx };
}

/************ Remove blank work lines (B..M empty) ************/
function removeBlankLines_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < FIRST_DATA_ROW) return;

  const num = lastRow - FIRST_DATA_ROW + 1;
  const display = sheet.getRange(FIRST_DATA_ROW, 1, num, lastCol).getDisplayValues();

  function isBlank_(row) {
    for (let c = LINE_FIRST_COL; c <= Math.min(LINE_LAST_COL, row.length); c++) {
      if (clean_(row[c - 1]) !== "") return false;
    }
    return true;
  }

  for (let i = display.length - 1; i >= 0; i--) {
    if (isBlank_(display[i])) sheet.deleteRow(FIRST_DATA_ROW + i);
  }
}

/************ Last visible content column in header rows ************/
function lastVisibleContentCol_(sheet, maxRow) {
  const lastCol = sheet.getLastColumn();
  const vals = sheet.getRange(1, 1, maxRow, lastCol).getDisplayValues();

  for (let c = lastCol; c >= 1; c--) {
    if (sheet.isColumnHiddenByUser(c)) continue;
    for (let r = 1; r <= maxRow; r++) {
      if (clean_(vals[r - 1][c - 1]) !== "") return c;
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
    format: "pdf",
    size: "A3",
    portrait: "false",
    fitw: "true",
    sheetnames: "false",
    printtitle: "false",
    pagenumbers: "false",
    gridlines: "false",
    fzr: "false",
    r1: "0",
    c1: "0",
    r2: String(lastRowToPrint),
    c2: String(lastColToPrint),
    top_margin: "0.50",
    bottom_margin: "0.50",
    left_margin: "0.50",
    right_margin: "0.50"
  };

  const query = Object.keys(params).map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`).join("&");
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?gid=${gid}&${query}`;

  for (let attempt = 1; attempt <= EXPORT_MAX_ATTEMPTS; attempt++) {
    Utilities.sleep(1200 + attempt * 900);

    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    });

    const code = resp.getResponseCode();
    const ct = String(resp.getHeaders()["Content-Type"] || "").toLowerCase();
    const blob = resp.getBlob();
    const bytes = blob.getBytes();

    const looksPdf = bytes.length >= 4 &&
      bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46; // %PDF

    if (code === 200 && (ct.includes("pdf") || looksPdf)) {
      const folder = getPlanningenFolder_();
      return folder.createFile(blob.setName(`${fileName}.pdf`));
    }

    const text = resp.getContentText() || "";
    const isHtml = ct.includes("text/html") || text.startsWith("<!DOCTYPE html") || text.startsWith("<html");

    if (attempt < EXPORT_MAX_ATTEMPTS && (code === 429 || code === 503 || isHtml)) {
      Utilities.sleep(2500 * attempt);
      continue;
    }

    throw new Error(`PDF export mislukt (HTTP ${code}). ${isHtml ? "Rate limit/HTML terug." : text.substring(0, 250)}`);
  }

  throw new Error("PDF export mislukt na meerdere pogingen.");
}

/************ Variants (base / base-<nummer>) ************/
function findVariants_(sheet, base) {
  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return [];

  const colA = sheet.getRange(FIRST_DATA_ROW, COL_WN, lastRow - FIRST_DATA_ROW + 1, 1).getValues();
  const re = new RegExp("^" + escapeRegex_(base) + "(?:-\\d+)?$");
  const set = new Set();

  for (const [v] of colA) {
    const s = String(v || "").trim();
    if (!s) continue;
    if (re.test(s)) set.add(s);
  }

  const arr = Array.from(set);
  arr.sort((a, b) => {
    const na = a.includes("-") ? parseInt(a.split("-")[1], 10) : 0;
    const nb = b.includes("-") ? parseInt(b.split("-")[1], 10) : 0;
    return na - nb;
  });

  return arr;
}

function escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/************ Links dialog ************/
function showLinks_(items, title) {
  const list = items.map(x => `<li><a href="${x.url}" target="_blank">${escapeHtml_(x.label)}</a></li>`).join("");
  const html = HtmlService.createHtmlOutput(`<p><b>Klaar ✅</b></p><ol>${list}</ol>`)
    .setWidth(420)
    .setHeight(Math.min(600, 140 + items.length * 22));
  SpreadsheetApp.getUi().showModelessDialog(html, title);
}

function escapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function uiResumeBatch() {
  runBatchChunk_();
}

function runBatchChunk_() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  const raw = props.getProperty(BATCH_STATE_KEY);
  if (!raw) return ui.alert("Geen batch om te hervatten. Start eerst via 'Batch: print alle locaties'.");

  const state = JSON.parse(raw);
  const base = state.base;
  let index = state.index || 0;
  const links = state.links || [];

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Tabblad "${SHEET_NAME}" niet gevonden.`);

  const variants = findVariants_(sheet, base);
  if (variants.length === 0) {
    props.deleteProperty(BATCH_STATE_KEY);
    return ui.alert(`Geen locaties gevonden voor ${base}.`);
  }

  // chunk bepalen
  const end = Math.min(index + BATCH_CHUNK_SIZE, variants.length);
  const slice = variants.slice(index, end);

  // uitvoer
  for (const wn of slice) {
    const file = printOne_(wn);
    links.push({ label: wn, url: file.getUrl() });

    // throttle tegen rate limits
    Utilities.sleep(SLEEP_BETWEEN_PDFS_MS);
  }

  // state opslaan
  index = end;
  props.setProperty(BATCH_STATE_KEY, JSON.stringify({ base, index, links }));

  if (index >= variants.length) {
    // klaar
    props.deleteProperty(BATCH_STATE_KEY);
    showLinks_(links, `PDF-links (${base})`);
    ui.alert(`Batch klaar ✅ (${variants.length} locaties).`);
  } else {
    // nog niet klaar: geef voortgang + instructie
    ui.alert(
      `Batch gedeeltelijk klaar: ${index}/${variants.length}.\n` +
      `Klik opnieuw op 'Print → Batch hervatten' om door te gaan.`
    );
  }
}

function getPlanningenFolder_() {
  const ss = SpreadsheetApp.getActive();
  const ssFile = DriveApp.getFileById(ss.getId());

  // Het spreadsheet kan in meerdere mappen staan; we pakken de "eerste" parent.
  const parents = ssFile.getParents();
  if (!parents.hasNext()) {
    // Als er geen parent is (zeldzaam), zet hem in root of maak daar de map.
    const root = DriveApp.getRootFolder();
    const it = root.getFoldersByName("planningen");
    return it.hasNext() ? it.next() : root.createFolder("planningen");
  }

  const parentFolder = parents.next();

  // Zoek submap "planningen" in dezelfde map als het spreadsheet
  const it = parentFolder.getFoldersByName("planningen");
  if (it.hasNext()) return it.next();

  // Bestaat nog niet? Dan maken (mag je ook weglaten als je liever een error wilt)
  return parentFolder.createFolder("planningen");
}