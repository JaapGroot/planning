function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function setupKey_(file, sheet) {
  return `${CONFIG.PROP_SETUP_PREFIX}${file.getId()}_${sheet.getSheetId()}`;
}

function lastRowKey_(file, sheet) {
  return `${CONFIG.PROP_LASTROW_PREFIX}${file.getId()}_${sheet.getSheetId()}`;
}

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
