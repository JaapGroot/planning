function insertDriveLinksFromFolder() {
  return withPlanningDocumentLock_('Drive-links invoegen', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    if (sheet.getName() !== CONFIG.PLANNING_SHEET) {
      throw new Error(`Drive-links kunnen alleen op tabblad "${CONFIG.PLANNING_SHEET}" worden ingevoegd.`);
    }

    const folderId = getDriveLinksFolderId_();
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const rows = [];

    while (files.hasNext()) {
      const file = files.next();
      rows.push([file.getUrl(), file.getName()]);
    }

    if (!rows.length) {
      SpreadsheetApp.getUi().alert(`Geen bestanden gevonden in map "${folder.getName()}".`);
      return;
    }

    // DriveApp geeft geen gegarandeerde volgorde. Sorteer op bestandsnaam zodat
    // herhaald invoegen altijd hetzelfde resultaat geeft.
    rows.sort((a, b) => String(a[1]).localeCompare(String(b[1]), 'nl', {
      numeric: true,
      sensitivity: 'base'
    }));

    const activeCell = sheet.getActiveCell();
    const startRow = activeCell.getRow();
    const startCol = activeCell.getColumn();

    ensureSheetHasEnoughSize_(
      sheet,
      startRow + rows.length - 1,
      startCol + 1
    );

    sheet.getRange(startRow, startCol, rows.length, 2).setValues(rows);
  });
}

function getDriveLinksFolderId_() {
  const override = PropertiesService
    .getDocumentProperties()
    .getProperty('DRIVE_LINKS_FOLDER_ID');

  const folderId = String(override || CONFIG.DRIVE_LINKS_FOLDER_ID || '').trim();
  if (!folderId) {
    throw new Error(
      'Geen Drive-map ingesteld voor DriveLinks. Stel CONFIG.DRIVE_LINKS_FOLDER_ID of DocumentProperty DRIVE_LINKS_FOLDER_ID in.'
    );
  }

  return folderId;
}
