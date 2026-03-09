function insertDriveLinksFromFolder() {

  const folderId = "1LW6UwuoaBKafbS7CM9P7Z6DQkvgZAw8J";

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();

  const sheet = SpreadsheetApp.getActiveSheet();
  const startRow = sheet.getActiveCell().getRow();
  const startCol = sheet.getActiveCell().getColumn();

  const rows = [];

  while (files.hasNext()) {
    const file = files.next();

    rows.push([
      file.getUrl(),   // kolom A (url)
      file.getName()   // kolom B (naam)
    ]);
  }

  if (rows.length > 0) {
    sheet.getRange(startRow, startCol, rows.length, 2).setValues(rows);
  }
}