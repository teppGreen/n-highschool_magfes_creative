function sheetId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  return ss.getId() + '.' + sheet.getSheetId();
}