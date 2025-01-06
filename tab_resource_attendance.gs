function inputAttendanceDate(e) {
  console.log('inputAttendanceDate');

  const ss = e.source;
  const sheet = ss.getActiveSheet();
  const targetSheetName = 'attendance';

  if (sheet.getName() !== targetSheetName || e.range.getRow() < 2) return;
  
  console.log('continue inputAttendanceDate');
  const range = sheet.getRange(1, e.range.getColumn());
  const now = new Date();
  
  if(!range.getValue()) range.setValue(now);
}