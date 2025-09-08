function receptionRequest_temp() {
  SpreadsheetApp.getUi().showModalDialog(startProcessingAnimation, "処理中"); 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const sheetName = sheet.getSheetName();


  if (sheetName !== 'form') {
    const formSheet = ss.getSheetByName('form');
    const lastRow = formSheet.getLastRow();
    formSheet.getRange(lastRow,1).activateAsCurrentCell();
    SpreadsheetApp.flush();
  }

  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('依頼受付（手動）','formタブの該当行番号を入力してください。',ui.ButtonSet.OK_CANCEL);
  const row = Number(prompt.getResponseText());
  
  if (row > 1 && prompt.getSelectedButton() === ui.Button.OK) {
    receptionRequest(row);
    SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, `${row}行目の処理が完了しました`);
  } else {
    SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, `処理を中断しました`);
  }

  activeRange.activateAsCurrentCell();
}