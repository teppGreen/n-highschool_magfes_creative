function receptionRequest_temp() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('依頼受付（手動）','formタブの該当行番号を入力してください。',ui.ButtonSet.OK_CANCEL);
  const row = Number(prompt.getResponseText());
  
  if (row > 1 && prompt.getSelectedButton() === ui.Button.OK) {
    SpreadsheetApp.getUi().showModalDialog(startProcessingAnimation, `formタブ${row}行目の依頼を受付中`);
    receptionRequest(row);
    SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, `${row}行目の処理が完了しました`);
  }
}