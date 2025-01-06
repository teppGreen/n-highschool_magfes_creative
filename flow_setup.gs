function setTriggers_workSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  // スプレッドシートにローディングアニメーションを表示
  SpreadsheetApp.getUi().showModalDialog(startProcessingAnimation, "処理中");

  const drawings = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDrawings();
  for (let drawing of drawings) {
    if (drawing.getOnAction() === 'setTriggers') drawing.remove()
  }

  ScriptApp.newTrigger('onEditFunctions').forSpreadsheet(sheet).onEdit().create();
  ScriptApp.newTrigger('onOpenFunctions').forSpreadsheet(sheet).onOpen().create();
  
  //ローディングアニメーションを閉じる
  SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, '処理が完了しました');

  sheet.toast('今後はリソース管理シートとデータが同期されます','同期を開始しました');
}

function isScriptPropery() {
  const scriptProperties = PropertiesService.getScriptProperties();

  if (scriptProperties.length === 0) {
    Browser.msgBox('初期設定がされていないため利用できません','システム管理者にスクリプトプロパティの設定を依頼してください。',Browser.Buttons.OK)
  }
}