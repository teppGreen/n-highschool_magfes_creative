function setTriggers() { //このシートのトリガーを設定
  deleteTriggers();
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onOpenFunctions').forSpreadsheet(sheet).onOpen().create();
  ScriptApp.newTrigger('onEditFunctions').forSpreadsheet(sheet).onEdit().create();
  ScriptApp.newTrigger('onFormSubmitFunctions').forSpreadsheet(sheet).onFormSubmit().create();
}

function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) ScriptApp.deleteTrigger(trigger);
}

function  createMenu() { 
  SpreadsheetApp.getUi().createMenu("GASメニュー")
    .addItem("制作依頼フォーム", "displayRequestForm_normal")
    .addSeparator()
    .addItem("制作依頼受付（手動）", "receptionRequest_temp")
    .addItem("Slack ワークフロー送信（手動）", "sendNotificationToSlack_fromResourceSheet")
    .addItem("リソース→制作管理シート 同期（手動）", "syncSheet_resourceToWork_temp")
    .addSeparator()
    .addItem("システムバグ報告/フィードバック", "contactSystemDeveloper")
    .addToUi();
}

function onFormSubmitFunctions(e) {
  try {
    console.log(`▼${arguments.callee.name}`);

    receptionRequest(e.range.getRow());
  } catch(error) {
    notifyError(error);
    throw new Error(error.stack);
  }
}

function onOpenFunctions(e) {
  try {
    console.log(`▼${arguments.callee.name}`);
    console.log('User: ' + e.user.getEmail());
    
    isScriptPropery();
    createMenu();
  } catch(error) {
    notifyError(error);
    throw new Error(error.stack);
  }
}

function onEditFunctions(e) {
  try {
    console.log(`▼${arguments.callee.name}`);
    console.log('User: ' + e.user.getEmail());
    
    integrityProjIdAndTitle(e);
    syncSheet_resourceToWork(e.source.getActiveSheet(), e.range.getRow());
    syncSheet_resourceToWork_status(e);
    inputAttendanceDate(e);
    openOrCloseFormResponse(e);
  } catch(error) {
    notifyError(error);
    throw new Error(error.stack);
  }
}

function onOpenFunctions_workSheet(e) {
  try {
    console.log(`▼${arguments.callee.name}`);
    console.log('User: ' + e.user.getEmail());
    
  } catch(error) {
    notifyError(error);
    throw new Error(error.stack);
  }
}

function onEditFunctions_workSheet(e) {
  try {
    console.log(`▼${arguments.callee.name}`);
    console.log('User: ' + e.user.getEmail() + '\nRange: ' + e.source.getActiveSheet().getName() + '!' + e.range.getA1Notation());

    syncSheet_workToResource(e);
    inputLastModifiedDate(e);
  } catch(error) {
    notifyError(error);
    throw new Error(error.stack);
  }
}

function notifyError(error) {
  const now = new Date();
  const datetime = Utilities.formatDate(now, 'JST', 'MM/dd HH:mm');
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const scriptUrl = 'https://script.google.com/home/projects/' + ScriptApp.getScriptId() + '/executions';

  const to = PropertiesService.getScriptProperties().getProperty('systemManagerEmails');
  const subject = `【${datetime}】【障害】磁実クリエイティブ班 業務システムでエラー発生`;
  const body = 
    '磁石祭実行委員会 クリエイティブ班 業務システムでエラーが発生しました。対応が必要な可能性がありますので、以下の内容を確認してください。' + 
    '\n\nエラー発生日時: ' + Utilities.formatDate(now, 'JST', 'yyyy/MM/dd(E) HH:mm') + 
    '\n' + error.stack + 
    '\n\nリソース管理シート: ' + sheetUrl + 
    '\nAppsScript: ' + scriptUrl
  const options = { name: 'リソース管理シート GASトリガー' };

  if (to) {
    GmailApp.sendEmail(to,subject,body,options);
  } else {
    console.error('スクリプトプロパティにシステム管理者のメールアドレスが設定されていなかったため、エラーメールを送信できませんでした。')
  }
}