function sendNotificationToSlack_fromResourceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resourceSheet = ss.getSheetByName('works');
  const paramSheet = ss.getSheetByName('parameters');
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('Slack ワークフロー送信（手動）','worksタブの該当行番号を入力してください。',ui.ButtonSet.OK_CANCEL);
  const row = Number(prompt.getResponseText());

  if (row > 1 && prompt.getSelectedButton() === ui.Button.OK) {
    SpreadsheetApp.getUi().showModalDialog(startProcessingAnimation, "処理中");
    
    const outputRange = getRangesByHeaderNames(resourceSheet, row, headerNames_work);
    let workInfo = getValuesByRanges(outputRange);

    //依頼者SlackIDの特定
    const contactSheetId = getValueRanges('contactSheet.id',paramSheet)[0].offset(0,1).getValue();
    const contactSheet = SpreadsheetApp.openById(contactSheetId).getSheetByName('persons');
    const emailCol = getColByHeaderName(contactSheet,'E-mail 1 - Value');
    const slackIdCol = getColByHeaderName(contactSheet,'Slack ID');
    const emailList = contactSheet.getRange(1,emailCol,contactSheet.getLastRow(),1).getValues().flat();
    const contactSheetRow = emailList.indexOf(workInfo.client.email) + 1;

    if (contactSheetRow > 0) {
      workInfo.client.slackId = contactSheet.getRange(contactSheetRow,slackIdCol).getValue();
    }

    sendNotificationToSlack(workInfo);
    SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, `${row}行目をSlack ワークフローに送信しました`);
  } else {
    ss.toast('処理を中断しました')
  }
}

function receptionRequest(formRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workSheet = ss.getSheetByName('works');
  const formSheet = ss.getSheetByName('form');
  const projSheet = ss.getSheetByName('projects');
  const paramSheet = ss.getSheetByName('parameters');

  const formResponse = formSheet.getRange(formRow, 1, 1, formSheet.getLastColumn()).getValues().flat();
  
  let workInfo = {
    genre: formResponse[4],
    projTitle: formResponse[2],
    workTitle: formResponse[3],
    status: '依頼受付', //初期値
    datetime: { request: formResponse[0], expected: formResponse[10] },
    client: { email: formResponse[1] },
    url: { footageFolder: formResponse[9] },
  }

  let requestInfo = {
    content: formResponse[5], 
    design: formResponse[6],
    note: formResponse[12],
    regulation: formResponse[8],
    hearingType: formResponse[11],
    hearingDatetime: [formResponse[14],formResponse[15],formResponse[16]],
    reference: formResponse[7],
    systemCommand: []
  }

  // システムコマンドを配列にして入れる
  requestInfo.systemCommand = formResponse[13].split(',').map(item => item.trim());

  //案件番号の決定
  const projIdCol = getColByHeaderName(projSheet,'案件番号');
  const projTitleCol = getColByHeaderName(projSheet,'案件タイトル');
  let projSheetRow = projSheet.getRange(1,projTitleCol).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;
    if (projSheetRow === projSheet.getMaxRows() + 1) projSheetRow = 2;
  const projTitles = projSheet.getRange(1, projTitleCol, projSheetRow, 1).getValues().flat();
  const projTitleIndex = projTitles.indexOf(workInfo.projTitle);
  
  if(projTitleIndex < 0) {
    workInfo.projId = projSheet.getRange(projSheetRow, projIdCol).getValue();
  } else {
    projSheetRow = projTitleIndex+1;
    workInfo.projId = projSheet.getRange(projSheetRow, projIdCol).getValue();
  }
  projSheet.getRange(projSheetRow,projTitleCol).setValue(workInfo.projTitle);

  //制作番号の決定
  const workIdCol = getColByHeaderName(workSheet,'制作番号');
  const workTitleCol = getColByHeaderName(workSheet,'制作タイトル');
  let workSheetRow = workSheet.getRange(1,workTitleCol).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;
    if (workSheetRow === workSheet.getMaxRows() + 1) workSheetRow = 2;
  
  workInfo.workId = workSheet.getRange(workSheetRow, workIdCol).getValue();

  //制作フォルダ・制作シートの作成
  const newFolder = createNewFolder(paramSheet, workInfo);
  workInfo.url.workFolder = newFolder.workFolder.getUrl();
  workInfo.url.deliveryFolder = newFolder.deliveryFolder.getUrl();
  workInfo.url.workSheet = createWorkSheet(paramSheet, workInfo.url.workFolder, workInfo, requestInfo).getUrl();

  writeResponseToSheet_resource(projSheet,workSheet,workSheetRow,workInfo);

  try {
    writeResponseToSheet_work(workInfo,requestInfo);
    syncSheet_resourceToWork(workSheet,workSheetRow);
  } catch(error) {
    console.error('Continue error: ' + error.stack);
  }

  try {
    sendNotificationToSlack(workInfo,requestInfo);
  } catch(error) {
    notifyError(error);
  }
  
  processSystemCommand(requestInfo);
}

function createNewFolder(paramSheet, workInfo) {
  const folderName = `${workInfo.workId}_${workInfo.projTitle}_${workInfo.workTitle}`;
  const parentFolderId = getValueRanges('workInfo.url.workFolder', paramSheet)[0].offset(0,1).getValue();
  const parentFolder = DriveApp.getFolderById(parentFolderId); //親フォルダを指定します
  
  let url = {};
  url.workFolder = parentFolder.createFolder(folderName);
  url.footageFolder = url.workFolder.createFolder('【素材】' + folderName);
  url.deliveryFolder = url.workFolder.createFolder('【納品】' + folderName);

  //フォーム回答の素材フォルダのショートカットの作成
  const existingFootageFolderId = extractFileId(workInfo.url.footageFolder);
  if (existingFootageFolderId) {
    url.footageFolder.createShortcut(existingFootageFolderId);
  }

  return url;
}

function createWorkSheet(paramSheet, folder, workInfo, requestInfo) {
  folder = DriveApp.getFolderById(extractFileId(folder));
  const sheetName = '【制作管理】' + workInfo.projId + String(workInfo.workId).padStart(3,'0') + '_' + workInfo.projTitle + '_' + workInfo.workTitle;
  const parentSheetId = getValueRanges('workInfo.url.workSheet', paramSheet)[0].offset(0,1).getValue();
  const sheet = DriveApp.getFileById(parentSheetId).makeCopy(sheetName,folder);
  
  return sheet;
}

function writeResponseToSheet_resource(projSheet, workSheet, workSheetRow, workInfo) {
  const range = getRangesByHeaderNames(workSheet, workSheetRow, headerNames_work);

  function processObject(obj,obj2) {
    for (let key in obj) {
      if (typeof obj[key] === 'object' && obj2[key]) {
        if (obj[key].getA1Notation) {
          obj[key].setValue(obj2[key]);
        } else {
          processObject(obj[key],obj2[key]); // ネストされたオブジェクトを再帰的に処理
        }
      }
    }

  }

  processObject(range,workInfo)
}

function writeResponseToSheet_work(workInfo,requestInfo){
  const resourceSheet = SpreadsheetApp.getActiveSpreadsheet();
  const workSheet = SpreadsheetApp.openByUrl(workInfo.url.workSheet)

  //クリエイティブ班員の名前を取得
  const chameleons = resourceSheet.getRange('members!E3:E30').getValues();
  
  const workSheet_main = workSheet.getSheetByName('main');
    const resourceSheetId = PropertiesService.getScriptProperties().getProperty('sheetId_resource');
    const resourceSheetUrl = `https://docs.google.com/spreadsheets/d/${resourceSheetId}/edit`;
    const resourceSheetLink = SpreadsheetApp.newRichTextValue().setText('リソース管理シート').setLinkUrl(resourceSheetUrl).build();
    getValueRanges('リソース管理シート', workSheet_main)[0].setRichTextValue(resourceSheetLink);
  
    getValueRanges('内容',workSheet_main)[0].offset(1,0).setValue(requestInfo.content);
    getValueRanges('デザイン要項',workSheet_main)[0].offset(1,0).setValue(requestInfo.design);
    getValueRanges('入稿規定',workSheet_main)[0].offset(1,0).setValue(requestInfo.regulation);
    getValueRanges('依頼備考',workSheet_main)[0].offset(1,0).setValue(requestInfo.note);
    getValueRanges('ヒアリング',workSheet_main)[0].offset(0,1).setValue(requestInfo.hearingType);
    getValueRanges('参考物',workSheet_main)[0].offset(1,0).setValue(requestInfo.reference);
    
    workSheet_main.getRange('C14:C41').setValues(chameleons);
    
    if(workInfo.url.footageFolder) {
      getValueRanges('その他メモ',workSheet_main)[0].offset(1,0).setValue('【注意】指定素材あり（/制作フォルダ/素材フォルダ）');
    }

  const workSheet_tasks = workSheet.getSheetByName('tasks');
  const statusCol = getColByHeaderName(workSheet_tasks, 'ステータス');
    if (requestInfo.hearingType.includes('不要')) {
      const targetStatusRow = getValueRanges('初回ヒアリング',workSheet_tasks)[0].getRow();
      workSheet_tasks.getRange(targetStatusRow, statusCol).clearContent();
    }
}

function sendNotificationToSlack(workInfo,requestInfo) {
  const paramSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('parameters');
  const token = PropertiesService.getScriptProperties().getProperty("slackWorkflow_notifyRequest_WebReqestUrl");
  //const token = PropertiesService.getScriptProperties().getProperty("slackWorkflow_test_WebReqestUrl");

  if (!token) {
    console.error('スクリプトプロパティが設定されていないため、Slack ワークフローに送信できません。')
    return;
  }

  let note = '';
  let datetime_hearing = '';

  if (requestInfo) {
    for (const command of requestInfo.systemCommand) {
      if (command === 'dontSendNotification') return;
    }

    if (requestInfo.hearingDatetime[0] !== '') {
      datetime_hearing = [];
      for (let i = 0; i < 3; i++) {
        if (requestInfo.hearingDatetime[i] !== '') {
          datetime_hearing.push(Utilities.formatDate(requestInfo.hearingDatetime[i], 'JST', 'MM/dd(E) HH:mm'));
        } else {
          datetime_hearing.push('-');
        }
      }
     datetime_hearing = `＜初回ヒアリング実施日時＞ 出席可能な候補をスタンプで教えてください。\n:one: ${datetime_hearing[0]}\n:two:${datetime_hearing[1]}\n:three:${datetime_hearing[2]}`; 
    }

    note = requestInfo.note;
  }

  //依頼者SlackIDの特定
  const slackAdminEmail = PropertiesService.getScriptProperties().getProperty('slackAdminEmail');
  const contactSheetId = getValueRanges('contactSheet.id',paramSheet)[0].offset(0,1).getValue();
  const contactSheet = SpreadsheetApp.openById(contactSheetId).getSheetByName('persons');
  const emailCol = getColByHeaderName(contactSheet,'E-mail 1 - Value');
  const slackIdCol = getColByHeaderName(contactSheet,'Slack ID');
  const emailList = contactSheet.getRange(1,emailCol,contactSheet.getLastRow(),1).getValues().flat();
  let contactSheetRow = emailList.indexOf(workInfo.client.email) + 1;
  let slackId;

  if (contactSheetRow > 0) {
    slackId = contactSheet.getRange(contactSheetRow,slackIdCol).getValue();
  } else {
    const registrationFormId = getValueRanges('registrationForm.id',paramSheet)[0].offset(0,1).getValue();
    const registrationFormUrl = `https://docs.google.com/forms/d/e/${registrationFormId}/viewform`;
    contactSheetRow = emailList.indexOf(slackAdminEmail) + 1;

    if (contactSheetRow > 0) {
      slackId = contactSheet.getRange(contactSheetRow,slackIdCol).getValue();
    }

    // 備考に基本情報収集フォームが送られていないことを書き足す
    if (note.length === 0) {
      note += '\n\n';
    }

    note += `依頼者(${workInfo.client.email})から基本情報収集フォームが提出されていなかったため、システム管理者をメンションしました。` + 
      `基本情報収集フォーム: ${registrationFormUrl}`;
  }
  
  const datetime_expected = workInfo.datetime.expected ? Utilities.formatDate(workInfo.datetime.expected, 'JST', 'yyyy/MM/dd(E) HH:mm') : '依頼時点では未指定';

  const params = {
    method : 'post',
    contentType: 'application/json',
    payload : JSON.stringify({
      "url_workSheet": workInfo.url.workSheet,
      "datetime_request": Utilities.formatDate(workInfo.datetime.request, 'JST', 'yyyy/MM/dd(E) HH:mm'),
      "slackChannel": workInfo.projId + '-' + workInfo.projTitle,
      "datetime_hearing": datetime_hearing,
      "workId": String(workInfo.workId).padStart(3,'0'),
      "note": note,
      "datetime_expected": datetime_expected,
      "projTitle": workInfo.projTitle,
      "client_slackId": slackId,
      "url_workFolder": workInfo.url.workFolder,
      "workTitle": workInfo.workTitle,
      "projId": workInfo.projId
    })
  };

  const res = UrlFetchApp.fetch(token, params);
  console.log(res);
}