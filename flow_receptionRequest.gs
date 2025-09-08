function sendNotificationToSlack_fromResourceSheet() {  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(startProcessingAnimation, "処理中");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheetName = ss.getActiveSheet().getSheetName();
  const workSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.WORKS);
  const paramSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PARAMETERS);

  const workIdCol = getColByHeaderName(workSheet,CONFIG.HEADER_NAMES.WORK_ID);
  const workTitleCol = getColByHeaderName(workSheet,CONFIG.HEADER_NAMES.WORK_TITLE);

  if (currentSheetName !== CONFIG.SHEET_NAMES.WORKS) {
    let workSheetRow = workSheet.getRange(1,workTitleCol).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
      if (workSheetRow === workSheet.getMaxRows()) workSheetRow = 2;
    workSheet.getRange(workSheetRow, workIdCol).activateAsCurrentCell();
    SpreadsheetApp.flush();
  }

  const prompt = ui.prompt('Slack ワークフロー送信（手動）','該当の制作番号を入力してください。',ui.ButtonSet.OK_CANCEL);
  const workId = Number(prompt.getResponseText());
  const row = getRowBySingleCol(workSheet, workIdCol, workId);

  if (row > 1 && prompt.getSelectedButton() === ui.Button.OK) {
    const outputRange = getRangesByHeaderNames(workSheet, row, headerNames_work);
    let workInfo = getValuesByRanges(outputRange);

    //依頼者SlackIDの特定
    const contactSheetId = getValueRanges(CONFIG.PARAM_KEYS.CONTACT_SHEET_ID,paramSheet)[0].offset(0,1).getValue();
    const contactSheet = SpreadsheetApp.openById(contactSheetId).getSheetByName(CONFIG.SHEET_NAMES.PERSONS);
    const emailCol = getColByHeaderName(contactSheet,CONFIG.HEADER_NAMES.EMAIL);
    const slackIdCol = getColByHeaderName(contactSheet,CONFIG.HEADER_NAMES.SLACK_ID);
    const emailList = contactSheet.getRange(1,emailCol,contactSheet.getLastRow(),1).getValues().flat();
    const contactSheetRow = emailList.indexOf(workInfo.client.email) + 1;

    if (contactSheetRow > 0) {
      workInfo.client.slackId = contactSheet.getRange(contactSheetRow,slackIdCol).getValue();
    }

    sendNotificationToSlack(workInfo);
    ui.showModalDialog(stopProcessingAnimation, `${row}行目をSlack ワークフローに送信しました`);
  } else {
    ui.showModalDialog(stopProcessingAnimation, `処理を中断しました`);
  }
}

function receptionRequest(formRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.WORKS);
  const formSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.FORM);
  const projSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROJECTS);
  const paramSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PARAMETERS);

  const formResponse = formSheet.getRange(formRow, 1, 1, formSheet.getLastColumn()).getValues().flat();
  
  let workInfo = {
    genre: formResponse[4],
    projTitle: formResponse[2],
    workTitle: formResponse[3],
    status: CONFIG.INITIAL_STATUS, //初期値
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
  const projIdCol = getColByHeaderName(projSheet,CONFIG.HEADER_NAMES.PROJECT_ID);
  const projTitleCol = getColByHeaderName(projSheet,CONFIG.HEADER_NAMES.PROJECT_TITLE);
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
  const workIdCol = getColByHeaderName(workSheet,CONFIG.HEADER_NAMES.WORK_ID);
  const workTitleCol = getColByHeaderName(workSheet,CONFIG.HEADER_NAMES.WORK_TITLE);
  let workSheetRow = workSheet.getRange(1,workTitleCol).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;
    if (workSheetRow === workSheet.getMaxRows() + 1) workSheetRow = 2;

  // 未発行の制作番号を求めるために、発行済みの制作番号の数字を全て足したものを、制作番号の要素数で割って、2倍する。
  const workIds = workSheet.getRange(2,workIdCol,workSheetRow-2,1).getValues().flat();
  let total = workIds.reduce(function(sum, element){
    return sum + element;
  });

  workInfo.workId = total / workIds.length * 2;
  

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

  // try {
  //   sendNotificationToSlack(workInfo,requestInfo);
  // } catch(error) {
  //   notifyError(error);
  // }
  
  processSystemCommand(requestInfo);
}

function createNewFolder(paramSheet, workInfo) {
  const systemStartYear = getValueRanges(CONFIG.PARAM_KEYS.SYSTEM_START_YEAR, paramSheet)[0].offset(0,1).getValue();
  const folderName = `${systemStartYear}-${String(workInfo.workId).padStart(4,"0")}_${workInfo.projTitle}_${workInfo.workTitle}`;
  const parentFolderId = getValueRanges(CONFIG.PARAM_KEYS.WORK_FOLDER_URL, paramSheet)[0].offset(0,1).getValue();
  const parentFolder = DriveApp.getFolderById(parentFolderId); //親フォルダを指定します
  
  let url = {};
  url.workFolder = parentFolder.createFolder(folderName);
  url.footageFolder = url.workFolder.createFolder(CONFIG.FOLDER_PREFIX.MATERIAL + folderName);
  url.deliveryFolder = url.workFolder.createFolder(CONFIG.FOLDER_PREFIX.DELIVERY + folderName);

  //フォーム回答の素材フォルダのショートカットの作成
  const existingFootageFolderId = extractFileId(workInfo.url.footageFolder);
  if (existingFootageFolderId) {
    url.footageFolder.createShortcut(existingFootageFolderId);
  }

  return url;
}

function createWorkSheet(paramSheet, folder, workInfo, requestInfo) {
  folder = DriveApp.getFolderById(extractFileId(folder));
  const sheetName = CONFIG.FOLDER_PREFIX.WORKSHEET + workInfo.projId + String(workInfo.workId).padStart(4,'0') + '_' + workInfo.projTitle + '_' + workInfo.workTitle;
  const parentSheetId = getValueRanges(CONFIG.PARAM_KEYS.WORK_SHEET_URL, paramSheet)[0].offset(0,1).getValue();
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
  const chameleons = resourceSheet.getRange(CONFIG.SHEET_NAMES.MEMBERS + '!E3:E30').getValues();
  
  const workSheet_main = workSheet.getSheetByName(CONFIG.SHEET_NAMES.MAIN);
    const resourceSheetId = PropertiesService.getScriptProperties().getProperty(CONFIG.PROPERTIES.RESOURCE_SHEET_ID);
    const resourceSheetUrl = `https://docs.google.com/spreadsheets/d/${resourceSheetId}/edit`;
    const resourceSheetLink = SpreadsheetApp.newRichTextValue().setText(CONFIG.UI.RESOURCE_SHEET_LINK_TEXT).setLinkUrl(resourceSheetUrl).build();
    getValueRanges(CONFIG.UI.RESOURCE_SHEET_LINK_TEXT, workSheet_main)[0].setRichTextValue(resourceSheetLink);
  
    getValueRanges(CONFIG.HEADER_NAMES.CONTENT,workSheet_main)[0].offset(1,0).setValue(requestInfo.content);
    getValueRanges(CONFIG.HEADER_NAMES.DESIGN,workSheet_main)[0].offset(1,0).setValue(requestInfo.design);
    getValueRanges(CONFIG.HEADER_NAMES.REGULATION,workSheet_main)[0].offset(1,0).setValue(requestInfo.regulation);
    getValueRanges(CONFIG.HEADER_NAMES.NOTE,workSheet_main)[0].offset(1,0).setValue(requestInfo.note);
    getValueRanges(CONFIG.HEADER_NAMES.HEARING,workSheet_main)[0].offset(0,1).setValue(requestInfo.hearingType);
    getValueRanges(CONFIG.HEADER_NAMES.REFERENCE,workSheet_main)[0].offset(1,0).setValue(requestInfo.reference);
    
    workSheet_main.getRange('C14:C41').setValues(chameleons);
    
    if(workInfo.url.footageFolder) {
      getValueRanges(CONFIG.HEADER_NAMES.OTHER_MEMO,workSheet_main)[0].offset(1,0).setValue('【注意】指定素材あり（/制作フォルダ/素材フォルダ）');
    }

  const workSheet_tasks = workSheet.getSheetByName(CONFIG.SHEET_NAMES.TASKS);
  const statusCol = getColByHeaderName(workSheet_tasks, CONFIG.HEADER_NAMES.STATUS);
  const statuslists = ['依頼受付','初回ヒアリング','制作','ブラッシュアップ','班長承認','納品'];
  const inputStatus = new Array();

  for (let i = 0; i < statuslists.length; i++) {
    if (i === 0) {
      inputStatus.push(['実行中']); //依頼受付のステータスを「実行中」に
      continue;
    } else if (i === 1 && requestInfo.hearingType.includes('不要')) {
      inputStatus.push(['']); //ヒアリングが「基本的に不要」の場合は、初回ヒアリングのステータスを空欄（対応不要の意）に
      continue;
    } else {
      inputStatus.push(['未着手']);
    }
  }
  
  workSheet_tasks.getRange(2,statusCol,inputStatus.length,1).setValues(inputStatus);
}

function sendNotificationToSlack(workInfo,requestInfo) {
  const paramSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.PARAMETERS);
  const token = PropertiesService.getScriptProperties().getProperty(CONFIG.PROPERTIES.SLACK_WORKFLOW_URL);
  //const token = PropertiesService.getScriptProperties().getProperty("slackWorkflow_test_WebReqestUrl");

  if (!token) {
    console.error('スクリプトプロパティが設定されていないため、Slack ワークフローに送信できません。')
    return;
  }

  let note = '';
  let datetime_hearing = '';

  if (requestInfo) {
    for (const command of requestInfo.systemCommand) {
      if (command === CONFIG.SYSTEM_COMMANDS.DONT_SEND_NOTIFICATION) return;
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
     datetime_hearing = `＜初回ヒアリング実施日時＞ 出席可能な候補をスタンプで教えてください。
      :one: ${datetime_hearing[0]}
      :two:${datetime_hearing[1]}
      :three:${datetime_hearing[2]}`;
    }

    note = requestInfo.note;
  }

  //依頼者SlackIDの特定
  const slackAdminEmail = PropertiesService.getScriptProperties().getProperty(CONFIG.PROPERTIES.SLACK_ADMIN_EMAIL);
  const contactSheetId = getValueRanges(CONFIG.PARAM_KEYS.CONTACT_SHEET_ID,paramSheet)[0].offset(0,1).getValue();
  const contactSheet = SpreadsheetApp.openById(contactSheetId).getSheetByName(CONFIG.SHEET_NAMES.PERSONS);
  const emailCol = getColByHeaderName(contactSheet,CONFIG.HEADER_NAMES.EMAIL);
  const slackIdCol = getColByHeaderName(contactSheet,CONFIG.HEADER_NAMES.SLACK_ID);
  const emailList = contactSheet.getRange(1,emailCol,contactSheet.getLastRow(),1).getValues().flat();
  let contactSheetRow = emailList.indexOf(workInfo.client.email) + 1;
  let slackId;

  if (contactSheetRow > 0) {
    slackId = contactSheet.getRange(contactSheetRow,slackIdCol).getValue();
  } else {
    const registrationFormId = getValueRanges(CONFIG.PARAM_KEYS.REGISTRATION_FORM_ID,paramSheet)[0].offset(0,1).getValue();
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