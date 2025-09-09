function sendNotificationToSlack_fromResourceSheet() {  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(startProcessingAnimation, "処理中");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheetName = ss.getActiveSheet().getSheetName();
  const workSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.WORKS);

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

    workInfo.client.slackId = getSlackIdByEmail(workInfo.client.email);

    sendNotificationToSlack(workInfo);
    ui.showModalDialog(stopProcessingAnimation, `${row}行目をSlack ワークフローに送信しました`);
  } else {
    ui.showModalDialog(stopProcessingAnimation, `処理を中断しました`);
  }
}

function receptionRequest(formRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.WORKS);
  const projSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROJECTS);
  const paramSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PARAMETERS);
  const formSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.FORM);

  const { workInfo, requestInfo } = getFormattedFormResponse(formSheet, formRow);

  setProjectId(projSheet, workInfo);

  const workSheetRow = setWorkId(workSheet, workInfo);
  
  const urls = setupWorkEnvironment(paramSheet, workInfo, requestInfo);
  Object.assign(workInfo.url, urls);

  writeResponseToSheet_resource(workSheet,workSheetRow,workInfo);

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

function getFormattedFormResponse(formSheet, formRow) {
  const formResponse = formSheet.getRange(formRow, 1, 1, formSheet.getLastColumn()).getValues().flat();
  return formatFormResponse(formResponse);
}

function setProjectId(projSheet, workInfo) {
  workInfo.projId = findOrCreateProjectId(projSheet, workInfo.projTitle);
}

function setWorkId(workSheet, workInfo) {
  const { workId, workSheetRow } = determineWorkId(workSheet);
  workInfo.workId = workId;
  return workSheetRow;
}

function formatFormResponse(formResponse) {
  const workInfo = {
    genre: formResponse[4],
    projTitle: formResponse[2],
    workTitle: formResponse[3],
    status: CONFIG.STATUS.STEP1, //初期値
    datetime: { request: formResponse[0], expected: formResponse[10] },
    client: { email: formResponse[1] },
    url: { footageFolder: formResponse[9] },
  };

  const requestInfo = {
    content: formResponse[5], 
    design: formResponse[6],
    note: formResponse[12],
    regulation: formResponse[8],
    hearingType: formResponse[11],
    hearingDatetime: [formResponse[14],formResponse[15],formResponse[16]],
    reference: formResponse[7],
    systemCommand: formResponse[13].split(',').map(item => item.trim()),
  };

  return { workInfo, requestInfo };
}

function findOrCreateProjectId(projSheet, projTitle) {
  const projIdCol = getColByHeaderName(projSheet, CONFIG.HEADER_NAMES.PROJECT_ID);
  const projTitleCol = getColByHeaderName(projSheet, CONFIG.HEADER_NAMES.PROJECT_TITLE);
  
  let projSheetRow = projSheet.getRange(1, projTitleCol).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;
  if (projSheetRow === projSheet.getMaxRows() + 1) {
    projSheetRow = 2;
  }
  
  const projTitles = projSheet.getRange(1, projTitleCol, projSheetRow, 1).getValues().flat();
  const projTitleIndex = projTitles.indexOf(projTitle);
  
  let projId;
  if (projTitleIndex < 0) {
    // 新しい案件の場合
    projId = projSheet.getRange(projSheetRow, projIdCol).getValue();
    projSheet.getRange(projSheetRow, projTitleCol).setValue(projTitle);
  } else {
    // 既存の案件の場合
    projSheetRow = projTitleIndex + 1;
    projId = projSheet.getRange(projSheetRow, projIdCol).getValue();
  }
  
  return projId;
}

function determineWorkId(workSheet) {
  const workIdCol = getColByHeaderName(workSheet, CONFIG.HEADER_NAMES.WORK_ID);
  const workTitleCol = getColByHeaderName(workSheet, CONFIG.HEADER_NAMES.WORK_TITLE);
  let workSheetRow = workSheet.getRange(1, workTitleCol).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;
  
  let workId;
  if (workSheetRow === workSheet.getMaxRows() + 1) {
    workSheetRow = 2;
    workId = 1;
  } else {
    // 未発行の制作番号を求めるために、発行済みの制作番号の数字を全て足したものを、制作番号の要素数で割って、2倍する。
    const workIds = workSheet.getRange(2, workIdCol, workSheetRow - 2, 1).getValues().flat();
    let total = workIds.reduce(function(sum, element) {
      return sum + element;
    });
    workId = total / workIds.length * 2;
  }

  return { workId, workSheetRow };
}

function setupWorkEnvironment(paramSheet, workInfo, requestInfo) {
  const folders = createNewFolder(paramSheet, workInfo);
  const workSheetUrl = createWorkSheet(paramSheet, folders.workFolder, workInfo, requestInfo).getUrl();

  return {
    workFolder: folders.workFolder.getUrl(),
    deliveryFolder: folders.deliveryFolder.getUrl(),
    workSheet: workSheetUrl
  };
}

function createNewFolder(paramSheet, workInfo) {
  const systemStartYear = getValueRanges(CONFIG.PARAM_KEYS.SYSTEM_START_YEAR, paramSheet)[0].offset(0,1).getValue();
  const folderName = `${String(systemStartYear).slice(-2)}-${String(workInfo.workId).padStart(4,"0")}_${workInfo.projTitle}_${workInfo.workTitle}`;
  const parentFolderId = getValueRanges(CONFIG.PARAM_KEYS.WORK_FOLDER_URL, paramSheet)[0].offset(0,1).getValue();
  const parentFolder = DriveApp.getFolderById(parentFolderId); //親フォルダを指定します
  
  let folders = {};
  folders.workFolder = parentFolder.createFolder(folderName);
  folders.footageFolder = folders.workFolder.createFolder(CONFIG.FOLDER_PREFIX.MATERIAL + folderName);
  folders.deliveryFolder = folders.workFolder.createFolder(CONFIG.FOLDER_PREFIX.DELIVERY + folderName);

  //フォーム回答の素材フォルダのショートカットの作成
  const existingFootageFolderId = extractFileId(workInfo.url.footageFolder);
  if (existingFootageFolderId) {
    folders.footageFolder.createShortcut(existingFootageFolderId);
  }

  return folders;
}

function createWorkSheet(paramSheet, folder, workInfo) {
  const sheetName = `${CONFIG.FOLDER_PREFIX.WORKSHEET}${String(systemStartYear).slice(-2)}-${workInfo.projId}-${String(workInfo.workId).padStart(4,'0')}_${workInfo.projTitle}_${workInfo.workTitle}`;
  const parentSheetId = getValueRanges(CONFIG.PARAM_KEYS.WORK_SHEET_URL, paramSheet)[0].offset(0,1).getValue();
  const sheet = DriveApp.getFileById(parentSheetId).makeCopy(sheetName,folder);
  
  return sheet;
}

function writeResponseToSheet_resource(workSheet, workSheetRow, workInfo) {
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
  const statuslists = Object.values(CONFIG.STATUS).filter(status => status !== CONFIG.STATUS.CANCELLED);
  const inputStatus = new Array();

  for (let i = 0; i < statuslists.length; i++) {
    if (i === 0) {
      inputStatus.push([CONFIG.TASK_STATUS.IN_PROGRESS]); //依頼受付のステータスを「実行中」に
      continue;
    } else if (i === 1 && requestInfo.hearingType.includes('不要')) {
      inputStatus.push(['']); //ヒアリングが「基本的に不要」の場合は、初回ヒアリングのステータスを空欄（対応不要の意）に
      continue;
    } else {
      inputStatus.push([CONFIG.TASK_STATUS.NOT_STARTED]);
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

  let slackId = getSlackIdByEmail(workInfo.client.email);

  if (!slackId) {
    const slackAdminEmail = PropertiesService.getScriptProperties().getProperty(CONFIG.PROPERTIES.SLACK_ADMIN_EMAIL);
    slackId = getSlackIdByEmail(slackAdminEmail);

    const registrationFormId = getValueRanges(CONFIG.PARAM_KEYS.REGISTRATION_FORM_ID,paramSheet)[0].offset(0,1).getValue();
    const registrationFormUrl = `https://docs.google.com/forms/d/e/${registrationFormId}/viewform`;
    
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