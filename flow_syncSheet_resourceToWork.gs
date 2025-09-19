function syncSheet_resourceToWork_temp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheetName = ss.getActiveSheet().getSheetName();
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(startProcessingAnimation, "処理中");

  if (currentSheetName !== 'works') {
    const formSheet = ss.getSheetByName('works');
    formSheet.getRange(2,1).activateAsCurrentCell();
    SpreadsheetApp.flush();
  }

  const prompt = ui.prompt('リソース→制作管理シート 同期（手動）','該当の制作番号を入力してください。',ui.ButtonSet.OK_CANCEL);

  if (prompt.getSelectedButton() === ui.Button.OK) {
    const workSheet = ss.getSheetByName('works');
    const workIdCol = getColByHeaderName(workSheet,'制作番号');
    const workId = Number(prompt.getResponseText());
    const row = getRowBySingleCol(workSheet, workIdCol, workId);
    if (row > 1) {
      syncSheet_resourceToWork(workSheet,row);
      SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, `同期が完了しました`);
    }
  } else {
    SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, `処理を中断しました`);
  }
}

function syncSheet_resourceToWork(sheet,row){
  if (sheet.getName() !== 'works' || row < 2) return;

  const outputRange = getRangesByHeaderNames(sheet, row, headerNames_work);
  const workInfo = getValuesByRanges(outputRange);
  console.log(workInfo);

  const workSheetUrl = workInfo.url.workSheet;
  if (!workSheetUrl || workSheetUrl == '') {
    console.error('[syncSheet_resourceToWork] workSheetUrl was not found.');
    return;
  }

  const workSheet = SpreadsheetApp.openByUrl(workSheetUrl);
  const workSheet_main = workSheet.getSheetByName('main');
  const workSheet_tasks = workSheet.getSheetByName('tasks');
  
  getValueRanges('管理番号', workSheet_main)[0].offset(0,1).setValue(workInfo.manageId);
  getValueRanges('制作タイトル', workSheet_main)[0].offset(0,1).setValue(workInfo.workTitle);
  getValueRanges('案件タイトル', workSheet_main)[0].offset(0,1).setValue(workInfo.projTitle);
  getValueRanges('ジャンル', workSheet_main)[0].offset(0,2).setValue(workInfo.genre);
  getValueRanges('依頼者', workSheet_main)[0].offset(0,2).setValue(workInfo.client.nickname);
  getValueRanges('依頼者', workSheet_main)[0].offset(0,3).setValue(workInfo.client.department);

  getValueRanges('制作アプリ', workSheet_main)[0].offset(0,1).setValue(workInfo.review.usedApp);
  getValueRanges('成果物数', workSheet_main)[0].offset(0,1).setValue(workInfo.review.deliverablesCount);
  getValueRanges('来年も作るべきか', workSheet_main)[0].offset(0,1).setValue(workInfo.review.willMakeNextYear);

  
  const generalSheetLabel = 'リソース管理シート';
  const generalSheetUrl = `https://docs.google.com/spreadsheets/d/${sheet.getParent().getId()}/edit#gid=${sheet.getSheetId()}&range=A${row}`;
  const generalSheetRichtext = SpreadsheetApp.newRichTextValue().setText(generalSheetLabel).setLinkUrl(generalSheetUrl).build();
  getValueRanges(generalSheetLabel, workSheet_main)[0].setRichTextValue(generalSheetRichtext);

  const urlLabels = ['制作フォルダ','納品フォルダ','Canva フォルダ','Slack チャンネル','Slack スレッド','案件ドキュメント','案件フォルダ'];
  for (let key in workInfo.url) {
    const urlIndex = urlLabels.indexOf(headerNames_work['url'][key]);
    if (urlIndex >= 0) {
      const url = workInfo['url'][key];
      const richtext = SpreadsheetApp.newRichTextValue().setText(urlLabels[urlIndex]).setLinkUrl(url).build();
      getValueRanges(urlLabels[urlIndex], workSheet_main)[0].setRichTextValue(richtext);
    }
  }

  getValueRanges(CONFIG.STATUS.STEP1, workSheet_tasks)[0].offset(0,1).setValue(workInfo.datetime.request);
  getValueRanges(CONFIG.STATUS.STEP2, workSheet_tasks)[0].offset(0,1).setValue(workInfo.datetime.hearing);
  getValueRanges(CONFIG.STATUS.STEP3, workSheet_tasks)[0].offset(0,1).setValue(workInfo.datetime.creating);
  getValueRanges(CONFIG.STATUS.STEP4, workSheet_tasks)[0].offset(0,1).setValue(workInfo.datetime.refining);
  getValueRanges(CONFIG.STATUS.STEP5, workSheet_tasks)[0].offset(0,1).setValue(workInfo.datetime.approval);
  getValueRanges(CONFIG.STATUS.STEP6, workSheet_tasks)[0].offset(0,1).setValue(workInfo.datetime.delivery);
  getValueRanges(CONFIG.STATUS.STEP6, workSheet_tasks)[0].offset(0,2).setValue(workInfo.datetime.expected);

  const oldJoinedMembers_range = workSheet_main.getRange('C14:C41');
  const newJoinedMembers = workInfo.joinedMembers.split(',').map(item => item.trim());

  oldJoinedMembers_range.offset(0,-1).setValue(false);

  for (const newJoinedMember of newJoinedMembers) {
    const currentNames = oldJoinedMembers_range.getValues().flat();
    let rowIndex = currentNames.indexOf(newJoinedMember);

    if (rowIndex < 0) {
      rowIndex = currentNames.indexOf('');
        if (rowIndex < 0) {
          throw new Error(`管理番号: ${workInfo.manageId}\n新しく担当となったmemberの名前の追加を試みましたが、人数の上限に達していたためできませんでした。`);
        }
      oldJoinedMembers_range.offset(rowIndex,0,1,1).setValue(newJoinedMember);
    }

    if (newJoinedMember) {
      oldJoinedMembers_range.offset(rowIndex,-1,1,1).setValue(true);
    }
  }

  changeFileName_work(workInfo);
}

function syncSheet_resourceToWork_status(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'works' || !e.value) return;

  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();
  if (sheet.getRange(1,editedCol).getValue() !== 'ステータス') return;
  
  const workSheet_url = sheet.getRange(editedRow,getColByHeaderName(sheet,'制作シート')).getValue();
  const workSheet = SpreadsheetApp.openByUrl(workSheet_url);
  const workSheet_tasks = workSheet.getSheetByName('tasks');
  const newStatusRow = getValueRanges(e.value, workSheet_tasks)[0].getRow();
  const statusCol = getColByHeaderName(workSheet_tasks, 'ステータス');
  const endDatetimeCol = getColByHeaderName(workSheet_tasks, '終了日時');
  const now = new Date();
  
  if (e.value === CONFIG.STATUS.STEP6 || e.value === CONFIG.STATUS.CANCELLED) { 
    workSheet_tasks.getRange(newStatusRow, statusCol).setValue(CONFIG.TASK_STATUS.DONE);
  } else {
    workSheet_tasks.getRange(newStatusRow, statusCol).setValue(CONFIG.TASK_STATUS.IN_PROGRESS);
  }

  if (e.oldValue && e.oldValue !== CONFIG.STATUS.STEP6 && e.oldValue !== CONFIG.STATUS.CANCELLED) {
    const oldStatusRow = getValueRanges(e.oldValue, workSheet_tasks)[0].getRow();
    workSheet_tasks.getRange(oldStatusRow, statusCol).setValue(CONFIG.TASK_STATUS.DONE);
    workSheet_tasks.getRange(oldStatusRow, endDatetimeCol).setValue(now);
  }
}