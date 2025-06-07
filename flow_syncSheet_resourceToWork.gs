function syncSheet_resourceToWork_temp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resourceSheet = ss.getSheetByName('works');
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('リソース→制作管理シート 同期（手動）','worksタブの該当行番号を入力してください。',ui.ButtonSet.OK_CANCEL);
  const row = Number(prompt.getResponseText());

  if (row > 1 && prompt.getSelectedButton() === ui.Button.OK) {
    SpreadsheetApp.getUi().showModalDialog(startProcessingAnimation, "同期中");
    syncSheet_resourceToWork(resourceSheet,row);
    SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, `${row}行目の同期が完了しました`);
  } else {
    ss.toast('処理を中断しました')
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
  
  getValueRanges('管理番号', workSheet_main)[0].offset(0,2).setValue(workInfo.projId);
  getValueRanges('管理番号', workSheet_main)[0].offset(0,3).setValue(workInfo.workId);
  getValueRanges('タイトル', workSheet_main)[0].offset(0,2).setValue(workInfo.projTitle);
  getValueRanges('タイトル', workSheet_main)[0].offset(0,3).setValue(workInfo.workTitle);
  getValueRanges('ジャンル', workSheet_main)[0].offset(0,2).setValue(workInfo.genre);
  getValueRanges('依頼者', workSheet_main)[0].offset(0,2).setValue(workInfo.client.nickname);
  getValueRanges('依頼者', workSheet_main)[0].offset(0,3).setValue(workInfo.client.department);
  getValueRanges('依頼者メアド', workSheet_main)[0].offset(0,2).setValue(workInfo.client.email);

  getValueRanges('制作アプリ', workSheet_main)[0].offset(0,1).setValue(workInfo.review.usedApp);
  getValueRanges('成果物数', workSheet_main)[0].offset(0,1).setValue(workInfo.review.deliverablesCount);
  getValueRanges('来年も作るべきか', workSheet_main)[0].offset(0,1).setValue(workInfo.review.willMakeNextYear);

  const urlLabels = ['制作フォルダ','納品フォルダ','Canva フォルダ','Slack チャンネル','Slack スレッド'];
  for (let key in workInfo.url) {
    const urlIndex = urlLabels.indexOf(headerNames_work['url'][key]);
    if (urlIndex >= 0) {
      const url = workInfo['url'][key];
      if (url) {
        const richtext = SpreadsheetApp.newRichTextValue().setText(urlLabels[urlIndex]).setLinkUrl(url).build();
        getValueRanges(urlLabels[urlIndex], workSheet_main)[0].setRichTextValue(richtext);
      }
    }
  }

  getValueRanges('依頼受付', workSheet_tasks)[0].offset(0,1).setValue(workInfo.datetime.request);
  getValueRanges('初回ヒアリング', workSheet_tasks)[0].offset(0,1).setValue(workInfo.datetime.hearing);
  getValueRanges('制作', workSheet_tasks)[0].offset(0,1).setValue(workInfo.datetime.creating);
  getValueRanges('ブラッシュアップ', workSheet_tasks)[0].offset(0,1).setValue(workInfo.datetime.refining);
  getValueRanges('班長承認', workSheet_tasks)[0].offset(0,1).setValue(workInfo.datetime.approval);
  getValueRanges('納品', workSheet_tasks)[0].offset(0,1).setValue(workInfo.datetime.delivery);
  getValueRanges('納品', workSheet_tasks)[0].offset(0,2).setValue(workInfo.datetime.expected);

  const oldJoinedMembers_range = workSheet_main.getRange('C14:C41');
  const newJoinedMembers = workInfo.joinedMembers.split(',').map(item => item.trim());

  oldJoinedMembers_range.offset(0,-1).setValue(false);

  for (const newJoinedMember of newJoinedMembers) {
    const nameRange = getValueRanges(newJoinedMember, oldJoinedMembers_range);
    if (nameRange) {
      const inputRange = nameRange[0].offset(0,-1);
      inputRange.setValue(true);
    }
  }

  changeFileName(workInfo);
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
  
  if (e.value === '納品' || e.value === '依頼取消') { 
    workSheet_tasks.getRange(newStatusRow, statusCol).setValue('完了');
  } else {
    workSheet_tasks.getRange(newStatusRow, statusCol).setValue('実行中');
  }

  if (e.oldValue !== '依頼取消') {
    const oldStatusRow = getValueRanges(e.oldValue, workSheet_tasks)[0].getRow();
    workSheet_tasks.getRange(oldStatusRow, statusCol).setValue('完了');
    workSheet_tasks.getRange(oldStatusRow, endDatetimeCol).setValue(now);
  }
}