function syncSheet_workToResource(e) {
  const resourceSheetId = PropertiesService.getScriptProperties().getProperty('sheetId_resource');
  const resourceSheet = SpreadsheetApp.openById(resourceSheetId).getSheetByName('works');
  const workSheet = e.source;
  const workSheet_main = workSheet.getSheetByName('main');
  const workSheet_active = e.source.getActiveSheet();

  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();
  const editedHeader = workSheet_active.getRange(1, editedCol).getValue();

  const workId = getValueRanges('管理番号',workSheet_main)[0].offset(0,3).getValue();
  const workIds = resourceSheet.getRange(1,getColByHeaderName(resourceSheet,'制作番号'),resourceSheet.getLastRow(),1).getValues().flat();
  const workSheetRow = workIds.indexOf(workId) + 1;

  const targetDataValidation = e.range.getDataValidation();
  let targetCreteriaType;
  if (targetDataValidation) {
    targetCreteriaType = targetDataValidation.getCriteriaType();
  }

  if (workSheetRow === 0) return;
  
  if (workSheet_active.getName() === 'main') {
    if (targetCreteriaType === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      const nickname = e.range.offset(0,1).getValue();
      const membersCol = getColByHeaderName(resourceSheet,'担当者');
      const membersRange = resourceSheet.getRange(workSheetRow,membersCol);
      let members = membersRange.getValue();
      if (members !== '') {
        members = members.split(',').map(item => item.trim());
      } else {
        members = [];
      }

      if (e.value == 'TRUE') {
        members.push(nickname);
      } else if (e.value == 'FALSE') {
        const index = members.indexOf(nickname);
        if (index >= 0) members.splice(index, 1);
      }
      
      members = members.join(',');
      membersRange.setValue(members);
    }

    const urlLabels = ['制作フォルダ','納品フォルダ','Canva フォルダ','Slack チャンネル','Slack スレッド'];
    const urlIndex = urlLabels.indexOf(e.value);
    if (urlIndex >= 0) {
      const urlLabel = urlLabels[urlIndex];
      const url = e.range.getRichTextValue().getLinkUrl();
      const urlRange = resourceSheet.getRange(workSheetRow,getColByHeaderName(resourceSheet,urlLabel));
      urlRange.setValue(url);
    }

    const reviewLabels = ['制作アプリ','成果物数','来年も作るべきか'];
    const editedLabel = e.range.offset(0,-1).getValue();
    const labelIndex = reviewLabels.indexOf(editedLabel);
    console.log(reviewLabels);
    console.log('editedLabel: ' + editedLabel)
    console.log('labelIndex: ' + labelIndex)

    if (labelIndex >= 0) {
      console.log('input review');
      const inputCol = getColByHeaderName(resourceSheet,reviewLabels[labelIndex]);
      resourceSheet.getRange(workSheetRow,inputCol).setValue(e.value);
    }
  }

  if (workSheet_active.getName() === 'tasks') {
    const statusList = ['依頼受付','初回ヒアリング','制作','ブラッシュアップ','班長承認','納品'];
    const editedTitleRange = workSheet_active.getRange(editedRow,getColByHeaderName(workSheet_active,'タイトル'));
    const editedTitle = editedTitleRange.getValue();
    
    const startDatetimeRange = workSheet_active.getRange(editedRow,getColByHeaderName(workSheet_active,'開始日時'));
    const startDatetime = startDatetimeRange.getValue();
    const endDatetimeRange = workSheet_active.getRange(editedRow,getColByHeaderName(workSheet_active,'終了日時'));
    const endDatetime = endDatetimeRange.getValue();
    
    if (statusList.includes(editedTitle)) {
      // ステータス変更を処理
      if (editedHeader === 'ステータス') {
        if ((e.value === '実行中') || (editedTitle === '納品' && e.value === '完了')) {
          const resourceSheetStatusCol = getColByHeaderName(resourceSheet,'ステータス');
          resourceSheet.getRange(workSheetRow, resourceSheetStatusCol).setValue(editedTitle);
        }
      }

      // 日時変更を処理
      if (editedTitle === '納品') {
        resourceSheet.getRange(workSheetRow, getColByHeaderName(resourceSheet, `${editedTitle}日時`)).setValue(startDatetime);
        resourceSheet.getRange(workSheetRow, getColByHeaderName(resourceSheet, `${editedTitle}期限日時`)).setValue(endDatetime);
      } else {
        resourceSheet.getRange(workSheetRow, getColByHeaderName(resourceSheet, `${editedTitle}開始日時`)).setValue(startDatetime);
      }
    }
  }
}