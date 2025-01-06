function generateProjectNumbers() { //AA-ZZの案件番号を作成
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('projects');

  let prefixes = [];
  for (let i = 'A'.charCodeAt(0); i <= 'Z'.charCodeAt(0); i++) {
    prefixes.push(String.fromCharCode([i]));
  }

  let numbers = [];
  for (let i = 0; i < prefixes.length; i++) {
    for (let j = 0; j < prefixes.length; j++) {
      numbers.push([prefixes[i] + prefixes[j]]);
      console.log(prefixes[i] + prefixes[j]);
    }
  }

  sheet.getRange(2,1,numbers.length,1).setValues(numbers);
}

function integrityProjIdAndTitle(e) {
  const ui = SpreadsheetApp.getUi();
  const sheet = e.source.getActiveSheet();
  const sheet_works = e.source.getSheetByName('works');
  const sheet_projects = e.source.getSheetByName('projects');
  const sheet_works_lastRow = sheet_works.getLastRow();
  const sheet_projects_lastRow = sheet_projects.getLastRow();

  const editedColumn = e.range.getColumn();
  const editedHeader = sheet.getRange(1,editedColumn).getValue();

  //worksタブの列番号取得
  const projIdColumn_works = getColByHeaderName(sheet_works, '案件番号');
  const projTitleColumn_works = getColByHeaderName(sheet_works, '案件タイトル');

  //projectsタブの列番号取得
  const projIdColumn_projects = getColByHeaderName(sheet_projects, '案件番号');
  const projTitleColumn_projects = getColByHeaderName(sheet_projects, '案件タイトル');
  const projIds_projects = sheet_projects.getRange(2,projIdColumn_projects,sheet_projects_lastRow,1).getValues().flat();
  const projTitles_projects = sheet_projects.getRange(2,projTitleColumn_projects,sheet_projects_lastRow,1).getValues().flat();

  if (sheet.getName() === 'works') {
    if (editedHeader === '案件番号') {
      SpreadsheetApp.getUi().showModalDialog(startProcessingAnimation, "案件番号を変更しています");
      const projIdIndex = projIds_projects.indexOf(e.value);
      const projTitle = projTitles_projects[projIdIndex];
      sheet_works.getRange(e.range.getRow(), projTitleColumn_works).setValue(projTitle);
      SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, "処理が完了しました");
    }

    if (editedHeader === '案件タイトル') {
      SpreadsheetApp.getUi().showModalDialog(startProcessingAnimation, "案件タイトルを変更しています");
      const projId = sheet_works.getRange(e.range.getRow(), projIdColumn_works).getValue();

      let duplicationProjId;
      for (let i = 0; i < projTitles_projects.length; i++) {
        if (projTitles_projects[i] === e.value) {
          duplicationProjId = projIds_projects[i];
          break;
        }
      }

      if (duplicationProjId) {
        const confirmation = ui.alert(`「${e.value}」は既に使用されています`,`案件番号を ${duplicationProjId} に変更しますか？`,ui.ButtonSet.YES_NO);
        if (confirmation === ui.Button.YES) {
          sheet_works.getRange(e.range.getRow(),projIdColumn_works).setValue(duplicationProjId);
        } else {
          e.range.setValue(e.oldValue);
        }
      } else {
        const projSheet_inputRow = projIds_projects.indexOf(projId) + 2;
        sheet_projects.getRange(projSheet_inputRow,projTitleColumn_projects).setValue(e.value);

        const projIdRange_works = sheet_works.getRange(1,projIdColumn_works,sheet_works_lastRow,1);
        const workSheet_inputRange = getValueRanges(projId,projIdRange_works);

        if (workSheet_inputRange) {
          for (const range of workSheet_inputRange) {
            sheet_works.getRange(range.getRow(),projTitleColumn_works).setValue(e.value);
          }
        }
      }

      SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, "処理が完了しました");
    }
  }

  if (sheet.getName() === 'projects') {
    if (editedHeader === '案件タイトル') {
      SpreadsheetApp.getUi().showModalDialog(startProcessingAnimation, "案件タイトルを変更しています");
      const projId = sheet_projects.getRange(e.range.getRow(), projIdColumn_projects).getValue();
      
      let duplicationProjId = [];
      for (let i = 0; i < projTitles_projects.length; i++) {
        if (projTitles_projects[i] === e.value) duplicationProjId.push(projIds_projects[i]);
      }

      if (duplicationProjId.length > 1) {
        ui.alert(`「${e.value}」は既に使用されています`,`案件番号：${duplicationProjId.join(',')}と重複しているため変更できません`,ui.ButtonSet.OK);
        e.range.setValue(e.oldValue);
      } else {
        const projIdRange_works = sheet_works.getRange(2,projIdColumn_works,sheet_works_lastRow-1,1);
        const workSheet_inputRange = getValueRanges(projId,projIdRange_works);

        for (const range of workSheet_inputRange) {
          const row = range.getRow();
          sheet_works.getRange(row, projTitleColumn_works).setValue(e.value);
          syncSheet_resourceToWork(sheet_works,row);
        }
      }
      SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, "処理が完了しました");
    }
  }
}