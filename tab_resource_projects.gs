function generateProjectNumbers() { //AA-ZZの案件番号を作成
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.PROJECTS);

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
  const ss = e.source;
  const sheet = ss.getActiveSheet();
  const sheet_works = e.source.getSheetByName(CONFIG.SHEET_NAMES.WORKS);
  const sheet_projects = e.source.getSheetByName(CONFIG.SHEET_NAMES.PROJECTS);
  const sheet_works_lastRow = sheet_works.getLastRow();
  const sheet_projects_lastRow = sheet_projects.getLastRow();

  const editedColumn = e.range.getColumn();
  const editedHeader = sheet.getRange(1,editedColumn).getValue();

  //worksタブの列番号取得
  const projIdColumn_works = getColByHeaderName(sheet_works, CONFIG.HEADER_NAMES.PROJECT_ID);
  const projTitleColumn_works = getColByHeaderName(sheet_works, CONFIG.HEADER_NAMES.PROJECT_TITLE);

  //projectsタブの列番号取得
  const projIdColumn_projects = getColByHeaderName(sheet_projects, CONFIG.HEADER_NAMES);
  const projTitleColumn_projects = getColByHeaderName(sheet_projects, CONFIG.HEADER_NAMES.PROJECT_TITLE);
  const projIds_projects = sheet_projects.getRange(2,projIdColumn_projects,sheet_projects_lastRow,1).getValues().flat();
  const projTitles_projects = sheet_projects.getRange(2,projTitleColumn_projects,sheet_projects_lastRow,1).getValues().flat();

  if (sheet.getName() === CONFIG.SHEET_NAMES.WORKS) {
    if (editedHeader === CONFIG.HEADER_NAMES.PROJECT_ID) {
      if (e.value) {
        ss.toast('案件タイトルは自動で変更されます。そのままお待ちください。',`案件番号が${e.oldValue}→${e.value}に変更されました`,-1);
        const projIdIndex = projIds_projects.indexOf(e.value);
        const projTitle = projTitles_projects[projIdIndex];
        sheet_works.getRange(e.range.getRow(), projTitleColumn_works).setValue(projTitle);
        ss.toast('案件番号・案件タイトルの変更が完了しました');
      } else {
        ss.toast('処理を中断しました');
        ui.alert('案件番号は削除できません',
          `案件を変更したい場合は、${CONFIG.SHEET_NAMES.PROJECTS}タブに記載されている該当の案件番号に変更してください。\n` + 
          `新しい案件を作成したい場合は、未使用の案件番号を指定した上で、案件タイトルを設定してください。`,
          ui.ButtonSet.OK);
        e.range.setValue(e.oldValue);
      }
    }

    if (editedHeader === CONFIG.HEADER_NAMES.PROJECT_TITLE) {
      const projId = sheet_works.getRange(e.range.getRow(), projIdColumn_works).getValue();
      
      ss.toast('他の制作物の案件タイトルは自動で変更されます。そのままお待ちください。',`案件番号${projId}のタイトルが変更されました`,-1);

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

      ss.toast('案件番号・案件タイトルの変更が完了しました');
    }
  }

  if (sheet.getName() === CONFIG.SHEET_NAMES.PROJECTS) {
    if (editedHeader === CONFIG.HEADER_NAMES.PROJECT_TITLE) {
      const projId = sheet_projects.getRange(e.range.getRow(), projIdColumn_projects).getValue();
      ss.toast('worksタブの案件タイトルは自動で変更されます。そのままお待ちください。',`案件番号${projId}のタイトルが変更されました`,-1);
      
      let duplicationProjId = [];
      for (let i = 0; i < projTitles_projects.length; i++) {
        if (projTitles_projects[i] === e.value) duplicationProjId.push(projIds_projects[i]);
      }

      if (duplicationProjId.length > 1) {
        ss.toast('処理を中断しました');
        ui.alert(`「${e.value}」は既に使用されています`,`案件番号：${duplicationProjId.join(',')}と重複しているため変更できません。他のタイトルを設定してください。`,ui.ButtonSet.OK);
        e.range.setValue(e.oldValue);
      } else {
        const projIdRange_works = sheet_works.getRange(2,projIdColumn_works,sheet_works_lastRow-1,1);
        const workSheet_inputRange = getValueRanges(projId,projIdRange_works);

        for (const range of workSheet_inputRange) {
          const row = range.getRow();
          sheet_works.getRange(row, projTitleColumn_works).setValue(e.value);
          syncSheet_resourceToWork(sheet_works,row);
        }
        ss.toast('案件タイトルの変更が完了しました');
      }
    }
    if (editedHeader === CONFIG.HEADER_NAMES.PROJECT_ID) {
      ss.toast('処理を中断しました');
      ui.alert('案件番号は削除・変更できません',
        `制作物の案件番号を変更したい場合は、${CONFIG.SHEET_NAMES.WORKS}タブから変更してください。`,
        ui.ButtonSet.OK);
      e.range.setValue(e.oldValue);
    }
  }
}

function updateProjLinks(e) {
  const ss = e.source;
  const sheet = ss.getActiveSheet();
  const sheet_works = e.source.getSheetByName('works');
  const sheet_projects = e.source.getSheetByName('projects');
  const sheet_works_lastRow = sheet_works.getLastRow();

  const editedColumn = e.range.getColumn();
  const editedHeader = sheet.getRange(1,editedColumn).getValue();
  
  const projIdColumn_works = getColByHeaderName(sheet_works, CONFIG.HEADER_NAMES.PROJECT_ID);
  const projIdColumn_projects = getColByHeaderName(sheet_projects, CONFIG.HEADER_NAMES);

  if (sheet.getName() !== CONFIG.SHEET_NAMES.PROJECTS) return;

  if (editedHeader === CONFIG.HEADER_NAMES.PROJECT_FOLDER || editedHeader === CONFIG.HEADER_NAMES.PROJECT_DOCUMENT) {
    const projId = sheet_projects.getRange(e.range.getRow(), projIdColumn_projects).getValue();
    const projIdRange_works = sheet_works.getRange(2,projIdColumn_works,sheet_works_lastRow-1,1);
    const workSheet_inputRange = getValueRanges(projId,projIdRange_works);

    for (const range of workSheet_inputRange) {
      const row = range.getRow();
      syncSheet_resourceToWork(sheet_works,row);
    }
  }
}