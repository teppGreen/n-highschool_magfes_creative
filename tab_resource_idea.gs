function displayRequestForm_normal() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('parameters');
  const key = getValueRanges('requestForm.id',sheet)[0];
  const value = key.offset(0,1).getValue();
  const form = FormApp.openById(value);
  const url = form.getPublishedUrl();

  displayRequestForm(`${url}?`,'制作依頼フォーム')
}

function processSystemCommand(requestInfo) {
  // ideaタブから提出された場合
  const regexp_ideas = /^idea[0-9]+$/;
  const result_ideas = requestInfo.systemCommand.find(item => regexp_ideas.test(item));

  if (result_ideas) {
    const ideaId = Number(result_ideas.match(/^idea([0-9]+)$/)[1]);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ideas');
    const lastRow = sheet.getLastRow();
    const ideaIdColumn = getColByHeaderName(sheet,'番号');
    const checkboxColumn = getColByHeaderName(sheet,'提出');
    const ideaIds = sheet.getRange(2,ideaIdColumn,lastRow-1,1).getValues().flat();

    const row = ideaIds.indexOf(ideaId) + 2; 
    sheet.getRange(row,checkboxColumn).setValue(true);
  }

  // 制作依頼シートから提出された場合
  const regexp_draft = /^reqFromDraftSheet\./;
  const result_draft = requestInfo.systemCommand.find(item => regexp_draft.test(item));

  if (result_draft) {
    let draftSheetId = result_draft.match(/^reqFromDraftSheet\.(.+)$/)[1];
    draftSheetId = draftSheetId.split('.').map(item => item.trim());
    
    const sheets = SpreadsheetApp.openById(draftSheetId[0]).getSheets();

    for (const sheet of sheets) {
      if (sheet.getSheetId() == draftSheetId[1]) {
        sheet.getRange('I1').setValue('提出済');
      }
    }
  }
}