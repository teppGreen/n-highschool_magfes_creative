function getSpreadsheetId() {
  const id = SpreadsheetApp.getActiveSpreadsheet().getId();
  console.log(`spreadsheetId: ${id}`);
  
  return id;
}

function getBindFormId() {
  const formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  const formId = FormApp.openByUrl(formUrl).getId();
  console.log(formUrl);

  return formId;
}