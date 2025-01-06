function test1() {
    let num = 500
    let str = String(num).padStart(3,'0');
    console.log(str);
}

function test2() {
  const url1 = "https://drive.google.com/drive/folders/1YKB66PXJ53jeUADX1fz2VBusq6HskkN6?usp=drive_link";
  const url2 = "https://drive.google.com/drive/folders/1sCpilHP3NyTSaWz2jgxWgO8p1n_LeCJC";

  console.log(extractFileId(url1)); // "1YKB66PXJ53jeUADX1fz2VBusq6HskkN6"
  console.log(extractFileId(url2)); // "1sCpilHP3NyTSaWz2jgxWgO8p1n_LeCJC"
}

function test3() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('works');
  getRangesByHeaderNames(sheet, 5, headerNames_work)
}

function test4() {
  receptionRequest(3);
}

function test5() {
  const workSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('works');
  if (!workSheet) {
    console.error("Sheet 'works' not found");
    return;
  }

  const range = getRangesByHeaderNames(workSheet, 2, headerNames_work); // Deep copy
  console.log('Processed Ranges:', range);

  // Rangeオブジェクトにアクセスして処理
  function processObject(obj) {
    for (let key in obj) {
      if (typeof obj[key] === 'object' && obj[key] !== null && obj[key].getA1Notation) {
        console.log(`Key: ${key}, Range: ${obj[key].getA1Notation()}`);
      } else if (typeof obj[key] === 'object' && obj[key] !== null) {
        processObject(obj[key]); // ネストされたオブジェクトを再帰的に処理
      }
    }
  }

  processObject(range);
}

function test6() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('works');
  syncSheet_resourceToWork(sheet,2);
}

function test7() {
  const editedLabel = '制作アプリ';
  const reviewLabels = ['制作アプリ','成果物数','来年も作るべきか'];
  const labelIndex = reviewLabels.indexOf(editedLabel);
  console.log(labelIndex)
}

function test8() {
  const range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('works').getRange('D6');
  const richtext = SpreadsheetApp.newRichTextValue().setText('かわべ').setLinkUrl('https://google.com').build();
  range.setRichTextValue(richtext);
}

function test9() {
  const workSheet_main = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1h5f1ivbApp6bu5LnhDS53RgSTmo1UcfVDx_ty0YGOcM/edit?usp=drive_link');

  getValueRanges('内容',workSheet_main)[0].offset(1,0).setValue('内容です');
  getValueRanges('デザイン要項',workSheet_main)[0].offset(1,0).setValue('デザイン要項です');
  getValueRanges('入稿規定',workSheet_main)[0].offset(1,0).setValue('入稿規定です');
  getValueRanges('依頼備考',workSheet_main)[0].offset(1,0).setValue('依頼備考です');
}

function test10() {
  test11(1);
}

function test11(num1, num2) {
  console.log(num1);
  console.log(num2);
}

function test12() {
  const token = PropertiesService.getScriptProperties().getProperty("slackWorkflow_notifyRequest_WebReqestUrl");
  console.log(token);
  let requestInfo = 'A';
  console.log(requestInfo.hearingDatetime.length);

  for (i = 0; i < requestInfo.hearingDatetime.length; i++) {
    console.log('Hi');
  }
}