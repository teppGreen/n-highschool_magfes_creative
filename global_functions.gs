const CONFIG = {
  SHEET_NAMES: {
    WORKS: 'works',
    PARAMETERS: 'parameters',
    FORM: 'form',
    PROJECTS: 'projects',
    PERSONS: 'persons',
    MEMBERS: 'members',
    MAIN: 'main',
    TASKS: 'tasks',
  },
  HEADER_NAMES: {
    WORK_ID: '制作番号',
    WORK_TITLE: '制作タイトル',
    PROJECT_ID: '案件番号',
    PROJECT_TITLE: '案件タイトル',
    EMAIL: 'E-mail 1 - Value',
    SLACK_ID: 'Slack ID',
    STATUS: 'ステータス',
    CONTENT: '内容',
    DESIGN: 'デザイン要項',
    REGULATION: '入稿規定',
    NOTE: '依頼備考',
    HEARING: 'ヒアリング',
    REFERENCE: '参考物',
    OTHER_MEMO: 'その他メモ',
  },
  PARAM_KEYS: {
    CONTACT_SHEET_ID: 'contactSheet.id',
    SYSTEM_START_YEAR: 'system.startYear',
    WORK_FOLDER_URL: 'workInfo.url.workFolder',
    WORK_SHEET_URL: 'workInfo.url.workSheet',
    REGISTRATION_FORM_ID: 'registrationForm.id',
  },
  PROPERTIES: {
    SLACK_WORKFLOW_URL: 'slackWorkflow_notifyRequest_WebReqestUrl',
    SLACK_ADMIN_EMAIL: 'slackAdminEmail',
    RESOURCE_SHEET_ID: 'sheetId_resource',
  },
  INITIAL_STATUS: '依頼受付',
  FOLDER_PREFIX: {
    MATERIAL: '【素材】',
    DELIVERY: '【納品】',
    WORKSHEET: '【制作管理】',
  },
  UI: {
    RESOURCE_SHEET_LINK_TEXT: 'リソース管理シート',
  },
  SYSTEM_COMMANDS: {
    DONT_SEND_NOTIFICATION: 'dontSendNotification',
  },
};

// ローディングアニメーション
const startProcessingAnimation = HtmlService.createHtmlOutputFromFile('processingAnimation').setWidth(400).setHeight(300);
const stopProcessingAnimation = HtmlService.createHtmlOutput('<script>google.script.host.close()</script>');
// SpreadsheetApp.getUi().showModalDialog(startProcessingAnimation, "処理中"); 
// SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, "処理完了")

function previewProcessingAnimation() {
  SpreadsheetApp.getUi().showModalDialog(startProcessingAnimation, "処理中"); 
}

function getValueRanges(targetValue, searchRange) {
  try{
    if (!targetValue || !searchRange) return;
    const targetRanges = searchRange.createTextFinder(targetValue).matchEntireCell(true).findAll().map(range => range);
    console.log(`targetValue: ${targetValue}`);
    return targetRanges;
  } catch(error) {
    console.log(`targetValue: ${targetValue}\ntargetRanges: null\n${error.message}`);
    return null;
  }
}

function extractFileId(url) { //フォルダのURLにも対応しています
  try {
    console.log('Url: ' + url);
    if (/^[-\w]{25,}$/.test(url)) {
      return url; // ファイルIDだけが渡されるケース
    }

    const patterns = [
      /\/d\/([-\w]{25,})/, // "/d/" パターン
      /id=([-\w]{25,})/,   // "id=" パターン
      /\/open\?id=([-\w]{25,})/, // "/open?id=" パターン
      /\/file\/d\/([-\w]{25,})/, // "/file/d/" パターン
      /drive.google.com\/uc\?export=download&id=([-\w]{25,})/, // ダウンロードリンク
      /\/folders\/([-\w]{25,})/, // フォルダの場合のパターン
      /drive\/folders\/([-\w]{25,})/, // "/drive/folders/" パターン
      /spreadsheets\/d\/([-\w]{25,})/, // Googleスプレッドシートの場合
      /document\/d\/([-\w]{25,})/, // Googleドキュメントの場合
      /presentation\/d\/([-\w]{25,})/, // Googleスライドの場合
    ];

    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match) return match[1]; // マッチした場合、IDを返す
    }

    return null;
  } catch {
    return null;
  }
}

function generateHexRandom(digits) {
  const hexChars = '0123456789ABCDEF';
  let hexRandom;
  for (const i = 0; i < digits; i++) {
    hexRandom += hexChars.charAt(Math.floor(Math.random() * hexChars.length));
  }
  return hexRandom;
}

function getRangesByHeaderNames(sheet, row, headerNames) {
  // 再帰的にオブジェクトを処理して、新しいオブジェクトを生成
  function processObject(sheet, row, obj) {
    const result = {}; // 新しいオブジェクトを生成
    for (let key in obj) {
      if (typeof obj[key] === 'object' && obj[key] !== null) {
        result[key] = processObject(sheet, row, obj[key]); // ネストされたオブジェクトを再帰的に処理
      } else {
        const column = getColByHeaderName(sheet, obj[key]);
        if (column === 0) {
          result[key] = null; // 該当ヘッダーが見つからない場合はnull
        } else {
          const range = sheet.getRange(row, column);
          result[key] = range; // Rangeオブジェクトを新しいオブジェクトに設定
          console.log(`Header: ${obj[key]}, Range: ${range.getA1Notation()}`);
        }
      }
    }
    return result; // 作成した新しいオブジェクトを返す
  }

  return processObject(sheet, row, headerNames); // 新しいオブジェクトを返す
}

function getValuesByRanges(ranges) {
  // 再帰的にオブジェクトの各プロパティを処理する
  function processObject(obj) {
    let result = {};

    for (let key in obj) {
      if (typeof obj[key] === 'object' && obj[key]) {
        if (obj[key].getA1Notation) {
          result[key] = obj[key].getValue();
        } else {
          result[key] = processObject(obj[key]); // ネストされたオブジェクトを再帰的に処理
        }
      }
    }

    return result;
  }

  return processObject(ranges); //Valueを引数で受け取った連想配列の構造のまま返す
}

function getColByHeaderName(sheet, headerName) {
  const header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues().flat();
  const column = header.indexOf(headerName) + 1;
  return column;
}

function getRowBySingleCol(sheet, col, targetText) {
  const lastRow = sheet.getLastRow();
  const rangeValues = sheet.getRange(1,col,lastRow,1).getValues().flat();

  let row;
  for (let i = 0; rangeValues.length; i++) {
    if (rangeValues[i] === targetText) {
      row = i + 1;
      return row;
    }
  }

  return null;
}

function getRowBySingleCol(sheet, colIndex, targetText) {
  const lastRow = sheet.getLastRow();
  const rangeValues = sheet.getRange(1,colIndex,lastRow,1).getValues().flat();

  let row;
  for (let i = 0; i < rangeValues.length; i++) {
    if (rangeValues[i] === targetText) {
      row = i + 1;
      return row;
    }
  }

  return null;
}

function deleteDrawings() {
  const drawings = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDrawings();
  for (let drawing of drawings) {
    if (drawing.getOnAction() === 'deleteDrawings') drawing.remove();
  }
}

function displayRequestForm(title,url,params) {
  const param = `?embedded=true&${params ? params.join('&') : ''}`;
  const iframeSrc = url + param;
  console.log(`Display form: ${iframeSrc}`)

  const htmlOutput = HtmlService.createHtmlOutput(
    `
    <style>
      { box-sizing: border-box; }
      body, html { margin: 0; padding: 0; height: 100%; overflow: hidden; }
      iframe { width: 100%; height: 100%; border: none; }
    </style>
    <iframe src="${iframeSrc}"></iframe>
    `
  ).setWidth(720).setHeight(10000);
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(htmlOutput,title);
}
