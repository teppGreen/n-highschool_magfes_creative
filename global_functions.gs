// ローディングアニメーション
const startProcessingAnimation = HtmlService.createHtmlOutputFromFile('processingAnimation').setWidth(400).setHeight(300);
const stopProcessingAnimation = HtmlService.createHtmlOutput('<script>google.script.host.close()</script>');
// SpreadsheetApp.getUi().showModalDialog(startProcessingAnimation, "処理中"); で呼び出して使用します。

function getValueRanges(targetValue, searchRange) {
  try{
    if (!targetValue || !searchRange) return;
    const targetRanges = searchRange.createTextFinder(targetValue).matchEntireCell(true).findAll().map(range => range);
    return targetRanges;
  } catch {
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

function deleteDrawings() {
  const drawings = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDrawings();
  for (let drawing of drawings) {
    if (drawing.getOnAction() === 'deleteDrawings') drawing.remove();
  }
}

function displayRequestForm(url,title) {
  const html = `<iframe src="${url}&embedded=true" width="640" height="5000" frameborder="0" marginheight="0" marginwidth="0">Loading…</iframe>`;
  
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(720).setHeight(3000);
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(htmlOutput,title);
}