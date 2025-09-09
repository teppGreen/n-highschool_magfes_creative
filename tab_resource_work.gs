const headerNames_work = { //リソース管理シート workタブの各列のヘッダー名を指定
  manageId: '管理番号',
  projId: '案件番号',
  workId: '制作番号',

  genre: 'ジャンル',
  projTitle: '案件タイトル',
  workTitle: '制作タイトル',
  status: 'ステータス',

  datetime: {
    request: '依頼受付開始日時',
    hearing: '初回ヒアリング開始日時',
    creating: '制作開始日時',
    refining: 'ブラッシュアップ開始日時',
    approval: '班長承認開始日時',
    expected: '納品期限日時',
    delivery: '納品日時'
  },

  client: {
    email: '依頼者メアド',
    department: '依頼班',
    nickname: '依頼者'
  },

  url: {
    workSheet: '制作シート',
    workFolder: '制作フォルダ',
    deliveryFolder: '納品フォルダ',
    canvaFolder: 'Canva フォルダ',
    slackChannel: 'Slack チャンネル',
    slackThread: 'Slack スレッド'
  },

  joinedMembers: '担当者',
  joinedMembersCount: '担当者数',

  review: {
    usedApp: '制作アプリ',
    deliverablesCount: '成果物数',
    willMakeNextYear: '来年も作るべきか'
  }
}

function changeFileName_work(workInfo) {
  const manageId = workInfo.manageId.split('-');

  if (workInfo.url.workSheet) {
    const fileId = extractFileId(workInfo.url.workSheet);
    DriveApp.getFileById(fileId).setName(`${CONFIG.FOLDER_PREFIX.WORKSHEET}${manageId[0]}-${manageId[1]}_${workInfo.projTitle}_${workInfo.workTitle}`);
  }

  if (workInfo.url.workFolder) {
    const fileId = extractFileId(workInfo.url.workFolder);
    DriveApp.getFileById(fileId).setName(`${manageId[0]}-${manageId[1]}_${workInfo.projTitle}_${workInfo.workTitle}`);
  }
}

function inputStatusChangedDatetime_resource(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'works' || !e.value) return;

  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();
  if (sheet.getRange(1,editedCol).getValue() !== 'ステータス') return;
  
  let newStatus, datetimeKey, inputValue;
  if (e.value == CONFIG.STATUS.CANCELLED) {
    newStatus = CONFIG.STATUS.STEP6;
  } else {
    newStatus = e.value;
  }
  
  if (newStatus === CONFIG.STATUS.STEP6) { 
    datetimeKey = newStatus + '日時';
  } else {
    datetimeKey = newStatus + '開始日時'
    inputValue = CONFIG.TASK_STATUS.IN_PROGRESS;
  }

  const datetimeRange = sheet.getRange(editedRow,getColByHeaderName(sheet,datetimeKey));
  const datetime = datetimeRange.getValue();
  const now = new Date();

  if (!datetime) {
    datetimeRange.setValue(now);
  }
}

function filterJoinedMembers() {
  SpreadsheetApp.getUi().showModalDialog(startProcessingAnimation, "処理中")

  // 1. アクティブなスプレッドシートとシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_active = ss.getActiveSheet();
  const ssId = ss.getId();
  const sheetId = sheet_active.getSheetId();
  const ui = SpreadsheetApp.getUi();

  if (sheet_active.getName() !== 'projects' && sheet_active.getName() !== 'works') return;

  // フィルタービューの名前を作成
  const sheet_members = ss.getSheetByName('members');
  const user_email = Session.getActiveUser().getEmail();

  console.log(`user_email: ${user_email}`);

  const membersNameColIndex = 5;
  const membersEmailColIndex = 8;
  const memberRow = getRowBySingleCol(sheet_members, membersEmailColIndex, user_email);

  if (!memberRow) {
    SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, "エラー")
    ui.alert('フィルタービューを作成できません','membersタブに名前とメールアドレスを正しく登録してください。',ui.ButtonSet.OK);
    sheet_members.getRange('E2').activateAsCurrentCell();
    return;
  }

  const user_name = sheet_members.getRange(memberRow, membersNameColIndex).getValue();
  const datetime = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd(Z) HH:mm:ss');
  const filterViewName = `担当者_${user_name}`

  // フィルターする対象の列番号を指定
  const joinedMembersColIndex = getColByHeaderName(sheet_active, '担当者');

  // データ範囲を取得（ヘッダー行を除く）
  // 最終行が1行以下（データなし）の場合は処理を終了
  const lastRow = sheet_active.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('フィルター対象のデータがありません。');
    return;
  }

  const dataRange = sheet_active.getRange(2, joinedMembersColIndex, lastRow - 1, 1);

  // --- ここからSpreadsheet APIを使用した処理 ---

  // 既存の同名フィルタービューがあれば削除するためのリクエストを作成
  const requests = [];
  const sheetInfo = Sheets.Spreadsheets.get(ssId);
  const targetSheetInfo = sheetInfo.sheets.find(s => s.properties.sheetId === sheetId);

  if (targetSheetInfo && targetSheetInfo.filterViews) {
    const existingViews = targetSheetInfo.filterViews;

    for (let i = 0; i < existingViews.length; i++) {
      const existingViewName = existingViews[i].title;

      if (typeof existingViewName === 'string') {
        const match = existingViewName.match(/^(.*) \(.*\)$/);
        if (match) {
          if (match[1] !== filterViewName) continue;
        } else if(existingViewName !== filterViewName) {
          continue;
        }
        requests.push({
          deleteFilterView: {
            filterId: existingViews[i].filterViewId
          }
        });
      }
    }
  }

  // 非表示にする値を格納するリストを初期化
  const valuesToHide = [];
  const allCellValues = dataRange.getValues();

  // データ範囲のすべてのセルを一つずつチェック
  for (const row of allCellValues) {
    const cellValue = row[0];

    // セルが空白（''）の場合、無条件で非表示リストに追加する
    // これにより (Blank) がフィルターから除外される
    if (cellValue === '') {
      valuesToHide.push(cellValue);
      continue; // 次のセルの処理へ
    }

    if (typeof cellValue !== 'string' || cellValue === '') {
      continue; // 空のセルや文字列でない場合はスキップ
    }

    // セル内の名前をカンマで分割し、前後の空白を削除して配列にする
    const names = cellValue.split(',').map(name => name.trim());

    // 作成した名前の配列にGASの実行者名が "含まれていない" 場合
    if (!names.includes(user_name)) {
      // そのセルの値を「非表示リスト」に追加する
      valuesToHide.push(cellValue);
    }
  }

  // 万が一同じ値が複数リストに入った場合を想定し、重複を削除
  const uniqueMembersToHide = [...new Set(valuesToHide)];
  const createdDatetime = Utilities.formatDate(new Date(), 'JST', 'MM/dd HH:mm:ss');

  // 6. 新しいフィルタービューを追加するリクエストを作成
  requests.push({
    addFilterView: {
      filter: {
        title: `${filterViewName} (${createdDatetime})`,
        range: {
          sheetId: sheetId,
          startRowIndex: 0, // 範囲はヘッダーを含むシート全体
          endRowIndex: sheet_active.getMaxRows(),
          startColumnIndex: 0,
          endColumnIndex: sheet_active.getMaxColumns()
        },
        criteria: {
          // APIの列インデックスは0から始まるため -1 する
          [joinedMembersColIndex - 1]: {
            // 新しいロジックで作成した「非表示リスト」を使用
            hiddenValues: uniqueMembersToHide
          }
        }
      }
    }
  });

  // リクエストをまとめて実行し、レスポンスを受け取る
  const response = Sheets.Spreadsheets.batchUpdate({ requests: requests }, ssId);

  // レスポンスから作成されたフィルタービューのIDを取得
  const addFilterViewResponse = response.replies.find(reply => 'addFilterView' in reply);
  if (!addFilterViewResponse) {
    throw new Error('フィルタービューの作成に失敗しました。');
  }
  const filterViewId = addFilterViewResponse.addFilterView.filter.filterViewId;

  // フィルタービューを適用し、UIを最小限にするURLを生成
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/edit?rm=minimal#gid=${sheetId}&fvid=${filterViewId}`;

  // HTMLでダイアログを作成し、iframeでフィルタービュー適用済みのシートを表示
  const htmlOutput = HtmlService.createHtmlOutput(
    `
    <style>
      body, html { margin: 0; padding: 0; height: 100%; overflow: hidden; }
      iframe { width: 100%; height: 100%; border: none; }
    </style>
    <iframe src="${url}"></iframe>
    `
  )
  .setWidth(10000)
  .setHeight(10000);

  SpreadsheetApp.getUi().showModalDialog(stopProcessingAnimation, "処理が完了しました")
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `フィルタービュー適用中: ${filterViewName}`);
}