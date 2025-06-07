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

function changeFileName(workInfo) {
  const fileName = `${workInfo.projId}${String(workInfo.workId).padStart(3,'0')}_${workInfo.projTitle}_${workInfo.workTitle}`;

  if (workInfo.url.workSheet) {
    const fileId = extractFileId(workInfo.url.workSheet);
    DriveApp.getFileById(fileId).setName('【制作管理】' + fileName);
  }

  if (workInfo.url.workFolder) {
    const fileId = extractFileId(workInfo.url.workFolder);
    DriveApp.getFileById(fileId).setName(fileName);
  }
}

function inputStatusChangedDatetime_resource(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'works' || !e.value) return;

  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();
  if (sheet.getRange(1,editedCol).getValue() !== 'ステータス') return;
  
  let newStatus, datetimeKey, inputValue;
  if (e.value == '依頼取消') {
    newStatus = '納品';
  } else {
    newStatus = e.value;
  }
  
  if (newStatus === '納品') { 
    datetimeKey = newStatus + '日時';
  } else {
    datetimeKey = newStatus + '開始日時'
    inputValue = '実行中'
  }

  const datetimeRange = sheet.getRange(editedRow,getColByHeaderName(sheet,datetimeKey));
  const datetime = datetimeRange.getValue();
  const now = new Date();

  if (!datetime) {
    datetimeRange.setValue(now);
  }
}