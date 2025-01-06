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
  let fileName,fileId

  if (workInfo.url.workSheet) {
    fileName = `【制作管理】${workInfo.projId}${String(workInfo.workId).padStart(3,'0')}_${workInfo.projTitle}_${workInfo.workTitle}`;
    fileId = extractFileId(workInfo.url.workSheet);
    DriveApp.getFileById(fileId).setName(fileName);
  }

  if (workInfo.url.workFolder) {
    fileName = `${workInfo.projId}${String(workInfo.workId).padStart(3,'0')}_${workInfo.projTitle}_${workInfo.workTitle}`;
    fileId = extractFileId(workInfo.url.workFolder);
    DriveApp.getFileById(fileId).setName(fileName);
  }
}