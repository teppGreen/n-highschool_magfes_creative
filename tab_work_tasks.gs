function inputStatusChangedDatetime_work(e) {
  const workSheet_active = e.source.getActiveSheet();
  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();
  const editedHeader = workSheet_active.getRange(1, editedCol).getValue();

  if (workSheet_active.getName() !== 'tasks' || editedHeader !== 'ステータス') return;
  
  const statusList = ['依頼受付','初回ヒアリング','制作','ブラッシュアップ','班長承認','納品'];
  const editedTitleRange = workSheet_active.getRange(editedRow,getColByHeaderName(workSheet_active,'タイトル'));
  const editedTitle = editedTitleRange.getValue();
  const startDatetimeRange = workSheet_active.getRange(editedRow,getColByHeaderName(workSheet_active,'開始日時'));
  const startDatetime = startDatetimeRange.getValue();
  const endDatetimeRange = workSheet_active.getRange(editedRow,getColByHeaderName(workSheet_active,'終了日時'));
  const endDatetime = endDatetimeRange.getValue();
  const now = new Date();
  
  if (!statusList.includes(editedTitle)) return;

  if (e.value === '実行中' && editedTitle === '納品') {
    e.range.setValue('完了');
  }
  
  if (((e.value === '実行中') || (e.value === '完了' && editedTitle === '納品')) && startDatetime === '') {
    startDatetimeRange.setValue(now);
  }

  if (e.value === '完了' && endDatetime === '') {
    endDatetimeRange.setValue(now);
  }
}