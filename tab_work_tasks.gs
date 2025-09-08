function inputStatusChangedDatetime_work(e) {
  const workSheet_active = e.source.getActiveSheet();
  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();
  const editedHeader = workSheet_active.getRange(1, editedCol).getValue();

  if (workSheet_active.getName() !== 'tasks' || editedHeader !== 'ステータス') return;
  
  const statusList = Object.values(CONFIG.STATUS).filter(status => status !== CONFIG.STATUS.CANCELLED);
  const editedTitleRange = workSheet_active.getRange(editedRow,getColByHeaderName(workSheet_active,'タイトル'));
  const editedTitle = editedTitleRange.getValue();
  const startDatetimeRange = workSheet_active.getRange(editedRow,getColByHeaderName(workSheet_active,'開始日時'));
  const startDatetime = startDatetimeRange.getValue();
  const endDatetimeRange = workSheet_active.getRange(editedRow,getColByHeaderName(workSheet_active,'終了日時'));
  const endDatetime = endDatetimeRange.getValue();
  const now = new Date();
  
  if (!statusList.includes(editedTitle)) return;

  if (e.value === CONFIG.TASK_STATUS.IN_PROGRESS && editedTitle === CONFIG.STATUS.STEP6) {
    e.range.setValue(CONFIG.TASK_STATUS.DONE);
  }
  
  if (((e.value === CONFIG.TASK_STATUS.IN_PROGRESS) || (e.value === CONFIG.TASK_STATUS.DONE && editedTitle === CONFIG.STATUS.STEP6)) && startDatetime === '') {
    startDatetimeRange.setValue(now);
  }

  if (e.value === CONFIG.TASK_STATUS.DONE && endDatetime === '') {
    endDatetimeRange.setValue(now);
  }
}