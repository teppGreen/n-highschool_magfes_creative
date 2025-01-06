function inputLastModifiedDate(e) { //制作シートで動作。各項目の最終更新日時を入力
  if (e.source.getActiveSheet().getName() !== 'main' || e.range.getRow() < 2) return;

  const targetRange = e.range.offset(-1,2);
  const targetDataValidation = targetRange.getDataValidation();

  if (targetDataValidation) {
    const targetCreteriaType = targetDataValidation.getCriteriaType();
    const creteriaType = SpreadsheetApp.DataValidationCriteria.DATE_IS_VALID_DATE;
    const now = new Date();
  
    if (targetCreteriaType === creteriaType) targetRange.setValue(now);
  }
}