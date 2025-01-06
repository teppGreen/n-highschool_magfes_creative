function openOrCloseFormResponse(e) { //毎分実行
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== 'members' && e.range.offset(-1,0).getValue() !== '新規依頼受付') return;

    const paramSheet = e.source.getSheetByName('parameters');
    const formId = getValueRanges('requestForm.id',paramSheet)[0].offset(0,1).getValue();

    let isAcceptingResponses;
    if (e.value === '可') isAcceptingResponses = true;
    if (e.value === '不可') isAcceptingResponses = false;

    const form = FormApp.openById(formId);
    form.setAcceptingResponses(isAcceptingResponses);

    e.source.toast(`制作依頼フォームの提出を「${e.value}」に変更しました`)
  } catch(error) {
    console.log(error.message)
    return;
  }
}