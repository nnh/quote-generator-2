function setYears(){
  const setYear = new SetYears();
  setYear.setCheckBox();
  setYear.setList();
}
class SetYears{
  constructor(){
    const thisYear = new Date().getFullYear();
    this.yearList = [...Array(10).keys()].map((_, idx) => `${thisYear + idx}年`); 
  }
  setList(){
    this.setChoiceItems(FormApp.ItemType.LIST, '年');
  }
  setCheckBox(){
    this.setChoiceItems(FormApp.ItemType.CHECKBOX, '中間解析の頻度');
  }
  setChoiceItems(itemType, targetTitle){
    const items = itemType === FormApp.ItemType.LIST 
      ? FormApp.getActiveForm().getItems(FormApp.ItemType.LIST).map(x => x.asListItem())
      : FormApp.getActiveForm().getItems(FormApp.ItemType.CHECKBOX).map(x => x.asCheckboxItem());
    const target = items.filter(x => new RegExp(targetTitle).test(x.getTitle()));
    target.forEach(list => list.setChoices(this.yearList.map(x => list.createChoice(x))));
  }
}
function createSs(){
  const user = Session.getActiveUser().getUserLoginId();
  const resList = FormApp.getActiveForm().getResponses(); 
  // Retrieve the most recent input information.
  const target = resList.filter(x => x.getRespondentEmail() === user).filter((_, idx, arr) => idx === arr.length - 1)[0];
  const items = quotegenerator2.getItemsFromFormRequests(target);
  const res = quotegenerator2.createSpreadsheet(items);
  if (typeof(res) !== 'string'){
    console.log(`${res.name}:${res.message}`);
    return;
  }
  GmailApp.sendEmail(user, 'test', `${res}の作成が完了しました。Googleドライブのマイドライブをご確認ください。`);
}
