function doGet(e) {
  const param = e.parameter;
  const page = param.page ? param.page : 'index';
  let htmlOutput = HtmlService.createTemplateFromFile(page).evaluate();
  if (page === 'index'){
    htmlOutput
      .setTitle('title_index')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    return htmlOutput;
  } 
  if (page === 'quote'){
    htmlOutput
      .setTitle('title_quote')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    return htmlOutput;
  }
}
function getAppUrl_() {
  return ScriptApp.getService().getUrl();
}
class SetInputForm{
  constructor(arg){
    this.arg = arg;
  }
  setRadio(){
    
  }

}
function getItemsList(){
  let targetIndex = {};
  targetIndex.itemValueHeading = 0;
  targetIndex.itemValueItem = 1;
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('inputSsId'));
  const itemSheet = ss.getSheetByName('Items');
  const itemValues = itemSheet.getDataRange().getValues();
  let itemHeadingAndPrice = [];
  const _dummy = itemValues.reduce((x, currentValue) => {
    let res = currentValue;
    if (res[targetIndex.itemValueHeading] === ''){
      res[targetIndex.itemValueHeading] = x[targetIndex.itemValueHeading];
    }
    itemHeadingAndPrice.push(res);
    return res;
  });
  itemHeadingAndPrice = itemHeadingAndPrice.filter(x => x[targetIndex.itemValueItem] !== '').map(x => x.splice(0, 4));
  return itemHeadingAndPrice;
}

