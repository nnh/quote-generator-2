function execTestBySs(){
  const test = forTest_(2);
  createSpreadsheet(test);
}
function forTest_(targetRowIndex = 1){
  const quotationRequest = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ssIdForTest')).getSheets()[0].getDataRange().getValues();
  const temp = quotationRequest.filter((_, idx) => idx === 0 || idx === targetRowIndex);  
  const tempItems = temp[0].map((x, idx) => {
    const tempValue = temp[1][idx];
    const value = x === '中間解析の頻度' && /,/.test(tempValue) ? tempValue.split(', ') : tempValue;
    return [x, value];
  });
  const items = new Map(tempItems);
  return items;
}
function execTestByForm(){
  const test = forTestByForm_(0);
  createSpreadsheet(test);
}
function forTestByForm_(targetIdx = 0){
  const form = FormApp.openById(PropertiesService.getScriptProperties().getProperty('formId'));
  const formResponses = form.getResponses();
  const formResponse = formResponses[targetIdx];
  return getItemsFromFormRequests(formResponse);
}
