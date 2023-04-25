function execCompare(){
  const compareFiles = driveCommon.getFilesArrayByFolderId(PropertiesService.getScriptProperties().getProperty('comparisonFolderId'));
  const testFileFiles = driveCommon.getFilesArrayByFolderId(PropertiesService.getScriptProperties().getProperty('testFolderId'));
  const fileNameHeader = 'Quote test-index';
  const startIndex = 1;
  const endIndex = 24;
  for (let i = startIndex; i <= endIndex; i++){
    const targetFileName = new RegExp(`${fileNameHeader}${i} `);
    const compare = compareFiles.filter(x => targetFileName.test(x));
    const test = testFileFiles.filter(x => targetFileName.test(x));
    if (compare.length === 1 && test.length === 1){
      const target1 = SpreadsheetApp.openById(compare[0].getId()).getSheetByName('Total').getRange('H96').getValue();
      const target2 = SpreadsheetApp.openById(test[0].getId()).getSheetByName('Total').getRange('H96').getValue();     
      console.log(`${targetFileName} compare : ${target1 === target2} : target1=${target1}, target2=${target2}`);
    } else {
      console.log(`※対象ファイルなし ： ${targetFileName}`);
    }
  }

}
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
  const startIndex = 0;
  const endIndex = 24;
  for (let i = startIndex; i <= endIndex; i++){
    try{
      const test = forTestByForm_(i);
      createSpreadsheet(test);
    } catch (error){
      console.log(`${error.name}:${error.message}`);
    }
  }
}
function forTestByForm_(targetIdx = 0){
  const form = FormApp.openById(PropertiesService.getScriptProperties().getProperty('formId'));
  const formResponses = form.getResponses();
  const formResponse = formResponses[targetIdx];
  return getItemsFromFormRequests(formResponse);
}
