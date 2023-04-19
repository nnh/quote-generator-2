function myFunction(){
  const test = forTest_(4);
  createSpreadsheet(test);
}
function forTest_(targetRowIndex = 1){
  const quotationRequest = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ssIdForTest')).getSheets()[0].getDataRange().getValues();
  const temp = quotationRequest.filter((_, idx) => idx === 0 || idx === targetRowIndex);  
  const tempItems = temp[0].map((x, idx) => [x, temp[1][idx]]);
  const items = new Map(tempItems);
  return items;
}
