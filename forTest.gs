function forTest(targetRowIndex = 1){
  const quotationRequest = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ssIdForTest')).getSheets()[0].getDataRange().getValues();
  const temp = quotationRequest.filter((_, idx) => idx === 0 || idx === targetRowIndex);  
  const tempItems = temp[0].map((x, idx) => [x, temp[1][idx]]);
  const items = new Map(tempItems);
  items.set('症例登録開始年', `${items.get('症例登録開始日').getFullYear()}年`);
  items.set('症例登録開始月', `${items.get('症例登録開始日').getMonth()}月`);
  items.set('試験終了年', `${items.get('試験終了日').getFullYear()}年`);
  items.set('試験終了月', `${items.get('試験終了日').getMonth()}月`);
  return items;
}
