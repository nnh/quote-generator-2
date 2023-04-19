/**
 * Re-set the formulas on the copied Items sheet.
 * @param {Object} ss The sheet object.
 * @param {Object} inputData Map object of the information entered from the form.
 * @return {Object} request object.
 */
function setItemsSheet_(ss, inputData){
  const itemsSheet = ss.sheets.filter(x => x.properties.title === itemsInfo.get('sheetName'));
  if (itemsSheet.length !== 1){
    return;
  }
  const items = itemsSheet[0];
  const itemsColIdxList = itemsInfo.get('colItemNameAndIdx');
  const formulaColsIdx = [
    itemsColIdxList.get('secondaryItem'),
    itemsColIdxList.get('price'),
    itemsColIdxList.get('baseUnitPrice'),
  ];
  const secondaryItemColName = colNamesConstant[getNumber_(itemsColIdxList.get('secondaryItem'))];
  // Corresponding to items for which unit prices need to be set (e.g., insurance premiums).
  const secondaryItemValue = spreadSheetBatchUpdate.rangeGetValue(ss.spreadsheetId, `${itemsInfo.get('sheetName')}!${secondaryItemColName}1:${secondaryItemColName}${items.properties.gridProperties.rowCount}`);
  if (secondaryItemValue.length !== 1){
    return;
  }
  const secondaryItem = secondaryItemValue[0].values;
  const setPriceTarget = new Map(
    [
      ['保険料', '保険料'],
      ['試験開始準備費用', '試験開始準備費用'],
      ['症例最終報告書提出毎の支払', '症例報告'],
      ['症例登録毎の支払', '症例登録'],
    ]
  );
  const setPriceTargetNameAndIdxMap = new Map();
  setPriceTarget.forEach((itemName, inputTitleName) => {
    const idxArray = secondaryItem.map((x, idx) => x[0] === itemName ? idx: null).filter(x => x);
    if (idxArray.length === 1 && Number.isSafeInteger(inputData.get(inputTitleName))){
      setPriceTargetNameAndIdxMap.set(inputTitleName, idxArray[0]);
    } 
  });
  const setFormulaRequest = formulaColsIdx.map(formulaColIdx => {
    const colString = commonGas.getColumnStringByIndex(formulaColIdx);
    const setItems = spreadSheetBatchUpdate.rangeGetValue(ss.spreadsheetId, `${items.properties.title}!${colString}:${colString}`, 'FORMULA')[0].values.map(x => x.length === 1 ? x : ['']);
    return spreadSheetBatchUpdate.getRangeSetValueRequest(items.properties.sheetId, 
                                                          0, 
                                                          formulaColIdx, 
                                                          setItems);
  });
  let setPriceRequest = [];
  setPriceTargetNameAndIdxMap.forEach((targetRowIdx, itemName) => 
    setPriceRequest.push(
      spreadSheetBatchUpdate.getRangeSetValueRequest(items.properties.sheetId, 
                                                     targetRowIdx, 
                                                     itemsColIdxList.get('price'), 
                                                     [[inputData.get(itemName)]])
    )
  );
  let requests = [...setFormulaRequest];
  if (setPriceRequest.length > 0){
    requests.push(...setPriceRequest);
  }
  return requests;                                                                           
}
