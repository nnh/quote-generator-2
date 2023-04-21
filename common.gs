const colNamesConstant = [null, 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ'];
function getCommonInfo_(){
  const info = new Map();
  info.set('facilitiesItemName', '実施施設数');
  info.set('crfItemName', 'CRF項目数');
  info.set('casesItemName', '目標症例数');
  info.set('trialTypeItemName', '試験種別');
  info.set('sourceOfFundsTextItemName', '原資');
  info.set('commercialCompany', '営利企業原資（製薬企業等）');
  const trialType = new Map([
    ['investigatorInitiatedTrial', '医師主導治験'],
    ['specifiedClinicalTrial', '特定臨床研究'],
  ]);
  info.set('trialType', trialType);
  info.set('totalSheetName', 'Total');
  info.set('total2SheetName', 'Total2');
  info.set('templateSheetName', 'Setup');
  return info;
}
function getTrialsInfo_(){
  const info = new Map();
  info.set('trialStartText', '症例登録開始');
  info.set('trialEndText', '試験終了');
  info.set('yearText', '年');
  info.set('monthText', '月');
  const trialTermsAddress = new Map([
    ['rowIdx', 39],
    ['startColIdx', 3],
    ['endColIdx', 4],
  ]);
  info.set('trialTermsAddress', trialTermsAddress);
  info.set('trialStart', null);
  info.set('trialEnd', null);
  info.set('setupStart', null);
  info.set('closingEnd', null);
  info.set('setupTerm', null);
  info.set('closingTerm', null);
  info.set('cases', null);
  info.set('facilities', null);
  info.set('registrationStartYear', null);
  info.set('registrationEndYear', null)
  info.set('registrationYearsCount', null);
  info.set('commentStartRowIdx', 11);
  info.set('commentEndRowIdx', 25);
  info.set('discountRateAddress', 'B47');
  info.set('taxAddress', 'B45');
  info.set('sheetName', 'Trial');
  return info;
}
function getItemsInfo_(){
  const info = new Map();
  const colItemNameAndIdx = new Map([
    ['primaryItem', 0],
    ['secondaryItem', 1],
    ['price', 2],
    ['unit', 3],
    ['baseUnitPrice', 17],
  ]);
  info.set('colItemNameAndIdx', colItemNameAndIdx);
  info.set('bodyStartRowIdx', 2);
  info.set('sheetName', 'Items');
  return info;
}
function getTemplateInfo_(){
  const info = new Map();
  const colItemNameAndIdx = new Map([
    ['primaryItem', 1],
    ['secondaryItem', 2],
    ['price', 3],
    ['x', 4],
    ['count', 5],  
    ['amount', 7],
    ['sum', 8],
    ['rightBorderEnd', 9],
    ['filter', 11],
  ]);
  info.set('sheetName', 'templateByYear');
  info.set('headStartRowIdx', 1);
  info.set('bodyStartRowIdx', 4);
  info.set('startColIdx', 1);
  info.set('colItemNameAndIdx', colItemNameAndIdx);
  return info;
}
function getRowAndColLength_(array){
  return [array.length, array[0].length];
}
function getNumber_(idx){
  return idx + 1;
}
/**
 * Set conditional formatting.
 * @param {Object} targetRange Range object to set the conditional formatting. 
 * @return {Object} Request body.
 */
function editConditionalFormatRuleRequest(targetRange){
  const rgbColor = new spreadSheetBatchUpdate.createRgbColor();
  return [
    spreadSheetBatchUpdate.setAddConditionalFormatRuleNumberEq(targetRange, '0', rgbColor.white(), rgbColor.gray()),
    spreadSheetBatchUpdate.setAddConditionalFormatRuleNumberEq(targetRange, '0', rgbColor.white(), rgbColor.white(), 'NUMBER_NOT_EQ'),
  ];
}
/**
 * Set number formatting.
 * @param {Object} sheet The sheet object.
 * @param {number} startRow
 * @param {number} startCol
 * @param {number} lastRow
 * @param {number} lastCol
 * @return {Object} Request body.
 */
function setNumberFormat_(sheet, startRow=0, startCol=templateInfo.get('colItemNameAndIdx').get('amount'), lastRow, lastCol=templateInfo.get('colItemNameAndIdx').get('sum')){
  const request = [spreadSheetBatchUpdate.getRangeSetFormatRequest(sheet.properties.sheetId, 
                                                                   startRow, 
                                                                   startCol,
                                                                   lastRow, 
                                                                   lastCol, 
                                                                   spreadSheetBatchUpdate.editNumberFormat('NUMBER', '#,###'), 
                                                                   'userEnteredFormat.numberFormat'),
                  ];
  return request;
}