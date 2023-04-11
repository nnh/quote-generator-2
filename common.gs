const colNamesConstant = [null, 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ'];
/**
 * 共通情報
 */
function getCommonInfo_(){
  const info = new Map();
  info.set('trialTypeItemName', '試験種別');
  info.set('sourceOfFundsTextItemName', '原資');
  info.set('commercialCompany', '営利企業原資（製薬企業等）');
  const trialType = new Map([
    ['investigatorInitiatedTrial', '医師主導治験'],
    ['specifiedClinicalTrial', '特定臨床研究'],
  ]);
  info.set('trialType', trialType);
  info.set('totalSheetName', 'total');
  info.set('total2SheetName', 'total2');
  return info;
}
/**
 * Trialsの列情報とか
 */
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
  info.set('discountRateAddress', 'B47');
  info.set('sheetName', 'Trial');
  return info;
}
/**
 * itemsの列情報とか
 */
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
/**
 * テンプレートファイルの列情報とか
 */
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
 * Set the filter.
 * @param {Object} sheet Sheet object.
 * @param {Object} targetRange Range object.
 * @return none.
 */
/*
function setFilter_(sheet, targetRange){
  const newRule1 = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([targetRange])
    .whenNumberGreaterThanOrEqualTo(1)
    .setBackground('#FFFFFF')
    .setFontColor('#FFFFFF')
    .build();
  const newRule2 = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([targetRange])
    .whenNumberEqualTo(0)
    .setBackground('#CCCCCC')
    .setFontColor('#FFFFFF')
    .build();
  const newRules = [newRule1, newRule2];
  sheet.setConditionalFormatRules(newRules);
  targetRange.createFilter();
}
*/
/**
 * Project management is handled separately since the formula is different from other items.
 */
class ProjectManagement{
  constructor(ss){
    this.ss = ss;
    this.itemName = 'プロジェクト管理';
    this.secondaryItemColNumber = getNumber_(templateInfo.get('colItemNameAndIdx').get('secondaryItem'));
    this.priceItemColNumber = getNumber_(templateInfo.get('colItemNameAndIdx').get('price'));
  }
  getCountColName(){
    return colNamesConstant[getNumber_(templateInfo.get('colItemNameAndIdx').get('count'))];
  }
  getRowIdx(){
    const secondaryItems = spreadSheetBatchUpdate.rangeGetValue(this.ss.spreadsheetId, `${this.sheet.properties.title}!${colNamesConstant[this.secondaryItemColNumber]}1:${colNamesConstant[this.secondaryItemColNumber]}${this.sheet.properties.gridProperties.rowCount}`)[0].values;
    const projectManagementIdx = secondaryItems.map((x, idx) => x[0] === this.itemName ? idx : null).filter(x => x)[0];
    return projectManagementIdx;
  }
  getRowNumber(){
    return getNumber_(this.getRowIdx());
  }
  /**
   * Edit the Template sheet.
   * @param {string} sheetId the sheet id.
   * @return {Object} Request body.
   */
  setTemplate_(sheetId){
    this.sheet = this.ss.sheets.filter(sheet => sheet.properties.sheetId === sheetId)[0];
    const rowNumber = this.getRowNumber();
    const targetStartRowNumber = getNumber_(templateInfo.get('bodyStartRowIdx'));
    const targetLastRowNumber = 63;
    const countColName = this.getCountColName();
    const amountColName = colNamesConstant[getNumber_(templateInfo.get('colItemNameAndIdx').get('amount'))];
    const formulaText = `=((sumif($${countColName}$${targetStartRowNumber}:$${countColName}$${rowNumber - 1}, ">0", $${amountColName}$${targetStartRowNumber}:$${amountColName}$${rowNumber - 1}) + sumif($${countColName}$${rowNumber + 1}:$${countColName}$${targetLastRowNumber}, ">0", $${amountColName}$${rowNumber + 1}:$${amountColName}$${targetLastRowNumber})) * 0.1) / ${countColName}${rowNumber}`; 
    const requests = [
      spreadSheetBatchUpdate.getRangeSetValueRequest(sheetId, this.getRowIdx(), templateInfo.get('colItemNameAndIdx').get('price'), [[formulaText]]),
    ];
    return requests;
  }
  /**
   * Edit the Total sheet.
   * @param {Object} sheet Sheet object.
   * @param <Array>{string} yearList Array of sheet names.
   * @return none.
   */
  setTotal_(sheet, yearList){
    this.sheet = sheet;
    const rowNumber = this.getRowNumber();
    const priceItemColName = getColumnString_(this.priceItemColNumber, sheet);
    const formulaText = yearList.map(year => `${year}!$${priceItemColName}$${rowNumber}`).join('+');
    sheet.getRange(`${priceItemColName}${rowNumber}`).setFormula(formulaText);
    const countColName = this.getCountColName();
    sheet.getRange(`${countColName}${rowNumber}`).setValue(1);
  }
}
/**
 * Set conditional formatting.
 * @param {Object} targetRange Range object to set the conditional formatting. 
 * @return {Object} Request body.
 * @see library spreadSheetBatchUpdate
 */
function editConditionalFormatRuleRequest(targetRange){
  const rgbColor = new spreadSheetBatchUpdate.createRgbColor();
  return [
    spreadSheetBatchUpdate.setAddConditionalFormatRuleNumberEq(targetRange, '0', rgbColor.white(), rgbColor.gray()),
    spreadSheetBatchUpdate.setAddConditionalFormatRuleNumberEq(targetRange, '0', rgbColor.white(), rgbColor.white(), 'NUMBER_NOT_EQ'),
  ];
}
