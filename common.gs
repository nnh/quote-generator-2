const colNamesConstant = [null, 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ'];
/**
 * registrationの年度で等分する必要のある項目リスト
 * 
 */
function getRegistrationDivisionInfo_(){
  const info = new Map();
  info.set('症例登録毎の支払', null);
  return info;
}
/**
 * 共通情報
 */
function getCommonInfo_(){
  const info = new Map();
  info.set('facilitiesItemName', '目標施設数');
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
  info.set('case', null);
  info.set('facilities', null);
  info.set('registrationStartYear', null);
  info.set('registrationEndYear', null)
  info.set('registrationYearsCount', null);
  info.set('discountRateAddress', 'B47');
  info.set('taxAddress', 'B45');
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
 * Project management is handled separately since the formula is different from other items.
 * @see library spreadSheetBatchUpdate
 */
class ProjectManagement{
  /**
   * @param {Object} ss The spreadsheet object.
   * @return none.
   */
  constructor(ss){
    this.ss = ss;
    this.itemName = 'プロジェクト管理';
    this.secondaryItemColNumber = getNumber_(templateInfo.get('colItemNameAndIdx').get('secondaryItem'));
    this.priceItemColNumber = getNumber_(templateInfo.get('colItemNameAndIdx').get('price'));
  }
  /**
   * Returns the column name of the count column on the template sheet as a string, e.g. 'F'.
   * @param none.
   * @return {string} the column name.
   */
  getCountColName(){
    return colNamesConstant[getNumber_(templateInfo.get('colItemNameAndIdx').get('count'))];
  }
  /**
   * Looks for "project management" in the list of subitem names and returns its index.
   * @param none.
   * @return {number} The row index.
   */
  getRowIdx(){
    const secondaryItems = spreadSheetBatchUpdate.rangeGetValue(this.ss.spreadsheetId, `${this.sheet.properties.title}!${colNamesConstant[this.secondaryItemColNumber]}1:${colNamesConstant[this.secondaryItemColNumber]}${this.sheet.properties.gridProperties.rowCount}`)[0].values;
    const projectManagementIdx = secondaryItems.map((x, idx) => x[0] === this.itemName ? idx : null).filter(x => x)[0];
    return projectManagementIdx;
  }
  /**
   * Looks for "project management" in the list of subitem names and returns its index.
   * @param none.
   * @return {number} The row number.
   */
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