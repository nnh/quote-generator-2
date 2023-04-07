/**
 * Create a Total, Total2 sheet.
 */
class CreateTotalSheet{
  /**
   * @param {Object} ss Spreadsheet object.
   * @param {string[]} yearList Array of sheet names.
   * @param {Object} template Sheet object.
   */
  constructor(ss, targets){
    this.ss = ss;
    this.yearList = [];
    this.targetSheetList = [];
    targets.forEach((value, key) => {
      if (/\d{4}/.test(key)){
        this.yearList.push(key);
        this.targetSheetList.push(value);
      } else if (key === commonInfo.get('totalSheetName')){
        this.totalSheet = value;
      } else if (key === commonInfo.get('total2SheetName')){
        this.total2Sheet = value;
      }
    });
    this.template = ss.sheets.filter(x => x.properties.title === templateInfo.get('sheetName'))[0];    
    this.totalHeadText = '【見積明細：総期間】';
    this.countColName = colNamesConstant[getNumber_(templateInfo.get('colItemNameAndIdx').get('count'))];
  }
  exec(){
    let res = [];
    res.push(this.editTotal2Sheet_());
    res.push(this.editTotalSheet_());
    return [res];
  }
  editTotal2Sheet_(){
    const test = this.total2Sheet;
    // 1行目を削除する
  const delRowsRequest = [
    spreadSheetBatchUpdate.getdelRowColRequest(this.total2Sheet.sheetId, 'ROWS', 0, 1),
  ];
    // D列以降を一旦削除し、年数分+3列追加する
  const outputStartIdx = templateInfo.get('colItemNameAndIdx').get('price');
  const delColRequest = spreadSheetBatchUpdate.getdelRowColRequest(this.total2Sheet.sheetId, 'COLUMNS', outputStartIdx, this.total2Sheet.gridProperties.columnCount - outputStartIdx);
  const insertColRequest = spreadSheetBatchUpdate.getInsertRowColRequest(this.total2Sheet.sheetId, 'COLUMNS', outputStartIdx, this.yearList.length + 3);
  const insertRowRequest = spreadSheetBatchUpdate.getInsertRowColRequest(this.total2Sheet.sheetId, 'ROWS', 3, 4);
  // B1セルに「期間別見積」
  // B3セルに'【見積明細：総期間】'
  // ４行目に年数を出力
  const primaryItemColName = colNamesConstant[getNumber_(templateInfo.get('colItemNameAndIdx').get('primaryItem'))];
  const secondaryItemColName = colNamesConstant[getNumber_(templateInfo.get('colItemNameAndIdx').get('secondaryItem'))];
  const primaryItemRange = `${commonInfo.get('total2SheetName')}!${primaryItemColName}1:${secondaryItemColName}${this.total2Sheet.gridProperties.rowCount}`;
  const itemsValues = spreadSheetBatchUpdate.rangeGetValue(this.ss.spreadsheetId, primaryItemRange)[0].values;
  let primaryRowIndex = [];
  let secondaryRowIndex =[];
  itemsValues.forEach((value, idx) => {
    if (value.length === 0){
      return null;
    }
    if (value.length === 2){
      secondaryRowIndex.push(idx);
    }
    if (value[0] !== ''){
      primaryRowIndex.push(idx);
    }
  });
  //const primaryRowIndex = primaryItemValues.map((value, idx) => value.length > 0 ? idx : null).filter(x => x);
  // 5行目以降年数分、該当シートの該当セルを出す
  const sumRowIdx = itemsValues.map((value, idx) => value.length > 0 ? value[0] === '合計' ? idx : null : null).filter(x => x)[0];  
  let bodyRowsArray = [];
  for (let i = 5; i < sumRowIdx; i++){
    bodyRowsArray.push(i);
  }
  const outputStartColName = colNamesConstant[getNumber_(outputStartIdx)];
  const outputEndColName = colNamesConstant[getNumber_(outputStartIdx + this.yearList.length - 1)];
  const sumColName = colNamesConstant[getNumber_(outputStartIdx + this.yearList.length)];
  const setBodyFormulas = bodyRowsArray.map(row => {
    const yearsFormula = this.yearList.map(year => `=${String(year)}!$H$${row}`);
    const sumFormula = `=if(sum(${outputStartColName}${row}:${outputEndColName}${row})=0, "", sum(${outputStartColName}${row}:${outputEndColName}${row}))`;
    const filterFormula = `="dummy"`;
    return [...yearsFormula, sumFormula, '', filterFormula];
  });
  const headerValuesArray = new Array(this.yearList.length + outputStartIdx + 1).fill('');
  const headerValues = [
    ['', '【期間別見積】', ...headerValuesArray.slice(2)],
    headerValuesArray,
    ['', this.totalHeadText, ...headerValuesArray.slice(2)],
    ['', '', '',　...this.yearList.map(x => String(x)), '合計'],
  ];
  const setBodyRequest = [
    spreadSheetBatchUpdate.getRangeSetValueRequest(this.total2Sheet.sheetId,
                                                   4,
                                                   3,
                                                   setBodyFormulas), 
    spreadSheetBatchUpdate.getRangeSetValueRequest(this.total2Sheet.sheetId,
                                                   0,
                                                   0,
                                                   headerValues),
  ];
  return [...delRowsRequest, delColRequest, insertColRequest, insertRowRequest, ...setBodyRequest];
    // フィルタはTotalから値をとってくる

  }
  /**
   * Edit total sheet.
   * @param none.
   * @return {Object}
   */
  editTotalSheet_(){
    const formulas = [];
    for (let i = templateInfo.get('bodyStartRowIdx'); i < this.totalSheet.gridProperties.rowCount - templateInfo.get('bodyStartRowIdx'); i++){
      const formula = this.yearList.map(sheetName => `'${sheetName}'!${this.countColName}${i + 1}`).join(' + ');
      formulas.push([`=if(${formula} > 0, ${formula}, "")`]);
    }
    const setFormulasRequest = spreadSheetBatchUpdate.getRangeSetValueRequest(this.totalSheet.sheetId,
                                                                              templateInfo.get('bodyStartRowIdx'),
                                                                              templateInfo.get('colItemNameAndIdx').get('count'),
                                                                              formulas);
    const setHeadTextRequest = spreadSheetBatchUpdate.getRangeSetValueRequest(this.totalSheet.sheetId,
                                                                        templateInfo.get('headStartRowIdx'),
                                                                        templateInfo.get('startColIdx'),
                                                                        [[this.totalHeadText]]);
    return [setFormulasRequest, setHeadTextRequest];
    // Project management is calculated only once during the entire period.
    //new ProjectManagement().setTotal_(sheet, this.yearList);
  }
}
/**
 * @extends CreateTotalSheet
 */
class CreateTotal2Sheet extends CreateTotalSheet{
  constructor(ss, yearList, template){
    super(ss, yearList, template);
    this.headText = '【期間別見積】';
    this.sheetName = 'Total2';
  }
  /**
   * Set the filter.
   * @param none.
   * @return none.
   */
  setFilter_(){
    const properties = this.getProperties_();
    const sheet = properties.get('sheet');
    const filterColIdx = properties.get('filterColIdx');
    const filterColNumber = getNumber_(filterColIdx);
    const filterColName = getColumnString_(getNumber_(filterColIdx), sheet);
    const filterRange = sheet.getRange(1, filterColNumber, properties.get('endRow'), 1);
    const totalFilterArray = filterRange.getValues().map((_, idx) => [`=Total!$L${getNumber_(idx)}`]);
    filterRange.setValues(totalFilterArray);
    setFilter_(sheet, sheet.getRange(`${filterColName}${properties.get('yearRow')}:${filterColName}`));
  }
  /**
   * Property Settings
   * @param none.
   * @return {Object} Map object.
   */
  getProperties_(){
    const properties = new Map([
      ['sheet', this.ss.getSheetByName(this.sheetName)],
      ['startColIdx', templateInfo.get('colItemNameAndIdx').get('x')],
    ]);
    properties.set('startColName', getColumnString_(properties.get('startColIdx'), properties.get('sheet')));
    properties.set('endColIdx', properties.get('startColIdx') + this.yearList.length - 1);
    properties.set('endColName', getColumnString_(properties.get('endColIdx'), properties.get('sheet')));
    properties.set('bodyEndRow', templateInfo.get('totalRowNumber'));
    properties.set('endRow', properties.get('bodyEndRow') + 1);
    properties.set('yearSheetSumColName', getColumnString_(templateInfo.get('colItemNameAndIdx').get('sum'), properties.get('sheet')));
    properties.set('yearRow', templateInfo.get('bodyStartRowIdx') + 1);
    properties.set('bodyRowCount', properties.get('endRow') - properties.get('yearRow'));
    properties.set('filterColIdx', properties.get('endColIdx') + 2);
    return properties;
  }
  /**
   * Edit the Total2 sheet.
   * @param none.
   * @return none.
   */
  editSheet_(){
    const properties = this.getProperties_();
    const sheet = properties.get('sheet');
    sheet.deleteColumns(properties.get('startColIdx'), sheet.getLastColumn());
    this.setFilter_();
    const formulas = [...new Array(properties.get('endRow'))].map((_, idx) => {
      const row = getNumber_(idx);
      // 割引後合計は未対応
      const yearsFormula = this.yearList.map((year) => row < properties.get('endRow') ? `=IF(${year}!$${properties.get('yearSheetSumColName')}${row}<>"",${year}!$${properties.get('yearSheetSumColName')}${row},"")` : null);
      const sumFormula = row < properties.get('endRow') ? `=if(sum(${properties.get('startColName')}${row}:${properties.get('endColName')}${row})<>0, sum(${properties.get('startColName')}${row}:${properties.get('endColName')}${row}),"")` : null;
      return [...yearsFormula, sumFormula];
    });    
    sheet.getRange(1, properties.get('startColIdx'), formulas.length, formulas[0].length).setFormulas(formulas);
    sheet.getRange(`B${templateInfo.get('bodyStartRowIdx')}:C${templateInfo.get('bodyStartRowIdx')}`).setValues([['【見積明細：総期間】', null]]);
    const headerYear = [...this.yearList.map(x => String(x)), '合計'];
    sheet.getRange(templateInfo.get('bodyStartRowIdx'), properties.get('startColIdx'), 1, headerYear.length).clear();
    sheet.insertRowAfter(templateInfo.get('bodyStartRowIdx'));
    const yearRange = sheet.getRange(properties.get('yearRow'), properties.get('startColIdx'), 1, headerYear.length);
    yearRange.setValues([headerYear]);
    this.setFormat_(headerYear.length, yearRange);
    sheet.deleteRow(1);
  }
  /**
   * Set formatting.
   * @param {Number} headerLength Array length for the "Year" heading.
   * @param {Object} yearRange Range Object.
   * @return none.
   */
  setFormat_(headerLength, yearRange){
    const properties = this.getProperties_();
    const sheet = properties.get('sheet');
    sheet.setColumnWidth(getNumber_(properties.get('endColIdx')) + 1, 18);
    sheet.setColumnWidth(getNumber_(properties.get('filterColIdx')), 35);
    sheet.getDataRange().setBorder(false, false, false, false, false, false);
    sheet.getRange(templateInfo.get('bodyStartRowIdx'), getNumber_(templateInfo.get('colItemNameAndIdx').get('primaryItem')), 1, properties.get('endColIdx')).setBorder(true, true, true, true, false, false);
    sheet.getRange(properties.get('yearRow'), getNumber_(templateInfo.get('colItemNameAndIdx').get('primaryItem')), properties.get('bodyRowCount'), properties.get('endColIdx')).setBorder(true, true, true, true, false, false);
    sheet.getRange(properties.get('yearRow'), getNumber_(templateInfo.get('colItemNameAndIdx').get('price')), 1, headerLength).setBorder(true, true, true, true, true, true);
    sheet.getRange(properties.get('yearRow'), getNumber_(templateInfo.get('colItemNameAndIdx').get('price')), properties.get('bodyRowCount'), headerLength).setBorder(true, true, true, true, true, null);
    sheet.getRange(properties.get('endRow'), getNumber_(templateInfo.get('colItemNameAndIdx').get('primaryItem')), 2, properties.get('endColIdx')).setBorder(true, true, true, true, null, true);
    sheet.getRange(properties.get('endRow'), getNumber_(templateInfo.get('colItemNameAndIdx').get('price')), 2, headerLength).setBorder(true, true, true, true, true, true);
    yearRange.setHorizontalAlignment('center');
    sheet.getRange(templateInfo.get('bodyStartRowIdx'), getNumber_(templateInfo.get('colItemNameAndIdx').get('price')), properties.get('bodyRowCount'), properties.get('endColIdx')).setNumberFormat('#,##0');
  }
}