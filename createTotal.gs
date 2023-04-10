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
    const primaryItemColName = colNamesConstant[getNumber_(templateInfo.get('colItemNameAndIdx').get('primaryItem'))];
    const secondaryItemColName = colNamesConstant[getNumber_(templateInfo.get('colItemNameAndIdx').get('secondaryItem'))];
    const primaryItemRange = `${commonInfo.get('total2SheetName')}!${primaryItemColName}1:${secondaryItemColName}${this.total2Sheet.gridProperties.rowCount}`;
    const itemsValues = spreadSheetBatchUpdate.rangeGetValue(this.ss.spreadsheetId, primaryItemRange)[0].values;
    let primaryRowIndex = [];
    let secondaryRowIndex =[];
    let discountedSumRowIdx;
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
      if (value[0] === '割引後合計'){
        discountedSumRowIdx = idx;
      }
    });
    let bodyRowsArray = [];
    for (let i = 5; i < this.total2Sheet.gridProperties.rowCount; i++){
      bodyRowsArray.push(i);
    }
    const outputStartColName = colNamesConstant[getNumber_(outputStartIdx)];
    const outputEndColName = colNamesConstant[getNumber_(outputStartIdx + this.yearList.length - 1)];
    const setBodyFormulas = bodyRowsArray.map(row => {
    // The rows after the discounted total fills in another formula.
      const yearsFormula = this.yearList.map((year, idx) => row <= getNumber_(discountedSumRowIdx) ? `=${String(year)}!$H${row - 1}` : `=if(and(${colNamesConstant[getNumber_(idx + outputStartIdx)]}${row - 1} <> "", ${trialInfo.get('sheetName')}!${trialInfo.get('discountRateAddress')} <> 0), (${colNamesConstant[getNumber_(idx + outputStartIdx)]}${row - 1} * (1 - ${trialInfo.get('sheetName')}!${trialInfo.get('discountRateAddress')})), ${colNamesConstant[getNumber_(idx + outputStartIdx)]}${row - 1})`);
      const sumFormula = `=if(sum(${outputStartColName}${row}:${outputEndColName}${row})=0, "", sum(${outputStartColName}${row}:${outputEndColName}${row}))`;
      const filterFormula = `=${commonInfo.get('totalSheetName')}!${colNamesConstant[getNumber_(templateInfo.get('colItemNameAndIdx').get('filter'))]}${row - 1}`;
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
                                                     1,
                                                     0,
                                                     headerValues),
    ];
    return [delColRequest, insertColRequest, insertRowRequest, ...setBodyRequest, ...delRowsRequest];
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