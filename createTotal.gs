/**
 * Create a Total, Total2 sheet.
 */
class CreateTotalSheet{
  /**
   * @param {Object} ss The spreadsheet object.
   * @param {Object} targets The map object of the sheets.
   */
  constructor(ss, targets){
    this.ss = ss;
    this.yearList = [];
    targets.forEach((_, sheetName) => {
      this.yearList.push(sheetName)
    });
    this.totalSheet = this.getSheet(commonInfo.get('totalSheetName'));
    this.total2Sheet = this.getSheet(commonInfo.get('total2SheetName'));
    this.templateSheet = this.getSheet(commonInfo.get('templateSheetName'));
    this.totalHeadText = '【見積明細：総期間】';
    this.countColName = colNamesConstant[getNumber_(templateInfo.get('colItemNameAndIdx').get('count'))];
  }
  getSheet(sheetName){
    const sheet = this.ss.sheets.filter(x => x.properties.title === sheetName);
    if (sheet.length !== 1){
      return null;
    }
    return sheet[0];
  }
  /**
   * Edit total, total2 sheet.
   * @param none.
   * @return {Object} Request body.
   */
  exec(){
    let res = [];
    res.push(this.editTotal2Sheet_());
    res.push(this.editTotalSheet_());
    return [res];
  }
  /**
   * Edit total2 sheet.
   * @param none.
   * @return {Object} Request body.
   */
  editTotal2Sheet_(){
    // Delete columns D and after and add years + 3 columns.
    this.outputStartIdx = templateInfo.get('colItemNameAndIdx').get('price');
    this.sumColIdx = this.yearList.length + this.outputStartIdx + 1;
    this.sheetId = this.total2Sheet.properties.sheetId;
    const delColRequest = spreadSheetBatchUpdate.getdelRowColRequest(this.sheetId, 'COLUMNS', this.outputStartIdx, this.total2Sheet.properties.gridProperties.columnCount - this.outputStartIdx);
    const insertColRequest = spreadSheetBatchUpdate.getInsertRowColRequest(this.sheetId, 'COLUMNS', this.outputStartIdx, this.yearList.length + 3);
    const insertRowRequest = spreadSheetBatchUpdate.getInsertRowColRequest(this.sheetId, 'ROWS', 3, 4);
    const primaryItemColName = colNamesConstant[getNumber_(templateInfo.get('colItemNameAndIdx').get('primaryItem'))];
    const secondaryItemColName = colNamesConstant[getNumber_(templateInfo.get('colItemNameAndIdx').get('secondaryItem'))];
    const primaryItemRange = `${commonInfo.get('total2SheetName')}!${primaryItemColName}1:${secondaryItemColName}${this.total2Sheet.properties.gridProperties.rowCount}`;
    const itemsValues = spreadSheetBatchUpdate.rangeGetValue(this.ss.spreadsheetId, primaryItemRange)[0].values;
    this.lastRowIdx = itemsValues.length;
    let primaryRowIndex = [];
    let secondaryRowIndex =[];
    let discountedSumRowIdx;
    this.sumRowIdx;
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
      if (value[0] === '合計' && value[1] === '（税抜）'){
        this.sumRowIdx = idx;
      }
    });
    let bodyRowsArray = [];
    for (let i = 6; i <= this.total2Sheet.properties.gridProperties.rowCount; i++){
      bodyRowsArray.push(i);
    }
    const outputStartColName = colNamesConstant[getNumber_(this.outputStartIdx)];
    const outputEndColName = colNamesConstant[getNumber_(this.outputStartIdx + this.yearList.length - 1)];
    const setBodyFormulas = bodyRowsArray.map(row => {
    // The rows after the discounted total fills in another formula.
      const yearsFormula = this.yearList.map((year, idx) => row <= getNumber_(discountedSumRowIdx) 
        ? `=${String(year)}!$H${row - 1}` 
        : row === getNumber_(discountedSumRowIdx) + 1
          ? `=if(and(${colNamesConstant[getNumber_(idx + this.outputStartIdx)]}${row - 1} <> "", ${trialInfo.get('sheetName')}!${trialInfo.get('discountRateAddress')} <> 0), (${colNamesConstant[getNumber_(idx + this.outputStartIdx)]}${row - 1} * (1 - ${trialInfo.get('sheetName')}!${trialInfo.get('discountRateAddress')})), ${colNamesConstant[getNumber_(idx + this.outputStartIdx)]}${row - 1})`
          : `=${colNamesConstant[getNumber_(idx + this.outputStartIdx)]}${row - 2} * (1 + ${trialInfo.get('sheetName')}!${trialInfo.get('taxAddress')})`);
      const sumFormula = `=if(sum(${outputStartColName}${row}:${outputEndColName}${row})=0, "", sum(${outputStartColName}${row}:${outputEndColName}${row}))`;
      const filterFormula = `=${commonInfo.get('totalSheetName')}!${colNamesConstant[getNumber_(templateInfo.get('colItemNameAndIdx').get('filter'))]}${row - 1}`;
      return [...yearsFormula, sumFormula, '', filterFormula];
    });
    const headerValuesArray = new Array(this.yearList.length + this.outputStartIdx + 1).fill('');
    const headerValues = [
      ['', '【期間別見積】', ...headerValuesArray.slice(2)],
      headerValuesArray,
      ['', this.totalHeadText, ...headerValuesArray.slice(2)],
      ['', '', '',　...this.yearList.map(x => String(x)), '合計'],
    ];
    const setBodyRequest = [
      spreadSheetBatchUpdate.getRangeSetValueRequest(this.sheetId,
                                                     5,
                                                     this.outputStartIdx,
                                                     setBodyFormulas), 
      spreadSheetBatchUpdate.getRangeSetValueRequest(this.sheetId,
                                                     0,
                                                     0,
                                                     headerValues),
    ];
    // Column width setting
    const yearsColWidths = 81;
    const colWidths = [25, 38, 447, ...this.yearList.map(_ => yearsColWidths), yearsColWidths, 18, 35];
    const filterColIdx = colWidths.length - 1;
    const setColWidthRequest = colWidths.map((width, idx) => spreadSheetBatchUpdate.getSetColWidthRequest(this.sheetId, width, idx, idx + 1));
    // Border setting
    const bordersRequest = this.setBorders_();
    const delRowsRequest = [
      spreadSheetBatchUpdate.getdelRowColRequest(this.sheetId, 'ROWS', 4, 5),
    ];
    const formatRequest = [
      spreadSheetBatchUpdate.getRangeSetFormatRequest(this.sheetId, 
                                                      0, 
                                                      templateInfo.get('colItemNameAndIdx').get('primaryItem'),
                                                      1, 
                                                      templateInfo.get('colItemNameAndIdx').get('primaryItem'), 
                                                      spreadSheetBatchUpdate.getFontBoldRequest(), 
                                                      'userEnteredFormat.textFormat.bold'),
      spreadSheetBatchUpdate.getRangeSetFormatRequest(this.sheetId, 
                                                      3,
                                                      this.outputStartIdx,
                                                      3, 
                                                      this.sumColIdx, 
                                                      spreadSheetBatchUpdate.getHorizontalAlignmentRequest('CENTER'), 
                                                      'userEnteredFormat.horizontalAlignment'),
      spreadSheetBatchUpdate.getSetRowHeightRequest(this.sheetId, 21, 0, 1),
      setNumberFormat_(this.total2Sheet, this.sumRowIdx, this.outputStartIdx, this.lastRowIdx, this.sumColIdx),
    ];
    const addConditionalFormatRuleTarget = spreadSheetBatchUpdate.getRangeGridByIdx(this.sheetId, 0, filterColIdx, this.lastRowIdx, filterColIdx);
    const addConditionalFormatRuleRequest = editConditionalFormatRuleRequest([addConditionalFormatRuleTarget,]);
    return [delColRequest, insertColRequest, insertRowRequest, ...setBodyRequest, ...delRowsRequest, ...setColWidthRequest, bordersRequest, formatRequest, ...addConditionalFormatRuleRequest];
  }
  setBorders_(){
    let request = [];
    const borderStyle = spreadSheetBatchUpdate.createBorderStyle();
    let borders = {
      'top': borderStyle.setBorderNone(),
      'bottom' : borderStyle.setBorderNone(),
      'left': borderStyle.setBorderNone(),
      'right': borderStyle.setBorderNone(),
      'innerHorizontal': borderStyle.setBorderNone(),
      'innerVertical': borderStyle.setBorderNone(),
    }
    let rowCol = {
      'startRowIndex': null,
      'endRowIndex' : null,
      'startColumnIndex' : null,
      'endColumnIndex': null,
    }
    rowCol = {
      'startRowIndex': 0,
      'endRowIndex' : this.total2Sheet.properties.gridProperties.rowCount,
      'startColumnIndex' : 0,
      'endColumnIndex': this.total2Sheet.properties.gridProperties.columnCount,
    }
    request.push(spreadSheetBatchUpdate.getUpdateBordersRequest(this.sheetId, rowCol, borders)); 
    borders = {
      'top': borderStyle.setBorderSolid(),
      'bottom' : borderStyle.setBorderSolid(),
      'left': borderStyle.setBorderSolid(),
      'right': borderStyle.setBorderSolid(),
    }
    delete borders.innerHorizontal;
    delete borders.innerVertical;
    rowCol = {
      'startRowIndex': 2,
      'endRowIndex' : this.sumRowIdx,
      'startColumnIndex' : 1,
      'endColumnIndex': this.sumColIdx,
    }
    request.push(spreadSheetBatchUpdate.getUpdateBordersRequest(this.sheetId, rowCol, borders)); 
    rowCol = {
      'startRowIndex': 2,
      'endRowIndex' : 3,
      'startColumnIndex' : 1,
      'endColumnIndex': this.sumColIdx,
    }
    request.push(spreadSheetBatchUpdate.getUpdateBordersRequest(this.sheetId, rowCol, borders)); 
    rowCol = {
      'startRowIndex': 3,
      'endRowIndex' : 4,
      'startColumnIndex' : this.outputStartIdx,
      'endColumnIndex': this.sumColIdx,
    }
    borders.innerHorizontal = borderStyle.setBorderSolid();
    borders.innerVertical = borderStyle.setBorderSolid();
    request.push(spreadSheetBatchUpdate.getUpdateBordersRequest(this.sheetId, rowCol, borders)); 
    rowCol = {
      'startRowIndex': 4,
      'endRowIndex' : this.lastRowIdx,
      'startColumnIndex' : this.outputStartIdx,
      'endColumnIndex': this.sumColIdx,
    }
    delete borders.innerHorizontal;
    request.push(spreadSheetBatchUpdate.getUpdateBordersRequest(this.sheetId, rowCol, borders)); 
    rowCol = {
      'startRowIndex': this.sumRowIdx,
      'endRowIndex' : this.lastRowIdx,
      'startColumnIndex' : this.outputStartIdx,
      'endColumnIndex': this.sumColIdx,
    }
    borders.innerHorizontal = borderStyle.setBorderSolid();
    request.push(spreadSheetBatchUpdate.getUpdateBordersRequest(this.sheetId, rowCol, borders)); 
    rowCol = {
      'startRowIndex': this.sumRowIdx,
      'endRowIndex' : this.lastRowIdx,
      'startColumnIndex' : 1,
      'endColumnIndex': this.outputStartIdx,
    }
    delete borders.innerVertical;
    request.push(spreadSheetBatchUpdate.getUpdateBordersRequest(this.sheetId, rowCol, borders)); 
    return request;
  }
  /**
   * Edit total sheet.
   * @param none.
   * @return {Object} Request body.
   */
  editTotalSheet_(){
    const formulas = [];
    for (let i = templateInfo.get('bodyStartRowIdx'); i < this.totalSheet.properties.gridProperties.rowCount - templateInfo.get('bodyStartRowIdx'); i++){
      const formula = this.yearList.map(sheetName => `'${sheetName}'!${this.countColName}${i + 1}`).join(' + ');
      formulas.push([`=if(${formula} > 0, ${formula}, "")`]);
    }
    this.sheetId = this.totalSheet.properties.sheetId;
    const setFormulasRequest = spreadSheetBatchUpdate.getRangeSetValueRequest(this.sheetId,
                                                                              templateInfo.get('bodyStartRowIdx'),
                                                                              templateInfo.get('colItemNameAndIdx').get('count'),
                                                                              formulas);
    const setHeadTextRequest = spreadSheetBatchUpdate.getRangeSetValueRequest(this.sheetId,
                                                                        templateInfo.get('headStartRowIdx'),
                                                                        templateInfo.get('startColIdx'),
                                                                        [[this.totalHeadText]]);
    return [setFormulasRequest, setHeadTextRequest];
  }
}