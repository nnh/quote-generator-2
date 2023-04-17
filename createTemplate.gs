/**
 * Build a template sheet.
 * @param {Object} template Sheet object.
 * @param {Object} items Sheet object.
 */
function createTemplate_(ss, template, items){
  const itemsSheetRawValues = spreadSheetBatchUpdate.rangeGetValue(ss.spreadsheetId, [`${items.properties.title}!R${getNumber_(itemsInfo.get('bodyStartRowIdx'))}C${getNumber_(itemsInfo.get('colItemNameAndIdx').get('primaryItem'))}:R${items.properties.gridProperties.rowCount}C${getNumber_(itemsInfo.get('colItemNameAndIdx').get('unit'))}`])[0].values;
  const maxLength = Math.max(...itemsSheetRawValues.map(x => x.length));
  const itemsSheetValues = itemsSheetRawValues.map(x => x.length < maxLength ? [...x, ...new Array(maxLength - x.length).fill('')] : x);
  const primaryItemFlagIdx = maxLength;
  const primaryItemExcludedIdx = primaryItemFlagIdx + 1;
  const headValues = [
                       ['【見積明細：1年毎(xxxx年度)】', '', '', '', '', '', '', ''],
                       ['', '', '単価', '', '', '', '', ''],
                       ['', '項目', '摘要', '', '', '', '金額', '合計金額'],
                     ];
  const setHeadRequest = spreadSheetBatchUpdate.getRangeSetValueRequest(template.properties.sheetId,
                                                                        templateInfo.get('headStartRowIdx'),
                                                                        templateInfo.get('startColIdx'),
                                                                        headValues);
  const outputBodyStartRowNumber = getNumber_(templateInfo.get('headStartRowIdx')) + headValues.length + 1;
  const editTemplateFormulas = new EditTemplateFormulas(items);
  const itemsTemplateRowDiff = getNumber_(templateInfo.get('bodyStartRowIdx') - itemsInfo.get('bodyStartRowIdx'));
  const itemsRangesValues = editItemValues_(itemsSheetValues);
  const amountCol = commonGas.getColumnStringByIndex(templateInfo.get('colItemNameAndIdx').get('amount'));
  const countCol = commonGas.getColumnStringByIndex(templateInfo.get('colItemNameAndIdx').get('count'));
  const formulas = itemsRangesValues.map((rows, rowIdx, itemsValues) => {
    const itemsRowNumber = rowIdx + itemsTemplateRowDiff;
    const rowNumber = rowIdx + outputBodyStartRowNumber;
    if (rows[primaryItemFlagIdx]){
      // Create a formula for the total amount.
      const sumStartIdx = rowIdx + 1;
      let sumEndIdx = sumStartIdx;
      let sumExcludedEndIdx = sumStartIdx;
      for (let i = sumStartIdx; i < itemsValues.length - 1; i++){
        if (itemsValues[i + 1][primaryItemFlagIdx] && itemsValues[i + 1][primaryItemExcludedIdx]){
            break;
        }
        sumEndIdx++;
      }
      for (let i = sumStartIdx; i < sumEndIdx; i++){
        if (!itemsValues[i + 1][primaryItemExcludedIdx]){
          break;
        }
        sumExcludedEndIdx++;
      }      
      const sumFormula = rows[primaryItemExcludedIdx] && rows[itemsInfo.get('colItemNameAndIdx').get('primaryItem')] !== '' 
                         ? `=sum(${amountCol}${getNumber_(sumStartIdx) + outputBodyStartRowNumber}:${amountCol}${getNumber_(sumEndIdx) + outputBodyStartRowNumber})`
                         : "";
      const sumExcludedFilter = rows[primaryItemExcludedIdx] ? `=if(sum(${commonGas.getColumnStringByIndex(templateInfo.get('colItemNameAndIdx').get('sum'))}${getNumber_(rowNumber)}) > 0, 1, 0)` : `=if(sum(${countCol}${getNumber_(sumStartIdx) + outputBodyStartRowNumber}:${countCol}${getNumber_(sumExcludedEndIdx) + outputBodyStartRowNumber})>0, 1, 0)`;
      return editTemplateFormulas.editPrimaryItem(itemsRowNumber, sumFormula, sumExcludedFilter);
    } else {
      return editTemplateFormulas.editSecondaryItem(getNumber_(rowNumber), itemsRowNumber);
    }
  });
  const itemsLastRowNumber = formulas.length - 1 + outputBodyStartRowNumber;
  const totalRowNumber = itemsLastRowNumber + 1;
  const itemsTotal = [
    ['合計', '（税抜）', '', '', '', '', `=sum(${amountCol}${outputBodyStartRowNumber + 1}:${amountCol}${totalRowNumber})`, '', '', '', 1],
    ['割引後合計','', '', '', '', '', `=${amountCol}${totalRowNumber + 1}*(1-${trialInfo.get('sheetName')}!${trialInfo.get('discountRateAddress')})`, '', '', '', `=if(${trialInfo.get('sheetName')}!${trialInfo.get('discountRateAddress')}=0, 0, 1)`],
    ['合計', '（税込）', '', '', '', '', `=${amountCol}${totalRowNumber + 1}*(1+${trialInfo.get('sheetName')}!${trialInfo.get('taxAddress')})`, '', '', '', `=if(${amountCol}${totalRowNumber + 1}<>${amountCol}${totalRowNumber + 3}, 1, 0)`],
  ];
  const itemsBody = [...formulas, ...itemsTotal];
  const setBodyRequest = spreadSheetBatchUpdate.getRangeSetValueRequest(template.properties.sheetId,
                                                                        outputBodyStartRowNumber,
                                                                        templateInfo.get('startColIdx'),
                                                                        itemsBody);
  const delRowsRequest = [
    spreadSheetBatchUpdate.getdelRowColRequest(template.properties.sheetId, 'ROWS', 4, 6),
    spreadSheetBatchUpdate.getdelRowColRequest(template.properties.sheetId, 'ROWS', totalRowNumber + 2, template.properties.gridProperties.rowCount),
  ];
  const setColWidthRequest = [21, 50, 453, 76, 13, 35, 46, 81, 75, 22, 18, 35].map((width, idx) => spreadSheetBatchUpdate.getSetColWidthRequest(template.properties.sheetId, width, idx, idx + 1));
  const autoResizeColRequest = ['C', 'D', 'H', 'I'].map(colName => {
    const idx = commonGas.getColumnIndex(colName);
    return spreadSheetBatchUpdate.getAutoResizeRowRequest(template.properties.id, idx, idx)
  });
  const lastRow = totalRowNumber + 3;
  const bordersRequest = setTemplateBorders_(template, totalRowNumber, lastRow);
  const boldRequest = setTemplateBold_(template, totalRowNumber, lastRow);
  const horizontalAlignmentRequest = setTemplateHorizontalAlignment_(template);
  const numberFormatRequest = setNumberFormat_(template, 0, templateInfo.get('colItemNameAndIdx').get('amount'), lastRow, templateInfo.get('colItemNameAndIdx').get('sum'));
  const addConditionalFormatRuleTarget = spreadSheetBatchUpdate.getRangeGridByIdx(template.properties.sheetId, 0, templateInfo.get('colItemNameAndIdx').get('filter'), template.properties.gridProperties.rowCount, templateInfo.get('colItemNameAndIdx').get('filter'));
  const addConditionalFormatRuleRequest = editConditionalFormatRuleRequest([addConditionalFormatRuleTarget,]);
  const rowHeightArray = [...[13, 21, 21, 21], ...formulas.map(x => x[0] !== '' ? 36 : 21)];
  const setRowHeightRequest = rowHeightArray.map((height, idx) => spreadSheetBatchUpdate.getSetRowHeightRequest(template.properties.sheetId, height, idx, idx + 1));
  const requests = [setHeadRequest, setBodyRequest, ...setColWidthRequest, autoResizeColRequest, bordersRequest, boldRequest, horizontalAlignmentRequest, numberFormatRequest, horizontalAlignmentRequest, ...delRowsRequest, ...addConditionalFormatRuleRequest, ...setRowHeightRequest];
  return requests;
}
function setTemplateHorizontalAlignment_(template){
  const itemNameRowIdx = templateInfo.get('bodyStartRowIdx') - 1; 
  const request = [spreadSheetBatchUpdate.getRangeSetFormatRequest(template.properties.sheetId, 
                                                                   itemNameRowIdx, 
                                                                   templateInfo.get('colItemNameAndIdx').get('amount'),
                                                                   itemNameRowIdx, 
                                                                   templateInfo.get('colItemNameAndIdx').get('sum'), 
                                                                   spreadSheetBatchUpdate.getHorizontalAlignmentRequest('CENTER'), 
                                                                   'userEnteredFormat.horizontalAlignment'),
                  spreadSheetBatchUpdate.getRangeSetFormatRequest(template.properties.sheetId, 
                                                                   itemNameRowIdx, 
                                                                   templateInfo.get('colItemNameAndIdx').get('secondaryItem'),
                                                                   itemNameRowIdx, 
                                                                   templateInfo.get('colItemNameAndIdx').get('secondaryItem'), 
                                                                   spreadSheetBatchUpdate.getHorizontalAlignmentRequest('CENTER'), 
                                                                   'userEnteredFormat.horizontalAlignment'),
                  spreadSheetBatchUpdate.getRangeSetFormatRequest(template.properties.sheetId, 
                                                                   itemNameRowIdx, 
                                                                   templateInfo.get('colItemNameAndIdx').get('price'),
                                                                   itemNameRowIdx, 
                                                                   templateInfo.get('colItemNameAndIdx').get('price'), 
                                                                   spreadSheetBatchUpdate.getHorizontalAlignmentRequest('RIGHT'), 
                                                                   'userEnteredFormat.horizontalAlignment'),
                  ];
  return request;
}
function setTemplateBold_(template, totalRowNumber, lastRow){
  const request = [spreadSheetBatchUpdate.getRangeSetFormatRequest(template.properties.sheetId, 
                                                                   templateInfo.get('headStartRowIdx'), 
                                                                   templateInfo.get('colItemNameAndIdx').get('primaryItem'),
                                                                   lastRow, 
                                                                   templateInfo.get('colItemNameAndIdx').get('primaryItem'), 
                                                                   spreadSheetBatchUpdate.getFontBoldRequest(), 
                                                                   'userEnteredFormat.textFormat.bold'),
                   spreadSheetBatchUpdate.getRangeSetFormatRequest(template.properties.sheetId, 
                                                                   templateInfo.get('bodyStartRowIdx'), 
                                                                   templateInfo.get('colItemNameAndIdx').get('sum'),
                                                                   lastRow, 
                                                                   templateInfo.get('colItemNameAndIdx').get('sum'), 
                                                                   spreadSheetBatchUpdate.getFontBoldRequest(), 
                                                                   'userEnteredFormat.textFormat.bold'),
                   spreadSheetBatchUpdate.getRangeSetFormatRequest(template.properties.sheetId, 
                                                                   totalRowNumber, 
                                                                   templateInfo.get('colItemNameAndIdx').get('amount'),
                                                                   lastRow, 
                                                                   templateInfo.get('colItemNameAndIdx').get('amount'), 
                                                                   spreadSheetBatchUpdate.getFontBoldRequest(), 
                                                                   'userEnteredFormat.textFormat.bold'),
                  ];
  return request;
}
function setTemplateBorders_(template, totalRowNumber, lastRow){
  const itemNameRowIdx = templateInfo.get('bodyStartRowIdx') - 1; 
  let request = [];
  const borderStyle = spreadSheetBatchUpdate.createBorderStyle();
  let rowCol = {
    'startRowIndex': null,
    'endRowIndex' : null,
    'startColumnIndex' : null,
    'endColumnIndex': null,
  }
  let borders = {
    'top': borderStyle.setBorderSolid(),
    'bottom' : borderStyle.setBorderSolid(),
    'left': borderStyle.setBorderSolid(),
    'right': borderStyle.setBorderSolid(),
  }
  rowCol = {
    'startRowIndex': templateInfo.get('headStartRowIdx'),
    'endRowIndex' : lastRow,
    'startColumnIndex' : templateInfo.get('startColIdx'),
    'endColumnIndex': templateInfo.get('startColIdx') + templateInfo.get('colItemNameAndIdx').size,
  }
  request.push(spreadSheetBatchUpdate.getUpdateBordersRequest(template.properties.id, rowCol, borders)); 
  rowCol.startRowIndex = itemNameRowIdx;
  rowCol.endRowIndex = templateInfo.get('bodyStartRowIdx'); 
  request.push(spreadSheetBatchUpdate.getUpdateBordersRequest(template.properties.id, rowCol, borders));  
  rowCol.startRowIndex = totalRowNumber;
  rowCol.endRowIndex = lastRow;
  borders.innerHorizontal = borderStyle.setBorderSolid();
  request.push(spreadSheetBatchUpdate.getUpdateBordersRequest(template.properties.id, rowCol, borders));  
  delete borders.innerHorizontal;
  rowCol.startRowIndex = itemNameRowIdx;
  rowCol.endRowIndex = totalRowNumber;
  rowCol.startColumnIndex = templateInfo.get('colItemNameAndIdx').get('price');
  request.push(spreadSheetBatchUpdate.getUpdateBordersRequest(template.properties.id, rowCol, borders));  
  rowCol.startColumnIndex = templateInfo.get('colItemNameAndIdx').get('amount');
  rowCol.endColumnIndex = templateInfo.get('colItemNameAndIdx').get('sum');
  request.push(spreadSheetBatchUpdate.getUpdateBordersRequest(template.properties.id, rowCol, borders));  
  return request;
}
/**
 * Set the heading information for the template sheet.
 * @param {string[][]} Value of the items sheet.
 * @return {string[][]}
 */
function editItemValues_(itemsSheetValues){
  const primaryItemExcludedList = ['準備作業', 'EDC構築', '中央モニタリング', 'データセット作成'];
  const itemsColInfo = itemsInfo.get('colItemNameAndIdx');
  const itemsRangesItemToUnitValues = itemsSheetValues.map(x => [
    ...x, 
    x[itemsColInfo.get('secondaryItem')] === '', 
    primaryItemExcludedList.every(excluded => x[itemsColInfo.get('primaryItem')] !== excluded)
  ]);
  // Delete rows below the total.
  const sumRowIdx = itemsRangesItemToUnitValues.map((x, idx) => x[itemsColInfo.get('primaryItem')] === '合計' ? idx : null).filter(x => x)[0];
  const itemsRangesValues = itemsRangesItemToUnitValues.filter((_, idx) => idx < sumRowIdx);
  return itemsRangesValues;
}
/**
 * Set a Map object with key:item name (e.g., 'primaryItem') and value:column name (e.g., 'A').
 * @param {Object} sheet Sheet object.
 * @param {Object} targetMap Map object. 
 * @return none.
 */
function setColNamesInfo_(ss, targetMap){
  const colNames = new Map();
  targetMap.get('colItemNameAndIdx').forEach((idx, itemName) => colNames.set(itemName, getColumnString_(getNumber_(idx), ss)));
  targetMap.set('colNames', colNames);
}
/**
 * Create a list of functions to be set up on the template sheet.
 */
class EditTemplateFormulas{
  constructor(items){
    this.items = items;
    this.itemsSheetName = this.items.properties.title;
    this.countCol = commonGas.getColumnStringByIndex(templateInfo.get('colItemNameAndIdx').get('count')); 
  }
  /**
   * Create a function list of major items.
   * @param {number} itemsRowNumber Row number of the items sheet.
   * @param {string} sumFormula Total column function.
   * @param {string} sumExcludedFilter Function of the filter column.
   * @return {string[]} a function list of major items.
   */
  editPrimaryItem(itemsRowNumber, sumFormula, sumExcludedFilter){
    return [
      `=${this.itemsSheetName}!$${commonGas.getColumnStringByNumber(1)}$${itemsRowNumber}`,
      '',
      '',
      '',
      '',
      '',
      '',
      sumFormula,
      '',
      '',
      sumExcludedFilter,
    ];
  }
  editSecondaryItem(rowNumber, itemsRowNumber){
    return [
      '',
      `=${this.itemsSheetName}!$${commonGas.getColumnStringByNumber(2)}$${itemsRowNumber}`,
      `=${this.itemsSheetName}!$${commonGas.getColumnStringByNumber(3)}$${itemsRowNumber}`,
      'x',
      '',
      `=${this.itemsSheetName}!$${commonGas.getColumnStringByNumber(4)}$${itemsRowNumber}`,
      `=if(${this.countCol}${rowNumber}="", "", ${commonGas.getColumnStringByIndex(templateInfo.get('colItemNameAndIdx').get('price'))}${rowNumber} * ${this.countCol}${rowNumber})`,
      '',
      '',
      '',
      `=if(sum(${this.countCol}${rowNumber}) > 0, 1, 0)`,
    ];
  }
}