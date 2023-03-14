/**
 * Build a template sheet.
 * @param {Object} template Sheet object.
 * @param {Object} items Sheet object.
 */
function createTemplate_(ss, template, items){
  const itemsSheetRawValues = spreadSheetBatchUpdate.rangeGetValue(ss.spreadsheetId, [`${items.title}!R${getNumber_(itemsInfo.get('bodyStartRowIdx'))}C${getNumber_(itemsInfo.get('colItemNameAndIdx').get('primaryItem'))}:R${items.gridProperties.rowCount}C${getNumber_(itemsInfo.get('colItemNameAndIdx').get('unit'))}`])[0].values;
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
  const editTemplateFormulas = new EditTemplateFormulas();
  const itemsTemplateRowDiff = getNumber_(templateInfo.get('bodyStartRowIdx') - itemsInfo.get('bodyStartRowIdx'));
  const itemsRangesValues = editItemValues_(itemsSheetValues);
  const amountCol = getColumnString_(getNumber_(templateInfo.get('colItemNameAndIdx').get('amount')));
  const countCol = getColumnString_(getNumber_(templateInfo.get('colItemNameAndIdx').get('count')));
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
      const sumExcludedFilter = rows[primaryItemExcludedIdx] ? `=if(sum(${getColumnString_(getNumber_(templateInfo.get('colItemNameAndIdx').get('sum')))}${getNumber_(rowNumber)}) > 0, 1, 0)` : `=if(sum(${countCol}${getNumber_(sumStartIdx) + outputBodyStartRowNumber}:${countCol}${getNumber_(sumExcludedEndIdx) + outputBodyStartRowNumber})>0, 1, 0)`;
      return editTemplateFormulas.editPrimaryItem(itemsRowNumber, sumFormula, sumExcludedFilter, ss);
    } else {
      return editTemplateFormulas.editSecondaryItem(getNumber_(rowNumber), itemsRowNumber, ss);
    }
  });
  const itemsLastRowNumber = formulas.length - 1 + outputBodyStartRowNumber;
  const totalRowNumber = itemsLastRowNumber + 1;
  const itemsTotal = [
    ['合計', '（税抜）', '', '', '', '', `=sum(${amountCol}${outputBodyStartRowNumber + 1}:${amountCol}${totalRowNumber})`, '', '', '', 1],
    ['割引後合計','', '', '', '', '', `=${amountCol}${totalRowNumber + 1}*(1-Trial!$B$47)`, '', '', '', '=if(Trial!$B$47=0, 0, 1)'],
  ];
  const itemsBody = [...formulas, ...itemsTotal];
  const setBodyRequest = spreadSheetBatchUpdate.getRangeSetValueRequest(template.properties.sheetId,
                                                                        outputBodyStartRowNumber,
                                                                        templateInfo.get('startColIdx'),
                                                                        itemsBody);
  // Set up formulas individually only for project management.
  // 後にしよ
//  new ProjectManagement().setTemplate_(template);

//  setTemplateFilter_(template);
//  const setBorders = setTemplateFormat_(template, outputBodyStartRowNumber - 1, totalRowNumber + 1);
  const requests = [setHeadRequest, setBodyRequest];
  return requests;
}
/**
 * Set the format of the template sheet.
 * @param {Object} template Sheet object.
 * @return none.
 */
function setTemplateFormat_(template, outputHeaderBodyRow, lastRow){
  //templateInfo.get('colWidths').forEach((value, key) => template.setColumnWidth(key, value));
//  const outputHeaderBodyRow = templateInfo.get('outputBodyStartRowNumber') - 1;
//  const lastRow = templateInfo.get('totalRowNumber') + 1;
  const bSolid = spreadSheetBatchUpdate.setBorderInfo();
  const bNone = spreadSheetBatchUpdate.delBorderInfo();
  const test = spreadSheetBatchUpdate.setBorders(template.id, {startRow: 1, startCol:1, endRow:1, endCol: 1}, {top: bSolid, bottom: bSolid, left: bSolid, right: bSolid, innerHorizontal: bNone, innerVertical: bNone});
  return;
  template.getRangeList([
    `${templateInfo.get('colNames').get('primaryItem')}${templateInfo.get('outputStartRowNumber')}:${templateInfo.get('colNames').get('rightBorderEnd')}${templateInfo.get('outputStartRowNumber') + 1}`,
    `${templateInfo.get('colNames').get('primaryItem')}${outputHeaderBodyRow}:${templateInfo.get('colNames').get('rightBorderEnd')}${outputHeaderBodyRow}`,
    `${templateInfo.get('colNames').get('primaryItem')}${templateInfo.get('outputBodyStartRowNumber')}:${templateInfo.get('colNames').get('rightBorderEnd')}${templateInfo.get('totalRowNumber') - 1}`,
    `${templateInfo.get('colNames').get('primaryItem')}${templateInfo.get('totalRowNumber')}:${templateInfo.get('colNames').get('rightBorderEnd')}${templateInfo.get('totalRowNumber')}`,
    `${templateInfo.get('colNames').get('primaryItem')}${lastRow}:${templateInfo.get('colNames').get('rightBorderEnd')}${lastRow}`,
  ]).setBorder(true, true, true, true, null, null);
  template.getRangeList([
    `${templateInfo.get('colNames').get('primaryItem')}${outputHeaderBodyRow}:${templateInfo.get('colNames').get('secondaryItem')}${lastRow}`,
    `${templateInfo.get('colNames').get('amount')}${outputHeaderBodyRow}:${templateInfo.get('colNames').get('amount')}${templateInfo.get('totalRowNumber') - 1}`,
  ]).setBorder(null, true, null, true, null, null);
  template.getRangeList([
    `${templateInfo.get('colNames').get('primaryItem')}:${templateInfo.get('colNames').get('primaryItem')}`, 
    `${templateInfo.get('colNames').get('sum')}${templateInfo.get('outputBodyStartRowNumber')}:${templateInfo.get('colNames').get('sum')}`, 
    `${templateInfo.get('colNames').get('amount')}${templateInfo.get('totalRowNumber')}:${templateInfo.get('colNames').get('amount')}${lastRow}`,
  ]).setFontWeight('bold');
  template.getRangeList([
    `${templateInfo.get('colNames').get('x')}:${templateInfo.get('colNames').get('x')}`, 
    `${templateInfo.get('colNames').get('count')}:${templateInfo.get('colNames').get('count')}`, 
    `${templateInfo.get('colNames').get('amount')}:${templateInfo.get('colNames').get('sum')}`, 
    `${templateInfo.get('colNames').get('price')}${outputHeaderBodyRow}`,
  ]).setHorizontalAlignment('right');
  template.getRangeList([
    `${templateInfo.get('colNames').get('secondaryItem')}${outputHeaderBodyRow}`, 
    `${templateInfo.get('colNames').get('filter')}:${templateInfo.get('colNames').get('filter')}`, 
    `${templateInfo.get('colNames').get('amount')}${outputHeaderBodyRow}:${templateInfo.get('colNames').get('sum')}${outputHeaderBodyRow}`,
  ]).setHorizontalAlignment('center');
  template.getRangeList([
    `${templateInfo.get('colNames').get('price')}${outputHeaderBodyRow}:${templateInfo.get('colNames').get('price')}`,
    `${templateInfo.get('colNames').get('amount')}${outputHeaderBodyRow}:${templateInfo.get('colNames').get('sum')}`,
  ]).setNumberFormat('#,##0');
  // Sets the row height of the primary item.
  for (let i = templateInfo.get('outputBodyStartRowNumber'); i < templateInfo.get('totalRowNumber'); i++){
    if (template.getRange(i, templateInfo.get('colItemNameAndIdx').get('amount')).getValue() === ''){
      template.setRowHeight(i, 36);
    }
  }
}
/**
 * Set filter.
 * @param {Object} Sheet object.
 * @return none.
 */
function setTemplateFilter_(template){
  const filterColName = getColumnString_(getNumber_(templateInfo.get('colItemNameAndIdx').get('filter')), template);
  const targetRange = template.getRange(`${filterColName}${templateInfo.get('outputBodyStartRowNumber') -1}:${filterColName}`);
  setFilter_(template, targetRange);
}
/**
 * Set the heading information for the template sheet.
 * @param {Object} ss Spreadsheet object.
 * @param {Object} template Sheet object.
 * @return none.
 */
/*
function setTemplateHeader_(ss, template){
  const itemsHead = [
    ['【見積明細：1年毎(xxxx年度)】', '', '', '', '', '', '', ''],
    ['', '', '単価', '', '', '', '', ''],
    ['', '項目', '摘要', '', '', '', '金額', '合計金額'],
  ];
  const batchUpdate = new SpreadSheetBatchUpdate();
  batchUpdate.rangeSetValue(ss.spreadsheetId, template.properties.sheetId,
                            batchUpdate.getRangeSetValueRequest(template.properties.sheetId,
                                                                templateInfo.get('headStartRowIdx'),
                                                                templateInfo.get('startColIdx'),
                                                                itemsHead));
}*/
/**
 * Set the heading information for the template sheet.
 * @param {Array.<string, string>} Value of the items sheet.
 * @return {Array.<string, string>}
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
class EditTemplateFormulas{
  constructor(){
    this.items = itemsInfo.get('sheet');
    this.itemsSheetName = this.items.title.replace(' のコピー', '');
    this.countCol = getColumnString_(getNumber_(templateInfo.get('colItemNameAndIdx').get('count'))); 
  }
  editPrimaryItem(itemsRowNumber, sumFormula, sumExcludedFilter, ss){
      return [
        `=${this.itemsSheetName}!$${getColumnString_(1, ss)}$${itemsRowNumber}`,
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
      ]
  }
  editSecondaryItem(rowNumber, itemsRowNumber, ss){
    return [
      '',
      `=${this.itemsSheetName}!$${getColumnString_(2, ss)}$${itemsRowNumber}`,
      `=${this.itemsSheetName}!$${getColumnString_(3, ss)}$${itemsRowNumber}`,
      '"x"',
      '',
      `=${this.itemsSheetName}!$${getColumnString_(4, ss)}$${itemsRowNumber}`,
      `=if(${this.countCol}${rowNumber}="", "", ${getColumnString_(getNumber_(templateInfo.get('colItemNameAndIdx').get('price')))}${rowNumber} * ${this.countCol}${rowNumber})`,
      '',
      '',
      '',
      `=if(sum(${this.countCol}${rowNumber}) > 0, 1, 0)`,
    ];
  }
}