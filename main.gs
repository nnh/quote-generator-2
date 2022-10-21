/**
 * @param {Object} Request Parameters.
 * @return {Object} The generated HtmlService object.
 */
function doGet(e) {
  const param = e.parameter;
  const page = param.page ? param.page : 'index';
  let htmlOutput = HtmlService.createTemplateFromFile(page).evaluate();
  if (page === 'index'){
    htmlOutput
      .setTitle('title_index')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    return htmlOutput;
  } 
  if (page === 'quote'){
    htmlOutput
      .setTitle('title_quote')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    return htmlOutput;
  }
}
/**
 * @return {String} Returns the URL for this web application.
 */
function getAppUrl_() {
  return ScriptApp.getService().getUrl();
}
/**
 * Obtain the necessary information from the items sheet of 'Quote Template'.
 * @param none.
 * @return {Object} Returns an associative array of heading, item name, unit, unit price (per person-day), number of days, and number of persons.
 */
function getItemsList() {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('inputSsId'));
  const itemSheet = ss.getSheetByName('Items');
  const ColInfo = new ColumnInfo(itemSheet);
  const targetIndex = {
    itemValueHeading: 0,
    itemValueItem: 1,
    itemValueUnit: 3,
    itemValueUnitPrice: ColInfo.getColumnIndex('S'),
    itemValueDays: ColInfo.getColumnIndex('T'),
    itemValueNumberOfPeople: ColInfo.getColumnIndex('U')
  };
  const itemValues = itemSheet.getDataRange().getValues();
  let itemHeadingAndPrice = [];
  const _dummy = itemValues.reduce((x, currentValue) => {
    let res = currentValue;
    if (res[targetIndex.itemValueHeading] === ''){
      res[targetIndex.itemValueHeading] = x[targetIndex.itemValueHeading];
    }
    itemHeadingAndPrice.push(res);
    return res;
  });
  const res = itemHeadingAndPrice.filter(x => x[targetIndex.itemValueItem] !== '').map(x => {
    return {
      heading: x[targetIndex.itemValueHeading],
      item: x[targetIndex.itemValueItem],
      unit: x[targetIndex.itemValueUnit],
      unitPrice: x[targetIndex.itemValueUnitPrice],
      days: x[targetIndex.itemValueDays],
      numberOfPeople: x[targetIndex.itemValueNumberOfPeople]
    }
  });
  return res;
}
class ColumnInfo {
  /**
  * Returns information about the columns of the spreadsheet.
  * @param {Object} The sheet object in a spreadsheet.
  * @return none.
  */
  constructor(sheet) {
    this.sheet = sheet;
  }
  /**
  * Return column number from column name.
  * @param {String} The column name (as in 'A')
  * @return {Number} Column number, such as 1 for A.
  */
  getColumnNumber(columnName) {
    return this.sheet.getRange(columnName + '1').getColumn();
  }
  /**
  * Return column index from column name.
  * @param {String} The column name (as in 'A')
  * @return {Number} Column index, such as 0 for A.
  */
  getColumnIndex(columnName) {
    return this.getColumnNumber(columnName) - 1;
  }
} 

/**
* 列名から列番号を返す
* @param {string} column_name 列名（'A'など）
* @return Aなら1、のような列番号
*/
function getColumnNumber(column_name){ 
  const temp_sheet = SpreadsheetApp.getActiveSheet();
  const temp_range = temp_sheet.getRange(column_name + '1').getColumn();
  return(temp_range);
}
