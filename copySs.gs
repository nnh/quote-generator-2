/**
 * Create a quotation using the information entered from the form.
 * @param {Object} inputData Map object of the information entered from the form.
 * @return none.
 */
function createSpreadsheet(inputData=null){
  if (!inputData){
    console.log('No information was submitted from the form.');
    return;
  }
  const resSetProperties = setPropertiesByInputData_(inputData);
  if (!resSetProperties){
    console.log(resSetProperties);
    return;
  }
  const templateFolder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('templateFolderId'));
  if (!templateFolder){
    return;
  }
  const tempFiles = templateFolder.getFilesByName('Quote Template for nmc oscr');
  if (!tempFiles){
    return;
  }
  const templateFile = tempFiles.next();
  if (templateFile.getId() !== PropertiesService.getScriptProperties().getProperty('templateFileId')){
    return;
  }
  const now = driveCommon.todayYyyymmdd();
  const newFile = templateFile.makeCopy(`Quote ${inputData.get('試験実施番号')} ${now}`, DriveApp.getRootFolder());
  const newSs = Sheets.Spreadsheets.get(newFile.getId());
  const targetSheetsName = ['Total', 'Total2', 'Items', 'Setup', 'Trial', 'Quote', 'Quotation Request'];
  const sheets = new Map();
  const tempDeleteSheetRequests = newSs.sheets.map(sheet => {
    const checkTarget = targetSheetsName.filter(sheetName => sheetName === sheet.properties.title).some(x => x);
    if (checkTarget){
      sheets.set(sheet.properties.title, sheet.properties.sheetId);
      return;
    }
    const request = spreadSheetBatchUpdate.editDeleteSheetRequest(sheet.properties.sheetId);
    return request;
  }).filter(x => x);
  const renameRequests = spreadSheetBatchUpdate.editRenameSheetRequest(sheets.get('Setup'), templateInfo.get('sheetName'));
  const setQuotationRequestRequests = spreadSheetBatchUpdate.getRangeSetValueRequest(
    sheets.get('Quotation Request'), 
    1, 
    0, 
    [[now]]
  );
  const setTemplateRequests = spreadSheetBatchUpdate.getRangeSetValueRequest(
    sheets.get('Setup'), 
    1, 
    1, 
    [['']]
  );
  const setTrialRequest = setTrialSheet_(inputData, sheets.get('Trial'), newSs);
  const setItemsRequest = setItemsSheet_(newSs, inputData);
  if (!setItemsRequest){
    console.log('error: setItemsSheet_');
    return;
  }
  // Create years, total, total2 sheet.
  const targetYearsSheet = createYearsAndTotalSheet_(newSs, sheets.get('Setup'));
  const targetYearsRename = Array.from(targetYearsSheet.keys()).map(key => [targetYearsSheet.get(key).sheetId, String(key)]);
  const targetYearsRenameRequests = targetYearsRename.map(x => spreadSheetBatchUpdate.editRenameSheetRequest(x[0], x[1]));
  targetYearsSheet.forEach((sheet, sheetName) => sheets.set(sheetName, sheet.sheetId));
  const totalSheetRequest = new CreateTotalSheet(newSs, targetYearsSheet).exec();
  const request1 = spreadSheetBatchUpdate.editBatchUpdateRequest([...tempDeleteSheetRequests, renameRequests, setQuotationRequestRequests, setTemplateRequests, setTrialRequest, setItemsRequest, targetYearsRenameRequests, totalSheetRequest]);
  spreadSheetBatchUpdate.execBatchUpdate(request1, newSs.spreadsheetId);
  let targetYears = [];
  targetYearsSheet.forEach((_, year) => targetYears.push(year));
  // Get the start and end year of the registration.
  trialInfo.set('registrationStartYear', trialInfo.get('trialStart').getMonth() === 3 ? targetYears[1] : targetYears[0]);
  trialInfo.set('registrationEndYear', trialInfo.get('trialEnd').getMonth() === 2 ? targetYears[targetYears.length - 2] : targetYears[targetYears.length - 1]);
  trialInfo.set('registrationYearsCount', trialInfo.get('registrationEndYear') - trialInfo.get('registrationStartYear') + 1);  
  const setValuesRegistration = new SetValuesRegistrationSheet(inputData, newSs);
  // Edit the sheet for each fiscal year.
  const filterRequestsYears = targetYears.map((year, idx) => {
    let res = setValuesRegistration.exec_(year);
    if (idx === 0){
      const _ = new SetValuesSetupSheet(inputData, newSs).exec_(year);
    }
    if (idx === targetYears.length - 1){
      const _ = new SetValuesClosingSheet(inputData, newSs).exec_(year);
    }
    return res;
  });
  const setFilterTotal = new SetFilterTotalSheet(inputData, newSs)
  const filterRequestTotal = [commonInfo.get('totalSheetName'), commonInfo.get('total2SheetName')].map(year => setFilterTotal.exec_(year));
  const moveSheetRequest = setMoveSheetRequest_(newSs, [commonInfo.get('totalSheetName'), commonInfo.get('total2SheetName'), ...targetYears.map(x => String(x))]);
  const filterRequests = [...filterRequestsYears, ...filterRequestTotal, ...moveSheetRequest];
  spreadSheetBatchUpdate.execBatchUpdate(spreadSheetBatchUpdate.editBatchUpdateRequest(filterRequests), newSs.spreadsheetId);
}
/**
 * Move the work sheet backward.
 * @param {Object} ss The spreadsheet object.
 * @return {Object} request object.
 */
function setMoveSheetRequest_(ss, targetSheetNames){
  const sheets = Sheets.Spreadsheets.get(ss.spreadsheetId).sheets;
  const res = sheets.map(sheet => {
    const idx = targetSheetNames.indexOf(sheet.properties.title);
    if (idx > -1){
      return spreadSheetBatchUpdate.moveSheetRequest(sheet.properties.sheetId, idx);      
    } else {
      return null;
    }
  }).filter(x => x);
  return res;
}
/**
 * Copy the template sheet by the number of contract years, total, total2.
 * @param {Object} ss Spreadsheet object.
 * @param {number} template The template sheet id.
 * @return Sheet object.
 */
function createYearsAndTotalSheet_(ss, template){
  const startYear = trialInfo.get('setupStart').getFullYear();
  const endYear = trialInfo.get('trialEnd').getMonth() === 2 ? trialInfo.get('closingEnd').getFullYear() : trialInfo.get('trialEnd').getFullYear();
  let targetYears = [...Array(endYear - startYear + 1)].map((_, idx) => startYear + idx);
  const targets = new Map();
  targetYears.forEach(year => targets.set(year, spreadSheetCommon.copySheet(ss.spreadsheetId, ss, template)));
  return targets;
}
