function createSpreadsheet(inputData=null){
  if (!inputData){
    console.log('No information was submitted from the form.');
    return;
  }
  if (!setPropertiesByInputData_(inputData)){
    console.log('The information submitted on the form is missing.');
    return;
  }
  const ss = {};
  const now = driveCommon.todayYyyymmdd();
  ss.newSs = spreadSheetCommon.createNewSpreadSheet(`Quote ${inputData.get('試験実施番号')} ${now}`);
  ss.template = ss.newSs.sheets[0];
  const sheetIdMap = spreadSheetCommon.getSheetIdMap(Sheets.Spreadsheets.get(PropertiesService.getScriptProperties().getProperty('templateFileId')));
  // Copy the 'Items', 'Trial', and 'Quotation Request' sheets from the source file.
  const copySheetNames = [itemsInfo.get('sheetName'), trialInfo.get('sheetName'), 'Quotation Request'];
  const copySheets = copySheetNames.map(x => spreadSheetCommon.copySheet(PropertiesService.getScriptProperties().getProperty('templateFileId'), ss.newSs, sheetIdMap.get(x)));
  const renameRequests = [
                          [0, templateInfo.get('sheetName')],
                          ...copySheetNames.map((sheetName, idx) => [copySheets[idx].sheetId, sheetName]),
                         ].map(x => spreadSheetBatchUpdate.editRenameSheetRequest(x[0], x[1]));  
  spreadSheetBatchUpdate.execBatchUpdate(spreadSheetBatchUpdate.editBatchUpdateRequest(renameRequests), ss.newSs.spreadsheetId);
  // Get the spreadsheet object again because the added sheet is not reflected.
  ss.newSs = Sheets.Spreadsheets.get(ss.newSs.spreadsheetId);
  [ss.items, ss.trial, ss.quotationRequest] = copySheetNames.map(x => ss.newSs.sheets.filter(sheet => sheet.properties.title === x)[0]);
  const setTrialRequest = setTrialSheet_(inputData, ss.trial.properties.sheetId, ss.newSs);
  const setQuotationRequestRequests = [
    spreadSheetBatchUpdate.getRangeSetValueRequest(ss.quotationRequest.properties.sheetId, 
                                                   1, 
                                                   0, 
                                                   [[now]]),
  ];
  ss.newSs = Sheets.Spreadsheets.get(ss.newSs.spreadsheetId);
  const setItemsRequest = setItemsSheet_(ss.newSs, inputData);
  if (!setItemsRequest){
    console.log('error: setItemsSheet_');
    return;
  }
  const createTemplateRequest = createTemplate_(ss.newSs, ss.template, ss.items);
  const requestsArray = [
                         setTrialRequest,
                         setQuotationRequestRequests,
                         setItemsRequest,
                         createTemplateRequest,
                        ];
  const requests = spreadSheetBatchUpdate.editBatchUpdateRequest(requestsArray);
  // Fix the template and then copy the sheets for each year, total, total2.
  spreadSheetBatchUpdate.execBatchUpdate(requests, ss.newSs.spreadsheetId);
  // Set up formulas individually only for project management.
  const projectManagement = new ProjectManagement(ss.newSs);
  const projectManagementPriceRequest = projectManagement.setTemplate_(ss.template.properties.sheetId);
  const numberFormatRequest = setNumberFormat_(ss.template, projectManagement.getRowIdx(), templateInfo.get('colItemNameAndIdx').get('price'), projectManagement.getRowIdx(), templateInfo.get('colItemNameAndIdx').get('price'));
  spreadSheetBatchUpdate.execBatchUpdate(spreadSheetBatchUpdate.editBatchUpdateRequest([projectManagementPriceRequest, numberFormatRequest]), ss.newSs.spreadsheetId);
  // Create years, total, total2 sheet.
  const targetYearsSheet = createYearsAndTotalSheet_(ss.newSs, ss.template);
  const targetYearsRename = Array.from(targetYearsSheet.keys()).map(key => [targetYearsSheet.get(key).sheetId, String(key)]);
  const targetYearsRenameRequests = targetYearsRename.map(x => spreadSheetBatchUpdate.editRenameSheetRequest(x[0], x[1]));
  spreadSheetBatchUpdate.execBatchUpdate(spreadSheetBatchUpdate.editBatchUpdateRequest(targetYearsRenameRequests), ss.newSs.spreadsheetId);
  // Get the spreadsheet object again because the added sheet is not reflected.
  ss.newSs = Sheets.Spreadsheets.get(ss.newSs.spreadsheetId);
  targetYearsSheet.forEach((_, sheetName) => targetYearsSheet.set(sheetName, ss.newSs.sheets.filter(sheet => sheet.properties.title === String(sheetName))[0]));
  const totalSheetRequest = new CreateTotalSheet(ss.newSs, targetYearsSheet).exec();
  const totalRequestsArray = [
                               totalSheetRequest,
                             ];
  const totalRequests = spreadSheetBatchUpdate.editBatchUpdateRequest(totalRequestsArray);
  spreadSheetBatchUpdate.execBatchUpdate(totalRequests, ss.newSs.spreadsheetId);
  // Edit the sheet for each fiscal year.
  let targetYears = [];
  let targetTotal = [];
  targetYearsSheet.forEach((_, year) => {
    const targetSheetCheck = /^\d{4}$/.test(String(year));
    if (targetSheetCheck){
      targetYears.push(year);
    } else {
      targetTotal.push(year);
    }
  });
  // Get the start and end year of the registration.
  trialInfo.set('registrationStartYear', trialInfo.get('trialStart').getMonth() === 3 ? targetYears[1] : targetYears[0]);
  trialInfo.set('registrationEndYear', trialInfo.get('trialEnd').getMonth() === 2 ? targetYears[targetYears.length - 2] : targetYears[targetYears.length - 1]);
  trialInfo.set('registrationYearsCount', trialInfo.get('registrationEndYear') - trialInfo.get('registrationStartYear') + 1);  
  const setValuesRegistration = new SetValuesRegistrationSheet(inputData, ss.newSs);
  const filterRequestsYears = targetYears.map((year, idx) => {
    let res = setValuesRegistration.exec_(year);
    if (idx === 0){
      const _ = new SetValuesSetupSheet(inputData, ss.newSs).exec_(year);
    }
    if (idx === targetYears.length - 1){
      const _ = new SetValuesClosingSheet(inputData, ss.newSs).exec_(year);
    }
    return res;
  });
  const setFilterTotal = new SetFilterTotalSheet(inputData, ss.newSs)
  const filterRequestTotal = targetTotal.map(year => setFilterTotal.exec_(year));
  const moveSheetRequest = ss.newSs.sheets.map(sheet => {
    const sheetId = sheet.properties.sheetId;
    return sheet.properties.title === commonInfo.get('totalSheetName') 
      ? spreadSheetBatchUpdate.moveSheetRequest(sheetId, 0)
      : sheet.properties.title === commonInfo.get('total2SheetName')
        ? spreadSheetBatchUpdate.moveSheetRequest(sheetId, 1)
        : !new RegExp(sheet.properties.title).test(/^[0-9]{4}$/) 
          ? spreadSheetBatchUpdate.moveSheetRequest(sheetId)
          : null;
  }).filter(x => x);
  spreadSheetBatchUpdate.moveSheetRequest(0, 3);
  const filterRequests = [...filterRequestsYears, ...filterRequestTotal, ...moveSheetRequest];
  spreadSheetBatchUpdate.execBatchUpdate(spreadSheetBatchUpdate.editBatchUpdateRequest(filterRequests), ss.newSs.spreadsheetId);
}
/**
 * Copy the template sheet by the number of contract years, total, total2.
 * @param {Object} ss Spreadsheet object.
 * @param {Object} template Sheet object.
 * @return Sheet object.
 */
function createYearsAndTotalSheet_(ss, template){
  const startYear = trialInfo.get('setupStart').getFullYear();
  const endYear = trialInfo.get('trialEnd').getMonth() === 2 ? trialInfo.get('closingEnd').getFullYear() : trialInfo.get('trialEnd').getFullYear();
  let targetYears = [...Array(endYear - startYear + 1)].map((_, idx) => startYear + idx);
  targetYears.push(commonInfo.get('totalSheetName'));
  targetYears.push(commonInfo.get('total2SheetName'));
  const targets = new Map();
  targetYears.forEach(year => targets.set(year, spreadSheetCommon.copySheet(ss.spreadsheetId, ss, template.properties.sheetId)));
  return targets;
}
