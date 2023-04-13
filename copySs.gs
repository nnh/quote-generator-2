function createSpreadsheet(inputData){
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
  spreadSheetBatchUpdate.execBatchUpdate(spreadSheetBatchUpdate.editBatchUpdateRequest([renameRequests]), ss.newSs.spreadsheetId);
  // Get the spreadsheet object again because the added sheet is not reflected.
  ss.newSs = Sheets.Spreadsheets.get(ss.newSs.spreadsheetId);
  [ss.items, ss.trial, ss.quotationRequest] = copySheetNames.map(x => ss.newSs.sheets.filter(sheet => sheet.properties.title === x)[0]);
  editTrialTerm_(inputData);
  const setTrialRequest = setTrialSheet_(inputData, ss.trial.properties.sheetId);
  const quotationRequestRequests = [
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
                         quotationRequestRequests,
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
  const targetYearsSheet = copyTemplate_(ss.newSs, ss.template);
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
  setPropertiesByTrialType_(inputData);
  const setValuesRegistration = new SetValuesRegistrationSheet(inputData, ss.newSs);
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
  const filterRequests = [...filterRequestsYears, ...filterRequestTotal];
  spreadSheetBatchUpdate.execBatchUpdate(spreadSheetBatchUpdate.editBatchUpdateRequest(filterRequests), ss.newSs.spreadsheetId);
}
/**
 * Configure settings for each TrialType.
 * @param {Object} inputData The map object of the information entered from the form.
 * @return none.
 */
function setPropertiesByTrialType_(inputData){
  commonInfo.set('investigatorInitiatedTrialFlag', inputData.get(commonInfo.get('trialTypeItemName')) === commonInfo.get('trialType').get('investigatorInitiatedTrial'));
  commonInfo.set('specifiedClinicalTrialFlag', inputData.get(commonInfo.get('trialTypeItemName')) === commonInfo.get('trialType').get('specifiedClinicalTrial'));
  // Establishment of secretariat operation: Add if the presence of company source or coordinating secretariat is "Yes" or if it is an investigator-initiated clinical trial.
  const clinicalTrialsOfficeFlag = inputData.get(commonInfo.get('sourceOfFundsTextItemName')) === commonInfo.get('commercialCompany') ||
                                   inputData.get('調整事務局の有無') === 'あり' ||
                                   commonInfo.get('investigatorInitiatedTrialFlag');
  commonInfo.set('clinicalTrialsOfficeFlag', clinicalTrialsOfficeFlag);
}
/**
 * Copy the template sheet by the number of contract years, total, total2.
 * @param {Object} ss Spreadsheet object.
 * @param {Object} template Sheet object.
 * @return Sheet object.
 */
function copyTemplate_(ss, template){
  const startYear = trialInfo.get('setupStart').getFullYear();
  const endYear = trialInfo.get('trialEnd').getMonth() === 2 ? trialInfo.get('closingEnd').getFullYear() : trialInfo.get('trialEnd').getFullYear();
  let targetYears = [...Array(endYear - startYear + 1)].map((_, idx) => startYear + idx);
  targetYears.push(commonInfo.get('totalSheetName'));
  targetYears.push(commonInfo.get('total2SheetName'));
  const targets = new Map();
  targetYears.forEach(year => targets.set(year, spreadSheetCommon.copySheet(ss.spreadsheetId, ss, template.properties.sheetId)));
  return targets;
}
/**
 * Calculate the contract start end date.
 * @param {Object} inputData Map object of the information entered from the form.
 * @return none.
 */
function editTrialTerm_(inputData){
  const trialStartYear = inputData.get(`${trialInfo.get('trialStartText')}${trialInfo.get('yearText')}`).replace(trialInfo.get('yearText'), '');
  const trialStartMonth = inputData.get(`${trialInfo.get('trialStartText')}${trialInfo.get('monthText')}`).replace(trialInfo.get('monthText'), '') - 1;
  const trialEndYear = inputData.get(`${trialInfo.get('trialEndText')}${trialInfo.get('yearText')}`).replace(trialInfo.get('yearText'), '');
  const trialEndMonth = inputData.get(`${trialInfo.get('trialEndText')}${trialInfo.get('monthText')}`).replace(trialInfo.get('monthText'), '');
  const target = [commonInfo.get('trialType').get('investigatorInitiatedTrial'), commonInfo.get('trialType').get('specifiedClinicalTrial')];
  const setupTerm = target.some(x => inputData.get(commonInfo.get('trialTypeItemName')) === x) ? 6 : 3; 
  const closingTerm = target.some(x => inputData.get(commonInfo.get('trialTypeItemName')) === x) ? 6 : 3; 
  const trialStart = new Date(trialStartYear, trialStartMonth, 1);
  const trialEnd = new Date(trialEndYear, trialEndMonth, 0);
  const setupStart = new Date(trialStart.getFullYear(), trialStart.getMonth() - setupTerm, trialStart.getDate());
  const closingEnd = new Date(trialEnd.getFullYear(), trialEnd.getMonth() + closingTerm + 1, 0);
  trialInfo.set('trialStart', trialStart);
  trialInfo.set('trialEnd', trialEnd);
  trialInfo.set('setupStart', setupStart);
  trialInfo.set('closingEnd', closingEnd);
  trialInfo.set('setupTerm', setupTerm);
  trialInfo.set('closingTerm', closingTerm);
}

/**
 * Set the Trial sheet values from the information entered.
 * @param {Object} inputData Map object of the information entered from the form.
 * @param {number} sheetId
 * @return {Object} request object.
 */
function setTrialSheet_(inputData, sheetId){
  const monthsCount = (trialInfo.get('closingEnd').getFullYear() * 12 + trialInfo.get('closingEnd').getMonth())
                       - (trialInfo.get('setupStart').getFullYear() * 12 + trialInfo.get('setupStart').getMonth());
  const targetItems = [commonInfo.get('trialTypeItemName'), '目標症例数', commonInfo.get('facilitiesItemName'), 'CRF項目数'];  
  const targetItemValues = targetItems.map(key => [inputData.get(key)]);
  const sourceOfFunds = inputData.get(commonInfo.get('sourceOfFundsTextItemName')) === commonInfo.get('commercialCompany') ? 1.5 : 1;
  const requests = [
    spreadSheetBatchUpdate.getRangeSetValueRequest(sheetId, 
                                                   trialInfo.get('trialTermsAddress').get('rowIdx') - 1, 
                                                   trialInfo.get('trialTermsAddress').get('startColIdx'), 
                                                   [
                                                    ['', '', Number(monthsCount) + 1],
                                                    [
                                                      Utilities.formatDate(trialInfo.get('setupStart'), 'JST', 'yyyy/MM/dd'), 
                                                      Utilities.formatDate(trialInfo.get('closingEnd'), 'JST', 'yyyy/MM/dd'), 
                                                      '=datedif(D40,(E40+1), "m")'
                                                    ]
                                                   ]),
    spreadSheetBatchUpdate.getRangeSetValueRequest(sheetId, 
                                                   26, 
                                                   1, 
                                                   targetItemValues),
    spreadSheetBatchUpdate.getRangeSetValueRequest(sheetId, 
                                                   43, 
                                                   1, 
                                                   [[sourceOfFunds]]),
  ];
  return requests;
}
/**
 * Re-set the formulas on the copied Items sheet.
 * @param {Object} ss The sheet object.
 * @param {Object} inputData Map object of the information entered from the form.
 * @return {Object} request object.
 */
function setItemsSheet_(ss, inputData){
  const itemsSheet = ss.sheets.filter(x => x.properties.title === itemsInfo.get('sheetName'));
  if (itemsSheet.length !== 1){
    return;
  }
  const items = itemsSheet[0];
  // 保険料とかは単価の設定がいる
  const secondaryItemValue = spreadSheetBatchUpdate.rangeGetValue(ss.spreadsheetId, `${itemsInfo.get('sheetName')}!B1:B85`);
  if (secondaryItemValue.length !== 1){
    return;
  }
  const secondaryItem = secondaryItemValue[0].values;
  const setPriceTarget = ['保険料'];
  const setPriceTargetNameAndIdx = setPriceTarget.map(itemText => {
    const idxArray = secondaryItem.map((x, idx) => x[0] === itemText ? idx: null).filter(x => x);
    return idxArray.length === 1 ? [itemText, idxArray[0]] : null; 
  }).filter(x => x);
  const setPriceTargetNameAndIdxMap = new Map(setPriceTargetNameAndIdx);
  itemsInfo.set('sheet', items);
  const itemsColIdxList = itemsInfo.get('colItemNameAndIdx');
  const formulaColsIdx = [
    itemsColIdxList.get('secondaryItem'),
    itemsColIdxList.get('price'),
    itemsColIdxList.get('baseUnitPrice'),
  ];
  const setFormulaRequest = formulaColsIdx.map(formulaColIdx => {
    const colString = commonGas.getColumnStringByIndex(formulaColIdx);
    const setItems = spreadSheetBatchUpdate.rangeGetValue(ss.spreadsheetId, `${items.properties.title}!${colString}:${colString}`, 'FORMULA')[0].values.map(x => x.length === 1 ? x : ['']);
    return spreadSheetBatchUpdate.getRangeSetValueRequest(items.properties.sheetId, 
                                                          0, 
                                                          formulaColIdx, 
                                                          setItems);
  });
  const setPriceRequest = setPriceTarget.map(itemText => {
    if (Number.isSafeInteger(inputData.get(itemText))){
      const targetRowIdx = setPriceTargetNameAndIdxMap.get(itemText); 
      return spreadSheetBatchUpdate.getRangeSetValueRequest(items.properties.sheetId, 
                                                            targetRowIdx, 
                                                            itemsColIdxList.get('price'), 
                                                            [[inputData.get(itemText)]]);
    } else {
      return null;
    }
  }).filter(x => x);
  let requests = [...setFormulaRequest];
  if (setPriceRequest.length > 0){
    requests.push(...setPriceRequest);
  }
  return requests;                                                                           
}
