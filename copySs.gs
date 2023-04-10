function testCreateSs(inputData){
  const ss = {};
  const now = driveCommon.todayYyyymmdd();
  ss.newSs = spreadSheetCommon.createNewSpreadSheet(`test${now}`);
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
  const setItemsRequest = setItemsSheet_(ss.newSs, ss.items);
  const createTemplateRequest = createTemplate_(ss.newSs, ss.template, ss.items);
  const requestsArray = [
                         setTrialRequest,
                         quotationRequestRequests,
                         setItemsRequest,
                         createTemplateRequest,
                        ];
  const requests = spreadSheetBatchUpdate.editBatchUpdateRequest(requestsArray);
  spreadSheetBatchUpdate.execBatchUpdate(requests, ss.newSs.spreadsheetId);
  // Fix the template and then copy the sheets for each year, total, total2.
  const targetYearsSheet = copyTemplate_(ss.newSs, ss.template);
  const targetYearsRename = Array.from(targetYearsSheet.keys()).map(key => [targetYearsSheet.get(key).sheetId, String(key)]);
  const targetYearsRenameRequests = targetYearsRename.map(x => spreadSheetBatchUpdate.editRenameSheetRequest(x[0], x[1]));
  spreadSheetBatchUpdate.execBatchUpdate(spreadSheetBatchUpdate.editBatchUpdateRequest(targetYearsRenameRequests), ss.newSs.spreadsheetId);
  const totalSheetRequest = new CreateTotalSheet(ss.newSs, targetYearsSheet).exec();
  const totalRequestsArray = [
                               totalSheetRequest,
                             ];
  const totalRequests = spreadSheetBatchUpdate.editBatchUpdateRequest(totalRequestsArray);
  spreadSheetBatchUpdate.execBatchUpdate(totalRequests, ss.newSs.spreadsheetId);
  setPropertiesByTrialType_(inputData);
  const setValuesRegistration = new SetValuesRegistrationSheet(inputData, ss.newSs);
  let idx = 0;
  let filterRequests = [];
  targetYearsSheet.forEach((_, year, arr) => {
    const targetSheetCheck = /^\d{4}$/.test(String(year));
    let res;
    if (targetSheetCheck){
      res = setValuesRegistration.exec_(year);
      if (idx === 0){
        res = new SetValuesSetupSheet(inputData, ss.newSs).exec_(year);
      }
      if (idx === arr.length - 1){
        res = new SetValuesClosingSheet(inputData, ss.newSs).exec_(year);
      }
    } else {
      res = new SetFilterTotalSheet(inputData, ss.newSs).exec_(year);
    }
    idx++;
    if (res){
      filterRequests.push(res);
    }
  });
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
  const targetItems = [commonInfo.get('trialTypeItemName'), '目標症例数', '施設数', 'CRF項目数'];  
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
 * @param {string} sheetId Sheet Id.
 * @return {Object} request object.
 */
function setItemsSheet_(ss, items){
  itemsInfo.set('sheet', items);
  const sheetId = items.properties.sheetId;
  const sheetName = items.properties.title;
  const itemsColIdxList = itemsInfo.get('colItemNameAndIdx');
  const formulaColsIdx = [
    itemsColIdxList.get('secondaryItem'),
    itemsColIdxList.get('price'),
    itemsColIdxList.get('baseUnitPrice'),
  ];
  const requests = formulaColsIdx.map(formulaColIdx => {
    const colString = commonGas.getColumnStringByIndex(formulaColIdx);
    const setItems = spreadSheetBatchUpdate.rangeGetValue(ss.spreadsheetId, `${sheetName}!${colString}:${colString}`, 'FORMULA')[0].values.map(x => x.length === 1 ? x : ['']);
    return spreadSheetBatchUpdate.getRangeSetValueRequest(sheetId, 
                                                          0, 
                                                          formulaColIdx, 
                                                          setItems);
  });
  return requests;                                                                           
}
