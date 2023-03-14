/**
 * @param {string} title Spreadsheet name.
 * @return {Object} Spreadsheet object.
 */
function createNewSpreadSheet_(title){
  const newSheet = Sheets.newSpreadsheet();
  newSheet.properties = Sheets.newSpreadsheetProperties();
  newSheet.properties.title = title;
  const ss = Sheets.Spreadsheets.create(newSheet);
  return ss;
}
/**
 * @param {Object} ss Spreadsheet object.
 * @param {number} sheetId
 * @return {Object} Sheet object.
 */
function copySheet_(ss, sheetId){
  const sheet = Sheets.Spreadsheets.Sheets.copyTo(
    {
      destinationSpreadsheetId: ss.spreadsheetId
    },
    PropertiesService.getScriptProperties().getProperty('templateFileId'),
    sheetId    
  );
  return sheet;
}
function testCreateSs(inputData){
  const ss = {};
  ss.ss = createNewSpreadSheet_('test20230215');
  ss.template = ss.ss.sheets[0];
  const copyFromSs = Sheets.Spreadsheets.get(PropertiesService.getScriptProperties().getProperty('templateFileId'));
  const sheetIdMap = new Map(copyFromSs.sheets.map(x => [x.properties.title, x.properties.sheetId]));
  const copySheetNames = ['Items', 'Trial', 'Quotation Request'];
  const copySheets = copySheetNames.map(x => copySheet_(ss.ss, sheetIdMap.get(x)));
  [ss.items, ss.trial, ss.quotationRequest] = copySheets;
  ss.ss = Sheets.Spreadsheets.get(ss.ss.spreadsheetId);
  const renameRequests = [
                          [0, templateInfo.get('sheetName')],
                          ...copySheetNames.map((sheetName, idx) => [copySheets[idx].sheetId, sheetName]),
                         ].map(x => spreadSheetBatchUpdate.editRenameSheetRequest(x[0], x[1]));  
  editTrialTerm_(inputData);
  const setTrialRequest = setTrialSheet_(inputData, ss.trial.sheetId);
  const quotationRequestRequests = [
    spreadSheetBatchUpdate.getRangeSetValueRequest(ss.quotationRequest.sheetId, 
                                                   1, 
                                                   0, 
                                                   [[Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd')]]),
  ];
  const setItemsRequest = setItemsSheet_(ss.ss, ss.items);
  const createTemplateRequest = createTemplate_(ss.ss, ss.template, ss.items);
  const requestsArray = [
                         ...renameRequests,
                         ...setTrialRequest,
                         ...quotationRequestRequests,
                         ...setItemsRequest,
                         ...createTemplateRequest,
                        ];
  const requests = spreadSheetBatchUpdate.editBatchUpdateRequest(requestsArray);
  spreadSheetBatchUpdate.execBatchUpdate(requests, ss.ss.spreadsheetId);
  return;

  return;
  const targetYearsSheet = copyTemplate_(newSs, templateSheet);
  const targetYears = targetYearsSheet.map(x => x.getName());
  new CreateTotalSheet(newSs, targetYears, templateSheet).exec('Total');
  new CreateTotal2Sheet(newSs, targetYears, templateSheet).exec();
  templateSheet.hideSheet();
  setPropertiesByTrialType_(inputData);
  const setValuesRegistration = new SetValuesRegistrationSheet(inputData);
  targetYears.forEach((year, idx, arr) => {
    const targetSheet = newSs.getSheetByName(year);
    setValuesRegistration.exec_(targetSheet);
    if (idx === 0){
      new SetValuesSetupSheet(inputData).exec_(targetSheet);
    }
    if (idx === arr.length - 1){
      new SetValuesClosingSheet(inputData).exec_(targetSheet);
    }
  });
}
/**
 * 
 */
function setPropertiesByTrialType_(inputData){
  commonInfo.set('investigatorInitiatedTrialFlag', inputData.get(commonInfo.get('trialTypeItemName')) === commonInfo.get('trialType').get('investigatorInitiatedTrial'));
  commonInfo.set('specifiedClinicalTrialFlag', inputData.get(commonInfo.get('trialTypeItemName')) === commonInfo.get('trialType').get('specifiedClinicalTrial'));
  // 事務局運営の設定：企業原資または調整事務局の有無が「あり」または医師主導治験の場合に積む
  const clinicalTrialsOfficeFlag = inputData.get(commonInfo.get('sourceOfFundsTextItemName')) === commonInfo.get('commercialCompany') ||
                                   inputData.get('調整事務局の有無') === 'あり' ||
                                   commonInfo.get('investigatorInitiatedTrialFlag');
  commonInfo.set('clinicalTrialsOfficeFlag', clinicalTrialsOfficeFlag);
}
/**
 * Copy the template sheet by the number of contract years.
 * @param {Object} Spreadsheet object.
 * @param {Object} Sheet object.
 * @return Sheet object.
 */
function copyTemplate_(ss, template){
  const startYear = trialInfo.get('setupStart').getFullYear();
  const endYear = trialInfo.get('trialEnd').getMonth() === 2 ? trialInfo.get('closingEnd').getFullYear() : trialInfo.get('trialEnd').getFullYear();
  const targetYears = [...Array(endYear - startYear + 1)].map((_, idx) => startYear + idx);
  return targetYears.map(year => copyFromTemplate_(ss, template, year, `【見積明細：1年毎(${year}年度)】`));
}
function copyFromTemplate_(ss, template, sheetName, headValue, headAddress='B2'){
  const sheet = template.copyTo(ss);
  sheet.setName(sheetName);
  sheet.getRange(headAddress).setValue(headValue);
  return sheet;
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
  const target = ['医師主導治験', '特定臨床研究'];
  const setupTerm = target.some(x => inputData.get('試験種別') === x) ? 6 : 3; 
  const closingTerm = target.some(x => inputData.get('試験種別') === x) ? 6 : 3; 
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
  const sheetId = items.sheetId;
  const sheetName = items.title;
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
