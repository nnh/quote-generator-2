/**
 * Set the Trial sheet values from the information entered.
 * @param {Object} inputData Map object of the information entered from the form.
 * @param {number} sheetId
 * @param {Object} ss The spreadsheet object.
 * @return {Object} request object.
 */
function setTrialSheet_(inputData, sheetId, ss){
  const monthsCount = (trialInfo.get('closingEnd').getFullYear() * 12 + trialInfo.get('closingEnd').getMonth())
                       - (trialInfo.get('setupStart').getFullYear() * 12 + trialInfo.get('setupStart').getMonth());
  let commentList = spreadSheetBatchUpdate.rangeGetValue(ss.spreadsheetId, `${trialInfo.get('sheetName')}!B${getNumber_(trialInfo.get('commentStartRowIdx'))}:B${getNumber_(trialInfo.get('commentEndRowIdx'))}`, 'FORMULA')[0].values;
  const crfItemName = commonInfo.get('crfItemName');
  const targetItems = [commonInfo.get('trialTypeItemName'), commonInfo.get('casesItemName'), commonInfo.get('facilitiesItemName'), crfItemName];  
  // If there is CDISC compliance, multiply the number of CRF items by 7.
  if (inputData.has('CDISC対応')){
    if (inputData.get('CDISC対応') === 'あり'){
      inputData.set(crfItemName, `=${inputData.get(crfItemName)} * 7`);
      commentList = commentList.map(
        x => x[0] === '="CRFのべ項目数を一症例あたり"&$B$30&"項目と想定しております。"' 
          ?['="CDISC SDTM変数へのプレマッピングを想定し、CRFのべ項目数を一症例あたり"&$B$30&"項目と想定しております。"']
          : x
      );
    }
  }
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
                                                      `=datedif(D${getNumber_(trialInfo.get('trialTermsAddress').get('rowIdx'))},(E${getNumber_(trialInfo.get('trialTermsAddress').get('rowIdx'))}+1), "m")`
                                                    ]
                                                   ]),
    spreadSheetBatchUpdate.getRangeSetValueRequest(sheetId, 
                                                   trialInfo.get('commentEndRowIdx') + 1, 
                                                   1, 
                                                   targetItemValues),
    spreadSheetBatchUpdate.getRangeSetValueRequest(sheetId, 
                                                   43, 
                                                   1, 
                                                   [[sourceOfFunds]]),
    spreadSheetBatchUpdate.getRangeSetValueRequest(sheetId, 
                                                   trialInfo.get('commentStartRowIdx'), 
                                                   1, 
                                                   commentList),
  ];
  return requests;
}
