/**
 * Set the properties.
 * @param {Object} inputData Map object of the information entered from the form.
 * @return {boolean}
 */
function setPropertiesByInputData_(inputData){
  if (!editTrialTerm_(inputData)){
    return false;
  }
  if (!setPropertiesByTrialType_(inputData)){
    return false;
  } 
  return true;
}
/**
 * Calculate the contract start end date.
 * @param {Object} inputData Map object of the information entered from the form.
 * @return {boolean}
 */
function editTrialTerm_(inputData){
  try{
    const trialStartYear = inputData.get(`${trialInfo.get('trialStartText')}${trialInfo.get('yearText')}`).replace(trialInfo.get('yearText'), '');
    const trialStartMonth = inputData.get(`${trialInfo.get('trialStartText')}${trialInfo.get('monthText')}`).replace(trialInfo.get('monthText'), '');
    const trialEndYear = inputData.get(`${trialInfo.get('trialEndText')}${trialInfo.get('yearText')}`).replace(trialInfo.get('yearText'), '');
    const trialEndMonth = inputData.get(`${trialInfo.get('trialEndText')}${trialInfo.get('monthText')}`).replace(trialInfo.get('monthText'), '');
    const target = [commonInfo.get('trialType').get('investigatorInitiatedTrial'), commonInfo.get('trialType').get('specifiedClinicalTrial')];
    const setupTerm = target.some(x => inputData.get(commonInfo.get('trialTypeItemName')) === x) ? 6 : 3; 
    const closingTerm = target.some(x => inputData.get(commonInfo.get('trialTypeItemName')) === x) ? 6 : 3; 
    const trialStart = new Date(trialStartYear, trialStartMonth - 1, 1);
    const trialEnd = new Date(trialEndYear, Number(trialEndMonth), 0);
    if (trialStart >= trialEnd){
      throw new Error('The end date must be after the start date.');
    }
    const setupStart = new Date(trialStart.getFullYear(), trialStart.getMonth() - setupTerm, trialStart.getDate());
    const closingEnd = new Date(trialEnd.getFullYear(), trialEnd.getMonth() + closingTerm + 1, 0);
    trialInfo.set('trialStart', trialStart);
    trialInfo.set('trialEnd', trialEnd);
    trialInfo.set('setupStart', setupStart);
    trialInfo.set('closingEnd', closingEnd);
    trialInfo.set('setupTerm', setupTerm);
    trialInfo.set('closingTerm', closingTerm);
    trialInfo.set('cases', inputData.get('目標症例数'));
    trialInfo.set('facilities', inputData.get(commonInfo.get('facilitiesItemName')));
  } catch (error){
    return error;
  }
  return true;
}
/**
 * Configure settings for each TrialType.
 * @param {Object} inputData The map object of the information entered from the form.
 * @return {boolean}
 */
function setPropertiesByTrialType_(inputData){
  try{
    commonInfo.set('investigatorInitiatedTrialFlag', inputData.get(commonInfo.get('trialTypeItemName')) === commonInfo.get('trialType').get('investigatorInitiatedTrial'));
    commonInfo.set('specifiedClinicalTrialFlag', inputData.get(commonInfo.get('trialTypeItemName')) === commonInfo.get('trialType').get('specifiedClinicalTrial'));
    // Establishment of secretariat operation: Add if the presence of company source or coordinating secretariat is "Yes" or if it is an investigator-initiated clinical trial.
    const clinicalTrialsOfficeFlag = inputData.get(commonInfo.get('sourceOfFundsTextItemName')) === commonInfo.get('commercialCompany') ||
                                   inputData.get('調整事務局の有無') === 'あり' ||
                                   commonInfo.get('investigatorInitiatedTrialFlag');
    commonInfo.set('clinicalTrialsOfficeFlag', clinicalTrialsOfficeFlag);
  } catch (error){
    return false;
  }
  return true;
}
