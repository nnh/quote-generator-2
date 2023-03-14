const commonInfo = getCommonInfo_();
const trialInfo = getTrialsInfo_();
const itemsInfo = getItemsInfo_();
const templateInfo = getTemplateInfo_();
function myFunction(){
  const test = new Map([
                   [commonInfo.get('trialTypeItemName'), '特定臨床研究'],
                   ['目標症例数', 500],
                   ['施設数', 600],
                   ['CRF項目数', 700],
                   ['症例登録開始年', '2020年'],
                   ['症例登録開始月', '10月'],
                   ['試験終了年', '2022年'],
                   ['試験終了月', '9月'],
                   [commonInfo.get('sourceOfFundsTextItemName'), '公的資金（税金由来）'],
                   ['調整事務局の有無', 'なし'],
                 ]);
  testCreateSs(test);
}
