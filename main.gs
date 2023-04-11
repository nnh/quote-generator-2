const commonInfo = getCommonInfo_();
const trialInfo = getTrialsInfo_();
const itemsInfo = getItemsInfo_();
const templateInfo = getTemplateInfo_();
function myFunction(){
  const test = new Map([
                   [commonInfo.get('trialTypeItemName'), '医師主導治験'],
                   ['目標症例数', 30],
                   ['施設数', 20],
                   ['CRF項目数', 500],
                   ['症例登録開始年', '2023年'],
                   ['症例登録開始月', '4月'],
                   ['試験終了年', '2026年'],
                   ['試験終了月', '3月'],
                   [commonInfo.get('sourceOfFundsTextItemName'), '公的資金（税金由来）'],
                   ['調整事務局の有無', 'あり'],
                   ['試験実施番号', 'test-shiken'],
                   ['症例検討会', 'あり'],
                   ['AMED申請資料作成支援', 'あり'],
                   ['PMDA相談資料作成支援', 'なし'],
                   ['統計解析に必要な図表数', 100],
                   ['安全性管理事務局設置', 'あり'],
                   ['効安事務局設置', 'あり'],
                   ['キックオフミーティング', 'あり'],
                 ]);
  createSpreadsheet(test);
}
