const commonInfo = getCommonInfo_();
const trialInfo = getTrialsInfo_();
const itemsInfo = getItemsInfo_();
const templateInfo = getTemplateInfo_();
const registrationDivisionInfo = getRegistrationDivisionInfo_();
function myFunction(){
  const test = new Map([
                   [commonInfo.get('trialTypeItemName'), '医師主導治験'],
                   ['目標症例数', 30],
                   [commonInfo.get('facilitiesItemName'), 20],
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
                   ['治験薬管理', 'あり'],
                   ['保険料', 1000000],
                   ['試験開始準備費用', 5000],
                   ['症例登録毎の支払', 10000],
                   ['治験薬運搬', 'あり'],
                   ['監査対象施設数', 6],
                   ['1例あたりの実地モニタリング回数', 5],
                   ['年間1施設あたりの必須文書実地モニタリング回数', 1],
                 ]);
  createSpreadsheet(test);
}
