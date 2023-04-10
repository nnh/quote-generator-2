class SetValuesSheetByYear{
  constructor(inputData, ss){
    this.formulas = new Map([
      ['cases', '=Trial!$B$28'],
      ['facilities', '=Trial!$B$29'],
      ['crfItems', '=Trial!$B$30'],
    ]);
    this.inputData = inputData;
    this.ss = ss;
  }
  exec_(year){
    this.appSheet = SpreadsheetApp.openById(this.ss.spreadsheetId).getSheetByName(year);
    const outputData = this.getRowNumberAndCount_(year);
    this.setSheetValues_(outputData);
    this.appSheet.getRange('B2').setValue(`【見積明細：1年毎(${year}年度)】`);
    SpreadsheetApp.flush();
    const filterRange = spreadSheetBatchUpdate.getRangeGridByIdx(this.appSheet.getSheetId(), templateInfo.get('bodyStartRowIdx') - 1, templateInfo.get('colItemNameAndIdx').get('filter'), null, templateInfo.get('colItemNameAndIdx').get('filter'));
    const filterRequest = spreadSheetBatchUpdate.getBasicFilterRequest(['0'], templateInfo.get('colItemNameAndIdx').get('filter'), filterRange);
    return filterRequest;
  }
  getRowNumberAndCount_(year){
    this.itemNameAndCount = this.editValues(year);
    const [itemNameIdx, countIdx] = [0, 1];
    const itemColIdx = templateInfo.get('colItemNameAndIdx').get('secondaryItem');
    const sheetValues = this.appSheet.getRange(1, 1, this.appSheet.getLastRow(), this.appSheet.getLastColumn()).getValues();
    let outputValues = [...Array(this.appSheet.getLastRow())].map(_ => [null]);
    sheetValues.forEach((rows, idx) => {
      this.itemNameAndCount.forEach(item =>{
        if (item[itemNameIdx] === rows[itemColIdx]){
          outputValues[idx][0] = item[countIdx];
        }
      });
    });
    return outputValues;
  }
  setSheetValues_(values){
    values.forEach((value, idx) => {
      if (value[0]){
        this.appSheet.getRange(getNumber_(idx), getNumber_(templateInfo.get('colItemNameAndIdx').get('count'))).setValue(value[0]);
      }
    });
  }
  editValues(){
    return;
  }
  getItemNameByTrialType_(trialType, itemList){
    const targetIdx = trialType === trialType ? 0 : 1;
    return itemList.map(x => x[targetIdx]);
  }
}
class SetValuesRegistrationSheet extends SetValuesSheetByYear{
  getMonthDiff(startDate, endDate){
    const monthUnit = 1000 * 60 * 60 * 24 * 30;
    return Math.trunc(Math.abs(endDate - startDate) / monthUnit);
  }
  editValues(year){
    const [interimAnalysis, centralMonitoring] = this.getItemNameByTrialType_(commonInfo.get('investigatorInitiatedTrialFlag'), 
      [
        ['中間解析プログラム作成、解析実施（ダブル）', '中間解析プログラム作成、解析実施（シングル）'],
        ['中央モニタリング', '中央モニタリング、定期モニタリングレポート作成'],
      ]
    );
    // Obtain the number of months of registration for the relevant year.
    const targetSheetStartDay = new Date(year, 3, 1);
    const targetSheetEndDay = new Date(parseInt(year) + 1, 2, 31);
    const thisYearStart = trialInfo.get('trialEnd') < targetSheetStartDay 
                          ? null 
                          : targetSheetStartDay < trialInfo.get('trialStart') 
                            ? trialInfo.get('trialStart') 
                            : targetSheetStartDay;
    const thisYearEnd = !thisYearStart 
                        ? null 
                        : targetSheetEndDay < trialInfo.get('trialEnd') 
                          ? targetSheetEndDay 
                          : trialInfo.get('trialEnd');
    const registrationMonth = thisYearStart ? this.getMonthDiff(thisYearStart, thisYearEnd) : null;
    const crb = this.inputData.get('CRB申請') === 'あり';
    const interimAnalysisFlag = this.inputData.get('中間解析業務の依頼') === 'あり';
    const itemNameAndCount = [
      ['名古屋医療センターCRB申請費用(初年度)', crb ? 1 : null],
      ['名古屋医療センターCRB申請費用(2年目以降)', crb ? 1 : null],
      ['治験薬運搬', this.inputData.get('治験薬運搬') === 'あり' ? this.formulas.get('facilities') : null],
      ['開始前モニタリング・必須文書確認', 0],
      ['統計解析計画書・出力計画書・解析データセット定義書・解析仕様書作成', interimAnalysisFlag ? 1 : null],
      [interimAnalysis, interimAnalysisFlag ? this.inputData.get('中間解析に必要な図表数') : null],
      ['中間解析報告書作成（出力結果＋表紙）', interimAnalysisFlag ? 1 : null],
      ['データクリーニング', interimAnalysisFlag ? 1 : null],
      ['症例モニタリング・SAE対応', 0],
      ['施設監査費用', 0],
      ['症例登録', 0],
      [centralMonitoring, registrationMonth],
      ['安全性管理事務局業務', this.inputData.get('安全性管理事務局設置') === 'あり' ? registrationMonth : null],
      ['効果安全性評価委員会事務局業務', this.inputData.get('効安事務局設置') === 'あり' ? registrationMonth : null],
      ['事務局運営（試験開始後から試験終了まで）', commonInfo.get('clinicalTrialsOfficeFlag') ? registrationMonth : null],
      ['データベース管理料', registrationMonth],
      ['プロジェクト管理', 1],
    ];
    return itemNameAndCount;
  }
  setSheetValues_(values){
    const targetRange = this.appSheet.getRange(1, getNumber_(templateInfo.get('colItemNameAndIdx').get('count')), values.length, 1);
    const targetValue = targetRange.getValues().map((_, idx) => values[idx]);
    targetRange.setValues(targetValue);
  }
}
class SetValuesSetupSheet extends SetValuesSheetByYear{
  editValues(){
    const [officeIrbStr, setAccounts] = this.getItemNameByTrialType_(commonInfo.get('investigatorInitiatedTrialFlag'), 
      [
        ['IRB承認確認、施設管理', 'IRB準備・承認確認'],
        ['初期アカウント設定（施設・ユーザー）', '初期アカウント設定（施設・ユーザー）、IRB承認確認'],
      ]
    );  
    const itemNameAndCount = [
      ['プロトコルレビュー・作成支援（図表案、統計解析計画書案を含む）', 1],
      ['検討会実施（TV会議等）', 4],
      ['PMDA相談資料作成支援', this.inputData.get('PMDA相談資料作成支援') === 'あり' ? 1 : null],
      ['AMED申請資料作成支援', this.inputData.get('AMED申請資料作成支援') === 'あり' ? 1 : null],
      ['特定臨床研究法申請資料作成支援', commonInfo.get('specifiedClinicalTrialFlag') ? this.formulas.get('facilities') : null],
      ['ミーティング準備・実行', this.inputData.get('キックオフミーティング') === 'あり' ? 1 : null],
      ['SOP一式、CTR登録案、TMF雛形', commonInfo.get('investigatorInitiatedTrialFlag') ? 1 : null],
      ['事務局運営（試験開始前）', commonInfo.get('clinicalTrialsOfficeFlag') ? 1 : null],
      [officeIrbStr, commonInfo.get('investigatorInitiatedTrialFlag') ? this.formulas.get('facilities') : null],
      ['薬剤対応', commonInfo.get('investigatorInitiatedTrialFlag') ? this.formulas.get('facilities') : null],
      ['モニタリング準備業務（関連資料作成、キックオフ参加）', 0],
      ['EDCライセンス・データベースセットアップ', 1],
      ['業務分析・DM計画書の作成・CTR登録案の作成', 1],
      ['DB作成・eCRF作成・バリデーション', 1],
      ['バリデーション報告書', 1],
      [setAccounts, this.formulas.get('cases')],
      ['入力の手引作成', 1],
      ['外部監査費用', 0],
      ['保険料', 0],
      ['治験薬管理（中央）', this.inputData.get('治験薬管理') === 'あり' ? 1 : null],
    ];
    return itemNameAndCount;
  }
}
class SetValuesClosingSheet extends SetValuesSheetByYear{
  editValues(){
    const [finalAnalysis, csr] = this.getItemNameByTrialType_(commonInfo.get('investigatorInitiatedTrialFlag'), 
      [
        ['最終解析プログラム作成、解析実施（ダブル）', '最終解析プログラム作成、解析実施（シングル）'],
        ['CSRの作成支援', '研究結果報告書の作成'],
      ]
    );
    const csrCount = commonInfo.get('investigatorInitiatedTrialFlag') || this.inputData.get('研究結果報告書作成支援') === 'あり' ? 1 : null;
    const finalAnalysisTableCount = this.inputData.get('最終解析業務の依頼') === 'あり' 
                                    ? commonInfo.get('investigatorInitiatedTrialFlag') && this.inputData.get('統計解析に必要な図表数') < 50 ? 50 : this.inputData.get('統計解析に必要な図表数')
                                    : null;
    const itemNameAndCount = [
      ['ミーティング準備・実行', this.inputData.get('症例検討会') === 'あり' ? 1 : null],
      ['データクリーニング', 1],
      ['事務局運営（試験終了時）', commonInfo.get('clinicalTrialsOfficeFlag') ? 1 : null],
      ['PMDA対応、照会事項対応', commonInfo.get('clinicalTrialsOfficeFlag') ? 1 : null],
      ['監査対応', commonInfo.get('clinicalTrialsOfficeFlag') ? 1 : null],
      ['データベース固定作業、クロージング', 1],
      ['症例検討会資料作成', this.inputData.get('症例検討会') === 'あり' ? 1 : null],
      ['統計解析計画書・出力計画書・解析データセット定義書・解析仕様書作成', this.inputData.get('最終解析業務の依頼') === 'あり' ? 1 :null],
      [finalAnalysis, finalAnalysisTableCount],
      ['最終解析報告書作成（出力結果＋表紙）', this.inputData.get('最終解析業務の依頼') === 'あり' ? 1 :null],
      ['監査対応', 0],
      [csr, csrCount],
      ['症例報告', this.inputData.get('症例最終報告書提出毎の支払') === 'あり' ? setValuesSheetByYear.formulas.get('cases') : null],
      ['外部監査費用', 0],
    ];
    return itemNameAndCount;
  }
}
class SetFilterTotalSheet extends SetValuesSheetByYear{
  exec_(year){
    this.appSheet = SpreadsheetApp.openById(this.ss.spreadsheetId).getSheetByName(year);
    const filterCol = year === commonInfo.get('total2SheetName') ? this.appSheet.getLastColumn() - 1 : templateInfo.get('colItemNameAndIdx').get('filter');
    const filterRange = spreadSheetBatchUpdate.getRangeGridByIdx(this.appSheet.getSheetId(), templateInfo.get('bodyStartRowIdx') - 1, filterCol, null, filterCol);
    const filterRequest = spreadSheetBatchUpdate.getBasicFilterRequest(['0'], filterCol, filterRange);
    return filterRequest;
  }
}