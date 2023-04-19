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
        const targetRange = this.appSheet.getRange(getNumber_(idx), getNumber_(templateInfo.get('colItemNameAndIdx').get('count')));
        const targetValue = targetRange.getValue() !== '' && value[0] !== '' ? targetRange.getValue() + value[0] : value[0];
        targetRange.setValue(targetValue);
      }
    });
  }
  editValues(){
    return;
  }
  getItemNameByTrialType_(trialType, itemList){
    const targetIdx = trialType ? 0 : 1;
    return itemList.map(x => x[targetIdx]);
  }
    setValueOrNull_(itemName, value){
    if (!this.inputData.has(itemName)){
      return null;
    }
    if (!Number.isSafeInteger(this.inputData.get(itemName))){
      return null;
    }
    return this.inputData.get(itemName) > 0 ? value : null;  
  }
}
class SetValuesRegistrationSheet extends SetValuesSheetByYear{
  constructor(inputData, ss){
    super(inputData, ss);
    // interim analysis
    this.interimAnalysisCount = null;
    this.interimYears = null;
    if (Number.isSafeInteger(this.inputData.get('中間解析に必要な図表数'))){
      this.interimAnalysisCount = this.setValueOrNull_('中間解析に必要な図表数', this.inputData.get('中間解析に必要な図表数'));
      const tempInterim = this.inputData.has('中間解析の頻度')
        ? this.inputData.get('中間解析の頻度').split(', ').map(x => x.replace('年', '')).filter(x => /^[0-9]{4}$/.test(x)) : null; 
      this.interimYears = tempInterim 
        ? tempInterim.filter(x => trialInfo.get('registrationStartYear') <= x && x <= trialInfo.get('registrationEndYear')) 
        : null;
      this.interimFirstYear = this.interimYears 
        ? this.interimYears.length > 0 ? this.interimYears[0] : null
        : null;
    }
  }
  getMonthDiff(startDate, endDate){
    const monthUnit = 1000 * 60 * 60 * 24 * 30;
    const res = Math.trunc(Math.abs(endDate - startDate) / monthUnit);
    return res > 0 ? res : null;
  }
  editValues(year){
    const [interimAnalysis, centralMonitoring] = this.getItemNameByTrialType_(commonInfo.get('investigatorInitiatedTrialFlag'), 
      [
        ['中間解析プログラム作成、解析実施（ダブル）', '中間解析プログラム作成、解析実施（シングル）'],
        ['ロジカルチェック、マニュアルチェック、クエリ対応', 'ロジカルチェック、マニュアルチェック、クエリ対応'],
      ]
    );
    const interimAnalysisFlag = this.interimAnalysisCount && this.interimYears.filter(x => x === String(year)).length > 0;
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
    const itemNameAndCount = [
      ['名古屋医療センターCRB申請費用(初年度)', crb 
        ? trialInfo.get('registrationStartYear') === year ? 1 : null 
        : null],
      ['名古屋医療センターCRB申請費用(2年目以降)', crb 
        ? trialInfo.get('registrationStartYear') < year && year <= trialInfo.get('registrationEndYear') && registrationMonth > 6 ? 1 : null
        : null],
      ['治験薬運搬', this.inputData.get('治験薬運搬') === 'あり' 
        ? trialInfo.get('registrationStartYear') <= year && year <= trialInfo.get('registrationEndYear')
           ? this.formulas.get('facilities') 
           : null
        : null],
      ['施設監査費用', this.setValueOrNull_('監査対象施設数', this.getDivisionCount_(this.inputData.get('監査対象施設数'), year))],
      ['症例モニタリング・SAE対応', this.setValueOrNull_('1例あたりの実地モニタリング回数', this.getDivisionCount_(this.inputData.get('1例あたりの実地モニタリング回数') * trialInfo.get('cases'), year))],
      ['開始前モニタリング・必須文書確認', this.setValueOrNull_('年間1施設あたりの必須文書実地モニタリング回数', this.getDivisionCount_(this.inputData.get('年間1施設あたりの必須文書実地モニタリング回数') * trialInfo.get('facilities') * trialInfo.get('registrationYearsCount'), year))],
      // If the interim analysis is performed more than once, set it once for the first year only.
      ['統計解析計画書・出力計画書・解析データセット定義書・解析仕様書作成', interimAnalysisFlag && this.interimFirstYear === year ? 1 : null],
      [interimAnalysis, interimAnalysisFlag ? this.interimAnalysisCount : null],
      ['中間解析報告書作成（出力結果＋表紙）', interimAnalysisFlag ? 1 : null],
      ['データクリーニング', interimAnalysisFlag ? 1 : null],
      ['症例登録', this.setValueOrNull_('症例登録毎の支払', this.getDivisionCount_(trialInfo.get('cases'), year))],
      [centralMonitoring, registrationMonth],
      ['安全性管理事務局業務', this.inputData.get('安全性管理事務局設置') === '設置・委託する' ? registrationMonth : null],
      ['効果安全性評価委員会事務局業務', this.inputData.get('効安事務局設置') === '設置・委託する' ? registrationMonth : null],
      ['事務局運営（試験開始後から試験終了まで）', commonInfo.get('clinicalTrialsOfficeFlag') ? registrationMonth : null],
      ['データベース管理料', registrationMonth],
      ['プロジェクト管理', registrationMonth],
    ];
    return itemNameAndCount;
  }
  getDivisionCount_(inputData, year){
    if (!Number.isSafeInteger(inputData)){
      return null;
    }
    if (year < trialInfo.get('registrationStartYear') || trialInfo.get('registrationEndYear') < year){
      return null;
    }
    const count = Math.floor(inputData / trialInfo.get('registrationYearsCount'));
    return year === trialInfo.get('registrationEndYear') ? inputData - count * (trialInfo.get('registrationYearsCount') - 1) : count;
  }
  setSheetValues_(values){
    const targetRange = this.appSheet.getRange(1, getNumber_(templateInfo.get('colItemNameAndIdx').get('count')), values.length, 1);
    const targetValue = targetRange.getValues().map((x, idx) => x !== '' && values[idx] !== '' ? [x + values[idx]] : [values[idx]]); 
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
      ['事務局運営（試験開始前）', commonInfo.get('clinicalTrialsOfficeFlag') ? trialInfo.get('setupTerm') : null],
      [officeIrbStr, commonInfo.get('investigatorInitiatedTrialFlag') ? this.formulas.get('facilities') : null],
      ['薬剤対応', commonInfo.get('investigatorInitiatedTrialFlag') ? this.formulas.get('facilities') : null],
      ['モニタリング準備業務（関連資料作成）', this.setValueOrNull_('1例あたりの実地モニタリング回数', 1)],
      ['EDCライセンス・データベースセットアップ', 1],
      ['業務分析・DM計画書の作成・CTR登録案の作成', 1],
      ['DB作成・eCRF作成・バリデーション', 1],
      ['バリデーション報告書', 1],
      [setAccounts, this.formulas.get('facilities')],
      ['入力の手引作成', 1],
      ['外部監査費用', this.setValueOrNull_('監査対象施設数', 1)],
      ['保険料', this.setValueOrNull_('保険料', 1)],
      ['治験薬管理（中央）', this.inputData.get('治験薬管理') === 'あり' ? 1 : null],
      ['プロジェクト管理', trialInfo.get('setupTerm')],
      ['試験開始準備費用', this.setValueOrNull_('試験開始準備費用', this.formulas.get('facilities'))],
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
    const finalAnalysisTableCount = this.inputData.get('統計解析に必要な図表数') > 0 
                                    ? commonInfo.get('investigatorInitiatedTrialFlag') && this.inputData.get('統計解析に必要な図表数') < 50 ? 50 : this.inputData.get('統計解析に必要な図表数')
                                    : null;
    const itemNameAndCount = [
      ['ミーティング準備・実行', this.inputData.get('症例検討会') === 'あり' ? 1 : null],
      ['データクリーニング', 1],
      ['事務局運営（試験終了時）', commonInfo.get('clinicalTrialsOfficeFlag') ? 1 : null],
      ['監査対応', commonInfo.get('clinicalTrialsOfficeFlag') && commonInfo.get('investigatorInitiatedTrialFlag') ? 1 : null],
      ['データベース固定作業、クロージング', 1],
      ['症例検討会資料作成', this.inputData.get('症例検討会') === 'あり' ? 1 : null],
      ['統計解析計画書・出力計画書・解析データセット定義書・解析仕様書作成', finalAnalysisTableCount ? 1 :null],
      [finalAnalysis, finalAnalysisTableCount],
      ['最終解析報告書作成（出力結果＋表紙）', finalAnalysisTableCount ? 1 :null],
      [csr, csrCount],
      ['症例報告', this.setValueOrNull_('症例最終報告書提出毎の支払', this.formulas.get('cases'))],
      ['外部監査費用', this.setValueOrNull_('監査対象施設数', 1)],
      ['プロジェクト管理', trialInfo.get('closingTerm')],
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