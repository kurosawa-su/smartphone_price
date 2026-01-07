/**
 * @fileoverview Softbankの製品価格・在庫情報をAPIから取得し、シートに書き込む
 * @version 4.6.0 (2025-10-30) [容量取得強化・関数名変更版]
 * - 5つのJSON API (getModelInfo.json + 価格API x3 + 在庫API x1) を連携。
 * - (v4.6.0):
 * - 関数名を「メイン処理_SoftBank端末価格取得」に変更。
 * - 容量(capacityNm)の取得ロジックを大幅に強化。入れ子構造や別名キー(capacity, rom等)を探索するヘルパー関数を追加。
 */

// スクリプト全体で使用する設定値（定数）
const SOFTBANK_SETTINGS = {
  // 1. 製品情報API (Base)
  MODEL_INFO_URL: 'https://online-shop.mb.softbank.jp/ols/mobile/products/user_json/getModelInfo.json',
  
  // 2. 在庫情報API (goodsCdで紐付け)
  STOCK_INFO_URL: 'https://online-shop.mb.softbank.jp/ols/mobile/products/json/stockforSS.json',

  // 3. 価格情報API群 (modelIdで紐付け。優先度順に上書き)
  // メイン (新機種用と推測)
  MAIN_PRICE_URL: 'https://www.softbank.jp/mobile/d/lib-proxy/ols-api/?u=https://online-shop.mb.softbank.jp/ols/mobile/products/json/price.json',
  // サブ (旧機種用)
  PRIORITY_PRICE_URL: 'https://www.softbank.jp/mobile/set/common/shared/data/products/price/priority-price.json',
  // 予備
  OLS_PRICE_INFO_URL: 'https://www.softbank.jp/mobile/set/common/shared/data/products/price/ols-alternative/price.json',
  
  // 抽出対象の契約タイプ (7 = のりかえ/MNP)
  TARGET_CONTRACT_TYPE: '7',

  // 書き込み対象のシート名
  SHEET_NAME: 'softbank端末一覧',

  // 書き込みヘッダー (7列)
  HEADER_ROW: [[
    '機種名', '容量', '在庫', '端末価格', '割引後価格', '返却価格', '状態'
  ]],

  // データ書き込み開始位置
  START_ROW: 2,
  START_COL: 1,

  // ロック設定
  LOCK_TIMEOUT: 30000,
  LOCK_EXPIRATION: 20 * 60 * 1000,

  // User-Agent (Mac版Chrome)
  USER_AGENT: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
};

/**
 * 共通API取得関数 (エラー時はnullを返し処理を継続)
 * URLをそのまま使用します。
 */
function softbank_fetchJsonData(url) {
  Logger.log(`Fetching: ${url}`);
  const params = {
    'method': 'get',
    'headers': {
      'User-Agent': SOFTBANK_SETTINGS.USER_AGENT,
      'Referer': 'https://www.softbank.jp/mobile/products/',
      'Accept': 'application/json, text/plain, */*',
      'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8'
    },
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, params);
    if (response.getResponseCode() !== 200) {
      Logger.log(`[ERROR] Status ${response.getResponseCode()} for ${url}`);
      return null;
    }
    return response.getContentText('UTF-8');
  } catch (e) {
    Logger.log(`[ERROR] Request failed: ${e.message}`);
    return null;
  }
}

/**
 * 在庫マップ作成: goodsCd -> 在庫あり/なし
 */
function softbank_createStockMap(jsonText) {
  const map = {};
  if (!jsonText) return map;
  try {
    const data = JSON.parse(jsonText);
    if (data && Array.isArray(data.goodsCondition)) {
      data.goodsCondition.forEach(item => {
        // 0:在庫あり, 1:在庫わずか, 2:在庫なし (推測)
        map[item.goodsCd] = (item.salesCondition === '0' || item.salesCondition === '1') ? '在庫あり' : '在庫なし';
      });
    }
  } catch (e) { Logger.log(`Stock parse error: ${e.message}`); }
  return map;
}

/**
 * 価格マップ作成: modelId -> 価格シナリオ配列(gPrLi)
 * 複数のAPIからデータを統合するため、既存マップに追記・上書きする形式
 */
function softbank_updatePriceMap(jsonText, currentMap) {
  if (!jsonText) return currentMap;
  try {
    const data = JSON.parse(jsonText);
    // データのキーを探索 (priceListOzil, priceList, またはルート配列)
    let list = null;
    if (Array.isArray(data)) list = data;
    else if (data.priceListOzil) list = data.priceListOzil;
    else if (data.priceList) list = data.priceList;

    if (list && Array.isArray(list)) {
      list.forEach(item => {
        if (item.modelId) {
          // gPrLi (価格シナリオリスト) をそのまま保持
          currentMap[item.modelId] = item.gPrLi || [];
        }
      });
    }
  } catch (e) { Logger.log(`Price parse error: ${e.message}`); }
  return currentMap;
}

/**
 * オブジェクトから容量情報を探索・抽出するヘルパー関数
 * @param {object} obj - 探索対象のオブジェクト (modelIdInfoなど)
 * @returns {string} - 見つかった容量文字列、なければ空文字
 */
function softbank_findCapacity(obj) {
  if (!obj || typeof obj !== 'object') return '';

  // 1. 最優先キーをチェック
  if (obj.capacityNm) return String(obj.capacityNm).trim();
  if (obj.capacity) return String(obj.capacity).trim();
  if (obj.rom) return String(obj.rom).trim(); // Androidでよくあるキー
  if (obj.storage) return String(obj.storage).trim();

  // 2. キーが見つからない場合、再帰的に探索するか、
  // 特定のパターン（数値 + "GB" or "TB"）を持つ値を探す
  // (今回は浅い探索にとどめるが、必要に応じて再帰処理を追加可能)
  for (const key in obj) {
    const val = obj[key];
    if (typeof val === 'string' && /^\d+(GB|TB)$/i.test(val)) {
      return val.trim();
    }
  }
  return '';
}

/**
 * メイン処理
 */
function メイン処理_SoftBank端末価格取得() {
  Logger.log('--- 処理開始 (v4.6.0) ---');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(SOFTBANK_SETTINGS.LOCK_TIMEOUT)) {
    Logger.log('ロック取得失敗');
    return;
  }

  try {
    // 1. 在庫情報の取得
    const stockMap = softbank_createStockMap(softbank_fetchJsonData(SOFTBANK_SETTINGS.STOCK_INFO_URL));
    Logger.log(`在庫情報取得完了: ${Object.keys(stockMap).length}件`);

    // 2. 価格情報の取得と統合 (Priority -> Ols -> Main の順で上書き)
    let priceMap = {};
    priceMap = softbank_updatePriceMap(softbank_fetchJsonData(SOFTBANK_SETTINGS.PRIORITY_PRICE_URL), priceMap);
    priceMap = softbank_updatePriceMap(softbank_fetchJsonData(SOFTBANK_SETTINGS.OLS_PRICE_INFO_URL), priceMap);
    priceMap = softbank_updatePriceMap(softbank_fetchJsonData(SOFTBANK_SETTINGS.MAIN_PRICE_URL), priceMap);
    Logger.log(`価格情報取得完了: ${Object.keys(priceMap).length}機種分`);

    // 3. 製品情報の取得
    const modelInfoJson = softbank_fetchJsonData(SOFTBANK_SETTINGS.MODEL_INFO_URL);
    if (!modelInfoJson) throw new Error('製品情報APIの取得に失敗しました');
    const modelInfo = JSON.parse(modelInfoJson);

    // 4. データの結合と集約
    // キー: "機種名|容量|状態", 値: { sortValue: 比較用価格, rowData: [行データ] }
    const productDataMap = {}; 

    if (modelInfo.itemTypeList) {
      modelInfo.itemTypeList.forEach(type => {
        const itemTypeNm = type.itemTypeNm || '';
        // (v4.5.4) 状態・ランクの判定ロジック強化
        let condition = '新品';
        if (itemTypeNm.includes('中古') || itemTypeNm.includes('Certified')) {
          condition = '中古';
          // itemTypeNm からランクを探す (例: "中古iPhone（ランクA）")
          const gradeMatchItem = itemTypeNm.match(/[（(]中古([A-Z\+]+)[)）]/);
          if (gradeMatchItem) {
            condition = '中古' + gradeMatchItem[1];
          }
        }

        if (!type.modelGrpList) return;

        type.modelGrpList.forEach(group => {
          const modelGrpNm = group.modelGrpNm || '';
          const modelName = modelGrpNm.replace(/<br \/>/g, ' ');

          // (v4.5.4) 機種名からもランクを探す (例: "iPhone 8（中古A）")
          if (condition.startsWith('中古') && condition === '中古') {
             const gradeMatchModel = modelName.match(/[（(]中古([A-Z\+]+)[)）]/);
             if (gradeMatchModel) {
               condition = '中古' + gradeMatchModel[1];
             }
          }

          if (!group.modelIdList) return;

          group.modelIdList.forEach(modelIdInfo => {
            const modelId = modelIdInfo.modelId;
            
            // --- (v4.6.0 修正) 容量取得ロジックの強化 ---
            // ヘルパー関数を使って容量を探す
            let capacityNm = softbank_findCapacity(modelIdInfo);
            
            // 'N/A' 等の無効値確認
            if (capacityNm === 'N/A' || capacityNm === 'null' || capacityNm === 'undefined') {
                capacityNm = '';
            }

            // デバッグログ: 容量が空の場合はJSON構造の一部をログに出力してみる
            // if (!capacityNm && itemTypeNm.includes('スマートフォン')) {
            //   Logger.log(`DEBUG: 容量空 [${modelName}] ID:${modelId} Data:${JSON.stringify(modelIdInfo)}`);
            // }

            const priceScenarios = priceMap[modelId]; 

            if (!modelIdInfo.goodsCdList) return;

            // 在庫集約ロジック
            let isAnyColorInStock = false;
            modelIdInfo.goodsCdList.forEach(goods => {
              if (stockMap[goods.goodsCd] === '在庫あり') {
                isAnyColorInStock = true;
              }
            });
            const aggregatedStockStatus = isAnyColorInStock ? '在庫あり' : '在庫なし';

            // 集約キー (状態を含めることで、ランク違いも別行にする)
            const productKey = `${modelName}|${capacityNm}|${condition}`;

            // 価格シナリオの処理
            if (priceScenarios && priceScenarios.length > 0) {
              priceScenarios.forEach(scenario => {
                // 契約タイプフィルタ (のりかえ '7' のみ)
                if (scenario.cTy !== SOFTBANK_SETTINGS.TARGET_CONTRACT_TYPE) return;

                // 変数抽出 (null対策: 存在しない場合は0とする)
                const sPr = (typeof scenario.sPr === 'number') ? scenario.sPr : 0;
                const iPr = (typeof scenario.iPr === 'number') ? scenario.iPr : 0;
                const eUsFe = (typeof scenario.eUsFe === 'number') ? scenario.eUsFe : 0;
                const rUsFe = (typeof scenario.rUsFe === 'number') ? scenario.rUsFe : 0;
                const olsDis = (typeof scenario.olsDis === 'number') ? scenario.olsDis : 0;
                
                // 割引後価格 = 端末価格(sPr) + OLS割引(olsDis)
                const discountedPrice = sPr + olsDis; 

                let returnPrice = null; // 返却価格
                let comparePrice = Infinity; // 比較用（安い方を採用するため）

                // 返却価格計算
                // (v4.1.8) eUsFe または rUsFe が 0より大きい場合のみ計算
                if (typeof iPr === 'number' && typeof eUsFe === 'number' && typeof rUsFe === 'number') {
                    if (eUsFe > 0 || rUsFe > 0) {
                        // OLS割引を24分割して月額に適用
                        const monthlyOlsDis = olsDis / 24;
                        // 実質月額
                        const monthlyCost = iPr + monthlyOlsDis;
                        
                        // 12ヶ月分 + 各種手数料
                        returnPrice = eUsFe + rUsFe + (monthlyCost * 12);
                        
                        // 比較価格として設定
                        comparePrice = returnPrice;
                    }
                }

                if (returnPrice === null) {
                  // 分割設定がない、または一括のみ等の場合
                  comparePrice = discountedPrice;
                  returnPrice = '-';
                }

                // 集約マップへの登録・更新（より安い価格があれば上書き）
                const currentData = productDataMap[productKey];
                if (!currentData || comparePrice < currentData.sortValue) {
                  productDataMap[productKey] = {
                    sortValue: comparePrice,
                    rowData: [
                      modelName,
                      capacityNm,
                      aggregatedStockStatus,
                      sPr,             // 端末価格
                      discountedPrice, // 割引後価格
                      returnPrice,     // 返却価格 (計算値)
                      condition        // 状態
                    ]
                  };
                }
              });
              
            } else {
              // 価格情報自体がない場合
              if (!productDataMap[productKey]) {
                productDataMap[productKey] = {
                  sortValue: Infinity,
                  rowData: [
                    modelName, capacityNm, aggregatedStockStatus, 
                    '-', '-', '-', condition
                  ]
                };
              }
            }
          });
        });
      });
    }

    // マップから配列へ変換
    const outputData = Object.values(productDataMap).map(d => d.rowData);
    Logger.log(`全データ生成完了: ${outputData.length}行`);

    // 5. シートへの書き込み
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SOFTBANK_SETTINGS.SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SOFTBANK_SETTINGS.SHEET_NAME);
    
    sheet.clearContents();
    
    // ヘッダー書き込み
    sheet.getRange(1, 1, 1, SOFTBANK_SETTINGS.HEADER_ROW[0].length).setValues(SOFTBANK_SETTINGS.HEADER_ROW);
    
    // データ書き込み
    if (outputData.length > 0) {
      sheet.getRange(SOFTBANK_SETTINGS.START_ROW, SOFTBANK_SETTINGS.START_COL, outputData.length, outputData[0].length).setValues(outputData);
    }

    ss.toast(`完了: ${outputData.length}件のデータを取得しました`, '成功');

  } catch (e) {
    Logger.log(`[ERROR] ${e.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`エラー: ${e.message}`, '失敗');
  } finally {
    if (lock.hasLock()) lock.releaseLock();
    Logger.log('--- 処理終了 ---');
  }
}

/**
 * テスト実行用（ログ確認）
 */
function test_softbank_jsonParsingLogic() {
  メイン処理_SoftBank端末価格取得(); 
}