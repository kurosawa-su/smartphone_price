/**
 * Google Apps Scriptのメイン設定値（定数）
 * Y!mobile「ヤフー店」の端末価格・在庫取得スクリプト全体で使用される設定値を定義します。
 */
const YMY_SETTINGS = {
  // データを書き込む対象のシート名 (ヤフー店であることを明記)
  SHEET_NAME: 'Y!mobileヤフー店端末一覧',

  // ヘッダー行 (Plan H 最終版: 7列構成)
  HEADER_ROW: [
    [
      '機種名',             // A列
      '容量',               // B列
      '在庫',               // C列
      '端末価格',           // D列
      '割引後価格',         // E列
      '返却価格',           // F列
      '状態'                // G列
    ]
  ],

  // データ書き込み開始行 (1行目はヘッダーなので2行目から)
  START_ROW: 2,
  // データ書き込み開始列 (A列から)
  START_COL: 1,

  // --- ▼ データソースの統合 (Plan H) ▼ ---
  // データソース1: Y!mobile公式API (価格・商品情報用)
  JSON_API_URLS: [
    { name: 'iPhone', url: 'https://www.ymobile.jp/lineup/common/json/v2/iphone.json' },
    { name: 'Android', url: 'https://www.ymobile.jp/lineup/common/json/v2/android.json' }
  ],
  
  // データソース2: Y!mobileヤフー店 在庫API (在庫情報用)
  STOCK_API_URL: 'https://reference-stock.api-ymobile.net/stock01.json',
  // --- ▲ データソース ▲ ---

  // Webサイト側からのブロックを避けるため宣言するUser-Agentヘッダー
  USER_AGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',

  // JSON内の販売種別キーの日本語変換マップ
  SALES_TYPE_MAP: {
    new: '新規',
    mnp: 'MNP',
    number: '番号移行',
    change: '機種変更'
  },

  // JSON内のプランキーの日本語変換マップ
  PLAN_TYPE_MAP: {
    plan_m: 'シンプルM',
    plan_l: 'シンプルL'
  }
};

/**
 * 複数の製品コードを結合してJSON APIの完全なURLを生成します。（本バージョンでは未使用）
 * @param {string[]} productCodes - 対象の製品コードの配列。
 * @returns {string} 完全なJSON APIのURL。
 */
function ymy_getJsonApiUrl(productCodes) {
  return YMY_SETTINGS.JSON_API_URLS[0].url;
}

/**
 * メイン処理：Y!mobileヤフー店の価格・在庫情報を取得し、シートに書き込みます。
 * 2つのAPIをSKUでマージし、最安値プランを抽出する Plan H 方式で実行します。
 */
function メイン処理_Ymobile_Yahoo店端末価格取得() {
  Logger.log('メイン処理_Y!mobile-Yahoo!店端末価格取得 (Plan H) を開始します。');
  const SETTINGS = YMY_SETTINGS;
  let allOutputData = [];

  try {
    // Plan H ステップ1: Yahoo!在庫APIから在庫マップを作成
    Logger.log('Plan H ステップ1: Yahoo!在庫APIから在庫マップを作成します。');
    const stockJson = ymy_fetchJsonData(SETTINGS.STOCK_API_URL);
    const stockMap = stockJson.stocks || {};
    Logger.log(`在庫マップの作成完了。${Object.keys(stockMap).length} 件のSKU在庫情報を取得。`);

    // Plan H ステップ2: 公式APIから商品・価格情報を取得し、在庫マップとマージ
    for (const api of SETTINGS.JSON_API_URLS) {
      Logger.log(`Plan H ステップ2: [${api.name}] 公式API (${api.url}) から商品・価格情報を取得します。`);
      const jsonObject = ymy_fetchJsonData(api.url);

      Logger.log(`[${api.name}] 取得したJSONデータと在庫マップをマージします。`);
      const extractedData = ymy_extractDataFromJson(jsonObject, stockMap);
      Logger.log(`[${api.name}] 総抽出件数: ${extractedData.length} 件`);

      allOutputData = allOutputData.concat(extractedData);
    }
    Logger.log(`全APIからの統合データ総件数 (集約前): ${allOutputData.length} 件`);

    // 最安値の集約処理
    Logger.log('機種・容量・状態ごとに最安値レコードへの集約を開始します。');
    const aggregatedData = ymy_aggregateLowestPrice(allOutputData);
    Logger.log(`集約後の出力データ件数: ${aggregatedData.length} 件`);
    
    allOutputData = aggregatedData;

    // Plan H ステップ3: スプレッドシート出力処理
    Logger.log(`シート (${SETTINGS.SHEET_NAME}) への書き込みを開始します。`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SETTINGS.SHEET_NAME);

    if (!sheet) {
      Logger.log(`シート「${SETTINGS.SHEET_NAME}」が存在しないため、新規作成します。`);
      sheet = ss.insertSheet(SETTINGS.SHEET_NAME);
    } else {
      Logger.log(`既存シート「${SETTINGS.SHEET_NAME}」を取得しました。`);
      sheet.clearContents();
      Logger.log('既存データをクリアしました。');
    }

    const headerRange = sheet.getRange(
      1,
      SETTINGS.START_COL,
      SETTINGS.HEADER_ROW.length,
      SETTINGS.HEADER_ROW[0].length
    );
    headerRange.setValues(SETTINGS.HEADER_ROW);
    Logger.log('ヘッダー行を書き込みました。');

    if (allOutputData.length > 0) {
      const range = sheet.getRange(
        SETTINGS.START_ROW, // 2行目から
        SETTINGS.START_COL, // A列から
        allOutputData.length, // 行数
        SETTINGS.HEADER_ROW[0].length // 列数
      );

      range.setValues(allOutputData);
      Logger.log(`${allOutputData.length} 件のデータをシートに書き込みました。`);
    } else {
      Logger.log('書き込むデータがありませんでした。');
    }

    ss.toast('Y!mobileヤフー店 端末一覧の更新が完了しました。', '処理完了', 5);

  } catch (err) {
    Logger.log('【ERROR】メイン処理中に重大なエラーが発生しました。');
    Logger.log(`エラーメッセージ: ${err.message}`);
    if (err.stack) { Logger.log(`スタックトレース: ${err.stack}`); }
    SpreadsheetApp.getActiveSpreadsheet().toast(`エラーが発生しました: ${err.message}`, '処理失敗', 10);
  }
  Logger.log('メイン処理_Y!mobile-Yahoo!店端末価格取得 が終了しました。');
}

/**
 * 指定されたURLからJSONデータを取得します。
 */
function ymy_fetchJsonData(url) {
  try {
    Logger.log(`ymy_fetchJsonData: ${url} からデータを取得中...`);
    const params = {
      'method': 'get',
      'headers': { 'User-Agent': YMY_SETTINGS.USER_AGENT },
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, params);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      throw new Error(`データ取得に失敗しました。ステータスコード: ${responseCode}, URL: ${url}`);
    }

    const jsonText = response.getContentText('UTF-8');
    return JSON.parse(jsonText);

  } catch (err) {
    Logger.log(`【ERROR】ymy_fetchJsonData 処理中にエラーが発生しました: ${err.message}`);
    throw err;
  }
}

/**
 * 公式API(JSON)と在庫API(stockMap)をマージし、シート書き込み用の形式に整形します。
 */
function ymy_extractDataFromJson(jsonObject, stockMap) {
  Logger.log('ymy_extractDataFromJson: JSONデータと在庫マップのマージ処理を開始します。');
  const outputData = [];

  try {
    const productList = jsonObject.orders;
    if (!Array.isArray(productList) || productList.length === 0) {
      Logger.log("【WARN】JSONオブジェクトに 'orders' キーが見つからないか、配列が空です。");
      return [];
    }

    // 1. 製品 (機種) レベルのループ
    for (const product of productList) {
      let modelName = product.model_name || ''; 
      const orderId = product.order_id || '';

      // 機種名の整形ロジック
      if (modelName === '' || modelName.includes('_') || /^[a-z0-9]+$/i.test(modelName)) {
          let tempName = (modelName || orderId);
          tempName = tempName.replace(/_used/gi, '');
          tempName = tempName.replace(/_/g, ' ');
          tempName = tempName.replace(/\biphone\b/gi, 'iPhone');
          tempName = tempName.replace(/\bse\b/gi, 'SE');
          modelName = tempName.trim();
      }

      // 2. 容量 (ストレージ) レベルのループ
      const storageList = product.storages;
      if (!Array.isArray(storageList)) continue;

      // 状態（新品/中古）の判定ロジック (order_id に _used が含まれるか)
      let condition = '新品';
      if (orderId.includes('_used')) {
          condition = '中古';
      }

      for (const storage of storageList) {
        let capacity = storage.storage; 
        
        // 容量が取れていない場合（null, undefined, ""）の代替処理を追加
        if (!capacity) {
            const md = storage.md;
            if (md) {
                const capacityMatch = md.match(/(\d+(gb|GB|tb|TB))$/i);
                if (capacityMatch && capacityMatch[1]) {
                    capacity = capacityMatch[1].toUpperCase(); 
                } else {
                    capacity = ''; // 不明の場合は空欄に
                }
            } else {
                capacity = ''; // 不明の場合は空欄に
            }
        }
        capacity = String(capacity); 

        // 容量が数値のみ（例: "128"）の場合、"GB" を付与する
        if (capacity && /^\d+$/.test(capacity)) {
            capacity += 'GB';
        }
        
        const basePrice = storage.price && storage.price.product ? storage.price.product : '';

        // --- 在庫集約ロジック ---
        const itemCodeList = product.item_code;
        let hasStock = false; 
        let hasTextInfo = false; 
        let textInfo = '在庫なし'; 
        
        if (Array.isArray(itemCodeList)) {
          for (const item of itemCodeList) {
            const sku = item.id;
            const stockInfo = stockMap[sku]; 

            if (stockInfo) {
              if (stockInfo.stock > 0) {
                hasStock = true; 
                break; 
              }
              if (stockInfo.text) {
                hasTextInfo = true;
                textInfo = stockInfo.text; 
              }
            }
          }
        } else {
             // SKUリストがない場合は、公式APIの sale_flg にフォールバック
            const officialStock = product.sale_flg === 1 ? '在庫あり' : '在庫なし';
            Logger.log(`【WARN】機種 ${modelName} にSKUリストがないため、公式APIの在庫(${officialStock})を使用します。`);
            if (officialStock === '在庫あり') hasStock = true;
        }
        
        let aggregatedStockStatus = '在庫なし';
        if (hasStock) {
          aggregatedStockStatus = '在庫あり';
        } else if (hasTextInfo) {
          aggregatedStockStatus = textInfo; 
        }
        // ------------------------

        // 3. 価格取得・最安値判定ロジック
        const salesPrices = storage.price;
        if (!salesPrices) continue;

        let minPrice = Infinity;
        let bestPlanData = null;

        const mnpPrices = salesPrices['mnp'];
        if (mnpPrices) {
          for (const planTypeKey in YMY_SETTINGS.PLAN_TYPE_MAP) {
            if (!mnpPrices[planTypeKey]) continue; 

            const planPrices = mnpPrices[planTypeKey];

            const total = planPrices.total; 
            const tokusapo24 = planPrices.tokusapo_24 || '';

            let tokusapoTotalPayment = '';
            if (tokusapo24 !== '') {
                const tokusapoValue = Number(tokusapo24) || 0;
                tokusapoTotalPayment = tokusapoValue * 24;
            }

            const currentPrice = (total !== undefined && total !== null && total !== '') ? Number(total) : Infinity;

            if (currentPrice < minPrice) {
              minPrice = currentPrice;
              bestPlanData = {
                total: total,
                tokusapoTotalPayment: tokusapoTotalPayment
              };
            }
          }
        }

        if (bestPlanData) {
          // データ配列の作成 (ヘッダーの7項目に対応)
          outputData.push([
            modelName,              // 0: 機種名
            capacity,               // 1: 容量
            aggregatedStockStatus,  // 2: 在庫
            basePrice,              // 3: 端末価格
            bestPlanData.total,     // 4: 割引後価格
            bestPlanData.tokusapoTotalPayment, // 5: 返却価格
            condition               // 6: 状態
          ]);
        }
      } 
    } 

  } catch (err) {
    Logger.log(`【ERROR】ymy_extractDataFromJson 処理中にエラーが発生しました: ${err.message}`);
    throw err;
  }

  return outputData;
}

/**
 * 抽出されたデータを最安値に集約します。
 */
function ymy_aggregateLowestPrice(rawData) {
    Logger.log('ymy_aggregateLowestPrice: 最安値レコードへの集約処理を開始します。');
    const aggregatedMap = new Map();

    const IDX_MODEL_NAME = 0;
    const IDX_CAPACITY = 1;
    const IDX_STOCK = 2; 
    const IDX_BASE_PRICE = 3; 
    const IDX_TOTAL_PRICE = 4; // 割引後価格
    const IDX_TOKUSAPO_TOTAL = 5; 
    const IDX_STATUS = 6; 

    for (const rawRow of rawData) {
        const modelName = rawRow[IDX_MODEL_NAME];
        const capacity = rawRow[IDX_CAPACITY];
        const status = rawRow[IDX_STATUS];
        
        const totalPriceString = rawRow[IDX_TOTAL_PRICE]; 
        const currentPrice = (totalPriceString !== '' && totalPriceString !== null) ? Number(totalPriceString) : Infinity;

        const key = `${modelName}|${capacity}|${status}`; 

        if (isNaN(currentPrice) || currentPrice === Infinity) {
            Logger.log(`【WARN】集約対象外の価格データが検出されたためスキップ: ${key} (価格が無効: ${totalPriceString})`);
            continue;
        }

        if (!aggregatedMap.has(key) || currentPrice < Number(aggregatedMap.get(key)[IDX_TOTAL_PRICE])) {
            // 現在のレコード（rawRow）を格納する前に、価格列を数値に変換して格納
            const rowToStore = [...rawRow];
            rowToStore[IDX_TOTAL_PRICE] = currentPrice; 
            rowToStore[IDX_BASE_PRICE] = Number(rowToStore[IDX_BASE_PRICE]) || rowToStore[IDX_BASE_PRICE]; 
            
            aggregatedMap.set(key, rowToStore);
            Logger.log(`更新: ${key} を価格 ${currentPrice} で設定`);
        }
    }

    const finalOutput = [];
    for (const rawRow of aggregatedMap.values()) {
        finalOutput.push([
            rawRow[IDX_MODEL_NAME], 
            rawRow[IDX_CAPACITY], 
            rawRow[IDX_STOCK], 
            rawRow[IDX_BASE_PRICE], 
            rawRow[IDX_TOTAL_PRICE], 
            rawRow[IDX_TOKUSAPO_TOTAL], 
            rawRow[IDX_STATUS] 
        ]);
    }
    return finalOutput;
}

/**
 * テスト関数
 */
function ymy_test_parsingLogic() {
  Logger.log('ymy_test_parsingLogic (Plan H 実装) を開始します。');
  
  try {
    const stockJson = ymy_fetchJsonData(YMY_SETTINGS.STOCK_API_URL);
    const stockMap = stockJson.stocks || {};
    Logger.log(`在庫マップ作成完了: ${Object.keys(stockMap).length} 件`);

    const targetApi = YMY_SETTINGS.JSON_API_URLS[0]; // iPhone
    const jsonObject = ymy_fetchJsonData(targetApi.url);

    const extractedData = ymy_extractDataFromJson(jsonObject, stockMap);
    Logger.log(`抽出件数 (iPhone): ${extractedData.length} 件`);

    const aggregatedData = ymy_aggregateLowestPrice(extractedData);
    Logger.log(`集約後件数: ${aggregatedData.length} 件`);
    
    Logger.log(`--- 抽出データ（先頭5件） ---`);
    if (aggregatedData.length > 0) {
      const header = YMY_SETTINGS.HEADER_ROW[0]; 
      
      aggregatedData.slice(0, 5).forEach((row, index) => {
        Logger.log(`[${index + 1}]`);
        Logger.log(`  ${header[0]} (機種名): ${row[0]}`);
        Logger.log(`  ${header[1]} (容量): ${row[1]}`);
        Logger.log(`  ${header[2]} (在庫): ${row[2]}`); 
        Logger.log(`  ${header[3]} (端末価格): ${row[3]}`); 
        Logger.log(`  ${header[4]} (割引後価格): ${row[4]}`); 
        Logger.log(`  ${header[5]} (返却価格): ${row[5]}`); 
        Logger.log(`  ${header[6]} (状態): ${row[6]}`); 
      });
    } else {
      Logger.log('データは抽出されませんでした。');
    }

  } catch (err) {
    Logger.log(`【ERROR】${err.message}`);
  } finally {
    Logger.log('ymy_test_parsingLogic 終了');
  }
}