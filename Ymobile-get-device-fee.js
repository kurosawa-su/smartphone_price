/**
 * Google Apps Scriptのメイン設定値（定数）
 * Y!mobileのiPhoneおよびAndroid価格取得スクリプト全体で使用される設定値を定義します。
 */
const Y_MOBILE_SETTINGS = {
  // データを書き込む対象のシート名
  SHEET_NAME: 'Y!mobile端末一覧',

  // スプレッドシートに出力するヘッダー行の定義 (全7列)
  HEADER_ROW: [
    ['機種名', '容量', '在庫', '端末価格', '割引後価格', '返却価格', '状態']
  ],

  // データ書き込み開始行 (1行目はヘッダーなので2行目から)
  START_ROW: 2,
  // データ書き込み開始列 (A列から)
  START_COL: 1,

  // データソースとなるY!mobile JSON APIのURLを配列で定義
  JSON_API_URLS: [
    { name: 'iPhone', url: 'https://www.ymobile.jp/lineup/common/json/v2/iphone.json' },
    { name: 'Android', url: 'https://www.ymobile.jp/lineup/common/json/v2/android.json' }
  ],

  // Webサイト側からのブロックを避けるため宣言するUser-Agentヘッダー
  USER_AGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',

  // JSON内の販売種別キーの日本語変換マップ
  // フィルタリングのために定義は維持します。
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
function yMobile_getJsonApiUrl(productCodes) {
    // 本バージョンでは未使用だが、関数の定義は残す
    // 互換性のために最初のURLを返す
    return Y_MOBILE_SETTINGS.JSON_API_URLS[0].url;
}

/**
 * メイン処理：Y!mobileのiPhoneおよびAndroid価格情報を取得し、指定されたシートに統合して書き込みます。
 * この関数をGASエディタから手動で実行することを想定しています。
 */
function メイン処理_Ymobile端末価格取得() {
  Logger.log('メイン処理_Ymobile端末価格取得 を開始します。');
  let allOutputData = []; // 全APIからの統合データ
  // ヘッダー行の定義に基づき、列数を取得 (7列)
  const numColumns = Y_MOBILE_SETTINGS.HEADER_ROW[0].length; 

  try {
    // 1. 全APIからのデータ取得と解析
    for (const api of Y_MOBILE_SETTINGS.JSON_API_URLS) {
      Logger.log(`[${api.name}] JSONデータURL (${api.url}) からデータを取得します。`);
      // JSONデータの取得
      const jsonObject = yMobile_fetchJsonData(api.url);

      // データの解析と抽出
      Logger.log(`[${api.name}] 取得したJSONデータから必要な情報を抽出・整形します。`);
      // yMobile_extractDataFromJsonは集約のために一時的に8列のデータ配列を返します
      const extractedData = yMobile_extractDataFromJson(jsonObject); 
      Logger.log(`[${api.name}] 総抽出件数: ${extractedData.length} 件`);

      // 全体データに追加
      allOutputData = allOutputData.concat(extractedData);
    }
    Logger.log(`全APIからの統合データ総件数: ${allOutputData.length} 件`);
    
    // === 最安値の集約処理 ===
    // 抽出された8列のデータから、機種・容量・状態ごとに最安値を選び出し、最終的な7列のデータに整形します
    Logger.log('機種・容量・状態ごとに最安値レコードへの集約を開始します。');
    const aggregatedData = yMobile_aggregateLowestPrice(allOutputData);
    Logger.log(`集約後の出力データ件数: ${aggregatedData.length} 件`);
    
    // 集約後のデータを使用するように置き換え
    allOutputData = aggregatedData;
    // ========================


    // 2. スプレッドシート出力処理
    Logger.log(`シート (${Y_MOBILE_SETTINGS.SHEET_NAME}) への書き込みを開始します。`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(Y_MOBILE_SETTINGS.SHEET_NAME);

    // シートが存在しない場合は新規作成し、存在する場合はデータをクリア
    if (!sheet) {
      Logger.log(`シート「${Y_MOBILE_SETTINGS.SHEET_NAME}」が存在しないため、新規作成します。`);
      sheet = ss.insertSheet(Y_MOBILE_SETTINGS.SHEET_NAME);
    } else {
      Logger.log(`既存シート「${Y_MOBILE_SETTINGS.SHEET_NAME}」を取得しました。`);
      // 既存データをクリアして、常に最新のデータに上書きできるようにします
      sheet.clearContents();
      Logger.log('既存データをクリアしました。');
    }

    // ヘッダー行を書き込み (1行目)
    // 定義したヘッダー配列のサイズに合わせてRangeを取得します
    const headerRange = sheet.getRange(
      1,
      Y_MOBILE_SETTINGS.START_COL,
      Y_MOBILE_SETTINGS.HEADER_ROW.length,
      numColumns
    );
    headerRange.setValues(Y_MOBILE_SETTINGS.HEADER_ROW);
    Logger.log('ヘッダー行を書き込みました。');

    // データの整形とエラー値のハイライト
    if (allOutputData.length > 0) {
      // 抽出データが1件以上ある場合のみ書き込み処理を実行
      const range = sheet.getRange(
        Y_MOBILE_SETTINGS.START_ROW, // 2行目から
        Y_MOBILE_SETTINGS.START_COL, // A列から
        allOutputData.length, // 行数
        numColumns // 列数 (7列)
      );

      // シートへの書き込み
      range.setValues(allOutputData);
      Logger.log(`${allOutputData.length} 件のデータをシートに書き込みました。`);
    } else {
      Logger.log('書き込むデータがありませんでした。');
    }

    // 完了通知
    ss.toast('Y!mobile端末一覧の更新が完了しました。', '処理完了', 5);

  } catch (err) {
    // 主要な処理全体でのエラーハンドリング
    Logger.log('【ERROR】メイン処理中に重大なエラーが発生しました。');
    Logger.log(`エラーメッセージ: ${err.message}`);
    // スタックトレースがあれば出力
    if (err.stack) { Logger.log(`スタックトレース: ${err.stack}`); }
    // ユーザーにシート上でエラー発生を通知
    SpreadsheetApp.getActiveSpreadsheet().toast(`エラーが発生しました: ${err.message}`, '処理失敗', 10);
  }
  Logger.log('メイン処理_Ymobile端末価格取得 が終了します。');
}

/**
 * 指定されたURLからJSONデータを取得します。
 * @param {string} url - 取得対象のJSON API URL。
 * @returns {Object} 取得したパース済みのJSONオブジェクト。
 * @throws {Error} 取得またはパースに失敗した場合。
 */
function yMobile_fetchJsonData(url) {
  try {
    Logger.log(`yMobile_fetchJsonData: ${url} からデータを取得中...`);
    // User-Agentを設定し、HTTPエラーを発生させずにレスポンスを取得
    const params = {
      'method': 'get',
      'headers': { 'User-Agent': Y_MOBILE_SETTINGS.USER_AGENT },
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, params);
    const responseCode = response.getResponseCode();

    // ステータスコードが200 (成功) 以外の場合はエラーとして処理
    if (responseCode !== 200) {
      throw new Error(`データ取得に失敗しました。ステータスコード: ${responseCode}, URL: ${url}`);
    }

    // レスポンスのテキストを取得し、JSONオブジェクトにパース
    const jsonText = response.getContentText('UTF-8');
    // JSON文字列をオブジェクトに変換
    return JSON.parse(jsonText);

  } catch (err) {
    // 取得またはパースエラーをキャッチし、ログに出力
    Logger.log(`【ERROR】yMobile_fetchJsonData 処理中にエラーが発生しました: ${err.message}`);
    // 呼び出し元にエラーを再スロー
    throw err;
  }
}

/**
 * パース済みのJSONオブジェクトから価格データを抽出・解析し、シート書き込み用の形式に整形します。
 * この関数は、MNP/シンプルM/シンプルLにフィルタリングされた、プラン情報と状態情報を含む8列のデータ（一時配列）を返します。
 * @param {Object} jsonObject - 取得したパース済みのJSONオブジェクト。
 * @returns {any[][]} プラン情報と状態情報を含む8列のデータ配列。条件に合うデータがない場合は空の配列を返します。
 * @throws {Error} 解析に失敗した場合。
 */
function yMobile_extractDataFromJson(jsonObject) {
  Logger.log('yMobile_extractDataFromJson: JSONデータ抽出処理を開始します。');
  const outputData = [];

  try {
    // JSONのトップレベルのリストを 'orders' キーから取得 (iPhone/Android共通)
    const productList = jsonObject.orders;
    if (!Array.isArray(productList) || productList.length === 0) {
      Logger.log("【WARN】JSONオブジェクトに 'orders' キーが見つからないか、配列が空です。");
      return [];
    }

    // 1. 製品 (機種) レベルのループ
    for (const product of productList) {
      // 在庫状況の表記を「在庫あり/在庫なし」に修正
      const stock = product.sale_flg === 1 ? '在庫あり' : '在庫なし'; // sale_flgで在庫状況を判断
      let modelName = product.model_name || ''; // 例: iPhone SE (第3世代), null, または空文字列
      const orderId = product.order_id || '';   // order_idを取得

      // === 機種名抽出ロジックの強化 (model_nameが空の場合、order_idから代替抽出) ===
      if (!modelName) {
        if (orderId) {
            // order_idから機種名（例: "iphone12_mini_used" -> "iPhone 12 mini"）を抽出・整形
            // 1. "_used"や"-"をスペースに置き換え
            modelName = orderId
                .replace(/_used/g, '')
                .replace(/_/g, ' ')
                .replace(/-/g, ' ');
                
            // 2. 各単語の先頭を大文字にする
            modelName = modelName.toLowerCase().split(' ').map(word => {
                // 空文字列対策
                if (word.length === 0) return '';
                return word.charAt(0).toUpperCase() + word.slice(1);
            }).join(' ');
            
            // 3. iPhoneなど、連続した大文字を修正し、前後のスペースを削除
            modelName = modelName.replace(/Iphone/g, 'iPhone').trim();
            
            Logger.log(`【INFO】機種名が空欄のため、order_idから代替抽出しました: ${modelName}`);
        } else {
            // order_idも無い場合はスキップ
            Logger.log('【WARN】JSONデータで機種名およびorder_idが空欄のレコードが検出されたためスキップします。');
            continue;
        }
      }
      // ===================================================================

      // === 状態（新品/中古）の判定 (order_idに基づく) ===
      // ユーザー指示: order_idに "_used" が含まれているかどうかで判断
      // 大文字小文字を区別せずチェックして判定
      const isUsed = orderId.toLowerCase().includes('_used');
      const deviceStatus = isUsed ? '中古' : '新品'; 
      // ===========================================

      // 2. 容量 (ストレージ) レベルのループ
      const storageList = product.storages;
      if (!Array.isArray(storageList)) continue;

      for (const storage of storageList) {
        let capacity = storage.storage;
        
        // 容量が取れていない場合（null, undefined, ""）の代替処理を追加
        if (!capacity) {
            const md = storage.md;
            
            // mdフィールドに容量情報が含まれているか確認する代替ロジック
            if (md) {
                // 例: 'iphone14_used_128gb' -> '128gb' を抽出
                // /(\d+(gb|GB|tb|TB))$/i は、文字列の末尾にある '数字 + gb/GB/tb/TB' を抽出する
                const capacityMatch = md.match(/(\d+(gb|GB|tb|TB))$/i);
                if (capacityMatch && capacityMatch[1]) {
                    // 抽出した容量を大文字に変換して設定 (例: 128gb -> 128GB)
                    capacity = capacityMatch[1].toUpperCase(); 
                    Logger.log(`【INFO】機種 ${modelName} の容量をmdフィールドから代替抽出しました: ${capacity}`);
                } else {
                    capacity = ''; // 不明な場合は空欄
                    Logger.log(`【WARN】機種 ${modelName} の容量データがJSONで空または欠損していたため空欄と設定しました。`);
                }
            } else {
                capacity = ''; // 不明な場合は空欄
                Logger.log(`【WARN】機種 ${modelName} の容量データがJSONで空または欠損していたため空欄と設定しました。`);
            }
        }
        
        // 念のため文字列として扱う
        capacity = String(capacity); 

        // === 容量に単位(GB)を付与する処理 ===
        // 容量が空でなく、かつGBやTBという単位が含まれていない場合、末尾にGBを付与する
        if (capacity && !capacity.toUpperCase().includes('GB') && !capacity.toUpperCase().includes('TB')) {
            capacity += 'GB';
        }
        // ===================================
        
        const basePrice = storage.price && storage.price.product ? storage.price.product : ''; // 本体価格(一括/割引前)

        // 3. 販売種別 (new, mnp, number, change) レベルのループ
        const salesPrices = storage.price;
        if (!salesPrices) continue;

        // SALES_TYPE_MAPに定義されたキーのみを処理
        for (const salesTypeKey in Y_MOBILE_SETTINGS.SALES_TYPE_MAP) {
          if (!salesPrices[salesTypeKey]) continue; // その販売種別が存在しない場合はスキップ
          
          // === MNP（Mobile Number Portability）のみにフィルタリング ===
          if (salesTypeKey !== 'mnp') continue; 
          // ==========================================================

          const salesTypeJp = Y_MOBILE_SETTINGS.SALES_TYPE_MAP[salesTypeKey]; // フィルタリングログ用
          const typePrices = salesPrices[salesTypeKey];

          // 4. プラン (plan_m, plan_l のみ) レベルのループ (PLAN_TYPE_MAPに定義されたキーのみ処理)
          for (const planTypeKey in Y_MOBILE_SETTINGS.PLAN_TYPE_MAP) {
            if (!typePrices[planTypeKey]) continue; // そのプランが存在しない場合はスキップ

            const planTypeJp = Y_MOBILE_SETTINGS.PLAN_TYPE_MAP[planTypeKey]; // 集約処理用
            const planPrices = typePrices[planTypeKey];

            // 価格詳細の抽出
            const total = planPrices.total || '';                               // 実質一括価格(割引後)
            const tokusapo24 = planPrices.tokusapo_24 || '';                    // 計算に必要

            // === 返却価格 (トクするサポート) ===
            // 仕様訂正: 月々の支払額(tokusapo_24) × 24回 として算出
            let tokusapoTotalPayment;
            
            if (tokusapo24 === '') {
                tokusapoTotalPayment = '';
            } else {
                const tokusapoValue = Number(tokusapo24) || 0;
                tokusapoTotalPayment = tokusapoValue * 24; 
            }
            // =======================================

            // データ配列の作成 (集約処理のためにプラン情報、状態を含む8列として一時的に出力)
            outputData.push([
              modelName,          // 0: 機種名
              capacity,           // 1: 容量
              stock,              // 2: 在庫状況
              deviceStatus,       // 3: 状態 (集約処理のキーとして保持)
              planTypeJp,         // 4: プラン (集約処理で比較するために保持)
              basePrice,          // 5: 本体価格(一括/割引前) -> 端末価格
              total,              // 6: 実質一括価格(割引後) -> 割引後価格
              tokusapoTotalPayment // 7: トクするサポート支払総額 -> 返却価格
            ]);
            
            Logger.log(`データ抽出: ${modelName} (${capacity}) - ${salesTypeJp}/${planTypeJp}/${deviceStatus}, 実質: ${total}, 返却価格: ${tokusapoTotalPayment}`);
          }
        }
      }
    }

  } catch (err) {
    Logger.log(`【ERROR】yMobile_extractDataFromJson 処理中にエラーが発生しました: ${err.message}`);
    if (err.stack) { Logger.log(`スタックトレース: ${err.stack}`); }
    throw err;
  }

  Logger.log('yMobile_extractDataFromJson: データ抽出処理を終了します。');
  return outputData;
}

/**
 * 抽出されたデータ（機種名、容量、プラン、状態を含む8列）を基に、
 * 機種、容量、状態が同じレコードの中から、最も「実質一括価格(割引後)」が安いものを選び出し、
 * 最終的な7列の形式に整形して返します。
 * * @param {any[][]} rawData - yMobile_extractDataFromJsonから返された8列のデータ配列。
 * @returns {any[][]} 機種・容量・状態ごとに最安値に集約された7列のデータ配列。
 */
function yMobile_aggregateLowestPrice(rawData) {
    Logger.log('yMobile_aggregateLowestPrice: 最安値レコードへの集約処理を開始します。');
    const aggregatedMap = new Map();

    // 抽出されたデータ（8列）のインデックス
    const IDX_MODEL_NAME = 0;
    const IDX_CAPACITY = 1;
    const IDX_STATUS = 3;      // 状態 (集約キーとして使用)
    const IDX_TOTAL_PRICE = 6; // 実質一括価格(割引後)

    for (const row of rawData) {
        const modelName = row[IDX_MODEL_NAME];
        const capacity = row[IDX_CAPACITY];
        const status = row[IDX_STATUS];
        const totalPrice = Number(row[IDX_TOTAL_PRICE]); // 数値として比較
        
        // === 集約キーに「状態」を追加 ===
        const key = `${modelName}|${capacity}|${status}`; // 集約キー
        // ===============================

        // totalPriceが数値として無効な場合はスキップ（空文字列の場合は0となるため、ここでは価格が有効かのみ確認）
        if (row[IDX_TOTAL_PRICE] === '' || isNaN(totalPrice)) {
            Logger.log(`【WARN】集約対象外の価格データが検出されたためスキップ: ${key}`);
            continue;
        }

        // Mapにキーが存在しない場合、または現在のレコードの方が安い場合
        if (!aggregatedMap.has(key) || totalPrice < Number(aggregatedMap.get(key)[IDX_TOTAL_PRICE])) {
            // 現在のレコードを格納
            aggregatedMap.set(key, row);
            Logger.log(`更新: ${key} を価格 ${totalPrice} で設定`);
        }
        // 価格が同じ場合は、先に処理したレコード（シンプルMが優先される想定）を維持します。
    }

    const finalOutput = [];
    
    // 集約されたMapの値を最終的な7列の形式（新しいヘッダー順）に整形
    for (const rawRow of aggregatedMap.values()) {
        finalOutput.push([
            rawRow[0], // 0: 機種名
            rawRow[1], // 1: 容量
            rawRow[2], // 2: 在庫 (在庫状況)
            rawRow[5], // 3: 端末価格 (本体価格(一括/割引前))
            rawRow[6], // 4: 割引後価格 (実質一括価格(割引後))
            rawRow[7], // 5: 返却価格 (トクするサポート支払総額)
            rawRow[3]  // 6: 状態 (状態)
        ]);
    }
    Logger.log('yMobile_aggregateLowestPrice: 集約処理を終了しました。');
    return finalOutput;
}

/**
 * テスト関数：JSONデータの取得と解析ロジックが正しく動作するかログ出力で確認します。
 * 主たる機能の動作確認用として提案します。
 */
function test_yMobileParsingLogic() {
  Logger.log('test_yMobileParsingLogic を開始します。');
  let extractedData = [];

  try {
    // 全てのAPIをテスト
    for (const api of Y_MOBILE_SETTINGS.JSON_API_URLS) {
      Logger.log(`--- [${api.name}] API テスト開始 ---`);
      // 1. JSONデータの取得
      const jsonObject = yMobile_fetchJsonData(api.url);

      // 2. データの解析と抽出
      // この時点ではプラン情報、状態情報を含む8列が出力されます
      extractedData = extractedData.concat(yMobile_extractDataFromJson(jsonObject));
    }
    
    Logger.log(`統合データ総件数 (集約前): ${extractedData.length} 件`);
    
    // 3. 最安値集約のテスト
    const aggregatedData = yMobile_aggregateLowestPrice(extractedData);
    Logger.log(`集約後の総件数 (7列): ${aggregatedData.length} 件`);

    // 抽出データ（先頭5件まで）をログに出力して確認
    Logger.log(`--- 統合抽出データ（集約後、先頭${Math.min(5, aggregatedData.length)}件まで） ---`);
    if (aggregatedData.length > 0) {
      aggregatedData.slice(0, 5).forEach((row, index) => {
        // 集約後の7列のインデックス: 機種名: 0, 割引後: 4, 返却価格: 5, 状態: 6
        Logger.log(`[${index + 1}] 機種名: ${row[0]} (${row[1]} / ${row[6]}), 割引後価格: ${row[4]}, 返却価格: ${row[5]}`);
      });
    } else {
      Logger.log('データは抽出されませんでした。');
    }
    Logger.log('-----------------------------------');

  } catch (err) {
    Logger.log('【ERROR】テスト処理中にエラーが発生しました。');
    Logger.log(`エラーメッセージ: ${err.message}`);
    if (err.stack) { Logger.log(`スタックトレース: ${err.stack}`); }
  } finally {
    Logger.log('test_yMobileParsingLogic が終了しました。');
  }
}