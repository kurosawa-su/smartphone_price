/**
 * @fileoverview Apple Store (JP) の製品価格・在庫API(JSON)から情報を取得し、シートに書き込む
 * @version 2.4.1 (2025-12-19)
 * - (v2.4.1) メイン実行関数名を「メイン処理_Apple端末価格取得」に変更。
 * - (v2.4.0) ヘッダー構成を7列（機種名, 容量, 在庫, 端末価格, 割引後価格, 返却価格, 状態）に変更。
 * 「割引後価格」「返却価格」は空欄、「状態」は「新品」固定で出力。
 * - (v2.3.4) 度重なる構文エラー(不要文字 't', 'C')の削除と、不可視Unicode文字の完全クリーンアップ
 * - (v2.3.3) apple_parseProductName のバグ修正（カラーなし製品の対応）と、trim()による堅牢性向上
 * - (v2.3.2) v2.3.1で修正漏れだった構文エラー(不要文字)を削除
 * - (v2.3.1) v2.3.0で混入した不要文字(構文エラーの原因)を削除
 * - (v2.3.0) ご指示に基づき、「機種名＋容量」での集約ロジックを実装。列構成を4列に変更。
 * - (v2.2.0) B案 (全SKU調査) 完了。最終的なハイブリッド方式を実装。
 */

// スクリプト全体で使用する設定値（定数）
const APPLE_SETTINGS = {
  // (API 1: 全製品グループリスト取得API - GET)
  DIGITAL_MAT_URL: 'https://www.apple.com/jp/shop/api/digital-mat?path=library/step0_iphone/digitalmat',
  
  // (API 2: v2.2.0で廃止)
  // PURCHASE_OPTIONS_URL: 'https://www.apple.com/jp/shop/api/purchase-options?fae=true&basePart=',
  
  // (API 3: SKU詳細・製品情報取得API - GET)
  // (v2.2.0) 在庫状況(isBuyable)の取得にのみ使用
  UPDATE_SUMMARY_URL: 'https://www.apple.com/jp/shop/updateSummary?fae=true&product=',

  // (v1.9.0 削除) API 4 は使用しない

  // API負荷軽減のための待機時間 (ミリ秒)
  API_WAIT_MS: 500, // 0.5秒 (API 3ループで使用)

  // 書き込み対象のシート名
  SHEET_NAME: 'apple端末一覧',

  // (v2.4.0) ヘッダー行構成変更
  HEADER_ROW: [['機種名', '容量', '在庫', '端末価格', '割引後価格', '返却価格', '状態']],

  // データ書き込み開始行（ヘッダーが1行目なので、データは2行目から）
  START_ROW: 2,
  // データ書き込み開始列（A列から）
  START_COL: 1,

  // ブロック対策として追加するUser-Agentヘッダー
  USER_AGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
};

/**
 * 共通のAPI取得処理 (GET) - JSONを返す
 * @param {string} url 取得対象のURL
 * @returns {object} APIから返されたJSONオブジェクト
 * @throws {Error} APIの取得またはJSONの解析に失敗した場合
 */
function apple_fetchApi_GET(url) {
  Logger.log(`API呼び出し (GET) を開始します: ${url}`);
  
  // API呼び出しのためのパラメーターを設定
  const params = {
    'method': 'get',
    'headers': {
      // ブロック対策としてUser-Agentを設定
      'User-Agent': APPLE_SETTINGS.USER_AGENT,
    },
    'muteHttpExceptions': true // HTTPエラー時もレスポンスを取得
  };

  let response;
  try {
    // API呼び出しを実行
    response = UrlFetchApp.fetch(url, params);
  } catch (e) {
    // ネットワークエラーなど、fetch自体が失敗した場合のエラーハンドリング
    Logger.log(`【ERROR】API (${url}) の呼び出し自体に失敗しました: ${e.message}`);
    throw new Error(`API呼び出しに失敗しました: ${e.message}`);
  }

  // HTTPステータスコードを取得
  const responseCode = response.getResponseCode();
  // レスポンスボディをUTF-8で取得
  const jsonText = response.getContentText('UTF-8');

  Logger.log(`HTTPステータスコード: ${responseCode}`);

  // 成功 (200 OK) 以外の場合はエラーとして処理
  if (responseCode !== 200) {
    Logger.log(`【ERROR】API (${url}) がエラーを返しました。`);
    Logger.log(`取得したテキスト(一部): ${jsonText.substring(0, 500)}`);
    throw new Error(`API (${url}) の取得に失敗しました。ステータスコード: ${responseCode}`);
  }

  try {
    // JSONとして解析して返す
    return JSON.parse(jsonText);
  } catch (e) {
    // JSON解析失敗のエラーハンドリング
    Logger.log(`【ERROR】JSONの解析に失敗しました: ${e.message}`);
    Logger.log(`取得したテキスト(一部): ${jsonText.substring(0, 500)}`);
    throw new Error(`JSONの解析に失敗しました: ${e.message}`);
  }
}

/**
 * (API 1) digital-mat API から「サブ機種ID」 (例: IPHONE17PRO) と「代表価格」「購入ページURL」のリストを取得します。
 * (v2.2.0 B案調査: JSDoc修正)
 * @returns {object[]} サブ機種ID (basePartId), 機種名 (productName), 代表価格 (priceString), 購入ページURL (buyPageUrl) を含むオブジェクト配列。
 */
function apple_fetchAllProductBaseParts() {
  Logger.log('(API 1) 全iPhoneグループの サブ機種ID (BasePart) および代表価格の取得を開始します...');
  // BasePart IDの重複排除用
  const processedBaseParts = new Set();
  // 最終的に返す、バリエーション情報を含む BasePart の配列
  const productMetaList = [];
  
  try {
    // API 1 を呼び出し
    const response = apple_fetchApi_GET(APPLE_SETTINGS.DIGITAL_MAT_URL);
    
    let digitalMatItems = []; // digitalMatの配列を格納する変数
    
    // JSON構造の最下層を確定させるためのロジック
    if (response && response.body && Array.isArray(response.body.digitalMat)) {
      digitalMatItems = response.body.digitalMat;
      Logger.log(`(API 1) "body.digitalMat" が配列 (${digitalMatItems.length} 件) であることを検出しました。`);
    } else {
      Logger.log('WARN: (API 1) レスポンスのJSON構造が想定外です。body.digitalMat配列が見つかりませんでした。');
      return []; // 処理を中断
    }

    // 取得した機種グループ (digitalMatItems) をループ
    for (const item of digitalMatItems) {
      if (!item || typeof item !== 'object') continue;
      
      // --- (v1.8.0 修正) ---
      // 親ID (item.partNumber) ではなく、familyTypes の中にあるサブ機種ID (omnitureData.partNumber) を抽出する
      
      const familyTypes = Array.isArray(item.familyTypes) ? item.familyTypes : [];
      
      if (familyTypes.length > 0) {
        Logger.log(`(API 1) [${item.partNumber || 'N/A'}] familyTypes (${familyTypes.length} 件) の内部からサブ機種IDを探索中...`);
        
        // familyTypes の各要素をループ
        for (let i = 0; i < familyTypes.length; i++) {
          const familyItem = familyTypes[i];
          if (familyItem && typeof familyItem === 'object') {
            
            // omnitureData内のpartNumber (サブ機種ID) を探索
            const productLink = familyItem.productLink?.link;
            if (productLink && productLink.omnitureData) {
              const omnitureData = productLink.omnitureData;
              
              // (v1.8.0 修正) これが BasePart ID (サブ機種ID) であると断定
              const subBasePartId = omnitureData.partNumber;
              
              // (v1.9.0 修正) API 1 から代表価格 (priceString) も同時に抽出
              const priceString = familyItem?.productPrice?.priceData?.fullPrice?.priceString || 'N/A';
              
              // (v2.0.0 B案調査) API 1 から「購入ページURL」も抽出
              const buyPageUrl = productLink.url || null;
              
              if (subBasePartId && typeof subBasePartId === 'string' && !processedBaseParts.has(subBasePartId)) {
                processedBaseParts.add(subBasePartId);

                // 取得したメタデータをリストに追加
                productMetaList.push({
                  basePartId: subBasePartId, // 例: "IPHONE17PRO"
                  productName: familyItem.productName || item.productName || 'N/A', // 例: "iPhone 17 Pro"
                  priceString: priceString, // (v2.2.0) この代表価格はもう使用しないが、ログ用に残す
                  buyPageUrl: buyPageUrl, // (v2.0.0 B案調査) 例: ".../buy-iphone/iphone-17-pro"
                  // (v1.8.0) カラーと容量は API 2, 3 で取得するため、ここでは取得しない
                  colorList: [],
                  capacityList: [], 
                });
                Logger.log(`(API 1) サブ機種ID [${subBasePartId}] (機種名: ${familyItem.productName}, 価格: ${priceString}, URL: ${buyPageUrl}) を抽出しました。`);
              }
              
            } else {
              // (v1.8.0 修正) ログレベルを WARN から ERROR に変更
              Logger.log(`【ERROR】(API 1) [${item.partNumber || 'N/A'}] familyType #${i} からサブ機種ID (omnitureData.partNumber) が見つかりません。処理をスキップします。`);
            }
          }
        }
      }
    }

    Logger.log(`(API 1) 取得完了。合計 ${productMetaList.length} 件のユニークなサブ機種ID (BasePart) を取得しました。`);
    
    return productMetaList;
    
  } catch (e) {
    Logger.log(`【ERROR】(API 1) digital-mat の処理中にエラー: ${e.message}`);
    return []; 
  }
}

/**
 * (v2.2.0) HTMLスクレイピングとAPI 3を併用し、全SKUの詳細情報を取得します。
 * @param {object[]} productMetaList - API 1で取得した サブ機種ID (basePartId), 機種名 (productName), 購入ページURL (buyPageUrl) を含むオブジェクト配列。
 * @returns {object[]} 全SKUの情報が格納されたオブジェクトの配列。
 */
function apple_fetchSkuDetails(productMetaList) {
  Logger.log(`(HTML+API 3) ${productMetaList.length} 件の機種グループ (buyPageUrl) のスクレイピングを開始します...`);
  const allSkuData = [];
  
  // (v2.2.0) API 1 から取得した機種グループ (buyPageUrl) ごとにループ
  for (const meta of productMetaList) {
    const buyPageUrl = meta.buyPageUrl;
    if (!buyPageUrl) {
      Logger.log(`WARN: [${meta.basePartId}] の buyPageUrl が見つからないため、スキップします。`);
      continue;
    }
    
    try {
      // [ステップA: HTMLスクレイピング]
      // (v2.2.0) 機種ごとの購入ページHTML (600KB+) を取得
      Logger.log(`(HTML 調査) [${meta.basePartId}] ページ (${buyPageUrl}) のHTMLを取得します...`);
      const pageHtml = apple_fetchPage_GET(buyPageUrl);
      Logger.log(`(HTML 調査) [${meta.basePartId}] HTML (全長: ${pageHtml.length}文字) を取得しました。`);
      
      // (v2.2.0) 正規表現で 'id="metrics"' の <script> タグを探す
      // [\s\S]*? は改行を含む任意の文字列（最短一致）
      const regex = /<script type="application\/json" id="metrics">([\s\S]*?)<\/script>/;
      const match = pageHtml.match(regex);

      if (!match || !match[1]) {
        Logger.log(`【ERROR】(HTML 調査) [${meta.basePartId}] 埋め込みJSON (id="metrics") がHTML内から見つかりませんでした。`);
        continue; // この機種グループはスキップ
      }
      
      Logger.log(`(HTML 調査) [${meta.basePartId}] 埋め込みJSON (id="metrics") を発見しました。`);
      
      // 埋め込みJSONをパース
      const metricsJson = JSON.parse(match[1]);
      
      // "data.products" 配列 (全SKUリスト) を取得
      const products = metricsJson?.data?.products;
      
      if (!Array.isArray(products) || products.length === 0) {
        Logger.log(`【ERROR】(HTML 調査) [${meta.basePartId}] 埋め込みJSON内に "data.products" 配列が見つかりませんでした。`);
        continue; // この機種グループはスキップ
      } else {
        Logger.log(`(HTML 調査) [${meta.basePartId}] "data.products" から ${products.length} 件のSKUを発見しました。`);
      }

      // [ステップB: API 3 ループ]
      // 抽出したSKUリスト (例: 21件) をループ
      for (const product of products) {
        const skuId = product.partNumber; // "MG854J/A"
        const price = product.price?.fullPrice || 'N/A'; // 179800.00
        const name = product.name || 'N/A'; // "iPhone 17 Pro 256GB Silver"
        
        if (!skuId || price === 'N/A' || name === 'N/A') {
          Logger.log(`WARN: (HTML 調査) SKU情報が不完全です。SKU: ${skuId}, Price: ${price}, Name: ${name}`);
          continue;
        }
        
        // (v2.2.0) name から 機種名・容量・色 を抽出
        const parsedName = apple_parseProductName(name);
        
        let isBuyable = false; // 在庫状況 (デフォルト: なし)
        
        try {
          // (v2.2.0) API 3 を呼び出し、在庫状況 (isBuyable) のみを取得
          const apiUrl_Api3 = APPLE_SETTINGS.UPDATE_SUMMARY_URL + skuId;
          // API 3 を呼び出し
          const responseApi3 = apple_fetchApi_GET(apiUrl_Api3);

          const summary = responseApi3?.body?.response?.summarySection?.summary; 
          
          if (!summary) {
            Logger.log(`WARN: (API 3) [${skuId}] summary オブジェクトが取得できませんでした。在庫を「なし」として扱います。`);
          } else {
            isBuyable = summary.isBuyable || false;
          }
          
          // 情報を結合して配列に追加
          allSkuData.push({
            modelName: parsedName.model,
            capacity: parsedName.capacity,
            color: parsedName.color,
            stock: isBuyable, // (v2.3.0) このstock (boolean) を後で集約に使う
            price: price, // (v2.2.0 修正) HTMLから取得した個別価格 (数値)
            skuId: skuId
          });
          
          Logger.log(` (HTML+API 3) SKU [${skuId}] の価格 [${price}] / 在庫 [${isBuyable}] / モデル [${parsedName.model}] を取得しました。`);

        } catch (eApi3) {
          // API 3 が 1件失敗しても処理を続行
          Logger.log(`【ERROR】(API 3) SKU [${skuId}] の在庫取得に失敗: ${eApi3.message}`);
        }
        // API負荷軽減のため待機
        Utilities.sleep(APPLE_SETTINGS.API_WAIT_MS);
      } 
      
    } catch (eHtml) {
      // HTML 調査 (スクレイピング) が失敗しても処理を続行
      Logger.log(`【ERROR】(HTML 調査) [${meta.basePartId}] のHTML解析またはJSON抽出に失敗: ${eHtml.message}`);
    }
    // HTML取得のループ間 (API 1 のループ) にも待機を入れる
    Utilities.sleep(APPLE_SETTINGS.API_WAIT_MS);
  } 
  
  Logger.log(`(HTML+API 3) 全SKU詳細の取得が完了しました。合計 ${allSkuData.length} 件。`);
  return allSkuData;
}

/**
 * (v2.3.0 修正) 在庫のブール値 (true/false) を、日本語のステータス文字列に変換します。
 * (true: SKU単体の在庫、または集約グループの在庫)
 * @param {boolean} hasStock - 在庫の有無 (true/false)
 * @returns {string} 日本語の在庫ステータス ("在庫あり" / "在庫なし")
 */
function apple_mapStockStatus(hasStock) {
  // ご指示に基づき、true の場合のみ "在庫あり"
  return (hasStock === true) ? '在庫あり' : '在庫なし';
}

/**
 * (v2.2.0 新設) "iPhone 17 Pro 256GB Silver" のような文字列を解析し、構造化します。
 * (v2.3.3 修正) カラーがない製品名 ("iPhone 16 128GB") に対応し、trim()を追加
 * @param {string} productName - HTMLの 'name' フィールドから取得した製品名
 * @returns {{model: string, capacity: string, color: string}} 解析されたオブジェクト
 */
function apple_parseProductName(productName) {
  // "iPhone 17 Pro Max 256GB Silver"
  // "iPhone 16 128GB (PRODUCT)RED"
  // "iPhone 16 128GB" (v2.3.3 対応ケース)
  
  const parts = productName.split(' ');
  
  // 簡易的な解析ロジック
  // (v2.3.3 修正) 最小パーツ数を2 ("iPhone 16") ではなく 3 ("iPhone 16 128GB") と想定
  if (parts.length >= 2) { 
    let modelParts = [];
    let capacity = 'N/A';
    let colorParts = [];
    
    let capacityFound = false;
    for (const part of parts) {
      if (part.includes('GB') || part.includes('TB')) {
        capacity = part.trim(); // (v2.3.3) trim追加
        capacityFound = true;
      } else if (capacityFound) {
        // 容量が見つかった後のパーツはすべて色とする
        colorParts.push(part);
      } else {
        // 容量が見つかる前のパーツはすべてモデル名
        modelParts.push(part);
      }
    }
    
    const model = modelParts.join(' ').trim(); // (v2.3.3) trim追加
    const color = colorParts.join(' ').trim(); // (v2.3.3) trim追加
    
    // (v2.3.3 修正) 「機種名」と「容量」が見つかれば解析成功とする (カラーは空でもOK)
    if (model && capacity !== 'N/A') {
       return {
         model: model,     // "iPhone 17 Pro Max"
         capacity: capacity, // "256GB"
         color: color      // "Silver" または "" (空文字)
       };
    }
  }
  
  // 解析失敗時
  Logger.log(`WARN: (解析) 製品名 "${productName}" の解析に失敗しました。`);
  return {
    model: productName,
    capacity: 'N/A',
    color: 'N/A'
  };
}

/**
 * (v2.3.0 修正) 取得した全SKUデータを「機種名＋容量」で集約し、シート書き込み用の2次元配列に変換します。
 * (v2.3.3 修正) 集約キー作成時に trim() と null チェックを追加
 * @param {object[]} allSkuData - 結合済みのSKU情報オブジェクトの配列
 * @returns {any[][]} シート書き込み用の2次元配列 (7列)
 */
function apple_parseAndExtractData(allSkuData) {
  Logger.log('データ解析と「機種名＋容量」での集約を開始します...');
  
  // 集約用オブジェクト
  // キー: "機種名::容量" (例: "iPhone 17 Pro::256GB")
  // 値: { modelName, capacity, price, hasStock (boolean) }
  const aggregatedData = {};

  // 結合済みのSKUデータをループ
  for (const sku of allSkuData) {
    try {
      // (v2.3.3 修正) 空白やnull値を除去してから集約キーを生成
      const modelNameKey = sku.modelName ? sku.modelName.trim() : 'N/A';
      const capacityKey = sku.capacity ? sku.capacity.trim() : 'N/A';
      const groupKey = `${modelNameKey}::${capacityKey}`;

      if (!aggregatedData[groupKey]) {
        // このグループの最初のSKUの場合
        aggregatedData[groupKey] = {
          modelName: modelNameKey, // (v2.3.3) クリーンなキーを使用
          capacity: capacityKey, // (v2.3.3) クリーンなキーを使用
          // ご指示: 最初の価格を代表価格とする
          price: sku.price, 
          // ご指示: 在庫は「1件でもあれば true」とするため、SKUの在庫(boolean)で初期化
          hasStock: sku.stock 
        };
      } else {
        // 2件目以降のSKUの場合 (在庫状況の更新)
        if (sku.stock === true) {
          // このSKUに在庫があれば、グループの在庫を true (在庫あり) に確定させる
          aggregatedData[groupKey].hasStock = true;
        }
        // 価格は最初のSKUのものを採用するため、更新しない
      }
    } catch (e) {
      Logger.log(`【ERROR】データ解析(集約)中にエラー (SKU: ${sku.skuId}): ${e.message}`);
    }
  } // SKUループの終わり

  Logger.log(`集約が完了しました。 ${Object.keys(aggregatedData).length} 件のグループが見つかりました。`);

  // 最終的なシート出力配列
  const outputData = [];
  
  // 集約済みデータをシート配列に変換
  for (const key in aggregatedData) {
    const group = aggregatedData[key];
    
    // 在庫ステータスを日本語に変換 (hasStock: true/false を "在庫あり"/"在庫なし")
    const stockStatus = apple_mapStockStatus(group.hasStock);
    
    // 価格を日本円の表示形式にフォーマット (例: 179800.00 -> ¥179,800)
    const priceNumber = Number(group.price);
    const formattedPrice = isNaN(priceNumber) ? group.price : `¥${priceNumber.toLocaleString('ja-JP')}`;
    
    // (v2.4.0) 7列の順序で配列に追加
    outputData.push([
      group.modelName,    // 機種名
      group.capacity,     // 容量
      stockStatus,        // 在庫
      formattedPrice,     // 端末価格
      '',                 // 割引後価格 (空欄)
      '',                 // 返却価格 (空欄)
      '新品'              // 状態 (固定値)
    ]);
  }

  Logger.log(`データ解析処理が完了しました。合計 ${outputData.length} 件の行を抽出しました。`);
  return outputData;
}


/**
 * 抽出したデータをスプレッドシートに書き込みます。
 * @param {any[][]} outputData - 書き込むデータ（2次元配列）。
 */
function apple_writeToSheet(outputData) {
  Logger.log(`シート (${APPLE_SETTINGS.SHEET_NAME}) への書き込みを開始します。`);
  try {
    // アクティブなスプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // 指定された名前のシートを取得
    let sheet = ss.getSheetByName(APPLE_SETTINGS.SHEET_NAME);
    
    // シートが存在しない場合
    if (!sheet) {
      Logger.log(`シート「${APPLE_SETTINGS.SHEET_NAME}」が存在しないため、新規作成します。`);
      // 新しいシートを作成
      sheet = ss.insertSheet(APPLE_SETTINGS.SHEET_NAME);
    } else {
      // シートが既に存在する場合
      Logger.log(`既存シート「${APPLE_SETTINGS.SHEET_NAME}」を取得しました。`);
    }

    // シートの既存の内容をすべてクリア
    sheet.clearContents();
    Logger.log('既存データをクリアしました。');

    // ヘッダー行の列数を取得
    const headerCols = APPLE_SETTINGS.HEADER_ROW[0].length;
    // 1行目にヘッダーを書き込み
    sheet.getRange(1, 1, APPLE_SETTINGS.HEADER_ROW.length, headerCols)
      .setValues(APPLE_SETTINGS.HEADER_ROW)
      .setBackground('#f3f3f3') // ヘッダーの背景色
      .setFontWeight('bold'); // ヘッダーを太字に
    Logger.log('ヘッダー行を書き込みました。');

    // 書き込むデータがある場合
    if (outputData.length > 0) {
      // データを指定された開始行から書き込み
      sheet.getRange(APPLE_SETTINGS.START_ROW, APPLE_SETTINGS.START_COL, outputData.length, headerCols)
        .setValues(outputData);
      Logger.log(`${outputData.length} 件のデータをシートに書き込みました。`);
      
      // 書き込んだ列の幅を自動調整
      sheet.autoResizeColumns(APPLE_SETTINGS.START_COL, headerCols);
      
    } else {
      // 書き込むデータがなかった場合
      Logger.log('書き込むデータがありませんでした。');
    }

    // 右下に完了通知を表示
    ss.toast('Apple製品価格一覧の更新が完了しました。', '処理完了', 5);

  } catch (e) {
    // シート書き込み中にエラーが発生した場合
    Logger.log(`【ERROR】シート書き込み中にエラーが発生しました: ${e.message}`);
    Logger.log(`スタックトレース: ${e.stack}`);
    try {
      // エラー通知を表示
      SpreadsheetApp.getActiveSpreadsheet().toast(`シート書き込みエラー: ${e.message}`, '処理失敗', 10);
    } catch (toastErr) {
      // 通知の表示自体に失敗した場合
      Logger.log(`【ERROR】エラー通知(toast)の表示中にさらにエラー: ${toastErr.message}`);
    }
  }
}

/**
 * (v2.0.0 B案調査) 共通のページ取得処理 (GET) - JSON解析を行わない
 * @param {string} url 取得対象のURL
 * @returns {string} APIから返されたテキスト（HTMLなど）
 * @throws {Error} APIの取得に失敗した場合
 */
function apple_fetchPage_GET(url) {
  Logger.log(`ページ取得 (GET) を開始します: ${url}`);
  
  // API呼び出しのためのパラメーターを設定
  const params = {
    'method': 'get',
    'headers': {
      // ブロック対策としてUser-Agentを設定
      'User-Agent': APPLE_SETTINGS.USER_AGENT,
    },
    'muteHttpExceptions': true // HTTPエラー時もレスポンスを取得
  };

  let response;
  try {
    // ページ呼び出しを実行
    response = UrlFetchApp.fetch(url, params);
  } catch (e) {
    // ネットワークエラーなど、fetch自体が失敗した場合のエラーハンドリング
    Logger.log(`【ERROR】ページ (${url}) の呼び出し自体に失敗しました: ${e.message}`);
    throw new Error(`ページ呼び出しに失敗しました: ${e.message}`);
  }

  // HTTPステータスコードを取得
  const responseCode = response.getResponseCode();
  // レスポンスボディをUTF-8で取得
  const responseText = response.getContentText('UTF-8');

  Logger.log(`HTTPステータスコード: ${responseCode}`);

  // 成功 (200 OK) 以外の場合はエラーとして処理
  if (responseCode !== 200) {
    Logger.log(`【ERROR】ページ (${url}) がエラーを返しました。`);
    Logger.log(`取得したテキスト(一部): ${responseText.substring(0, 500)}`);
    throw new Error(`API (${url}) の取得に失敗しました。ステータスコード: ${responseCode}`);
  }
  
  // テキストをそのまま返す
  return responseText;
}


/**
 * メイン処理：Apple製品価格情報を取得し、シートに書き込みます。
 * この関数をGASエディタから手動で実行してください。
 */
function メイン処理_Apple端末価格取得() {
  Logger.log('メイン処理_Apple端末価格取得 を開始します。');
  
  try {
    // 1. (API 1) 全iPhoneの basePart キーと代表価格を取得
    const productMetaList = apple_fetchAllProductBaseParts();
    if (!productMetaList || productMetaList.length === 0) {
      throw new Error('API 1 から対象機種の BasePart メタデータが1件も取得できませんでした。');
    }

    // 2. (v2.2.0) HTMLスクレイピングとAPI 3で全SKUの詳細を取得
    const allSkuData = apple_fetchSkuDetails(productMetaList);
    if (!allSkuData || allSkuData.length === 0) {
      throw new Error('HTMLスクレイピングおよびAPI 3 からSKU情報が1件も取得できませんでした。');
    }

    // 3. (v2.3.3) データを解析・集約
    const outputData = apple_parseAndExtractData(allSkuData);

    // 4. シートに書き込み
    apple_writeToSheet(outputData);

  } catch (err) {
    // メイン処理全体でエラーが発生した場合
    Logger.log('【ERROR】メイン処理中に重大なエラーが発生しました。');
    Logger.log(`エラーメッセージ: ${err.message}`);
    Logger.log(`スタックトレース: ${err.stack}`);
      try {
        // エラー通知を表示
        SpreadsheetApp.getActiveSpreadsheet().toast(`エラーが発生しました: ${err.message}`, '処理失敗', 10);
      } catch (toastErr) {
        // 通知の表示自体に失敗した場合
        Logger.log(`【ERROR】エラー通知(toast)の表示中にさらにエラー: ${toastErr.message}`);
      }
  } finally {
    Logger.log('メイン処理_Apple端末価格取得 が終了しました。');
  }
}