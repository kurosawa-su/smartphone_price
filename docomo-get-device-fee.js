/**
 * @fileoverview ドコモオンラインショップの製品価格・在庫API(JSON)から情報を取得し、シートに書き込む
 * @version 3.8.1 (2025-12-19)
 * - v3.8.1: メイン関数名を `メイン処理_docomo端末価格取得` に変更。
 * - v3.8.0: 【容量取得ロジック刷新】ユーザー提供のスペック比較API (`get-mobile-spec-comparison`) を導入。
 * - API 4から取得した `mobileCode` を使用して新APIを呼び出し、正確なROM容量を取得して容量列に反映するよう修正。
 * - これにより、機種名からの不安定な抽出や表記揺れに依存せず、正確な容量データを出力可能に。
 * - v3.7.0: 【データ整形・表示改善】機種名からの容量分離ロジック強化、N/Aの空欄化。
 * - v3.6.3: 【表示改善】Certified品（中古）の機種名から「docomo Certified」およびランク表記を削除。
 */

// スクリプト全体で使用する設定値（定数）
const DOCOMO_SETTINGS = {
  // (API 1: 認証トークン取得API: GETメソッド)
  TOKEN_API_URL: 'https://onlineshop.docomo.ne.jp/common/auth/auth-anonymus-token-create',

  // (API 2: トランザクションID取得API: GETメソッド)
  TRANSACTION_API_URL: 'https://onlineshop.docomo.ne.jp/common-ui/common/common-info/auth-transaction-id-create',

  // (API 3: 製品情報・価格・在庫API: POSTメソッド)
  API_URL: 'https://onlineshop.docomo.ne.jp/common-ui/ols/get-mobile-price-stock-list',
  
  // (API 4: 端末価格取得API: POSTメソッド)
  TERMINAL_PRICE_API_URL: 'https://onlineshop.docomo.ne.jp/common-ui/ols/get-cart-reserve-if',

  // (API 5: スペック比較API: POSTメソッド) v3.8.0追加
  SPEC_API_URL: 'https://onlineshop.docomo.ne.jp/common-ui/ols/get-mobile-spec-comparison',

  // 書き込み対象のシート名
  SHEET_NAME: 'docomo端末一覧',

  // ヘッダー行
  HEADER_ROW: [['機種名', '容量', '在庫', '端末価格', '割引後価格', '返却価格', '状態']],

  // データ書き込み開始行（ヘッダーが1行目なので、データは2行目から）
  START_ROW: 2,
  // データ書き込み開始列（A列から）
  START_COL: 1,

  // --- APIリクエスト設定 ---
  // (API 3用) 取得対象の契約形態 (03 = のりかえ/MNP)
  TARGET_ORDER_DIV: '03',
  // 並び順 (1 = 新着順)
  TARGET_SORT: '1',

  // ブロック対策として追加するUser-Agentヘッダー
  USER_AGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  
  // 403/401エラー対策の参照元URL
  REFERER_URL: 'https://onlineshop.docomo.ne.jp/products/mobile/price-stock',

  // 403/401エラー対策のオリジンURL
  ORIGIN_URL: 'https://onlineshop.docomo.ne.jp'
};

/**
 * (v1.1.0) 匿名認証トークン(idToken)を取得します。
 */
function docomo_getAuthToken() {
  Logger.log('docomo_getAuthToken: 認証トークン(idToken)とCookieの取得を開始します。');

  // キャッシュを回避するためにタイムスタンプをURLに追加
  const url = DOCOMO_SETTINGS.TOKEN_API_URL + '?nocache=' + new Date().getTime();

  // APIリクエストのパラメータを設定
  const params = {
    'method': 'get',
    'headers': {
      'User-Agent': DOCOMO_SETTINGS.USER_AGENT,
      'Referer': DOCOMO_SETTINGS.REFERER_URL,
      'X-Requested-With': 'XMLHttpRequest',
      'Origin': DOCOMO_SETTINGS.ORIGIN_URL
    },
    'muteHttpExceptions': true
  };

  let response;
  try {
    response = UrlFetchApp.fetch(url, params);
  } catch (e) {
    Logger.log(`【ERROR】認証API (${url}) の呼び出しに失敗しました: ${e.message}`);
    throw new Error(`認証APIの呼び出しに失敗しました: ${e.message}`);
  }

  const responseCode = response.getResponseCode();
  const jsonText = response.getContentText('UTF-8');

  // レスポンスヘッダーから Set-Cookie を取得
  const headers = response.getAllHeaders();
  const setCookieHeaders = headers['Set-Cookie'] || headers['set-cookie'];
  let cookieString = '';

  if (Array.isArray(setCookieHeaders)) {
    cookieString = setCookieHeaders.map(cookie => cookie.split(';')[0]).join('; ');
  } else if (typeof setCookieHeaders === 'string') {
    cookieString = setCookieHeaders.split(';')[0];
  }

  Logger.log(`認証API HTTPステータスコード: ${responseCode}`);
  
  if (responseCode !== 200) {
    Logger.log(`認証API 取得したJSONテキスト(一部): ${jsonText.substring(0, 500)}`);
    throw new Error(`認証API (${url}) の取得に失敗しました。ステータスコード: ${responseCode}`);
  }
  
  try {
    const data = JSON.parse(jsonText);
    if (data.status === 'SUCCESS' && data.result && data.result.idToken && cookieString) { 
      Logger.log('idTokenとCookieの取得に成功しました。');
      return {
        idToken: data.result.idToken,
        cookies: cookieString
      };
    } else {
      Logger.log(`【ERROR】JSONまたはCookieの構造が予期したものではありません: ${jsonText.substring(0, 500)}`);
      throw new Error('idTokenまたはCookieがレスポンスに含まれていません。');
    }
  } catch (e) {
    Logger.log(`【ERROR】認証JSONの解析に失敗しました: ${e.message}`);
    throw new Error(`認証JSONの解析に失敗しました: ${e.message}`);
  }
}

/**
 * (API 2) 認証トークンを使ってトランザクションIDを取得します。
 */
function docomo_getTransactionId(authResult) {
  Logger.log('docomo_getTransactionId: トランザクションIDの取得を開始します。');

  const params = {
    'method': 'get',
    'headers': {
      'User-Agent': DOCOMO_SETTINGS.USER_AGENT,
      'Referer': DOCOMO_SETTINGS.ORIGIN_URL, // v2.1.1 修正済み
      'X-Requested-With': 'XMLHttpRequest',
      'Origin': DOCOMO_SETTINGS.ORIGIN_URL,
      'Authorization': 'Bearer ' + authResult.idToken,
      'Cookie': authResult.cookies,
      'channel-type': '1',
      'lamp-country-region': 'EAST'
    },
    'muteHttpExceptions': true
  };

  let response;
  try {
    response = UrlFetchApp.fetch(DOCOMO_SETTINGS.TRANSACTION_API_URL, params);
  } catch (e) {
    Logger.log(`【ERROR】Transaction API (${DOCOMO_SETTINGS.TRANSACTION_API_URL}) の呼び出しに失敗しました: ${e.message}`);
    throw new Error(`Transaction APIの呼び出しに失敗しました: ${e.message}`);
  }

  const responseCode = response.getResponseCode();
  const jsonText = response.getContentText('UTF-8');

  Logger.log(`Transaction API HTTPステータスコード: ${responseCode}`);

  if (responseCode !== 200) {
    Logger.log(`Transaction API 取得したJSONテキスト(一部): ${jsonText.substring(0, 500)}`);
    throw new Error(`Transaction API (${DOCOMO_SETTINGS.TRANSACTION_API_URL}) の取得に失敗しました。ステータスコード: ${responseCode}`);
  }
  
  try {
    const data = JSON.parse(jsonText);
    if (data.status === 'SUCCESS' && data.result && data.result.transactionId) {
      Logger.log('transactionIdの取得に成功しました。');
      return data.result;
    } else {
      Logger.log(`【ERROR】Transaction JSONの構造が予期したものではありません: ${jsonText.substring(0, 500)}`);
      throw new Error('transactionIdがJSONレスポンスに含まれていません。');
    }
  } catch (e) {
    Logger.log(`【ERROR】Transaction JSONの解析に失敗しました: ${e.message}`);
    throw new Error(`Transaction JSONの解析に失敗しました: ${e.message}`);
  }
}


/**
 * (API 3) APIから指定されたページのJSONデータを取得します (POST)。
 */
function docomo_fetchApiPage(pageNum, authResult) {
  Logger.log(`docomo_fetchApiPage: ページ ${pageNum} の取得を開始します。`);

  const payload = {
    'sort': DOCOMO_SETTINGS.TARGET_SORT,
    'pageNum': pageNum,
    'orderDiv': DOCOMO_SETTINGS.TARGET_ORDER_DIV,
    'category': '',
    'maker': ''
  };

  const params = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'User-Agent': DOCOMO_SETTINGS.USER_AGENT,
      'Referer': DOCOMO_SETTINGS.REFERER_URL,
      'X-Requested-With': 'XMLHttpRequest',
      'Origin': DOCOMO_SETTINGS.ORIGIN_URL,
      'Authorization': 'Bearer ' + authResult.idToken,
      'Cookie': authResult.cookies,
      'channel-type': '1',
      'lamp-country-region': 'EAST'
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  let response;
  try {
    response = UrlFetchApp.fetch(DOCOMO_SETTINGS.API_URL, params);
  } catch (e) {
    Logger.log(`【ERROR】API (${DOCOMO_SETTINGS.API_URL}) の呼び出しに失敗しました: ${e.message}`);
    throw new Error(`API (${DOCOMO_SETTINGS.API_URL}) の呼び出しに失敗しました: ${e.message}`);
  }

  const responseCode = response.getResponseCode();
  const jsonText = response.getContentText('UTF-8');

  Logger.log(`API 3 (Price/Stock) HTTPステータスコード: ${responseCode}`);

  if (responseCode !== 200 && responseCode !== 201) {
    Logger.log(`取得したJSONテキスト(一部): ${jsonText.substring(0, 500)}`);
    throw new Error(`API (${DOCOMO_SETTINGS.API_URL}) の取得に失敗しました。ステータスコード: ${responseCode}`);
  }

  Logger.log(`ページ ${pageNum} の取得に成功しました。`);
  
  try {
    return JSON.parse(jsonText);
  } catch (e) {
    Logger.log(`【ERROR】JSONの解析に失敗しました。ページ: ${pageNum}, 内容(一部): ${jsonText.substring(0, 500)}`);
    throw new Error(`JSONの解析に失敗しました: ${e.message}`);
  }
}

/**
 * (API 3) 全てのページを巡回してAPIデータを取得します。
 */
function docomo_fetchAllApiData(authResult) {
  Logger.log('docomo_fetchAllApiData: 全ページのデータ取得を開始します。');
  const allResponses = [];
  let maxPageCount = 1;

  try {
    const firstPageResponse = docomo_fetchApiPage('1', authResult);
    allResponses.push(firstPageResponse);

    if (firstPageResponse.result && firstPageResponse.result.maxPagingCount) {
      maxPageCount = parseInt(firstPageResponse.result.maxPagingCount, 10);
      Logger.log(`最大ページ数: ${maxPageCount} を確認しました。`);
    } else {
      Logger.log('WARN: maxPagingCount が見つかりません。1ページのみ処理します。');
    }

    if (maxPageCount > 1) {
      for (let i = 2; i <= maxPageCount; i++) {
        Utilities.sleep(500); // 0.5秒待機
        Logger.log(`(${i}/${maxPageCount}) ページ目を取得します...`);
        const pageResponse = docomo_fetchApiPage(String(i), authResult);
        allResponses.push(pageResponse);
      }
    }

    Logger.log(`全 ${allResponses.length} ページ分のデータ取得が完了しました。`);
    return allResponses;

  } catch (e) {
    Logger.log(`【ERROR】docomo_fetchAllApiData 処理中にエラーが発生しました: ${e.message}`);
    if (allResponses.length > 0) {
      Logger.log('エラーが発生しましたが、それまでに取得したデータで処理を続行します。');
      return allResponses;
    } else {
      throw e;
    }
  }
}

/**
 * 在庫フラグを日本語のステータスに変換します。
 */
function docomo_mapStockStatus(saleStockFlag, purchaseFlag) {
  if ((saleStockFlag === '0' || saleStockFlag === '3') && purchaseFlag === '2') {
    return '在庫あり'; // 予約受付中
  }
  if (saleStockFlag === '1' || saleStockFlag === '2') {
    return '在庫あり';
  }
  return '在庫なし';
}

/**
 * 在庫ステータスに優先順位を割り当てます。
 */
function docomo_getStockPriority(stockStatus) {
  switch (stockStatus) {
    case '在庫あり': return 1;
    case '在庫なし': return 2;
    default: return 99;
  }
}

/**
 * (API 3) APIレスポンスの配列を解析し、*集約前*のデータリストに変換します。
 * (v3.7.0) 機種名からの容量分離ロジック強化、N/Aの空欄化
 * @param {object[]} allResponses - 全ページ分のAPIレスポンスオブジェクトの配列。
 * @returns {object[]} 集約前（カラーごと）のデータオブジェクトの配列。
 */
function docomo_parseAndExtractData(allResponses) {
  Logger.log('docomo_parseAndExtractData: データ抽出処理(集約前)を開始します。');
  const parsedData = [];
  const processedItemCodes = new Set();

  for (const response of allResponses) {
    if (!response.result || !Array.isArray(response.result.csOlsLmd04MobileList)) {
      continue;
    }

    const modelList = response.result.csOlsLmd04MobileList;

    for (const model of modelList) {
      let rawModelName = model.mobileNameNoModel || '不明な機種'; // 元の機種名
      let modelName = rawModelName; // 表示用に加工する機種名
      let capacity = model.modelData || ''; // N/A -> 空文字
      const price = model.price || ''; 

      // --- データクリーニング処理 ---
      modelName = modelName.trim();
      let isCertified = modelName.includes('Certified'); 
      modelName = modelName.replace(/docomo[\s　\u00A0]*Certified[\s　\u00A0]*/gi, '');

      // 1. 機種名に容量が含まれている場合の分離
      const nameMatch = modelName.match(/[\s　\u00A0]+(\d{1,4}[GT]B)$/i);
      let capacityFromTitle = null;
      if (nameMatch) {
        capacityFromTitle = nameMatch[1].toUpperCase();
        modelName = modelName.replace(nameMatch[0], '');
      }

      // 2. ランク表記の削除 (Certified品のみ)
      if (isCertified) {
        modelName = modelName.replace(/[\s　\u00A0]+[AB]\+?$/i, '');
      }

      modelName = modelName.trim();

      // 3. 容量フィールドの整形
      if (capacity !== '') {
        const capMatch = capacity.match(/(\d{1,4}[GT]B)/i);
        if (capMatch) {
            capacity = capMatch[1].toUpperCase();
        } else {
            capacity = ''; // 不正な値は空にする
        }
      }

      // 4. 容量の補完 (API 5で上書きされる可能性が高いが、念のため)
      if (capacity === '' && capacityFromTitle) {
        capacity = capacityFromTitle;
      }
      // ------------------------------------

      if (!Array.isArray(model.colorList)) {
        continue;
      }

      for (const color of model.colorList) {
        const itemCode = color.itemCode;
        
        if (!itemCode || processedItemCodes.has(itemCode)) {
          continue;
        }
        processedItemCodes.add(itemCode);

        const stockStatus = docomo_mapStockStatus(color.saleStockFlag, color.purchaseFlag);

        parsedData.push({
          modelName: modelName,
          originalModelName: rawModelName,
          capacity: capacity,
          price: price,
          stockStatus: stockStatus,
          itemCode: itemCode
        });
      }
    }
  }
  Logger.log(`データ抽出処理(集約前)が完了しました。合計 ${parsedData.length} 件のカラーSKUを抽出しました。`);
  return parsedData;
}

/**
 * (v3.8.0) 指定されたmobileCodeのスペック(ROM)を取得します。(API 5)
 * @param {object} authResult - 認証情報
 * @param {string} mobileCode - モバイルコード (例: "004Ly")
 * @returns {string} ROM容量 (例: "256GB") または空文字
 */
function docomo_getSpecRom(authResult, mobileCode) {
  if (!mobileCode) return '';

  const payload = {
    "mobileCodeList": [mobileCode]
  };

  const params = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'User-Agent': DOCOMO_SETTINGS.USER_AGENT,
      'Referer': DOCOMO_SETTINGS.REFERER_URL,
      'X-Requested-With': 'XMLHttpRequest',
      'Origin': DOCOMO_SETTINGS.ORIGIN_URL,
      'Authorization': 'Bearer ' + authResult.idToken,
      'Cookie': authResult.cookies,
      'channel-type': '1',
      'lamp-country-region': 'EAST'
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(DOCOMO_SETTINGS.SPEC_API_URL, params);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200 || responseCode === 201) {
      const jsonText = response.getContentText('UTF-8');
      const data = JSON.parse(jsonText);
      
      // result配列から対象のmobileCodeを探す
      const targetItem = data.result?.find(item => item.mobileCode === mobileCode);
      if (targetItem && Array.isArray(targetItem.mobileSpecInfoList)) {
        // specId: "spec_rom" を探す
        const romSpec = targetItem.mobileSpecInfoList.find(spec => spec.specId === 'spec_rom');
        if (romSpec && Array.isArray(romSpec.valueListCsName)) {
           // 値を結合 (例: value="256", valueUnitBack="GB" -> "256GB")
           // 複数の値がある場合は最初に見つかった有効なものを採用
           for (const valObj of romSpec.valueListCsName) {
             if (valObj.value) {
               return valObj.value + (valObj.valueUnitBack || '');
             }
           }
        }
      }
    }
  } catch (e) {
    Logger.log(`WARN: docomo_getSpecRom (${mobileCode}) - 実行エラー: ${e.message}`);
  }
  
  return '';
}

/**
 * (v3.4.0) 端末価格(定価)、割引後価格、返却価格を取得します。（API 4）
 * (v3.8.0) mobileCode を戻り値に追加
 * @param {object} authResult - 認証情報
 * @param {object} transactionInfo - トランザクション情報
 * @param {string} itemCode - 商品コード
 * @param {string} modelName - 機種名
 * @returns {object} { terminalPrice, discountedPrice, returnPrice, mobileCode }
 */
function docomo_getTerminalPrice(authResult, transactionInfo, itemCode, modelName) {
  const TARGET_TKUBUN = 5; 
  
  const payload = {
    'transactionId': transactionInfo.transactionId,
    'branchNumber': transactionInfo.branchNumber,
    'quick': '1',
    'hflag': '1',
    'scd': itemCode,
    'tkubun': TARGET_TKUBUN,
    'orderDiv': DOCOMO_SETTINGS.TARGET_ORDER_DIV,
  };

  const params = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'User-Agent': DOCOMO_SETTINGS.USER_AGENT,
      'Referer': DOCOMO_SETTINGS.REFERER_URL,
      'X-Requested-With': 'XMLHttpRequest',
      'Origin': DOCOMO_SETTINGS.ORIGIN_URL,
      'Authorization': 'Bearer ' + authResult.idToken,
      'Cookie': authResult.cookies,
      'channel-type': '1',
      'lamp-country-region': 'EAST'
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  const defaultPrice = { terminalPrice: '', discountedPrice: '', returnPrice: '', mobileCode: '' }; 

  try {
    const response = UrlFetchApp.fetch(DOCOMO_SETTINGS.TERMINAL_PRICE_API_URL, params);
    const responseCode = response.getResponseCode();
    const jsonText = response.getContentText('UTF-8'); 
    
    if (responseCode === 200 || responseCode === 201) {
      const data = JSON.parse(jsonText);
      const cartDate = data.result?.cartResult?.csOls20CartDate; 
      const itemType = data.result?.cartResult?.csOls20ItemType; 

      const isCertified = modelName.includes('Certified');
      
      let priceResult = {};
      const getValue = (obj, path) => typeof obj?.[path] !== 'undefined' ? String(obj[path]) : '';
      
      // mobileCodeの取得 (v3.8.0)
      priceResult.mobileCode = getValue(itemType, 'mobileCode');

      // 端末価格 (定価)
      priceResult.terminalPrice = getValue(cartDate, 'itemCashSalePriceBulk'); 
      if (priceResult.terminalPrice === '') {
          priceResult.terminalPrice = getValue(cartDate, 'totalPrice'); 
      }
      
      // 割引後価格
      priceResult.discountedPrice = getValue(cartDate, 'totalPrice');
      
      // 返却価格
      let returnPriceValue = '';
      const programCode = getValue(itemType, 'deviceReturnPgTargetModel');

      if (programCode === '2') {
          returnPriceValue = getValue(cartDate, 'residualBondsPgEarlyReplPrice');
      } else if (programCode === '3') {
          const firstDiscountPrice = Number(getValue(cartDate, 'residualBondsPgFirstDiscountPrice') || 0);
          const monthlyDiscountPrice = Number(getValue(cartDate, 'residualBondsPgDiscountPriceIsm') || 0);
          const calculatedPrice = firstDiscountPrice + (monthlyDiscountPrice * 11);
          
          if (calculatedPrice > 0) {
              returnPriceValue = String(calculatedPrice);
          }
      } else if (programCode === '0') {
          returnPriceValue = ''; 
      } else {
          returnPriceValue = '';
      }
      
      priceResult.returnPrice = returnPriceValue;

      return priceResult;
    } else {
      Logger.log(`WARN: docomo_getTerminalPrice (${itemCode}) - API取得失敗: ${responseCode}`);
      return defaultPrice;
    }
  } catch (e) {
    Logger.log(`WARN: docomo_getTerminalPrice (${itemCode}) - 実行エラー: ${e.message}`);
    return defaultPrice;
  }
}

/**
 * 機種名から端末の状態（グレード）を抽出します。
 * (v3.6.3) 判定には元の機種名を使用。
 * @param {string} modelName - 機種名（例: "docomo Certified iPhone 15 A+"）。
 * @returns {string} 端末の状態（例: "新品", "中古A+", "中古A"）。
 */
function docomo_extractItemGrade(modelName) {
  if (modelName.includes('Certified')) {
    const match = modelName.match(/(\s[A-Z]\+?)$/);
    if (match && match[1]) {
      return '中古' + match[1].trim();
    }
    return '中古';
  }
  return '新品';
}

/**
 * (v3.6.3) 抽出したカラーごとのデータを集約します。
 * (v3.8.0) docomo_getSpecRom を使用して容量を正確なものに更新。
 * @param {object[]} parsedData - docomo_parseAndExtractDataから返されたオブジェクト配列。
 * @returns {any[][]} シート書き込み用の2次元配列。
 */
function docomo_aggregateData(parsedData, authResult, transactionInfo) {
  Logger.log('docomo_aggregateData: データの集約処理を開始します。');
  const aggregationMap = new Map();

  // 1. データをMapに集約
  for (const item of parsedData) {
    const key = item.originalModelName + '::' + item.capacity;
    
    if (!aggregationMap.has(key)) {
      aggregationMap.set(key, {
        modelName: item.modelName,
        originalModelName: item.originalModelName,
        capacity: item.capacity, // API 3由来の容量（不正確な可能性あり）
        price: item.price,
        stockStatusList: [],
        firstItemCode: item.itemCode
      });
    }
    
    aggregationMap.get(key).stockStatusList.push(item.stockStatus);
  }

  Logger.log(`集約キー ${aggregationMap.size} 件のデータを検出しました。在庫ステータスを集約します...`);

  const outputData = [];
  let counter = 0;
  
  // 2. Mapをループし、出力配列を作成
  for (const [key, aggregatedItem] of aggregationMap.entries()) {
    counter++;
    if (counter % 50 === 0) {
       Logger.log(`集約処理 (${counter}/${aggregationMap.size})...`);
    }

    let bestStockStatus = '不明';
    let bestPriority = 99;

    if (aggregatedItem.stockStatusList.length > 0) {
      for (const status of aggregatedItem.stockStatusList) {
        const priority = docomo_getStockPriority(status);
        if (priority < bestPriority) {
          bestPriority = priority;
          bestStockStatus = status;
        }
      }
    } else {
      bestStockStatus = '在庫なし';
    }

    // 4. 価格・mobileCode取得
    const priceResult = docomo_getTerminalPrice(
        authResult, 
        transactionInfo, 
        aggregatedItem.firstItemCode, 
        aggregatedItem.originalModelName 
    );
    
    // (v3.8.0) 正確な容量の取得
    // mobileCodeが取得できていれば、新APIからスペック(ROM)を取得してcapacityを上書きする
    let finalCapacity = aggregatedItem.capacity;
    if (priceResult.mobileCode) {
        // 通信回数削減のため、少し待機を入れる
        Utilities.sleep(100); 
        const specRom = docomo_getSpecRom(authResult, priceResult.mobileCode);
        if (specRom) {
            finalCapacity = specRom;
            // Logger.log(`INFO: 容量をAPI更新 (${aggregatedItem.capacity} -> ${finalCapacity}) 機種:${aggregatedItem.modelName}`);
        }
    }

    // 5. 状態抽出
    const itemGrade = docomo_extractItemGrade(aggregatedItem.originalModelName);

    // 6. シート出力
    outputData.push([
      aggregatedItem.modelName,   // 1. 機種名
      finalCapacity,              // 2. 容量 (v3.8.0: 正確な値に更新)
      bestStockStatus,            // 3. 在庫
      priceResult.terminalPrice,  // 4. 端末価格
      priceResult.discountedPrice,// 5. 割引後価格
      priceResult.returnPrice,    // 6. 返却価格
      itemGrade                   // 7. 状態
    ]);
  }

  Logger.log(`データ集約処理が完了しました。合計 ${outputData.length} 件の行を作成しました。`);
  return outputData;
}


/**
 * 抽出したデータをスプレッドシートに書き込みます。
 */
function docomo_writeToSheet(outputData) {
  Logger.log(`シート (${DOCOMO_SETTINGS.SHEET_NAME}) への書き込みを開始します。`);
  
  SpreadsheetApp.flush();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DOCOMO_SETTINGS.SHEET_NAME);
  
  if (!sheet) {
    Logger.log(`シート「${DOCOMO_SETTINGS.SHEET_NAME}」が存在しないため、新規作成します。`);
    sheet = ss.insertSheet(DOCOMO_SETTINGS.SHEET_NAME);
  } else {
    Logger.log(`既存シート「${DOCOMO_SETTINGS.SHEET_NAME}」を取得しました。データをクリアします。`);
    try {
      sheet.clear();
      SpreadsheetApp.flush();
    } catch (e) {
      Logger.log(`WARN: シートのクリアに失敗しました (${e.message})。シートを再作成します。`);
      try { ss.deleteSheet(sheet); } catch (delErr) {}
      SpreadsheetApp.flush();
      sheet = ss.insertSheet(DOCOMO_SETTINGS.SHEET_NAME);
    }
  }

  try {
    const headerCols = DOCOMO_SETTINGS.HEADER_ROW[0].length;
    sheet.getRange(1, 1, DOCOMO_SETTINGS.HEADER_ROW.length, headerCols)
         .setValues(DOCOMO_SETTINGS.HEADER_ROW)
         .setBackground('#f3f3f3')
         .setFontWeight('bold');
    Logger.log('ヘッダー行を書き込みました。');

    if (outputData.length > 0) {
      sheet.getRange(DOCOMO_SETTINGS.START_ROW, DOCOMO_SETTINGS.START_COL, outputData.length, headerCols)
           .setValues(outputData);
      Logger.log(`${outputData.length} 件のデータをシートに書き込みました。`);
      
      try {
        sheet.autoResizeColumns(DOCOMO_SETTINGS.START_COL, headerCols);
      } catch (resizeError) {
        Logger.log(`WARN: 列幅調整に失敗しましたが、データは書き込まれました: ${resizeError.message}`);
      }
      
    } else {
      Logger.log('書き込むデータがありませんでした。');
    }

    ss.toast('ドコモ製品価格一覧の更新が完了しました。', '処理完了', 5);

  } catch (e) {
    Logger.log(`【ERROR】シート書き込み中にエラーが発生しました: ${e.message}`);
    Logger.log(`スタックトレース: ${e.stack}`);
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(`シート書き込みエラー: ${e.message}`, '処理失敗', 10);
    } catch (toastErr) {
      Logger.log(`【ERROR】エラー通知の表示中にさらにエラー: ${toastErr.message}`);
    }
  }
}

/**
 * メイン処理
 */
function メイン処理_docomo端末価格取得() {
  Logger.log('メイン処理_docomo端末価格取得 を開始します。');
  
  try {
    const authResult = docomo_getAuthToken();
    const transactionInfo = docomo_getTransactionId(authResult);
    const allApiResponses = docomo_fetchAllApiData(authResult);
    const parsedData = docomo_parseAndExtractData(allApiResponses);
    const outputData = docomo_aggregateData(parsedData, authResult, transactionInfo);
    docomo_writeToSheet(outputData);

  } catch (err) {
    Logger.log('【ERROR】メイン処理中に重大なエラーが発生しました。');
    Logger.log(`エラーメッセージ: ${err.message}`);
    Logger.log(`スタックトレース: ${err.stack}`);
     try {
       SpreadsheetApp.getActiveSpreadsheet().toast(`エラーが発生しました: ${err.message}`, '処理失敗', 10);
     } catch (toastErr) {
       Logger.log(`【ERROR】エラー通知の表示中にさらにエラー: ${toastErr.message}`);
     }
  } finally {
    Logger.log('メイン処理_docomo端末価格取得 が終了しました。');
  }
}

/**
 * テスト関数
 */
function test_docomo_apiLogic() {
  Logger.log('test_docomo_apiLogic を開始します。(v3.8.0: 容量API対応)');

  try {
    const authResult = docomo_getAuthToken();
    const transactionInfo = docomo_getTransactionId(authResult); 
    const firstPageResponse = docomo_fetchApiPage('1', authResult);
    
    if (!firstPageResponse || !firstPageResponse.result || !Array.isArray(firstPageResponse.result.csOlsLmd04MobileList)) {
      throw new Error('APIからレスポンスが取得できないか、機種リストが空です。');
    }
    
    const parsedData = docomo_parseAndExtractData([firstPageResponse]);
    const outputData = docomo_aggregateData(parsedData, authResult, transactionInfo);

    Logger.log(`--- ヘッダー: ${JSON.stringify(DOCOMO_SETTINGS.HEADER_ROW[0])} ---`);
    
    if (outputData.length > 0) {
      const sampleData = outputData.slice(0, 5);
      Logger.log(`抽出データ (先頭 ${sampleData.length} 件):`);
      sampleData.forEach((row, index) => {
        Logger.log(`[${index + 1}] ${JSON.stringify(row)}`);
      });
    } else {
      Logger.log('抽出データなし');
    }
    Logger.log('-----------------------------------');

  } catch (err) {
    Logger.log('【ERROR】テスト処理中にエラーが発生しました。');
    Logger.log(`エラーメッセージ: ${err.message}`);
    Logger.log(`スタックトレース: ${err.stack}`);
  } finally {
    Logger.log('test_docomo_apiLogic が終了しました。');
  }
}