/**
 * @fileoverview 楽天モバイルの製品価格・在庫API(JSON)から情報を取得し、シートに書き込む
 * @version 1.2.6 (2025-11-18)
 * - (v1.2.6) すべての関数を含む完全版コード（省略なし）。関数未定義エラーを修正
 * - (v1.2.5) メイン関数名を `メイン処理_Rakuten端末価格取得` に変更
 * - (v1.2.4) HTMLテキスト解析機能を「取得済みHTMLの再利用」として復活
 * - (v1.2.1) 認定中古の状態表記を「中古」に変更。割引(クーポン)価格の取得ロジックを強化
 * - (v1.2.0) 「Rakuten 認定中古」ページをスクレイピングする機能を実装
 */

// スクリプト全体で使用する設定値（定数）
const RAKUTEN_SETTINGS = {
  // (API 1: 全製品グループリスト取得API)
  EQUIPMENTS_URL: 'https://onboarding.mobile.rakuten.co.jp/api/equipment/equipments',
  // (API 2: SKUリスト取得API)
  GROUP_URL: 'https://onboarding.mobile.rakuten.co.jp/api/equipment/getEquipmentGroup',
  // (API 3: SKU詳細・価格取得API)
  DETAILS_URL: 'https://onboarding.mobile.rakuten.co.jp/api/equipment/getEquipmentDetails',
  
  // (v1.2.0) 認定中古ページURL
  CERTIFIED_URL: 'https://www.rakuten.ne.jp/gold/rakutenmobile-store/product/rakuten-certified/',

  // (API 1) 取得対象のカテゴリコード
  CATEGORY_CODES: ['Smartphones', 'Apple Smartphones'],
  
  // (API 1) 1度のAPI呼び出しで取得する件数
  PAGE_LIMIT: 50,

  // API負荷軽減のための待機時間 (ミリ秒)
  API_WAIT_MS: 500,

  // JSファイル取得を行う最大件数 (タイムアウト防止用)
  MAX_JS_FETCH_COUNT: 20,

  // 書き込み対象のシート名
  SHEET_NAME: 'rakuten端末一覧',

  // ヘッダー行
  HEADER_ROW: [['機種名', '容量', '在庫', '端末価格', '割引後価格', '返却価格', '状態']],

  // データ書き込み開始行
  START_ROW: 2,
  // データ書き込み開始列
  START_COL: 1,

  // User-Agent
  USER_AGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  ORIGIN_URL: 'https://onboarding.mobile.rakuten.co.jp'
};

// --- グローバル変数 ---
let globalJsFetchCounter = 0;

/**
 * 共通のAPI取得処理 (POST)
 */
function rakuten_fetchApi(url, payload) {
  Logger.log(`API呼び出し (POST): ${url}`);
  const params = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'User-Agent': RAKUTEN_SETTINGS.USER_AGENT,
      'Origin': RAKUTEN_SETTINGS.ORIGIN_URL,
      'Referer': RAKUTEN_SETTINGS.ORIGIN_URL + '/'
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, params);
    if (response.getResponseCode() !== 200) {
      Logger.log(`【ERROR】ステータスコード: ${response.getResponseCode()}`);
      return null;
    }
    return JSON.parse(response.getContentText('UTF-8'));
  } catch (e) {
    Logger.log(`【ERROR】API呼び出し失敗: ${e.message}`);
    return null;
  }
}

/**
 * (v1.2.4) 製品ページ(HTML/JS)から価格情報を強力に抽出する関数
 * 改良点: HTMLテキスト解析を復活させ、JS変数が見つからない場合の保険とする
 */
function rakuten_fetchJsData(productPageUrl, capacity) {
  if (!productPageUrl || globalJsFetchCounter >= RAKUTEN_SETTINGS.MAX_JS_FETCH_COUNT) {
    return { discountAmount: 0, discountedPrice: '', returnPrice: '' };
  }

  Logger.log(`(Fetch) データ取得開始 (${globalJsFetchCounter + 1}/${RAKUTEN_SETTINGS.MAX_JS_FETCH_COUNT}): ${productPageUrl} (Cap: ${capacity})`);
  globalJsFetchCounter++;

  try {
    // 1. 製品ページ(HTML)を取得
    const pageResponse = UrlFetchApp.fetch(productPageUrl, {
      'muteHttpExceptions': true,
      'headers': { 'User-Agent': RAKUTEN_SETTINGS.USER_AGENT }
    });
    const htmlText = pageResponse.getContentText('UTF-8');

    let priceInJs = 0;
    let discountInJs = 0;
    let division48InJs = 0;
    let isReturnProgramTarget = false;
    
    // HTMLから取得した価格（保険用）
    let htmlDiscountedPrice = '';
    let htmlReturnPrice = '';

    // ---------------------------------------------------------
    // (A) HTMLテキスト解析 (v1.2.4 復活)
    // ---------------------------------------------------------
    
    // 1. 割引後価格 (「値引き後価格」または「支払い総額」)
    const discountLabelMatch = htmlText.match(/値引き後価格[\s\S]{0,300}?([0-9]{1,3}(?:,[0-9]{3})*)\s*(?:<\/span>)?\s*(?:<span>)?円/);
    if (discountLabelMatch && discountLabelMatch[1]) {
      const val = parseFloat(discountLabelMatch[1].replace(/,/g, ''));
      if (!isNaN(val)) htmlDiscountedPrice = val;
    }
    if (htmlDiscountedPrice === '') {
       const totalPaymentMatch = htmlText.match(/支払い総額[\s\S]{0,300}?([0-9]{1,3}(?:,[0-9]{3})*)\s*(?:<\/span>)?\s*(?:<span>)?円/);
       if (totalPaymentMatch && totalPaymentMatch[1]) {
          const val = parseFloat(totalPaymentMatch[1].replace(/,/g, ''));
          if (!isNaN(val)) htmlDiscountedPrice = val;
       }
    }

    // 2. 返却価格 (「実質負担額」または「48回払い」)
    const returnLabelMatch = htmlText.match(/実質負担額[\s\S]{0,300}?([0-9]{1,3}(?:,[0-9]{3})*)\s*(?:<\/span>)?\s*(?:<span>)?円/);
    if (returnLabelMatch && returnLabelMatch[1]) {
       const val = parseFloat(returnLabelMatch[1].replace(/,/g, ''));
       if (!isNaN(val)) htmlReturnPrice = val;
    }
    if (htmlReturnPrice === '') {
       const monthly48Match = htmlText.match(/48回払い[\s\S]{0,300}?([0-9]{1,3}(?:,[0-9]{3})*)\s*(?:<\/span>)?\s*(?:<span>)?円/);
       if (monthly48Match && monthly48Match[1]) {
         const val = parseFloat(monthly48Match[1].replace(/,/g, ''));
         if (!isNaN(val)) {
            htmlReturnPrice = val * 24;
            isReturnProgramTarget = true;
         }
       }
    }


    // ---------------------------------------------------------
    // (B) JSファイル解析 (Android詳細データ用)
    // ---------------------------------------------------------
    
    // JSファイルパス検索
    const jsPathMatch = htmlText.match(/src=["'](\/_next\/static\/chunks\/pages\/product\/[^"']+\.js)["']/);

    if (jsPathMatch && jsPathMatch[1]) {
        const jsUrl = 'https://network.mobile.rakuten.co.jp' + jsPathMatch[1];
        Logger.log(`  -> JSファイル特定: ${jsUrl}`);
        
        try {
            const jsResponse = UrlFetchApp.fetch(jsUrl, { 'muteHttpExceptions': true, 'headers': { 'User-Agent': RAKUTEN_SETTINGS.USER_AGENT } });
            const jsText = jsResponse.getContentText('UTF-8');

            const discountMatch = jsText.match(/\bfirstTimeApplyPoint\s*:\s*([0-9eE\.]+)/);
            if (discountMatch) discountInJs = parseFloat(discountMatch[1]);

            if (!isReturnProgramTarget) {
                if (jsText.indexOf('買い替え超トクプログラム') !== -1 || jsText.indexOf('replacement-program') !== -1) {
                    isReturnProgramTarget = true;
                }
            }

            const capPattern = capacity && capacity !== '' ? capacity.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') : null;
            let foundCapacityData = false;

            // 容量別データの検索
            if (capPattern) {
                 const pricesRegex = new RegExp(`["']?${capPattern}["']?\\s*:\\s*\\{([^}]+)\\}`, 'i');
                 const blockMatch = jsText.match(pricesRegex);
                 if (blockMatch) {
                    const lumpSumMatch = blockMatch[1].match(/\blumpSum\s*:\s*([0-9eE\.]+)/);
                    if (lumpSumMatch) { priceInJs = parseFloat(lumpSumMatch[1]); foundCapacityData = true; }
                    const divMatch = blockMatch[1].match(/\bdivision48\s*:\s*([0-9eE\.]+)/);
                    if (divMatch) { division48InJs = parseFloat(divMatch[1]); foundCapacityData = true; }
                 } else {
                     const storageRegex = new RegExp(`\\{[^}]*?${capPattern}[^}]*?\\}`, 'i');
                     const storageMatch = jsText.match(storageRegex);
                     if (storageMatch) {
                        const priceMatch = storageMatch[0].match(/\bprice\s*:\s*([0-9eE\.]+)/);
                        if (priceMatch) { priceInJs = parseFloat(priceMatch[1]); foundCapacityData = true; }
                        const divMatch = storageMatch[0].match(/\bdivision48\s*:\s*([0-9eE\.]+)/);
                        if (divMatch) { division48InJs = parseFloat(divMatch[1]); foundCapacityData = true; }
                     }
                 }
                 
                 const capLower = capPattern.toLowerCase();
                 const beforeRegex = new RegExp(`division48Before${capLower}\\s*:\\s*([0-9eE\\.]+)`, 'i');
                 const beforeMatch = jsText.match(beforeRegex);
                 if (beforeMatch) {
                    division48InJs = parseFloat(beforeMatch[1]);
                    foundCapacityData = true;
                    isReturnProgramTarget = true;
                    const priceRegex = new RegExp(`priceOf${capLower}\\s*:\\s*([0-9eE\\.]+)`, 'i');
                    const priceMatch = jsText.match(priceRegex);
                    if (priceMatch) priceInJs = parseFloat(priceMatch[1]);
                 }
            }

            // 容量データがなければ単純構造から
            if (!foundCapacityData) {
                if (priceInJs === 0) {
                    const simplePriceMatch = jsText.match(/(?:[{,]\s*)\bprice\s*:\s*([0-9eE\.]+)/);
                    if (simplePriceMatch) priceInJs = parseFloat(simplePriceMatch[1]);
                }
                if (division48InJs === 0) {
                    const simpleDivMatch = jsText.match(/\bdivision48\s*:\s*([0-9eE\.]+)/);
                    if (simpleDivMatch) division48InJs = parseFloat(simpleDivMatch[1]);
                }
            }
            
            // iPhone用 JSXハードコード月額 (変数がない場合)
            if (isReturnProgramTarget && division48InJs === 0) {
               const jsxMonthlyMatch = jsText.match(/children\s*:\s*["']\s*([0-9,]+)\s*["']\s*\}\s*\)\s*,\s*["']円\/月["']/);
               if (jsxMonthlyMatch && jsxMonthlyMatch[1]) {
                  const extractedPrice = parseFloat(jsxMonthlyMatch[1].replace(/[, \s]/g, ''));
                  if (!isNaN(extractedPrice)) division48InJs = extractedPrice;
               }
            }

            // JSで容量固有データが取れた場合、HTMLの値を無効化してJSを優先
            if (foundCapacityData) {
                 htmlReturnPrice = '';
                 htmlDiscountedPrice = '';
            }

        } catch (jsErr) {
            Logger.log(`  -> JS取得エラー: ${jsErr.message}`);
        }
    } else {
        Logger.log('  -> JSファイルパスが見つかりませんでした。HTML解析値を採用します。');
    }

    // ---------------------------------------------------------
    // 最終的な値の決定
    // ---------------------------------------------------------
    
    let finalDiscountedPrice = '';
    let finalReturnPrice = '';

    // 1. 割引後価格
    if (priceInJs > 0 && discountInJs > 0) {
        const calcPrice = priceInJs - discountInJs;
        if (calcPrice >= 0) finalDiscountedPrice = calcPrice;
    } 
    // JSで決まらなかった場合のみHTMLを採用
    if (finalDiscountedPrice === '' && htmlDiscountedPrice !== '') {
        finalDiscountedPrice = htmlDiscountedPrice;
    }

    // 2. 返却価格
    if (isReturnProgramTarget && division48InJs > 0) {
        finalReturnPrice = division48InJs * 24;
    }
    // JSで決まらなかった場合のみHTMLを採用
    if (finalReturnPrice === '' && htmlReturnPrice !== '') {
        finalReturnPrice = htmlReturnPrice;
    }

    Utilities.sleep(RAKUTEN_SETTINGS.API_WAIT_MS);

    return { 
      discountAmount: discountInJs,
      discountedPrice: finalDiscountedPrice,
      returnPrice: finalReturnPrice
    };

  } catch (e) {
    Logger.log(`【WARN】データ抽出失敗: ${e.message}`);
    return { discountAmount: 0, discountedPrice: '', returnPrice: '' };
  }
}

/**
 * (v1.2.1) Rakuten 認定中古情報をスクレイピングする関数
 */
function rakuten_fetchCertifiedUsed() {
  Logger.log(`(Certified) 認定中古ページの取得を開始します: ${RAKUTEN_SETTINGS.CERTIFIED_URL}`);
  const outputData = [];

  try {
    const response = UrlFetchApp.fetch(RAKUTEN_SETTINGS.CERTIFIED_URL, {
      'muteHttpExceptions': true,
      'headers': { 'User-Agent': RAKUTEN_SETTINGS.USER_AGENT }
    });
    const htmlText = response.getContentText('UTF-8');
    const productCardRegex = /<div class="i-Product-card" id="([^"]+)">([\s\S]*?)<\/div>\s*<\/div>\s*<\/div>/g;
    let match;

    while ((match = productCardRegex.exec(htmlText)) !== null) {
      const cardContent = match[2];
      const titleMatch = cardContent.match(/<h4[^>]*>([\s\S]*?)<\/h4>/);
      let title = titleMatch ? titleMatch[1].trim().replace(/\s+/g, ' ') : "不明な端末";
      title = title.replace(/【.*?】/g, '').trim();

      const normalPriceBlockRegex = /<dl class="i-Product-card_Price-normal[^"]*">([\s\S]*?)<\/dl>/g;
      let normalPriceBlock;
      let hasPriceBlock = false; // dlブロックがあったか

      while ((normalPriceBlock = normalPriceBlockRegex.exec(cardContent)) !== null) {
        hasPriceBlock = true;
        const blockInner = normalPriceBlock[1];
        const capacityBlockRegex = /<dt>([^<]+)<\/dt>\s*<dd>[\s\S]*?<span>([^<]+)<\/span>円/g;
        let capMatch;

        while ((capMatch = capacityBlockRegex.exec(blockInner)) !== null) {
          const capacity = capMatch[1].trim();
          let priceStr = capMatch[2].replace(/,/g, '').trim();
          
          const ddContentRegex = new RegExp(`<dt>${capacity}<\/dt>\\s*<dd>([\\s\\S]*?)<\/dd>`);
          const ddMatch = blockInner.match(ddContentRegex);
          
          if (ddMatch) {
             const ddInner = ddMatch[1];
             const arrowMatch = ddInner.match(/<span>→\s*([0-9,]+)<\/span>/);
             if (arrowMatch) {
               priceStr = arrowMatch[1].replace(/,/g, '');
             }
          }
          const price = parseInt(priceStr, 10);
          let couponPrice = '';
          const couponRegex = /class="[^"]*Price-coupon[^"]*"[\s\S]*?<span>([0-9,]+)<\/span>/;
          const couponMatch = cardContent.match(couponRegex);
          if (couponMatch) {
             couponPrice = parseInt(couponMatch[1].replace(/,/g, ''), 10);
          }

          outputData.push([
            title, capacity, '在庫あり', price, couponPrice, '', '中古'
          ]);
        }
      }
    }
    Logger.log(`(Certified) ${outputData.length} 件の中古端末を取得しました。`);
    return outputData;
  } catch (e) {
    Logger.log(`【ERROR】認定中古取得失敗: ${e.message}`);
    return [];
  }
}

/**
 * (API 1) カテゴリ取得
 */
function rakuten_fetchCategoryEquipments(categoryCode) {
  Logger.log(`(API 1) カテゴリ [${categoryCode}] の製品グループ取得を開始します...`);
  const allEquipments = [];
  let offset = 1;
  let total = 0;
  let hasMore = true;
  try {
    do {
      const payload = { 'categoryCodes': categoryCode, 'sort': 'groupOrderNumber,orderNumber', 'topOfferId': '', 'offset': offset, 'limit': RAKUTEN_SETTINGS.PAGE_LIMIT, 'filters': {}, 'searchText': '', 'offeringGroupingEnabled': true };
      const response = rakuten_fetchApi(RAKUTEN_SETTINGS.EQUIPMENTS_URL, payload);
      if (response && Array.isArray(response.equipments)) {
        allEquipments.push(...response.equipments);
        if (total === 0 && response.total) total = parseInt(response.total, 10) || 0;
        offset += RAKUTEN_SETTINGS.PAGE_LIMIT;
        hasMore = (offset <= total && response.equipments.length > 0);
      } else { hasMore = false; }
      if (hasMore) Utilities.sleep(RAKUTEN_SETTINGS.API_WAIT_MS);
    } while (hasMore);
  } catch (e) { Logger.log(`【ERROR】rakuten_fetchCategoryEquipments 処理中にエラー: ${e.message}`); }
  return allEquipments;
}

/**
 * (API 2) SKU一覧取得
 */
function rakuten_fetchEquipmentGroup(groupId) {
  try {
    const payload = { 'value': groupId };
    const response = rakuten_fetchApi(RAKUTEN_SETTINGS.GROUP_URL, payload);
    return (response && Array.isArray(response.equipmentBases)) ? response.equipmentBases : [];
  } catch (e) { return []; }
}

/**
 * (API 3) 詳細情報取得
 */
function rakuten_fetchEquipmentDetails(skuId) {
  try {
    const payload = { 'value': skuId };
    return rakuten_fetchApi(RAKUTEN_SETTINGS.DETAILS_URL, payload);
  } catch (e) { return null; }
}

function rakuten_mapStockStatus(isInStock) { return isInStock ? '在庫あり' : '在庫なし'; }
function rakuten_escapeRegExp(str) { if (typeof str !== 'string' || str.length === 0) return ''; return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }


/**
 * APIレスポンスを解析し、シート書き込み用の2次元配列に変換します。
 */
function rakuten_parseAndExtractData(allSkusMap, allDetailsMap, skuUrlMap) {
  Logger.log('全データの解析とシート配列への変換を開始します...');
  const aggregatedMap = new Map();

  for (const [skuId, sku] of allSkusMap.entries()) {
    try {
      const originalSkuName = sku.name || '';
      const capacity = sku.memorySize || '';
      const currentSkuIsInStock = (sku.isAvailableInStock === true);

      let cleanedSkuName = originalSkuName;
      try {
        const capacityStr = rakuten_escapeRegExp(capacity);
        const color = (sku.color && sku.color.name) ? sku.color.name : '';
        const colorStr = rakuten_escapeRegExp(color);
        if (capacityStr !== '' && capacityStr.length > 0) cleanedSkuName = cleanedSkuName.replace(new RegExp(`\\s*${capacityStr}\\s*`, 'g'), ' ').trim();
        if (colorStr !== '' && colorStr.length > 0) cleanedSkuName = cleanedSkuName.replace(new RegExp(`\\s*${colorStr}\\s*`, 'g'), ' ').trim();
        cleanedSkuName = cleanedSkuName.replace(/\s+/g, ' ');
      } catch (e_clean) { cleanedSkuName = originalSkuName; }

      const details = allDetailsMap.get(skuId);
      let terminalPrice = '';
      let returnPrice = '';
      let discountedPrice = ''; 
      const condition = '新品';

      if (details) {
        if (details.smartphoneFreeDetails?.amountWithTax) {
          terminalPrice = Math.floor(parseFloat(String(details.smartphoneFreeDetails.amountWithTax).replace(/,/g, '')));
        }
        if (details.smfInstallments?.smfFirstPartInstallment) {
           const firstInstallment = parseFloat(String(details.smfInstallments.smfFirstPartInstallment).replace(/,/g, ''));
           if (!isNaN(firstInstallment)) returnPrice = firstInstallment * 24;
        }
      }

      const aggregationKey = `${cleanedSkuName}|${capacity}`;

      if (!aggregatedMap.has(aggregationKey)) {
        aggregatedMap.set(aggregationKey, {
          name: cleanedSkuName,
          capacity: capacity,
          isInStock: currentSkuIsInStock,
          terminalPrice: terminalPrice,
          discountedPrice: discountedPrice, 
          returnPrice: returnPrice,
          condition: condition,
          url: skuUrlMap.get(skuId)
        });
      } else {
        const entry = aggregatedMap.get(aggregationKey);
        if (currentSkuIsInStock) entry.isInStock = true;
        if (!entry.url && skuUrlMap.get(skuId)) entry.url = skuUrlMap.get(skuId);
      }
    } catch (e) { Logger.log(`解析エラー: ${e.message}`); }
  }
  
  const outputData = [];
  for (const entry of aggregatedMap.values()) {
    if (entry.discountedPrice === '' && entry.url && globalJsFetchCounter < RAKUTEN_SETTINGS.MAX_JS_FETCH_COUNT) {
      const jsData = rakuten_fetchJsData(entry.url, entry.capacity);
      if (jsData.discountedPrice !== '') {
         entry.discountedPrice = jsData.discountedPrice;
      } else if (entry.terminalPrice !== '' && jsData.discountAmount > 0) {
         const calcPrice = entry.terminalPrice - jsData.discountAmount;
         if (calcPrice >= 0) entry.discountedPrice = calcPrice;
      }
      // iPhoneの場合、APIの返却価格を優先
      const isIphone = entry.name.indexOf('iPhone') !== -1;
      if (isIphone && entry.returnPrice !== '') {
          // 維持
      } else if (jsData.returnPrice !== '') {
          entry.returnPrice = jsData.returnPrice;
      }
    }

    outputData.push([
      entry.name,
      entry.capacity,
      rakuten_mapStockStatus(entry.isInStock),
      entry.terminalPrice,
      entry.discountedPrice,
      entry.returnPrice,
      entry.condition
    ]);
  }
  return outputData;
}

/**
 * データをシートに書き込み
 */
function rakuten_writeToSheet(outputData) {
  Logger.log(`シート (${RAKUTEN_SETTINGS.SHEET_NAME}) への書き込みを開始します。`);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(RAKUTEN_SETTINGS.SHEET_NAME);
    if (!sheet) { sheet = ss.insertSheet(RAKUTEN_SETTINGS.SHEET_NAME); }
    sheet.clearContents();
    const headerCols = RAKUTEN_SETTINGS.HEADER_ROW[0].length;
    sheet.getRange(1, 1, RAKUTEN_SETTINGS.HEADER_ROW.length, headerCols).setValues(RAKUTEN_SETTINGS.HEADER_ROW).setBackground('#f3f3f3').setFontWeight('bold');

    if (outputData.length > 0) {
      const dataCols = outputData[0].length; 
      sheet.getRange(RAKUTEN_SETTINGS.START_ROW, RAKUTEN_SETTINGS.START_COL, outputData.length, dataCols).setValues(outputData);
      Logger.log(`${outputData.length} 件のデータを書き込みました。`);
      sheet.autoResizeColumns(RAKUTEN_SETTINGS.START_COL, dataCols);
    } else { Logger.log('書き込むデータがありませんでした。'); }
    ss.toast('完了しました。', '処理完了', 5);
  } catch (e) { Logger.log(`【ERROR】シート書き込みエラー: ${e.message}`); }
}

/**
 * メイン処理
 */
function rakuten_fetchAllApiData() {
  const allSkusMap = new Map();
  const allDetailsMap = new Map();
  const skuUrlMap = new Map();
  globalJsFetchCounter = 0;

  // 1. 新品APIデータの取得
  for (const categoryCode of RAKUTEN_SETTINGS.CATEGORY_CODES) {
    const groups = rakuten_fetchCategoryEquipments(categoryCode);
    for (const group of groups) {
      if (!group.id) continue;
      const groupUrl = group.detailsLink || null;
      const skus = rakuten_fetchEquipmentGroup(group.id);
      for (const sku of skus) {
        if (sku.id && !allSkusMap.has(sku.id)) {
          allSkusMap.set(sku.id, sku);
          if (groupUrl) skuUrlMap.set(sku.id, groupUrl);
        }
      }
      Utilities.sleep(RAKUTEN_SETTINGS.API_WAIT_MS);
    }
  }

  // 詳細取得
  let count = 0;
  for (const skuId of allSkusMap.keys()) {
    const details = rakuten_fetchEquipmentDetails(skuId);
    if (details) allDetailsMap.set(skuId, details);
    count++;
    if (count % 20 === 0) Utilities.sleep(RAKUTEN_SETTINGS.API_WAIT_MS);
  }
  
  // 新品データの解析
  const newProductData = rakuten_parseAndExtractData(allSkusMap, allDetailsMap, skuUrlMap);
  
  // 2. 認定中古データの取得 (v1.2.1)
  const usedProductData = rakuten_fetchCertifiedUsed();
  
  // データの結合
  const finalData = [...newProductData, ...usedProductData];
  
  return finalData;
}

/**
 * メイン処理_Rakuten端末価格取得
 */
function メイン処理_Rakuten端末価格取得() {
  Logger.log('メイン処理_Rakuten端末価格取得 を開始します。');
  try {
    const outputData = rakuten_fetchAllApiData();
    rakuten_writeToSheet(outputData);
  } catch (err) { Logger.log(`【ERROR】${err.message}`); } 
  finally { Logger.log('処理終了'); }
}

function test_rakuten_apiLogic() {
    Logger.log('test function skipped for v1.2.6');
}