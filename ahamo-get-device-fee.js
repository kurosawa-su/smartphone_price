/**
 * @fileoverview ahamoデータ取得・整形スクリプト (中古ランク紐付け復旧版)
 * @version 5.4.0 (2025-11-14)
 * - 1. (v5.4.0) 【中古修正】IDで紐付かない中古データについて、機種名による辞書再検索(逆引き)ロジックを復旧し、ランク判定ができるように修正。
 * - 2. (v5.3.0) 新品価格ロジック修正。
 */

// スクリプト全体で使用する設定値（定数）
const AHAMO_SETTINGS = {
  API_URL_NEW: 'https://ahamo.com/price_info/mobilephone.json',
  API_URL_USED: 'https://ahamo.com/price_info/used_devices.json',
  API_URL_STOCK: 'https://ahamo.com/api/cil/tra/ptscf/v3.2/olstermlistget',
  
  SHEET_NAME_COMMON: 'ahamo端末一覧',
  HEADER_ROW_COMMON: [['機種名', '容量', '在庫', '端末価格', '割引後価格', '返却価格', '状態']],
  
  USER_AGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  CAPACITY_REGEX: /^(.*?) *(\d+ *(GB|TB)) *(.*)$/i,
};

/**
 * 容量表記を正規化するヘルパー関数
 */
function normalizeCapacity(cap) {
  if (!cap) return ''; 
  let s = String(cap).trim().toUpperCase().replace(/\s+/g, '');
  if (s === '1000' || s === '1000GB') return '1TB';
  if (s === '2000' || s === '2000GB') return '2TB';
  if (/^\d+$/.test(s)) return s + 'GB';
  return s;
}

/**
 * 機種名を整形するヘルパー関数
 */
function cleanProductNameString(name) {
  if (!name) return '不明な機種';
  let cleaned = name;
  const match = name.match(AHAMO_SETTINGS.CAPACITY_REGEX);
  if (match && match[1]) {
    cleaned = match[1].trim();
  }
  cleaned = cleaned.replace(/\(/g, '（').replace(/\)/g, '）');
  return cleaned;
}

/**
 * 文字列からランク情報を抽出するヘルパー関数
 */
function extractRankFromText(text) {
  if (!text) return null;
  const patterns = [
    /ランク\s*([SAB][+\uff0b]?)/i,      // ランクA
    /([SAB][+\uff0b]?)\s*ランク/i,      // Aランク
    /【([SAB][+\uff0b]?)】/             // 【A】
  ];
  
  for (const regex of patterns) {
    const match = text.match(regex);
    if (match) {
      let r = match[1].toUpperCase();
      if (r.includes('\uff0b')) r = r.replace('\uff0b', '+');
      return r;
    }
  }
  return null;
}

/**
 * HTTP GETリクエスト
 */
function ahamo_fetchApiData(url) {
  Logger.log(`[GET] Requesting: ${url}`);
  const params = { 'method': 'get', 'headers': { 'User-Agent': AHAMO_SETTINGS.USER_AGENT }, 'muteHttpExceptions': true };
  const response = UrlFetchApp.fetch(url, params);
  if (response.getResponseCode() !== 200) throw new Error(`API取得失敗: ${response.getResponseCode()}`);
  return JSON.parse(response.getContentText('UTF-8'));
}

/**
 * HTTP POSTリクエスト
 */
function ahamo_fetchApiDataPost(url, payload) {
  Logger.log(`[POST] Requesting: ${url}`);
  const params = { 'method': 'post', 'contentType': 'application/json', 'headers': { 'User-Agent': AHAMO_SETTINGS.USER_AGENT }, 'payload': JSON.stringify(payload), 'muteHttpExceptions': true };
  const response = UrlFetchApp.fetch(url, params);
  if (response.getResponseCode() !== 200) throw new Error(`API取得失敗: ${response.getResponseCode()}`);
  return JSON.parse(response.getContentText('UTF-8'));
}

/**
 * 在庫APIからmobileInfoリストを抽出
 */
function extractMobileInfoList(stockApiResponse) {
  if (!stockApiResponse || !stockApiResponse.terminalInfo) return [];
  const allMobileInfo = [];
  stockApiResponse.terminalInfo.forEach(term => {
    const terminalName = term.terminalNameExcludeModelNumber || term.terminalName || '';
    if (term.mobileInfo && Array.isArray(term.mobileInfo)) {
      term.mobileInfo.forEach(mobile => {
        mobile._derivedName = terminalName;
        allMobileInfo.push(mobile);
      });
    }
  });
  return allMobileInfo;
}

/**
 * 新品用マップ作成
 */
function createNewProductMap(apiResponse) {
  Logger.log('新品製品情報の詳細解析を開始します。');
  const productMap = new Map();
  const productList = Array.isArray(apiResponse.mobilephone) ? apiResponse.mobilephone : [];

  for (const model of productList) {
    if (!model || !model.id) continue;
    
    const productId = String(model.id).trim();
    const productName = cleanProductNameString(model.productName);

    if (Array.isArray(model.itemInfo) && model.itemInfo.length > 0) {
      for (const item of model.itemInfo) {
        let rawCapacity = item.name;
        if (!rawCapacity) {
            const match = (model.productName || '').match(AHAMO_SETTINGS.CAPACITY_REGEX);
            if (match && match[2]) rawCapacity = match[2];
        }
        const capacity = normalizeCapacity(rawCapacity); 
        
        const basePrice = (item.price && item.price.amount) ? item.price.amount : 0;
        let discountPrice = basePrice;
        let kaedokiPrice = 0;

        if (item.finalPrice && item.finalPrice.mnp) {
          const mnpInfo = item.finalPrice.mnp;
          if (mnpInfo.price && mnpInfo.price.amount) discountPrice = mnpInfo.price.amount;
          if (mnpInfo.priceKaedoki && mnpInfo.priceKaedoki.amount) kaedokiPrice = mnpInfo.priceKaedoki.amount;
        }
        
        if (kaedokiPrice === 0) {
            if (item.priceKaedoki && item.priceKaedoki.amount) kaedokiPrice = item.priceKaedoki.amount;
            else kaedokiPrice = basePrice;
        }
        
        if (!productMap.has(productId)) {
            productMap.set(productId, { 
                name: productName, 
                capacity: capacity,
                price: basePrice, 
                discountPrice: discountPrice, 
                kaedokiPrice: kaedokiPrice,
                rank: '新品'
            });
        }
      }
    } else {
      let rawCapacity = '';
      const match = (model.productName || '').match(AHAMO_SETTINGS.CAPACITY_REGEX);
      if (match && match[2]) rawCapacity = match[2];
      const capacity = normalizeCapacity(rawCapacity);

      const basePrice = (model.price && model.price.amount) ? model.price.amount : 0;
      let kaedokiPrice = 0;
      if (model.priceKaedoki && model.priceKaedoki.amount) kaedokiPrice = model.priceKaedoki.amount;
      if (kaedokiPrice === 0) kaedokiPrice = basePrice;

      productMap.set(productId, {
        name: productName,
        capacity: capacity,
        price: basePrice,
        discountPrice: basePrice,
        kaedokiPrice: kaedokiPrice,
        rank: '新品'
      });
    }
  }
  Logger.log(`新品辞書作成完了: ${productMap.size}件のIDを登録`);
  return productMap;
}

/**
 * 中古用マップ作成
 */
function createUsedProductMap(apiResponse) {
  Logger.log('中古製品情報の解析を開始します。');
  const productMap = new Map();
  let rankCount = { '中古A+': 0, '中古A': 0, '中古B': 0, '中古': 0 };
  
  function traverse(obj, parentId = null, inheritedRank = null, inheritedName = null) {
    if (!obj || typeof obj !== 'object') return;

    let currentId = obj.id || obj.modelCode || obj.productCode || obj.sku;
    
    if (!currentId && obj.image && typeof obj.image === 'string') {
       const match = obj.image.match(/\/([A-Za-z0-9]+)_[A-Z]\.jpg/);
       if (match) currentId = match[1];
    }
    
    const effectiveId = currentId ? String(currentId).trim() : parentId;

    let rawName = obj.productName || obj.modelName || inheritedName;
    let name = rawName ? cleanProductNameString(rawName) : null;

    const register = (key) => {
        if (!productMap.has(key)) {
             productMap.set(key, { 
                 name: name || '名称不明', 
                 rank: '中古', 
                 price: 0,
                 priceRankMap: {} 
             });
        } else if (name) {
             const entry = productMap.get(key);
             if (entry.name === '名称不明') entry.name = name;
        }
        return productMap.get(key);
    };

    let entryById = effectiveId ? register(effectiveId) : null;
    let entryByName = name ? register('NAME_' + name) : null;

    const currentRank = obj.rank || inheritedRank;

    let price = 0;
    if (obj.amount) price = obj.amount;
    else if (obj.price && typeof obj.price === 'object' && obj.price.amount) price = obj.price.amount;
    else if (typeof obj.price === 'number') price = obj.price;

    if (currentRank && price > 0) {
        if (entryById) entryById.priceRankMap[String(price)] = currentRank;
        if (entryByName) entryByName.priceRankMap[String(price)] = currentRank;
        
        if (rankCount[currentRank] !== undefined) rankCount[currentRank]++;
    }

    if (Array.isArray(obj)) {
      obj.forEach(item => traverse(item, effectiveId, currentRank, name));
    } else {
      Object.keys(obj).forEach(key => {
        let nextRank = currentRank;
        const keyLower = key.toLowerCase();
        
        if (/rank.?a.?plus/i.test(key)) nextRank = '中古A+';
        else if (/rank.?a(?!.?plus)/i.test(key)) nextRank = '中古A';
        else if (/rank.?b/i.test(key)) nextRank = '中古B';
        
        traverse(obj[key], effectiveId, nextRank, name);
      });
    }
  }

  traverse(apiResponse);
  Logger.log(`中古辞書作成完了: ${productMap.size}件のキーを登録`);
  Logger.log(`ランク別価格登録数: ${JSON.stringify(rankCount)}`);
  return productMap;
}

/**
 * 在庫リストを処理して出力データを作る関数
 */
function processStockList(stockList, nameMap, label) {
  const rows = [];
  
  for (const info of stockList) {
    const modelCode = String(info.modelCode).trim();
    
    // 辞書から製品情報を検索
    let productInfo = nameMap.get(modelCode);

    // 容量決定
    let rawCapacity = info.capacity;
    if (!rawCapacity && productInfo && productInfo.capacity) {
        rawCapacity = productInfo.capacity;
    }
    const capacity = normalizeCapacity(rawCapacity);

    // 在庫・価格
    let stockStatus = '在庫なし';
    if (info.saleStockFlag === '1' || info.saleStockFlag === '2') stockStatus = '在庫あり';

    let stockPrice = 0;
    if (info.priceInfo && info.priceInfo[0] && info.priceInfo[0].price) {
        stockPrice = info.priceInfo[0].price;
    }

    // 機種名
    let outName = '';
    if (info._derivedName) {
        outName = cleanProductNameString(info._derivedName); 
    } else if (productInfo && productInfo.name) {
        outName = productInfo.name;
    } else {
        outName = `(ID: ${modelCode})`;
    }
    
    // (v5.4.0) 中古でID紐付け失敗時に、名前で辞書を再検索（復旧）
    if (!productInfo && label === '中古' && outName) {
        productInfo = nameMap.get('NAME_' + outName);
    }
    
    // 状態 (ランク) 判定
    let outRank = label === '新品' ? '新品' : '中古';
    
    if (label === '中古') {
        const rankFromName = extractRankFromText(info._derivedName) || (productInfo ? extractRankFromText(productInfo.name) : null);
        if (rankFromName) {
            outRank = rankFromName;
        }
        else if (productInfo && productInfo.priceRankMap) {
            // 価格照合 (productInfoがあるから可能)
            const rankFromPrice = productInfo.priceRankMap[String(stockPrice)];
            if (rankFromPrice) {
                outRank = rankFromPrice;
            } else if (productInfo.rank) {
                outRank = productInfo.rank; 
            }
        }
        else if (productInfo && productInfo.rank) {
            outRank = productInfo.rank;
        }
    } else {
        if (productInfo && productInfo.rank) outRank = productInfo.rank;
    }
    
    if (label === '中古' && !outRank.startsWith('中古')) {
        outRank = '中古' + outRank;
    }

    // 価格情報の決定
    let outPrice = 0;
    let outDiscount = 0;
    let outKaedoki = '-';

    if (label === '新品') {
        if (productInfo) {
            // 新品: 辞書の定価を正とする
            outPrice = productInfo.price;
            outDiscount = productInfo.discountPrice;
            outKaedoki = productInfo.kaedokiPrice;
        } else {
            // 辞書がない場合のみ在庫API価格(MNP価格)を定価として代用
            outPrice = stockPrice;
            outDiscount = stockPrice;
        }
    } else {
        // 中古: 在庫API価格が正
        outPrice = stockPrice > 0 ? stockPrice : (productInfo ? productInfo.price : 0);
        outDiscount = outPrice;
    }

    // 価格逆転補正
    if (outDiscount > outPrice) {
        outDiscount = outPrice;
    }

    rows.push([
      outName,      // 機種名
      capacity,     // 容量
      stockStatus,  // 在庫
      outPrice,     // 端末価格
      outDiscount,  // 割引後価格
      outKaedoki,   // 返却価格
      outRank       // 状態
    ]);
  }
  
  Logger.log(`[${label}] ${rows.length}件の在庫データを処理しました。`);
  return rows;
}

// ----------------------------------------------------------------------
// メイン処理
// ----------------------------------------------------------------------

function メイン処理_ahamo端末価格取得() {
  Logger.log('=== メイン処理_ahamo端末価格取得 (v5.4.0) ===');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const newJson = ahamo_fetchApiData(AHAMO_SETTINGS.API_URL_NEW);
    const newNameMap = createNewProductMap(newJson);
    
    const usedJson = ahamo_fetchApiData(AHAMO_SETTINGS.API_URL_USED);
    const usedNameMap = createUsedProductMap(usedJson);
    
    const newStockJson = ahamo_fetchApiDataPost(AHAMO_SETTINGS.API_URL_STOCK, { "orderDiv": ["03"], "usedFlag": "0" });
    const newStockList = extractMobileInfoList(newStockJson);
    
    const usedStockJson = ahamo_fetchApiDataPost(AHAMO_SETTINGS.API_URL_STOCK, { "orderDiv": ["03"], "usedFlag": "1" });
    const usedStockList = extractMobileInfoList(usedStockJson);
    
    const newRows = processStockList(newStockList, newNameMap, '新品');
    const usedRows = processStockList(usedStockList, usedNameMap, '中古');
    
    const allRows = newRows.concat(usedRows);
    Logger.log(`=== 合計出力件数: ${allRows.length}件 ===`);
    
    if (allRows.length > 0) {
      let sheet = ss.getSheetByName(AHAMO_SETTINGS.SHEET_NAME_COMMON);
      if (!sheet) sheet = ss.insertSheet(AHAMO_SETTINGS.SHEET_NAME_COMMON);
      sheet.clearContents();
      
      const range = sheet.getRange(1, 1, 1, AHAMO_SETTINGS.HEADER_ROW_COMMON[0].length);
      range.setValues(AHAMO_SETTINGS.HEADER_ROW_COMMON);
      range.setBackground('#d9ead3').setFontWeight('bold');
      
      sheet.getRange(2, 1, allRows.length, AHAMO_SETTINGS.HEADER_ROW_COMMON[0].length).setValues(allRows);
      
      sheet.autoResizeColumn(1); 
      sheet.autoResizeColumn(2); 
      
      ss.toast(`完了: ${allRows.length}件のデータを取得しました`, '成功', 5);
    } else {
      ss.toast('データが1件も取得できませんでした', '警告', 10);
    }
    
  } catch (e) {
    Logger.log(`【エラー】: ${e.message}`);
    Logger.log(e.stack);
    ss.toast(`エラー発生: ${e.message}`, '失敗', 10);
  }
}