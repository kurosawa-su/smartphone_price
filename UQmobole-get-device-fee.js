/**
 * @fileoverview UQ WiMAXの製品一覧から価格・在庫情報を取得する解析スクリプト
 * @version 17.1.0 (2025-12-19)
 * - 【修正】製品リストとの紐付けに失敗した機種（iPhone 15等）についても、機種名からURLを推測して在庫取得を試みるよう変更
 * - 【解決】名前の不一致により在庫情報が取れない（紐付け失敗となる）主要機種のデータ取得率を向上
 */

// スクリプト全体で使用する設定値（定数）
const UQ_SETTINGS = {
  // データ埋め込み元の製品一覧ページURL（メイン）
  TARGET_PAGE_URL: 'https://www.uqwimax.jp/mobile/products/',
  
  // 認定中古品のリンク探索用ページURL
  AUCERTIFIED_PAGE_URL: 'https://shop.uqmobile.jp/shop/aucertified/',

  // ドメイン
  BASE_DOMAIN: 'https://www.uqwimax.jp',
  SHOP_DOMAIN: 'https://shop.uqmobile.jp',

  // 1. HTML内の変数 _productsTxt (製品リスト) の抽出マーカー
  DATA_START_MARKER: 'let _productsTxt = `',
  DATA_END_MARKER: '`;',

  // 2. 価格計算ロジックJSのパス抽出パターン
  UTIL_JS_PATH_PATTERN: /src="([^"]*?productprice2024util\.js[^"]*)"/,

  // 3. 価格JSONのパス抽出パターン
  PRICE_JSON_PATH_PATTERN: /['"]([^'"]*?product_prices[^'"]*?\.json)['"]/,
  
  // 4. 在庫情報JSのパス抽出パターン
  INVENTORY_JS_PATH_PATTERN: /src=["']([^"']*?\/data\/[^"']*?\.js[^"']*?)["']/,

  // 5. 製品紹介ページ内の購入ボタン(ショップリンク)抽出パターン
  SHOP_LINK_PATTERN: /href=["']((?:https:\/\/shop\.uqmobile\.jp)?\/detail\/[^"']+)["']/,

  // フィルタリング条件（ユーザー指定）
  FILTER: {
    ENABLED: true,       
    KEY_ORDER: 'オーダー',
    VAL_ORDER: 'MNP',
    KEY_OPTION: '増量オプション',
    VAL_OPTION: 'あり',
    KEY_POWER: 'Power',
    VAL_POWERS: ['トクトク2', 'コミプラ', 'コミコミ'] 
  },

  // 並列リクエストのバッチサイズ
  FETCH_BATCH_SIZE: 10,
  
  // 書き込み設定
  SHEET_NAME: 'UQ端末価格一覧',
  START_ROW: 2,
  START_COL: 1,
  // 固定ヘッダー定義
  HEADERS: ['機種名', '容量', '在庫', '端末価格', '割引後価格', '返却価格', '状態'],

  USER_AGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
};

/**
 * 文字列の全角英数字を半角に変換する
 */
function uq_toHalfWidth(str) {
  if (!str) return '';
  return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
}

/**
 * 製品名を正規化するヘルパー関数（名寄せ用）
 */
function uq_normalizeName(name) {
  if (!name) return '';
  
  let s = name.toLowerCase();
  s = uq_toHalfWidth(s);

  // 特定キーワードの処理
  s = s.replace(/【uq】/g, '')
       .replace(/au\s*certified/g, 'certified') 
       .replace(/認定中古品/g, 'certified')      
       .replace(/\(x\)/g, '').replace(/（x）/g, '') 
       .replace(/\(v\)/g, '').replace(/（v）/g, '') 
       .replace(/（第[0-9]+世代）/g, '') 
       .replace(/\(.*?generation\)/g, '') 
       .replace(/[0-9]+gb/g, '') // 容量削除
       .replace(/5g/g, '') 
       .replace(/[（(].*?[)）]/g, ''); 

  // 型番削除
  if (!s.includes('iphone') && !s.includes('pixel')) {
      s = s.replace(/\s+[a-z]{3,}[0-9]{2,}$/, '');
  }

  // スペース・記号削除
  s = s.replace(/[\s\u3000]/g, '')
       .replace(/[!-/:-@[-`{-~]/g, '');

  return s;
}

/**
 * URLを正規化するヘルパー関数（紐付け用）
 */
function uq_normalizeUrl(url) {
  if (!url) return '';
  let path = url.replace(/^https?:\/\/[^\/]+/, '');
  path = path.split('?')[0];
  path = path.replace(/\/index\.html$/, '/');
  if (path.length > 1 && path.endsWith('/')) {
    path = path.slice(0, -1);
  }
  return path;
}

/**
 * 【変更】製品名からショップURLを推測生成する（Certified以外も対応）
 * 例: "iPhone 14" -> "https://shop.uqmobile.jp/detail/iphone14/"
 * 例: "au Certified iPhone 14 Pro" -> "https://shop.uqmobile.jp/detail/aucertified_iphone14_pro/"
 */
function uq_guessShopUrl(productName) {
  let name = productName.toLowerCase();
  name = uq_toHalfWidth(name);
  
  // 不要な装飾を削除
  name = name.replace(/【uq】/g, '')
             .replace(/\(x\)/g, '')
             .replace(/（x）/g, '')
             .replace(/\(v\)/g, '')
             .replace(/（v）/g, '')
             .replace(/（第([0-9]+)世代）/g, 'se$1') 
             .replace(/[^a-z0-9\s]/g, '') 
             .replace(/[0-9]+gb/g, ''); 

  // "au Certified" -> "aucertified" (スペース詰め)
  name = name.replace(/au\s*certified/g, 'aucertified');
  name = name.replace(/認定中古品/g, 'aucertified');
  
  // iPhone/Pixelと数字の間のスペースを削除 (iphone 14 -> iphone14)
  name = name.replace(/(iphone|pixel)\s+([0-9])/g, '$1$2');

  // 残りのスペースをアンダースコアに置換 (pro max -> pro_max)
  name = name.trim().replace(/\s+/g, '_');
  
  // 重複アンダースコア削除
  name = name.replace(/_+/g, '_');

  return UQ_SETTINGS.SHOP_DOMAIN + '/detail/' + name + '/';
}

/**
 * ページHTMLを取得する共通関数
 */
function uq_fetchPageHtml(url) {
  try {
    const response = UrlFetchApp.fetch(url, {
      'method': 'get',
      'headers': { 'User-Agent': UQ_SETTINGS.USER_AGENT },
      'muteHttpExceptions': true
    });
    if (response.getResponseCode() !== 200) {
      Logger.log(`アクセス失敗: ${response.getResponseCode()} URL: ${url}`);
      return null;
    }
    return response.getContentText();
  } catch (e) {
    Logger.log(`URL取得エラー: ${e.message} URL: ${url}`);
    return null;
  }
}

/**
 * 複数のURLを並列で取得する関数
 */
function uq_fetchPagesParallel(urls) {
  if (!urls || urls.length === 0) return [];
  
  const results = [];
  for (let i = 0; i < urls.length; i += UQ_SETTINGS.FETCH_BATCH_SIZE) {
    const batchUrls = urls.slice(i, i + UQ_SETTINGS.FETCH_BATCH_SIZE);
    const requests = batchUrls.map(url => ({
      url: url,
      method: 'get',
      headers: { 'User-Agent': UQ_SETTINGS.USER_AGENT },
      muteHttpExceptions: true
    }));

    try {
      const responses = UrlFetchApp.fetchAll(requests);
      responses.forEach(res => {
        if (res.getResponseCode() === 200) {
          results.push(res.getContentText());
        } else {
          results.push(null);
        }
      });
      Utilities.sleep(500); 
    } catch (e) {
      Logger.log(`バッチリクエストエラー: ${e.message}`);
      batchUrls.forEach(() => results.push(null));
    }
  }
  return results;
}

/**
 * STEP 1: HTMLから製品リスト(_products)を抽出
 */
function uq_extractProductsFromHtml(html) {
  Logger.log(`STEP 1: 製品リスト抽出中...`);
  const startIdx = html.indexOf(UQ_SETTINGS.DATA_START_MARKER);
  if (startIdx === -1) throw new Error('製品リスト開始位置が見つかりません');
  
  const dataStart = startIdx + UQ_SETTINGS.DATA_START_MARKER.length;
  const endIdx = html.indexOf(UQ_SETTINGS.DATA_END_MARKER, dataStart);
  if (endIdx === -1) throw new Error('製品リスト終了位置が見つかりません');

  let jsonString = html.substring(dataStart, endIdx);
  jsonString = jsonString.replace(/\[an error occurred while processing this directive\](,)?/g, "");
  jsonString = jsonString.replace(/([^:]|^)\/\/.*/g, '$1');
  jsonString = jsonString.replace(/[\r\n\t]+/g, ' ');

  try {
    const data = new Function("return (" + jsonString + ");")();
    if (!Array.isArray(data)) throw new Error('配列ではありません');
    
    Logger.log(`  -> 抽出成功: ${data.length}件`);
    return data;
  } catch (e) {
    throw new Error(`製品リスト解析失敗: ${e.message}`);
  }
}

/**
 * 認定中古品ページからリンクを取得し、製品リストを補完する
 */
function uq_complementCertifiedLinks(products) {
  Logger.log('STEP 1.5: 認定中古品ページからリンク情報を補完します。');
  const html = uq_fetchPageHtml(UQ_SETTINGS.AUCERTIFIED_PAGE_URL);
  if (!html) return;

  const linkRegex = /href=["']([^"']+)["']/g;
  const links = new Set();
  let match;
  while ((match = linkRegex.exec(html)) !== null) {
    let url = match[1];
    if (url.includes('detail')) {
        if (url.startsWith('/')) {
            url = UQ_SETTINGS.SHOP_DOMAIN + url;
        } else if (!url.startsWith('http')) {
            const base = UQ_SETTINGS.AUCERTIFIED_PAGE_URL;
            url = base.endsWith('/') ? base + url : base + '/' + url;
        }
        url = url.split('?')[0].split('#')[0];
        links.add(url);
    }
  }
  
  const linkList = Array.from(links);
  Logger.log(`  -> 検出された詳細リンク候補: ${linkList.length}件`);

  let fixedCount = 0;
  products.forEach(p => {
    if ((!p.link_shop || p.link_shop === '') && 
        (p.name.includes('Certified') || p.name.includes('認定中古品'))) {
        
        const normName = uq_normalizeName(p.name); 
        
        const targetLink = linkList.find(link => {
            let normLink = uq_normalizeUrl(link).toLowerCase();
            normLink = normLink.replace(/detail/g, '')
                               .replace(/au_certified/g, 'certified')
                               .replace(/aucertified/g, 'certified')
                               .replace(/certified/g, '')
                               .replace(/[^a-z0-9]/g, '');
            return normLink.includes(normName) || (normName.length > 8 && normLink.includes(normName.substring(0, 8)));
        });

        if (targetLink) {
            p.link_shop = targetLink;
            fixedCount++;
        }
    }
  });
  
  Logger.log(`  -> リンク補完成功: ${fixedCount}件`);
}

/**
 * STEP 2: 価格データの取得とフィルタリング
 */
function uq_fetchAndFilterPriceData(html) {
  Logger.log('STEP 2: 価格データの特定・取得・フィルタリングを行います。');
  
  const jsMatch = html.match(UQ_SETTINGS.UTIL_JS_PATH_PATTERN);
  if (!jsMatch || !jsMatch[1]) return [];
  let jsUrl = jsMatch[1].startsWith('/') ? UQ_SETTINGS.BASE_DOMAIN + jsMatch[1] : jsMatch[1];
  
  const jsContent = uq_fetchPageHtml(jsUrl);
  if (!jsContent) return [];

  const jsonMatch = jsContent.match(UQ_SETTINGS.PRICE_JSON_PATH_PATTERN);
  if (!jsonMatch || !jsonMatch[1]) return [];
  
  let jsonUrl = jsonMatch[1];
  if (!jsonUrl.startsWith('http')) {
     if (jsonUrl.startsWith('/')) jsonUrl = UQ_SETTINGS.BASE_DOMAIN + jsonUrl;
     else jsonUrl = 'https://www.uqwimax.jp' + (jsonUrl.startsWith('json/') ? '/' : '/json/') + jsonUrl;
  }
  Logger.log(`価格JSON: ${jsonUrl}`);

  const jsonContent = uq_fetchPageHtml(jsonUrl);
  let prices = [];
  try {
    prices = JSON.parse(jsonContent);
    Logger.log(`全価格データ: ${prices.length}件`);
  } catch (e) {
    Logger.log(`JSONパースエラー: ${e.message}`);
    return [];
  }

  if (!UQ_SETTINGS.FILTER.ENABLED) return prices;

  const filtered = prices.filter(p => {
    const orderVal = p[UQ_SETTINGS.FILTER.KEY_ORDER];
    const optionVal = p[UQ_SETTINGS.FILTER.KEY_OPTION];
    const powerVal = p[UQ_SETTINGS.FILTER.KEY_POWER] || "";

    const isMnp = (orderVal === UQ_SETTINGS.FILTER.VAL_ORDER);
    const isAddOp = (optionVal === UQ_SETTINGS.FILTER.VAL_OPTION);
    const isTargetPlan = UQ_SETTINGS.FILTER.VAL_POWERS.some(target => powerVal.includes(target));
    
    return isMnp && isAddOp && isTargetPlan;
  });

  Logger.log(`フィルタリング結果: ${filtered.length}件`);

  const uniquePrices = [];
  const seen = new Set();

  filtered.forEach(p => {
    const name = p['体系表機種名'] || p['PC表示名'];
    const price = p['端末代金一括'];
    const key = `${name}_${price}`; 
    if (!seen.has(key)) {
      seen.add(key);
      uniquePrices.push(p);
    }
  });

  Logger.log(`重複排除後: ${uniquePrices.length}件`);
  return uniquePrices;
}

/**
 * STEP 3: 在庫情報の取得
 */
function uq_fetchInventoryData(products, targetPrices) {
  Logger.log('STEP 3: 対象機種の在庫情報を取得します。');
  
  const productMap = new Map();
  products.forEach(p => {
    if (p.name) productMap.set(uq_normalizeName(p.name), p);
  });

  const targetUrls = new Set();    
  const missingNames = [];

  targetPrices.forEach(priceItem => {
    const rawName = priceItem['体系表機種名'] || priceItem['PC表示名'];
    const name = uq_normalizeName(rawName);
    const htmlItem = productMap.get(name);
    
    // HTMLデータが見つかる場合
    if (htmlItem) {
      if (htmlItem.link_shop && htmlItem.link_shop.includes('shop.uqmobile.jp/detail/')) {
        targetUrls.add(htmlItem.link_shop);
      } else if (htmlItem.entryiphonepagelink || htmlItem.url) {
        // Certified の場合
        if (name.includes('certified')) {
           const guessedUrl = uq_guessShopUrl(rawName);
           targetUrls.add(guessedUrl);
        } else {
           // 新品の場合は紹介ページを探すロジック(今回は省略、必要なら復活)
           // ここでも念のため推測URLを追加してみる（iPhone 15等対策）
           const guessedUrl = uq_guessShopUrl(rawName);
           targetUrls.add(guessedUrl);
        }
      }
    } 
    // HTMLデータが見つからない場合（Certifiedなど）
    else {
      // 【修正】HTMLにない場合でも、URLを推測してチェック対象に追加する
      const guessedUrl = uq_guessShopUrl(rawName);
      targetUrls.add(guessedUrl);
      
      // ログには残す
      if (!missingNames.includes(rawName)) missingNames.push(rawName);
    }
  });

  if (missingNames.length > 0) {
    Logger.log(`紐付け失敗(一部): ${missingNames.slice(0, 5).join(', ')} など`);
    // これらは「HTMLリストとの紐付け」には失敗したが、上記の処理でURL推測により在庫チェック対象には入っている
  }

  const shopUrls = Array.from(targetUrls);
  
  if (shopUrls.length === 0) {
    Logger.log('対象の購入ページが見つかりません。');
    return {};
  }

  Logger.log(`巡回対象購入ページ(確定): ${shopUrls.length}件`);

  const shopPages = uq_fetchPagesParallel(shopUrls);
  
  const jsRequests = [];
  
  shopPages.forEach((html, index) => {
    if (!html) return;
    const match = html.match(UQ_SETTINGS.INVENTORY_JS_PATH_PATTERN);
    if (match && match[1]) {
      let jsPath = match[1];
      jsPath = jsPath.replace(/&amp;/g, '&');
      
      // パス解決
      if (!jsPath.startsWith('http')) {
        if (jsPath.startsWith('/')) {
          jsPath = UQ_SETTINGS.SHOP_DOMAIN + jsPath;
        } else {
          // 相対パスの場合
          const baseUrl = shopUrls[index].split('?')[0]; 
          const baseDir = baseUrl.substring(0, baseUrl.lastIndexOf('/') + 1);
          jsPath = baseDir + jsPath;
        }
      }

      // クエリ付与
      if (!jsPath.includes('contract=')) {
        const joinChar = jsPath.includes('?') ? '&' : '?';
        const contractType = UQ_SETTINGS.FILTER.VAL_ORDER === 'MNP' ? 'mnp' : 'new';
        jsPath += `${joinChar}contract=${contractType}`;
      }
      
      jsRequests.push({
        shopUrl: shopUrls[index], 
        jsUrl: jsPath
      });
    }
  });

  Logger.log(`特定された在庫JS: ${jsRequests.length}件。取得開始...`);

  const inventoryJsContents = uq_fetchPagesParallel(jsRequests.map(r => r.jsUrl));
  
  const inventoryDataMap = {};

  inventoryJsContents.forEach((jsCode, index) => {
    if (!jsCode) return;
    const shopUrl = jsRequests[index].shopUrl;
    const normalizedKey = uq_normalizeUrl(shopUrl);

    try {
      const plMatch = jsCode.match(/const\s+productList\s*=\s*({[\s\S]*?});/);
      const datMatch = jsCode.match(/var\s+dat\s*=\s*({[\s\S]*?});/);

      if (plMatch && plMatch[1] && datMatch && datMatch[1]) {
        const cleanPl = plMatch[1].replace(/\/\/.*/g, '');
        const cleanDat = datMatch[1].replace(/\/\/.*/g, '');

        const productList = new Function("return " + cleanPl)();
        const dat = new Function("return " + cleanDat)();

        let stockList = [];

        if (dat.storage_types && Array.isArray(dat.storage_types)) {
           dat.storage_types.forEach(storage => {
             const capName = storage.name; 
             if (storage.colorMap && Array.isArray(storage.colorMap)) {
               storage.colorMap.forEach(colorInfo => {
                 const colorName = colorInfo.name;
                 const sku = colorInfo.deviceCode; 
                 
                 const stockInfo = productList[sku];
                 let status = '不明';
                 if (stockInfo) {
                   status = stockInfo.supplement || (stockInfo.salesStatus == 1 ? '在庫あり' : '在庫なし');
                 }
                 
                 stockList.push({
                   capacity: capName,
                   color: colorName,
                   status: status
                 });
               });
             }
           });
        }

        if (stockList.length > 0) {
          inventoryDataMap[normalizedKey] = stockList;
        }
      }
    } catch (e) {
      Logger.log(`在庫JS解析エラー (${shopUrl}): ${e.message}`);
    }
  });

  Logger.log(`在庫情報生成完了: ${Object.keys(inventoryDataMap).length} ページ分`);
  return inventoryDataMap;
}


/**
 * STEP 4: 全データ統合
 */
function uq_mergeAllData(products, prices, inventoryMap) {
  Logger.log('STEP 4: データを統合します。');
  
  const productHtmlMap = new Map();
  products.forEach(p => {
    if (p.name) productHtmlMap.set(uq_normalizeName(p.name), p);
  });

  let matchCount = 0;

  const mergedList = prices.map(priceItem => {
    const merged = { ...priceItem };
    
    // 表示用名称から機種名を取得
    const dispName = priceItem['体系表機種名'] || priceItem['PC表示名'] || '';
    const nameNorm = uq_normalizeName(dispName);
    const htmlInfo = productHtmlMap.get(nameNorm);
    
    if (htmlInfo) {
      Object.keys(htmlInfo).forEach(k => {
        if (!merged.hasOwnProperty(k)) merged[k] = htmlInfo[k];
        else merged[`html_${k}`] = htmlInfo[k];
      });
    }

    // 在庫情報結合
    // 優先1: HTML情報のリンクから
    // 優先2: 推測URLから
    let stocks = null;
    
    if (htmlInfo && htmlInfo.link_shop) {
        const shopUrl = htmlInfo.link_shop;
        const normalizedKey = uq_normalizeUrl(shopUrl);
        if (inventoryMap[normalizedKey]) {
            stocks = inventoryMap[normalizedKey];
        }
    }
    
    if (!stocks) {
        const guessedUrl = uq_guessShopUrl(dispName);
        const guessedKey = uq_normalizeUrl(guessedUrl);
        if (inventoryMap[guessedKey]) {
            stocks = inventoryMap[guessedKey];
        }
    }

    if (stocks) {
      // 容量抽出ロジック（全角半角対応）
      const halfDispName = uq_toHalfWidth(dispName);
      const capMatch = halfDispName.match(/([0-9]+gb)/i);
      const targetCap = capMatch ? capMatch[1].toUpperCase() : null;

      const filteredStocks = stocks.filter(s => {
        if (!targetCap) return true;
        return uq_toHalfWidth(s.capacity).toUpperCase() === targetCap;
      });

      if (filteredStocks.length > 0) {
          merged['stock_msg'] = filteredStocks.map(s => `[${s.capacity}] ${s.color}: ${s.status}`).join('\n');
          // 整形用生データ
          merged['stocks_raw'] = filteredStocks;
          matchCount++;
      } else if (stocks.length > 0) {
          merged['stock_msg'] = stocks.map(s => `[${s.capacity}] ${s.color}: ${s.status}`).join('\n');
          merged['stocks_raw'] = stocks;
          matchCount++;
      } else {
          merged['stock_msg'] = '-';
      }
    } else {
      merged['stock_msg'] = '-';
    }

    return merged;
  });

  Logger.log(`在庫情報紐付け成功: ${matchCount} 件`);
  return mergedList;
}

/**
 * 【追加】出力用にデータを整形する関数
 */
function uq_formatDataForOutput(mergedData) {
  Logger.log('STEP 5: 出力用にデータを整形します。');
  
  const formatted = mergedData.map(item => {
    let rawName = item['体系表機種名'] || item['PC表示名'] || '';
    // 全角半角統一
    rawName = uq_toHalfWidth(rawName);

    let name = rawName;
    let capacity = '';

    // 装飾削除
    name = name.replace(/【UQ】/g, '');
    
    // 【修正】(X), (V) などの管理記号を削除
    name = name.replace(/\(X\)/gi, '')
               .replace(/\(V\)/gi, '')
               .replace(/（X）/gi, '')
               .replace(/（V）/gi, '');

    // 容量抽出 (優先: 在庫データから)
    if (item['stocks_raw'] && item['stocks_raw'].length > 0 && item['stocks_raw'][0].capacity) {
        // 機種名から容量を推測
        const capMatch = name.match(/(\d+GB)$/i);
        if (capMatch) {
            capacity = capMatch[1];
        } else {
            capacity = item['stocks_raw'][0].capacity;
        }
    } else {
        const capMatch = name.match(/(\d+GB)$/i);
        if (capMatch) {
          capacity = capMatch[1];
        }
    }

    // 機種名から容量部分を削除してトリム
    if (capacity) {
      name = name.replace(new RegExp(capacity + '$', 'i'), '').trim();
    }
    
    // 状態判定
    let condition = '新品';
    if (name.includes('Certified') || name.includes('認定中古品') || rawName.includes('Certified')) {
        condition = '中古';
        name = name.replace(/au\s*Certified/gi, '').replace(/（認定中古品）/g, '').trim();
    }

    // 在庫判定
    let stockStatus = '在庫なし';
    if (item['stocks_raw'] && Array.isArray(item['stocks_raw'])) {
        // 機種名の容量と一致する在庫データに絞り込む
        let targetStocks = item['stocks_raw'];
        if (capacity) {
            const capNorm = uq_toHalfWidth(capacity).toUpperCase();
            const filtered = item['stocks_raw'].filter(s => uq_toHalfWidth(s.capacity).toUpperCase() === capNorm);
            if (filtered.length > 0) {
                targetStocks = filtered;
            }
        }
        
        const hasStock = targetStocks.some(s => s.status.includes('在庫あり'));
        if (hasStock) {
            stockStatus = '在庫あり';
        }
    }

    // 価格マッピング
    const price = item['割賦代金'] || '';
    const priceDiscounted = item['端末代金一括'] || '';
    const priceReturn = item['スマトク_24回実質負担金'] || '';

    return {
      '機種名': name,
      '容量': capacity,
      '在庫': stockStatus,
      '端末価格': price,
      '割引後価格': priceDiscounted,
      '返却価格': priceReturn,
      '状態': condition
    };
  });

  // 重複排除ロジックを追加
  const uniqueData = [];
  const seen = new Set();

  formatted.forEach(item => {
    // 一意性を判定するキーを作成
    const key = `${item['機種名']}_${item['容量']}_${item['端末価格']}_${item['割引後価格']}_${item['状態']}`;
    
    if (!seen.has(key)) {
      seen.add(key);
      uniqueData.push(item);
    }
  });
  
  Logger.log(`整形・重複排除後: ${uniqueData.length}件 (元データ: ${formatted.length}件)`);
  
  return uniqueData;
}

/**
 * シート書き込み
 */
function uq_writeToSheet(formattedData) {
  Logger.log(`シート (${UQ_SETTINGS.SHEET_NAME}) へ書き込みます。`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(UQ_SETTINGS.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(UQ_SETTINGS.SHEET_NAME);

  sheet.clear();
  
  // ヘッダー書き込み
  const headers = UQ_SETTINGS.HEADERS;
  sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setBackground('#d9ead3')
        .setFontWeight('bold');

  // データ書き込み
  if (formattedData.length > 0) {
    const rows = formattedData.map(item => headers.map(h => item[h]));
    
    sheet.getRange(UQ_SETTINGS.START_ROW, UQ_SETTINGS.START_COL, rows.length, headers.length)
          .setValues(rows);
    sheet.autoResizeColumn(1);
  }
  Logger.log(`${formattedData.length}件 書き込み完了`);
}

/**
 * メイン処理
 */
function メイン処理_UQmobile端末価格取得() {
  Logger.log('メイン処理_UQmobile端末価格取得 を開始します。');
  try {
    const html = uq_fetchPageHtml(UQ_SETTINGS.TARGET_PAGE_URL);
    const products = uq_extractProductsFromHtml(html);
    
    uq_complementCertifiedLinks(products);

    const targetPrices = uq_fetchAndFilterPriceData(html);
    
    if (targetPrices.length === 0) {
      throw new Error('条件に合致する価格データが見つかりませんでした。');
    }

    const inventoryMap = uq_fetchInventoryData(products, targetPrices);
    const mergedData = uq_mergeAllData(products, targetPrices, inventoryMap);

    const formattedData = uq_formatDataForOutput(mergedData);

    uq_writeToSheet(formattedData);
    SpreadsheetApp.getActiveSpreadsheet().toast('処理が完了しました。', '完了', 5);

  } catch (e) {
    Logger.log(`エラー発生: ${e.message}`);
    Logger.log(e.stack);
    SpreadsheetApp.getActiveSpreadsheet().toast(`エラー: ${e.message}`, '失敗', 10);
  } finally {
    Logger.log('メイン処理_UQmobile端末価格取得 が終了しました。');
  }
}