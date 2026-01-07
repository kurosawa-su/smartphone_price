/**
 * @fileoverview auの価格・在庫情報をAPIから取得し、シートに書き込む (v3.3.46)
 *
 * @version 3.3.46 (2025-12-19)
 * @author (お客様の開発パートナー)
 *
 * @history
 * v3.3.46 (2025-12-19):
 * - [主変更] 関数名変更: メイン関数名を 'メイン処理_au価格在庫取得' から 'メイン処理_au端末価格取得' に変更。
 * v3.3.45 (2025-12-17):
 * - [主変更] 耐障害性向上: 在庫API (v2) で504エラー等が発生した際、全体を停止させず、そのチャンクのみスキップして継続するよう修正。
 * v3.3.44 (2025-12-17):
 * - [主変更] 中古ランク表示: 中古品の場合、商品名からランク情報（S, A, B等）を抽出し、「中古A」「中古B」のように表示するよう修正。
 * - [主変更] 機種名整形: シート書き込み直前に、機種名に含まれる半角括弧を全角括弧に変換する処理を追加。
 * v3.3.43 (2025-12-17):
 * - [主変更] 商品状態表記変更: 認定中古品のステータス定義を 'au Certified' から '中古' に変更。
 * v3.3.42 (2025-11-19):
 * - [主変更] スプレッドシート名の変更: 'AU_TARGET_SHEET_NAME' を 'au価格情報' から 'au端末一覧' に更新。
 */

// --- 定義 (v3.3.46) ---

// 【基本設計方針】: マジックナンバーの禁止

/**
 * [v3.0.0 新規] ステップ1: olsProductCode(ID)リスト取得用 端末情報API (Android)
 */
const AU_PRODUCT_SMARTPHONE_JSON_URL = 'https://www.au.com/content/dam/au-com/mobile/onlineshop/stock_list/json/product_smartphone.json';

/**
 * [v3.0.0 新規] ステップ1: olsProductCode(ID)リスト取得用 端末情報API (iPhone)
 */
const AU_PRODUCT_IPHONE_JSON_URL = 'https://www.au.com/content/dam/au-com/mobile/onlineshop/stock_list/json/product_iphone.json';

/**
 * [v3.3.0 新規] ステップ1: olsProductCode(ID)リスト取得用 端末情報API (中古)
 * (v3.3.0-debug の調査で特定)
 */
const AU_PRODUCT_CERTIFIED_JSON_URL = 'https://www.au.com/content/dam/au-com/mobile/onlineshop/stock_list/js/au_certified_product_data.js';

/**
 * [v3.3.17 復活] 中古品 (Certified) 用のデフォルトパス (フォールバック)
 */
const AU_CERTIFIED_DEFAULT_PATH = '/content/au-com/mobile/product/certified/';


/**
 * [v3.0.0 新規] ステップ3: 在庫・価格API (v2) のURLベース
 */
const AU_STOCK_V2_API_BASE_URL = 'https://www.au.com/bin/wcm/au-com/ols/product/v2.';

/**
 * [v3.0.1 新規] URL長制限エラー(Limit Exceeded)対策。
 * [v3.3.13 修正] チャンクサイズをより安全な 30件 に減らす。
 */
const AU_ID_CHUNK_SIZE = 30;

/**
 * [v3.3.15 新規] API連続アクセスを避けるための待機時間 (ミリ秒)
 * [v3.3.26 変更] 504 エラー対策のため、3000ms (3秒) に延長。
 */
const AU_API_CALL_INTERVAL_MS = 3000;

/**
 * [v3.3.25 新規] 実質負担額としてディープ割引価格 (例: 1円) を採用するための価格しきい値 (5,500円未満)
 */
const AU_DEEP_DISCOUNT_THRESHOLD = 5500;

/**
 * 書き出し対象のスプレッドシート名
 * [v3.3.42 変更] 'au価格情報' -> 'au端末一覧' に変更
 */
const AU_TARGET_SHEET_NAME = 'au端末一覧';

/**
 * [v3.0.0 新規] 403/401エラー対策の参照元URL (ドコモのスクリプトを参考)
 */
const AU_REFERER_URL = 'https://www.au.com/mobile/product/price/smartphone/';

/**
 * [v3.0.0 新規] 403/401エラー対策のオリジンURL (ドコモのスクリプトを参考)
 */
const AU_ORIGIN_URL = 'https://www.au.com';

/**
 * [v3.0.0 変更] APIアクセス用の強化版ヘッダー (ドコモのスクリプトを参考)
 */
const AU_REQUEST_HEADERS_V2 = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  'Referer': AU_REFERER_URL,
  'Origin': AU_ORIGIN_URL,
  'X-Requested-With': 'XMLHttpRequest'
};

/**
 * [v3.3.40 変更] シートに書き込むヘッダー行の定義 (列の追加と名称変更)
 */
const AU_NEW_HEADER_ROW = [[
  '機種名',             // productName (重複削除後)
  '容量',               // [v3.3.32] 容量を2列目に移動
  '在庫',               // [v3.3.33] '在庫状況' -> '在庫' に変更
  '端末価格',           // [v3.3.33] '本体価格(MNP)' -> '端末価格' に変更
  '割引後価格',         // [v3.3.40 新規] simpleCourseAfterDiscountPriceWithMnpInTax
  '返却価格',           // [v3.3.40 改名] 旧: 実質負担額 (残価控除後の総支払額)
  '状態'                // [v3.3.33] '商品状態' -> '状態' に変更
]];


// --- メイン関数 (v3.3.46) ---

/**
 * [v3.3.46 変更] auの価格・在庫情報をAPIから取得し、シートに書き込みます。
 */
function メイン処理_au端末価格取得() {
  Logger.log('メイン処理_au端末価格取得 (v3.3.46) を開始します。'); // v3.3.45 -> v3.3.46
  let cookieJar = null; // [v3.0.7] Cookie Jar 用の変数を定義
  
  try {
    // --- [v3.0.7 新規] ステップ0: 初期Cookieの取得 ---
    Logger.log('ステップ0: 初期セッションCookieの取得を開始します...');
    cookieJar = fetchInitialCookie_(); // (v3.0.8でバグ修正)
    if (cookieJar) {
      Logger.log('ステップ0: 完了。セッションCookie (新品Cookie) を取得しました。');
    } else {
      Logger.log('ステップ0: 完了。ただし、セッションCookieの取得に失敗しました。処理を続行します。');
    }

    // --- ステップ1: 端末情報APIから全IDリストとパスを取得 ---
    Logger.log('ステップ1: 端末情報API (product_*.json) から全IDとパスの取得を開始します...');
    const productInfoMap = fetchAllOlsProductCodesAndPaths_(cookieJar); // [v3.3.22] バグ修正版を呼び出し
    if (!productInfoMap || productInfoMap.size === 0) {
      throw new Error('端末情報APIから olsProductCode を1件も抽出できませんでした。処理を中断します。');
    }
    Logger.log(`ステップ1: 完了。合計 ${productInfoMap.size} 件のユニークなID/Path情報を取得しました。`);
    
    // --- [v3.2.0 新規] IDリストを「パス(path)」ごとにグループ化 ---
    Logger.log('   ... 取得したIDを機種別パス (currentPagePath) ごとにグループ化します...');
    const groupedByPath = new Map();
    // [v3.3.15] Map の値が {path, status} オブジェクトに変更されたため、 .entries() で受ける
    for (const [id, info] of productInfoMap.entries()) {
      const path = info.path; // [v3.3.15] path を info オブジェクトから取得
      if (!groupedByPath.has(path)) {
        groupedByPath.set(path, []); // 新しいパスのグループを作成
      }
      groupedByPath.get(path).push(id); // IDのみをグループに追加
    }
    Logger.log(`   ... ${groupedByPath.size} 個のユニークな機種別パス（グループ）に分類しました。`);


    // --- ステップ2: 在庫・価格API (v2) を「グループごと」に動的に呼び出し ---
    Logger.log('ステップ2: 在庫・価格API (v2....json) の「グループ別・分割呼び出し」を開始します (チャンクサイズ: %s)', AU_ID_CHUNK_SIZE);
    const allStockData = []; // 全てのAPI呼び出し結果を格納する配列
    
    let groupCount = 1;
    // グループ (例: '/.../galaxy_a25_5g/', ['SCG33SFA', ...]) ごとにループ
    for (const [path, idList] of groupedByPath.entries()) {
      Logger.log(`   ... グループ %s / %s を処理中 (Path: %s, ID: %s 件)`, groupCount, groupedByPath.size, path, idList.length);

      // (v3.0.1 チャンク対応) グループ内のIDリストが50件を超える場合、さらに分割
      for (let i = 0; i < idList.length; i += AU_ID_CHUNK_SIZE) {
        const idChunk = idList.slice(i, i + AU_ID_CHUNK_SIZE);
        Logger.log(`     ... チャンク %s / %s を処理中 (ID %s 件)`, (i / AU_ID_CHUNK_SIZE) + 1, Math.ceil(idList.length / AU_ID_CHUNK_SIZE), idChunk.length);

        // 分割したIDリストと、グループの「パス」でAPIを呼び出し
        const chunkData = fetchStockAndPriceData_(idChunk, cookieJar, path); // [v3.2.0] path を渡す
        
        if (Array.isArray(chunkData)) {
          allStockData.push(...chunkData);
        }

        // [v3.3.26] 504 エラー対策のため、3秒待機
        Utilities.sleep(AU_API_CALL_INTERVAL_MS); 
      }
      groupCount++;
    }
    
    if (allStockData.length === 0) {
      Logger.log('【WARN】在庫・価格API(v2)からデータを1件も取得できませんでした。');
    }
    Logger.log(`ステップ2: 完了。全グループ・全チャンクから合計 ${allStockData.length} 件の在庫・価格情報を取得しました。`);


    // --- ステップ3: データを解析しシートに書き出し ---
    Logger.log('ステップ3: データの解析とシートへの書き出しを開始します...');
    // [v3.3.32] 集約ロジックを実装した新バージョンを呼び出し
    parseAndWriteStockData_(allStockData, productInfoMap);
    Logger.log('ステップ3: 完了。シートへの書き出しが完了しました。');

    // 完了通知
    SpreadsheetApp.getActiveSpreadsheet().toast('au価格在庫一覧の更新が完了しました。', '処理完了 (v3.3.46)', 5);

  } catch (e) {
    // 【基本設計方針】: エラーハンドリング
    Logger.log('【ERROR】メイン処理 (メイン処理_au端末価格取得) 中に予期せず例外が発生しました。');
    Logger.log('エラーメッセージ: %s', e.message);
    if (e.fileName) Logger.log('ファイル: %s', e.fileName);
    if (e.lineNumber) Logger.log('行番号: %s', e.lineNumber);
    if (e.stack) Logger.log('スタックトレース: \n%s', e.stack);
    // エラー通知
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(`エラー: ${e.message}`, '処理失敗 (v3.3.46)', 10);
    } catch(toastErr) {
      Logger.log('【ERROR】エラー通知(toast)の表示に失敗しました: %s', toastErr.message);
    }

  } finally {
    Logger.log('メイン処理_au端末価格取得 (v3.3.46) が終了しました。');
  }
}

// --- ヘルパー関数群 (v3.3.46) ---

/**
 * [v3.0.8-debug 変更]
 * ステップ0: メインのHTMLページにアクセスし、サーバーから発行されるセッションCookieを取得（キャプチャ）する。
 * @returns {string | null} サーバーから返された 'Set-Cookie' ヘッダーの連結文字列。取得失敗時は null。
 * @private
 */
function fetchInitialCookie_() {
  const url = AU_REFERER_URL; // HTMLページ (price/smartphone/)
  
  // ヘッダーは User-Agent と 強化ヘッダー を使用
  const options = {
    'method': 'get',
    'headers': AU_REQUEST_HEADERS_V2, // 既存の強化ヘッダーを流用
    'muteHttpExceptions': true
  };

  let response;
  try {
    response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log(`【WARN】初期Cookieの取得に失敗しました。HTTP ${responseCode} (URL: ${url})`);
      return null;
    }

    const headers = response.getHeaders();

    // [v3.0.8-debug 修正]
    const setCookieValue = headers['Set-Cookie'];

    if (setCookieValue) {
      let cookies = [];
      if (Array.isArray(setCookieValue)) {
        cookies = setCookieValue;
      } else {
        cookies = [setCookieValue];
      }

      const cookieString = cookies.map(cookie => cookie.split(';')[0]).join('; ');
      Logger.log('   ... Set-Cookie を受信: %s', cookieString.substring(0, 100) + '...');
      return cookieString;
    } else {
      Logger.log('【WARN】サーバーから Set-Cookie ヘッダーが返されませんでした。(HTMLページ: %s)', url);
      return null;
    }
  } catch (e) {
    Logger.log(`【ERROR】fetchInitialCookie_ 実行中にエラー: ${e.message}`);
    return null; 
  }
}


/**
 * [v3.3.44 変更]
 * 端末情報API (smartphone.json, iphone.json, certified.js) にアクセスし、
 * すべての `olsProductCode` と、それが属する `productDetailUrl` (パス) 、「商品状態」を抽出し、Mapとして返します。
 * @param {string | null} cookieJar - [v3.0.7] ステップ0で取得したCookie文字列
 * @returns {Map<string, {path: string, status: string, nameRaw: string}>} { id => {path, status, nameRaw} } のMap
 * @private
 */
function fetchAllOlsProductCodesAndPaths_(cookieJar) {
  // [v3.3.15] Mapの値を {path: string, status: string} のオブジェクトに変更
  const productInfoMap = new Map();
  // [v3.3.0 変更] 中古品 (Certified) を読み込み対象に追加
  const targetUrls = [
    { name: 'Smartphone', url: AU_PRODUCT_SMARTPHONE_JSON_URL },
    { name: 'iPhone', url: AU_PRODUCT_IPHONE_JSON_URL },
    { name: 'Certified', url: AU_PRODUCT_CERTIFIED_JSON_URL }
  ];
  
  // [v3.2.1 新規] 'productDetailUrl' が null だった場合のフォールバックパスを準備
  const newPhoneDefaultPath = extractPathFromUrl_(AU_REFERER_URL);
  if (!newPhoneDefaultPath) {
    throw new Error(`デフォルトパスの抽出に失敗しました (${AU_REFERER_URL})。処理を中断します。`);
  }
  Logger.log('   ... 新品用のデフォルトパスを準備しました: %s', newPhoneDefaultPath);
  // [v3.3.17] 中古用のデフォルトパス (v3.3.14 相当)
  Logger.log('   ... 中古用のデフォルトパスを準備しました: %s', AU_CERTIFIED_DEFAULT_PATH); 

  // [v3.0.7] ヘッダーオブジェクトをコピーして、動的にCookieを追加
  const options = {
    'method': 'get',
    'headers': { ...AU_REQUEST_HEADERS_V2 }, // 強化ヘッダーをコピー
    'muteHttpExceptions': true
  };
  if (cookieJar) {
    options.headers['Cookie'] = cookieJar; 
    Logger.log('   ... fetchAllOlsProductCodesAndPaths_: Cookieヘッダーを設定しました。');
  }

  // [v3.3.0 変更] targetUrls をループ
  for (const target of targetUrls) {
    Logger.log(`   ... 端末情報API (%s) をフェッチします: %s`, target.name, target.url);
    let response;
    let content;
    let jsonData;

    try {
      response = UrlFetchApp.fetch(target.url, options);
      content = response.getContentText('UTF-8');
      const responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        Logger.log(`【WARN】端末情報API (${target.name}) の取得に失敗しました。HTTPステータス: ${responseCode}`);
        continue;
      }
      
      // [v3.3.1 修正]
      // 1. 先頭の /* ... */ コメントブロックを削除
      if (content.trim().startsWith('/*')) {
        Logger.log('   ... JSコメントブロック (/* ... */) を検出。削除します。');
        content = content.replace(/^\s*\/\*[\s\S]*?\*\/\s*/, '');
      }
      
      // 2. [v3.3.0] 'var ... =' の部分を削除
      if (content.trim().startsWith('var')) {
        Logger.log('   ... JS変数形式 (var ...) を検出。JSONに変換します。');
        content = content.replace(/^var\s+.*?=\s*/, '').replace(/;$/, '');
      }

      // [v3.3.2 修正] 'const ... =' の部分を削除 (v3.3.1のログで 'const auCe...' が検出されたため)
      if (content.trim().startsWith('const')) {
        Logger.log('   ... JS変数形式 (const ...) を検出。JSONに変換します。');
        content = content.replace(/^const\s+.*?=\s*/, '').replace(/;$/, '');
      }

      // [v3.3.5 修正]
      // 'const ... =' 除去後に、さらにコメント (/*...*/) が残っている場合 (v3.3.4 ログ)
      content = content.trim(); // 先頭の空白を除去
      if (content.startsWith('/*')) {
        Logger.log('   ... 2番目のJSコメントブロック (/* ... */) を検出。削除します。');
        content = content.replace(/^\s*\/\*[\s\S]*?\*\/\s*/, '');
      }

      // [v3.3.4 修正]
      if (target.url.endsWith('.js')) {
        // .js ファイルはキーが引用符で囲まれていない (例: { productCode: ... })
        // 'new Function()' が JSON エラーをスローしたため (v3.3.3 ログ)、
        // 正規表現でキーに引用符を追加し、'JSON.parse()' で安全に解析する。
        Logger.log('   ... .js ファイルのため、JS Object Literal -> JSON 文字列に変換します...');
        // 'productCode:' -> '"productCode":'
        // (注: 'http:' などの値に含まれるコロンを誤変換しないよう、キー名のパターンを限定する)
        let jsonContent = content.replace(/([{,]\s*)([a-zA-Z0-9_]+)\s*:/g, '$1"$2":'); 
        
        // [v3.3.5 修正] Trailing Comma (例: '...}, ]' や '..., ]') を除去する
        // (v3.3.4 のログ 'Unexpected token ]' 対策)
        Logger.log('   ... Trailing Comma (余分な末尾カンマ) を除去します...');
        // '...},]' -> '...}]'
        // '...",]' -> '..."]'
        jsonContent = jsonContent.replace(/,(\s*[}\]])/g, '$1'); 

        jsonData = JSON.parse(jsonContent);

        // [v3.3.19 削除] v3.3.18 の調査用ログコードを削除 (役目を終えたため)

      } else {
        jsonData = JSON.parse(content);
      }

      Logger.log(`   ... 端末情報API (${target.name}) の解析に成功しました。`);

    } catch (e) {
      Logger.log(`【ERROR】端末情報API (${target.name}) のフェッチまたは解析中にエラー: ${e.message}`);
      Logger.log('      ... 内容(一部): %s', content ? content.substring(0, 100) : 'N/A');
      continue;
    }

    // [v3.3.10 修正]
    // 504 (Gateway Timeout) エラー対策。
    // [v3.3.26] 504 エラー対策のため、3秒待機
    Logger.log('   ... サーバー負荷軽減のため 3秒 待機します...');
    Utilities.sleep(AU_API_CALL_INTERVAL_MS);

    // JSON構造の解析 (v3.2.1)
    if (Array.isArray(jsonData)) {
      // jsonData = [ (端末), (端末), ... ]
      for (const item of jsonData) {
        
        // [v3.3.44] API名から、このアイテムが「新品」か「中古」かを定義する
        let productStatus = (target.name === 'Certified') ? '中古' : '新品';

        // [v3.3.44] 中古の場合、商品名からランク情報があれば抽出して付加する
        // (例: "iPhone... (ランクA)" -> status: "中古A")
        if (productStatus === '中古') {
          const nameRaw = item.productName || item.title || item.petName || '';
          // ランクA, Rank A, (A) などのパターンを抽出 (アルファベット部分を取り出す)
          const rankMatch = nameRaw.match(/(?:ランク|Rank)[\s:：]*([SABC][+]?)/i);
          if (rankMatch) {
             productStatus = `中古${rankMatch[1].toUpperCase()}`;
          }
        }

        // [v3.3.19 修正] 'Bug C' 修正
        // 新品APIは 'productDetailUrl'、中古APIは 'url' (v3.3.18 調査ログ) を参照
        const itemUrl = item.productDetailUrl || item.url;
        // [v3.3.19 修正] 'Bug D' 修正 (extractPathFromUrl_ が相対パス 'item.url' を処理可能)
        let itemPath = extractPathFromUrl_(itemUrl);

        // [v3.3.17 修正] v3.3.16 のログ分析に基づき、v3.3.15 の「単一パス化」ロジックを破棄。
        // v3.3.14 の「分岐ロジック」に戻す。
        if (!itemPath) {
          if (target.name === 'Certified') {
            itemPath = AU_CERTIFIED_DEFAULT_PATH;
            Logger.log('   ... [INFO] (中古) item.url が見つかりません。中古用デフォルトパス (%s) を使用します。(productCode: %s)', itemPath, (item.productCode || 'N/A'));
          } else {
            itemPath = newPhoneDefaultPath;
            Logger.log('   ... [INFO] (新品) productDetailUrl が見つかりません。新品用デフォルトパス (%s) を使用します。(productCode: %s)', itemPath, (item.productCode || 'N/A'));
          }
        }
        
        // [v3.3.32 新規] 生の機種名・容量・カラー名を保存
        const nameRaw = item.productName || item.title || item.petName || 'N/A';
        
        // [v3.3.22 修正] 'Bug G' 修正。
        // API (target.name) ごとに、正しいネスト構造を解析する。
        if (target.name === 'Certified') {
          // --- Certified API Parser ---
          if (Array.isArray(item.detail)) {
            for (const capacityItem of item.detail) {
              if (Array.isArray(capacityItem.colorsAndOlsCode)) {
                for (const color of capacityItem.colorsAndOlsCode) {
                  if (color.olsProductCode) {
                    // [v3.3.32] 生の機種名と容量を保存
                    productInfoMap.set(color.olsProductCode, { path: itemPath, status: productStatus, nameRaw: nameRaw, capacityRaw: capacityItem.capacity });
                  }
                }
              }
            }
          }
          // --- End of Certified Parser ---
        } else {
          // --- New (Smartphone/iPhone) API Parser ---
          
          // colorVariations から探す (新品API用)
          if (Array.isArray(item.colorVariations)) {
            for (const color of item.colorVariations) {
              if (color.olsProductCode) {
                // [v3.3.32] 容量情報は color.name に含まれているため、ここでは color.name を容量の生のデータとして保存
                productInfoMap.set(color.olsProductCode, { path: itemPath, status: productStatus, nameRaw: nameRaw, capacityRaw: color.name });
              }
            }
          }
          
          // colorsAndCapacities からも探す (新品API用)
          if (Array.isArray(item.colorsAndCapacities)) {
            for (const colorCap of item.colorsAndCapacities) {
              if (Array.isArray(colorCap.capacities)) {
                for (const capacity of colorCap.capacities) {
                  const currentId = capacity.value || capacity.olsProductCode;
                  if (currentId) {
                    // [v3.3.32] 容量情報は capacity.capacityName/capacity.value に含まれている
                    productInfoMap.set(currentId, { path: itemPath, status: productStatus, nameRaw: nameRaw, capacityRaw: capacity.capacityName || capacity.value });
                  }
                }
              }
            }
          }
          // --- End of New Parser ---
        } // [v3.3.22] End of if/else
        
        // [v3.3.22 削除] v3.3.21 以前の、トップレベルID (item.productCode) を
        // 誤って取得していたパーサーロジック (L443-L451@v3.3.21) を削除。

      } // for (item of jsonData)
    } // if (Array.isArray)
  } // for (url of targetUrls)

  return productInfoMap;
}

/**
 * [v3.3.45 変更]
 * olsProductCode のリスト (チャンク) と「機種別パス」を元に v2 API のURLを構築し、
 * 在庫と価格のデータを取得して返します。(Cookieヘッダーを追加)
 * [v3.3.45] エラー発生時は例外を投げず、空配列を返してスキップさせる。
 * @param {string[]} productIds - olsProductCode の IDリスト (チャンク)
 * @param {string | null} cookieJar - ステップ0で取得した「新品Cookie」
 * @param {string} path - このIDグループが属する機種別パス (例: '/.../galaxy_a25_5g/')
 * @returns {any[]} v2 APIから取得した在庫・価格データの配列。エラー時は空配列。
 * @private
 */
function fetchStockAndPriceData_(productIds, cookieJar, path) {
  if (!productIds || productIds.length === 0) {
    Logger.log('【WARN】fetchStockAndPriceData_ に渡されたIDリストが空です。スキップします。');
    return [];
  }

  // [v3.2.0 変更] チャンクのURLを「機種別パス」を使って動的に構築
  // [v3.2.1 修正] path が null だった場合に備え、'path' をログ出力
  if (!path) {
     Logger.log('【ERROR】fetchStockAndPriceData_ に渡された path が null または undefined です。API呼び出しをスキップします。');
     return [];
  }
  // [v3.2.3] extractPathFromUrl_ が末尾 '/' を保証するため、 'currentPagePath=${path}' でOK
  const dynamicSuffix = `.json?currentPagePath=${path}`;
  const dynamicUrl = AU_STOCK_V2_API_BASE_URL + productIds.join('.') + dynamicSuffix;
  
  Logger.log(`    ... 動的URL (v2 API) をフェッチします: ${AU_STOCK_V2_API_BASE_URL}[...${productIds.length}件のID...]${dynamicSuffix}`);
  
  // [v3.0.7] ヘッダーオブジェクトをコピーして、動的にCookieを追加
  const options = {
    'method': 'get', 
    'headers': { ...AU_REQUEST_HEADERS_V2 }, // 強化ヘッダーをコピー
    'muteHttpExceptions': true
  };

  // [v3.3.20 修正] 'Bug E' 修正。Refererを動的パスに上書き
  options.headers['Referer'] = AU_ORIGIN_URL + path;
  
  // [v3.3.22 修正] 'Bug F' (Cookie動的再取得) ロジックを削除。新品Cookieの使い回しに。
  if (cookieJar) {
    options.headers['Cookie'] = cookieJar; 
    Logger.log('    ... fetchStockAndPriceData_: Cookieヘッダーと動的Refererを設定しました。');
  } else {
     Logger.log('    ... fetchStockAndPriceData_: 動的Refererのみ設定しました (Cookieなし)。');
  }

  let response;
  let content;
  let jsonData;

  try {
    response = UrlFetchApp.fetch(dynamicUrl, options);
    content = response.getContentText('UTF-8');
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log(`【ERROR】在庫・価格API (v2) の取得に失敗しました。HTTPステータス: ${responseCode}`);
      Logger.log(`   ... URL(一部): ${dynamicUrl.substring(0, 200)}...`);
      // [v3.3.45] エラー時は空配列を返してスキップさせる
      return []; 
    }

    jsonData = JSON.parse(content);
    Logger.log('    ... 在庫・価格API (v2) の解析に成功しました。');

  } catch (e) {
    Logger.log(`【ERROR】在庫・価格API (v2) のフェッチまたは解析中にエラー: ${e.message}`);
    // [v3.3.45] エラー時は空配列を返してスキップさせる (throwしない)
    return []; 
  }

  // v2.2.8 のログで「配列」であることを確認済み
  if (Array.isArray(jsonData)) {
    return jsonData;
  } else {
    Logger.log('【ERROR】在庫・価格API (v2) のレスポンが、想定された「配列」ではありませんでした。');
    Logger.log('   ... 取得したデータ型: %s', typeof jsonData);
    // 形式エラーの場合はスキップ
    return [];
  }
}

/**
 * [v3.3.44 変更] 取得した在庫・価格データを解析し、シートに書き込みます。
 * (機種名・容量・状態ごとの集約と在庫統合判定を導入。新列対応)
 * @param {any[]} stockDataArray - v2 APIから取得したデータの配列
 * @param {Map<string, {path: string, status: string, nameRaw: string, capacityRaw: string}>} productInfoMap - IDと商品状態などを突合するためのMap
 * @private
 */
function parseAndWriteStockData_(stockDataArray, productInfoMap) {
  if (!stockDataArray || stockDataArray.length === 0) {
    Logger.log('【WARN】書き込むデータが0件のため、シート処理をスキップします。');
    return;
  }

  // --- スプレッドシートの準備 ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(AU_TARGET_SHEET_NAME);
  const isNewSheet = sheet.getLastRow() === 0; // [v3.3.38] 初回実行判定フラグ

  if (!sheet) {
    sheet = ss.insertSheet(AU_TARGET_SHEET_NAME);
    Logger.log('シート "%s" が存在しなかったため、新規作成しました。', AU_TARGET_SHEET_NAME);
  } else {
    Logger.log('既存のシート "%s" を使用します。', AU_TARGET_SHEET_NAME);
  }

  // [v3.3.38 修正] シート全体ではなく、データ範囲のみをクリアし、設定を維持
  if (sheet.getLastRow() > 1) {
    // 2行目から最終行までをクリア
    sheet.getRange(2, 1, sheet.getLastRow() - 1, AU_NEW_HEADER_ROW[0].length).clearContent();
    Logger.log('既存のデータ範囲 (%s行目以降) をクリアしました。', 2);
  }


  // --- ヘッダー行の書き込み ---
  sheet.getRange(1, 1, AU_NEW_HEADER_ROW.length, AU_NEW_HEADER_ROW[0].length)
       .setValues(AU_NEW_HEADER_ROW)
       .setFontWeight('bold');
  Logger.log('ヘッダー行を書き込みました。');

  // --- 集約用 Map の準備 ---
  // キー: [クリーン機種名]_[容量]_[商品状態]
  const aggregatedData = new Map();

  // データ行の処理 (集約)
  for (const item of stockDataArray) {
    
    try {
      // 1. データ整形 (v3.3.31 のロジックを流用)
      let productNameRaw = item.productName || 'N/A';
      const colorNameRaw = (item.colorVariationName !== undefined) ? item.colorVariationName : (item.productName || 'N/A');
      
      // 容量の分離 (v3.3.31 修正: 全角括弧対応)
      const capacityMatch = colorNameRaw.match(/（(.*?)）$/);
      const capacity = capacityMatch ? capacityMatch[1].trim() : 'N/A';
      const colorName = colorNameRaw.replace(/\s*（.*?）$/, '').trim();
      
      // [v3.3.39 新規] カラー名自体のクリーンアップ (全角文字/スペースの統一)
      let colorNameCleaned = colorName;
      if (colorNameCleaned !== 'N/A') {
        colorNameCleaned = colorNameCleaned.replace(/　/g, ' ')
                                         .replace(/[０-９Ａ-Ｚａ-ｚ]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
                                         .replace(/\s{2,}/g, ' ').trim();
      }


      // 機種名のクリーンアップ (v3.3.31/v3.3.34/v3.3.35 修正)
      let productName = productNameRaw;
      if (productName !== 'N/A') {
        // [v3.3.35 修正] 1. クリーンアップ処理の冒頭で全角を半角に統一
        // 全角数字、全角アルファベット、全角括弧、全角スペースなどを半角に統一
        productName = productName.replace(/　/g, ' ') 
                                 .replace(/（/g, '(') 
                                 .replace(/）/g, ')') 
                                 .replace(/ＧＢ/g, 'GB') 
                                 .replace(/ＴＢ/g, 'TB') 
                                 .replace(/(\d)Ｇ/g, '$1G') 
                                 .replace(/[０-９Ａ-Ｚａ-ｚ]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)); 
        
        // 2. 容量を削除 (半角になった容量文字列を使って削除)
        if (capacity !== 'N/A' && capacity.length > 0) {
          const escapedCapacity = capacity.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
          // [v3.3.35] 括弧やスペースが残らないよう処理
          productName = productName.replace(new RegExp(`\\s*\\(?${escapedCapacity}\\)?`, 'ig'), '').trim(); 
        }
        
        // 3. カラー名を削除 (大文字小文字を区別しない & 正規表現エスケープ)
        if (colorNameCleaned.length > 0 && colorNameCleaned !== 'N/A') {
          // クリーンアップ後のカラー名を使用して機種名から削除
          const escapedColorName = colorNameCleaned.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
          // [v3.3.36 修正] Bug O 対策: 削除前に機種名からも全角スペースを半角に統一
          productName = productName.replace(/　/g, ' ');

          // [v3.3.35] 前後のスペースを徹底的に削除
          productName = productName.replace(new RegExp(`\\s*${escapedColorName}\\s*`, 'ig'), ' ').trim();
        }

        // 4. [v3.3.35] iPhoneの特殊な全角/半角表記を削除
        productName = productName.replace(/\(PRODUCT\)RED/ig, '').trim();
        productName = productName.replace(/（ＰＲＯＤＵＣＴ）ＲＥＤ/ig, '').trim();

        // 5. [v3.3.31 修正] 中古品特有の冗長な文字列やスペースを削除 (Bug M 修正)
        productName = productName.replace(/\(認定中古品\)/ig, '').trim();
        productName = productName.replace(/au Certified/ig, '').trim();
        productName = productName.replace(/\s{2,}/g, ' ').trim();

        // 6. [v3.3.44 新規] 最後に半角括弧を全角括弧に変換 (表示用)
        productName = productName.replace(/\(/g, '（').replace(/\)/g, '）');
      }
      
      // 商品状態の取得
      const olsProductCode = item.olsProductCode || 'N/A';
      const info = productInfoMap.get(olsProductCode);
      const productStatus = info ? info.status : '不明';

      // 在庫状況の判定 (統合用フラグ)
      const isAvailable = (item.olsStatus === '在庫あり' || (item.olsSalesAttribute && item.olsSalesAttribute.stockQuantity > 0));

      // 2. 集約キーの作成
      const aggregationKey = `${productName}_${capacity}_${productStatus}`;

      // 3. データ集約の実行
      if (aggregatedData.has(aggregationKey)) {
        // 既存のデータがある場合: 在庫フラグを統合する
        const existingData = aggregatedData.get(aggregationKey);
        // [v3.3.32] 一つでも在庫があれば 'isAvailable' は true にする
        existingData.isAvailable = existingData.isAvailable || isAvailable;
      } else {
        // 新しいデータの場合: 初回のエントリとして登録 (価格はこのエントリのものを代表値とする)
        const salesPrice = item.salesPrice;
        let priceMnp = 'N/A';
        let formattedBurden = 'N/A';
        let discountedPriceFormatted = 'N/A'; // [v3.3.40 新規] 割引後価格の格納用

        // 本体価格(MNP)の抽出 (v3.3.24)
        if (salesPrice && salesPrice.simpleCoursePriceWithMnpInTax !== undefined) {
          priceMnp = salesPrice.simpleCoursePriceWithMnpInTax;
        } else if (item.price !== undefined) {
          priceMnp = item.price;
        }

        // [v3.3.40 新規] 割引後価格の抽出 (新列用: simpleCourseAfterDiscountPriceWithMnpInTax)
        if (salesPrice && salesPrice.simpleCourseAfterDiscountPriceWithMnpInTax !== undefined) {
          const discountedPriceValue = salesPrice.simpleCourseAfterDiscountPriceWithMnpInTax;
          if (typeof discountedPriceValue === 'number') {
            discountedPriceFormatted = `¥${discountedPriceValue.toLocaleString()}`;
          } else {
            discountedPriceFormatted = discountedPriceValue;
          }
        }

        // 実質負担額（返却価格）の抽出 (v3.3.41 修正)
        if (salesPrice) {
          const discountedPrice = salesPrice.simpleCourseAfterDiscountPriceWithMnpInTax;
          // 優先度1: ディープ割引 (1円)
          if (discountedPrice !== undefined && discountedPrice < AU_DEEP_DISCOUNT_THRESHOLD) {
            formattedBurden = `¥${discountedPrice.toLocaleString()}`;
          // 優先度2: 残価控除後の総支払額 (iPhoneなど)
          } else if (Array.isArray(salesPrice.residualValueInstallmentList) && 
                      salesPrice.residualValueInstallmentList.length > 0 && 
                      salesPrice.residualValueInstallmentList[0].mnpResidualValueTotalInstallmentPaymentInTax !== undefined) {
            formattedBurden = `¥${salesPrice.residualValueInstallmentList[0].mnpResidualValueTotalInstallmentPaymentInTax.toLocaleString()}`;
            // [v3.3.41 修正] 優先度3: その他の割引後価格 (37000円) は削除。空欄 ('N/A') のまま保持する。
          }
        }
        
        aggregatedData.set(aggregationKey, {
          productName: productName,
          capacity: capacity,
          discountedPriceFormatted: discountedPriceFormatted, // [v3.3.40] 新しい値をMapに追加
          productStatus: productStatus,
          isAvailable: isAvailable, // 初回在庫フラグ
          priceMnp: priceMnp,
          formattedBurden: formattedBurden
        });
      }

    } catch (parseErr) {
      Logger.log(`【WARN】データ1件の集約中にエラーが発生しました。スキップします。 (ID: ${item.olsProductCode || 'N/A'}) - ${parseErr.message}`);
    }
  } // for (item of stockDataArray)


  // 4. 集約後のデータをシート出力用の配列に変換
  const outputData = [];
  for (const [key, data] of aggregatedData.entries()) {
    // 最終的な在庫状況の文字列を決定
    const finalStockStatus = data.isAvailable ? '在庫あり' : '在庫なし';
    
    // [v3.3.40] 新しい7列の並び順でプッシュ (実質負担額 -> 返却価格 に変更)
    outputData.push([
      data.productName,              // 0: 機種名
      data.capacity,                 // 1: 容量
      finalStockStatus,              // 2: 在庫 (統合判定後)
      data.priceMnp,                 // 3: 端末価格 (MNP)
      data.discountedPriceFormatted, // 4: 割引後価格 (新規)
      data.formattedBurden,          // 5: 返却価格 (旧: 実質負担額)
      data.productStatus             // 6: 状態
    ]);
  }


  // --- シートへの一括書き込み ---
  if (outputData.length > 0) {
    sheet.getRange(2, 1, outputData.length, AU_NEW_HEADER_ROW[0].length)
         .setValues(outputData);
    
    // [v3.3.38 修正] 列幅の自動調整は、新規シートの場合のみ行う (手動設定を維持するため)
    const isNewSheet = sheet.getLastRow() === outputData.length + 1; // last rowがデータ行数+ヘッダー行の場合
    if (isNewSheet) {
      sheet.autoResizeColumns(1, AU_NEW_HEADER_ROW[0].length);
      Logger.log('新規シートのため、列幅を自動調整しました。');
    }
    
    Logger.log('%s 行のデータをシートに書き込みました。', outputData.length);
  } else {
    Logger.log('【WARN】抽出可能なデータが0件でした。');
  }
}

/**
 * [v3.3.19 変更]
 * 完全なURL (例: 'https://www.au.com/mobile/product/price/smartphone/') または
 * 相対URL (例: '/content/au-com/mobile/product/certified/...') から、
 * v2 API の 'currentPagePath' に必要なパス部分を抽出します。
 * @param {string} fullUrl - 'https://' から始まるURL、または '/content/...' から始まる相対パス。
 * @returns {string | null} 抽出したパス。
 * @private
 */
function extractPathFromUrl_(fullUrl) {
  try {
    if (!fullUrl) {
      return null;
    }

    // [v3.3.19 修正] 'Bug D' 修正。相対パスを優先的に処理。
    if (fullUrl.startsWith('/content/au-com/')) {
      const pathCleaned = fullUrl.split('?')[0]; // クエリパラメータを除去
      return pathCleaned.endsWith('/') ? pathCleaned : pathCleaned + '/';
    }

    // [v3.3.3 修正] 外部リンクは WARN を出さずに null を返す
    if (!fullUrl.startsWith(AU_ORIGIN_URL)) {
      return null;
    }

    // 'https://www.au.com' の部分を削除
    const pathCleaned = fullUrl.replace(AU_ORIGIN_URL, '').split('?')[0];
    
    // [v3.3.19] パターンC (レガシーパス, /mobile/...)
    if (pathCleaned.startsWith('/mobile/') || pathCleaned.startsWith('/iphone/')) { // パターン1
      const contentPath = '/content/au-com' + pathCleaned;
      
      // [v3.2.2 修正] ... 末尾の '/' は「必須」
      return contentPath.endsWith('/') ? contentPath : contentPath + '/';
    }

    // [v3.3.4 修正] 外部リンク以外の「想定外」の au.com パスの場合のみ WARN を出す
    Logger.log('【WARN】extractPathFromUrl_: 想定外の au.com パス形式です: %s', fullUrl);
    // [v3.2.2 修正] 不明な形式でも、末尾の '/' を保証する
    return pathCleaned.endsWith('/') ? pathCleaned : pathCleaned + '/';

  } catch (e) {
    Logger.log('【ERROR】extractPathFromUrl_ でエラー: %s', e.message);
    return null;
  }
}