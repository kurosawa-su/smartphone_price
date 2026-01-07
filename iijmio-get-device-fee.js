/**
 * @fileoverview IIJmioの端末API(JSON)から端末情報を取得し、シートに書き込む
 * @version 3.6.3 (2025-12-19)
 * - 11:47の要望に基づき、メイン関数名を「メイン処理_IIJmio端末価格取得」に変更。
 * @version 3.6.2 (2025-12-17)
 * - 13:49の指摘に基づき、状態判定（中古/新品）の参照先を「端末名」から「メーカー名(device.manufacturer)」に変更。
 * メーカー名に「品」が含まれる場合（例：「未使用品」など）を「中古」とし、それ以外を「新品」とする。
 * @version 3.6.1 (2025-12-17)
 * - 状態判定ロジックの動作検証強化版。
 * @version 3.6.0 (2025-12-17)
 * - ヘッダー構成を8列化。['機種名', '容量', '在庫', '端末価格', '割引後価格', '返却価格', '状態', 'メーカー名']
 * - 機種名の半角括弧()を全角（）に置換。
 * - 返却価格列を追加（空欄）。
 */

// スクリプト全体で使用する設定値（定数）
const IIJMIO_SETTINGS = {
  // スクレイピング対象のURL
  TARGET_URL: 'https://www.iijmio.jp/call/api?serviceType=common&apiName=terminal&action=list&range=all',

  // 書き込み対象のシート名
  SHEET_NAME: 'iijmio端末一覧',

  // 書き込みヘッダー（全8列）
  // ['機種名', '容量', '在庫', '端末価格', '割引後価格', '返却価格', '状態', 'メーカー名']
  HEADER_ROW: [['機種名', '容量', '在庫', '端末価格', '割引後価格', '返却価格', '状態', 'メーカー名']],

  // データ書き込み開始行（ヘッダーが1行目なので、データは2行目から）
  START_ROW: 2,
  // データ書き込み開始列（A列から）
  START_COL: 1,

  // ロックのタイムアウト時間（ミリ秒）。30秒
  LOCK_TIMEOUT: 30000,
  // 古いロックとみなす有効期限（ミリ秒）。20分
  LOCK_EXPIRATION: 20 * 60 * 1000,
  
  // シート上のエラー値（#N/Aなど）を発見した場合のハイライト色
  ERROR_CELL_COLOR: '#f8d7da', // 薄い赤色
  
  // ブロック対策として追加するUser-Agentヘッダー
  USER_AGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
};


/**
 * メイン処理：キャンペーン端末情報を取得し、シートに書き込みます。
 * この関数をGASエディタから手動で実行してください。
 */
function メイン処理_IIJmio端末価格取得() {
  // 処理開始ログ
  Logger.log('メイン処理_IIJmio端末価格取得 を開始します。');
  
  // 排他制御用のロック変数を初期化
  let lock = null;

  try {
    // --- 排他制御 開始 ---
    Logger.log('スクリプトの排他制御（ロック）を開始します。');
    lock = LockService.getScriptLock();
    lock.waitLock(IIJMIO_SETTINGS.LOCK_TIMEOUT);
    Logger.log('ロックを取得しました。処理を実行します。');
    // --- 排他制御 完了 ---

    // --- JSON API 取得処理 ---
    Logger.log(`対象API URL (${IIJMIO_SETTINGS.TARGET_URL}) からJSONデータを取得します。`);
    
    // User-Agentヘッダーを追加して、ブラウザからのアクセスを装う
    const params = {
      'method': 'get',
      'headers': {
        'User-Agent': IIJMIO_SETTINGS.USER_AGENT
      },
      'muteHttpExceptions': true // 403(ブロック)などのエラーでも例外を投げず、応答を取得する
    };
    
    // APIからデータを取得
    const response = UrlFetchApp.fetch(IIJMIO_SETTINGS.TARGET_URL, params);
    const jsonText = response.getContentText('UTF-8');
    
    // レスポンスコードをログに出力
    const responseCode = response.getResponseCode();
    Logger.log(`HTTPステータスコード: ${responseCode}`);

    // 取得失敗時（200 OK以外）は処理を中断
    if (responseCode !== 200) {
      throw new Error(`JSON APIの取得に失敗しました。ステータスコード: ${responseCode}`);
    }
    Logger.log('JSON APIの取得に成功しました。');


    // --- JSON解析処理 ---
    const outputData = []; // シートに書き込むための2次元配列 (8列)
    Logger.log('JSONデータの解析を開始します...');
    
    // 取得したテキストをJSONオブジェクトに変換
    const jsonResponse = JSON.parse(jsonText);

    if (jsonResponse && jsonResponse.result && jsonResponse.result.result && Array.isArray(jsonResponse.result.result)) {
      const devices = jsonResponse.result.result;
      Logger.log(`${devices.length} 件の端末データをJSONから取得しました。`);

      // 端末データをループ処理
      for (const device of devices) {
        try {
          // --- 0. 除外判定 ---
          if (device.category === 'closedsale') continue;
          if (device.manufacturer === 'ネットチャート') continue;

          // 1. 機種名 (旧: 端末名)
          const rawName = device.name || null;
          
          // []削除と、半角括弧()を全角（）に置換
          let name = rawName ? rawName.replace(/\[.*?\]/g, '').trim() : null;
          if (name) {
            name = name.replace(/\(/g, '（').replace(/\)/g, '）');
          }

          // 2. 容量 (GB/TB付与)
          const rawCapacity = IIJMIO_getMetadata(device.metadata, 'ROM') || null;
          let capacity = null;
          if (rawCapacity === '1000') {
            capacity = '1TB';
          } else if (rawCapacity) {
            capacity = rawCapacity + 'GB';
          }

          // 3. 在庫
          let totalStock = 0;
          if (device.colors && Array.isArray(device.colors)) {
            try {
              for (const color of device.colors) {
                if (color.number_of_stock && typeof color.number_of_stock === 'number' && color.number_of_stock > 0) {
                  totalStock += color.number_of_stock;
                }
              }
            } catch(e) {
              totalStock = -1;
            }
          }
          const stockDisplay = (totalStock > 0) ? '在庫あり' : (totalStock === 0) ? '在庫なし' : 'エラー';
          
          // 4. 端末価格
          const terminalPrice = IIJMIO_getMetadata(device.metadata, 'discount_after_1') || 
                                (device.payments && Array.isArray(device.payments) && device.payments[0] 
                                  ? device.payments[0].amount : null) 
                                || null;

          // 5. 割引後価格 (旧: 実質負担額)
          const discountPrice = IIJMIO_getMetadata(device.metadata, 'voice_set_1') || null;

          // 6. 返却価格 (空欄)
          const returnPrice = '';

          // 8. メーカー名 (APIから取得した生のメーカー名)
          const maker = device.manufacturer || null;

          // 7. 状態 (判定ロジック)
          // 13:49の指摘に基づき修正: 判定対象を「メーカー名(maker)」に変更
          // メーカー名に「品」が含まれていれば「中古」、そうでなければ「新品」
          const condition = (maker && maker.includes('品')) ? '中古' : '新品';

          // 出力配列に追加
          if (name) {
            outputData.push([name, capacity, stockDisplay, terminalPrice, discountPrice, returnPrice, condition, maker]);
          }

        } catch (e) {
          Logger.log(`【WARN】解析エラー: ${e.message} (端末: ${device.name})`);
        }
      }
      
      Logger.log(`JSONの解析が完了しました。${outputData.length} 件のデータを正常に抽出しました。`);

    } else {
      Logger.log('【ERROR】JSONの構造が予期したものではありません');
    }


    // --- スプレッドシート出力処理 ---
    Logger.log(`シート (${IIJMIO_SETTINGS.SHEET_NAME}) への書き込みを開始します。`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(IIJMIO_SETTINGS.SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(IIJMIO_SETTINGS.SHEET_NAME);
    }
    sheet.clearContents();
    
    // ヘッダー行を書き込み (8列)
    sheet.getRange(1, 1, IIJMIO_SETTINGS.HEADER_ROW.length, IIJMIO_SETTINGS.HEADER_ROW[0].length)
         .setValues(IIJMIO_SETTINGS.HEADER_ROW);

    // 抽出したデータがある場合のみ、シートに書き込み
    if (outputData.length > 0) {
      sheet.getRange(IIJMIO_SETTINGS.START_ROW, IIJMIO_SETTINGS.START_COL, outputData.length, outputData[0].length)
           .setValues(outputData);
      Logger.log(`${outputData.length} 件のデータをシートに書き込みました。`);
    }

    ss.toast('端末一覧の更新が完了しました。', '処理完了', 5);

  } catch (err) {
    Logger.log('【ERROR】処理中に重大なエラーが発生しました。');
    Logger.log(`詳細: ${err.message}`);
  } finally {
    if (lock) lock.releaseLock();
    Logger.log('メイン処理_IIJmio端末価格取得 が終了しました。');
  }
}


/**
 * テスト関数：ロジック検証用（モックデータ使用）
 */
function IIJMIO_test_parsingLogic() {
  Logger.log('IIJMIO_test_parsingLogic を開始します。');

  // テスト用モックデータ (メーカー名での判定を確認)
  const mockDevices = [
    {
      name: "iPhone 15 [128GB]",
      metadata: [{name: 'ROM', value: '128'}],
      manufacturer: "Apple", // 「品」なし -> 新品
      colors: [{number_of_stock: 10}]
    },
    {
      name: "iPhone SE (第3世代) [64GB]",
      metadata: [{name: 'ROM', value: '64'}],
      manufacturer: "未使用品", // 「品」あり -> 中古
      colors: [{number_of_stock: 5}]
    },
    {
      name: "Reno9 A",
      metadata: [{name: 'ROM', value: '128'}],
      manufacturer: "美品", // 「品」あり -> 中古
      colors: [{number_of_stock: 0}]
    }
  ];

  Logger.log('--- テスト実行: モックデータの解析 ---');
  const outputData = [];

  for (const device of mockDevices) {
    try {
      // 1. 機種名
      const rawName = device.name || null;
      let name = rawName ? rawName.replace(/\[.*?\]/g, '').trim() : null;
      if (name) name = name.replace(/\(/g, '（').replace(/\)/g, '）');

      // 2. 容量
      const rawCapacity = IIJMIO_getMetadata(device.metadata, 'ROM');
      let capacity = (rawCapacity === '1000') ? '1TB' : (rawCapacity ? rawCapacity + 'GB' : null);

      // 3. 在庫
      let totalStock = 0;
      if (device.colors) {
        for (const c of device.colors) totalStock += c.number_of_stock || 0;
      }
      const stockDisplay = (totalStock > 0) ? '在庫あり' : '在庫なし';

      // 4-6. 価格系 (テストでは省略)
      const terminalPrice = 10000;
      const discountPrice = 5000;
      const returnPrice = '';

      // 8. メーカー名 (APIから取得した値)
      const maker = device.manufacturer;

      // 7. 状態 (判定ロジック)
      // メーカー名に「品」が含まれていれば「中古」、そうでなければ「新品」
      const condition = (maker && maker.includes('品')) ? '中古' : '新品';

      if (name) {
        outputData.push([name, capacity, stockDisplay, terminalPrice, discountPrice, returnPrice, condition, maker]);
      }
    } catch (e) {
      Logger.log(e);
    }
  }

  Logger.log('--- 抽出データ確認 (全8列) ---');
  Logger.log('ヘッダー: 機種名, 容量, 在庫, 端末価格, 割引後価格, 返却価格, 状態, メーカー名');
  outputData.forEach(row => {
    Logger.log(JSON.stringify(row));
  });
  Logger.log('-----------------------------------');
}

/**
 * metadata検索用ヘルパー
 */
function IIJMIO_getMetadata(metadataArray, key) {
  if (!Array.isArray(metadataArray) || !key) return null;
  const item = metadataArray.find(i => i && i.name === key);
  return item ? item.value : null;
}