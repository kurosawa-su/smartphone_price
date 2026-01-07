/**
 * スマホ価格比較集計スクリプト
 * * 概要:
 * 指定された各社のデータ取得関数を順次実行した後、
 * 「端末一覧」を含む各シートからデータを収集し、まとめシートを作成します。
 * * * 比較項目:
 * 1. 端末価格（定価）の最安値
 * 2. 割引後価格（実質）の最安値と割引率
 * 3. 返却価格の最安値と割引率
 * * * 仕様:
 * - 各社データ取得関数の連続実行（表のみ作成モードあり）
 * - 排他制御なし
 * - 在庫なしデータの除外
 * - 価格文字列の数値変換処理
 * - 「中古」端末の行背景色変更
 * - 区切り線（返却割引率の右）と複合ソート（機種名＞状態）の実装
 * - SpreadsheetApp.flush() による画面描画更新
 * - 完了通知メールの送信機能（宛先指定可）
 * - 実行ログ記録機能
 */

// ===========================================================================
// 定数定義 (マジックナンバー防止)
// ===========================================================================

// 出力先のシート名
const SUMMARY_SHEET_NAME = 'スマホ価格比較';

// ログ記録用のシート名
const LOG_SHEET_NAME = '実行ログ';

// 処理対象シートを判定するキーワード
const TARGET_SHEET_KEYWORD = '端末一覧';

// 在庫切れ判定のキーワード
const OUT_OF_STOCK_VALUE = '在庫なし';

// 入力データのヘッダー項目名（各シート共通）
const HEADER_MODEL_NAME = '機種名';
const HEADER_CAPACITY = '容量';
const HEADER_STOCK = '在庫';
const HEADER_PRICE_FULL = '端末価格';
const HEADER_PRICE_DISCOUNT = '割引後価格';
const HEADER_PRICE_RETURN = '返却価格';
const HEADER_CONDITION = '状態';

// エラー時のセル背景色（警告色）
const ERROR_BG_COLOR = '#FFCCCC'; // 薄い赤

// 特定条件（中古）のセル背景色
const USED_BG_COLOR = '#FFF2CC'; // 薄い黄色

// 完了通知メールの送信先（複数指定可能）
// 例: const NOTIFICATION_RECIPIENTS = ['user1@example.com', 'user2@example.com'];
// ※ 空配列 [] の場合は、スクリプト実行者本人に送信されます。
const NOTIFICATION_RECIPIENTS = ['ogura@starcraft-n.co.jp', 'kurosawa@starcraft-n.co.jp'];

// ===========================================================================
// メイン処理 (エントリーポイント)
// ===========================================================================

/**
 * [メイン関数] 各社データ取得を実行後、スマホ価格比較まとめシートを作成・更新します。
 */
function メイン処理_スマホ価格比較作成() {
  executeMainProcess_(true);
}

/**
 * [メイン関数] データ取得をスキップし、既存のシートデータからスマホ価格比較まとめシートを作成・更新します。
 */
function メイン処理_スマホ価格比較表のみ作成() {
  executeMainProcess_(false);
}

// ===========================================================================
// 共通処理ロジック
// ===========================================================================

/**
 * 共通メイン処理
 * @param {boolean} needsFetch - データ取得を行うかどうか (true:行う, false:行わない)
 */
function executeMainProcess_(needsFetch) {
  const actionLabel = needsFetch ? 'データ取得および価格比較作成' : '価格比較表のみ作成';
  console.log('処理を開始します: ' + actionLabel);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    // ---------------------------------------------------------
    // 0. 事前処理: 各社のデータ取得関数を順次実行 (needsFetchがtrueの場合のみ)
    // ---------------------------------------------------------
    if (needsFetch) {
      console.log('各社の最新データ取得を開始します...');
      
      // ※ 以下の関数は同一プロジェクト内に存在する必要があります
      メイン処理_IIJmio端末価格取得();
      メイン処理_au端末価格取得();
      メイン処理_SoftBank端末価格取得();
      メイン処理_Ymobile端末価格取得();
      メイン処理_Ymobile_Yahoo店端末価格取得();
      メイン処理_docomo端末価格取得();
      メイン処理_Apple端末価格取得();
      メイン処理_UQmobile端末価格取得();
      メイン処理_ahamo端末価格取得();
      メイン処理_Rakuten端末価格取得();
      
      console.log('全社のデータ取得が完了しました。集計処理へ移行します。');
    } else {
      console.log('データ取得をスキップし、集計処理を開始します。');
    }


    // ---------------------------------------------------------
    // 1. 全対象シートからデータを収集
    // ---------------------------------------------------------
    // 戻り値: { headerCompanies: string[], dataMap: Map }
    const gatheredData = gatherAllSheetData_(ss);
    
    // データが見つからなかった場合のガード節
    if (gatheredData.headerCompanies.length === 0) {
      console.warn('処理対象となるシートが見つかりませんでした。');
      ui.alert('対象シート（シート名に「' + TARGET_SHEET_KEYWORD + '」を含む）が見つかりませんでした。');
      return;
    }


    // ---------------------------------------------------------
    // 2. 収集したデータを集計し、最安値比較ロジックを適用
    // ---------------------------------------------------------
    // 戻り値: { outputRows: any[][], companyColumns: number }
    const aggregatedResult = aggregateAndCalculate_(gatheredData);


    // ---------------------------------------------------------
    // 3. 結果をシートに出力し、書式を設定（ソート含む）
    // ---------------------------------------------------------
    writeToSummarySheet_(ss, aggregatedResult.outputRows, gatheredData.headerCompanies);


    // ---------------------------------------------------------
    // 4. ログ記録 (実行日時と内容を記録)
    // ---------------------------------------------------------
    recordExecutionLog_(ss, actionLabel);


    // ---------------------------------------------------------
    // 5. 完了処理
    // ---------------------------------------------------------
    // シートへの変更を即座に適用し、UI描画を更新させる（処理中表示対策）
    SpreadsheetApp.flush();

    // メール送信機能
    try {
      // 指定された宛先がある場合は結合して使用、なければ実行ユーザー
      let recipient = '';
      if (NOTIFICATION_RECIPIENTS.length > 0) {
        recipient = NOTIFICATION_RECIPIENTS.join(',');
      } else {
        recipient = Session.getActiveUser().getEmail();
      }

      if (recipient) {
        const subject = '【スマホ価格比較】集計処理完了のお知らせ';
        const body = 'スマホ価格比較シートの作成処理が正常に完了しました。\n' +
                     '処理内容: ' + actionLabel + '\n\n' +
                     '確認はこちら:\n' + ss.getUrl();
        
        MailApp.sendEmail(recipient, subject, body);
        console.log('完了メールを送信しました: ' + recipient);
      }
    } catch (mailError) {
      // メール送信エラーはメイン処理の致命的なエラーとはせず、警告ログのみ残す
      console.warn('メール送信に失敗しました: ' + mailError.message);
    }

    console.log('処理が正常に完了しました。');
    ui.alert(actionLabel + 'が完了しました。\n完了メールを送信しました。');

  } catch (e) {
    // エラーハンドリング
    console.error('【ERROR】メイン処理でエラーが発生しました: ' + e.message);
    console.error('Stack: ' + e.stack);
    SpreadsheetApp.getUi().alert('エラーが発生しました。ログを確認してください。\n' + e.message);
  }
}

// ===========================================================================
// サブ関数 (ログ記録)
// ===========================================================================

/**
 * 実行ログをシートに記録します。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - 対象スプレッドシート
 * @param {string} content - 実行内容の説明
 */
function recordExecutionLog_(ss, content) {
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  
  // シートが存在しない場合は新規作成し、ヘッダーを設定
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    // ログシートは末尾に追加するのが望ましいが、insertSheetはアクティブシートの隣などに作られる場合があるため
    // 必要であれば moveActiveSheet 等で移動も検討可能。今回は作成のみ。
    logSheet.appendRow(['実行日時', '実行内容']);
    logSheet.getRange(1, 1, 1, 2).setBackground('#EFEFEF').setFontWeight('bold');
    logSheet.setFrozenRows(1);
    // 列幅調整
    logSheet.setColumnWidth(1, 150); // 日時
    logSheet.setColumnWidth(2, 300); // 内容
  }

  // 日時フォーマット
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  
  // 追記
  logSheet.appendRow([formattedDate, content]);
}

// ===========================================================================
// サブ関数 (データ収集・整形)
// ===========================================================================

/**
 * 全シートを走査し、対象シートからデータを抽出してMapに格納します。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - 対象スプレッドシート
 * @returns {Object} 会社名リストと、キーごとに整形されたデータMap
 */
function gatherAllSheetData_(ss) {
  const sheets = ss.getSheets();
  const dataMap = new Map(); // Key: "機種_容量_状態", Value: { baseInfo: {}, prices: { companyName: { full:..., discount:..., return:... } } }
  const companyNames = [];

  // 全シートをループ
  for (const sheet of sheets) {
    const sheetName = sheet.getName();

    // 対象シート判定（キーワードが含まれているか）
    if (sheetName.indexOf(TARGET_SHEET_KEYWORD) === -1) {
      continue;
    }

    // 会社名の抽出（シート名からキーワードを除去）
    const companyName = sheetName.replace(TARGET_SHEET_KEYWORD, '').trim();
    companyNames.push(companyName);
    
    console.log('データ収集開始: ' + companyName + ' (' + sheetName + ')');

    // シートデータの取得
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) continue; // ヘッダーのみまたは空の場合はスキップ

    // ヘッダー行（1行目）から列インデックスを特定
    const headerRow = values[0];
    const colIdx = getColumnIndices_(headerRow);

    // 必須列が見つからない場合はログを出してスキップ
    if (!colIdx) {
      console.error('【ERROR】必須列が見つかりません: ' + sheetName);
      continue;
    }

    // 2行目以降のデータ行をループ
    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      // 在庫チェック: 在庫なしならスキップ
      const stockVal = row[colIdx.stock];
      if (stockVal === OUT_OF_STOCK_VALUE) {
        continue;
      }

      // キー項目の取得
      const model = row[colIdx.model];
      const capacity = row[colIdx.capacity];
      const condition = row[colIdx.condition];

      // 一意なキーを生成 (機種_容量_状態)
      const uniqueKey = [model, capacity, condition].join('_');

      // 価格データのパース（数値化）
      const priceFull = parsePrice_(row[colIdx.priceFull]);
      const priceDiscount = parsePrice_(row[colIdx.priceDiscount]);
      const priceReturn = parsePrice_(row[colIdx.priceReturn]);

      // Mapへの登録
      if (!dataMap.has(uniqueKey)) {
        dataMap.set(uniqueKey, {
          model: model,
          capacity: capacity,
          condition: condition,
          prices: {} // 会社ごとの価格情報を格納
        });
      }

      const entry = dataMap.get(uniqueKey);
      entry.prices[companyName] = {
        full: priceFull,
        discount: priceDiscount,
        returnVal: priceReturn // returnは予約語のため returnVal とする
      };
    }
  }

  return {
    headerCompanies: companyNames,
    dataMap: dataMap
  };
}

/**
 * ヘッダー行から必要な列のインデックスを取得します。
 * @param {Array} headerRow - ヘッダー行の配列
 * @returns {Object|null} 列インデックスのオブジェクト。失敗時はnull
 */
function getColumnIndices_(headerRow) {
  const idx = {
    model: headerRow.indexOf(HEADER_MODEL_NAME),
    capacity: headerRow.indexOf(HEADER_CAPACITY),
    stock: headerRow.indexOf(HEADER_STOCK),
    priceFull: headerRow.indexOf(HEADER_PRICE_FULL),
    priceDiscount: headerRow.indexOf(HEADER_PRICE_DISCOUNT),
    priceReturn: headerRow.indexOf(HEADER_PRICE_RETURN),
    condition: headerRow.indexOf(HEADER_CONDITION)
  };

  // 必須項目が一つでも欠けていたらnullを返す（簡易チェック）
  for (const key in idx) {
    if (idx[key] === -1) return null;
  }
  return idx;
}

/**
 * 価格文字列を数値に変換します。
 * 例: "¥10,000" -> 10000, "10,000円" -> 10000
 * @param {string|number} priceRaw - 元の価格データ
 * @returns {number|null} 変換後の数値。変換不可の場合はnull
 */
function parsePrice_(priceRaw) {
  if (priceRaw === '' || priceRaw === undefined || priceRaw === null) return null;
  if (typeof priceRaw === 'number') return priceRaw;

  // 文字列の場合、数字以外を除去してパース
  const numStr = String(priceRaw).replace(/[^0-9]/g, '');
  if (numStr === '') return null;
  
  return parseInt(numStr, 10);
}

// ===========================================================================
// サブ関数 (集計・計算)
// ===========================================================================

/**
 * 集計データから最安値比較テーブルを生成します。
 * @param {Object} gatheredData - gatherAllSheetData_の戻り値
 * @returns {Object} 出力用2次元配列
 */
function aggregateAndCalculate_(gatheredData) {
  const dataMap = gatheredData.dataMap;
  const companies = gatheredData.headerCompanies;
  const outputRows = [];

  // Mapの全エントリをループ
  for (const [key, info] of dataMap) {
    const rowData = [];

    // 1. 基本情報 (A, B, C列)
    rowData.push(info.model);
    rowData.push(info.capacity);
    rowData.push(info.condition);

    // --- 最安値計算の準備 ---
    let minFull = Infinity;
    let minDiscount = Infinity;
    let minReturn = Infinity;

    let companiesMinFull = [];
    let companiesMinDiscount = [];
    let companiesMinReturn = [];

    // 各社の価格をチェックして最安値を探索
    for (const company of companies) {
      const priceObj = info.prices[company];
      if (!priceObj) continue; // その会社のデータがない場合はスキップ

      // 定価(端末価格)の最安チェック
      if (priceObj.full !== null) {
        if (priceObj.full < minFull) {
          minFull = priceObj.full;
          companiesMinFull = [company]; // 更新
        } else if (priceObj.full === minFull) {
          companiesMinFull.push(company); // 同額なら追加
        }
      }

      // 実質(割引後価格)の最安チェック
      if (priceObj.discount !== null) {
        if (priceObj.discount < minDiscount) {
          minDiscount = priceObj.discount;
          companiesMinDiscount = [company];
        } else if (priceObj.discount === minDiscount) {
          companiesMinDiscount.push(company);
        }
      }

      // 返却価格の最安チェック
      if (priceObj.returnVal !== null) {
        if (priceObj.returnVal < minReturn) {
          minReturn = priceObj.returnVal;
          companiesMinReturn = [company];
        } else if (priceObj.returnVal === minReturn) {
          companiesMinReturn.push(company);
        }
      }
    }

    // 値が更新されなかった(データなし)場合の処理
    if (minFull === Infinity) minFull = null;
    if (minDiscount === Infinity) minDiscount = null;
    if (minReturn === Infinity) minReturn = null;

    // --- 割引率の計算 ---
    // 割引率 = 1 - (比較価格 / 定価最安)
    // ※定価最安が存在し、0でない場合のみ計算
    let rateDiscount = null;
    if (minFull && minDiscount !== null) {
      rateDiscount = 1 - (minDiscount / minFull);
    }

    let rateReturn = null;
    if (minFull && minReturn !== null) {
      rateReturn = 1 - (minReturn / minFull);
    }

    // --- 比較列の格納 (D〜K列) ---
    
    // 定価
    rowData.push(minFull);
    rowData.push(companiesMinFull.join(', '));
    
    // 実質
    rowData.push(minDiscount);
    rowData.push(companiesMinDiscount.join(', '));
    rowData.push(rateDiscount); // 実質割引率

    // 返却
    rowData.push(minReturn);
    rowData.push(companiesMinReturn.join(', '));
    rowData.push(rateReturn); // 返却割引率

    // --- 各社詳細データの格納 (L列以降) ---
    for (const company of companies) {
      const p = info.prices[company];
      if (p) {
        rowData.push(p.full);
        rowData.push(p.discount);
        rowData.push(p.returnVal);
      } else {
        rowData.push(null); // データなし
        rowData.push(null);
        rowData.push(null);
      }
    }

    outputRows.push(rowData);
  }

  return {
    outputRows: outputRows
  };
}

// ===========================================================================
// サブ関数 (出力・書式)
// ===========================================================================

/**
 * データをシートに出力し、書式設定を行います。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - 対象スプレッドシート
 * @param {Array} rows - 出力データ配列
 * @param {Array} companies - 会社名リスト（ヘッダー生成用）
 */
function writeToSummarySheet_(ss, rows, companies) {
  // シート取得または新規作成
  let sheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (sheet) {
    sheet.clear(); // 既存の場合はクリア
  } else {
    sheet = ss.insertSheet(SUMMARY_SHEET_NAME); // 新規作成
  }

  // --- ヘッダー行の作成 ---
  const headerRow = [
    HEADER_MODEL_NAME, HEADER_CAPACITY, HEADER_CONDITION, // A-C
    '定価(最安)', '定価最安の会社', // D, E
    '実質(最安)', '実質最安の会社', '実質割引率', // F, G, H
    '返却(最安)', '返却最安の会社', '返却割引率'  // I, J, K
  ];

  // 各社ごとの列ヘッダーを追加
  for (const company of companies) {
    headerRow.push(company + '_' + HEADER_PRICE_FULL);
    headerRow.push(company + '_' + HEADER_PRICE_DISCOUNT);
    headerRow.push(company + '_' + HEADER_PRICE_RETURN);
  }

  // データがない場合の処理
  if (rows.length === 0) {
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    SpreadsheetApp.getUi().alert('条件に一致するデータがありませんでした。');
    return;
  }

  // --- まとめて書き込み ---
  // ヘッダー + データ
  const allValues = [headerRow, ...rows];
  const totalRows = allValues.length;
  const totalCols = allValues[0].length;
  const range = sheet.getRange(1, 1, totalRows, totalCols);
  range.setValues(allValues);

  // --- 書式設定 ---
  
  // 1. 固定ヘッダーの装飾
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange.setBackground('#D9EAD3'); // 薄い緑
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1); // 行固定
  sheet.setFrozenColumns(3); // A-C列（端末情報）固定

  // 2. 「中古」行の背景色設定
  // データ部分（2行目以降）に対して背景色を計算して一括適用
  const dataRowsCount = rows.length;
  if (dataRowsCount > 0) {
    const backgrounds = [];
    
    for (const row of rows) {
      const condition = String(row[2]); // C列「状態」
      const rowColors = new Array(totalCols).fill(null); // デフォルト（設定なし）
      
      // 「中古」という文字列が含まれているか判定
      if (condition.indexOf('中古') !== -1) {
        rowColors.fill(USED_BG_COLOR);
      }
      backgrounds.push(rowColors);
    }
    
    // 背景色を一括設定
    sheet.getRange(2, 1, dataRowsCount, totalCols).setBackgrounds(backgrounds);
  }

  // 3. 数値フォーマット（通貨）の設定
  // 対象列: D(定価), F(実質), I(返却), および各社の3列ごと
  // 通貨フォーマットを設定する列インデックス(1始まり)のリストを作成
  // H(実質割引率), K(返却割引率)はパーセント
  
  const numRows = rows.length;
  if (numRows > 0) {
    // 列ごとのフォーマット定義
    // A:1, B:2, C:3
    // D:4(円), E:5
    // F:6(円), G:7, H:8(%)
    // I:9(円), J:10, K:11(%)
    // L以降: 3列ごとに(円, 円, 円)
    
    // 円マークフォーマット範囲の設定
    const currencyFormat = '¥#,##0';
    const percentFormat = '0.0%';

    // D列(4) 定価最安
    sheet.getRange(2, 4, numRows, 1).setNumberFormat(currencyFormat);
    // F列(6) 実質最安
    sheet.getRange(2, 6, numRows, 1).setNumberFormat(currencyFormat);
    // H列(8) 実質割引率
    sheet.getRange(2, 8, numRows, 1).setNumberFormat(percentFormat);
    // I列(9) 返却最安
    sheet.getRange(2, 9, numRows, 1).setNumberFormat(currencyFormat);
    // K列(11) 返却割引率
    sheet.getRange(2, 11, numRows, 1).setNumberFormat(percentFormat);

    // 各社の詳細列の設定 (L列=12列目以降)
    // 会社数 * 3列分
    const startColCompany = 12;
    const totalCompanyCols = companies.length * 3;
    if (totalCompanyCols > 0) {
      sheet.getRange(2, startColCompany, numRows, totalCompanyCols).setNumberFormat(currencyFormat);
    }
  }

  // 4. 列幅の自動調整 (データ量が多いと時間がかかるため、主要列のみ調整または一括調整)
  sheet.autoResizeColumns(1, totalCols);

  // 5. 区切り線（返却割引率と詳細データの間）
  // K列(11列目)の右側に実線を引く
  // 引数: top, left, bottom, right, vertical, horizontal, color, style
  sheet.getRange(1, 11, totalRows, 1)
    .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID);

  // 6. 複合ソート実行
  // 優先順位1: 機種名(1列目) 降順
  // 優先順位2: 状態(3列目) 昇順
  if (dataRowsCount > 0) {
    sheet.getRange(2, 1, dataRowsCount, totalCols).sort([
      {column: 1, ascending: false}, // 機種名 (第一優先)
      {column: 3, ascending: true}   // 状態 (第二優先)
    ]);
  }
}