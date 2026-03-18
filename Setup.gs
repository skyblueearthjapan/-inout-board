/**
 * 初期設定ヘルパー
 * 
 * 初回セットアップ時にこのファイルの関数を実行してください。
 */

/**
 * スクリプトプロパティを初期設定
 * 
 * 使い方：
 * 1. この関数内の値を実際の環境に合わせて編集
 * 2. GASエディタで initializeProperties を実行
 * 3. デプロイ
 */
function initializeProperties() {
  const props = PropertiesService.getScriptProperties();
  
  // ========================================
  // ここを編集してください
  // ========================================
  const config = {
    // スプレッドシートID（URLの /d/ と /edit の間の部分）
    SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',
    
    // シート名
    SHEET_NAME: '外出ホワイトボード',
    
    // 日付セル
    DATE_CELL: 'D2',
    
    // ヘッダー範囲
    HEADER_RANGE: 'A3:E3',
    
    // データ範囲
    DATA_RANGE: 'A4:E19'
  };
  // ========================================
  
  // プロパティを設定
  props.setProperties(config);
  
  Logger.log('スクリプトプロパティを設定しました:');
  Logger.log(JSON.stringify(config, null, 2));
  
  return '設定完了';
}

/**
 * 現在のスクリプトプロパティを確認
 */
function viewProperties() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  
  Logger.log('現在のスクリプトプロパティ:');
  Logger.log(JSON.stringify(all, null, 2));
  
  return all;
}

/**
 * スクリプトプロパティをクリア
 */
function clearProperties() {
  const props = PropertiesService.getScriptProperties();
  props.deleteAllProperties();
  
  Logger.log('スクリプトプロパティをクリアしました');
  
  return 'クリア完了';
}

/**
 * 接続テスト - シートにアクセスできるか確認
 */
function testConnection() {
  try {
    const config = SheetService.getConfig();
    
    if (!config.spreadsheetId || config.spreadsheetId === 'YOUR_SPREADSHEET_ID_HERE') {
      throw new Error('SPREADSHEET_ID が設定されていません。initializeProperties を実行してください。');
    }
    
    const ss = SpreadsheetApp.openById(config.spreadsheetId);
    const sheet = ss.getSheetByName(config.sheetName);
    
    if (!sheet) {
      throw new Error(`シート「${config.sheetName}」が見つかりません。`);
    }
    
    // 日付セルを読み取り
    const dateValue = sheet.getRange(config.dateCell).getValue();
    
    // データ範囲を読み取り
    const dataRange = sheet.getRange(config.dataRange);
    const rowCount = dataRange.getNumRows();
    
    const result = {
      status: 'OK',
      spreadsheetName: ss.getName(),
      sheetName: sheet.getName(),
      dateCell: config.dateCell,
      dateCellValue: dateValue,
      dataRange: config.dataRange,
      dataRowCount: rowCount
    };
    
    Logger.log('接続テスト成功:');
    Logger.log(JSON.stringify(result, null, 2));
    
    return result;
    
  } catch (e) {
    Logger.log('接続テスト失敗: ' + e.message);
    throw e;
  }
}

/**
 * デプロイ後のURL確認用
 */
function getWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  Logger.log('WebアプリURL: ' + url);
  return url;
}

/**
 * 所在地列（C列）を挿入するマイグレーション
 *
 * 変更前: A:人員 B:行先 C:出社日 D:帰社時刻 E:備考
 * 変更後: A:人員 B:行先 C:所在地 D:出社日 E:帰社時刻 F:備考
 *
 * 実行すると：
 * 1. C列の前に新しい列を挿入（既存C〜E列がD〜F列にシフト）
 * 2. ヘッダー行（3行目）のC列に「所在地」を設定
 * 3. ScriptProperties の HEADER_RANGE を A3:F3 に更新
 * 4. ScriptProperties の DATA_RANGE を A4:F19 に更新
 *
 * ※ C3が既に「所在地」の場合はスキップします
 */
function migrateAddLocationColumn() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  const sheetName = props.getProperty('SHEET_NAME') || '外出ホワイトボード';

  if (!spreadsheetId || spreadsheetId === 'YOUR_SPREADSHEET_ID_HERE') {
    throw new Error('SPREADSHEET_ID が設定されていません。先に initializeProperties を実行してください。');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('シート「' + sheetName + '」が見つかりません。');
  }

  // C3が既に「所在地」なら列挿入をスキップ
  const headerC = String(sheet.getRange('C3').getValue()).trim();
  if (headerC === '所在地') {
    Logger.log('C3は既に「所在地」です。列挿入をスキップします。');
  } else {
    // C列の前に1列挿入（既存C〜E列 → D〜F列にシフト）
    sheet.insertColumnBefore(3);
    Logger.log('C列を挿入しました（既存C〜E列がD〜F列にシフト）。');

    // ヘッダーを設定
    sheet.getRange('C3').setValue('所在地');
    Logger.log('C3に「所在地」ヘッダーを設定しました。');
  }

  // ScriptProperties を更新
  props.setProperty('HEADER_RANGE', 'A3:F3');
  props.setProperty('DATA_RANGE', 'A4:F19');

  Logger.log('ScriptProperties を更新しました:');
  Logger.log('  HEADER_RANGE: A3:F3');
  Logger.log('  DATA_RANGE: A4:F19');
  Logger.log('マイグレーション完了。新しいバージョンをデプロイしてください。');

  return 'マイグレーション完了';
}

/**
 * 所在地列と行先列を入れ替えるマイグレーション
 *
 * 変更前: A:人員 B:行先 C:所在地 D:出社日 E:帰社時刻 F:備考
 * 変更後: A:人員 B:所在地 C:行先 D:出社日 E:帰社時刻 F:備考
 *
 * 実行すると：
 * 1. ヘッダー行（3行目）のB列とC列を入れ替え
 * 2. データ行（4〜19行目）のB列とC列を入れ替え
 *
 * ※ B3が既に「所在地」の場合はスキップします
 */
function migrateSwapLocationDestination() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  const sheetName = props.getProperty('SHEET_NAME') || '外出ホワイトボード';

  if (!spreadsheetId || spreadsheetId === 'YOUR_SPREADSHEET_ID_HERE') {
    throw new Error('SPREADSHEET_ID が設定されていません。先に initializeProperties を実行してください。');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('シート「' + sheetName + '」が見つかりません。');
  }

  // B3が既に「所在地」なら入れ替え済み
  const headerB = String(sheet.getRange('B3').getValue()).trim();
  if (headerB === '所在地') {
    Logger.log('B3は既に「所在地」です。入れ替え済みのためスキップします。');
    return '既に入れ替え済みです';
  }

  // ヘッダー入れ替え
  sheet.getRange('B3').setValue('所在地');
  sheet.getRange('C3').setValue('行先');
  Logger.log('ヘッダーを入れ替えました: B3=所在地, C3=行先');

  // データ行（4〜19行目）のB列とC列を入れ替え
  var dataRange = sheet.getRange('B4:C19');
  var values = dataRange.getValues();

  var swapped = values.map(function(row) {
    return [row[1], row[0]];
  });

  dataRange.setValues(swapped);
  Logger.log('データ行（4〜19行目）のB列とC列を入れ替えました。');

  SpreadsheetApp.flush();
  Logger.log('マイグレーション完了: 所在地↔行先 列入れ替え');

  return 'マイグレーション完了';
}

/**
 * 部署列（B列）を挿入するマイグレーション
 *
 * 変更前: A:人員 B:所在地 C:内線 D:行先 E:出社日 F:帰社時刻 G:備考
 * 変更後: A:人員 B:部署 C:所在地 D:内線 E:行先 F:出社日 G:帰社時刻 H:備考
 *
 * 実行すると：
 * 1. B列の前に新しい列を挿入（既存B〜G列がC〜H列にシフト）
 * 2. ヘッダー行（3行目）のB列に「部署」を設定
 * 3. ScriptProperties の HEADER_RANGE を A3:H3 に更新
 * 4. ScriptProperties の DATA_RANGE を A4:H19 に更新
 *
 * ※ B3が既に「部署」の場合はスキップします
 */
function migrateAddDepartmentColumn() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheetId = props.getProperty('SPREADSHEET_ID');
  var sheetName = props.getProperty('SHEET_NAME') || '外出ホワイトボード';

  if (!spreadsheetId || spreadsheetId === 'YOUR_SPREADSHEET_ID_HERE') {
    throw new Error('SPREADSHEET_ID が設定されていません。先に initializeProperties を実行してください。');
  }

  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('シート「' + sheetName + '」が見つかりません。');
  }

  // B3が既に「部署」なら列挿入をスキップ
  var headerB = String(sheet.getRange('B3').getValue()).trim();
  if (headerB === '部署') {
    Logger.log('B3は既に「部署」です。列挿入をスキップします。');
  } else {
    // B列の前に1列挿入（既存B〜G列 → C〜H列にシフト）
    sheet.insertColumnBefore(2);
    Logger.log('B列を挿入しました（既存B〜G列がC〜H列にシフト）。');

    // ヘッダーを設定
    sheet.getRange('B3').setValue('部署');
    Logger.log('B3に「部署」ヘッダーを設定しました。');
  }

  // ScriptProperties を更新
  props.setProperty('HEADER_RANGE', 'A3:H3');
  props.setProperty('DATA_RANGE', 'A4:H19');

  Logger.log('ScriptProperties を更新しました:');
  Logger.log('  HEADER_RANGE: A3:H3');
  Logger.log('  DATA_RANGE: A4:H19');
  Logger.log('マイグレーション完了。新しいバージョンをデプロイしてください。');

  return 'マイグレーション完了';
}

/**
 * 内線番号列（C列）を挿入するマイグレーション
 *
 * 変更前: A:人員 B:所在地 C:行先 D:出社日 E:帰社時刻 F:備考
 * 変更後: A:人員 B:所在地 C:内線 D:行先 E:出社日 F:帰社時刻 G:備考
 *
 * 実行すると：
 * 1. C列の前に新しい列を挿入（既存C〜F列がD〜G列にシフト）
 * 2. ヘッダー行（3行目）のC列に「内線」を設定
 * 3. ScriptProperties の HEADER_RANGE を A3:G3 に更新
 * 4. ScriptProperties の DATA_RANGE を A4:G19 に更新
 *
 * ※ C3が既に「内線」の場合はスキップします
 */
function migrateAddExtensionColumn() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheetId = props.getProperty('SPREADSHEET_ID');
  var sheetName = props.getProperty('SHEET_NAME') || '外出ホワイトボード';

  if (!spreadsheetId || spreadsheetId === 'YOUR_SPREADSHEET_ID_HERE') {
    throw new Error('SPREADSHEET_ID が設定されていません。先に initializeProperties を実行してください。');
  }

  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('シート「' + sheetName + '」が見つかりません。');
  }

  // C3が既に「内線」なら列挿入をスキップ
  var headerC = String(sheet.getRange('C3').getValue()).trim();
  if (headerC === '内線') {
    Logger.log('C3は既に「内線」です。列挿入をスキップします。');
  } else {
    // C列の前に1列挿入（既存C〜F列 → D〜G列にシフト）
    sheet.insertColumnBefore(3);
    Logger.log('C列を挿入しました（既存C〜F列がD〜G列にシフト）。');

    // ヘッダーを設定
    sheet.getRange('C3').setValue('内線');
    Logger.log('C3に「内線」ヘッダーを設定しました。');
  }

  // ScriptProperties を更新
  props.setProperty('HEADER_RANGE', 'A3:G3');
  props.setProperty('DATA_RANGE', 'A4:G19');

  Logger.log('ScriptProperties を更新しました:');
  Logger.log('  HEADER_RANGE: A3:G3');
  Logger.log('  DATA_RANGE: A4:G19');
  Logger.log('マイグレーション完了。新しいバージョンをデプロイしてください。');

  return 'マイグレーション完了';
}
