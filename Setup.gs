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
