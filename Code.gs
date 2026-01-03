/**
 * 外出記入板 Webアプリ - メインエントリポイント
 * 
 * 初回デプロイ前に、スクリプトプロパティに以下を設定してください：
 * - SPREADSHEET_ID: 対象スプレッドシートのID
 * - SHEET_NAME: シート名（例: 外出ホワイトボード）
 * - DATE_CELL: 日付セル（例: D2）
 * - HEADER_RANGE: ヘッダー範囲（例: A3:E3）
 * - DATA_RANGE: データ範囲（例: A4:E19）
 */

/**
 * Webアプリのエントリポイント
 * @param {Object} e - イベントオブジェクト
 * @returns {HtmlOutput}
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  
  // URLパラメータから初期日付を取得（?date=YYYY-MM-DD）
  const initialDate = e && e.parameter && e.parameter.date 
    ? e.parameter.date 
    : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  
  template.initialDate = initialDate;
  
  return template.evaluate()
    .setTitle('外出記入板')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * HTMLファイルをインクルードするためのヘルパー
 * @param {string} filename - ファイル名
 * @returns {string}
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 指定日のデータを取得
 * @param {string} dateString - 日付文字列（YYYY-MM-DD形式、省略時は今日）
 * @returns {Object} DayData
 */
function getDayData(dateString) {
  const lock = LockService.getDocumentLock();
  
  try {
    // ロック取得（最大5秒待機）
    if (!lock.tryLock(5000)) {
      throw new Error('他のユーザーがデータを更新中です。数秒後に再試行してください。');
    }
    
    const result = SheetService.getDayData(dateString);
    return result;
    
  } finally {
    lock.releaseLock();
  }
}

/**
 * 1行分のデータを更新
 * @param {string} dateString - 日付文字列（YYYY-MM-DD形式）
 * @param {Object} payload - 更新データ
 * @returns {Object} 更新後のOutingRow
 */
function updateRow(dateString, payload) {
  const lock = LockService.getDocumentLock();

  try {
    // ロック取得（最大10秒待機）
    if (!lock.tryLock(10000)) {
      throw new Error('他のユーザーがデータを更新中です。数秒後に再試行してください。');
    }

    // バリデーション
    const validationResult = Validation.validateRowPayload(payload);
    if (!validationResult.valid) {
      throw new Error(validationResult.message);
    }

    // 更新実行
    const result = SheetService.updateRow(dateString, payload);
    return result;

  } finally {
    lock.releaseLock();
  }
}

/**
 * 複数行のデータを一括更新（A+C+B方式のB層API）
 * @param {string} dateString - 日付文字列（YYYY-MM-DD形式）
 * @param {Array<Object>} changes - 変更配列
 * @returns {Array<Object>} 更新結果
 */
function api_applyPatch(dateString, changes) {
  const lock = LockService.getDocumentLock();

  try {
    // ロック取得（最大15秒待機、複数行更新のため長め）
    if (!lock.tryLock(15000)) {
      throw new Error('他のユーザーがデータを更新中です。数秒後に再試行してください。');
    }

    Logger.log('api_applyPatch called with ' + changes.length + ' changes');

    // 一括更新実行
    const results = SheetService.applyPatch(dateString, changes);
    return results;

  } catch (error) {
    Logger.log('api_applyPatch error: ' + error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 設定値を取得（デバッグ用）
 * @returns {Object}
 */
function getConfig() {
  return SheetService.getConfig();
}
