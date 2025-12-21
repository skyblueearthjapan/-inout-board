/**
 * バリデーション - 入力検証と正規化
 */
const Validation = (function() {
  
  /**
   * 日付形式をチェック
   * 許容: 空 / YYYY-MM-DD
   * @param {string} dateStr
   * @returns {boolean}
   */
  function isValidDateFormat(dateStr) {
    if (!dateStr || dateStr.trim() === '') {
      return true; // 空はOK
    }

    const trimmed = dateStr.trim();

    // YYYY-MM-DD 形式をチェック
    const pattern = /^\d{4}-\d{2}-\d{2}$/;
    if (!pattern.test(trimmed)) {
      return false;
    }

    // 実際に有効な日付かチェック
    const date = new Date(trimmed);
    return !isNaN(date.getTime());
  }

  /**
   * 時刻形式をチェック
   * 許容: 空 / H:MM / HH:MM（0-23時、0-59分）
   * @param {string} timeStr
   * @returns {boolean}
   */
  function isValidTimeFormat(timeStr) {
    if (!timeStr || timeStr.trim() === '') {
      return true; // 空はOK
    }

    const trimmed = timeStr.trim();

    // H:MM または HH:MM 形式をチェック
    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9])$/;
    return pattern.test(trimmed);
  }
  
  /**
   * 時刻を正規化（HH:MM形式に変換）
   * @param {string} timeStr
   * @returns {string}
   */
  function normalizeTime(timeStr) {
    if (!timeStr || timeStr.trim() === '') {
      return '';
    }
    
    const trimmed = timeStr.trim();
    const match = trimmed.match(/^(\d{1,2}):(\d{2})$/);
    
    if (!match) {
      return trimmed; // フォーマットが合わない場合はそのまま返す
    }
    
    const hours = match[1].padStart(2, '0');
    const minutes = match[2];
    
    return `${hours}:${minutes}`;
  }
  
  /**
   * 行データのバリデーション
   * @param {Object} payload
   * @returns {Object} {valid: boolean, message: string}
   */
  function validateRowPayload(payload) {
    const errors = [];

    // rowIndexの存在チェック
    if (payload.rowIndex === undefined || payload.rowIndex === null) {
      errors.push('行番号が指定されていません');
    }

    // 出社日のフォーマットチェック（YYYY-MM-DD）
    if (!isValidDateFormat(payload.startTime)) {
      errors.push('出社日の形式が不正です（例: 2025-12-25）');
    }

    // 帰社時刻のフォーマットチェック（HH:MM）
    if (!isValidTimeFormat(payload.endTime)) {
      errors.push('帰社時刻の形式が不正です（例: 17:30 または 09:00）');
    }

    if (errors.length > 0) {
      return {
        valid: false,
        message: errors.join('\n')
      };
    }

    return {
      valid: true,
      message: ''
    };
  }
  
  // 公開API
  return {
    isValidDateFormat: isValidDateFormat,
    isValidTimeFormat: isValidTimeFormat,
    normalizeTime: normalizeTime,
    validateRowPayload: validateRowPayload
  };
  
})();
