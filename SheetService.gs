/**
 * シートサービス - スプレッドシートの読み書き処理
 */
const SheetService = (function() {
  
  /**
   * 設定値を取得
   * @returns {Object}
   */
  function getConfig() {
    const props = PropertiesService.getScriptProperties();
    return {
      spreadsheetId: props.getProperty('SPREADSHEET_ID') || '',
      sheetName: props.getProperty('SHEET_NAME') || '外出ホワイトボード',
      dateCell: props.getProperty('DATE_CELL') || 'D2',
      headerRange: props.getProperty('HEADER_RANGE') || 'A3:E3',
      dataRange: props.getProperty('DATA_RANGE') || 'A4:E19'
    };
  }
  
  /**
   * シートを取得
   * @returns {Sheet}
   */
  function getSheet() {
    const config = getConfig();
    
    if (!config.spreadsheetId) {
      throw new Error('SPREADSHEET_ID がスクリプトプロパティに設定されていません。');
    }
    
    const ss = SpreadsheetApp.openById(config.spreadsheetId);
    const sheet = ss.getSheetByName(config.sheetName);
    
    if (!sheet) {
      throw new Error(`シート「${config.sheetName}」が見つかりません。`);
    }
    
    return sheet;
  }
  
  /**
   * 日付をシートにセット
   * @param {Sheet} sheet
   * @param {string} dateString - YYYY-MM-DD形式
   */
  function setSheetDate(sheet, dateString) {
    const config = getConfig();
    const dateCell = sheet.getRange(config.dateCell);
    
    // 日付をDate型に変換してセット
    const dateParts = dateString.split('-');
    const dateObj = new Date(
      parseInt(dateParts[0]),
      parseInt(dateParts[1]) - 1,
      parseInt(dateParts[2])
    );
    
    dateCell.setValue(dateObj);
  }
  
  /**
   * シートから日付を取得（表示用フォーマット）
   * @param {Sheet} sheet
   * @returns {string}
   */
  function getSheetDateDisplay(sheet) {
    const config = getConfig();
    const dateCell = sheet.getRange(config.dateCell);
    const dateValue = dateCell.getValue();
    
    if (dateValue instanceof Date) {
      return Utilities.formatDate(dateValue, 'Asia/Tokyo', 'M/d/yyyy');
    }
    return String(dateValue);
  }
  
  /**
   * 行のステータスを計算
   * @param {string} startTime
   * @param {string} endTime
   * @returns {string} 'NONE' | 'OUT' | 'BACK'
   */
  function calculateStatus(startTime, endTime) {
    const hasStart = startTime && String(startTime).trim() !== '';
    const hasEnd = endTime && String(endTime).trim() !== '';
    
    if (hasEnd) {
      return 'BACK';
    } else if (hasStart) {
      return 'OUT';
    }
    return 'NONE';
  }
  
  /**
   * セル値を文字列に変換
   * @param {*} value
   * @returns {string}
   */
  function cellToString(value) {
    if (value === null || value === undefined) {
      return '';
    }
    if (value instanceof Date) {
      // 時刻として扱う場合
      return Utilities.formatDate(value, 'Asia/Tokyo', 'HH:mm');
    }
    return String(value).trim();
  }
  
  /**
   * 指定日のデータを取得
   * @param {string} dateString - YYYY-MM-DD形式（省略時は今日）
   * @returns {Object} DayData
   */
  function getDayData(dateString) {
    const config = getConfig();
    const sheet = getSheet();
    
    // 日付が指定されていない場合は今日
    if (!dateString) {
      dateString = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    }
    
    // シートに日付をセット
    setSheetDate(sheet, dateString);
    
    // SpreadsheetApp.flush() で変更を確定
    SpreadsheetApp.flush();
    
    // データ範囲を取得
    const dataRange = sheet.getRange(config.dataRange);
    const values = dataRange.getValues();
    
    // データ範囲の開始行番号を取得
    const startRow = dataRange.getRow();
    
    // OutingRow配列に変換
    const rows = values.map((row, index) => {
      const name = cellToString(row[0]);
      const destination = cellToString(row[1]);
      const startTime = cellToString(row[2]);
      const endTime = cellToString(row[3]);
      const note = cellToString(row[4]);
      
      return {
        rowIndex: startRow + index,
        name: name,
        destination: destination,
        startTime: startTime,
        endTime: endTime,
        note: note,
        status: calculateStatus(startTime, endTime)
      };
    });
    
    return {
      date: dateString,
      sheetDateDisplay: getSheetDateDisplay(sheet),
      rows: rows
    };
  }
  
  /**
   * 1行分のデータを更新
   * @param {string} dateString - YYYY-MM-DD形式
   * @param {Object} payload - 更新データ
   * @returns {Object} 更新後のOutingRow
   */
  function updateRow(dateString, payload) {
    const config = getConfig();
    const sheet = getSheet();
    
    // シートに日付をセット
    setSheetDate(sheet, dateString);
    SpreadsheetApp.flush();
    
    const rowIndex = payload.rowIndex;
    
    // B〜E列（2〜5列目）を更新
    // 時刻を正規化
    const normalizedStartTime = Validation.normalizeTime(payload.startTime || '');
    const normalizedEndTime = Validation.normalizeTime(payload.endTime || '');
    
    sheet.getRange(rowIndex, 2).setValue(payload.destination || '');
    sheet.getRange(rowIndex, 3).setValue(normalizedStartTime);
    sheet.getRange(rowIndex, 4).setValue(normalizedEndTime);
    sheet.getRange(rowIndex, 5).setValue(payload.note || '');
    
    SpreadsheetApp.flush();
    
    // 更新後の行を読み直す
    const updatedRow = sheet.getRange(rowIndex, 1, 1, 5).getValues()[0];
    
    return {
      rowIndex: rowIndex,
      name: cellToString(updatedRow[0]),
      destination: cellToString(updatedRow[1]),
      startTime: cellToString(updatedRow[2]),
      endTime: cellToString(updatedRow[3]),
      note: cellToString(updatedRow[4]),
      status: calculateStatus(
        cellToString(updatedRow[2]),
        cellToString(updatedRow[3])
      )
    };
  }
  
  // 公開API
  return {
    getConfig: getConfig,
    getDayData: getDayData,
    updateRow: updateRow
  };
  
})();
