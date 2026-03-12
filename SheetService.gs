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
   * セル値を文字列に変換（時刻用）
   * @param {*} value
   * @returns {string}
   */
  function cellToTimeString(value) {
    if (value === null || value === undefined) {
      return '';
    }
    if (value instanceof Date) {
      // 時刻として扱う
      return Utilities.formatDate(value, 'Asia/Tokyo', 'HH:mm');
    }
    return String(value).trim();
  }

  /**
   * セル値を文字列に変換（日付用）
   * @param {*} value
   * @returns {string}
   */
  function cellToDateString(value) {
    if (value === null || value === undefined) {
      return '';
    }
    if (value instanceof Date) {
      // 日付として扱う
      return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
    }
    return String(value).trim();
  }

  /**
   * セル値を文字列に変換（汎用）
   * @param {*} value
   * @returns {string}
   */
  function cellToString(value) {
    if (value === null || value === undefined) {
      return '';
    }
    if (value instanceof Date) {
      // デフォルトは時刻として扱う（後方互換性）
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
    // 列構成: A:人員 B:行先 C:所在地 D:出社日 E:帰社時刻 F:備考
    const rows = values.map((row, index) => {
      const name = cellToString(row[0]);
      const destination = cellToString(row[1]);
      const location = row.length > 2 ? cellToString(row[2]) : '';  // C列：所在地
      const startTime = row.length > 3 ? cellToDateString(row[3]) : '';  // D列：出社日
      const endTime = row.length > 4 ? cellToTimeString(row[4]) : '';    // E列：帰社時刻
      const note = row.length > 5 ? cellToString(row[5]) : '';           // F列：備考

      return {
        rowIndex: startRow + index,
        name: name,
        destination: destination,
        location: location,
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
   * 複数行のデータを一括更新（A+C+B方式のB層）
   * @param {string} dateString - YYYY-MM-DD形式
   * @param {Array<Object>} changes - 変更配列 [{rowIndex, destination, startTime, endTime, note}, ...]
   * @returns {Array<Object>} 更新結果 [{success: boolean, rowIndex: number, row: Object, error: string}, ...]
   */
  function applyPatch(dateString, changes) {
    Logger.log('applyPatch called with ' + changes.length + ' changes');

    const results = [];
    const config = getConfig();
    const sheet = getSheet();

    // シートに日付をセット
    setSheetDate(sheet, dateString);
    SpreadsheetApp.flush();

    // 各変更を適用
    changes.forEach(function(change) {
      try {
        // バリデーション
        const validation = Validation.validateRowPayload(change);
        if (!validation.valid) {
          results.push({
            success: false,
            rowIndex: change.rowIndex,
            row: null,
            error: validation.message
          });
          return;
        }

        const rowIndex = change.rowIndex;

        // B〜F列（2〜6列目）を更新
        // B:行先, C:所在地, D:出社日, E:帰社時刻, F:備考
        if (change.destination !== undefined) {
          sheet.getRange(rowIndex, 2).setValue(change.destination || '');
        }

        if (change.location !== undefined) {
          sheet.getRange(rowIndex, 3).setValue(change.location || '');
        }

        // 出社日セルを文字列として保存
        if (change.startTime !== undefined) {
          const startDateCell = sheet.getRange(rowIndex, 4);
          startDateCell.setNumberFormat('@');
          startDateCell.setValue(change.startTime || '');
        }

        // 帰社時刻セルを文字列として保存
        if (change.endTime !== undefined) {
          const normalizedEndTime = Validation.normalizeTime(change.endTime || '');
          const endTimeCell = sheet.getRange(rowIndex, 5);
          endTimeCell.setNumberFormat('@');
          endTimeCell.setValue(normalizedEndTime);
        }

        if (change.note !== undefined) {
          sheet.getRange(rowIndex, 6).setValue(change.note || '');
        }

        SpreadsheetApp.flush();

        // 更新後の行を読み直す
        const updatedRow = sheet.getRange(rowIndex, 1, 1, 6).getValues()[0];

        results.push({
          success: true,
          rowIndex: rowIndex,
          row: {
            rowIndex: rowIndex,
            name: cellToString(updatedRow[0]),
            destination: cellToString(updatedRow[1]),
            location: cellToString(updatedRow[2]),
            startTime: cellToDateString(updatedRow[3]),
            endTime: cellToTimeString(updatedRow[4]),
            note: cellToString(updatedRow[5]),
            status: calculateStatus(
              cellToDateString(updatedRow[3]),
              cellToTimeString(updatedRow[4])
            )
          },
          error: null
        });

      } catch (error) {
        Logger.log('Error updating row ' + change.rowIndex + ': ' + error.message);
        results.push({
          success: false,
          rowIndex: change.rowIndex,
          row: null,
          error: error.message
        });
      }
    });

    Logger.log('applyPatch completed: ' + results.filter(r => r.success).length + ' success, ' +
               results.filter(r => !r.success).length + ' failed');

    return results;
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

    // B〜F列（2〜6列目）を更新
    // B:行先, C:所在地, D:出社日, E:帰社時刻, F:備考
    const normalizedEndTime = Validation.normalizeTime(payload.endTime || '');

    sheet.getRange(rowIndex, 2).setValue(payload.destination || '');
    sheet.getRange(rowIndex, 3).setValue(payload.location || '');

    // 出社日セルを文字列として保存（日付変換を防ぐ）
    const startDateCell = sheet.getRange(rowIndex, 4);
    startDateCell.setNumberFormat('@');
    startDateCell.setValue(payload.startTime || '');

    // 帰社時刻セルを文字列として保存（タイムゾーン変換を防ぐ）
    const endTimeCell = sheet.getRange(rowIndex, 5);
    endTimeCell.setNumberFormat('@');
    endTimeCell.setValue(normalizedEndTime);

    sheet.getRange(rowIndex, 6).setValue(payload.note || '');

    SpreadsheetApp.flush();

    // 更新後の行を読み直す
    const updatedRow = sheet.getRange(rowIndex, 1, 1, 6).getValues()[0];

    return {
      rowIndex: rowIndex,
      name: cellToString(updatedRow[0]),
      destination: cellToString(updatedRow[1]),
      location: cellToString(updatedRow[2]),
      startTime: cellToDateString(updatedRow[3]),
      endTime: cellToTimeString(updatedRow[4]),
      note: cellToString(updatedRow[5]),
      status: calculateStatus(
        cellToDateString(updatedRow[3]),
        cellToTimeString(updatedRow[4])
      )
    };
  }
  
  // 公開API
  return {
    getConfig: getConfig,
    getDayData: getDayData,
    updateRow: updateRow,
    applyPatch: applyPatch
  };

})();
