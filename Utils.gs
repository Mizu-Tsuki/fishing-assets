/**
 * Utils: 通用工具函式庫
 * 負責處理試算表底層操作與資料格式轉換。
 */

const Utils = {

  /**
   * 透過分頁 A1 的 Tag 標記來尋找分頁物件
   * @param {string} tag - 在 SYSTEM_CONFIG.TAGS 中定義的標籤字串
   * @return {GoogleAppsScript.Spreadsheet.Sheet|null} - 回傳分頁物件或 null
   */
  getSheetByTag: function(tag) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    // 遍歷所有分頁，檢查 A1 儲存格內容是否與 Tag 相符
    const targetSheet = sheets.find(sheet => {
      // 使用 getRange(1, 1) 效能較佳
      return sheet.getRange(1, 1).getValue().toString().trim() === tag;
    });

    if (!targetSheet) {
      console.warn("警告：找不到標記為 [" + tag + "] 的分頁。");
      return null;
    }
    
    return targetSheet;
  },

  /**
   * 取得資料起始位置
   * 方便未來快速獲取第 11 列之後的範圍
   */
  getDataRange: function(sheet) {
    if (!sheet) return null;
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    // 如果沒有資料，回傳 null
    if (lastRow < SYSTEM_CONFIG.LAYOUT.DATA_START_ROW) return null;
    
    return sheet.getRange(
      SYSTEM_CONFIG.LAYOUT.DATA_START_ROW, 
      1, 
      lastRow - SYSTEM_CONFIG.LAYOUT.DATA_START_ROW + 1, 
      lastColumn
    );
  }
};
