/**
 * Data_System_Cache: 系統快取管理員
 */
const Data_System_Cache = {

  /**
   * 批次覆蓋快取 (效能核心 + 保險機制)
   * @param {Array<Array>} rows - 格式為 [[key, jsonString, timestamp], ...]
   */
  overwriteAll: function(rows) {
    // 保險機制：如果傳入的資料是空的，為了防止誤刪舊快取，直接中斷
    if (!rows || rows.length === 0) {
      console.warn("⚠️ 接收到空的資料集，取消覆蓋動作以保護舊有快取。");
      return;
    }

    const sheet = Utils.getSheetByTag(SYSTEM_CONFIG.TAGS.CACHE);
    if (!sheet) return;

    const startRow = SYSTEM_CONFIG.LAYOUT.DATA_START_ROW;
    const lastRow = sheet.getLastRow();

    // 1. 清空舊資料 (從第 11 列開始到最後一列，包含 A, B, C 三欄)
    if (lastRow >= startRow) {
      sheet.getRange(startRow, 1, lastRow - startRow + 1, 3).clearContent();
    }

    // 2. 一次性寫入新資料
    sheet.getRange(startRow, 1, rows.length, 3).setValues(rows);
    
    console.log(`✅ 成功完成批次入庫，總計 ${rows.length} 項貨架標籤。`);
  }
};
