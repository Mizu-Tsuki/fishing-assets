/**
 * Utils: 通用工具函式庫
 */
const Utils = {

  /**
   * 透過分頁 A1 的 Tag 標記來尋找分頁物件
   */
  getSheetByTag: function(tag) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const targetSheet = sheets.find(sheet => {
      return sheet.getRange(1, 1).getValue().toString().trim() === tag;
    });
    if (!targetSheet) console.warn("警告：找不到標記為 [" + tag + "] 的分頁。");
    return targetSheet;
  },

  /**
   * 將陣列切分成指定大小的區塊 (Chunking)
   * @param {Array} array - 原陣列
   * @param {number} size - 每個區塊的大小 (例如 20)
   * @return {Array<Array>}
   */
  chunkArray: function(array, size) {
    const results = [];
    for (let i = 0; i < array.length; i += size) {
      results.push(array.slice(i, i + size));
    }
    return results;
  }
};
