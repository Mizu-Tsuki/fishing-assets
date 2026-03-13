/**
 * Data_System_Cache: 系統快取管理員
 * 負責將 API 抓回來的原始 JSON 存入快取分頁，或從中讀取。
 */

const Data_System_Cache = {

  /**
   * 存入快取
   * @param {string} key - 快取的名稱 (例如 "BANK_DATA")
   * @param {Object|Array} data - 要存入的 JSON 物件或陣列
   */
  save: function(key, data) {
    const sheet = Utils.getSheetByTag(SYSTEM_CONFIG.TAGS.CACHE);
    if (!sheet) return;

    const timestamp = new Date();
    const jsonString = JSON.stringify(data);

    // 我們假設 A 欄放 Key，B 欄放 JSON 字串，C 欄放更新時間
    // 這裡使用一個簡單的查找邏輯：如果 Key 已存在就覆蓋，不存在就新增
    const lastRow = sheet.getLastRow();
    let targetRow = 0;

    if (lastRow >= SYSTEM_CONFIG.LAYOUT.DATA_START_ROW) {
      const keys = sheet.getRange(
        SYSTEM_CONFIG.LAYOUT.DATA_START_ROW, 
        1, 
        lastRow - SYSTEM_CONFIG.LAYOUT.DATA_START_ROW + 1, 
        1
      ).getValues();
      
      for (let i = 0; i < keys.length; i++) {
        if (keys[i][0] === key) {
          targetRow = i + SYSTEM_CONFIG.LAYOUT.DATA_START_ROW;
          break;
        }
      }
    }

    if (targetRow === 0) targetRow = Math.max(lastRow + 1, SYSTEM_CONFIG.LAYOUT.DATA_START_ROW);

    sheet.getRange(targetRow, 1, 1, 3).setValues([[key, jsonString, timestamp]]);
  },

  /**
   * 讀取快取
   * @param {string} key - 快取的名稱
   * @return {Object|Array|null}
   */
  get: function(key) {
    const sheet = Utils.getSheetByTag(SYSTEM_CONFIG.TAGS.CACHE);
    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    if (lastRow < SYSTEM_CONFIG.LAYOUT.DATA_START_ROW) return null;

    const data = sheet.getRange(
      SYSTEM_CONFIG.LAYOUT.DATA_START_ROW, 
      1, 
      lastRow - SYSTEM_CONFIG.LAYOUT.DATA_START_ROW + 1, 
      2
    ).getValues();

    const row = data.find(r => r[0] === key);
    return row ? JSON.parse(row[1]) : null;
  }
};
