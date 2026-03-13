/**
 * Data_Sync_Market: 市場價格同步
 */
const Data_Sync_Market = {
  runSync() {
    const { MARKET, WATCHLIST } = SYSTEM_CONFIG.SHEET_NAMES;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 從 Watchlist 或 Assets 搜集所有需要查價的 ID
    const ids = this._collectNeededIds();
    if (ids.length === 0) return;

    // 2. 呼叫 Conn_GW2 去拿價格 (v2/commerce/prices)
    // 這裡通常會分批抓取，因為 API 一次限制 200 個 ID
    const priceData = Conn_GW2.fetch(`/v2/commerce/prices?ids=${ids.join(",")}`);

    // 3. 整理成表格格式並寫入 Market 分頁
    const tableData = this._transformPricesToTable(priceData);
    const marketSheet = ss.getSheetByName(MARKET);
    if (marketSheet && tableData.length > 0) {
      marketSheet.getRange(2, 1, marketSheet.getLastRow() || 1, marketSheet.getLastColumn() || 1).clearContent();
      marketSheet.getRange(2, 1, tableData.length, tableData[0].length).setValues(tableData);
    }
  },

  _collectNeededIds() {
    // 這裡寫：去各分頁收集 ID 的邏輯
    return []; 
  },

  _transformPricesToTable(data) {
    // 這裡寫：將價格 JSON 轉成二維陣列的邏輯
    return [];
  }
};
