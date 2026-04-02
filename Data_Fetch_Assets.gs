/**
 * Data_Fetch_Assets: 負責從 GW2 API 抓取帳號資產並存入快取
 */
function run_Asset_Update() {
  Asset_Fetcher.run();
}
const Asset_Fetcher = {

  /**
   * 執行全帳號資產更新
   */
  run: function() {
    console.log("開始抓取帳號資產...");
    
    // 1. 取得 API Key (假設你存在 SYSTEM_CONFIG 或特定位置)
    const apiKey = SYSTEM_CONFIG.API_KEY; 
    if (!apiKey) return console.error("請先在 SYSTEM_CONFIG 中設定 API_KEY");

    const cacheSheet = Utils.getSheetByTag("SYS_CACHE");
    if (!cacheSheet) return console.error("找不到 SYS_CACHE 分頁");

    // 2. 定義要抓取的目標與對應的 API 路徑
    const targets = [
      { key: "ASSET_MATERIALS", url: "/account/materials" },
      { key: "ASSET_BANK", url: "/account/bank" },
      { key: "ASSET_WALLET", url: "/account/wallet" }
    ];

    const results = [];

    // --- A. 抓取共通資產 (素材、銀行、錢包) ---
    targets.forEach(target => {
      const data = this._fetch(target.url, apiKey);
      if (data) {
        results.push([target.key, JSON.stringify(data), new Date()]);
      }
    });

    // --- B. 抓取角色清單與各別包包 ---
    const characters = this._fetch("/characters?page=0", apiKey); // 抓取所有角色詳細資料
    if (characters && Array.isArray(characters)) {
      characters.forEach(char => {
        const charKey = `CHAR_${char.name}`;
        // 提取該角色的背包 (bags) 資料
        const bagData = char.bags || [];
        results.push([charKey, JSON.stringify(bagData), new Date()]);
      });
    }

    // 3. 寫入 SYS_CACHE (從 Row 11 開始，整張覆蓋)
    if (results.length > 0) {
      this._saveToCache(cacheSheet, results);
      console.log("SYS_CACHE 更新成功。");
    }

    // --- C. 【核心連動】自動觸發後續同步 ---
    console.log("偵測到資產更新，自動啟動素材庫同步...");
    try {
      Sync_Material.sync(); 
      console.log("所有自動化同步已完成！");
    } catch (e) {
      console.error("自動同步失敗: " + e.message);
    }
  },

  /**
   * 內部私有：API 請求工具
   */
  _fetch: function(endpoint, apiKey) {
    const url = `${SYSTEM_CONFIG.BASE_URL}${endpoint}`;
    const options = {
      headers: { "Authorization": `Bearer ${apiKey}` },
      muteHttpExceptions: true
    };
    try {
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        return JSON.parse(response.getContentText());
      } else {
        console.warn(`API 警告 (${endpoint}): ${response.getContentText()}`);
        return null;
      }
    } catch (e) {
      console.error(`Fetch 失敗 (${endpoint}): ${e.message}`);
      return null;
    }
  },

  /**
   * 內部私有：寫入快取表
   */
  _saveToCache: function(sheet, results) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 11) {
      sheet.getRange(11, 1, lastRow - 10, 3).clearContent();
    }
    sheet.getRange(11, 1, results.length, 3).setValues(results);
  }
};
