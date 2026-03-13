/**
 * Data_Fetch_Asset: 資產收割機 (完全體)
 */
const Data_Fetch_Asset = {

  fetchAll: function() {
    console.log("🚀 開始執行全域資產併發收割...");
    const allResults = [];
    const timestamp = new Date();

    // --- 第一階段：銀行與素材庫 (基本盤) ---
    const baseEndpoints = {
      "ASSET_BANK": "/account/bank",
      "ASSET_MATERIALS": "/account/materials"
    };
    
    for (let key in baseEndpoints) {
      const data = Conn_GW2.fetch(baseEndpoints[key]);
      if (data) allResults.push([key, JSON.stringify(data), timestamp]);
    }

    // --- 第二階段：角色背包 (併發分組盤) ---
    console.log("👤 正在獲取角色清單...");
    const charNames = Conn_GW2.fetch("/characters");
    
    if (charNames && Array.isArray(charNames)) {
      // 每 20 個角色分一組
      const nameChunks = Utils.chunkArray(charNames, 20);
      
      nameChunks.forEach((chunk, index) => {
        console.log(`  > 正在抓取第 ${index + 1} 組角色背包 (併發數: ${chunk.length})...`);
        
        // 準備這一組的併發請求
        const requests = chunk.map(name => {
          return {
            url: SYSTEM_CONFIG.BASE_URL + "/characters/" + encodeURIComponent(name) + "/inventory",
            method: "get",
            headers: { "Authorization": "Bearer " + SYSTEM_CONFIG.API_KEY },
            muteHttpExceptions: true
          };
        });

        // 執行平行請求
        const responses = UrlFetchApp.fetchAll(requests);
        
        responses.forEach((res, i) => {
          if (res.getResponseCode() === 200) {
            const data = JSON.parse(res.getContentText());
            allResults.push(["CHAR_" + chunk[i], JSON.stringify(data), timestamp]);
          } else {
            console.error(`  ❌ 角色 ${chunk[i]} 抓取失敗: ${res.getResponseCode()}`);
          }
        });
      });
    }

    // --- 第三階段：一次性入庫 ---
    Data_System_Cache.overwriteAll(allResults);
    console.log("🏁 全域資產收割流程結束。");
  }
};

/**
 * 手動執行按鈕
 */
function test_Asset_Fetch() {
  Data_Fetch_Asset.fetchAll();
}
