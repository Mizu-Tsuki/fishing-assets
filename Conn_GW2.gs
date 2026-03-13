/**
 * Conn_GW2: GW2 API 連線核心
 * 負責所有與官方 API 的通訊，具備自動授權與錯誤處理功能。
 */

const Conn_GW2 = {

  /**
   * 通用抓取函式
   * @param {string} endpoint - API 路徑 (例如 "/account/bank")
   * @return {Object|Array|null} - 回傳解析後的 JSON 資料
   */
  fetch: function(endpoint) {
    const apiKey = SYSTEM_CONFIG.API_KEY;
    
    if (!apiKey) {
      console.error("錯誤：未偵測到 API Key，請檢查 System_Config 分頁。");
      return null;
    }

    // 自動拼接完整的 URL
    const url = SYSTEM_CONFIG.BASE_URL + endpoint;
    
    // 設定連線參數，自動掛載 Bearer Token
    const options = {
      "method": "get",
      "headers": {
        "Authorization": "Bearer " + apiKey
      },
      "muteHttpExceptions": true // 讓程式在 API 報錯時不中斷，以便我們抓取錯誤訊息
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const content = response.getContentText();

      if (code === 200) {
        return JSON.parse(content);
      } else {
        console.warn("API 回報錯誤 (代碼: " + code + "): " + content);
        return null;
      }
    } catch (e) {
      console.error("連線發生異常: " + e.toString());
      return null;
    }
  }
};
