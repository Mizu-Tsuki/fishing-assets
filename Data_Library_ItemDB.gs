/**
 * Data_Library_ItemDB: 官方物品數據百科管理員
 * * 功能：
 * 1. 支援「雙語併發」抓取 (CN/EN)。
 * 2. 自動補全：檢查 A 欄 ID，若無資料則自動填充。
 * 3. 欄位保護：保留已有的中文名稱，不隨意覆蓋。
 */
const Data_Library_ItemDB = {

  /**
   * 驅動核心：更新 ItemDB 內容
   */
  updateItemDatabase: function() {
    const sheet = Utils.getSheetByTag(SYSTEM_CONFIG.TAGS.ITEMDB);
    if (!sheet) return;

    const startRow = 11; // 依照規範，資料從第 11 列開始
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) {
      console.warn("⚠️ ItemDB 目前沒有資料列。");
      return;
    }

    // 讀取目前的整張字典表 (A-H 欄)
    const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, 8);
    const values = range.getValues();
    
    // 找出需要補全的 ID 清單 (判斷標準：A欄有ID，但D欄英文名是空的)
    const missingIds = [];
    const idToIndicesMap = {};

    values.forEach((row, index) => {
      const id = String(row[0]).trim();
      const nameEn = row[3]; 
      if (id && id !== "" && (!nameEn || nameEn === "")) {
        missingIds.push(id);
        if (!idToIndicesMap[id]) idToIndicesMap[id] = [];
        idToIndicesMap[id].push(index);
      }
    });

    if (missingIds.length === 0) {
      console.log("✅ ItemDB 字典資料已齊全，不需要更新。");
      return;
    }

    // 將缺少的 ID 每 200 個分一組進行批次請求
    const chunks = Utils.chunkArray(missingIds, 200);
    const dbData = {}; 

    console.log(`🚀 正在補全 ${missingIds.length} 個物品資訊，分 ${chunks.length} 組執行...`);

    chunks.forEach((chunk, chunkIdx) => {
      const idsParam = chunk.join(",");
      
      // 同時發送雙語請求 (併發 2 條連線)
      const requests = [
        { 
          url: SYSTEM_CONFIG.BASE_URL + "/v2/items?ids=" + idsParam + "&lang=en", 
          method: "get", 
          muteHttpExceptions: true 
        },
        { 
          url: SYSTEM_CONFIG.BASE_URL + "/v2/items?ids=" + idsParam + "&lang=zh", 
          method: "get", 
          muteHttpExceptions: true 
        }
      ];

      const responses = UrlFetchApp.fetchAll(requests);
      
      // 處理英文資料 (基礎資料)
      if (responses[0].getResponseCode() === 200) {
        JSON.parse(responses[0].getContentText()).forEach(item => {
          dbData[item.id] = {
            en: item.name,
            icon: item.icon,
            type: item.type,
            rarity: item.rarity,
            level: item.level,
            link: item.chat_link
          };
        });
      }

      // 處理中文資料 (補強名稱)
      if (responses[1].getResponseCode() === 200) {
        try {
          JSON.parse(responses[1].getContentText()).forEach(item => {
            if (dbData[item.id]) {
              dbData[item.id].cn = item.name;
            }
          });
        } catch(e) {
          console.warn(`第 ${chunkIdx + 1} 組簡中資料解析異常，可能該批次無中文資料。`);
        }
      }
    });

    // 將抓到的資料填回原始 values 陣列中
    values.forEach((row, index) => {
      const id = String(row[0]).trim();
      const info = dbData[id];
      
      if (info) {
        // B 欄: Icon (若原本沒圖才補)
        if (!row[1] || row[1] === "") {
          row[1] = `=IMAGE("${info.icon}")`;
        }
        // C 欄: Name_CN (優先使用 API 抓到的，若 API 沒給則維持現狀)
        if (info.cn) {
          row[2] = info.cn;
        }
        // D 欄: Name_EN
        row[3] = info.en;
        // E 欄: Type
        row[4] = info.type;
        // F 欄: Rarity
        row[5] = info.rarity;
        // G 欄: Level
        row[6] = info.level;
        // H 欄: Chat_Link
        row[7] = info.link;
      }
    });

    // 將更新後的陣列一次寫回試算表
    range.setValues(values);
    console.log(`✨ 字典補全完畢。共處理 ${Object.keys(dbData).length} 個物品。`);
  }
};

/**
 * 手動更新 ItemDB 的按鈕函式
 */
function run_ItemDB_Update() {
  Data_Library_ItemDB.updateItemDatabase();
}
