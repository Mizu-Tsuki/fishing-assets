/**
 * Data_System_Itemdb: 物品字典管理器
 * 功能：自動檢查、補全、並同步儲存物品元數據
 */
const Item_Manager = {

  /**
   * 核心進入點：確保字典齊全並回傳完整物品對照表
   * @param {Array} ids - 需要查詢的物品 ID 清單 (數組或字串)
   * @param {Object} categoryMap - 選填，{ "id": category_id } 用於更新素材分類
   * @return {Object} 包含所有 ID 詳細資料的 Map
   */
  getOrUpdateItems: function(ids, categoryMap = {}) {
    if (!ids || ids.length === 0) return {};
    
    // 1. 讀取現有的字典 (從 A11 開始)
    const sheet = Utils.getSheetByTag("SYS_ITEMDB");
    if (!sheet) throw new Error("找不到標籤為 SYS_ITEMDB 的分頁");

    const lastRow = sheet.getLastRow();
    const existingData = lastRow >= 11 ? sheet.getRange(11, 1, lastRow - 10, 9).getValues() : [];
    
    // 將現有資料轉成 Map (以 ID 為 Key)
    const itemMap = {};
    existingData.forEach(row => {
      const id = String(row[0]);
      if (id) {
        itemMap[id] = {
          id: id,
          icon: row[1],
          name_cn: row[2],
          name_en: row[3],
          type: row[4],
          rarity: row[5],
          level: row[6],
          category: row[7],
          chat_link: row[8]
        };
      }
    });

    // 2. 找出缺失的 ID
    const uniqueIds = [...new Set(ids.map(id => String(id)))];
    const missingIds = uniqueIds.filter(id => !itemMap[id] || !itemMap[id].name_cn);

    let isDirty = false;

    // 3. 如果有缺失，分批向 API 請求資料
    if (missingIds.length > 0) {
      console.log(`偵測到 ${missingIds.length} 筆新物品，準備從 API 更新...`);
      const apiData = this._fetchFromAPI(missingIds);
      
      // 合併新資料到 Map
      Object.keys(apiData).forEach(id => {
        itemMap[id] = apiData[id];
        isDirty = true;
      });
    }

    // 4. 更新 Category (如果傳入的 categoryMap 有新資訊且原本為空)
    Object.keys(categoryMap).forEach(id => {
      const sId = String(id);
      if (itemMap[sId] && (!itemMap[sId].category || itemMap[sId].category === "")) {
        itemMap[sId].category = categoryMap[sId];
        isDirty = true;
      }
    });

    // 5. 如果有異動 (補了名字或補了 Category)，整張刷回試算表
    if (isDirty) {
      this._saveToSheet(sheet, itemMap);
    }

    return itemMap;
  },

  /**
   * 內部私有：分批從 GW2 API 獲取中英文資料
   */
  _fetchFromAPI: function(ids) {
    const chunks = [];
    for (let i = 0; i < ids.length; i += 200) {
      chunks.push(ids.slice(i, i + 200));
    }

    const results = {};

    chunks.forEach(chunk => {
      const idString = chunk.join(",");
      
      // 同時請求中文與英文 (透過兩次 Fetch)
      const urlZh = `${SYSTEM_CONFIG.BASE_URL}/items?ids=${idString}&lang=zh`;
      const urlEn = `${SYSTEM_CONFIG.BASE_URL}/items?ids=${idString}&lang=en`;

      try {
        const respZh = JSON.parse(UrlFetchApp.fetch(urlZh).getContentText());
        const respEn = JSON.parse(UrlFetchApp.fetch(urlEn).getContentText());
        
        // 將資料對齊
        respZh.forEach((item, index) => {
          const enItem = respEn[index];
          results[String(item.id)] = {
            id: String(item.id),
            icon: item.icon || "",
            name_cn: item.name || "",
            name_en: enItem ? enItem.name : "",
            type: item.type || "",
            rarity: item.rarity || "",
            level: item.level || 0,
            category: "", // 初始留空，等 Material Sync 補上
            chat_link: item.chat_link || ""
          };
        });
      } catch (e) {
        console.error("API 請求失敗: " + e.message);
      }
    });

    return results;
  },

  /**
   * 內部私有：將記憶體中的 Map 轉為矩陣並寫回 Sheet
   */
  _saveToSheet: function(sheet, itemMap) {
    const matrix = Object.values(itemMap).map(item => [
      item.id,
      item.icon,
      item.name_cn,
      item.name_en,
      item.type,
      item.rarity,
      item.level,
      item.category,
      item.chat_link
    ]);

    if (matrix.length > 0) {
      // 清空 Row 11 以下的舊資料 (避免重刷時殘留舊長度的資料)
      const lastRow = sheet.getLastRow();
      if (lastRow >= 11) {
        sheet.getRange(11, 1, lastRow - 10, 9).clearContent();
      }
      // 一次性寫入新矩陣
      sheet.getRange(11, 1, matrix.length, 9).setValues(matrix);
      console.log(`字典更新完成，共 ${matrix.length} 筆資料已同步。`);
    }
  }
};
