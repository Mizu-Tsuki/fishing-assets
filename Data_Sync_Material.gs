/**
 * Data_Sync_Material: 素材展示同步組件 (字典連動版)
 */
const Sync_Material = {

  /**
   * 執行同步主程式
   */
  sync: function() {
    const sheet = Utils.getSheetByTag("SYNC_MATERIAL");
    if (!sheet) return console.error("找不到 SYNC_MATERIAL 分頁");

    const startRow = 11;
    
    // 1. 取得全帳號資產與所有素材 ID / Category 映射
    // assetMap: { id: { mats: 0, bank: 0, chars: { "角色名": 0 }, totalChars: 0 } }
    const { assetMap, charNames, materialIds, categoryMap } = this._getAccountAssets();

    if (materialIds.length === 0) return console.warn("素材庫快取為空，請先執行資產抓取。");

    // 2. 呼叫字典經理：確保所有素材 ID 都有中文名、Icon 與 Category
    // 這一步會自動補全 SYS_ITEMDB 並在有異動時整張重刷字典
    const itemDbMap = Item_Manager.getOrUpdateItems(materialIds, categoryMap);

    // 3. 準備渲染矩陣 (以素材庫官方排序為準)
    const matrix = materialIds.map(id => {
      const item = itemDbMap[id] || {};
      const assets = assetMap[id] || { mats: 0, bank: 0, chars: {}, totalChars: 0 };
      
      const total = assets.mats + assets.bank + assets.totalChars;

      // 基礎欄位：ID, Icon, Name_CN, Total, Mats, Bank, Char_Sum
      const row = [
        id,                                // A: ID
        item.icon || "",                   // B: Icon (URL)
        item.name_cn || `未知物(ID:${id})`, // C: Name_CN
        total,                             // D: Total
        assets.mats,                       // E: Mats (素材庫)
        assets.bank,                       // F: Bank (銀行)
        assets.totalChars                  // G: Char_All (包包總計)
      ];

      // 動態欄位：依照角色清單填入各別包包數量 (從 H 欄開始)
      charNames.forEach(name => {
        row.push(assets.chars[name] || 0);
      });

      return row;
    });

    // 4. 寫入表格
    this._renderToSheet(sheet, matrix, charNames, startRow);
  },

  /**
   * 內部私有：拆解所有資產 JSON 並計算數量
   */
  _getAccountAssets: function() {
    const cacheSheet = Utils.getSheetByTag("SYS_CACHE");
    const data = cacheSheet.getRange(11, 1, cacheSheet.getLastRow() - 10, 2).getValues();
    
    const assetMap = {};      // 存 ID 對應的各處數量
    const charNamesSet = new Set();
    const materialIds = [];   // 存素材庫 ID (決定排序)
    const categoryMap = {};   // 存 ID 對應的素材分類

    data.forEach(row => {
      const key = row[0];
      const jsonStr = row[1];
      if (!jsonStr || jsonStr === "") return;

      const items = JSON.parse(jsonStr);

      if (key === "ASSET_MATERIALS") {
        items.forEach(item => {
          if (!item) return;
          const id = String(item.id);
          materialIds.push(id);
          categoryMap[id] = item.category; // 記錄分類編號給字典用
          this._addCount(assetMap, id, "mats", item.count);
        });
      } else if (key === "ASSET_BANK") {
        items.forEach(item => {
          if (item) this._addCount(assetMap, String(item.id), "bank", item.count);
        });
      } else if (key.startsWith("CHAR_")) {
        const charName = key.replace("CHAR_", "");
        charNamesSet.add(charName);
        items.forEach(bag => {
          if (bag && bag.inventory) {
            bag.inventory.forEach(item => {
              if (item) this._addCount(assetMap, String(item.id), "chars", item.count, charName);
            });
          }
        });
      }
    });

    return { 
      assetMap, 
      charNames: Array.from(charNamesSet).sort(), 
      materialIds, 
      categoryMap 
    };
  },

  /**
   * 累加數量工具
   */
  _addCount: function(map, id, type, count, charName = null) {
    if (!map[id]) map[id] = { mats: 0, bank: 0, chars: {}, totalChars: 0 };
    if (type === "chars") {
      map[id].chars[charName] = (map[id].chars[charName] || 0) + count;
      map[id].totalChars += count;
    } else {
      map[id][type] += count;
    }
  },

  /**
   * 執行最後的渲染與格式調整
   */
  _renderToSheet: function(sheet, matrix, charNames, startRow) {
    // 1. 處理標題 (H 欄以後的角色名)
    const headerRange = sheet.getRange(10, 8, 1, Math.max(sheet.getLastColumn() - 7, 1));
    headerRange.clearContent();
    if (charNames.length > 0) {
      sheet.getRange(10, 8, 1, charNames.length).setValues([charNames]);
    }

    // 2. 清空舊資料區 (避免新表比舊表短)
    const lastRow = sheet.getLastRow();
    if (lastRow >= startRow) {
      sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).clearContent();
    }

    // 3. 寫入主矩陣
    if (matrix.length > 0) {
      sheet.getRange(startRow, 1, matrix.length, matrix[0].length).setValues(matrix);
    }

    console.log(`素材庫同步完成，共更新 ${matrix.length} 筆物品。`);
  }
};
