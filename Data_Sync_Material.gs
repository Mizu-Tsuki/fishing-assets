/**
 * Data_Sync_Material: 素材展示同步組件 (對齊隱藏欄位與 IMAGE 公式)
 */
const Sync_Material = {

  /**
   * 執行同步主程式
   */
  sync: function() {
    const sheet = Utils.getSheetByTag("SYNC_MATERIAL");
    if (!sheet) return console.error("找不到 SYNC_MATERIAL 分頁");

    const startRow = 11;
    
    // 1. 取得原始資產資料 (從 SYS_CACHE 拆解)
    const { assetMap, charNames, materialIds, categoryMap } = this._getAccountAssets();

    // 2. 核心修正：去重並過濾無效 ID，確保筆數精準
    const cleanIds = [...new Set(materialIds)]
      .map(id => String(id))
      .filter(id => id && id !== "null" && id !== "undefined" && id !== "");

    if (cleanIds.length === 0) return console.warn("素材庫快取為空，請先執行資產抓取。");

    // 3. 呼叫字典經理補全缺失資料
    const itemDbMap = Item_Manager.getOrUpdateItems(cleanIds, categoryMap);

    // 4. 準備渲染矩陣 (對齊 A:ID, B:隱藏網址, C:圖示, D:中文, E:英文, F:總計...)
    const matrix = cleanIds.map(id => {
      const item = itemDbMap[id] || {};
      const assets = assetMap[id] || { mats: 0, bank: 0, chars: {}, totalChars: 0 };
      const total = assets.mats + assets.bank + assets.totalChars;

      const rawIconUrl = item.icon || "";
      // 生成給 C 欄使用的圖片公式
      const iconFormula = rawIconUrl ? `=IMAGE("${rawIconUrl}")` : "";

      return [
        id,                                // A: ID
        rawIconUrl,                        // B: 圖示連結 (隱藏欄)
        iconFormula,                       // C: 圖示 (IMAGE公式)
        item.name_cn || `未知物(ID:${id})`, // D: 中文名稱
        item.name_en || "",                // E: 英文名稱
        total,                             // F: 總計
        assets.mats,                       // G: 素材庫
        assets.bank,                       // H: 銀行
        assets.totalChars                  // I: 角色合計
      ];
    });

    // 5. 執行渲染
    this._renderToSheet(sheet, matrix, charNames, startRow);
  },

  /**
   * 內部私有：拆解所有資產 JSON 並計算數量
   */
  _getAccountAssets: function() {
    const cacheSheet = Utils.getSheetByTag("SYS_CACHE");
    const lastRow = cacheSheet.getLastRow();
    if (lastRow < 11) return { assetMap: {}, charNames: [], materialIds: [], categoryMap: {} };

    const data = cacheSheet.getRange(11, 1, lastRow - 10, 2).getValues();
    
    const assetMap = {};
    const charNamesSet = new Set();
    const materialIds = [];
    const categoryMap = {};

    data.forEach(row => {
      const key = row[0];
      const jsonStr = row[1];
      if (!jsonStr) return;

      const items = JSON.parse(jsonStr);

      if (key === "ASSET_MATERIALS") {
        items.forEach(item => {
          if (!item) return;
          const id = String(item.id);
          materialIds.push(id);
          categoryMap[id] = item.category;
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
   * 渲染與清除舊資料
   */
  _renderToSheet: function(sheet, matrix, charNames, startRow) {
    // 處理 J 欄以後的角色標題 (因為現在 I 欄是角色合計)
    const headerRange = sheet.getRange(10, 10, 1, Math.max(sheet.getLastColumn() - 9, 1));
    headerRange.clearContent();
    if (charNames.length > 0) {
      sheet.getRange(10, 10, 1, charNames.length).setValues([charNames]);
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow >= startRow) {
      sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).clearContent();
    }

    if (matrix.length > 0) {
      // 根據 matrix[0].length 自動決定寫入寬度
      sheet.getRange(startRow, 1, matrix.length, matrix[0].length).setValues(matrix);
    }
    console.log(`同步完成！素材總數：${matrix.length} 筆。`);
  }
};

/**
 * 手動啟動器
 */
function run_Material_Sync() {
  Sync_Material.sync();
}
