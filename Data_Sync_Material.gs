/**
 * Data_Sync_Material: 素材展示同步組件 (矩陣重排版)
 */
const Sync_Material = {

  sync: function() {
    const sheet = Utils.getSheetByTag("SYNC_MATERIAL");
    if (!sheet) return console.error("找不到 SYNC_MATERIAL 分頁");

    const startRow = 11;
    
    // 1. 取得 Key 導航圖 (第 2 行的配置)
    // 取得第 2 行所有定義好的 Key，並過濾掉空值
    const keyOrder = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(k => String(k).trim())
                       .filter(k => k !== "");

    if (keyOrder.length === 0) return console.error("第 2 行未定義資料 Key，同步中止。");

    // 2. 準備原始資料源
    const { assetMap, charNames, materialIds, categoryMap } = this._getAccountAssets();
    
    // 修正：去重並過濾無效 ID
    const cleanIds = [...new Set(materialIds)]
      .map(id => String(id))
      .filter(id => id && id !== "null" && id !== "");

    // 3. 呼叫字典經理取得靜態資訊 (如名稱、圖示)
    const itemDbMap = Item_Manager.getOrUpdateItems(cleanIds, categoryMap);

    // 4. 第一階段：建立「純資料物件池」 (Data Objects)
    const rawDataObjects = cleanIds.map(id => {
      const item = itemDbMap[id] || {};
      const assets = assetMap[id] || { mats: 0, bank: 0, chars: {}, totalChars: 0 };
      const total = assets.mats + assets.bank + assets.totalChars;

      // 這裡就是你說的：先把資料都準備好，不管順序
      return {
        "ID": id,
        "ICON_URL": item.icon || "",
        "ICON": item.icon ? `=IMAGE("${item.icon}")` : "",
        "NAME_CN": item.name_cn || `未知物(ID:${id})`,
        "NAME_EN": item.name_en || "",
        "TOTAL": total,
        "MATS": assets.mats,
        "BANK": assets.bank,
        "CHAR_ALL": assets.totalChars,
        // 額外預留：如果未來有動態欄位需求可擴充於此
        ...assets.chars 
      };
    });

    // 5. 第二階段：【矩陣重排】照著第 2 行的 KEY 順序，在記憶體中重新排隊
    const matrix = rawDataObjects.map(obj => {
      return keyOrder.map(key => {
        // 如果 Key 在資料池裡有對應，就填入；沒有則留白
        return obj[key] !== undefined ? obj[key] : "";
      });
    });

    // 6. 渲染與寫入 (此處 matrix 已經是排好順序的二維陣列)
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
   * 最終渲染：清空舊資料並寫入矩陣
   */
  _renderToSheet: function(sheet, matrix, charNames, startRow) {
    // 1. 自動修正角色標題 (從 J10 開始，假設 J10 的 Key 定義為 CHAR_ALL 之後的動態欄位)
    // 註：這部分根據你的實體表格需求調整，目前此程式碼會依照第 2 行的 Key 完整覆蓋
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    // 清除舊的資料內容 (從第 11 行開始)
    if (lastRow >= startRow) {
      sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).clearContent();
    }

    // 2. 寫入重排後的矩陣
    if (matrix.length > 0) {
      sheet.getRange(startRow, 1, matrix.length, matrix[0].length).setValues(matrix);
    }
    
    console.log(`同步完成！矩陣重排筆數：${matrix.length} 筆，欄位數：${matrix[0].length}`);
  }
};

/**
 * 手動啟動器
 */
function run_Material_Sync() {
  Sync_Material.sync();
}
