/**
 * Data_Sync_Material: 素材展示同步組件
 * 功能：
 * 1. 讀取 Row 2 標籤決定填入內容。
 * 2. 支援 CHAR_DYNAMIC 標籤，從該欄起自動展開所有角色背包數量。
 * 3. 自動從 ItemDB 補全中文與英文名稱。
 */
const Sync_Material = {

  sync: function() {
    // 使用你設定的 TAG: SYNC_MATERIAL
    const sheet = Utils.getSheetByTag("SYNC_MATERIAL");
    if (!sheet) {
      console.error("❌ 找不到標籤為 SYNC_MATERIAL 的分頁。請檢查 A1 是否正確。");
      return;
    }

    const startRow = 11; // 資料起始行
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) {
      console.warn("⚠️ 展示表內無 ID 資料。");
      return;
    }

    // 1. 讀取 Row 2 的標籤定義
    const configRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    const dynamicStartIdx = configRow.indexOf("CHAR_DYNAMIC");

    // 2. 獲取全帳號資產與角色清單 (從 Cache 讀取)
    const { assetMap, charNames } = this._getAccountAssets();
    
    // 3. 處理動態區：清空並重新寫入角色標題 (Row 10)
    if (dynamicStartIdx !== -1) {
      const maxCols = sheet.getMaxColumns();
      const dynamicColStart = dynamicStartIdx + 1; // 轉為 1-based index
      
      // 清空從 CHAR_DYNAMIC 開始往右的所有內容 (Row 10 及其下方所有列)
      if (maxCols >= dynamicColStart) {
        sheet.getRange(10, dynamicColStart, sheet.getMaxRows() - 9, maxCols - dynamicStartIdx).clearContent();
      }
      // 在 Row 10 寫入角色名稱
      if (charNames.length > 0) {
        sheet.getRange(10, dynamicColStart, 1, charNames.length).setValues([charNames]);
      }
    }

    // 4. 讀取 A 欄 ID
    const ids = sheet.getRange(startRow, 1, lastRow - startRow + 1, 1).getValues().map(r => String(r[0]).trim());
    
    // 5. 獲取 ItemDB 字典映射
    const itemDb = this._getItemDbMap();

    // 6. 構建資料矩陣
    const totalCols = (dynamicStartIdx !== -1) ? (dynamicStartIdx + charNames.length) : configRow.length;
    
    const finalValues = ids.map(id => {
      const rowData = new Array(totalCols).fill("");
      if (!id || id === "") return rowData;

      const assets = assetMap[id] || { total: 0, mats: 0, bank: 0, charAll: 0, chars: {} };
      const info = itemDb[id] || { cn: "", en: "" };

      // 依照 Row 2 的標籤填入固定區
      configRow.forEach((tag, idx) => {
        if (dynamicStartIdx !== -1 && idx >= dynamicStartIdx) return;
        
        if (tag === "ID") rowData[idx] = id;
        else if (tag === "NAME_CN") rowData[idx] = info.cn;
        else if (tag === "NAME_EN") rowData[idx] = info.en;
        else if (tag === "TOTAL") rowData[idx] = assets.total;
        else if (tag === "MATS") rowData[idx] = assets.mats;
        else if (tag === "BANK") rowData[idx] = assets.bank;
        else if (tag === "CHAR_ALL") rowData[idx] = assets.charAll;
      });

      // 填入動態角色包包數量
      if (dynamicStartIdx !== -1) {
        charNames.forEach((name, i) => {
          rowData[dynamicStartIdx + i] = assets.chars[name] || 0;
        });
      }

      return rowData;
    });

    // 7. 一次性回寫
    sheet.getRange(startRow, 1, finalValues.length, totalCols).setValues(finalValues);
    console.log(`✨ 同步成功！已處理 ${ids.length} 項物品，自動展開 ${charNames.length} 個角色。`);
  },

  /**
   * 私有函式：從 Cache 建立資產地圖與角色名單
   */
  _getAccountAssets: function() {
    const cacheSheet = Utils.getSheetByTag("SYS_CACHE");
    if (!cacheSheet) return { assetMap: {}, charNames: [] };

    const data = cacheSheet.getRange(11, 1, cacheSheet.getLastRow() - 10, 2).getValues();
    const assetMap = {};
    const charNamesSet = new Set();

    data.forEach(row => {
      const key = row[0];
      const jsonStr = row[1];
      if (!jsonStr || !key) return;

      try {
        const items = JSON.parse(jsonStr);
        if (key === "ASSET_MATERIALS") {
          this._merge(assetMap, items, "mats");
        } else if (key === "ASSET_BANK") {
          this._merge(assetMap, items, "bank");
        } else if (key.startsWith("CHAR_")) {
          const name = key.replace("CHAR_", "");
          charNamesSet.add(name);
          if (items.bags) {
            items.bags.forEach(bag => {
              if (bag && bag.inventory) {
                this._merge(assetMap, bag.inventory, "char", name);
              }
            });
          }
        }
      } catch (e) {
        console.warn(`解析 ${key} 失敗:`, e);
      }
    });

    return { assetMap, charNames: Array.from(charNamesSet).sort() };
  },

  _merge: function(map, items, type, charName) {
    items.forEach(item => {
      if (!item || !item.id) return;
      const id = String(item.id);
      if (!map[id]) map[id] = { total: 0, mats: 0, bank: 0, charAll: 0, chars: {} };
      
      const count = item.count || 0;
      map[id].total += count;
      if (type === "mats") map[id].mats += count;
      else if (type === "bank") map[id].bank += count;
      else if (type === "char") {
        map[id].charAll += count;
        map[id].chars[charName] = (map[id].chars[charName] || 0) + count;
      }
    });
  },

  _getItemDbMap: function() {
    const dbSheet = Utils.getSheetByTag("SYS_ITEMDB");
    if (!dbSheet) return {};
    const lastRow = dbSheet.getLastRow();
    if (lastRow < 11) return {};
    const data = dbSheet.getRange(11, 1, lastRow - 10, 4).getValues();
    const map = {};
    data.forEach(r => { map[String(r[0])] = { cn: r[2], en: r[3] }; });
    return map;
  }
};

/**
 * 手動同步入口
 */
function run_Material_Sync() {
  Sync_Material.sync();
}
