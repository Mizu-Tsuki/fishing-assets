const Data_Sync_Assets = {
  runSync() {
    const { ASSETS_SN } = SYSTEM_CONFIG.SHEET_NAMES;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ASSETS_SN);
    if (!sheet) return [["錯誤", "找不到分頁 " + ASSETS_SN]];

    const characters = Conn_GW2.fetch("/characters");
    if (!characters) return [["錯誤", "API 連線失敗"]];

    const endpoints = ["/account/bank", "/account/materials"];
    characters.forEach(c => endpoints.push(`/characters/${encodeURIComponent(c)}/inventory`));
    const res = Conn_GW2.fetchAll(endpoints);
    
    const [rawBank, rawMaterials] = [res[0], res[1]];
    const charInvs = res.slice(2);

    const inventoryMap = {};
    const ensureItem = (id) => {
      const sId = String(id);
      if (!inventoryMap[sId]) {
        inventoryMap[sId] = { material: 0, bank: 0, bagTotal: 0, chars: {} };
        characters.forEach(c => inventoryMap[sId].chars[c] = 0);
      }
      return sId;
    };

    const materialStorageIds = new Set();
    if (Array.isArray(rawMaterials)) {
      rawMaterials.forEach(item => {
        if (item && item.count > 0) {
          const sId = ensureItem(item.id);
          inventoryMap[sId].material += item.count;
          materialStorageIds.add(sId);
        }
      });
    }

    const processItems = (items, type, charName = null) => {
      if (!Array.isArray(items)) return;
      items.forEach(item => {
        if (item && item.id && item.count > 0) {
          const sId = ensureItem(item.id);
          if (charName) {
            inventoryMap[sId].chars[charName] += item.count;
            inventoryMap[sId].bagTotal += item.count;
          } else {
            inventoryMap[sId][type] += item.count;
          }
        }
      });
    };
    processItems(rawBank, "bank");
    charInvs.forEach((inv, i) => {
      if (inv && inv.bags) {
        inv.bags.forEach(bag => bag && bag.inventory && processItems(bag.inventory, null, characters[i]));
      }
    });

    const allIds = Object.keys(inventoryMap);
    const finalMaterialRows = [];
    const batchSize = 200;
    
    for (let i = 0; i < allIds.length; i += batchSize) {
      const batch = allIds.slice(i, i + batchSize);
      const details = Conn_GW2.fetch(`/items?ids=${batch.join(",")}`);
      
      if (details && Array.isArray(details)) {
        details.forEach(item => {
          const sId = String(item.id);
          const isMaterial = item.type === "Material" || item.type === "CraftingMaterial" || materialStorageIds.has(sId);
          if (isMaterial) {
            const d = inventoryMap[sId];
            const total = d.material + d.bank + d.bagTotal;
            const names = Data_Library_ItemDB.getItemNames(sId);
            const row = [sId, names.cn, names.en, total, d.material, d.bank, d.bagTotal];
            characters.forEach(c => row.push(d.chars[c] || 0));
            finalMaterialRows.push(row);
          }
        });
      }
    }

    const header = ["ID", "物品名稱", "英文名稱", "總數量", "素材庫", "銀行", "包包數量", ...characters];
    sheet.clearContents();
    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold");

    if (finalMaterialRows.length > 0) {
      finalMaterialRows.sort((a, b) => Number(a[0]) - Number(b[0]));
      sheet.getRange(2, 1, finalMaterialRows.length, header.length).setValues(finalMaterialRows);
    }

    return [["連線狀態", "✅ 成功"], ["偵測物品", allIds.length], ["材料入庫", finalMaterialRows.length]];
  }
};
