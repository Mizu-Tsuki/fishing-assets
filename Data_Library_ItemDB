const Data_Library_ItemDB = {
  _cache: null,
  getItemNames(id) {
    try {
      if (!this._cache) {
        const { ITEM_DB_SN } = SYSTEM_CONFIG.SHEET_NAMES;
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ITEM_DB_SN);
        this._cache = {};
        if (sheet) {
          const data = sheet.getDataRange().getValues();
          for (let i = 1; i < data.length; i++) {
            if (data[i][0]) this._cache[String(data[i][0])] = { cn: data[i][1] || "Unknown", en: data[i][2] || "" };
          }
        }
      }
      const res = this._cache[String(id)];
      return res ? res : { cn: "Unknown [" + id + "]", en: "" };
    } catch (e) {
      return { cn: "DB Error", en: "" };
    }
  }
};
