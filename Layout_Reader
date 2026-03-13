const Layout_Reader = {
  getLayout(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.CONFIG_SN);
    const layout = { frozen: [1, 0], columns: {}, styles: {}, charDefault: { width: 80, align: "Center" } };

    if (!configSheet) return layout;
    const data = configSheet.getDataRange().getValues();
    const row = data.find(r => r[0] === sheetName);
    if (!row) return layout;

    if (row[1]) layout.frozen = row[1].toString().split(",").map(s => parseInt(s.trim()));
    if (row[2]) {
      row[2].split("\n").forEach(line => {
        const p = line.split(":").map(s => s.trim());
        if (p[0] === "[角色預設]") layout.charDefault = { width: parseInt(p[1]), align: p[2] || "Center" };
        else if (p[0]) layout.columns[p[0]] = { width: parseInt(p[1]), align: p[2] || "Normal" };
      });
    }
    return layout;
  }
};
