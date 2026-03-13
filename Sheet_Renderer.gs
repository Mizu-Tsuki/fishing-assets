const Sheet_Renderer = {
  renderDashboard(logData) {
    const { DASHBOARD_SN } = SYSTEM_CONFIG.SHEET_NAMES;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(DASHBOARD_SN) || ss.insertSheet(DASHBOARD_SN);
    sheet.clear();
    sheet.getRange("A1").setValue("🚀 系統執行診斷中心").setFontSize(14).setFontWeight("bold");
    sheet.getRange("A1:B1").merge().setBackground("#34495e").setFontColor("#ffffff").setHorizontalAlignment("center");
    const range = sheet.getRange(3, 1, logData.length, 2);
    range.setValues(logData).setBorder(true, true, true, true, true, true, "#dcdde1", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(footerRow = logData.length + 4, 1).setValue("上次更新: " + new Date().toLocaleString());
    sheet.setColumnWidth(1, 150); sheet.setColumnWidth(2, 350);
  },

  render(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const layout = Layout_Reader.getLayout(sheetName);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1) return;

    sheet.setFrozenRows(layout.frozen[0]);
    sheet.setFrozenColumns(layout.frozen[1]);
    const fullRange = sheet.getRange(1, 1, lastRow, lastCol);
    fullRange.setVerticalAlignment("middle").setFontFamily("Microsoft JhengHei");
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    headers.forEach((h, idx) => {
      const config = layout.columns[String(h).trim()] || layout.charDefault;
      sheet.setColumnWidth(idx + 1, config.width);
      sheet.getRange(1, idx + 1, lastRow, 1).setHorizontalAlignment(config.align);
    });
    sheet.getRange(1, 1, 1, lastCol).setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold");
  }
};
