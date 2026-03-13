/**
 * App_Actions: 強化流程調度
 */
function action_syncAll() {
  const { DASHBOARD_SN } = SYSTEM_CONFIG.SHEET_NAMES;
  
  try {
    // 1. 執行同步邏輯 (僅處理資料抓取與寫入)
    const diagLog = Data_Sync_Assets.runSync();
    
    // 2. 渲染看板 (顯示診斷結果，讓你知道抓了多少資料)
    if (typeof Sheet_Renderer !== 'undefined') {
      Sheet_Renderer.renderDashboard(diagLog);
    }

    // 原本的自動排版邏輯已移除，現在同步完只會乖乖待在看板
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASHBOARD_SN).activate();

  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ 執行崩潰！錯誤原因：\n" + e.message);
  }
}

/**
 * 對應選單：🎨 僅重整表格格式
 * 當你覺得「肉眼可見」排版歪掉時，手動點這個
 */
function action_onlyFormat() {
  const { ASSETS_SN, MARKET_SN } = SYSTEM_CONFIG.SHEET_NAMES;
  const ui = SpreadsheetApp.getUi();
  
  try {
    [ASSETS_SN, MARKET_SN].forEach(name => {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
      if (sheet) {
        Sheet_Renderer.render(name);
      }
    });
    ui.alert("✅ 樣式重整完成！");
  } catch (e) {
    ui.alert("❌ 重整失敗：\n" + e.message);
  }
}
