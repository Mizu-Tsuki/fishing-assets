function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧪 PriceLab 實驗室')
    .addItem('🔄 執行完整同步', 'action_syncAll')
    .addSeparator()
    .addItem('🎨 僅重整表格格式', 'action_onlyFormat')
    .addToUi();
}
