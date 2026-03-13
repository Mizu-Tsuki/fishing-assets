/**
 * Data_Fetch_Asset: 資產收割機
 * 負責從 GW2 API 抓取玩家帳號內的物品資料並存入快取。
 */

const Data_Fetch_Asset = {

  /**
   * 執行全域收割 (主進入點)
   */
  fetchAll: function() {
    console.log("開始執行全域資產收割...");
    
    // 1. 收割銀行
    this.fetchBank();
    
    // 2. 收割素材倉庫
    this.fetchMaterials();

    // 💡 以後如果要加「角色包包」，就加在這裡
    // this.fetchCharactersInventory();

    console.log("全域資產收割完成。");
  },

  /**
   * 抓取銀行資料
   */
  fetchBank: function() {
    const endpoint = "/account/bank";
    const data = Conn_GW2.fetch(endpoint);
    
    if (data) {
      Data_System_Cache.save("ASSET_BANK", data);
      console.log("銀行資料已更新至快取。");
    }
  },

  /**
   * 抓取素材倉庫資料
   */
  fetchMaterials: function() {
    const endpoint = "/account/materials";
    const data = Conn_GW2.fetch(endpoint);
    
    if (data) {
      Data_System_Cache.save("ASSET_MATERIALS", data);
      console.log("素材倉庫資料已更新至快取。");
    }
  }
};

/**
 * 手動測試用函式
 * 在 GAS 編輯器上方選擇此函式並按下「執行」，即可測試收割效果。
 */
function test_Asset_Fetch() {
  Data_Fetch_Asset.fetchAll();
}
