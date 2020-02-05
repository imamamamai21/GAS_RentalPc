/**
 * シート: 内定者・インターンレンタル を整理するコード (人事用)
 */
var KintoneSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('PC台帳データ');
  this.values = this.sheet.getDataRange().getValues();
  
  /**
   * 元データの目次行から、行数を調べる
   * 列が追加されてもデータが狂わないようにするため
   */
  this.index = {};
  
  this.showTitleError = function(key) {
    Browser.msgBox('データが見つかりません', '表のタイトル名を変えていませんか？' + key, Browser.Buttons.OK);
  }
}
  
KintoneSheet.prototype = {
  getRowKey: function(target) {
    var alfabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
    var index = this.getIndex();
    var targetIndex = index[target];
    var returnKey = (targetIndex > -1) ? alfabet[targetIndex] : '';
    if (!returnKey || returnKey === '') this.showTitleError(target);
    return returnKey;
  },
  getIndex: function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  },
  createIndex: function() {
    const KEY_CAPC = 'CAグループPC番号';
    var filterData = this.values.filter(function(value) {
      return value.indexOf(KEY_CAPC) > -1;
    })[0];
    if(!filterData || filterData.length === 0) {
      this.showTitleError();
      return;
    }
    
    this.index = {// 念の為タイトル名から取得
      caPcNo     : filterData.indexOf(KEY_CAPC),
      rentalNo   : filterData.indexOf('レンタル番号'),
      status     : filterData.indexOf('ステータス'),
      place      : filterData.indexOf('保管場所')
    }
    return this.index;
  },
  /**
   * PC台帳シートから、最新のデータを取得して書き直す
   */
  updateKintoneData: function() {
    const START_ROW = 2;
    // 現在の書き込みを削除
    var sheet = this.sheet;
    sheet.getRange(START_ROW, 1, this.sheet.getLastRow(), this.sheet.getLastColumn()).clearContent();
    // レンタル番号があるデータをすべて書き出し
    var titles = KintonePCData.pcDataSheet.getTitles();
    var data = {caPcNo: [], rentalNo: [], status: [], place: []};
    
    KintonePCData.pcDataSheet.values
      .slice(KintonePCData.pcDataSheet.startRow - 1)
      .filter(function(value) {
        return value[titles.rentalid.index] != '';
      }).forEach(function(value) {
        // 項目ごとに配列直す
        data.caPcNo.push([value[titles.capc_id.index]]);
        data.rentalNo.push([value[titles.rentalid.index]]);
        data.status.push([value[titles.pc_status.index]]);
        data.place.push([value[titles.location.index]]);
      });
    
    var index = this.getIndex();
    Object.keys(data).forEach(function (key) { // 項目ごとに書き込む
      sheet.getRange(START_ROW, index[key] + 1, data[key].length, 1).setValues(data[key]);
    });
  }
}

var kintoneSheet = new KintoneSheet();

function testK() {
  kintoneSheet.updateKintoneData();
}