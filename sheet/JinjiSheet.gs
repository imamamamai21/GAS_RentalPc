/**
 * シート: 内定者・インターンレンタル を整理するコード (人事用)
 */
var JinjiSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('内定者・インターンレンタル');
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
  
JinjiSheet.prototype = {
  getSheet: function() {
    return this.sheet;
  },
  getValues: function() {
    return this.values;
  },
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
      agreementNo: filterData.indexOf('契約番号'),
      model      : filterData.indexOf('機種'),
      size       : filterData.indexOf('サイズ'),
      key        : filterData.indexOf('キー'),
      endDate    : filterData.indexOf('契約期限'),
      estimateNo : filterData.indexOf('見積NO'),
      administratorNo         : filterData.indexOf('社員番号'),
      administratorName       : filterData.indexOf('氏名'),
      administratorJob        : filterData.indexOf('種別'),
      administratorJoinDate   : filterData.indexOf('入社日'),
      administratorLeavingDate: filterData.indexOf('退職日'),
      dejieDate  : filterData.indexOf('手配状況'),
      dejiePlace : filterData.indexOf('デヂエ場所')
    }
    return this.index;
  },
  /**
   * 人事データを返す
   * @return [Object]: {rentalNo, agreementNo, caPcNo, model....(indexと同じ配列)}
   */
  getEndListByJinjiData: function() {
    var index = this.getIndex();
    var returnAry = [];
    var sortedEndList = getSortedEndList();
    
    var values = this.getValues();
    for(var i = 3; i < values.length; i++) {
      sortedEndList.forEach(function(value) {
        if (values[i][index.rentalNo] != value.rentalNo) return;
        // デヂエデータから対象のステータスははずす
        if(values[i][index.dejieDate] === 'レンタル返却済' || values[i][index.dejieDate] === '転用' || values[i][index.dejieDate] === 'CA本体完治対象外') return;
        
        var data = {};
        Object.keys(index).forEach(function(key) {
          data[key] = values[i][index[key]];
        });
        returnAry.push(data);
      });
    }
    return returnAry;
  }
}

var jinjiSheet;

function test() {
  jinjiSheet = new JinjiSheet();
  Logger.log(jinjiSheet.values.length);
}