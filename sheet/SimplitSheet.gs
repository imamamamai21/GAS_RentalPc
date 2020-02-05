/**
 * シート: Smplitデータを整理するコード
 */
var SimplitSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('Simplitデータ');
  this.values = this.sheet.getDataRange().getValues();
  this.titleRow = 0;
  this.index = {};
  
  this.createIndex = function() {
    const TEXT_NO = 'レンタル番号';
    var me = this;
    var filterData = (function() {
      for(var i = 0; i < me.values.length; i++) {
        if (me.values[i].indexOf(TEXT_NO) > -1) {
          me.titleRow = i + 1;
          return me.values[i];
        }
      }
    }());
    if(!filterData || filterData.length === 0) {
      showTitleError();
      return;
    }
    
    this.index = {
      rentalNo          : filterData.indexOf(TEXT_NO),
      agreementNo       : filterData.indexOf('契約番号'),
      startDate         : filterData.indexOf('レンタル開始日'),
      startExtensionDate: filterData.indexOf('レンタル延長開始日'),
      endDate           : filterData.indexOf('レンタル終了予定日'),
      contractedParty   : filterData.indexOf('契約先事業所名称'),
      money             : filterData.indexOf('レンタル月額(契約先)')
    };
    return this.index;
  }
}
  
SimplitSheet.prototype = {
  getRowKey: function(target) {
    var targetIndex = this.getIndex()[target];
    var returnKey = (targetIndex > -1) ? 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')[targetIndex] : '';
    if (!returnKey || returnKey === '') showTitleError(target);
    return returnKey;
  },
  getIndex: function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  }
};

var simplitSheet;

var Hoge = function() { this.h = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('Simplitデータ'); }

function t() {
  //var hoge = new Hoge();
  //Logger.log(hoge.h)
  var simplitSheet = new SimplitSheet();
  Logger.log(simplitSheet.values.length);
}