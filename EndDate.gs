/**
 * ２週間以内に終了するレンタルPCリスト
 * @return [Object] [{rentalNo, endDate, agreementNo}]
 */
function getEndList() {
  var returnAry = [];
  var simplitSheet = new SimplitSheet();
  var index = simplitSheet.getIndex();
  var today = new Date();
  var targetValues = simplitSheet.values.filter(function(value) {
    // 全シスのもののみ
    return value[index.contractedParty] === TEXT_CONTRACTED_PARTY_FOR_MINE;
  })
  
  for(var i = 0; i < targetValues.length; i++) {
    var endDate = targetValues[i][index.endDate];
    var dif = (endDate - today) / 1000 / 60 / 60 / 24; // 日で割る
    // 2週間以内のものをpush
    if (dif < 14) {
      returnAry.push({
        rentalNo: targetValues[i][index.rentalNo],
        endDate: endDate,
        agreementNo: targetValues[i][index.agreementNo]
      });
    }
  }
  Logger.log(returnAry);
  return returnAry;
}

/**
 * 終了日が早い順にソートして返す
 * @return [Object] [{rentalNo, endDate, agreementNo}]
 */
function getSortedEndList() {
  return getEndList().sort(function(a, b) {
    if (a.endDate < b.endDate) return -1;
    if (a.endDate > b.endDate) return 1;
    return 0;
  })
}

function test() {
  var t = getEndList()
  Logger.log(t);
}