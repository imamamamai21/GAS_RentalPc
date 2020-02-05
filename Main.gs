/**
 * v1.0.1
 * 外部に公開 ID: 1dGGkkpj2hWKBATp_EEdxvPl5LIoHjaYa3q6STd2XVnlRdRmoMwbgF_rB
 */

/**
 * エラーpopup 指定行が見つからないときに使う
 */
function showTitleError(key) {
  Browser.msgBox('データが見つかりません', '表のタイトル名を変えていませんか？ : ' + key, Browser.Buttons.OK);
}

/**
 * データを更新する
 * トリガー登録 毎日1時〜2時
 */
function updateData() {
  // PC台帳データを更新する
  kintoneSheet.updateKintoneData();
}