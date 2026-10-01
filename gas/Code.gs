/**
 * 賃上げ原資シミュレーターの「提案を保存」を受け取り、スプレッドシートに1行追加する。
 *
 * 設定手順:
 * 1. 保存先のGoogleスプレッドシートを開き、「拡張機能 > Apps Script」を選ぶ
 * 2. このファイルの内容を貼り付けて保存する
 * 3. 「デプロイ > 新しいデプロイ」で種類を「ウェブアプリ」にし、
 *    実行ユーザー「自分」、アクセスできるユーザー「全員」でデプロイする
 * 4. 表示されたウェブアプリURL（.../exec）を、ツールの「保存データ」タブに貼り付ける
 */
var SHEET_NAME = '提案データ';

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  // 見出し行がなければ作る。施策が増えて列が増えた場合は見出しを更新する
  var lastCol = sheet.getLastColumn();
  var current = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (sheet.getLastRow() === 0 || current.length < data.headers.length) {
    sheet.getRange(1, 1, 1, data.headers.length).setValues([data.headers]);
    sheet.setFrozenRows(1);
  }

  sheet.appendRow(data.row);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
