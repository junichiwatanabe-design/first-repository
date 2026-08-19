/**
 * 竹橋製造 タスク管理シート セットアップ
 *
 * 実行方法:
 *   拡張機能 → Apps Script → setupTaskSheet を実行
 */

var MEMBERS = ['全員', '佐野', '高橋', '平野', '吉野', '邱', '小畑', '渡辺', '上野', '早瀬', '服部'];
var STATUSES = ['未着手', '進行中', '完了'];
var HEADERS  = ['タスク名', '担当者', 'ステータス', '期限', 'メモ'];

function setupTaskSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.setName('タスク一覧');
  sheet.clearContents();
  sheet.clearFormats();
  sheet.clearConditionalFormatRules();

  // ヘッダー行
  var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);
  headerRange.setBackground('#1F4E79')
             .setFontColor('#FFFFFF')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // 列幅
  sheet.setColumnWidth(1, 260); // タスク名
  sheet.setColumnWidth(2, 90);  // 担当者
  sheet.setColumnWidth(3, 90);  // ステータス
  sheet.setColumnWidth(4, 90);  // 期限
  sheet.setColumnWidth(5, 220); // メモ

  // サンプル行（最初から何か入っていると入力イメージが掴みやすい）
  var samples = [
    ['朝礼・ミーティング',       '全員', '未着手', '',  '毎朝8:30'],
    ['発注書の確認・送付',        '佐野', '未着手', '',  ''],
    ['冷凍庫の温度チェック記録',  '全員', '未着手', '',  '1日2回'],
  ];
  sheet.getRange(2, 1, samples.length, HEADERS.length).setValues(samples);

  // ドロップダウン: 担当者（B列）
  var memberRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(MEMBERS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B2:B500').setDataValidation(memberRule);

  // ドロップダウン: ステータス（C列）
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUSES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('C2:C500').setDataValidation(statusRule);

  // 期限列: 日付フォーマット
  sheet.getRange('D2:D500').setNumberFormat('yyyy/MM/dd');

  // 条件付き書式: ステータスが「完了」の行をグレーアウト
  var doneRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C2="完了"')
    .setBackground('#E0E0E0')
    .setFontColor('#888888')
    .setRanges([sheet.getRange('A2:E500')])
    .build();

  // 条件付き書式: ステータスが「進行中」の行を薄い青に
  var inProgressRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C2="進行中"')
    .setBackground('#E8F0FE')
    .setRanges([sheet.getRange('A2:E500')])
    .build();

  sheet.setConditionalFormatRules([inProgressRule, doneRule]);

  Browser.msgBox(
    '✅ タスク管理シートを作成しました！\n\n' +
    '【使い方】\n' +
    '① タスク名を入力\n' +
    '② 担当者をプルダウンで選択（チーム全体は「全員」）\n' +
    '③ ステータスを更新するだけ\n\n' +
    '進行中 → 青, 完了 → グレーに自動で変わります。'
  );
}
