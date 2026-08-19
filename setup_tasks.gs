/**
 * 竹橋製造 タスク管理シート セットアップ
 *
 * 実行方法:
 *   拡張機能 → Apps Script → setupAll を実行
 */

// ============================================================
// メイン: 全シートをまとめてセットアップ
// ============================================================
function setupAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSettings(ss);
  setupTaskBoard(ss);
  setupSummary(ss);
  ss.setActiveSheet(ss.getSheetByName('📋 タスクボード'));
  Browser.msgBox(
    '✅ セットアップ完了！\n\n' +
    '① ⚙️ 設定シートでメンバー名・カテゴリを確認\n' +
    '② 📋 タスクボードにタスクを追加\n' +
    '③ 状態をプルダウンで更新するだけでOK\n\n' +
    '📊 進捗サマリーは自動集計されます。'
  );
}

// ============================================================
// Sheet 3: ⚙️ 設定
// ============================================================
function setupSettings(ss) {
  var sheet = ss.getSheetByName('⚙️ 設定') || ss.insertSheet('⚙️ 設定');
  sheet.clearContents();
  sheet.clearFormats();

  // メンバーリスト
  sheet.getRange('A1').setValue('メンバー').setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
  var members = ['全員','佐野','高橋','平野','吉野','邱','小畑','渡辺','上野','早瀬','服部'];
  sheet.getRange(2, 1, members.length, 1).setValues(members.map(function(m){ return [m]; }));

  // カテゴリリスト
  sheet.getRange('C1').setValue('カテゴリ').setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
  var categories = ['設備','品質','生産','その他'];
  sheet.getRange(2, 3, categories.length, 1).setValues(categories.map(function(c){ return [c]; }));

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(3, 100);

  // 名前付き範囲（ドロップダウンのソース）
  var existing = ss.getNamedRanges().map(function(r){ return r.getName(); });
  if (existing.indexOf('メンバーリスト') === -1) {
    ss.setNamedRange('メンバーリスト', sheet.getRange('A2:A' + (1 + members.length)));
  } else {
    ss.getRangeByName('メンバーリスト').getSheet();
    ss.removeNamedRange('メンバーリスト');
    ss.setNamedRange('メンバーリスト', sheet.getRange('A2:A' + (1 + members.length)));
  }
  if (existing.indexOf('カテゴリリスト') === -1) {
    ss.setNamedRange('カテゴリリスト', sheet.getRange('C2:C' + (1 + categories.length)));
  } else {
    ss.removeNamedRange('カテゴリリスト');
    ss.setNamedRange('カテゴリリスト', sheet.getRange('C2:C' + (1 + categories.length)));
  }
}

// ============================================================
// Sheet 1: 📋 タスクボード
// ============================================================
function setupTaskBoard(ss) {
  var sheet = ss.getSheetByName('📋 タスクボード') || ss.getSheets()[0];
  sheet.setName('📋 タスクボード');
  sheet.clearContents();
  sheet.clearFormats();
  sheet.clearConditionalFormatRules();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();

  var headers = ['タスクID','タスク名','カテゴリ','担当者','期限','状態','完了日','備考'];
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground('#1F4E79')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // 列幅
  sheet.setColumnWidth(1, 90);   // ID
  sheet.setColumnWidth(2, 260);  // タスク名
  sheet.setColumnWidth(3, 90);   // カテゴリ
  sheet.setColumnWidth(4, 90);   // 担当者
  sheet.setColumnWidth(5, 90);   // 期限
  sheet.setColumnWidth(6, 110);  // 状態
  sheet.setColumnWidth(7, 90);   // 完了日
  sheet.setColumnWidth(8, 200);  // 備考

  // サンプル行（グレーで区別）
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  sheet.getRange(2, 1, 1, headers.length).setValues([
    ['T-001', '【入力例】朝礼・ミーティング', '生産', '全員', today, '🔴 未着手', '', '毎朝8:30 ← この行を参考に入力してください']
  ]).setBackground('#F5F5F5').setFontColor('#999999').setFontStyle('italic');

  // タスクID自動採番（A3以降: T-002, T-003...）
  sheet.getRange('A3:A500').setFormula(''); // 手動IDに変更（番号は自分で振るか、後述のonEdit関数で自動化）

  // ドロップダウン: カテゴリ（C列）
  var settingSheet = ss.getSheetByName('⚙️ 設定');
  var categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settingSheet.getRange('C2:C5'), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('C3:C500').setDataValidation(categoryRule);

  // ドロップダウン: 担当者（D列）
  var memberRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settingSheet.getRange('A2:A12'), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('D3:D500').setDataValidation(memberRule);

  // 日付フォーマット: 期限（E列）・完了日（G列）
  sheet.getRange('E3:E500').setNumberFormat('yyyy/MM/dd');
  sheet.getRange('G3:G500').setNumberFormat('yyyy/MM/dd');

  // ドロップダウン: 状態（F列）
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['🔴 未着手', '🟡 進行中', '🟢 完了'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('F3:F500').setDataValidation(statusRule);

  // 条件付き書式
  var rules = [];
  var dataRange = sheet.getRange('A3:H500');

  // 期限超過（期限<今日 かつ 未完了）→ 行全体を濃い赤
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($E3<TODAY(),$F3<>"🟢 完了",$E3<>"")')
    .setBackground('#FFCCCC')
    .setFontColor('#CC0000')
    .setRanges([dataRange])
    .build());

  // 状態: 完了 → グレー
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F3="🟢 完了"')
    .setBackground('#E8E8E8')
    .setFontColor('#999999')
    .setRanges([dataRange])
    .build());

  // 状態: 進行中 → 薄い黄
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F3="🟡 進行中"')
    .setBackground('#FFF9C4')
    .setRanges([dataRange])
    .build());

  // 状態: 未着手 → 薄い赤
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F3="🔴 未着手"')
    .setBackground('#FFEBEE')
    .setRanges([dataRange])
    .build());

  sheet.setConditionalFormatRules(rules);
}

// ============================================================
// Sheet 2: 📊 進捗サマリー
// ============================================================
function setupSummary(ss) {
  var sheet = ss.getSheetByName('📊 進捗サマリー') || ss.insertSheet('📊 進捗サマリー');
  sheet.clearContents();
  sheet.clearFormats();

  var taskSheet = '📋 タスクボード';

  // タイトル
  sheet.getRange('A1').setValue('📊 進捗サマリー')
    .setFontSize(14).setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
  sheet.getRange('A1:F1').merge().setHorizontalAlignment('center');

  // 集計対象週
  sheet.getRange('A3').setValue('集計対象週（開始日）').setFontWeight('bold');
  sheet.getRange('B3').setBackground('#FFF9C4').setNumberFormat('yyyy/MM/dd')
    .setValue(new Date());
  sheet.getRange('C3').setValue('※ 手動で変更してください').setFontColor('#999999').setFontStyle('italic');

  // ---- 全体完了率 ----
  sheet.getRange('A5').setValue('■ 全体').setFontWeight('bold').setFontSize(12);
  sheet.getRange('A6:C6').setValues([['全タスク数', '完了数', '完了率']]);
  sheet.getRange('A6:C6').setBackground('#DAEEF3').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('A7').setFormula('=COUNTA(\''+taskSheet+'\'!B3:B500)');
  sheet.getRange('B7').setFormula('=COUNTIF(\''+taskSheet+'\'!F3:F500,"🟢 完了")');
  sheet.getRange('C7').setFormula('=IF(A7>0,B7/A7,"")').setNumberFormat('0%');
  sheet.getRange('A7:C7').setHorizontalAlignment('center');

  // ---- メンバー別 ----
  sheet.getRange('A9').setValue('■ メンバー別').setFontWeight('bold').setFontSize(12);
  sheet.getRange('A10:D10').setValues([['担当者', '担当タスク数', '完了数', '完了率']]);
  sheet.getRange('A10:D10').setBackground('#DAEEF3').setFontWeight('bold').setHorizontalAlignment('center');

  var settingSheet = ss.getSheetByName('⚙️ 設定');
  var members = settingSheet.getRange('A2:A12').getValues().flat().filter(function(m){ return m && m !== '全員'; });
  members.forEach(function(member, i) {
    var r = 11 + i;
    sheet.getRange(r, 1).setValue(member);
    sheet.getRange(r, 2).setFormula(
      '=COUNTIF(\''+taskSheet+'\'!D3:D500,"'+member+'")+COUNTIF(\''+taskSheet+'\'!D3:D500,"全員")'
    );
    sheet.getRange(r, 3).setFormula(
      '=COUNTIFS(\''+taskSheet+'\'!D3:D500,"'+member+'",\''+taskSheet+'\'!F3:F500,"🟢 完了")+COUNTIFS(\''+taskSheet+'\'!D3:D500,"全員",\''+taskSheet+'\'!F3:F500,"🟢 完了")'
    );
    sheet.getRange(r, 4).setFormula('=IF(B'+r+'>0,C'+r+'/B'+r+',"")').setNumberFormat('0%');
    if (i % 2 === 0) sheet.getRange(r, 1, 1, 4).setBackground('#F5F5F5');
  });

  var overdueStartRow = 11 + members.length + 2;

  // ---- 期限切れタスク ----
  sheet.getRange(overdueStartRow, 1).setValue('■ 期限超過（未完了）').setFontWeight('bold').setFontSize(12).setFontColor('#CC0000');
  sheet.getRange(overdueStartRow + 1, 1, 1, 4).setValues([['タスク名', '担当者', '期限', '状態']]);
  sheet.getRange(overdueStartRow + 1, 1, 1, 4).setBackground('#FFCCCC').setFontWeight('bold').setHorizontalAlignment('center');
  // FILTER: タスク名・担当者・期限(TEXT変換)・状態 の4列のみ返す
  sheet.getRange(overdueStartRow + 2, 1).setFormula(
    '=IFERROR(FILTER(' +
    'CHOOSE({1,2,3,4},' +
    '\''+taskSheet+'\'!B3:B500,' +
    '\''+taskSheet+'\'!D3:D500,' +
    'TEXT(\''+taskSheet+'\'!E3:E500,"yyyy/MM/dd"),' +
    '\''+taskSheet+'\'!F3:F500),' +
    '\''+taskSheet+'\'!E3:E500<TODAY(),' +
    '\''+taskSheet+'\'!F3:F500<>"🟢 完了",' +
    '\''+taskSheet+'\'!B3:B500<>""),"期限超過のタスクはありません")'
  );

  // 列幅
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 80);

  sheet.setFrozenRows(1);
}
