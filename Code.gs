// タスク一覧 列構成: A=タスク名, B=担当者, C=カテゴリ, D=優先度, E=ステータス, F=開始日, G=締切日, H=予定工数(h), I=実績工数(h), J=ブロッカー
// 週別ワークロード列構成: A=担当者
var TASK_SHEET_NAME = 'タスク一覧';

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('タスク管理', [
    {name: 'すべての改善を適用', functionName: 'applyAllImprovements'}
  ]);
}

function applyAllImprovements() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var taskSheet = ss.getSheetByName(TASK_SHEET_NAME);
  if (!taskSheet) {
    SpreadsheetApp.getUi().alert('シート「' + TASK_SHEET_NAME + '」が見つかりません。');
    return;
  }
  applyConditionalFormatting(taskSheet);
  addDataValidation(taskSheet);
  setupKpiSheet(ss, taskSheet);
  SpreadsheetApp.getUi().alert('適用完了。KPIダッシュボードを確認してください。');
}

function applyConditionalFormatting(taskSheet) {
  var rules = taskSheet.getConditionalFormatRules();
  var range = taskSheet.getRange('A2:K100');
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($G2<TODAY(),$E2<>"完了",$G2<>"")')
    .setBackground('#FFCCCC').setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$J2="あり"')
    .setBackground('#FFF3CD').setRanges([range]).build());
  taskSheet.setConditionalFormatRules(rules);
}

function addDataValidation(taskSheet) {
  taskSheet.getRange('E2:E100').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['未着手', '進行中', 'レビュー中', '完了', '保留'], true)
      .setAllowInvalid(false).build());
  taskSheet.getRange('D2:D100').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['高', '中', '低'], true)
      .setAllowInvalid(false).build());
  taskSheet.getRange('J2:J100').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['あり', 'なし'], true)
      .setAllowInvalid(false).build());
}

// 週別ワークロード用SUMPRODUCT数式を返す
// 週別ワークロード: A列=担当者（memberCell例: "$A4"）、週開始日セル例: "C3"
// タスク一覧: B列=担当者, F列=開始日, G列=締切日, H列=予定工数
function weeklyWorkloadFormula(memberCell, weekStartCell) {
  return 'IF(' + memberCell + '="","",IFERROR(SUMPRODUCT('
    + '(\'タスク一覧\'!$B$2:$B$1000=' + memberCell + ')'
    + '*(\'タスク一覧\'!$F$2:$F$1000<=' + weekStartCell + '+6)'
    + '*(\'タスク一覧\'!$G$2:$G$1000>=' + weekStartCell + ')'
    + '*(\'タスク一覧\'!$H$2:$H$1000)'
    + '),0))';
}

function setupKpiSheet(ss, taskSheet) {
  var kpiSheet = ss.getSheetByName('KPIダッシュボード') || ss.insertSheet('KPIダッシュボード');
  kpiSheet.clearContents();
  kpiSheet.clearConditionalFormatRules();
  var sn = taskSheet.getName();

  kpiSheet.getRange('A1').setValue('キッチン部門 KPIダッシュボード').setFontSize(16).setFontWeight('bold');
  var header = kpiSheet.getRange('A3:C3');
  kpiSheet.getRange('A3').setValue('指標');
  kpiSheet.getRange('B3').setValue('現在値');
  kpiSheet.getRange('C3').setValue('目標');
  header.setBackground('#4A90D9').setFontColor('#FFFFFF').setFontWeight('bold');

  kpiSheet.getRange('A4').setValue('今週の売上（円）※手動入力');
  kpiSheet.getRange('B4').setValue(0);

  var rows = [
    ['今週の総実績工数(h)', "=SUMIF('" + sn + "'!E:E,\"<>完了\",'" + sn + "'!I:I)", ''],
    ['人時生産性（円/h）', '=IFERROR(B4/B5,"入力してください")', '15000'],
    ['目標達成率', '=IFERROR(B6/15000,"---")', '100%'],
    ['締切超過タスク数', "=COUNTIFS('" + sn + "'!G:G,\"<\"&TODAY(),'" + sn + "'!E:E,\"<>完了\",'" + sn + "'!G:G,\"<>\"&\"\")", '0'],
    ['ブロック中タスク数', "=COUNTIF('" + sn + "'!J:J,\"あり\")", '0']
  ];
  for (var i = 0; i < rows.length; i++) {
    kpiSheet.getRange(5 + i, 1).setValue(rows[i][0]);
    kpiSheet.getRange(5 + i, 2).setValue(rows[i][1]);
    kpiSheet.getRange(5 + i, 3).setValue(rows[i][2]);
  }

  var kpiRange = kpiSheet.getRange('B6');
  kpiSheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(15000).setBackground('#C6EFCE').setRanges([kpiRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(15000).setBackground('#FFCCCC').setRanges([kpiRange]).build()
  ]);

  kpiSheet.setColumnWidth(1, 280);
  kpiSheet.setColumnWidth(2, 200);
  kpiSheet.setColumnWidth(3, 120);
}
