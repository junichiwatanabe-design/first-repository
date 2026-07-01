// タスク一覧 列構成: A=タスク名, B=担当者, C=カテゴリ, D=優先度, E=ステータス, F=開始日, G=締切日, H=予定工数(h), I=実績工数(h), J=メモ
var TASK_SHEET_NAME = 'タスク一覧';

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('タスク管理', [
    {name: 'すべての改善を適用', functionName: 'applyAllImprovements'},
    {name: 'シート構成を更新（列整理）', functionName: 'restructureTaskSheet'},
    {name: '週別ワークロードヘッダーを日付型に修正', functionName: 'fixWeeklyWorkloadHeaders'},
    {name: '週別ワークロード数式を書き込む', functionName: 'setupWeeklyWorkloadFormulas'}
  ]);
}

// タスク一覧の列を整理: J列(ブロッカーあり/なし)削除、K列(ブロッカー内容)→J列(メモ)、L列(メモ)削除
function restructureTaskSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var taskSheet = ss.getSheetByName(TASK_SHEET_NAME);
  if (!taskSheet) {
    SpreadsheetApp.getUi().alert('シート「' + TASK_SHEET_NAME + '」が見つかりません。');
    return;
  }
  taskSheet.deleteColumn(12); // L列（メモ）削除
  taskSheet.deleteColumn(10); // J列（ブロッカーあり/なし）削除
  taskSheet.getRange('J1').setValue('メモ');
  SpreadsheetApp.getUi().alert('列の整理が完了しました。続けて「すべての改善を適用」を実行してください。');
}

// 週別ワークロードのヘッダー行（"6/29週"等のテキスト）を日付型に変換し表示形式を m/d"週" に設定
function fixWeeklyWorkloadHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var workloadSheet = ss.getSheetByName('週別ワークロード');
  if (!workloadSheet) {
    SpreadsheetApp.getUi().alert('シート「週別ワークロード」が見つかりません。');
    return;
  }
  var data = workloadSheet.getDataRange().getValues();
  var fixed = 0;
  var weekPattern = /^(\d+)\/(\d+)週$/;
  var year = new Date().getFullYear();

  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < data[r].length; c++) {
      var val = String(data[r][c]);
      var m = val.match(weekPattern);
      if (m) {
        var month = parseInt(m[1], 10);
        var day = parseInt(m[2], 10);
        var date = new Date(year, month - 1, day);
        var cell = workloadSheet.getRange(r + 1, c + 1);
        cell.setValue(date);
        cell.setNumberFormat('m/d"週"');
        fixed++;
      }
    }
  }
  SpreadsheetApp.getUi().alert(fixed + '件のヘッダーを日付型に変換しました。');
}

// 週別ワークロードテーブルを探してSUMPRODUCT数式を書き込む
function setupWeeklyWorkloadFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var workloadSheet = ss.getSheetByName('週別ワークロード');
  if (!workloadSheet) {
    SpreadsheetApp.getUi().alert('シート「週別ワークロード」が見つかりません。');
    return;
  }
  var data = workloadSheet.getDataRange().getValues();
  var headerRow = -1;

  for (var r = 0; r < data.length; r++) {
    if (String(data[r][0]).trim() === '担当者') {
      for (var c = 1; c < data[r].length; c++) {
        if (data[r][c] instanceof Date) {
          headerRow = r;
          break;
        }
      }
      if (headerRow >= 0) break;
    }
  }

  if (headerRow < 0) {
    SpreadsheetApp.getUi().alert('週別ワークロードのヘッダー行が見つかりません。\nまず「週別ワークロードヘッダーを日付型に修正」を実行してください。');
    return;
  }

  var weekCols = [];
  for (var c = 1; c < data[headerRow].length; c++) {
    if (data[headerRow][c] instanceof Date) weekCols.push(c);
  }

  var written = 0;
  for (var r = headerRow + 1; r < data.length; r++) {
    var member = String(data[r][0]).trim();
    if (member === '' || member === '合計') break;
    for (var i = 0; i < weekCols.length; i++) {
      var c = weekCols[i];
      var weekRef = workloadSheet.getRange(headerRow + 1, c + 1).getA1Notation().replace(/\d+/, '') + (headerRow + 1);
      var rowNum = r + 1;
      var formula = '=IF($A' + rowNum + '="","",IFERROR(SUMPRODUCT('
        + "('" + TASK_SHEET_NAME + "'!$B$2:$B$1000=$A" + rowNum + ')'
        + "*('" + TASK_SHEET_NAME + "'!$F$2:$F$1000<=" + weekRef + "+6)"
        + "*('" + TASK_SHEET_NAME + "'!$G$2:$G$1000>=" + weekRef + ")"
        + "*('" + TASK_SHEET_NAME + "'!$H$2:$H$1000)"
        + '),0))';
      workloadSheet.getRange(r + 1, c + 1).setFormula(formula);
      written++;
    }
  }
  SpreadsheetApp.getUi().alert(written + '件のセルに数式を書き込みました。');
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
  var range = taskSheet.getRange('A2:J100');
  taskSheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($G2<TODAY(),$E2<>"完了",$G2<>"")')
      .setBackground('#FFCCCC').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$E2="ブロッカー"')
      .setBackground('#FF4444').setFontColor('#FFFFFF').setRanges([range]).build()
  ]);
}

function addDataValidation(taskSheet) {
  taskSheet.getRange('E2:E100').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['未着手', '進行中', 'レビュー中', 'ブロッカー', '完了'], true)
      .setAllowInvalid(false).build());
  taskSheet.getRange('D2:D100').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['高', '中', '低'], true)
      .setAllowInvalid(false).build());
}

function weeklyWorkloadFormula(memberCell, weekStartCell) {
  return 'IF(' + memberCell + '="","",IFERROR(SUMPRODUCT('
    + '(\'' + TASK_SHEET_NAME + '\'!$B$2:$B$1000=' + memberCell + ')'
    + '*(\'' + TASK_SHEET_NAME + '\'!$F$2:$F$1000<=' + weekStartCell + '+6)'
    + '*(\'' + TASK_SHEET_NAME + '\'!$G$2:$G$1000>=' + weekStartCell + ')'
    + '*(\'' + TASK_SHEET_NAME + '\'!$H$2:$H$1000)'
    + '),0))';
}

function setupKpiSheet(ss, taskSheet) {
  var kpiSheet = ss.getSheetByName('KPIダッシュボード') || ss.insertSheet('KPIダッシュボード');
  kpiSheet.clearContents();
  kpiSheet.clearConditionalFormatRules();
  var sn = taskSheet.getName();

  kpiSheet.getRange('A1').setValue('キッチン部門 KPIダッシュボード').setFontSize(16).setFontWeight('bold');
  kpiSheet.getRange('A3').setValue('指標');
  kpiSheet.getRange('B3').setValue('現在値');
  kpiSheet.getRange('C3').setValue('目標');
  kpiSheet.getRange('A3:C3').setBackground('#4A90D9').setFontColor('#FFFFFF').setFontWeight('bold');

  kpiSheet.getRange('A4').setValue('今週の売上（円）※手動入力');
  kpiSheet.getRange('B4').setValue(0);

  var rows = [
    ['今週の総実績工数(h)', "=SUMIF('" + sn + "'!E:E,\"<>完了\",'" + sn + "'!I:I)", ''],
    ['人時生産性（円/h）', '=IFERROR(B4/B5,"入力してください")', '15000'],
    ['目標達成率', '=IFERROR(B6/15000,"---")', '100%'],
    ['締切超過タスク数', "=COUNTIFS('" + sn + "'!G:G,\"<\"&TODAY(),'" + sn + "'!E:E,\"<>完了\",'" + sn + "'!G:G,\"<>\"&\"\")", '0'],
    ['ブロッカータスク数', "=COUNTIF('" + sn + "'!E:E,\"ブロッカー\")", '0']
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
