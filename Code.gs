// タスク一覧 列構成: A=タスク名, B=担当者, C=カテゴリ, D=優先度, E=ステータス, F=開始日, G=締切日, H=予定工数(h), I=実績工数(h), J=メモ
// 週別ワークロード列構成: A=担当者
var TASK_SHEET_NAME = 'タスク一覧';

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('タスク管理', [
    {name: 'すべての改善を適用', functionName: 'applyAllImprovements'},
    {name: 'シート構成を更新（列整理）', functionName: 'restructureTaskSheet'},
    {name: '集計テーブルを更新', functionName: 'setupSummaryTables'},
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
  taskSheet.deleteColumn(12);
  taskSheet.deleteColumn(10);
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

// 担当者別集計・カテゴリ別集計テーブルのヘッダーと数式を更新
function setupSummaryTables() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var personSheet = ss.getSheetByName('担当者別集計');
  var categorySheet = ss.getSheetByName('カテゴリ別集計');

  if (!personSheet && !categorySheet) {
    SpreadsheetApp.getUi().alert('「担当者別集計」「カテゴリ別集計」シートが見つかりません。');
    return;
  }

  if (personSheet) setupPersonSummary(personSheet);
  if (categorySheet) setupCategorySummary(categorySheet);

  SpreadsheetApp.getUi().alert('集計テーブルを更新しました。');
}

function setupPersonSummary(sheet) {
  var headers = ['担当者', '総タスク数', '未着手', '進行中', 'レビュー中', '完了', 'ブロック中', '予定工数(h)', '実績工数(h)'];
  var skipValues = ['担当者', 'カテゴリ'];

  // ヘッダー書き込み前にデータを読む
  var data = sheet.getDataRange().getValues();

  // ヘッダーを１行目に書き込む
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 2行目以降で本物の担当者名が入っている行にのみ数式を書き込む
  var written = 0;
  for (var r = 1; r < data.length; r++) {
    var val = String(data[r][0]).trim();
    if (val === '' || val.indexOf('※') === 0 || skipValues.indexOf(val) >= 0) continue;
    var row = r + 1;
    var a = '$A' + row;
    sheet.getRange(row, 2).setFormula("=COUNTIF('" + TASK_SHEET_NAME + "'!$B:$B," + a + ")");
    sheet.getRange(row, 3).setFormula("=COUNTIFS('" + TASK_SHEET_NAME + "'!$B:$B," + a + ",'" + TASK_SHEET_NAME + "'!$E:$E,\"未着手\")");
    sheet.getRange(row, 4).setFormula("=COUNTIFS('" + TASK_SHEET_NAME + "'!$B:$B," + a + ",'" + TASK_SHEET_NAME + "'!$E:$E,\"進行中\")");
    sheet.getRange(row, 5).setFormula("=COUNTIFS('" + TASK_SHEET_NAME + "'!$B:$B," + a + ",'" + TASK_SHEET_NAME + "'!$E:$E,\"レビュー中\")");
    sheet.getRange(row, 6).setFormula("=COUNTIFS('" + TASK_SHEET_NAME + "'!$B:$B," + a + ",'" + TASK_SHEET_NAME + "'!$E:$E,\"完了\")");
    sheet.getRange(row, 7).setFormula("=COUNTIFS('" + TASK_SHEET_NAME + "'!$B:$B," + a + ",'" + TASK_SHEET_NAME + "'!$E:$E,\"ブロッカー\")");
    sheet.getRange(row, 8).setFormula("=SUMIF('" + TASK_SHEET_NAME + "'!$B:$B," + a + ",'" + TASK_SHEET_NAME + "'!$H:$H)");
    sheet.getRange(row, 9).setFormula("=SUMIF('" + TASK_SHEET_NAME + "'!$B:$B," + a + ",'" + TASK_SHEET_NAME + "'!$I:$I)");
    written++;
  }

  if (written > 0) {
    var dataRange = sheet.getRange(2, 1, data.length - 1, headers.length);
    sheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND($A2<>"",$A2<>"※",$G2>0)')
        .setBackground('#FF4444').setFontColor('#FFFFFF').setRanges([dataRange]).build()
    ]);
  }
}

function setupCategorySummary(sheet) {
  var headers = ['カテゴリ', '総タスク数', '未着手', '進行中', 'レビュー中', '完了', 'ブロック中', '予定工数(h)', '実績工数(h)'];
  var skipValues = ['担当者', 'カテゴリ'];

  // ヘッダー書き込み前にデータを読む
  var data = sheet.getDataRange().getValues();

  // ヘッダーを１行目に書き込む
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 2行目以降で本物のカテゴリ名が入っている行にのみ数式を書き込む
  var written = 0;
  for (var r = 1; r < data.length; r++) {
    var val = String(data[r][0]).trim();
    if (val === '' || val.indexOf('※') === 0 || skipValues.indexOf(val) >= 0) continue;
    var row = r + 1;
    var a = '$A' + row;
    sheet.getRange(row, 2).setFormula("=COUNTIF('" + TASK_SHEET_NAME + "'!$C:$C," + a + ")");
    sheet.getRange(row, 3).setFormula("=COUNTIFS('" + TASK_SHEET_NAME + "'!$C:$C," + a + ",'" + TASK_SHEET_NAME + "'!$E:$E,\"未着手\")");
    sheet.getRange(row, 4).setFormula("=COUNTIFS('" + TASK_SHEET_NAME + "'!$C:$C," + a + ",'" + TASK_SHEET_NAME + "'!$E:$E,\"進行中\")");
    sheet.getRange(row, 5).setFormula("=COUNTIFS('" + TASK_SHEET_NAME + "'!$C:$C," + a + ",'" + TASK_SHEET_NAME + "'!$E:$E,\"レビュー中\")");
    sheet.getRange(row, 6).setFormula("=COUNTIFS('" + TASK_SHEET_NAME + "'!$C:$C," + a + ",'" + TASK_SHEET_NAME + "'!$E:$E,\"完了\")");
    sheet.getRange(row, 7).setFormula("=COUNTIFS('" + TASK_SHEET_NAME + "'!$C:$C," + a + ",'" + TASK_SHEET_NAME + "'!$E:$E,\"ブロッカー\")");
    sheet.getRange(row, 8).setFormula("=SUMIF('" + TASK_SHEET_NAME + "'!$C:$C," + a + ",'" + TASK_SHEET_NAME + "'!$H:$H)");
    sheet.getRange(row, 9).setFormula("=SUMIF('" + TASK_SHEET_NAME + "'!$C:$C," + a + ",'" + TASK_SHEET_NAME + "'!$I:$I)");
    written++;
  }

  if (written > 0) {
    var dataRange = sheet.getRange(2, 1, data.length - 1, headers.length);
    sheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND($A2<>"",$A2<>"※",$G2>0)')
        .setBackground('#FF4444').setFontColor('#FFFFFF').setRanges([dataRange]).build()
    ]);
  }
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
  setupSummaryTables();
  SpreadsheetApp.getUi().alert('適用完了。');
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
