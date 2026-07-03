/**
 * 竹橋製造 タスク管理シート - Apps Script
 *
 * 関数一覧:
 *   buildProductivityTable()  - 人時生産性シートを構築（初回 or 再構築時に実行）
 *   archiveCompletedTasks()   - 完了タスクをアーカイブシートへ移動（月次運用）
 *
 * 実行方法:
 *   拡張機能 → Apps Script → 関数を選択して ▷ 実行
 */

// ============================================================
// 設定
// ============================================================
const LABOR_SHEET_ID   = '1SrBKN0WboWVXT8xTHcNqGdAHxUywmzYmmwKry0y4HDw';
const LABOR_SHEET_NAME = 'シート1'; // 実際のシート名が違う場合はここを変更
const PERSON_COUNT     = 10;        // 担当者人数

const COLOR = {
  darkBlue  : '#1F4E79',
  midBlue   : '#2E75B6',
  lightBlue : '#DAEEF3',
  lightGray : '#F2F7FC',
  yellow    : '#FFF2CC',
  white     : '#FFFFFF',
  green     : '#00B050',
  red       : '#FF0000',
  orange    : '#FFA500',
};

// ============================================================
// buildProductivityTable - 人時生産性シートを構築
// ============================================================
function buildProductivityTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('人時生産性');
  if (!sheet) sheet = ss.insertSheet('人時生産性');
  sheet.clearContents();
  sheet.clearFormats();
  sheet.setConditionalFormatRules([]);

  const N = PERSON_COUNT;

  // ----------------------------------------------------------
  // ブロック1: 労働時間テーブル（行1〜14）
  // 列: A=担当者, B=1w, C=2w, D=3w, E=4w, F=5w, G=合計
  // ----------------------------------------------------------

  sheet.getRange(1, 1, 1, 8).merge()
    .setValue('▼ 労働時間（参照元: 勤怠シート）')
    .setBackground(COLOR.darkBlue).setFontColor(COLOR.white)
    .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');

  sheet.getRange(2, 1, 1, 7)
    .setValues([['担当者', '1w(h)', '2w(h)', '3w(h)', '4w(h)', '5w(h)', '合計(h)']])
    .setBackground(COLOR.midBlue).setFontColor(COLOR.white)
    .setFontWeight('bold').setHorizontalAlignment('center');

  // A3: 担当者名を IMPORTRANGE で取得
  sheet.getRange(3, 1).setFormula(
    `=ARRAYFORMULA(IMPORTRANGE("${LABOR_SHEET_ID}","${LABOR_SHEET_NAME}!A7:A16"))`
  );
  // B3: 週別労働時間を IMPORTRANGE で取得（×24 で時間数に変換）
  sheet.getRange(3, 2).setFormula(
    `=ARRAYFORMULA(IMPORTRANGE("${LABOR_SHEET_ID}","${LABOR_SHEET_NAME}!B7:G16")*24)`
  );
  sheet.getRange(3, 2, N, 6).setNumberFormat('0.0"h"');

  for (let i = 0; i < N; i++) {
    if (i % 2 === 0) sheet.getRange(3 + i, 1, 1, 7).setBackground(COLOR.lightGray);
  }

  // チーム合計行
  const sumRow = 3 + N;
  sheet.getRange(sumRow, 1).setValue('チーム合計').setFontWeight('bold').setBackground(COLOR.lightBlue);
  for (let c = 2; c <= 7; c++) {
    sheet.getRange(sumRow, c)
      .setFormula(`=SUM(${col(c)}3:${col(c)}${2 + N})`)
      .setNumberFormat('0.0"h"').setBackground(COLOR.lightBlue).setFontWeight('bold');
  }
  sheet.getRange(sumRow, 1, 1, 7).setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(sumRow + 1, 20);

  // ----------------------------------------------------------
  // ブロック2: 人時生産性テーブル（行15〜）
  // 列: A=担当者, B=総勤務時間, C=事務工数, D=製造時間,
  //     E=事務時間率, F=売上, G=人時生産性全体, H=製造のみ,
  //     I=目標差, J=製造機会損失額, K=必要削減時間
  // ----------------------------------------------------------
  const T = sumRow + 2;

  sheet.getRange(T, 1, 1, 11).merge()
    .setValue('▼ 人時生産性テーブル（売上を入力→自動計算）')
    .setBackground(COLOR.darkBlue).setFontColor(COLOR.white)
    .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');

  sheet.getRange(T + 1, 1, 1, 11).setValues([[
    '担当者', '総勤務時間(h)', '事務工数(h)', '製造時間(h)',
    '事務時間率', '売上(円)', '人時生産性・全体(¥/h)',
    '人時生産性・製造のみ(¥/h)', '目標差(¥/h)',
    '製造機会損失額(円)', '目標達成に必要な事務削減(h)'
  ]]).setBackground(COLOR.midBlue).setFontColor(COLOR.white)
    .setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  sheet.setRowHeight(T + 1, 52);

  const dataStart = T + 2;
  const dataEnd   = T + 1 + N;

  for (let i = 0; i < N; i++) {
    const r         = dataStart + i;
    const laborRow  = 3 + i;

    sheet.getRange(r, 1).setFormula(`=A${laborRow}`);
    sheet.getRange(r, 2).setFormula(`=G${laborRow}`).setNumberFormat('0.0"h"');
    sheet.getRange(r, 3)
      .setFormula(`=IFERROR(SUMIF('タスク一覧'!$B:$B,A${r},'タスク一覧'!$I:$I),0)`)
      .setNumberFormat('0.0"h"');
    setKpiFormulas(sheet, r);

    if (i % 2 === 0) sheet.getRange(r, 1, 1, 11).setBackground(COLOR.lightGray);
  }

  // チーム合計行
  const TR = dataEnd + 1;
  sheet.getRange(TR, 1).setValue('チーム合計').setFontWeight('bold');
  sheet.getRange(TR, 2).setFormula(`=SUM(B${dataStart}:B${dataEnd})`).setNumberFormat('0.0"h"');
  sheet.getRange(TR, 3).setFormula(`=SUM(C${dataStart}:C${dataEnd})`).setNumberFormat('0.0"h"');
  sheet.getRange(TR, 4).setFormula(`=MAX(B${TR}-C${TR},0)`).setNumberFormat('0.0"h"');
  sheet.getRange(TR, 5).setFormula(`=IF(B${TR}>0,C${TR}/B${TR},"")`).setNumberFormat('0.0%');
  sheet.getRange(TR, 6).setBackground(COLOR.yellow).setNumberFormat('¥#,##0');
  sheet.getRange(TR, 7).setFormula(`=IF(AND(ISNUMBER(F${TR}),B${TR}>0),F${TR}/B${TR},"")`).setNumberFormat('¥#,##0');
  sheet.getRange(TR, 8).setFormula(`=IF(AND(ISNUMBER(F${TR}),D${TR}>0),F${TR}/D${TR},"")`).setNumberFormat('¥#,##0');
  sheet.getRange(TR, 9).setFormula(`=IF(ISNUMBER(G${TR}),G${TR}-15000,"")`).setNumberFormat('¥#,##0');
  sheet.getRange(TR, 10).setFormula(`=IF(C${TR}>0,C${TR}*15000,"")`).setNumberFormat('¥#,##0');
  sheet.getRange(TR, 11).setFormula(`=IF(AND(ISNUMBER(F${TR}),B${TR}>0),B${TR}-(F${TR}/15000),"")`).setNumberFormat('0.0"h"');
  sheet.getRange(TR, 1, 1, 11).setBackground(COLOR.lightBlue).setFontWeight('bold')
    .setBorder(true, true, true, true, true, true);

  // 条件付き書式
  const rules = [];
  const eRng = sheet.getRange(`E${dataStart}:E${TR}`);
  const iRng = sheet.getRange(`I${dataStart}:I${TR}`);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0.3).setBackground(COLOR.red).setFontColor(COLOR.white).setRanges([eRng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.2, 0.3).setBackground(COLOR.orange).setRanges([eRng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setBackground(COLOR.red).setFontColor(COLOR.white).setRanges([iRng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0).setBackground(COLOR.green).setFontColor(COLOR.white).setRanges([iRng]).build());
  sheet.setConditionalFormatRules(rules);

  [100, 90, 90, 90, 80, 110, 145, 155, 110, 140, 165].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setFrozenRows(2);

  Browser.msgBox(
    '✅ 人時生産性シートを作成しました！\n\n' +
    '① 初回のみ: IMPORTRANGE許可ダイアログ → 「許可」\n' +
    '② 黄色セル（F列）に今週の売上を入力\n' +
    '③ 事務時間率 赤=30%超 / 橙=20%超\n' +
    '④ 目標差 赤=¥15,000/h未達 / 緑=達成'
  );
}

// ============================================================
// archiveCompletedTasks - 完了タスクをアーカイブシートへ移動
// ============================================================
function archiveCompletedTasks() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName('タスク一覧');

  if (!src) {
    Browser.msgBox('「タスク一覧」シートが見つかりません。');
    return;
  }

  let archive = ss.getSheetByName('アーカイブ');
  if (!archive) {
    archive = ss.insertSheet('アーカイブ');
    src.getRange(1, 1, 1, src.getLastColumn())
       .copyTo(archive.getRange(1, 1));
  }

  const lastRow  = src.getLastRow();
  if (lastRow < 2) {
    Browser.msgBox('タスクが登録されていません。');
    return;
  }

  const statusCol = 5; // E列=ステータス
  const data      = src.getRange(2, 1, lastRow - 1, src.getLastColumn()).getValues();

  const toDelete = [];
  data.forEach((row, i) => {
    if (row[statusCol - 1] === '完了') toDelete.push(i + 2);
  });

  if (toDelete.length === 0) {
    Browser.msgBox('完了タスクはありません。');
    return;
  }

  toDelete.forEach(r => {
    const rowData = src.getRange(r, 1, 1, src.getLastColumn()).getValues();
    archive.appendRow(rowData[0]);
  });

  toDelete.reverse().forEach(r => src.deleteRow(r));

  Browser.msgBox(`✅ ${toDelete.length}件の完了タスクをアーカイブしました。`);
}

// ============================================================
// ヘルパー関数
// ============================================================

function setKpiFormulas(sheet, r) {
  sheet.getRange(r, 4).setFormula(`=MAX(B${r}-C${r},0)`).setNumberFormat('0.0"h"');
  sheet.getRange(r, 5).setFormula(`=IF(B${r}>0,C${r}/B${r},"")`).setNumberFormat('0.0%');
  sheet.getRange(r, 6).setBackground(COLOR.yellow).setNumberFormat('¥#,##0');
  sheet.getRange(r, 7).setFormula(`=IF(AND(ISNUMBER(F${r}),B${r}>0),F${r}/B${r},"")`).setNumberFormat('¥#,##0');
  sheet.getRange(r, 8).setFormula(`=IF(AND(ISNUMBER(F${r}),D${r}>0),F${r}/D${r},"")`).setNumberFormat('¥#,##0');
  sheet.getRange(r, 9).setFormula(`=IF(ISNUMBER(G${r}),G${r}-15000,"")`).setNumberFormat('¥#,##0');
  sheet.getRange(r, 10).setFormula(`=IF(C${r}>0,C${r}*15000,"")`).setNumberFormat('¥#,##0');
  sheet.getRange(r, 11).setFormula(`=IF(AND(ISNUMBER(F${r}),B${r}>0),B${r}-(F${r}/15000),"")`).setNumberFormat('0.0"h"');
}

function col(n) {
  let s = '';
  while (n > 0) { s = String.fromCharCode(64 + (n - 1) % 26 + 1) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
