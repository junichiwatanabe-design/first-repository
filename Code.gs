/**
 * 特定フォルダ内の複数スプレッドシートファイル（案件ごと）から、
 * ファイル名に含まれる日付（例: 2026.7.13）をもとに
 * 起点日付から指定日数分の案件を抽出し、集計用スプレッドシートへ
 * 案件ごとに別タブとしてコピーする（印刷時に案件単位で改ページされるよう
 * 1案件＝1タブの構成にしている）。
 *
 * さらに、集計済みの案件タブから日付・企業名・メニュー名・数量を集めて、
 * ラベル専用の別スプレッドシートへ印刷用ラベル（A4・24面）を作成する機能も持つ。
 *
 * ファイル構成:
 *   1. 設定シート関連
 *   2. 日次データ取込（集計用スプレッドシートへの案件タブ作成）
 *   3. 印刷用ラベル作成
 *   4. 共通ヘルパー
 */

// ============================================================
// 1. 設定シート関連
// ============================================================

var CONFIG_SHEET_NAME = '設定';
var CONFIG_FOLDER_CELL = 'B1';       // 対象フォルダの URL/ID
var CONFIG_TAB_NAME_CELL = 'B2';     // 各ファイル内で読み込む固定タブ名
var CONFIG_START_DATE_CELL = 'B3';   // 起点日付
var CONFIG_DAYS_CELL = 'B4';         // 読み込み日数（1〜MAX_DAYS_TO_READ）
var CONFIG_LABEL_SS_CELL = 'B5';     // ラベル出力先スプレッドシートの URL/ID（自動設定）

function getOrCreateConfigSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (sheet) {
    return sheet;
  }

  sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  sheet.getRange('A1').setValue('対象フォルダURL/ID');
  sheet.getRange('A2').setValue('読み込むタブ名');
  sheet.getRange('A3').setValue('起点日付');
  sheet.getRange('A4').setValue('読み込み日数（1〜' + MAX_DAYS_TO_READ + '、空欄で' + DEFAULT_DAYS_TO_READ + '日）');
  sheet.getRange('A5').setValue('ラベル出力先スプレッドシートID（自動設定・空欄でOK）');
  sheet.getRange('A1:A5').setFontWeight('bold');

  SpreadsheetApp.getUi().alert(
    '「' + CONFIG_SHEET_NAME + '」シートを作成しました。' +
    CONFIG_FOLDER_CELL + '/' + CONFIG_TAB_NAME_CELL + '/' + CONFIG_START_DATE_CELL +
    ' に値を入力してから再度実行してください。'
  );
  return null;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日次データ取込')
    .addItem('データ読込', 'importData')
    .addItem('ラベル作成', 'createLabels')
    .addToUi();
}

// ============================================================
// 2. 日次データ取込（集計用スプレッドシートへの案件タブ作成）
// ============================================================

var DEFAULT_DAYS_TO_READ = 5;
var MAX_DAYS_TO_READ = 5;
var CASE_MENU_FONT_SIZE = 14; // 案件タブの「メニュー名」「数量」列のデータ行だけ適用するフォントサイズ

/**
 * メニュー「日次データ取込」→「データ読込」から呼び出されるメイン関数。
 * 「設定」シートの内容に従い、対象フォルダ内の案件ファイルから起点日付〜指定日数分を
 * 検出し、各案件を集計用スプレッドシート内の個別タブとしてコピーする。
 * 実行のたびに前回までの案件タブは全て削除してから作り直す（deleteAllCaseSheets_）。
 */
function importData() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = getOrCreateConfigSheet_(ss);
  if (!configSheet) {
    return;
  }

  var folderInput = configSheet.getRange(CONFIG_FOLDER_CELL).getValue();
  var tabName = configSheet.getRange(CONFIG_TAB_NAME_CELL).getValue();
  var startDateValue = configSheet.getRange(CONFIG_START_DATE_CELL).getValue();

  if (!folderInput || !tabName || !startDateValue) {
    ui.alert(
      '設定シートの「' + CONFIG_FOLDER_CELL + '」(フォルダURL/ID)、「' +
      CONFIG_TAB_NAME_CELL + '」(タブ名)、「' + CONFIG_START_DATE_CELL +
      '」(起点日付) をすべて入力してください。'
    );
    return;
  }

  var startDate = startDateValue instanceof Date ? startDateValue : new Date(startDateValue);
  if (isNaN(startDate.getTime())) {
    ui.alert('起点日付(' + CONFIG_START_DATE_CELL + ')が日付として認識できません。');
    return;
  }

  var daysValue = configSheet.getRange(CONFIG_DAYS_CELL).getValue();
  var daysToRead = DEFAULT_DAYS_TO_READ;
  if (daysValue !== '' && daysValue !== null) {
    daysToRead = Number(daysValue);
    if (!Number.isInteger(daysToRead) || daysToRead < 1 || daysToRead > MAX_DAYS_TO_READ) {
      ui.alert(
        '「' + CONFIG_DAYS_CELL + '」(読み込み日数) は1〜' + MAX_DAYS_TO_READ + 'の整数で入力してください。'
      );
      return;
    }
  }

  var folderId = extractFolderId_(folderInput);
  var folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    ui.alert('フォルダが見つかりません。フォルダURL/IDを確認してください。');
    return;
  }

  var timeZone = ss.getSpreadsheetTimeZone();
  var allFiles = collectSpreadsheetFiles_(folder, ss.getId());

  // 実行のたびに前回までの案件タブをすべて削除してから作り直す
  // （範囲外の古い日付のタブも含めて残さない）
  deleteAllCaseSheets_(ss);

  var summaryLines = [];
  var allWarnings = [];
  for (var i = 0; i < daysToRead; i++) {
    var targetDate = new Date(startDate.getTime());
    targetDate.setDate(targetDate.getDate() + i);
    var dateStr = formatDateForMatch_(targetDate, timeZone);

    var matchedFiles = allFiles.filter(function (f) {
      return f.name.indexOf(dateStr) !== -1;
    });

    if (matchedFiles.length === 0) {
      summaryLines.push(dateStr + ' : 該当ファイルなし');
      continue;
    }

    var createdDisplayNames = [];
    matchedFiles.forEach(function (f) {
      // 表示する名前は「企業名（またはファイル名から日付・記号を除いたもの）」にする
      var displayName = stripRedundantDatePrefix_(f.name, dateStr);
      try {
        var sourceSheet = SpreadsheetApp.openById(f.id).getSheetByName(tabName);
        if (!sourceSheet) {
          allWarnings.push(dateStr + ' ' + displayName + ' : タブ「' + tabName + '」が見つかりません');
          return;
        }
        var values = sourceSheet.getDataRange().getValues();
        if (values.length === 0) {
          allWarnings.push(dateStr + ' ' + displayName + ' : データがありません');
          return;
        }

        var companyName = findAdjacentValue_(values, '企業名');
        displayName = companyName || displayName;
        var baseName = sanitizeSheetName_(dateStr + ' ' + displayName);
        var newSheetName = uniqueSheetName_(ss, baseName);

        // 元シートの書式（フォント・背景色・罫線・セル結合・列幅など）を保つため
        // getValues()+setValues() ではなく copyTo() でシートごと複製する。
        var newSheet = sourceSheet.copyTo(ss);
        newSheet.setName(newSheetName);

        // ただし copyTo() は数式（他シート参照など）もそのままコピーしてしまい、
        // 集計先で参照が切れて #REF! 等のエラーになることがある。コピー前に
        // 評価済みの値（values）で上書きし、セルの中身を確定値にする（書式は変えない）。
        newSheet.getRange(1, 1, values.length, values[0].length).setValues(values);

        applyCaseMenuFontSize_(newSheet, values);
        createdDisplayNames.push(displayName);
      } catch (e) {
        allWarnings.push(dateStr + ' ' + displayName + ' : 処理中にエラーが発生しました（' + e.message + '）');
      }
    });

    summaryLines.push(dateStr + ' : ' + createdDisplayNames.length + '件（' + createdDisplayNames.join(' / ') + '）');
  }

  ui.alert('データ読込 結果', summaryLines.join('\n'), ui.ButtonSet.OK);

  // 警告・エラーは通常の結果に埋もれて見落とされないよう、別ダイアログで目立たせて表示する
  if (allWarnings.length > 0) {
    ui.alert('⚠️要確認', allWarnings.join('\n'), ui.ButtonSet.OK);
  }
}

/**
 * 「設定」シートを除く全ての案件タブを削除する。実行のたびに前回までの
 * 案件タブ（処理対象の日付範囲外だったものも含む）を一掃してから作り直すことで、
 * タブが際限なく増え続けるのを防ぐ。
 */
function deleteAllCaseSheets_(ss) {
  ss.getSheets().forEach(function (sheet) {
    if (sheet.getName() === CONFIG_SHEET_NAME) {
      return;
    }
    if (ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
    }
  });
}

/**
 * 案件タブの「メニュー名」「数量」列のデータ行（見出し行を除く）だけ
 * フォントサイズを設定する。太字・背景色などは変更しない（それ以外の書式は
 * 元ファイルのものをそのまま使う）。
 */
function applyCaseMenuFontSize_(sheet, values) {
  var menuHeader = findMenuTableHeader_(values);
  if (!menuHeader) {
    return;
  }
  var dataStartRow1 = menuHeader.row + 2;
  var numDataRows = values.length - dataStartRow1 + 1;
  if (numDataRows <= 0) {
    return;
  }
  setColumnFontSize_(sheet, menuHeader.colsByLabel['メニュー名'], dataStartRow1, numDataRows, CASE_MENU_FONT_SIZE);
  setColumnFontSize_(sheet, menuHeader.colsByLabel['数量'], dataStartRow1, numDataRows, CASE_MENU_FONT_SIZE);
}

function setColumnFontSize_(sheet, colIndex, startRow1, numRows, fontSize) {
  if (colIndex == null) {
    return;
  }
  sheet.getRange(startRow1, colIndex + 1, numRows, 1).setFontSize(fontSize);
}

function stripRedundantDatePrefix_(fileName, dateStr) {
  var cleaned = fileName.split(dateStr).join('').trim();
  cleaned = cleaned.replace(/^[【\[]\s*/, '').replace(/[】\]]\s*/, ' ').trim();
  return cleaned || fileName;
}

function sanitizeSheetName_(name) {
  var sanitized = String(name).replace(/[\[\]\*\?\/\\:]/g, '_').trim();
  if (sanitized.length > 100) {
    sanitized = sanitized.substring(0, 100);
  }
  return sanitized || 'シート';
}

function uniqueSheetName_(ss, baseName) {
  var name = baseName;
  var suffix = 2;
  while (ss.getSheetByName(name)) {
    name = baseName + ' (' + suffix + ')';
    suffix++;
  }
  return name;
}

// ============================================================
// 3. 印刷用ラベル作成
// ============================================================

var LABEL_SHEET_NAME = 'ラベル印刷'; // 旧バージョンが残していた集計用シート名（あれば案件扱いから除外する）
var LABEL_SPREADSHEET_SUFFIX = '（ラベル印刷）';

// 面付け（A-one マルチプリンタ用ラベルシール24面: 66mm×33.9mm、3列×8行）
var LABEL_COLS = 3;
var LABEL_ROWS = 8;
var LABEL_COPIES_PER_ENTRY = 2; // 同一ラベルを縦に2枚配置
var LABEL_COL_WIDTH_MM = 66;    // ラベル1枚の実寸幅。商品が異なる場合は要調整
var MM_TO_PX = 96 / 25.4;

// ラベル1枚（実寸33.9mm、96dpi換算で128px）を3段に分ける。
// セル内改行では段ごとに異なる書式（数量だけ中央寄せ等）を付けられないため、
// 「日付＋企業名／メニュー名／数量」を別々のセル（3段）にしている。
// 内訳は34px/44px/50px（元は34/38/56pxだったが、数量段を6px減らしメニュー名段に
// 6px足した。合計128pxは変わらないためラベル1枚の実寸33.9mmは維持される）
var LABEL_SUBROW_HEIGHTS_PX = [34, 44, 50];
var LABEL_SUBROWS = LABEL_SUBROW_HEIGHTS_PX.length;

var LABEL_FONT_SIZE = 14;      // 1段目（日付＋企業名）のフォントサイズ
var LABEL_MENU_FONT_SIZE = 12; // 2段目（メニュー名）。1段目より2pt小さい
var LABEL_QTY_FONT_SIZE = 28;  // 3段目（数量）。太字・大きめフォントで強調する

/**
 * メニュー「日次データ取込」→「ラベル作成」から呼び出されるメイン関数。
 * 「設定」シートを除く案件タブごとに、日付・企業名・メニュー名・数量を集めて、
 * ラベル専用スプレッドシート内の同名タブへ印刷用ラベルを作成する。
 * 案件タブの内容は実行のたびに変わるため、都度シートを走査して集計する。
 */
function createLabels() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = getOrCreateConfigSheet_(ss);
  if (!configSheet) {
    return;
  }
  var timeZone = ss.getSpreadsheetTimeZone();
  var labelSs = getOrCreateLabelSpreadsheet_(ss, configSheet);

  var caseSheetNames = [];
  var createdCases = [];
  var totalEntries = 0;
  var warnings = [];
  ss.getSheets().forEach(function (sheet) {
    var name = sheet.getName();
    if (name === CONFIG_SHEET_NAME || name === LABEL_SHEET_NAME) {
      return;
    }
    caseSheetNames.push(name);
    try {
      var entries = collectLabelEntries_(sheet, timeZone);
      if (entries.length === 0) {
        return;
      }
      writeLabelSheetForCase_(labelSs, name, entries);
      createdCases.push(name);
      totalEntries += entries.length;
    } catch (e) {
      warnings.push(name + ': 処理中にエラーが発生しました（' + e.message + '）');
    }
  });

  removeStaleLabelSheets_(labelSs, caseSheetNames);

  if (createdCases.length === 0) {
    ui.alert('ラベルに出力できる案件データが見つかりませんでした。');
  } else {
    var message = createdCases.length + '件の案件・計' + totalEntries + '件のメニューから' +
      'ラベル' + (totalEntries * LABEL_COPIES_PER_ENTRY) + '枚を作成しました。\n' +
      labelSs.getUrl();
    ui.alert('ラベル作成 結果', message, ui.ButtonSet.OK);
  }

  // 警告・エラーは通常の結果に埋もれて見落とされないよう、別ダイアログで目立たせて表示する
  if (warnings.length > 0) {
    ui.alert('⚠️要確認', warnings.join('\n'), ui.ButtonSet.OK);
  }
}

/**
 * ラベル専用スプレッドシートを取得する。「設定」シートに保存済みのURL/IDがあれば
 * それを再利用し、なければ新規作成して同じ親フォルダに置き、IDを保存する。
 */
function getOrCreateLabelSpreadsheet_(ss, configSheet) {
  var savedRaw = configSheet.getRange(CONFIG_LABEL_SS_CELL).getValue();
  var savedId = savedRaw ? extractSpreadsheetId_(savedRaw) : '';
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {
      // 保存済みIDが無効（削除済みなど）の場合は新規作成にフォールバックする
    }
  }

  var labelSs = SpreadsheetApp.create(ss.getName() + LABEL_SPREADSHEET_SUFFIX);
  try {
    var parents = DriveApp.getFileById(ss.getId()).getParents();
    if (parents.hasNext()) {
      var parentFolder = parents.next();
      var labelFile = DriveApp.getFileById(labelSs.getId());
      parentFolder.addFile(labelFile);
      DriveApp.getRootFolder().removeFile(labelFile);
    }
  } catch (e) {
    // フォルダ移動に失敗してもマイドライブ直下に作成されているため処理は継続する
  }

  configSheet.getRange(CONFIG_LABEL_SS_CELL).setValue(labelSs.getId());
  return labelSs;
}

/**
 * ラベル専用スプレッドシート内で、現在の案件タブ名に対応しないシートを削除する
 * （案件が削除・再作成された場合に古いラベルタブが残らないようにする）。
 */
function removeStaleLabelSheets_(labelSs, currentCaseNames) {
  var sheets = labelSs.getSheets();
  sheets.forEach(function (sheet) {
    if (currentCaseNames.indexOf(sheet.getName()) === -1 && labelSs.getSheets().length > 1) {
      labelSs.deleteSheet(sheet);
    }
  });
}

/**
 * 案件タブ1枚分から、ラベルに出力するエントリ（日付・企業名・メニュー名・数量）を集める。
 * メニュー名・数量のどちらかが空の行はスキップする。
 */
function collectLabelEntries_(sheet, timeZone) {
  var values = sheet.getDataRange().getValues();
  if (values.length === 0) {
    return [];
  }

  var dateText = formatLabelDate_(findAdjacentValue_(values, '案件実施日'), timeZone);
  var company = String(findAdjacentValue_(values, '企業名') || '').trim();

  var menuHeader = findMenuTableHeader_(values);
  if (!menuHeader) {
    return [];
  }
  var menuCol = menuHeader.colsByLabel['メニュー名'];
  var qtyCol = menuHeader.colsByLabel['数量'];
  if (menuCol == null || qtyCol == null) {
    return [];
  }

  var entries = [];
  for (var r = menuHeader.row + 1; r < values.length; r++) {
    var menuName = String(values[r][menuCol] || '').trim();
    var qty = values[r][qtyCol];
    if (!menuName || !qty) {
      continue;
    }
    entries.push({ date: dateText, company: company, menu: menuName, qty: qty });
  }
  return entries;
}

/**
 * 「案件実施日」の値をラベル表示用に整形する（ゼロ埋めなし・年なしの `M/d` 形式）。
 */
function formatLabelDate_(value, timeZone) {
  if (!value) {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, timeZone, 'M/d');
  }
  return normalizeDateText_(String(value).trim());
}

/**
 * 「7/16」「2026.7.16」「2026/7/17」など表記が揺れた月日の文字列を、
 * ゼロ埋めなし・年なしの「7/16」形式に統一する。区切り文字（`.`または`/`）で
 * 分割した末尾2つを月・日とみなす（先頭に年が付いていても無視される）。
 * 月日の形式として解釈できない場合は元の文字列をそのまま返す。
 */
function normalizeDateText_(text) {
  var parts = text.split(/[./]/).filter(function (p) { return p !== ''; });
  if (parts.length < 2) {
    return text;
  }
  var month = Number(parts[parts.length - 2]);
  var day = Number(parts[parts.length - 1]);
  if (isNaN(month) || isNaN(day)) {
    return text;
  }
  return month + '/' + day;
}

/**
 * 1案件分の entries を3列×8行（1件＝3段のセル）のグリッドに配置し、labelSs内の
 * 同名タブへ書き込む。1エントリにつき縦2行（2枚）を使い、24枚（12エントリ）ごとに
 * 次の8行ブロック＝次ページへ折り返す。
 * 既に同名タブがあれば削除してから作り直すため、再実行しても古い内容が残らない。
 */
function writeLabelSheetForCase_(labelSs, caseName, entries) {
  var existing = labelSs.getSheetByName(caseName);
  if (existing) {
    labelSs.deleteSheet(existing);
  }
  var sheet = labelSs.insertSheet(caseName);

  var entriesPerColumn = Math.floor(LABEL_ROWS / LABEL_COPIES_PER_ENTRY);
  var entriesPerPage = entriesPerColumn * LABEL_COLS;
  var totalPages = Math.ceil(entries.length / entriesPerPage);
  var totalRows = totalPages * LABEL_ROWS * LABEL_SUBROWS;

  // 最終ページで3列目が1件も埋まらないと、その列が空のままになり印刷範囲が
  // 2列分に縮んで中央寄せがずれてしまう。先に全セルへ空文字列の値を入れておくことで
  // 3列とも印刷範囲に含まれるようにする（実データは後段の書き込みで上書きされる）
  sheet.getRange(1, 1, totalRows, LABEL_COLS)
    .setValue('')
    .setFontSize(LABEL_FONT_SIZE)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  var index = 0;
  for (var page = 0; page < totalPages; page++) {
    for (var col = 0; col < LABEL_COLS; col++) {
      for (var slot = 0; slot < entriesPerColumn; slot++) {
        if (index >= entries.length) {
          break;
        }
        var entry = entries[index++];
        for (var copy = 0; copy < LABEL_COPIES_PER_ENTRY; copy++) {
          var physicalSlot = slot * LABEL_COPIES_PER_ENTRY + copy;
          var rowBase = page * LABEL_ROWS * LABEL_SUBROWS + physicalSlot * LABEL_SUBROWS;
          writeLabelCellGroup_(sheet, rowBase, col + 1, entry);
        }
      }
    }
  }

  for (var col1 = 1; col1 <= LABEL_COLS; col1++) {
    sheet.setColumnWidth(col1, Math.round(LABEL_COL_WIDTH_MM * MM_TO_PX));
  }
  for (var r = 0; r < totalRows; r++) {
    var subIndex = r % LABEL_SUBROWS;
    sheet.setRowHeight(r + 1, LABEL_SUBROW_HEIGHTS_PX[subIndex]);
  }
}

/**
 * ラベル1件分（3段）を書き込む。
 *   1段目: 日付＋企業名（左寄せ）
 *   2段目: メニュー名（左寄せ・1段目より2pt小さいフォント。中央寄せだと
 *          はみ出した際に先頭が見えず分かりにくいため左寄せにしている）
 *   3段目: 数量（中央寄せ・太字・大きめフォントで強調）
 * 3段とも WrapStrategy.CLIP のため、行の高さは常に固定（データの長さに応じて
 * 自動で広がらない）。はみ出した分は非表示になる。
 */
function writeLabelCellGroup_(sheet, rowBase, col1, entry) {
  sheet.getRange(rowBase + 1, col1).setValue(entry.date + '　' + entry.company)
    .setFontSize(LABEL_FONT_SIZE)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('bottom')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  sheet.getRange(rowBase + 2, col1).setValue(entry.menu)
    .setFontSize(LABEL_MENU_FONT_SIZE)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  sheet.getRange(rowBase + 3, col1).setValue(entry.qty)
    .setFontSize(LABEL_QTY_FONT_SIZE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

// ============================================================
// 4. 共通ヘルパー
// ============================================================

/**
 * 値の2次元配列からメニュー表のヘッダー行（「メニュー名」を含む行）を探し、
 * その行にある各列見出しの列インデックスを引けるようにして返す。
 */
function findMenuTableHeader_(values) {
  var headerCell = findCellByValue_(values, 'メニュー名');
  if (!headerCell) {
    return null;
  }
  var row = values[headerCell.row];
  var colsByLabel = {};
  for (var c = 0; c < row.length; c++) {
    var text = String(row[c]).trim();
    if (text) {
      colsByLabel[text] = c;
    }
  }
  return { row: headerCell.row, colsByLabel: colsByLabel };
}

/** ラベル文字列（例:「企業名」）が入ったセルの右隣のセルの値を返す。 */
function findAdjacentValue_(values, labelText) {
  var cell = findCellByValue_(values, labelText);
  if (!cell) {
    return null;
  }
  var row = values[cell.row];
  return cell.col + 1 < row.length ? row[cell.col + 1] : null;
}

/** 値の2次元配列から、指定文字列と完全一致するセルの位置（0始まり）を探す。 */
function findCellByValue_(values, targetText) {
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (String(values[r][c]).trim() === targetText) {
        return { row: r, col: c };
      }
    }
  }
  return null;
}

/** フォルダ内のGoogleスプレッドシートファイル一覧を返す（集計用スプレッドシート自身は除外）。 */
function collectSpreadsheetFiles_(folder, excludeFileId) {
  var files = [];
  var iterator = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (iterator.hasNext()) {
    var file = iterator.next();
    if (file.getId() === excludeFileId) {
      continue;
    }
    files.push({ id: file.getId(), name: file.getName() });
  }
  return files;
}

/** フォルダのURLまたは素のIDから、フォルダIDを取り出す。 */
function extractFolderId_(input) {
  var text = String(input).trim();
  var match = text.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (match) {
    return match[1];
  }
  return text;
}

/** スプレッドシートのURLまたは素のIDから、スプレッドシートIDを取り出す。 */
function extractSpreadsheetId_(input) {
  var text = String(input).trim();
  var match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match) {
    return match[1];
  }
  return text;
}

/** ファイル名とのマッチングに使う日付文字列（`yyyy.M.d`形式）を作る。 */
function formatDateForMatch_(date, timeZone) {
  return Utilities.formatDate(date, timeZone, 'yyyy.M.d');
}
