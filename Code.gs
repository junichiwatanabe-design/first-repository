/**
 * 特定フォルダ内の複数スプレッドシートファイル（案件ごと）から、
 * ファイル名に含まれる日付（例: 2026.7.13）をもとに
 * 起点日付から5日分の案件を抽出し、案件ごとに別タブへ出力する。
 * （印刷時に案件単位で改ページされるよう、1案件＝1タブの構成にしている）
 */

var CONFIG_SHEET_NAME = '設定';
var CONFIG_FOLDER_CELL = 'B1';
var CONFIG_TAB_NAME_CELL = 'B2';
var CONFIG_START_DATE_CELL = 'B3';
var CONFIG_DAYS_CELL = 'B4';
var CONFIG_LABEL_SS_CELL = 'B5';
var DEFAULT_DAYS_TO_READ = 5;
var MAX_DAYS_TO_READ = 5;
var BASE_FONT_SIZE = 14;
var SECONDARY_FONT_SIZE = 16;
var HIGHLIGHT_FONT_SIZE = 20;
var HEADER_BACKGROUND = '#f3f3f3';
var CHECKED_BACKGROUND = '#f4c7c3';
var DARK_BACKGROUND = '#434343';
var WHITE_FONT = '#ffffff';
var MENU_NAME_COLUMN_SPAN = 3; // メニュー名セルをE:G相当の3列分に横結合する
var HIDDEN_COLUMNS = [1, 4, 10]; // A, D, J
var COLUMN_WIDTHS = {
  2: 30,  // B (温製)
  3: 30,  // C (ベジ)
  5: 180, // E
  6: 480, // F
  7: 50,  // G
  9: 200, // I
  11: 200 // K
};
var TIME_COLUMN = 14; // N列

var LABEL_SHEET_NAME = 'ラベル印刷'; // 旧バージョンが残していた集計用シート名（あれば案件扱いから除外する）
var LABEL_SPREADSHEET_SUFFIX = '（ラベル印刷）';
var LABEL_COLS = 3;
var LABEL_ROWS = 8;
var LABEL_COPIES_PER_ENTRY = 2; // 同一ラベルを縦に2枚配置
var LABEL_COL_WIDTH_MM = 70;   // A4・24面ラベルの一般的な実寸（商品が異なる場合は要調整）
var LABEL_ROW_HEIGHT_MM = 33.9;
var MM_TO_PX = 96 / 25.4;
var LABEL_FONT_SIZE = 14;
var LABEL_QTY_FONT_SIZE = 28; // 数量行は太字・大きめフォントで強調する

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日次データ取込')
    .addItem('5日分読み込み実行', 'importFiveDaysData')
    .addItem('印刷用ラベル作成', 'generateMenuLabels')
    .addToUi();
}

function importFiveDaysData() {
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

  var summaryLines = [];
  for (var i = 0; i < daysToRead; i++) {
    var targetDate = new Date(startDate.getTime());
    targetDate.setDate(targetDate.getDate() + i);
    var dateStr = formatDateForMatch_(targetDate, timeZone);

    var matchedFiles = allFiles.filter(function (f) {
      return f.name.indexOf(dateStr) !== -1;
    });

    deleteSheetsForDate_(ss, dateStr);

    if (matchedFiles.length === 0) {
      summaryLines.push(dateStr + ': 該当ファイルなし');
      continue;
    }

    var createdTabNames = [];
    var warnings = [];
    matchedFiles.forEach(function (f) {
      try {
        var sourceSheet = SpreadsheetApp.openById(f.id).getSheetByName(tabName);
        if (!sourceSheet) {
          warnings.push(f.name + ': タブ「' + tabName + '」が見つかりません');
          return;
        }
        var values = sourceSheet.getDataRange().getValues();
        if (values.length === 0) {
          warnings.push(f.name + ': データがありません');
          return;
        }

        var companyName = findAdjacentValue_(values, '企業名');
        var displayName = companyName || stripRedundantDatePrefix_(f.name, dateStr);
        var baseName = sanitizeSheetName_(dateStr + ' ' + displayName);
        var newSheetName = uniqueSheetName_(ss, baseName);
        var newSheet = ss.insertSheet(newSheetName);
        newSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
        applyCaseFormatting_(newSheet, values);
        applySheetLayout_(newSheet);
        createdTabNames.push(newSheetName);
      } catch (e) {
        warnings.push(f.name + ': 処理中にエラーが発生しました（' + e.message + '）');
      }
    });

    var line = dateStr + ': ' + matchedFiles.length + '件のファイルを' +
      createdTabNames.length + '個のタブに書き込みました（' + createdTabNames.join(' / ') + '）';
    if (warnings.length > 0) {
      line += '（警告: ' + warnings.join(' / ') + '）';
    }
    summaryLines.push(line);
  }

  ui.alert(summaryLines.join('\n'));
}

function deleteSheetsForDate_(ss, dateStr) {
  ss.getSheets().forEach(function (sheet) {
    if (sheet.getName() === CONFIG_SHEET_NAME) {
      return;
    }
    if (sheet.getName().indexOf(dateStr) !== 0) {
      return;
    }
    if (ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
    }
  });
}

/**
 * 「設定」シートを除く案件タブごとに、日付・企業名・メニュー名・数量を集めて、
 * ラベル専用スプレッドシート内の同名タブへ A4・24面（3列×8行）のラベル
 * （1件あたり縦2枚）を作成する。案件タブの内容は実行のたびに変わるため、
 * 都度シートを走査して集計する。
 */
function generateMenuLabels() {
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
    ui.alert('ラベルに出力できる案件データが見つかりませんでした。' +
      (warnings.length > 0 ? '\n警告: ' + warnings.join(' / ') : ''));
    return;
  }

  var message = createdCases.length + '件の案件・計' + totalEntries + '件のメニューから' +
    'ラベル' + (totalEntries * LABEL_COPIES_PER_ENTRY) + '枚を作成しました。\n' +
    labelSs.getUrl();
  if (warnings.length > 0) {
    message += '\n警告: ' + warnings.join(' / ');
  }
  ui.alert(message);
}

/**
 * ラベル専用スプレッドシートを取得する。「設定」シートに保存済みのIDがあれば
 * それを再利用し、なければ新規作成して同じ親フォルダに置き、IDを保存する。
 */
function getOrCreateLabelSpreadsheet_(ss, configSheet) {
  var savedId = String(configSheet.getRange(CONFIG_LABEL_SS_CELL).getValue() || '').trim();
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

function formatLabelDate_(value, timeZone) {
  if (!value) {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, timeZone, 'M/d');
  }
  return String(value).trim();
}

/**
 * ラベル1件分の3行テキストを組み立てる。
 * 1行目: 日付＋企業名、2行目: メニュー名、3行目: 数量（呼び出し側でリッチテキスト装飾する）
 */
function formatLabelText_(entry) {
  var qtyText = entry.qty === '' || entry.qty == null ? '' : String(entry.qty);
  return entry.date + '　' + entry.company + '\n' + entry.menu + '\n' + qtyText;
}

/**
 * 1案件分の entries を3列×8行のグリッドに配置し、labelSs内の同名タブへ書き込む。
 * 1エントリにつき縦2行（2枚）を使い、24枚（12エントリ）ごとに次の8行ブロック＝次ページへ折り返す。
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

  var grid = [];
  for (var i = 0; i < totalPages * LABEL_ROWS; i++) {
    grid.push(new Array(LABEL_COLS).fill(null));
  }

  var index = 0;
  for (var page = 0; page < totalPages; page++) {
    for (var col = 0; col < LABEL_COLS; col++) {
      for (var slot = 0; slot < entriesPerColumn; slot++) {
        if (index >= entries.length) {
          break;
        }
        var entry = entries[index++];
        var richText = buildLabelRichText_(entry);
        var row0 = page * LABEL_ROWS + slot * LABEL_COPIES_PER_ENTRY;
        for (var copy = 0; copy < LABEL_COPIES_PER_ENTRY; copy++) {
          grid[row0 + copy][col] = richText;
        }
      }
    }
  }

  for (var r = 0; r < grid.length; r++) {
    for (var c = 0; c < LABEL_COLS; c++) {
      if (grid[r][c]) {
        sheet.getRange(r + 1, c + 1).setRichTextValue(grid[r][c]);
      }
    }
  }

  for (var col1 = 1; col1 <= LABEL_COLS; col1++) {
    sheet.setColumnWidth(col1, Math.round(LABEL_COL_WIDTH_MM * MM_TO_PX));
  }
  for (var row1 = 1; row1 <= grid.length; row1++) {
    sheet.setRowHeight(row1, Math.round(LABEL_ROW_HEIGHT_MM * MM_TO_PX));
  }

  // setRichTextValue後に setFontSize を呼ぶと行ごとのフォントサイズ指定が
  // 上書きされてしまうため、ここでは折り返し方式と垂直方向の配置のみ設定する
  sheet.getRange(1, 1, grid.length, LABEL_COLS)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment('middle');
}

/**
 * ラベルの3行テキストのうち、3行目（数量）だけ太字・大きめフォントにしたリッチテキストを作る。
 */
function buildLabelRichText_(entry) {
  var text = formatLabelText_(entry);
  var qtyStart = text.lastIndexOf('\n') + 1;
  return SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setTextStyle(0, qtyStart, SpreadsheetApp.newTextStyle().setFontSize(LABEL_FONT_SIZE).build())
    .setTextStyle(qtyStart, text.length, SpreadsheetApp.newTextStyle().setFontSize(LABEL_QTY_FONT_SIZE).setBold(true).build())
    .build();
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

/**
 * 案件シートのラベル文字列（案件実施日・企業名・メニュー名など）を検索して
 * 該当セルに書式を適用する。行番号を固定値で持たずラベル一致で探すことで、
 * テンプレートの行位置が案件ごとに多少ずれても崩れないようにしている。
 */
function applyCaseFormatting_(sheet, values) {
  var lastRow = values.length;
  var lastColAll = values[0].length;
  sheet.getRange(1, 1, lastRow, lastColAll)
    .setFontSize(BASE_FONT_SIZE)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  highlightLabelValue_(sheet, values, '案件実施日', HIGHLIGHT_FONT_SIZE, {
    background: DARK_BACKGROUND,
    fontColor: WHITE_FONT,
    align: 'center',
    numberFormat: 'M/d',
    labelBackground: HEADER_BACKGROUND
  });
  highlightLabelValue_(sheet, values, '企業名', HIGHLIGHT_FONT_SIZE, {
    background: DARK_BACKGROUND,
    fontColor: WHITE_FONT,
    align: 'center',
    labelBackground: HEADER_BACKGROUND
  });
  highlightLabelValue_(sheet, values, '参加人数', SECONDARY_FONT_SIZE, {
    labelBackground: HEADER_BACKGROUND
  });
  highlightLabelValue_(sheet, values, 'パーティー目的', SECONDARY_FONT_SIZE, {
    labelBackground: HEADER_BACKGROUND,
    mergeRight: true
  });
  highlightLabelValue_(sheet, values, 'プランナー', SECONDARY_FONT_SIZE, {
    labelBackground: HEADER_BACKGROUND,
    mergeRight: true
  });

  var menuHeader = findMenuTableHeader_(values);
  if (!menuHeader) {
    return;
  }

  var timeCell = findCellByValue_(values, '時刻');
  if (timeCell && menuHeader.row > timeCell.row) {
    var timeSectionRows = menuHeader.row - timeCell.row; // 時刻見出し行を含む行数
    sheet.getRange(timeCell.row + 1, timeCell.col + 1, timeSectionRows, 1)
      .setHorizontalAlignment('center');
    var itemCol = timeCell.col + 1; // 「項目」列（時刻の右隣）
    sheet.getRange(timeCell.row + 1, itemCol + 1, timeSectionRows, 2).mergeAcross();
    sheet.getRange(timeCell.row + 1, timeCell.col + 1).setBackground(HEADER_BACKGROUND); // 時刻見出し
    sheet.getRange(timeCell.row + 1, itemCol + 1, 1, 2).setBackground(HEADER_BACKGROUND); // 項目見出し
  }

  var headerRow1 = menuHeader.row + 1;
  var lastCol = values[menuHeader.row].length;
  sheet.getRange(headerRow1, 1, 1, lastCol).setFontWeight('bold').setBackground(HEADER_BACKGROUND);

  var dataStartRow1 = headerRow1 + 1;
  var numDataRows = values.length - dataStartRow1 + 1;
  if (numDataRows <= 0) {
    return;
  }

  var menuCol = menuHeader.colsByLabel['メニュー名'];
  applyColumnHighlight_(sheet, menuCol, dataStartRow1, numDataRows, MENU_NAME_COLUMN_SPAN);
  applyColumnHighlight_(sheet, menuHeader.colsByLabel['数量'], dataStartRow1, numDataRows, 1);

  sheet.getRange(headerRow1, 1, numDataRows + 1, lastCol)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  applyCheckboxHighlight_(sheet, menuHeader.colsByLabel['温製'], dataStartRow1, numDataRows);
  applyCheckboxHighlight_(sheet, menuHeader.colsByLabel['ベジ'], dataStartRow1, numDataRows);

  if (menuCol != null) {
    sheet.getRange(headerRow1, menuCol + 1, 1, MENU_NAME_COLUMN_SPAN).merge();
    sheet.getRange(dataStartRow1, menuCol + 1, numDataRows, MENU_NAME_COLUMN_SPAN).mergeAcross();
  }
}

function applySheetLayout_(sheet) {
  HIDDEN_COLUMNS.forEach(function (col) {
    sheet.hideColumns(col);
  });
  Object.keys(COLUMN_WIDTHS).forEach(function (col) {
    sheet.setColumnWidth(Number(col), COLUMN_WIDTHS[col]);
  });
  sheet.getRange(1, TIME_COLUMN, sheet.getMaxRows(), 1).setNumberFormat('H:mm');
}

function highlightLabelValue_(sheet, values, labelText, fontSize, opts) {
  var cell = findCellByValue_(values, labelText);
  if (!cell) {
    return;
  }
  var labelRange = sheet.getRange(cell.row + 1, cell.col + 1).setFontWeight('bold');
  if (opts && opts.labelBackground) {
    labelRange.setBackground(opts.labelBackground);
  }
  if (cell.col + 1 < values[cell.row].length) {
    var valueCols = (opts && opts.mergeRight) ? 2 : 1;
    var valueRange = sheet.getRange(cell.row + 1, cell.col + 2, 1, valueCols)
      .setFontSize(fontSize)
      .setFontWeight('bold');
    if (opts && opts.background) {
      valueRange.setBackground(opts.background);
    }
    if (opts && opts.fontColor) {
      valueRange.setFontColor(opts.fontColor);
    }
    if (opts && opts.border) {
      valueRange.setBorder(true, true, true, true, null, null);
    }
    if (opts && opts.align) {
      valueRange.setHorizontalAlignment(opts.align);
    }
    if (opts && opts.numberFormat) {
      valueRange.setNumberFormat(opts.numberFormat);
    }
    if (opts && opts.mergeRight) {
      valueRange.mergeAcross();
    }
  }
}

function applyColumnHighlight_(sheet, colIndex, startRow1, numRows, span) {
  if (colIndex == null) {
    return;
  }
  sheet.getRange(startRow1, colIndex + 1, numRows, span).setFontSize(HIGHLIGHT_FONT_SIZE).setFontWeight('bold');
}

function applyCheckboxHighlight_(sheet, colIndex, startRow1, numRows) {
  if (colIndex == null) {
    return;
  }
  var range = sheet.getRange(startRow1, colIndex + 1, numRows, 1);
  range.insertCheckboxes();
  var firstCellA1 = range.getCell(1, 1).getA1Notation();
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=' + firstCellA1 + '=TRUE')
    .setBackground(CHECKED_BACKGROUND)
    .setRanges([range])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

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

function findAdjacentValue_(values, labelText) {
  var cell = findCellByValue_(values, labelText);
  if (!cell) {
    return null;
  }
  var row = values[cell.row];
  return cell.col + 1 < row.length ? row[cell.col + 1] : null;
}

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

function extractFolderId_(input) {
  var text = String(input).trim();
  var match = text.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (match) {
    return match[1];
  }
  return text;
}

function formatDateForMatch_(date, timeZone) {
  return Utilities.formatDate(date, timeZone, 'yyyy.M.d');
}

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
