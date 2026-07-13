/**
 * 特定フォルダ内の複数スプレッドシートファイルから、
 * ファイル名に含まれる日付（例: 2026.7.13）をもとに
 * 起点日付から5日分のデータを抽出し、日付ごとのタブに集約する。
 */

var CONFIG_SHEET_NAME = '設定';
var CONFIG_FOLDER_CELL = 'B1';
var CONFIG_TAB_NAME_CELL = 'B2';
var CONFIG_START_DATE_CELL = 'B3';
var DAYS_TO_READ = 5;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日次データ取込')
    .addItem('5日分読み込み実行', 'importFiveDaysData')
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
  for (var i = 0; i < DAYS_TO_READ; i++) {
    var targetDate = new Date(startDate.getTime());
    targetDate.setDate(targetDate.getDate() + i);
    var dateStr = formatDateForMatch_(targetDate, timeZone);

    var matchedFiles = allFiles.filter(function (f) {
      return f.name.indexOf(dateStr) !== -1;
    });

    if (matchedFiles.length === 0) {
      summaryLines.push(dateStr + ': 該当ファイルなし');
      continue;
    }

    var aggregatedRows = [];
    var blocks = [];
    var warnings = [];
    matchedFiles.forEach(function (f) {
      var sourceSheet;
      try {
        sourceSheet = SpreadsheetApp.openById(f.id).getSheetByName(tabName);
      } catch (e) {
        warnings.push(f.name + ': 開けませんでした');
        return;
      }
      if (!sourceSheet) {
        warnings.push(f.name + ': タブ「' + tabName + '」が見つかりません');
        return;
      }
      var values = sourceSheet.getDataRange().getValues();
      if (values.length > 0) {
        blocks.push({ startIndex: aggregatedRows.length, rowCount: values.length });
        aggregatedRows = aggregatedRows.concat(values);
      }
    });

    var outputSheet = getOrClearSheet_(ss, dateStr);
    if (aggregatedRows.length > 0) {
      var maxCols = aggregatedRows.reduce(function (max, row) {
        return Math.max(max, row.length);
      }, 0);
      var paddedRows = aggregatedRows.map(function (row) {
        if (row.length === maxCols) {
          return row;
        }
        return row.concat(new Array(maxCols - row.length).fill(''));
      });
      var writeRange = outputSheet.getRange(1, 1, paddedRows.length, maxCols);
      writeRange.setValues(paddedRows);
      writeRange.setFontSize(10).setFontWeight('normal');
      blocks.forEach(function (block) {
        applyHighlightFormatting_(outputSheet, block.startIndex + 1, block.rowCount);
      });
    }

    var line = dateStr + ': ' + matchedFiles.length + '件のファイルから' + aggregatedRows.length + '行を書き込みました';
    if (warnings.length > 0) {
      line += '（警告: ' + warnings.join(' / ') + '）';
    }
    summaryLines.push(line);
  }

  ui.alert(summaryLines.join('\n'));
}

function applyHighlightFormatting_(sheet, blockStartRow, blockRowCount) {
  var FONT_SIZE = 14;
  function absRow(relRow) {
    return blockStartRow + relRow - 1;
  }

  if (blockRowCount >= 4) {
    sheet.getRange(absRow(4), 6).setFontSize(FONT_SIZE).setFontWeight('bold'); // F4 案件実施日
  }
  if (blockRowCount >= 6) {
    sheet.getRange(absRow(6), 6).setFontSize(FONT_SIZE).setFontWeight('bold'); // F6 時刻
  }
  if (blockRowCount >= 19) {
    sheet.getRange(absRow(11), 14, 9, 1).setFontSize(FONT_SIZE).setFontWeight('bold'); // N11:N19 メニュー名
  }
  if (blockRowCount >= 42) {
    sheet.getRange(absRow(27), 5, 16, 4).setFontSize(FONT_SIZE).setFontWeight('bold'); // E27:H42 数量
  }
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
  sheet.getRange('A1:A3').setFontWeight('bold');

  SpreadsheetApp.getUi().alert(
    '「' + CONFIG_SHEET_NAME + '」シートを作成しました。' +
    CONFIG_FOLDER_CELL + '/' + CONFIG_TAB_NAME_CELL + '/' + CONFIG_START_DATE_CELL +
    ' に値を入力してから再度実行してください。'
  );
  return null;
}

function getOrClearSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clearContents();
    return sheet;
  }
  return ss.insertSheet(sheetName);
}
