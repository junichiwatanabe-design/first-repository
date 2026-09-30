/**
 * 「設定」シートに手入力した案件ファイルのURLを1件ずつ開き、ファイル名から
 * 案件実施日・企業名を読み取って一覧に表示するとともに、リンク先ファイルの
 * 全タブの内容を集計用スプレッドシートへ案件ごとに別タブとしてコピーする
 * （印刷時に案件単位で改ページされるよう1案件＝1タブの構成にしている）。
 *
 * さらに、集計済みの案件タブから日付・企業名・メニュー名・数量を集めて、
 * ラベル専用の別スプレッドシートへ印刷用ラベル（A4・24面）を作成する機能も持つ。
 *
 * ファイル構成:
 *   1. 設定シート関連
 *   2. リンクからの案件読み込み
 *   3. 印刷用ラベル作成
 *   4. 共通ヘルパー
 */

// ============================================================
// 1. 設定シート関連
// ============================================================

var CONFIG_SHEET_NAME = '設定';
var CONFIG_LABEL_SS_CELL = 'B1';          // ラベル出力先スプレッドシートの URL/ID（自動設定）
var CONFIG_PRINT_MARGIN_CELL = 'B2';      // 印刷時余白のメモ（スクリプトからは読み書きしない）
var CONFIG_PRINT_MARGIN_DEFAULT = '上下12mm・左右5mm程度';
var CONFIG_LINKS_HEADER_ROW = 4;          // 案件ファイルのリンク一覧の見出し行
var CONFIG_LINKS_DATA_START_ROW = 5;      // 案件ファイルのリンク一覧のデータ開始行
var LEGACY_LINKS_SHEET_NAME = '案件リンク一覧'; // 旧バージョンで使っていたリンク一覧シート名（あれば「設定」へ自動移行して削除する）
var LEGACY_TAB_NAME_LABEL = '読み込むタブ名'; // 旧バージョンのA1見出し（タブ名を固定指定していた名残。あれば1行目ごと削除して移行する）

/**
 * 「設定」シートを取得する。無ければ新規作成し、見出しラベルを設定したうえで
 * アラートを出して処理を中断する（初回のみ）。既にある場合は、旧バージョンの
 * レイアウト（固定タブ名指定・「案件リンク一覧」シート分離）が残っていれば
 * 自動移行し、不足している見出しラベルだけを補完する（ユーザー入力済みの値は
 * 上書きしない）。
 */
function getOrCreateConfigSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  var isNew = !sheet;
  if (isNew) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  } else {
    migrateAwayFromTabNameLayout_(sheet);
  }

  fillConfigSheetLabels_(sheet);
  migrateLegacyLinksSheet_(ss, sheet);

  if (isNew) {
    SpreadsheetApp.getUi().alert(
      '「' + CONFIG_SHEET_NAME + '」シートを作成しました。' + CONFIG_LINKS_HEADER_ROW +
      '行目以降に案件ファイルのURLを入力してから再度実行してください。'
    );
    return null;
  }
  return sheet;
}

/**
 * 「設定」シートの見出しラベル（A1〜A2・A4・B4）とB2の初期メモ値を補完する。
 * 既に値が入っているセル（B1・B2・URL一覧など）は上書きしない。
 */
function fillConfigSheetLabels_(sheet) {
  setIfEmpty_(sheet.getRange('A1'), 'ラベル出力先スプレッドシートID（自動設定・空欄でOK）');
  setIfEmpty_(sheet.getRange('A2'), '印刷時の余白（メモ・スクリプトでは使用しません）');
  setIfEmpty_(sheet.getRange(CONFIG_PRINT_MARGIN_CELL), CONFIG_PRINT_MARGIN_DEFAULT);
  setIfEmpty_(sheet.getRange(CONFIG_LINKS_HEADER_ROW, 1), '案件ファイルのリンク（1行に1件、URLを貼り付け）');
  setIfEmpty_(sheet.getRange(CONFIG_LINKS_HEADER_ROW, 2), 'ファイル名');
  sheet.getRange('A1:A2').setFontWeight('bold');
  sheet.getRange(CONFIG_LINKS_HEADER_ROW, 1, 1, 2).setFontWeight('bold');
}

function setIfEmpty_(range, value) {
  if (range.getValue() === '') {
    range.setValue(value);
  }
}

/**
 * 旧バージョンでは1行目が「読み込むタブ名」の指定だったが、現在は全タブを
 * 読み込むためこの指定は不要になった。1行目がその名残であれば行ごと削除し、
 * 2行目以降（ラベルSS ID・印刷余白メモ・URL一覧）を1行分繰り上げる。
 */
function migrateAwayFromTabNameLayout_(sheet) {
  if (String(sheet.getRange('A1').getValue()).trim() === LEGACY_TAB_NAME_LABEL) {
    sheet.deleteRow(1);
  }
}

/**
 * 旧バージョンが使っていた「案件リンク一覧」シートが残っている場合、
 * そのURL・ファイル名（2行目以降）を「設定」シートのCONFIG_LINKS_DATA_START_ROW
 * 行目以降へ一度だけ移行してから、旧シートを削除する。移行先に既にデータがある
 * 場合（移行済みだが旧シートの削除だけ失敗した場合など）は、コピーを二重に
 * 行わず削除だけ行う。
 */
function migrateLegacyLinksSheet_(ss, configSheet) {
  var legacySheet = ss.getSheetByName(LEGACY_LINKS_SHEET_NAME);
  if (!legacySheet) {
    return;
  }
  var alreadyMigrated = configSheet.getRange(CONFIG_LINKS_DATA_START_ROW, 1).getValue() !== '';
  if (!alreadyMigrated) {
    var numRows = Math.max(legacySheet.getLastRow() - 1, 0);
    if (numRows > 0) {
      var legacyValues = legacySheet.getRange(2, 1, numRows, 2).getValues();
      configSheet.getRange(CONFIG_LINKS_DATA_START_ROW, 1, numRows, 2).setValues(legacyValues);
    }
  }
  ss.deleteSheet(legacySheet);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日次データ取込')
    .addItem('リンクから読込', 'importFromLinks')
    .addItem('ラベル作成', 'createLabels')
    .addToUi();
}

// ============================================================
// 2. リンクからの案件読み込み
// ============================================================

var CASE_MENU_NAME_FONT_SIZE = 13; // 案件タブの「メニュー名」列のデータ行に適用するフォントサイズ
var CASE_QTY_FONT_SIZE = 14;       // 案件タブの「数量」列のデータ行に適用するフォントサイズ

/**
 * メニュー「日次データ取込」→「リンクから読込」から呼び出されるメイン関数。
 * 「設定」シートのCONFIG_LINKS_DATA_START_ROW行目以降のA列に貼られたURLを
 * 1件ずつ開き、リンク先ファイルの名前（Driveのファイル名）を同じ行のB列に
 * 表示するとともに、そのファイル名から読み取った「案件実施日 + 企業名」で
 * 名前をつけた案件タブを、リンク先ファイルの**全タブ**についてそれぞれ作成し、
 * 内容をそのままコピーする（1ファイルに複数タブある場合はタブ名を付記して
 * 区別する）。実行のたびに前回までの案件タブは全て削除してから作り直す
 * （deleteAllCaseSheets_。案件タブが際限なく増え続けるのを防ぐ）。
 */
function importFromLinks() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = getOrCreateConfigSheet_(ss);
  if (!configSheet) {
    return;
  }

  var numRows = Math.max(configSheet.getLastRow() - CONFIG_LINKS_DATA_START_ROW + 1, 0);
  if (numRows === 0) {
    ui.alert('「' + CONFIG_SHEET_NAME + '」シートの' + CONFIG_LINKS_DATA_START_ROW +
      '行目以降に案件ファイルのURLを入力してください。');
    return;
  }

  var urls = configSheet.getRange(CONFIG_LINKS_DATA_START_ROW, 1, numRows, 1).getValues()
    .map(function (row) { return String(row[0] || '').trim(); });

  // 実行のたびに前回までの案件タブをすべて削除してから作り直す
  deleteAllCaseSheets_(ss);

  var createdDisplayNames = [];
  var allWarnings = [];
  urls.forEach(function (url, i) {
    var rowNum = CONFIG_LINKS_DATA_START_ROW + i;
    var fileNameCell = configSheet.getRange(rowNum, 2);
    if (!url) {
      fileNameCell.setValue('');
      return;
    }

    var displayLabel = url;
    try {
      var sourceSs = resolveLinkedSpreadsheet_(url);
      displayLabel = sourceSs.getName();
      var parsed = extractDateAndCompanyFromFileName_(displayLabel);
      fileNameCell.setValue(displayLabel);

      var rawBaseName = (parsed.date ? parsed.date + ' ' : '') + parsed.company;
      var sourceSheets = sourceSs.getSheets();
      var multipleTabs = sourceSheets.length > 1;

      // リンク先ファイルの全タブを、それぞれ別の案件タブとしてコピーする
      sourceSheets.forEach(function (sourceSheet) {
        try {
          var values = sourceSheet.getDataRange().getValues();
          if (values.length === 0) {
            allWarnings.push(displayLabel + '「' + sourceSheet.getName() + '」 : データがありません');
            return;
          }

          // 1ファイルに複数タブある場合は元のタブ名を付記して区別する
          var rawName = multipleTabs ? rawBaseName + '（' + sourceSheet.getName() + '）' : rawBaseName;
          var newSheetName = uniqueSheetName_(ss, sanitizeSheetName_(rawName));

          // 元シートの書式（フォント・背景色・罫線・セル結合・列幅など）を保つため
          // getValues()+setValues() ではなく copyTo() でシートごと複製する。
          var newSheet = sourceSheet.copyTo(ss);
          newSheet.setName(newSheetName);
          copyEvaluatedFormulaCells_(sourceSheet, newSheet, values);

          applyCaseMenuFontSize_(newSheet, values);
          createdDisplayNames.push(newSheetName);
        } catch (eTab) {
          allWarnings.push(displayLabel + '「' + sourceSheet.getName() + '」 : 処理中にエラーが発生しました（' + eTab.message + '）');
        }
      });
    } catch (e) {
      allWarnings.push(displayLabel + ' : 処理中にエラーが発生しました（' + e.message + '）');
      fileNameCell.setValue('');
    }
  });

  ui.alert('リンク読込 結果', createdDisplayNames.length + '件の案件を読み込みました。\n' +
    createdDisplayNames.join('\n'), ui.ButtonSet.OK);

  // 警告・エラーは通常の結果に埋もれて見落とされないよう、別ダイアログで目立たせて表示する
  if (allWarnings.length > 0) {
    ui.alert('⚠️要確認', allWarnings.join('\n'), ui.ButtonSet.OK);
  }
}

/**
 * copyTo()でコピーされたシートのうち、(a) 数式が入っていたセルと、(b) コピー後に
 * 空欄になってしまっているのに元シートでは値があったセルを、評価済みの値
 * （values）で上書きする。
 *
 * (a) は、copyTo()が数式（他シート参照など）もそのままコピーしてしまい、
 * コピー先で参照が切れて #REF! 等のエラーになることを防ぐため。
 * (b) は、IMPORTRANGE・QUERY・ARRAYFORMULAなど1つのセルの数式が周囲の
 * セルへ結果を展開する（スピルする）形式の場合、Apps Scriptの
 * getFormulas()では展開先のセルは「数式なし」として扱われ、copyTo()でも
 * 展開先セル自体には値が複製されない（数式の実体は先頭セルにしかないため）。
 * その結果、(a)の条件だけでは展開範囲のほとんどが空欄のままコピーされてしまう。
 * 元シートでは値があった（＝スピル結果を含む）のにコピー後は空欄、という
 * セルだけを補完することで、この抜け漏れを埋める。
 *
 * それ以外の、数式もなくコピー後も空欄でないセル（時刻・日付など通常の入力値）
 * には触れない。全セルを評価済みの値で上書きすると、Apps Scriptが値をいったん
 * Dateオブジェクトに変換してから書き込み先のタイムゾーンで解釈し直すため、
 * コピー元とコピー先のタイムゾーン設定が異なる場合に時刻がズレてしまう。
 * 触る必要のないセルには触れないことでこのズレを避ける。
 */
function copyEvaluatedFormulaCells_(sourceSheet, targetSheet, values) {
  var numRows = values.length;
  var numCols = values[0].length;
  var formulas = sourceSheet.getRange(1, 1, numRows, numCols).getFormulas();
  var copiedValues = targetSheet.getRange(1, 1, numRows, numCols).getValues();
  for (var r = 0; r < numRows; r++) {
    for (var c = 0; c < numCols; c++) {
      var hasFormula = !!formulas[r][c];
      var copiedIsBlank = copiedValues[r][c] === '' || copiedValues[r][c] == null;
      var sourceHasValue = values[r][c] !== '' && values[r][c] != null;
      if (hasFormula || (copiedIsBlank && sourceHasValue)) {
        targetSheet.getRange(r + 1, c + 1).setValue(values[r][c]);
      }
    }
  }
}

/**
 * ファイル名から日付部分を抜き出し、日付テキストと残りの文字列（企業名。前後の
 * 【】/[]は取り除く）を返す。対応する日付表記は以下の2パターン。
 *   - 区切りあり: 区切り文字が `.` `/` `-` のいずれか、月日はゼロ埋めあり/なしの両方
 *     （「2026.8.1」「2026/08/01」「2026-8-01」等）
 *   - 区切りなし: `yyyyMMdd` の8桁連結表記で月日は常に2桁ゼロ埋め固定（「20260801」）
 * 日付が見つからない場合、dateは空文字になりcompanyはファイル名そのまま（記号除去のみ）。
 */
function extractDateAndCompanyFromFileName_(fileName) {
  var m = fileName.match(/(^|\D)(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})(\D|$)/) ||
    fileName.match(/(^|\D)(\d{4})(\d{2})(\d{2})(\D|$)/);
  if (!m) {
    return { date: '', company: cleanCompanyText_(fileName) };
  }
  var dateText = m[2] + '.' + Number(m[3]) + '.' + Number(m[4]);
  var remaining = fileName.substring(0, m.index) + m[1] + m[5] +
    fileName.substring(m.index + m[0].length);
  return { date: dateText, company: cleanCompanyText_(remaining) };
}

/** ファイル名から前後の【】/[]などの記号を取り除く。 */
function cleanCompanyText_(text) {
  var original = String(text).trim();
  var cleaned = original.replace(/^[【\[]\s*/, '').replace(/[】\]]\s*/, ' ').trim();
  return cleaned || original;
}

/**
 * 「設定」シートを除く全ての案件タブを削除する。実行のたびに前回までの案件タブを
 * 一掃してから作り直すことで、タブが際限なく増え続けるのを防ぐ。
 */
function deleteAllCaseSheets_(ss) {
  deleteSheetsWhere_(ss, function (sheet) {
    return sheet.getName() !== CONFIG_SHEET_NAME;
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
  setColumnFontSize_(sheet, menuHeader.colsByLabel['メニュー名'], dataStartRow1, numDataRows, CASE_MENU_NAME_FONT_SIZE);
  setColumnFontSize_(sheet, menuHeader.colsByLabel['数量'], dataStartRow1, numDataRows, CASE_QTY_FONT_SIZE);
}

function setColumnFontSize_(sheet, colIndex, startRow1, numRows, fontSize) {
  if (colIndex == null) {
    return;
  }
  sheet.getRange(startRow1, colIndex + 1, numRows, 1).setFontSize(fontSize);
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
// 内訳は34px/48px/46px（メニュー名の2行折り返しのため、数量段からさらに2px
// 減らしてメニュー名段に足した。合計128pxは変わらないためラベル1枚の実寸33.9mmは維持される）
var LABEL_SUBROW_HEIGHTS_PX = [34, 48, 46];
var LABEL_SUBROWS = LABEL_SUBROW_HEIGHTS_PX.length;

var LABEL_FONT_SIZE = 14;      // 1段目（日付＋企業名）のフォントサイズ
var LABEL_MENU_FONT_SIZE = 12; // 2段目（メニュー名）。1段目より2pt小さい
var LABEL_QTY_FONT_SIZE = 28;  // 3段目（数量）。太字・大きめフォントで強調する
var LABEL_MENU_MAX_ZENKAKU_LEN = 30; // 2段目（メニュー名）は全角換算でこの文字数を超えたら切り捨てる

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
  deleteSheetsWhere_(labelSs, function (sheet) {
    return currentCaseNames.indexOf(sheet.getName()) === -1;
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
 * 既に同名タブがあれば中身だけ消して再利用するため、再実行しても古い内容は
 * 残らない（タブ自体を削除しないのは、印刷余白などタブに紐づく設定を
 * 失わないようにするため）。
 */
function writeLabelSheetForCase_(labelSs, caseName, entries) {
  // 削除して作り直すと、印刷余白などタブに紐づく設定が失われる可能性があるため、
  // 既存タブがあれば中身だけ消して（clear）再利用する。
  var sheet = labelSs.getSheetByName(caseName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = labelSs.insertSheet(caseName);
  }

  var entriesPerColumn = Math.floor(LABEL_ROWS / LABEL_COPIES_PER_ENTRY);
  var entriesPerPage = entriesPerColumn * LABEL_COLS;
  var totalPages = Math.ceil(entries.length / entriesPerPage);
  var totalRows = totalPages * LABEL_ROWS * LABEL_SUBROWS;

  // 最終ページで3列目が1件も埋まらないと、その列が空のままになり印刷範囲が
  // 2列分に縮んで中央寄せがずれてしまう。先に全セルへ空文字列の値を入れておくことに加え、
  // 薄い罫線も引いておくことで、3列とも「データがある列」として印刷範囲に確実に
  // 含まれるようにする（実データは後段の書き込みで上書きされる）
  sheet.getRange(1, 1, totalRows, LABEL_COLS)
    .setValue('')
    .setFontSize(LABEL_FONT_SIZE)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setBorder(true, true, true, true, true, true, '#f2f2f2', SpreadsheetApp.BorderStyle.SOLID);

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
 *   1段目: 日付＋企業名（左寄せ、WrapStrategy.CLIP＝折り返さず高さ固定）
 *   2段目: メニュー名（中央寄せ・1段目より2pt小さいフォント。
 *          WrapStrategy.WRAP＝2行まで折り返す。全角30文字
 *          （LABEL_MENU_MAX_ZENKAKU_LEN）を超える分は事前に切り捨てているため、
 *          2行に収まりきらず段の高さが崩れることを防いでいる）
 *   3段目: 数量（中央寄せ・太字・大きめフォントで強調、WrapStrategy.CLIP）
 */
function writeLabelCellGroup_(sheet, rowBase, col1, entry) {
  sheet.getRange(rowBase + 1, col1).setValue(entry.date + ' ' + entry.company)
    .setFontSize(LABEL_FONT_SIZE)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('bottom')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  sheet.getRange(rowBase + 2, col1).setValue(truncateByZenkakuWidth_(entry.menu, LABEL_MENU_MAX_ZENKAKU_LEN))
    .setFontSize(LABEL_MENU_FONT_SIZE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  sheet.getRange(rowBase + 3, col1).setValue(entry.qty)
    .setFontSize(LABEL_QTY_FONT_SIZE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

/**
 * 文字列を全角換算で指定した文字数までに切り詰める（半角文字は0.5文字分として
 * カウントする）。超えた分は末尾を単純に切り捨てる（省略記号は付けない）。
 */
function truncateByZenkakuWidth_(text, maxZenkakuWidth) {
  var result = '';
  var width = 0;
  for (var i = 0; i < text.length; i++) {
    var ch = text.charAt(i);
    var charWidth = isHalfWidthChar_(ch) ? 0.5 : 1;
    if (width + charWidth > maxZenkakuWidth) {
      break;
    }
    result += ch;
    width += charWidth;
  }
  return result;
}

/** 半角英数・記号（U+0000〜U+00FF）、半角カタカナ（U+FF61〜U+FF9F）を半角とみなす。 */
function isHalfWidthChar_(ch) {
  var code = ch.charCodeAt(0);
  return (code >= 0x0000 && code <= 0x00FF) || (code >= 0xFF61 && code <= 0xFF9F);
}

// ============================================================
// 4. 共通ヘルパー
// ============================================================

/**
 * spreadsheet内のシートのうち、shouldDelete(sheet)がtrueを返すものを削除する。
 * ただし最低1枚はシートを残す（スプレッドシートは全シート削除できないため）。
 */
function deleteSheetsWhere_(spreadsheet, shouldDelete) {
  var sheets = spreadsheet.getSheets();
  var remaining = sheets.length;
  sheets.forEach(function (sheet) {
    if (shouldDelete(sheet) && remaining > 1) {
      spreadsheet.deleteSheet(sheet);
      remaining--;
    }
  });
}

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

/**
 * スプレッドシートのURLまたは素のIDから、スプレッドシートIDを取り出す。
 * `.../spreadsheets/d/<ID>/edit`・`.../file/d/<ID>/view`・
 * `drive.google.com/open?id=<ID>` のいずれの形式にも対応する。
 */
function extractSpreadsheetId_(input) {
  var text = String(input).trim();
  var match = text.match(/\/(?:spreadsheets|file)\/d\/([a-zA-Z0-9_-]+)/) ||
    text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match) {
    return match[1];
  }
  return text;
}

/**
 * 「設定」シートのA列（案件ファイルのリンク一覧）の値からスプレッドシートを
 * 取得する。URL・IDに加えて、
 * リンクの代わりにファイル名がそのまま貼られてしまった場合の救済策として、
 * Drive内をそのファイル名で検索して開くこともできる。
 */
function resolveLinkedSpreadsheet_(input) {
  var id = extractSpreadsheetId_(input);
  if (/^[a-zA-Z0-9_-]{20,}$/.test(id)) {
    return SpreadsheetApp.openById(id);
  }

  // URLの形式にもIDの形式にも一致しなかった場合、そのままファイル名として扱う
  var iterator = DriveApp.getFilesByName(id);
  while (iterator.hasNext()) {
    var file = iterator.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      return SpreadsheetApp.open(file);
    }
  }
  throw new Error('URL・IDとして認識できず、ファイル名でも見つかりませんでした');
}
