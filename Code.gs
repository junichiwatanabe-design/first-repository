/**
 * 「案件リンク一覧」シートに手入力した案件ファイルのURLを1件ずつ開き、
 * ファイル名から案件実施日・企業名を読み取って一覧に表示するとともに、
 * 「設定」シートで指定したタブ名の内容を集計用スプレッドシートへ案件ごとに
 * 別タブとしてコピーする（印刷時に案件単位で改ページされるよう1案件＝1タブの
 * 構成にしている）。
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
var CONFIG_TAB_NAME_CELL = 'B1';  // 各ファイル内で読み込む固定タブ名
var CONFIG_LABEL_SS_CELL = 'B2';  // ラベル出力先スプレッドシートの URL/ID（自動設定）

function getOrCreateConfigSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (sheet) {
    return sheet;
  }

  sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  sheet.getRange('A1').setValue('読み込むタブ名');
  sheet.getRange('A2').setValue('ラベル出力先スプレッドシートID（自動設定・空欄でOK）');
  sheet.getRange('A1:A2').setFontWeight('bold');

  SpreadsheetApp.getUi().alert(
    '「' + CONFIG_SHEET_NAME + '」シートを作成しました。' +
    CONFIG_TAB_NAME_CELL + ' に読み込むタブ名を入力してから再度実行してください。'
  );
  return null;
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

var LINKS_SHEET_NAME = '案件リンク一覧';
var CASE_MENU_NAME_FONT_SIZE = 13; // 案件タブの「メニュー名」列のデータ行に適用するフォントサイズ
var CASE_QTY_FONT_SIZE = 14;       // 案件タブの「数量」列のデータ行に適用するフォントサイズ

/**
 * 「案件リンク一覧」シートを取得する。無ければ見出し付きで新規作成し、
 * アラートを出して処理を中断する（getOrCreateConfigSheet_と同じパターン）。
 * A列: URL（手入力）、B列: 案件実施日、C列: 企業名、D列: ファイル名
 * （B〜Dはどれもimportfromlinksが自動入力する。案件実施日・企業名はファイル名
 * から読み取ったもの、ファイル名はDriveのファイル名そのもの）。
 */
function getOrCreateLinksSheet_(ss) {
  var sheet = ss.getSheetByName(LINKS_SHEET_NAME);
  if (sheet) {
    return sheet;
  }

  sheet = ss.insertSheet(LINKS_SHEET_NAME);
  sheet.getRange('A1').setValue('案件ファイルのリンク（1行に1件、URLを貼り付け）');
  sheet.getRange('B1').setValue('案件実施日');
  sheet.getRange('C1').setValue('企業名');
  sheet.getRange('D1').setValue('ファイル名');
  sheet.getRange('A1:D1').setFontWeight('bold');

  SpreadsheetApp.getUi().alert(
    '「' + LINKS_SHEET_NAME + '」シートを作成しました。' +
    'A列2行目以降に案件ファイルのURLを1行に1件貼り付けてから再度実行してください。'
  );
  return null;
}

/**
 * メニュー「日次データ取込」→「リンクから読込」から呼び出されるメイン関数。
 * 「案件リンク一覧」のA列に貼られたURLを1件ずつ開き、リンク先ファイルの名前
 * （Driveのファイル名）を同じ行のD列に、そのファイル名から読み取った案件実施日・
 * 企業名をB・C列に表示するとともに、「案件実施日 + 企業名」で名前をつけた
 * 案件タブを作成し、「設定」シートB1で指定したタブ名の内容をそのままコピーする。
 * 実行のたびに前回までの案件タブは全て削除してから作り直す
 * （deleteAllCaseSheets_。案件タブが際限なく増え続けるのを防ぐ）。
 */
function importFromLinks() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = getOrCreateConfigSheet_(ss);
  if (!configSheet) {
    return;
  }
  var tabName = configSheet.getRange(CONFIG_TAB_NAME_CELL).getValue();
  if (!tabName) {
    ui.alert('設定シートの「' + CONFIG_TAB_NAME_CELL + '」(タブ名) を入力してください。');
    return;
  }

  var linksSheet = getOrCreateLinksSheet_(ss);
  if (!linksSheet) {
    return;
  }
  var numRows = Math.max(linksSheet.getLastRow() - 1, 0);
  if (numRows === 0) {
    ui.alert('「' + LINKS_SHEET_NAME + '」シートのA列2行目以降にURLを入力してください。');
    return;
  }

  var urls = linksSheet.getRange(2, 1, numRows, 1).getValues()
    .map(function (row) { return String(row[0] || '').trim(); });

  // 実行のたびに前回までの案件タブをすべて削除してから作り直す
  deleteAllCaseSheets_(ss);

  var createdDisplayNames = [];
  var allWarnings = [];
  urls.forEach(function (url, i) {
    var rowNum = i + 2;
    var infoRange = linksSheet.getRange(rowNum, 2, 1, 3); // B:D（案件実施日・企業名・ファイル名）
    if (!url) {
      infoRange.setValue('');
      return;
    }

    var displayLabel = url;
    try {
      var sourceSs = resolveLinkedSpreadsheet_(url);
      displayLabel = sourceSs.getName();
      var parsed = extractDateAndCompanyFromFileName_(displayLabel);

      // ファイル自体は開けたので、タブが見つからない場合でも診断用にファイル名は表示する
      linksSheet.getRange(rowNum, 4).setValue(displayLabel);

      var sourceSheet = sourceSs.getSheetByName(tabName);
      if (!sourceSheet) {
        allWarnings.push(displayLabel + ' : タブ「' + tabName + '」が見つかりません');
        linksSheet.getRange(rowNum, 2, 1, 2).setValue(''); // B:Cのみクリア
        return;
      }
      var values = sourceSheet.getDataRange().getValues();
      if (values.length === 0) {
        allWarnings.push(displayLabel + ' : データがありません');
        linksSheet.getRange(rowNum, 2, 1, 2).setValue(''); // B:Cのみクリア
        return;
      }

      // 「案件リンク一覧」の同じ行にファイル名由来の案件実施日・企業名を表示する
      linksSheet.getRange(rowNum, 2).setValue(parsed.date);
      linksSheet.getRange(rowNum, 3).setValue(parsed.company);

      // 案件タブを作成し、リンク先の内容をそのままコピーする
      var baseName = sanitizeSheetName_((parsed.date ? parsed.date + ' ' : '') + parsed.company);
      var newSheetName = uniqueSheetName_(ss, baseName);

      // 元シートの書式（フォント・背景色・罫線・セル結合・列幅など）を保つため
      // getValues()+setValues() ではなく copyTo() でシートごと複製する。
      var newSheet = sourceSheet.copyTo(ss);
      newSheet.setName(newSheetName);

      // copyTo() は数式（他シート参照など）もそのままコピーしてしまい、集計先で
      // 参照が切れて #REF! 等のエラーになることがある。コピー前に評価済みの値
      // （values）で上書きし、セルの中身を確定値にする（書式は変えない）。
      newSheet.getRange(1, 1, values.length, values[0].length).setValues(values);

      applyCaseMenuFontSize_(newSheet, values);
      createdDisplayNames.push(newSheetName);
    } catch (e) {
      allWarnings.push(displayLabel + ' : 処理中にエラーが発生しました（' + e.message + '）');
      infoRange.setValue('');
    }
  });

  ui.alert('リンク読込 結果', createdDisplayNames.length + '件の案件を読み込みました（' +
    createdDisplayNames.join(' / ') + '）', ui.ButtonSet.OK);

  // 警告・エラーは通常の結果に埋もれて見落とされないよう、別ダイアログで目立たせて表示する
  if (allWarnings.length > 0) {
    ui.alert('⚠️要確認', allWarnings.join('\n'), ui.ButtonSet.OK);
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
 * 「設定」「案件リンク一覧」シートを除く全ての案件タブを削除する。実行のたびに
 * 前回までの案件タブを一掃してから作り直すことで、タブが際限なく増え続けるのを防ぐ。
 */
function deleteAllCaseSheets_(ss) {
  var sheets = ss.getSheets();
  var remaining = sheets.length;
  sheets.forEach(function (sheet) {
    if (sheet.getName() === CONFIG_SHEET_NAME || sheet.getName() === LINKS_SHEET_NAME) {
      return;
    }
    if (remaining > 1) {
      ss.deleteSheet(sheet);
      remaining--;
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
    if (name === CONFIG_SHEET_NAME || name === LABEL_SHEET_NAME || name === LINKS_SHEET_NAME) {
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
  var remaining = sheets.length;
  sheets.forEach(function (sheet) {
    if (currentCaseNames.indexOf(sheet.getName()) === -1 && remaining > 1) {
      labelSs.deleteSheet(sheet);
      remaining--;
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

/** スプレッドシートのURLまたは素のIDから、スプレッドシートIDを取り出す。 */
function extractSpreadsheetId_(input) {
  var text = String(input).trim();
  var match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match) {
    return match[1];
  }
  return text;
}

/**
 * 「案件リンク一覧」のA列の値からスプレッドシートを取得する。URL・IDに加えて、
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
