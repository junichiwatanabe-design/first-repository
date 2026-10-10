/**
 * 「設定」シートに手入力した案件ファイルのURLを1件ずつ開き、ファイル名から
 * 案件実施日・企業名を読み取るとともに、リンク先ファイル内の指定タブから
 * 日付・企業名・メニュー名・数量を直接読み取って、印刷用ラベル（A-oneラベル
 * シール、設定で3×8/3×6を選択）のHTMLを生成し、ブラウザの印刷機能で
 * そのまま印刷できるダイアログを表示する。元の案件ファイルをコピーする工程も、
 * スプレッドシートへラベルを書き込む工程も持たないため、スプレッドシートは
 * これ1枚だけで運用できる。
 *
 * ファイル構成:
 *   1. 設定シート関連
 *   2. リンクからのラベル作成（HTML生成・印刷ダイアログ表示）
 *   3. 共通ヘルパー
 */

// ============================================================
// 1. 設定シート関連
// ============================================================

var CONFIG_SHEET_NAME = '設定';
var CONFIG_TAB_NAME_CELL = 'B1';          // 各ファイル内で読み込む固定タブ名
var CONFIG_PRINT_MARGIN_CELL = 'B2';      // 印刷時余白のメモ（スクリプトからは読み書きしない）
var CONFIG_PRINT_MARGIN_DEFAULT = '上下12mm・左右5mm程度';
var CONFIG_LABEL_SIZE_CELL = 'B3';        // ラベルシールのサイズ（プルダウンで3×8/3×6を選択）
var CONFIG_LINKS_HEADER_ROW = 4;          // 案件ファイルのリンク一覧の見出し行
var CONFIG_LINKS_DATA_START_ROW = 5;      // 案件ファイルのリンク一覧のデータ開始行
var LEGACY_LINKS_SHEET_NAME = '案件リンク一覧'; // 旧バージョンで使っていたリンク一覧シート名（あれば「設定」へ自動移行して削除する）
var TAB_NAME_LABEL = '読み込むタブ名';
var LEGACY_LABEL_SS_LABEL = 'ラベル出力先スプレッドシートID（自動設定・空欄でOK）';

/**
 * 「設定」シートを取得する。無ければ新規作成し、見出しラベルを設定したうえで
 * アラートを出して処理を中断する（初回のみ）。既にある場合は、1行目に
 * タブ名の行が無い旧レイアウト、または2行目がラベル出力先スプレッドシートID
 * の名残である旧レイアウトであれば自動移行し、不足している見出しラベル
 * だけを補完する（ユーザー入力済みの値は上書きしない）。
 */
function getOrCreateConfigSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  var isNew = !sheet;
  if (isNew) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  } else {
    migrateBackToTabNameLayout_(sheet);
    migrateAwayFromLabelSsLayout_(sheet);
  }

  fillConfigSheetLabels_(sheet);
  migrateLegacyLinksSheet_(ss, sheet);

  if (isNew) {
    SpreadsheetApp.getUi().alert(
      '「' + CONFIG_SHEET_NAME + '」シートを作成しました。' +
      CONFIG_TAB_NAME_CELL + ' に読み込むタブ名を、' + CONFIG_LINKS_HEADER_ROW +
      '行目以降に案件ファイルのURLを入力してから再度実行してください。'
    );
    return null;
  }
  return sheet;
}

/**
 * 「設定」シートの見出しラベル（A1〜A3・A4）とB2・B3の初期値を補完する。
 * 既に値が入っているセル（B1・B2・B3・URL一覧など）は上書きしない。
 */
function fillConfigSheetLabels_(sheet) {
  setIfEmpty_(sheet.getRange('A1'), TAB_NAME_LABEL);
  setIfEmpty_(sheet.getRange('A2'), '印刷時の余白（メモ・スクリプトでは使用しません）');
  setIfEmpty_(sheet.getRange(CONFIG_PRINT_MARGIN_CELL), CONFIG_PRINT_MARGIN_DEFAULT);
  setIfEmpty_(sheet.getRange('A3'), 'ラベルシールのサイズ（プルダウンで選択）');
  setIfEmpty_(sheet.getRange(CONFIG_LABEL_SIZE_CELL), LABEL_SIZE_3X8);
  sheet.getRange(CONFIG_LABEL_SIZE_CELL).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList([LABEL_SIZE_3X8, LABEL_SIZE_3X6], true)
      .setAllowInvalid(false)
      .build()
  );
  setIfEmpty_(sheet.getRange(CONFIG_LINKS_HEADER_ROW, 1), '案件ファイルのリンク（1行に1件、URLを貼り付け）');
  sheet.getRange('A1:A3').setFontWeight('bold');
  sheet.getRange(CONFIG_LINKS_HEADER_ROW, 1).setFontWeight('bold');
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
function migrateBackToTabNameLayout_(sheet) {
  if (String(sheet.getRange('A1').getValue()).trim() !== TAB_NAME_LABEL) {
    sheet.insertRowBefore(1);
  }
}

/**
 * 旧バージョン（単一スプレッドシート化する前）では2行目が「ラベル出力先
 * スプレッドシートID」の指定だったが、ラベル専用の別ファイルを作らなくなった
 * ため不要になった。2行目がその名残であれば行ごと削除し、3行目以降
 * （印刷余白メモ・ラベルサイズ・URL一覧）を1行分繰り上げる。
 */
function migrateAwayFromLabelSsLayout_(sheet) {
  if (String(sheet.getRange('A2').getValue()).trim() === LEGACY_LABEL_SS_LABEL) {
    sheet.deleteRow(2);
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
    .createMenu('ラベル作成')
    .addItem('作成', 'createLabels')
    .addToUi();
}

// ============================================================
// 2. リンクからのラベル作成
// ============================================================

// 面付け（A-one マルチプリンタ用ラベルシール。「設定」シートのプルダウンで
// 3×8（24面・66mm×33.9mm）と3×6（18面相当・66mm×45.2mm）を切り替えられる。
// 総高さ271.2mmは共通で、8行と6行のどちらで均等に割るかだけが違う）
var LABEL_COLS = 3;
var LABEL_COPIES_PER_ENTRY = 2; // 同一ラベルを縦に2枚配置
var LABEL_WIDTH_MM = 66;        // ラベル1枚の実寸幅。商品が異なる場合は要調整

var LABEL_SIZE_3X8 = '3×8（66mm×33.9mm）';
var LABEL_SIZE_3X6 = '3×6（66mm×45.2mm）';

// 内訳（3段の高さmm）は、3×8は9.0/12.7/12.2（メニュー名の2行分を広めに確保し、
// 合計33.9mmに揃えた）、3×6は45.2mmを同じ比率で配分した11.9/16.9/16.4を
// 初期値とする。
var LABEL_SIZE_PRESETS = {};
LABEL_SIZE_PRESETS[LABEL_SIZE_3X8] = { rows: 8, labelHeightMm: 33.9, segmentHeightsMm: [9.0, 12.7, 12.2] };
LABEL_SIZE_PRESETS[LABEL_SIZE_3X6] = { rows: 6, labelHeightMm: 45.2, segmentHeightsMm: [11.9, 16.9, 16.4] };

/**
 * 「設定」シートB3で選択されたラベルサイズのプリセット（行数・ラベル高さmm・
 * 3段の高さmm）を返す。未選択・不正な値の場合は3×8（従来のデフォルト）を返す。
 */
function getLabelSizePreset_(configSheet) {
  var raw = String(configSheet.getRange(CONFIG_LABEL_SIZE_CELL).getValue() || '').trim();
  return LABEL_SIZE_PRESETS[raw] || LABEL_SIZE_PRESETS[LABEL_SIZE_3X8];
}

var LABEL_FONT_SIZE = 14;      // 1段目（日付＋企業名）のフォントサイズ(pt)
var LABEL_MENU_FONT_SIZE = 12; // 2段目（メニュー名）。1段目より2pt小さい
var LABEL_QTY_FONT_SIZE = 25;  // 3段目（数量）。太字・大きめフォントで強調する

/**
 * メニュー「ラベル作成」→「作成」から呼び出されるメイン関数。
 * 「設定」シートのURL一覧を1件ずつ開き、リンク先ファイル内の指定タブから
 * 直接（コピーせずに）日付・企業名・メニュー名・数量を読み取り、案件ごとの
 * 結果をまとめてHTMLの印刷用ダイアログを表示する（スプレッドシートへの
 * 書き込みは行わない）。
 */
function createLabels() {
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

  var numRows = Math.max(configSheet.getLastRow() - CONFIG_LINKS_DATA_START_ROW + 1, 0);
  if (numRows === 0) {
    ui.alert('「' + CONFIG_SHEET_NAME + '」シートの' + CONFIG_LINKS_DATA_START_ROW +
      '行目以降に案件ファイルのURLを入力してください。');
    return;
  }

  var urls = configSheet.getRange(CONFIG_LINKS_DATA_START_ROW, 1, numRows, 1).getValues()
    .map(function (row) { return String(row[0] || '').trim(); });

  var timeZone = ss.getSpreadsheetTimeZone();
  var sizePreset = getLabelSizePreset_(configSheet);

  var caseResults = [];
  var totalEntries = 0;
  var warnings = [];
  urls.forEach(function (url) {
    if (!url) {
      return;
    }

    var displayLabel = url;
    try {
      var sourceSs = resolveLinkedSpreadsheet_(url);
      displayLabel = sourceSs.getName();
      var parsed = extractDateAndCompanyFromFileName_(displayLabel);

      var sourceSheet = sourceSs.getSheetByName(tabName);
      if (!sourceSheet) {
        warnings.push(displayLabel + ' : タブ「' + tabName + '」が見つかりません');
        return;
      }

      var entries = collectLabelEntries_(sourceSheet, timeZone);
      if (entries.length === 0) {
        warnings.push(displayLabel + ' : ラベルに出力できるデータがありません');
        return;
      }

      var caseLabel = ((parsed.date ? parsed.date + ' ' : '') + parsed.company).trim();
      caseResults.push({ caseLabel: caseLabel, entries: entries });
      totalEntries += entries.length;
    } catch (e) {
      warnings.push(displayLabel + ' : 処理中にエラーが発生しました（' + e.message + '）');
    }
  });

  // 警告・エラーは印刷ダイアログに埋もれて見落とされないよう、先に表示する
  if (warnings.length > 0) {
    ui.alert('⚠️要確認', warnings.join('\n'), ui.ButtonSet.OK);
  }

  if (caseResults.length === 0) {
    ui.alert('ラベルに出力できる案件データが見つかりませんでした。');
    return;
  }

  var html = buildLabelsHtml_(caseResults, sizePreset);
  var output = HtmlService.createHtmlOutput(html).setWidth(850).setHeight(650);
  ui.showModalDialog(output, 'ラベル印刷（' + caseResults.length + '件の案件・計' + totalEntries + '件のメニュー）');
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
 * リンク先ファイル内の指定タブ1枚分から、ラベルに出力するエントリ
 * （日付・企業名・メニュー名・数量）を集める。メニュー名・数量のどちらかが
 * 空の行はスキップする。
 */
function collectLabelEntries_(sheet, timeZone) {
  var values = sheet.getDataRange().getValues();
  if (values.length === 0) {
    return [];
  }

  var dateText = formatLabelDate_(findAdjacentValue_(values, '案件実施日'), timeZone);
  var company = normalizeLabelText_(findAdjacentValue_(values, '企業名'));

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
    var menuName = normalizeLabelText_(values[r][menuCol]);
    var qty = values[r][qtyCol];
    if (!menuName || !qty) {
      continue;
    }
    entries.push({ date: dateText, company: company, menu: menuName, qty: qty });
  }
  return entries;
}

/**
 * セルの値に含まれる改行文字を半角スペースに置き換えてから前後の空白を除く。
 * WrapStrategy.CLIPでも、セルの値自体に改行が含まれていればそのまま改行
 * として表示されてしまうため、ラベルに使う文字列から事前に取り除く。
 */
function normalizeLabelText_(value) {
  return String(value || '').replace(/[\r\n]+/g, ' ').trim();
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
 * 「7/16」「2026.7.16」「2026/7/17」「2026年7月16日（木）」など表記が揺れた
 * 月日の文字列を、ゼロ埋めなし・年なしの「7/16」形式に統一する。
 * 曜日の注記「（木）」などが末尾に付いていても、数字部分だけを拾うため
 * 影響しない。月日の形式として解釈できない場合は元の文字列をそのまま返す。
 */
function normalizeDateText_(text) {
  // 「2026年7月16日（木）」のような漢字区切り表記
  var kanji = text.match(/(\d{1,2})月(\d{1,2})日/);
  if (kanji) {
    return Number(kanji[1]) + '/' + Number(kanji[2]);
  }
  // 「2026.7.16（木）」「2026/7/16」のように年＋月＋日（区切りは`.`または`/`）
  var withYear = text.match(/\d{4}[./](\d{1,2})[./](\d{1,2})/);
  if (withYear) {
    return Number(withYear[1]) + '/' + Number(withYear[2]);
  }
  // 「7/16」「7.16」のように月日のみ（年なし）
  var withoutYear = text.match(/(\d{1,2})[./](\d{1,2})/);
  if (withoutYear) {
    return Number(withoutYear[1]) + '/' + Number(withoutYear[2]);
  }
  return text;
}

/** HTML文字列に埋め込むテキストをエスケープする（& < > " ' の5文字）。 */
function escapeHtml_(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * ラベル印刷ダイアログの<style>ブロックを返す。グリッドの行列数をラベル
 * サイズプリセット通りに固定し、`grid-auto-flow: column`でセルを出現順に
 * 流し込むだけで「1列を上から下まで埋めてから次の列へ」という面付け順を
 * 再現する（Sheets時代の空セル事前書き込みのような回避策が不要になる）。
 */
function buildLabelStyle_(sizePreset) {
  var segmentHeights = sizePreset.segmentHeightsMm;
  return '' +
    '<style>' +
    'body { margin: 0; font-family: sans-serif; }' +
    '.toolbar { padding: 12px 16px; background: #f2f2f2; font-size: 14px; }' +
    '.toolbar button { font-size: 14px; padding: 6px 16px; margin-right: 8px; }' +
    '.case-title { margin: 16px 16px 0; font-size: 14px; color: #555; }' +
    '.label-page {' +
    '  display: grid;' +
    '  grid-template-columns: repeat(' + LABEL_COLS + ', ' + LABEL_WIDTH_MM + 'mm);' +
    '  grid-template-rows: repeat(' + sizePreset.rows + ', ' + sizePreset.labelHeightMm + 'mm);' +
    '  grid-auto-flow: column;' +
    '  margin: 8px;' +
    '}' +
    '.label-page.page-break { break-after: page; page-break-after: always; }' +
    '.label-cell { display: flex; flex-direction: column; border: 1px solid #f2f2f2; box-sizing: border-box; }' +
    '.label-seg1 {' +
    '  height: ' + segmentHeights[0] + 'mm;' +
    '  font-size: ' + LABEL_FONT_SIZE + 'pt;' +
    '  font-weight: bold;' +
    '  white-space: nowrap;' +
    '  overflow: hidden;' +
    '  display: flex;' +
    '  align-items: flex-end;' +
    '  padding: 0 2px;' +
    '  box-sizing: border-box;' +
    '}' +
    '.label-seg2 {' +
    '  height: ' + segmentHeights[1] + 'mm;' +
    '  font-size: ' + LABEL_MENU_FONT_SIZE + 'pt;' +
    '  font-weight: bold;' +
    '  text-align: center;' +
    '  display: -webkit-box;' +
    '  -webkit-box-orient: vertical;' +
    '  -webkit-box-pack: center;' +
    '  -webkit-line-clamp: 2;' +
    '  overflow: hidden;' +
    '  padding: 0 2px;' +
    '  box-sizing: border-box;' +
    '}' +
    '.label-seg3 {' +
    '  height: ' + segmentHeights[2] + 'mm;' +
    '  font-size: ' + LABEL_QTY_FONT_SIZE + 'pt;' +
    '  font-weight: bold;' +
    '  display: flex;' +
    '  align-items: center;' +
    '  justify-content: center;' +
    '  box-sizing: border-box;' +
    '}' +
    '@media print { .no-print { display: none; } }' +
    '@page { margin: 0; }' +
    '</style>';
}

/**
 * ラベル1件分（3段）のHTMLを返す。
 *   1段目: 日付＋企業名＋「様」（左寄せ・下揃え、nowrap+overflow:hiddenで
 *          折り返さず末尾を切る。旧WrapStrategy.CLIPの挙動を再現する）
 *   2段目: メニュー名（中央寄せ、-webkit-line-clamp:2で2行まで表示し、
 *          超えた分は省略記号「…」で省略する）
 *   3段目: 数量（中央寄せ・太字・大きめフォントで強調）
 */
function buildLabelCellHtml_(entry) {
  var line1 = escapeHtml_(entry.date + ' ' + entry.company + ' 様');
  var line2 = escapeHtml_(entry.menu);
  var line3 = escapeHtml_(String(entry.qty));
  return '' +
    '<div class="label-cell">' +
    '<div class="label-seg1">' + line1 + '</div>' +
    '<div class="label-seg2">' + line2 + '</div>' +
    '<div class="label-seg3">' + line3 + '</div>' +
    '</div>';
}

/**
 * 1案件分のentriesを、1ページ分（LABEL_COLS × sizePreset.rows）ごとの
 * ページHTML文字列（`.label-page`の<div>）の配列に分割する。1エントリにつき
 * LABEL_COPIES_PER_ENTRY枚（同一内容を複製）を連続して配置し、
 * `grid-auto-flow: column`によって1列を上から埋めてから次の列へという順で
 * 面付けされる。
 */
function buildCasePages_(entries, sizePreset) {
  var entriesPerColumn = Math.floor(sizePreset.rows / LABEL_COPIES_PER_ENTRY);
  var entriesPerPage = entriesPerColumn * LABEL_COLS;
  var pages = [];
  for (var i = 0; i < entries.length; i += entriesPerPage) {
    var pageEntries = entries.slice(i, i + entriesPerPage);
    var cellsHtml = [];
    pageEntries.forEach(function (entry) {
      for (var copy = 0; copy < LABEL_COPIES_PER_ENTRY; copy++) {
        cellsHtml.push(buildLabelCellHtml_(entry));
      }
    });
    pages.push('<div class="label-page">' + cellsHtml.join('') + '</div>');
  }
  return pages;
}

/**
 * 印刷用ダイアログに表示するHTML文書全体を返す。案件ごとに画面確認用の
 * 見出し（no-print・印刷時は非表示）と、buildCasePages_で得たページを出力し、
 * 文書全体で最後のページ以外には改ページを付与する（最後のページに付けると
 * 余分な空白ページが印刷されてしまうため）。
 */
function buildLabelsHtml_(caseResults, sizePreset) {
  var allPages = [];
  caseResults.forEach(function (caseResult) {
    allPages.push(
      '<div class="case-title no-print">' + escapeHtml_(caseResult.caseLabel) +
      '（' + caseResult.entries.length + '件）</div>'
    );
    buildCasePages_(caseResult.entries, sizePreset).forEach(function (pageHtml) {
      allPages.push(pageHtml);
    });
  });

  var lastPageIndex = -1;
  for (var i = allPages.length - 1; i >= 0; i--) {
    if (allPages[i].indexOf('class="label-page"') !== -1) {
      lastPageIndex = i;
      break;
    }
  }
  for (var j = 0; j < allPages.length; j++) {
    if (j !== lastPageIndex && allPages[j].indexOf('class="label-page"') !== -1) {
      allPages[j] = allPages[j].replace('class="label-page"', 'class="label-page page-break"');
    }
  }

  return '' +
    '<!doctype html><html><head><meta charset="utf-8">' +
    buildLabelStyle_(sizePreset) +
    '</head><body>' +
    '<div class="toolbar no-print">' +
    '<button onclick="window.print()">印刷</button>' +
    '<span>プレビューを確認してから印刷してください（用紙サイズ・余白はブラウザの印刷ダイアログで調整できます）。</span>' +
    '</div>' +
    allPages.join('') +
    '</body></html>';
}

// ============================================================
// 3. 共通ヘルパー
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
