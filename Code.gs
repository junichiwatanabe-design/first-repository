/**
 * 「設定」シートに手入力した案件ファイルのURLを1件ずつ開き、ファイル名から
 * 案件実施日・企業名を読み取るとともに、リンク先ファイル内の指定タブから
 * 日付・企業名・メニュー名・数量を直接読み取って、このスプレッドシート自身に
 * 案件ごとの印刷用ラベル（A-oneラベルシール、設定で3×8/3×6を選択）タブを
 * 作成する。元の案件ファイルをコピーする工程は持たないため、スプレッドシートは
 * これ1枚だけで運用できる。
 *
 * ファイル構成:
 *   1. 設定シート関連
 *   2. リンクからのラベル作成
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
var LABEL_COL_WIDTH_MM = 66;    // ラベル1枚の実寸幅。商品が異なる場合は要調整
var MM_TO_PX = 96 / 25.4;

var LABEL_SIZE_3X8 = '3×8（66mm×33.9mm）';
var LABEL_SIZE_3X6 = '3×6（66mm×45.2mm）';

// 内訳（3段の高さpx）は、3×8は34/48/46（メニュー名の2行折り返しのため、数量段
// からさらに2px減らしてメニュー名段に足した。合計128px＝33.9mm）、3×6は
// 171px（＝45.2mm）を同じ比率で配分した45/64/62を初期値とする。
var LABEL_SIZE_PRESETS = {};
LABEL_SIZE_PRESETS[LABEL_SIZE_3X8] = { rows: 8, subrowHeightsPx: [34, 48, 46] };
LABEL_SIZE_PRESETS[LABEL_SIZE_3X6] = { rows: 6, subrowHeightsPx: [45, 64, 62] };

/**
 * 「設定」シートB3で選択されたラベルサイズのプリセット（行数・3段の高さpx）を返す。
 * 未選択・不正な値の場合は3×8（従来のデフォルト）を返す。
 */
function getLabelSizePreset_(configSheet) {
  var raw = String(configSheet.getRange(CONFIG_LABEL_SIZE_CELL).getValue() || '').trim();
  return LABEL_SIZE_PRESETS[raw] || LABEL_SIZE_PRESETS[LABEL_SIZE_3X8];
}

var LABEL_FONT_SIZE = 14;      // 1段目（日付＋企業名）のフォントサイズ
var LABEL_MENU_FONT_SIZE = 12; // 2段目（メニュー名）。1段目より2pt小さい
var LABEL_QTY_FONT_SIZE = 25;  // 3段目（数量）。太字・大きめフォントで強調する
var LABEL_MENU_MAX_ZENKAKU_LEN = 28; // 2段目（メニュー名）は全角換算でこの文字数を超えたら切り捨てる

/**
 * メニュー「ラベル作成」→「作成」から呼び出されるメイン関数。
 * 「設定」シートのURL一覧を1件ずつ開き、リンク先ファイル内の指定タブから
 * 直接（コピーせずに）日付・企業名・メニュー名・数量を読み取り、このスプレッド
 * シート自身に案件ごとのラベルタブを作成する。実行のたびに、今回のURL一覧に
 * 対応しなくなった古いラベルタブは削除される。
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

  var currentCaseNames = [];
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

      // 前回実行分の同名タブは重複とみなさず再利用・上書きするため、
      // 「今回すでに割り当てたタブ名」とだけ重複チェックする
      var baseName = sanitizeSheetName_((parsed.date ? parsed.date + ' ' : '') + parsed.company);
      var caseName = uniqueCaseNameForThisRun_(currentCaseNames, baseName);
      currentCaseNames.push(caseName);

      writeLabelSheetForCase_(ss, caseName, entries, sizePreset);
      totalEntries += entries.length;
    } catch (e) {
      warnings.push(displayLabel + ' : 処理中にエラーが発生しました（' + e.message + '）');
    }
  });

  removeStaleLabelSheets_(ss, currentCaseNames);

  if (currentCaseNames.length === 0) {
    ui.alert('ラベルに出力できる案件データが見つかりませんでした。');
  } else {
    var message = currentCaseNames.length + '件の案件・計' + totalEntries + '件のメニューから' +
      'ラベル' + (totalEntries * LABEL_COPIES_PER_ENTRY) + '枚を作成しました。\n' +
      currentCaseNames.join('\n');
    ui.alert('ラベル作成 結果', message, ui.ButtonSet.OK);
  }

  // 警告・エラーは通常の結果に埋もれて見落とされないよう、別ダイアログで目立たせて表示する
  if (warnings.length > 0) {
    ui.alert('⚠️要確認', warnings.join('\n'), ui.ButtonSet.OK);
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

function sanitizeSheetName_(name) {
  var sanitized = String(name).replace(/[\[\]\*\?\/\\:]/g, '_').trim();
  if (sanitized.length > 100) {
    sanitized = sanitized.substring(0, 100);
  }
  return sanitized || 'シート';
}

/**
 * 今回の実行で既に割り当てたタブ名（currentCaseNames）とだけ重複チェックする。
 * スプレッドシート全体の既存タブと比較しないのは、前回実行分の同名ラベルタブを
 * 「再利用・上書き」対象として正しく扱うため（重複とみなして(2)付きの別タブを
 * 作ってしまわないようにする）。今回の実行内で2つの案件が同じ名前になった
 * 場合だけ(2)が付く。
 */
function uniqueCaseNameForThisRun_(currentCaseNames, baseName) {
  var name = baseName;
  var suffix = 2;
  while (currentCaseNames.indexOf(name) !== -1) {
    name = baseName + ' (' + suffix + ')';
    suffix++;
  }
  return name;
}

/**
 * ラベル専用タブに対応しない（今回のURL一覧に案件が存在しない）シートを削除する
 * （案件が削除・再作成された場合に古いラベルタブが残らないようにする）。
 */
function removeStaleLabelSheets_(ss, currentCaseNames) {
  deleteSheetsWhere_(ss, function (sheet) {
    var name = sheet.getName();
    return name !== CONFIG_SHEET_NAME && currentCaseNames.indexOf(name) === -1;
  });
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

/**
 * 1案件分の entries を3列×N行（1件＝3段のセル。Nは「設定」シートで選んだ
 * sizePresetの行数）のグリッドに配置し、このスプレッドシート内の同名タブへ
 * 書き込む。1エントリにつき縦に隣接するLABEL_COPIES_PER_ENTRY枚（同一内容を
 * 複製）を配置し、1ページ分（3列×N行）ごとに次のブロック＝次ページへ折り返す。
 * 既に同名タブがあれば中身だけ消して再利用するため、再実行しても古い内容は
 * 残らない（タブ自体を削除しないのは、印刷余白などタブに紐づく設定を
 * 失わないようにするため）。
 */
function writeLabelSheetForCase_(ss, caseName, entries, sizePreset) {
  // 削除して作り直すと、印刷余白などタブに紐づく設定が失われる可能性があるため、
  // 既存タブがあれば中身だけ消して（clear）再利用する。
  var sheet = ss.getSheetByName(caseName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(caseName);
  }

  var labelRows = sizePreset.rows;
  var subrowHeightsPx = sizePreset.subrowHeightsPx;
  var subrows = subrowHeightsPx.length;

  var entriesPerColumn = Math.floor(labelRows / LABEL_COPIES_PER_ENTRY);
  var entriesPerPage = entriesPerColumn * LABEL_COLS;
  var totalPages = Math.ceil(entries.length / entriesPerPage);
  var totalRows = totalPages * labelRows * subrows;

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
          var rowBase = page * labelRows * subrows + physicalSlot * subrows;
          writeLabelCellGroup_(sheet, rowBase, col + 1, entry);
        }
      }
    }
  }

  for (var col1 = 1; col1 <= LABEL_COLS; col1++) {
    sheet.setColumnWidth(col1, Math.round(LABEL_COL_WIDTH_MM * MM_TO_PX));
  }
  for (var r = 0; r < totalRows; r++) {
    var subIndex = r % subrows;
    sheet.setRowHeight(r + 1, subrowHeightsPx[subIndex]);
  }
}

/**
 * ラベル1件分（3段）を書き込む。
 *   1段目: 日付＋企業名＋「様」（左寄せ、WrapStrategy.CLIP＝折り返さず高さ固定。
 *          文字数が多い場合は折り返さず末尾が切れる）
 *   2段目: メニュー名（中央寄せ・1段目より2pt小さいフォント。
 *          WrapStrategy.WRAP＝2行まで折り返す。全角文字数
 *          （LABEL_MENU_MAX_ZENKAKU_LEN）を超える分は事前に切り捨てているため、
 *          2行に収まりきらず段の高さが崩れることを防いでいる）
 *   3段目: 数量（中央寄せ・太字・大きめフォントで強調、WrapStrategy.CLIP）
 */
function writeLabelCellGroup_(sheet, rowBase, col1, entry) {
  sheet.getRange(rowBase + 1, col1).setValue(entry.date + ' ' + entry.company + ' 様')
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
// 3. 共通ヘルパー
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
