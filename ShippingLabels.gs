/**
 * 「出荷シート」（例:「【作成中2026/10/2（金）出荷分】motoGP_竹橋製造/出荷シート」）
 * フォーマット専用のラベル作成ツール。Code.gs（案件リンク一覧からの読み込み用）とは
 * 完全に独立したファイル・関数群で、出荷シートのスプレッドシートに直接バインドして
 * 使うコンテナバインドスクリプトとして書く（Code.gsと同じApps Scriptプロジェクトには
 * 貼り付けない。onOpen()が重複するため）。
 *
 * 各タブ（例:「102出荷①朝和」〜「102出荷⑥デザート」）は同じレイアウトを持つ前提。
 *   - G4:  出荷日（見出し情報）
 *   - B12: 提供日の見出し行（見出し情報）
 *   - 項目表: E列=メニュー名（新しい品目の先頭行にしか入らない。以降はforward-fillで
 *     引き継ぐ）、G列=梱包メモ等（見出しなし）、I列=数量、J列=場所
 * 項目表のうち数量・場所のどちらかに値がある行を1件のラベルとして集計し、
 * A-oneラベルシール18面相当（66mm×45.2mm・3列×6行。既存プロジェクトの
 * 24面版66mm×33.9mmと同じ総高さを6行で割った寸法）に印刷できる形式で
 * ラベル専用スプレッドシートへ出力する。
 *
 * ファイル構成:
 *   1. 設定シート関連
 *   2. 出荷シートからのデータ抽出
 *   3. 印刷用ラベル作成
 */

// ============================================================
// 1. 設定シート関連
// ============================================================

var SHIP_CONFIG_SHEET_NAME = '設定';
var SHIP_CONFIG_LABEL_SS_CELL = 'B1'; // ラベル出力先スプレッドシートのURL/ID（自動設定）

function getOrCreateShipConfigSheet_(ss) {
  var sheet = ss.getSheetByName(SHIP_CONFIG_SHEET_NAME);
  if (sheet) {
    return sheet;
  }
  sheet = ss.insertSheet(SHIP_CONFIG_SHEET_NAME);
  sheet.getRange('A1').setValue('ラベル出力先スプレッドシートID（自動設定・空欄でOK）');
  sheet.getRange('A1').setFontWeight('bold');
  return sheet;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('出荷ラベル')
    .addItem('ラベル作成', 'createShippingLabels')
    .addToUi();
}

// ============================================================
// 2. 出荷シートからのデータ抽出
// ============================================================

var SHIP_ITEM_MENU_COL = 4;  // E列（0始まり）: メニュー名（新しい品目の先頭行のみ）
var SHIP_ITEM_NOTE_COL = 6;  // G列: 梱包メモ等（見出しなし）
var SHIP_ITEM_QTY_COL = 8;   // I列: 数量
var SHIP_ITEM_PLACE_COL = 9; // J列: 場所
var SHIP_SHIP_DATE_CELL = 'G4';
var SHIP_PROVIDE_LINE_CELL = 'B12';

/**
 * タブ1枚分から、ラベルに出力するエントリ（メニュー名・梱包メモ・数量・場所）と、
 * 見出し情報（出荷日・提供日）を集める。項目表のヘッダー行（E列が「メニュー名」に
 * 一致する行）を探し、そこから下のデータ行を1行ずつ処理する。メニュー名(E)は
 * 新しい品目の先頭行にしか入っていないため、直前の非空値をforward-fillしながら
 * 「現在の品目名」として扱う。数量(I)・場所(J)のどちらかに値がある行だけを
 * 1件のラベルとして採用する。
 */
function collectShippingEntries_(sheet) {
  var values = sheet.getDataRange().getValues();
  var headerRow = -1;
  for (var r = 0; r < values.length; r++) {
    if (String(values[r][SHIP_ITEM_MENU_COL] || '').trim() === 'メニュー名') {
      headerRow = r;
      break;
    }
  }
  if (headerRow === -1) {
    return null;
  }

  var shipDate = String(sheet.getRange(SHIP_SHIP_DATE_CELL).getValue() || '').trim();
  var provideLine = String(sheet.getRange(SHIP_PROVIDE_LINE_CELL).getValue() || '').trim();

  var entries = [];
  var currentMenu = '';
  for (var i = headerRow + 1; i < values.length; i++) {
    var menuVal = String(values[i][SHIP_ITEM_MENU_COL] || '').trim();
    if (menuVal) {
      currentMenu = menuVal;
    }
    var qty = values[i][SHIP_ITEM_QTY_COL];
    var place = String(values[i][SHIP_ITEM_PLACE_COL] || '').trim();
    if (!qty && !place) {
      continue;
    }
    var note = String(values[i][SHIP_ITEM_NOTE_COL] || '').trim();
    entries.push({
      shipDate: shipDate, provideLine: provideLine,
      menu: currentMenu, note: note, qty: qty, place: place
    });
  }

  return { shipDate: shipDate, provideLine: provideLine, entries: entries };
}

// ============================================================
// 3. 印刷用ラベル作成
// ============================================================

var SHIP_LABEL_SPREADSHEET_SUFFIX = '（出荷ラベル印刷）';

// 面付け（A-one マルチプリンタ用ラベルシール18面相当: 66mm×45.2mm、3列×6行。
// 24面版（66mm×33.9mm、3列×8行）の総高さ271.2mmは変えず、8行ではなく6行で
// 均等に割った高さにしている。1段目が3行になったことで従来の33.9mmでは
// 収まらなくなったため）
var SHIP_LABEL_COLS = 3;
var SHIP_LABEL_ROWS = 6;
var SHIP_LABEL_COPIES_PER_ENTRY = 1; // 1件につき1枚（必要ならここを増やす）
var SHIP_LABEL_COL_WIDTH_MM = 66;
var SHIP_MM_TO_PX = 96 / 25.4;

// ラベル1枚（実寸45.2mm、96dpi換算で171px）を3段に分ける。
// 内訳は54px/58px/59px（1段目が出荷日・提供日・場所の3行分必要になったため、
// Code.gsの34/48/46から全体的に底上げしている）。実際に印刷して3行が窮屈な
// 場合は、2段目・3段目から少しずつ高さを移して調整する
var SHIP_LABEL_SUBROW_HEIGHTS_PX = [54, 58, 59];
var SHIP_LABEL_SUBROWS = SHIP_LABEL_SUBROW_HEIGHTS_PX.length;

var SHIP_LABEL_HEADER_FONT_SIZE = 10; // 1段目（出荷日・提供日・場所の3行、共通サイズ）
var SHIP_LABEL_MENU_FONT_SIZE = 12;   // 2段目（メニュー名＋梱包メモ）
var SHIP_LABEL_QTY_FONT_SIZE = 28;    // 3段目（数量）
var SHIP_LABEL_MENU_MAX_ZENKAKU_LEN = 30; // 2段目は全角換算でこの文字数を超えたら切り捨てる

/**
 * メニュー「出荷ラベル」→「ラベル作成」から呼び出されるメイン関数。
 * アクティブなスプレッドシートの全タブを走査し、各タブの項目データを
 * ラベル専用スプレッドシート内の同名タブへ出力する。
 */
function createShippingLabels() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = getOrCreateShipConfigSheet_(ss);
  var labelSs = getOrCreateShipLabelSpreadsheet_(ss, configSheet);

  var sourceTabNames = [];
  var resultLines = [];
  var warnings = [];
  var totalEntries = 0;

  ss.getSheets().forEach(function (sheet) {
    var name = sheet.getName();
    if (name === SHIP_CONFIG_SHEET_NAME) {
      return;
    }
    sourceTabNames.push(name);
    try {
      var data = collectShippingEntries_(sheet);
      if (!data) {
        warnings.push(name + ' : 項目表のヘッダー行（メニュー名）が見つかりません');
        return;
      }
      if (data.entries.length === 0) {
        warnings.push(name + ' : データがありません');
        return;
      }
      writeShipLabelSheetForTab_(labelSs, name, data.entries);
      totalEntries += data.entries.length;
      resultLines.push(name + '（出荷日: ' + data.shipDate + ' / 提供日: ' + data.provideLine + '）: ' +
        data.entries.length + '件');
    } catch (e) {
      warnings.push(name + ' : 処理中にエラーが発生しました（' + e.message + '）');
    }
  });

  removeStaleShipLabelSheets_(labelSs, sourceTabNames);

  if (resultLines.length === 0) {
    ui.alert('ラベルに出力できる出荷データが見つかりませんでした。');
  } else {
    var message = resultLines.length + 'タブ・計' + totalEntries + '件からラベルを作成しました。\n' +
      resultLines.join('\n') + '\n' + labelSs.getUrl();
    ui.alert('ラベル作成 結果', message, ui.ButtonSet.OK);
  }

  if (warnings.length > 0) {
    ui.alert('⚠️要確認', warnings.join('\n'), ui.ButtonSet.OK);
  }
}

/**
 * ラベル専用スプレッドシートを取得する。「設定」シートに保存済みのURL/IDがあれば
 * それを再利用し、なければ新規作成して同じ親フォルダに置き、IDを保存する。
 */
function getOrCreateShipLabelSpreadsheet_(ss, configSheet) {
  var savedRaw = configSheet.getRange(SHIP_CONFIG_LABEL_SS_CELL).getValue();
  var savedId = savedRaw ? shipExtractSpreadsheetId_(savedRaw) : '';
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {
      // 保存済みIDが無効（削除済みなど）の場合は新規作成にフォールバックする
    }
  }

  var labelSs = SpreadsheetApp.create(ss.getName() + SHIP_LABEL_SPREADSHEET_SUFFIX);
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

  configSheet.getRange(SHIP_CONFIG_LABEL_SS_CELL).setValue(labelSs.getId());
  return labelSs;
}

/** ラベル専用スプレッドシート内で、現在の出荷元タブ名に対応しないシートを削除する。 */
function removeStaleShipLabelSheets_(labelSs, currentTabNames) {
  var sheets = labelSs.getSheets();
  var remaining = sheets.length;
  sheets.forEach(function (sheet) {
    if (currentTabNames.indexOf(sheet.getName()) === -1 && remaining > 1) {
      labelSs.deleteSheet(sheet);
      remaining--;
    }
  });
}

/**
 * 1タブ分の entries を3列×6行（1件＝3段のセル）のグリッドに配置し、labelSs内の
 * 同名タブへ書き込む。既に同名タブがあれば中身だけ消して再利用する
 * （印刷余白などタブに紐づく設定が維持される可能性があるため）。
 */
function writeShipLabelSheetForTab_(labelSs, tabName, entries) {
  var sheet = labelSs.getSheetByName(tabName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = labelSs.insertSheet(tabName);
  }

  var entriesPerColumn = Math.floor(SHIP_LABEL_ROWS / SHIP_LABEL_COPIES_PER_ENTRY);
  var entriesPerPage = entriesPerColumn * SHIP_LABEL_COLS;
  var totalPages = Math.ceil(entries.length / entriesPerPage);
  var totalRows = totalPages * SHIP_LABEL_ROWS * SHIP_LABEL_SUBROWS;

  // 最終ページで3列とも「データがある列」として印刷範囲に含まれるよう、
  // 先に全セルへ空文字列と薄い罫線を設定しておく（Code.gsと同じ対策）。
  sheet.getRange(1, 1, totalRows, SHIP_LABEL_COLS)
    .setValue('')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setBorder(true, true, true, true, true, true, '#f2f2f2', SpreadsheetApp.BorderStyle.SOLID);

  var index = 0;
  for (var page = 0; page < totalPages; page++) {
    for (var col = 0; col < SHIP_LABEL_COLS; col++) {
      for (var slot = 0; slot < entriesPerColumn; slot++) {
        if (index >= entries.length) {
          break;
        }
        var entry = entries[index++];
        for (var copy = 0; copy < SHIP_LABEL_COPIES_PER_ENTRY; copy++) {
          var physicalSlot = slot * SHIP_LABEL_COPIES_PER_ENTRY + copy;
          var rowBase = page * SHIP_LABEL_ROWS * SHIP_LABEL_SUBROWS + physicalSlot * SHIP_LABEL_SUBROWS;
          writeShipLabelCellGroup_(sheet, rowBase, col + 1, entry);
        }
      }
    }
  }

  for (var col1 = 1; col1 <= SHIP_LABEL_COLS; col1++) {
    sheet.setColumnWidth(col1, Math.round(SHIP_LABEL_COL_WIDTH_MM * SHIP_MM_TO_PX));
  }
  for (var r = 0; r < totalRows; r++) {
    var subIndex = r % SHIP_LABEL_SUBROWS;
    sheet.setRowHeight(r + 1, SHIP_LABEL_SUBROW_HEIGHTS_PX[subIndex]);
  }
}

/**
 * ラベル1件分（3段）を書き込む。
 *   1段目: 出荷日・提供日・場所の3行（改行区切りで1セルにまとめる。左寄せ・
 *          太字・共通フォントサイズ、WrapStrategy.CLIP＝改行文字はそのまま
 *          改行として表示されるが、1行が長すぎる場合の自動折り返しはしない）
 *   2段目: メニュー名＋梱包メモ（中央寄せ。WrapStrategy.WRAP＝2行まで折り返す。
 *          全角30文字（SHIP_LABEL_MENU_MAX_ZENKAKU_LEN）を超える分は事前に
 *          切り捨てているため、2行に収まりきらず段の高さが崩れることを防いでいる）
 *   3段目: 数量（中央寄せ・太字・大きめフォントで強調、WrapStrategy.CLIP）
 */
function writeShipLabelCellGroup_(sheet, rowBase, col1, entry) {
  var headerText = '出荷日 ' + entry.shipDate + '\n' + entry.provideLine + '\n' + entry.place;
  sheet.getRange(rowBase + 1, col1).setValue(headerText)
    .setFontSize(SHIP_LABEL_HEADER_FONT_SIZE)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  var menuText = entry.note ? entry.menu + ' ' + entry.note : entry.menu;
  sheet.getRange(rowBase + 2, col1).setValue(shipTruncateByZenkakuWidth_(menuText, SHIP_LABEL_MENU_MAX_ZENKAKU_LEN))
    .setFontSize(SHIP_LABEL_MENU_FONT_SIZE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  sheet.getRange(rowBase + 3, col1).setValue(entry.qty)
    .setFontSize(SHIP_LABEL_QTY_FONT_SIZE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

/**
 * 文字列を全角換算で指定した文字数までに切り詰める（半角文字は0.5文字分として
 * カウントする）。超えた分は末尾を単純に切り捨てる（省略記号は付けない）。
 */
function shipTruncateByZenkakuWidth_(text, maxZenkakuWidth) {
  var result = '';
  var width = 0;
  for (var i = 0; i < text.length; i++) {
    var ch = text.charAt(i);
    var charWidth = shipIsHalfWidthChar_(ch) ? 0.5 : 1;
    if (width + charWidth > maxZenkakuWidth) {
      break;
    }
    result += ch;
    width += charWidth;
  }
  return result;
}

/** 半角英数・記号（U+0000〜U+00FF）、半角カタカナ（U+FF61〜U+FF9F）を半角とみなす。 */
function shipIsHalfWidthChar_(ch) {
  var code = ch.charCodeAt(0);
  return (code >= 0x0000 && code <= 0x00FF) || (code >= 0xFF61 && code <= 0xFF9F);
}

/**
 * スプレッドシートのURLまたは素のIDから、スプレッドシートIDを取り出す。
 */
function shipExtractSpreadsheetId_(input) {
  var text = String(input).trim();
  var match = text.match(/\/(?:spreadsheets|file)\/d\/([a-zA-Z0-9_-]+)/) ||
    text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match) {
    return match[1];
  }
  return text;
}
