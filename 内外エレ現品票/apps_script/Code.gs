/*
 * 現品票 自動作成スクリプト
 *
 * 「内外エレクトロニクス生産計画」からNEI注番号・型式を自動取得し、
 * QRコード付きの現品票をこのスプレッドシート内に自動レイアウトする。
 * QRコード生成は外部通信を一切行わず、このスクリプト内で完結する
 * （qrcodegen.gs に同梱した QR Code generator library (Project Nayuki, MIT License) を使用）。
 *
 * 使い方:
 *   1. スプレッドシート上部メニュー「管理№シート作成」>「今月の入力シートを開く」を実行する
 *      （「入力_202607」のように月ごとのシートが自動で作られる）
 *   2. そのシートのA列（2行目以降）に、現品票を作りたいNEI注番号を1行ずつ入力する
 *      （生産計画シートのC列「NEI注番号」の値。例: JJFK25008650）
 *   3. 「管理№シート作成」>「連番を記載」を実行する
 *      （まだB列に連番が無い行すべてに、生産計画シートのL列「管理No.」の値がそのまま取得される。
 *       同じNEI注番号が複数行に分かれている場合は、それぞれの行のL列の値をすべて取得する。
 *       印刷時は3桁ゼロ埋め（例: "1"→"001"）で表示される）
 *   4. 印刷したい行のC列「印刷対象」にチェックを入れる
 *   5. 「管理№シート作成」>「チェックした行を印刷シートに出力」を実行する
 *      （チェックした行だけが「印刷」シートに出力される。何度でも、別のチェックの組み合わせで実行し直せる）
 *   6. 「印刷」シートに印刷用レイアウトが生成されるので、内容を確認してから
 *      ファイル > 印刷 で「用紙: A4」「向き: 横」「余白: なし」「拡大縮小: 標準(100%)」を選んで
 *      印刷し、これまで通り手で切り分ける
 *      （拡大縮小を「幅に合わせる」等にすると、14行(2ブロック)ごとの改ページ位置がズレるので注意）
 */

/*==================== 設定 ====================*/

// コピー元「内外エレクトロニクス生産計画」のスプレッドシートID（URLの /d/ と /edit の間の文字列）
var SOURCE_SPREADSHEET_ID = '1VQoAyx76euNgH7DsQWljKPyp1oQbmQqjr-FWUL2-wEU';

// コピー元シートのgid（URLの #gid=... の数字）
var SOURCE_SHEET_GID = 332550892;

// コピー元シートの列番号（1=A, 2=B, 3=C, 4=D, 5=E, ... 12=L）
// 現品票に印字される注番はC列「NEI注番号」（例: JJFK25008650）であり、B列の「注番」ではない点に注意
var SOURCE_COL_ORDER_NO = 3;      // C列: NEI注番号
var SOURCE_COL_MODEL = 4;         // D列: 型式
var SOURCE_COL_MANAGEMENT_NO = 12; // L列: 管理No.（現品票の連番はここから取得する。自動採番はしない）

// このスプレッドシート内のシート名
// 入力シートは月ごとに「入力_202607」のように自動作成される（INPUT_SHEET_PREFIX + yyyyMM）
var INPUT_SHEET_PREFIX = '入力_';
// 同じNEI注番号が入力シート内に複数回入力されていた行のB列(連番)に表示するエラー文言
var DUPLICATE_ERROR_LABEL = 'エラー:重複';
// 印刷シートは1枚の連続したシートにする（ページごとにシートを分けない）。
// Google Sheetsには「ここで改ページする」という指定をスクリプトから行うAPIが無いため、
// 印刷時に「余白: なし」「拡大縮小: 標準(100%)」を選んだときに、行の高さの積み上げが
// ちょうどA4横向き1ページ分(約793px)に収まるようBLOCK_ROW_HEIGHTS×BLOCKS_PER_PAGEを
// 計算してあり、これによって自然に14行(2ブロック)ごとに改ページされる。
// ※「拡大縮小: 幅に合わせる/ページ数に合わせる」を選ぶとこの計算が崩れるので使わないこと。
var OUTPUT_SHEET_NAME = '印刷';

// 現品票レイアウト設定
// 元のExcelブック(内外エレクトロニクス_入庫処理_2026.xlsm)のスタイル定義を直接調べて割り出した実際の値。
// 注意: setRowHeight/setColumnWidthはピクセル指定のAPIのため、
// 「pt換算したい高さ×4/3」でここに指定する（例: 130pt相当にしたいなら173px）
// 実際に印刷してみると、理論値(96dpi換算)ぴったりだと余白なしでもわずかに収まりきらなかったため、
// 安全マージンを取って少し詰めてある（型式・連番・注番などフォントサイズに直結する行の高さは変えていない）。
var BLOCK_ROW_HEIGHTS = [32, 173, 8, 27, 8, 87, 15]; // 1ブロック=7行の各行の高さ(px)
// 1ブロック=350px、2ブロック=700px。A4横向き(210mm)の印刷可能高さは理論上は余白なしで約793.7pxだが、
// 実機では収まりきらなかったため700px(約88px の余裕)を目安に14行区切りで改ページされるようにしてある。
var BLOCKS_PER_PAGE = 2;
// A4横向き(297mm)の印刷可能幅は理論上は余白なしで約1122pxだが、安全マージンを見て
// 合計幅を1080px程度(A+B+C+D+E列)に収める。
// 注意: TEXT_COL_WIDTH_PXは注番("JJFK25008650-1/2"等)がORDER_NO_FONT_SIZEで
// 折り返さずに収まる最小幅が約500pxのため、これより狭くすると隣の列に文字がはみ出す。
var TEXT_COL_WIDTH_PX = 510; // 現品票テキスト列の幅(px)
var GAP_COL_WIDTH_PX = 20;      // 中央(C列)の隙間の幅(px)
var A_COL_WIDTH_PX = GAP_COL_WIDTH_PX;   // 左端(A列)の余白。中央の隙間と同じ幅に揃える
var E_COL_WIDTH_PX = GAP_COL_WIDTH_PX;   // 右端(D列の右、E列)の余白。同じく中央の隙間と同じ幅に揃える
var LEFT_COL = 2;  // B列を左側の現品票の開始列にする
var RIGHT_COL = 4; // D列を右側の現品票の開始列にする(左のB列+ギャップC列の次)
var QR_MODULE_PX = 8;   // QR 1モジュールあたりの解像度(px)。画質を上げたい場合は増やす
// QRの表示サイズ(px)。大きすぎたので、連番行の高さいっぱい(BLOCK_ROW_HEIGHTS[1]-6)ではなく
// ひとまわり小さい固定値にしてある。もっと大きく/小さくしたい場合はこの数値を直接調整する。
var QR_DISPLAY_SIZE_PX = 110;
// 連番("001"等)とQRコードの間の距離。連番テキストの右端からのおおよその位置(px)を直接指定する
// （列幅いっぱいの右端ではなく、数字にもっと近づけたい場合はこの値を小さくする）
// QRと連番をまとめて1cm(約38px)右にずらすためQR_OFFSET_X_PXに38pxを加算してある。
var QR_OFFSET_X_PX = 260 + 38;
// QRを連番の行の縦方向中央に合わせる（行の高さからQRの高さを引いた分の半分だけ上から空ける）
var QR_OFFSET_Y_PX = Math.round((BLOCK_ROW_HEIGHTS[1] - QR_DISPLAY_SIZE_PX) / 2);
var MONTH_FONT_SIZE = 16;    // 年月のフォントサイズ（Excel実物は太字）
var SERIAL_FONT_SIZE = 130;  // 連番のフォントサイズ
var MODEL_FONT_SIZE = 14;    // 型式のフォントサイズ
var ORDER_NO_FONT_SIZE = 44; // 注番のフォントサイズ
// 連番の文字列の前に付ける余白。セルの値である連番自体には「Xpx右にずらす」という
// 精密な指定ができるAPIが無いため、リッチテキストで先頭の空白だけ別の(小さい)フォントサイズに
// して、その空白の見た目の横幅で右にずらす。
// 半角スペース1文字の幅は、そのフォントサイズのおよそ半分(0.5em)になるため、
// 目安の計算式は: SERIAL_LEFT_PAD_FONT_SIZE ≈ 欲しいシフト量(px) ÷ (4/3) ÷ 0.5
// 目標は約38px(1cm)右にずらすことなので 38 ÷ (4/3) ÷ 0.5 ≈ 57pt にしてある。
// まだ左に寄っている/右に寄りすぎている場合は、この数値を増減して調整する。
var SERIAL_LEFT_PAD = ' ';
var SERIAL_LEFT_PAD_FONT_SIZE = 57;

/*==================== メニュー ====================*/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('管理№シート作成')
    .addItem('今月の入力シートを開く', 'openCurrentMonthInputSheet')
    .addItem('別の月の入力シートを開く', 'openOtherMonthInputSheet')
    .addItem('連番を記載', 'assignSerialNumbers')
    .addSeparator()
    .addItem('印刷対象をすべてチェック', 'checkAllPrintTargets')
    .addItem('印刷対象をすべて解除', 'uncheckAllPrintTargets')
    .addItem('チェックした行を印刷シートに出力', 'printSelectedTags')
    .addToUi();
}

// 今開いている「入力_yyyyMM」シートで、A列にNEI注番号が入っている行のC列(印刷対象)をすべてチェックする
function checkAllPrintTargets() {
  setAllPrintTargets_(true);
}

// 今開いている「入力_yyyyMM」シートで、A列にNEI注番号が入っている行のC列(印刷対象)をすべて解除する
function uncheckAllPrintTargets() {
  setAllPrintTargets_(false);
}

function setAllPrintTargets_(checked) {
  var inputSheet = resolveTargetInputSheet_().sheet;
  var lastRow = inputSheet.getLastRow();
  if (lastRow < 2) return;
  var aValues = inputSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var cValues = [];
  for (var i = 0; i < aValues.length; i++) {
    cValues.push([String(aValues[i][0]).trim() !== '' ? checked : false]);
  }
  inputSheet.getRange(2, 3, cValues.length, 1).setValues(cValues);
}

// 今月分の「入力_yyyyMM」シートを（なければ作成して）開く
function openCurrentMonthInputSheet() {
  var sheet = ensureInputSheet_();
  sheet.activate();
}

// 来月・再来月など、今月以外の「入力_yyyyMM」シートを（なければ作成して）開く
function openOtherMonthInputSheet() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '対象の年月を入力',
    '入力シートを開きたい年月を6桁で入力してください（例: 来月なら202608）',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var monthKey = response.getResponseText().trim();
  if (!/^\d{6}$/.test(monthKey)) {
    ui.alert('年月は "202608" のように6桁の数字で入力してください。');
    return;
  }
  var sheet = ensureInputSheet_(monthKey);
  sheet.activate();
}

// 連番の記載・印刷対象は「今アクティブになっている入力_yyyyMMシート」を対象にする。
// これにより、来月・再来月分を先行してこのシートで作業していても、そのシートに対して処理できる。
// アクティブなシートが入力シートでない場合（生産計画シートや印刷シートを見ている場合など）は、
// 今月の入力シートを対象にする。
function resolveTargetInputSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var active = ss.getActiveSheet();
  var monthKey = getMonthKeyFromInputSheetName_(active.getName());
  if (monthKey) {
    return { sheet: active, monthKey: monthKey };
  }
  var tz = ss.getSpreadsheetTimeZone();
  var todayKey = Utilities.formatDate(new Date(), tz, 'yyyyMM');
  return { sheet: ensureInputSheet_(todayKey), monthKey: todayKey };
}

// シート名が「入力_202608」のような形式ならその"202608"部分を返す。そうでなければnullを返す。
function getMonthKeyFromInputSheetName_(sheetName) {
  if (sheetName.indexOf(INPUT_SHEET_PREFIX) !== 0) return null;
  var rest = sheetName.substring(INPUT_SHEET_PREFIX.length);
  return /^\d{6}$/.test(rest) ? rest : null;
}

/*==================== メイン処理 ====================*/

// ステップ1: チェック状態は見ず、まだ連番が記載されていない行すべてに連番を記載する。
// 印刷シートは作らない（印刷は後でチェックした行だけをステップ2で行う）。
function assignSerialNumbers() {
  var ui = SpreadsheetApp.getUi();
  var target = resolveTargetInputSheet_();
  var inputSheet = target.sheet;
  var entries = readInputOrderNumbers_(inputSheet); // [{row, orderNo, existingSerial, checked}, ...]
  if (entries.length === 0) {
    ui.alert('「' + inputSheet.getName() + '」シートのA列（2行目以降）にNEI注番号を入力してください。');
    return;
  }

  // B列（連番）に既に値が入っている行は「採番済み」として飛ばす。
  var toProcess = entries.filter(function (e) { return !e.existingSerial; });
  if (toProcess.length === 0) {
    ui.alert('新しく連番を記載する行がありません（入力済みの行は全て連番が記載済みです）。');
    return;
  }

  // 同じNEI注番号がこの入力シート内に複数回入力されている（コピペミス等）行は、
  // 連番を記載せずにB列へエラー表示する。全行（記載済みの行も含む）を対象に重複を数える。
  var orderNoCounts = {};
  for (var c = 0; c < entries.length; c++) {
    var key = entries[c].orderNo;
    orderNoCounts[key] = (orderNoCounts[key] || 0) + 1;
  }
  var duplicateRows = toProcess.filter(function (e) { return orderNoCounts[e.orderNo] > 1; });
  toProcess = toProcess.filter(function (e) { return orderNoCounts[e.orderNo] === 1; });

  var orderNumbers = toProcess.map(function (e) { return e.orderNo; });
  var lookup = lookupOrdersFromSource_(orderNumbers);

  var notFound = [];
  var serialByRow = {};
  var assignedCount = 0;

  for (var d = 0; d < duplicateRows.length; d++) {
    serialByRow[duplicateRows[d].row] = DUPLICATE_ERROR_LABEL;
  }

  for (var i = 0; i < toProcess.length; i++) {
    var orderNo = toProcess[i].orderNo;
    var info = lookup[orderNo];
    if (info === undefined || info.managementNumbers.length === 0) {
      notFound.push(orderNo);
      continue;
    }
    // 連番は生産計画シートのL列「管理No.」の値をそのまま使う（自動採番はしない）。
    // 同じNEI注番号が複数行に分かれている場合は、それぞれの行のL列の値をすべて使う。
    // 印刷時は3桁ゼロ埋めするが、B列には見やすいよう連続する番号を"051-058"のような
    // 範囲表記にまとめ、連続していない場合はカンマ区切りで記録する。
    serialByRow[toProcess[i].row] = compressToRanges_(info.managementNumbers);
    assignedCount += info.managementNumbers.length;
  }

  writeBackSerials_(inputSheet, serialByRow);

  var msg = assignedCount + ' 件の管理No.を生産計画シートから取得しました。\n続けてC列「印刷対象」にチェックを入れ、' +
    '「チェックした行を印刷シートに出力」を実行してください。';
  if (notFound.length > 0) {
    msg += '\n\n生産計画シート内に見つからなかったNEI注番号（連番は記載していません）:\n' + notFound.join('\n');
  }
  if (duplicateRows.length > 0) {
    msg += '\n\n同じNEI注番号がこのシート内に複数回入力されていたため、連番を記載せず' +
      'B列に「' + DUPLICATE_ERROR_LABEL + '」と表示した行があります:\n' +
      duplicateRows.map(function (e) { return e.row + '行目: ' + e.orderNo; }).join('\n') +
      '\n重複している行を削除するか修正してから、B列のエラー表示を消して再実行してください。';
  }
  ui.alert(msg);
}

// ステップ2: C列「印刷対象」にチェックが入っている行（＝既に連番が記載済みの行）を
// 「印刷」シートに出力する。新しい連番は一切記載しない。
function printSelectedTags() {
  var ui = SpreadsheetApp.getUi();
  var target = resolveTargetInputSheet_();
  var inputSheet = target.sheet;
  // 現品票に印字する年月は、実行時の実際の日付ではなく、対象にしている入力シートの年月
  // （例: 入力_202608 なら "2026-08"）を使う。来月・再来月分を先行して作業できるようにするため。
  var monthLabel = target.monthKey.substring(0, 4) + '-' + target.monthKey.substring(4, 6);

  var entries = readInputOrderNumbers_(inputSheet);
  var selected = entries.filter(function (e) { return e.checked; });
  if (selected.length === 0) {
    ui.alert('C列「印刷対象」にチェックが入っている行がありません。印刷したい行にチェックを入れてから実行してください。');
    return;
  }

  var notReady = selected.filter(function (e) { return !e.existingSerial; });
  var hasError = selected.filter(function (e) { return e.existingSerial && !isValidSerialValue_(e.existingSerial); });
  var ready = selected.filter(function (e) { return isValidSerialValue_(e.existingSerial); });
  if (ready.length === 0) {
    ui.alert('チェックされている行はまだ連番が記載されていません。' +
      '先に「管理№シート作成」>「連番を記載」を実行してください。');
    return;
  }

  var orderNumbers = ready.map(function (e) { return e.orderNo; });
  var lookup = lookupOrdersFromSource_(orderNumbers);

  var records = [];
  var notFound = [];
  for (var i = 0; i < ready.length; i++) {
    var orderNo = ready[i].orderNo;
    var info = lookup[orderNo];
    if (info === undefined) {
      notFound.push(orderNo);
      continue;
    }
    var serials = expandSerialRange_(ready[i].existingSerial);
    var qty = serials.length;
    for (var u = 0; u < serials.length; u++) {
      var displayOrderNo = qty > 1 ? (orderNo + '-' + (u + 1) + '/' + qty) : orderNo;
      records.push({
        serial: serials[u],
        model: info.model,
        month: monthLabel,
        orderNo: displayOrderNo
      });
    }
  }

  var msg = records.length + ' 件を「' + OUTPUT_SHEET_NAME + '」シートに出力しました。';
  if (records.length > 0) {
    buildOutputSheet_(records);
    // 出力後は自動的に「印刷」シートに切り替える
    var outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OUTPUT_SHEET_NAME);
    if (outputSheet) outputSheet.activate();
    msg += '\n印刷する際は [ファイル > 印刷] を開き、余白「なし」・拡大縮小「標準(100%)」を選んでください' +
      '（「幅に合わせる」等の自動スケールを選ぶと、14行(2ブロック)ごとの改ページ位置がズレます）。';
  }
  if (notReady.length > 0) {
    msg += '\n\n以下はチェックされていますが、まだ連番が未記載のため出力していません' +
      '（先に「連番を記載」を実行してください）:\n' + notReady.map(function (e) { return e.orderNo; }).join('\n');
  }
  if (hasError.length > 0) {
    msg += '\n\n以下はB列が「' + DUPLICATE_ERROR_LABEL + '」等のエラー表示のため出力していません' +
      '（重複行を修正し、エラー表示を消してから「連番を記載」をやり直してください）:\n' +
      hasError.map(function (e) { return e.orderNo; }).join('\n');
  }
  if (notFound.length > 0) {
    msg += '\n\n生産計画シート内に見つからなかったNEI注番号:\n' + notFound.join('\n');
  }
  ui.alert(msg);
}

// 連番セル(B列)の値が"001"や"001-002"のような正しい連番の形式かどうかを判定する
// （重複エラーなどの文字列が入っている場合はfalseになる）
function isValidSerialValue_(s) {
  // "051"、"051-058"、"051-053,060"、"226,227"のような形式を許可する
  return /^\d+(-\d+)?(,\d+(-\d+)?)*$/.test(String(s));
}

/*==================== データ取得 ====================*/

// monthKeyを省略した場合は「今月」の入力シートを対象にする
function ensureInputSheet_(monthKey) {
  if (!monthKey) {
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    monthKey = Utilities.formatDate(new Date(), tz, 'yyyyMM');
  }
  var sheetName = INPUT_SHEET_PREFIX + monthKey;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.getRange(1, 1, 1, 3).setValues([['NEI注番号', '連番', '印刷対象']]);
  // C列（印刷対象）をチェックボックスにする。既にチェック状態が入っている行はそのまま保持される。
  var checkboxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 3, checkboxRows, 1).insertCheckboxes();
  return sheet;
}

function readInputOrderNumbers_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // A列(NEI注番号)・B列(連番)・C列(印刷対象チェック)をまとめて読む。
  var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var result = [];
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0]).trim();
    if (v !== '') {
      result.push({
        row: i + 2,
        orderNo: v,
        existingSerial: String(values[i][1]).trim(),
        checked: values[i][2] === true
      });
    }
  }
  return result;
}

// 新しく記載した連番（生産計画のL列から取得した管理No.）だけを「入力」シートのB列に書き戻す。
// 1件（数量1）なら"001"、数量2以上なら"001-002"のような範囲表記にする。
// serialByRowに含まれる行だけをピンポイントで更新し、途中にある他の行（既に処理済みの行）には触れない。
function writeBackSerials_(sheet, serialByRow) {
  for (var row in serialByRow) {
    if (!Object.prototype.hasOwnProperty.call(serialByRow, row)) continue;
    sheet.getRange(Number(row), 2).setNumberFormat('@').setValue(serialByRow[row]);
  }
}

function padSerial_(n) {
  return ('000' + n).slice(-3);
}

// 連番の値の配列(生産計画のL列の値そのまま)を、見やすい表示用文字列に圧縮する。
// 連続する番号は"051-058"のような範囲にまとめ、連続していないものはカンマ区切りにする。
// 例: ["051","052","053","060"] → "051-053,060"
function compressToRanges_(managementNumbers) {
  var nums = managementNumbers.map(function (n) { return parseInt(n, 10); })
    .filter(function (n) { return !isNaN(n); })
    .sort(function (a, b) { return a - b; });
  if (nums.length === 0) return '';
  var runs = [];
  var i = 0;
  while (i < nums.length) {
    var start = nums[i];
    var end = start;
    var j = i + 1;
    while (j < nums.length && nums[j] === end + 1) {
      end = nums[j];
      j++;
    }
    runs.push(end > start ? (padSerial_(start) + '-' + padSerial_(end)) : padSerial_(start));
    i = j;
  }
  return runs.join(',');
}

// 連番セルの値を、連番文字列の配列に展開する（再印刷時に使う）。
// "051-058"のような範囲、"226,227"のようなカンマ区切り、"051-053,060"のような
// 両方が混ざった形式のいずれにも対応する。
function expandSerialRange_(s) {
  var segments = String(s).split(',');
  var result = [];
  for (var i = 0; i < segments.length; i++) {
    var seg = segments[i].trim();
    var parts = seg.split('-');
    if (parts.length === 1) {
      result.push(parts[0]);
      continue;
    }
    var start = parseInt(parts[0], 10);
    var end = parseInt(parts[1], 10);
    for (var n = start; n <= end; n++) {
      result.push(padSerial_(n));
    }
  }
  return result;
}

function lookupOrdersFromSource_(orderNumbers) {
  var wanted = {};
  for (var i = 0; i < orderNumbers.length; i++) wanted[orderNumbers[i]] = true;

  var sourceSs = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var sheet = null;
  var sheets = sourceSs.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SOURCE_SHEET_GID) {
      sheet = sheets[i];
      break;
    }
  }
  if (!sheet) throw new Error('コピー元シート(gid=' + SOURCE_SHEET_GID + ')が見つかりません。');

  var lastRow = sheet.getLastRow();
  // 注番号・型式・管理No.の列をまとめて1回で読み込む
  var firstCol = Math.min(SOURCE_COL_ORDER_NO, SOURCE_COL_MODEL, SOURCE_COL_MANAGEMENT_NO);
  var lastCol = Math.max(SOURCE_COL_ORDER_NO, SOURCE_COL_MODEL, SOURCE_COL_MANAGEMENT_NO);
  var block = sheet.getRange(1, firstCol, lastRow, lastCol - firstCol + 1).getValues();
  var orderOffset = SOURCE_COL_ORDER_NO - firstCol;
  var modelOffset = SOURCE_COL_MODEL - firstCol;
  var mgmtOffset = SOURCE_COL_MANAGEMENT_NO - firstCol;

  // 生産計画側は、数量が2以上の場合、同じNEI注文番号の行を数量分だけ複数行に分けて登録し、
  // それぞれの行のL列(管理No.)に別々の番号を振る運用になっている。
  // そのため、現品票の連番は自動採番せず、同じNEI注文番号に対応する行すべてのL列の値を
  // 出現順にそのまま集めて使う（件数=行数がそのまま数量になる）。
  var result = {};
  for (var r = 0; r < lastRow; r++) {
    var orderVal = String(block[r][orderOffset]).trim();
    if (orderVal !== '' && wanted[orderVal]) {
      if (result[orderVal] === undefined) {
        result[orderVal] = {
          model: String(block[r][modelOffset]).trim(),
          managementNumbers: []
        };
      }
      var mgmtVal = String(block[r][mgmtOffset]).trim();
      if (mgmtVal !== '') {
        result[orderVal].managementNumbers.push(mgmtVal);
      }
    }
  }
  return result;
}

/*==================== 出力シート構築 ====================*/

function buildOutputSheet_(records) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 以前のバージョンで作られた「印刷_1」「印刷_2」...のページ分割シートが残っていれば削除する
  var stalePageNum = 1;
  var staleSheet;
  while ((staleSheet = ss.getSheetByName('印刷_' + stalePageNum))) {
    ss.deleteSheet(staleSheet);
    stalePageNum++;
  }

  var sheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (sheet) {
    // シートを削除して作り直すと、Googleスプレッドシートがシートごとに保持している印刷設定
    // (用紙サイズ・余白・拡大縮小など)もリセットされてしまうため、削除・再作成はせず
    // 同じシートの中身（値・書式・画像）だけをクリアして使い回す。
    sheet.clear();
    var oldImages = sheet.getImages();
    for (var oi = 0; oi < oldImages.length; oi++) {
      oldImages[oi].remove();
    }
  } else {
    sheet = ss.insertSheet(OUTPUT_SHEET_NAME);
  }
  sheet.setHiddenGridlines(true); // 印刷時にセルの罫線（グリッド線）が出ないようにする

  // 列幅設定（左端A・中央の隙間・右端の余白を同じ幅に揃える）
  sheet.setColumnWidth(1, A_COL_WIDTH_PX);
  sheet.setColumnWidth(LEFT_COL, TEXT_COL_WIDTH_PX);
  sheet.setColumnWidth(LEFT_COL + 1, GAP_COL_WIDTH_PX);
  sheet.setColumnWidth(RIGHT_COL, TEXT_COL_WIDTH_PX);
  sheet.setColumnWidth(RIGHT_COL + 1, E_COL_WIDTH_PX);

  var rowsPerBlock = BLOCK_ROW_HEIGHTS.length;
  var n = records.length;
  if (n === 0) {
    sheet.setFrozenRows(0);
    return;
  }
  var totalRows = n * rowsPerBlock;

  // --- 行の高さを設定 ---
  // (setRowHeightは1行ずつしか指定できないAPIのため、ここは件数に比例して呼び出すが、
  //  画像転送を伴わない軽い呼び出しなので値・書式のまとめ書きほどのボトルネックにはならない)
  for (var i = 0; i < n; i++) {
    var base = i * rowsPerBlock;
    for (var k = 0; k < rowsPerBlock; k++) {
      sheet.setRowHeight(base + k + 1, BLOCK_ROW_HEIGHTS[k]);
    }
  }

  // --- セルの値・書式は配列にまとめて一括反映（1件ずつ個別にsetValue()等を呼ばない） ---
  // 列は LEFT_COL(B) 〜 RIGHT_COL(D) の3列分（間のギャップ列Cも含む）をまとめて1回のRangeで扱う
  var numCols = RIGHT_COL - LEFT_COL + 1;
  var leftIdx = 0;
  var rightIdx = RIGHT_COL - LEFT_COL;

  var values = [];
  var fontSizes = [];
  var fontWeights = [];
  var hAligns = [];
  var vAligns = [];
  var numberFormats = [];
  for (var r = 0; r < totalRows; r++) {
    values.push(new Array(numCols).fill(''));
    fontSizes.push(new Array(numCols).fill(10));
    fontWeights.push(new Array(numCols).fill('normal'));
    hAligns.push(new Array(numCols).fill('general'));
    vAligns.push(new Array(numCols).fill('bottom'));
    numberFormats.push(new Array(numCols).fill('General'));
  }

  var setCell = function (rowIdx, colIdx, value, fontSize, fontWeight, hAlign, vAlign, numberFormat) {
    values[rowIdx][colIdx] = value;
    fontSizes[rowIdx][colIdx] = fontSize;
    fontWeights[rowIdx][colIdx] = fontWeight;
    hAligns[rowIdx][colIdx] = hAlign;
    vAligns[rowIdx][colIdx] = vAlign;
    numberFormats[rowIdx][colIdx] = numberFormat;
  };

  for (var i2 = 0; i2 < n; i2++) {
    var record = records[i2];
    var base2 = i2 * rowsPerBlock;

    // 行の並び: 0=年月, 1=連番(+QR), 2=空白, 3=型式, 4=空白, 5=注番, 6=空白
    // 年月("2026-07")・連番("001")とも、数値や日付に自動変換されないよう文字列指定('@')にする
    setCell(base2 + 0, leftIdx, record.month, MONTH_FONT_SIZE, 'bold', 'left', 'bottom', '@');
    setCell(base2 + 0, rightIdx, record.month, MONTH_FONT_SIZE, 'bold', 'left', 'bottom', '@');

    // 連番の値・文字サイズはここでは設定しない（後でリッチテキストとして個別に設定するため）。
    // 配置(揃え)・数値書式だけここで反映しておく。
    setCell(base2 + 1, leftIdx, '', SERIAL_FONT_SIZE, 'normal', 'general', 'bottom', '@');
    setCell(base2 + 1, rightIdx, '', SERIAL_FONT_SIZE, 'normal', 'general', 'bottom', '@');

    setCell(base2 + 3, leftIdx, record.model, MODEL_FONT_SIZE, 'normal', 'general', 'bottom', 'General');
    setCell(base2 + 3, rightIdx, record.model, MODEL_FONT_SIZE, 'normal', 'general', 'bottom', 'General');

    setCell(base2 + 5, leftIdx, record.orderNo, ORDER_NO_FONT_SIZE, 'normal', 'general', 'middle', 'General');
    setCell(base2 + 5, rightIdx, record.orderNo, ORDER_NO_FONT_SIZE, 'normal', 'general', 'middle', 'General');
  }

  var range = sheet.getRange(1, LEFT_COL, totalRows, numCols);
  // 先に書式（特に文字列指定'@'）を反映してから値を書き込む。
  // 逆順にすると、"001"や"2026-07"のような数字・日付に見える文字列が
  // 書き込み時点で数値や日付に自動変換されてしまい、後から'@'にしても元に戻らない。
  range.setFontSizes(fontSizes);
  range.setFontWeights(fontWeights);
  range.setHorizontalAlignments(hAligns);
  range.setVerticalAlignments(vAligns);
  range.setNumberFormats(numberFormats);
  range.setWrap(false);
  range.setFontFamily('MS Gothic');
  range.setFontColor('#000000');
  range.setValues(values);

  // --- 連番のセルにリッチテキストを設定（先頭の空白だけ小さいフォントにして約1cm右にずらす） ---
  for (var i4 = 0; i4 < n; i4++) {
    var record4 = records[i4];
    var serialRow4 = i4 * rowsPerBlock + 2; // ブロック内0-indexで1番目(連番行)。シート行は1始まりなので+2
    setSerialRichText_(sheet, serialRow4, LEFT_COL, record4.serial);
    setSerialRichText_(sheet, serialRow4, RIGHT_COL, record4.serial);
  }

  // --- QRコード画像を挿入 ---
  // 左右で同じ内容のため、QR画像の生成は1件につき1回だけ行い、挿入だけ2回行う
  // （画像挿入自体はSheets APIの仕様上1枚ずつしか呼び出せないため、ここは件数に比例する）
  // QRは列の右端ではなく、連番の数字にもっと近いQR_OFFSET_X_PXの位置に置く
  var offsetX = Math.max(QR_OFFSET_X_PX, 0);
  for (var i3 = 0; i3 < n; i3++) {
    var record3 = records[i3];
    var startRow3 = i3 * rowsPerBlock + 1;
    var qrText = record3.serial + '\n' + record3.model + '\n' + record3.month + '\n' + record3.orderNo;
    var blob = generateQrPngBlob(qrText, QR_MODULE_PX);

    insertQrImage_(sheet, startRow3, LEFT_COL, blob, offsetX);
    insertQrImage_(sheet, startRow3, RIGHT_COL, blob, offsetX);
  }

  // 印刷設定: A4横向き、余白なし
  // (Apps Script には Excel のような細かい PageSetup API がないため、
  //  実際の印刷は [ファイル > 印刷] のダイアログで「用紙: A4」「向き: 横」
  //  「余白: なし」を選んで行う)
  sheet.setFrozenRows(0);
}

function insertQrImage_(sheet, startRow, col, blob, offsetX) {
  // QRを連番の行の縦方向中央に来るようQR_OFFSET_Y_PXだけ下げて配置する
  var image = sheet.insertImage(blob, col, startRow + 1, offsetX, QR_OFFSET_Y_PX);
  image.setWidth(QR_DISPLAY_SIZE_PX);
  image.setHeight(QR_DISPLAY_SIZE_PX);
}

// 連番セルに、先頭のSERIAL_LEFT_PADだけSERIAL_LEFT_PAD_FONT_SIZE(小さめ)、
// 本体の連番はSERIAL_FONT_SIZEで表示するリッチテキストを設定する。
// これにより、空白の見た目の横幅を使って連番を約1cm右にずらす。
function setSerialRichText_(sheet, row, col, serial) {
  var text = SERIAL_LEFT_PAD + serial;
  var padLen = SERIAL_LEFT_PAD.length;
  var padStyle = SpreadsheetApp.newTextStyle()
    .setFontSize(SERIAL_LEFT_PAD_FONT_SIZE)
    .setFontFamily('MS Gothic')
    .setForegroundColor('#000000')
    .build();
  var serialStyle = SpreadsheetApp.newTextStyle()
    .setFontSize(SERIAL_FONT_SIZE)
    .setFontFamily('MS Gothic')
    .setForegroundColor('#000000')
    .build();
  var richValue = SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setTextStyle(0, padLen, padStyle)
    .setTextStyle(padLen, text.length, serialStyle)
    .build();
  sheet.getRange(row, col).setRichTextValue(richValue);
}

/*==================== QRコード → PNG画像 生成（外部通信なし） ====================*/
/* qrcodegen.gs の QR Code generator library (Project Nayuki, MIT License) を利用し、
 * QRコードのモジュール配列から白黒PNG画像をバイト列レベルで直接組み立てる。
 * DEFLATE は圧縮を行わない "stored"（無圧縮）ブロックのみを使用するため、
 * 追加ライブラリなしで正しい PNG ファイルを生成できる。
 */

function generateQrPngBlob(text, moduleSizePx) {
  var qr = qrcodegen.QrCode.encodeText(text, qrcodegen.QrCode.Ecc.MEDIUM);
  var quiet = 4; // QRコード仕様上必要なクワイエットゾーン(モジュール数)
  var modules = qr.size;
  var totalModules = modules + quiet * 2;
  var size = totalModules * moduleSizePx;

  var pngBytes = encodePngGrayscale_(size, size, function (x, y) {
    var mx = Math.floor(x / moduleSizePx) - quiet;
    var my = Math.floor(y / moduleSizePx) - quiet;
    if (mx < 0 || my < 0 || mx >= modules || my >= modules) return 255;
    return qr.getModule(mx, my) ? 0 : 255;
  });

  return Utilities.newBlob(bytesToSignedInt8Array_(pngBytes), 'image/png', 'qr.png');
}

function encodePngGrayscale_(width, height, getPixel) {
  var sig = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
  var ihdr = [
    (width >>> 24) & 0xFF, (width >>> 16) & 0xFF, (width >>> 8) & 0xFF, width & 0xFF,
    (height >>> 24) & 0xFF, (height >>> 16) & 0xFF, (height >>> 8) & 0xFF, height & 0xFF,
    8, 0, 0, 0, 0 // bit depth 8, color type 0 (grayscale), compression/filter/interlace = 0
  ];

  var raw = [];
  for (var y = 0; y < height; y++) {
    raw.push(0); // フィルタタイプ: None
    for (var x = 0; x < width; x++) {
      raw.push(getPixel(x, y));
    }
  }

  var idatData = zlibWrap_(raw);

  var out = sig.slice();
  out = out.concat(pngChunk_(strToBytes_('IHDR'), ihdr));
  out = out.concat(pngChunk_(strToBytes_('IDAT'), idatData));
  out = out.concat(pngChunk_(strToBytes_('IEND'), []));
  return out;
}

function pngChunk_(tagBytes, data) {
  var out = [];
  var len = data.length;
  out.push((len >>> 24) & 0xFF, (len >>> 16) & 0xFF, (len >>> 8) & 0xFF, len & 0xFF);
  var tagAndData = tagBytes.concat(data);
  out = out.concat(tagAndData);
  var crc = crc32_(tagAndData);
  out.push((crc >>> 24) & 0xFF, (crc >>> 16) & 0xFF, (crc >>> 8) & 0xFF, crc & 0xFF);
  return out;
}

function zlibWrap_(rawData) {
  var cmf = 0x78; // 32K window, deflate method
  var flg = 0;
  for (var candidate = 0; candidate < 256; candidate++) {
    // FLG は (CMF*256+FLG) が 31 の倍数になる必要がある。FDICTビット(0x20)は立てない。
    if (((cmf << 8) + candidate) % 31 === 0 && (candidate & 0x20) === 0) {
      flg = candidate;
      break;
    }
  }
  var out = [cmf, flg];
  out = out.concat(storedDeflate_(rawData));
  var adler = adler32_(rawData);
  out.push((adler >>> 24) & 0xFF, (adler >>> 16) & 0xFF, (adler >>> 8) & 0xFF, adler & 0xFF);
  return out;
}

function storedDeflate_(data) {
  // DEFLATE の "stored"（無圧縮）ブロックのみで構成する。
  // 各ブロックは最大65535バイトまで。
  var out = [];
  var n = data.length;
  if (n === 0) {
    out.push(0x01, 0x00, 0x00, 0xFF, 0xFF);
    return out;
  }
  var i = 0;
  while (i < n) {
    var len = Math.min(65535, n - i);
    var isFinal = (i + len) >= n;
    out.push(isFinal ? 0x01 : 0x00);
    out.push(len & 0xFF, (len >>> 8) & 0xFF);
    var nlen = (~len) & 0xFFFF;
    out.push(nlen & 0xFF, (nlen >>> 8) & 0xFF);
    for (var j = 0; j < len; j++) out.push(data[i + j]);
    i += len;
  }
  return out;
}

var CRC_TABLE_ = null;
function crcTable_() {
  if (CRC_TABLE_) return CRC_TABLE_;
  var table = [];
  for (var n = 0; n < 256; n++) {
    var c = n;
    for (var k = 0; k < 8; k++) {
      c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[n] = c >>> 0;
  }
  CRC_TABLE_ = table;
  return table;
}

function crc32_(bytes) {
  var table = crcTable_();
  var crc = 0xFFFFFFFF;
  for (var i = 0; i < bytes.length; i++) {
    crc = table[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function adler32_(bytes) {
  var MOD = 65521;
  var a = 1, b = 0;
  for (var i = 0; i < bytes.length; i++) {
    a = (a + bytes[i]) % MOD;
    b = (b + a) % MOD;
  }
  return ((b << 16) | a) >>> 0;
}

function strToBytes_(s) {
  var arr = [];
  for (var i = 0; i < s.length; i++) arr.push(s.charCodeAt(i));
  return arr;
}

function bytesToSignedInt8Array_(bytes) {
  var out = [];
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i];
    out.push(b > 127 ? b - 256 : b);
  }
  return out;
}
