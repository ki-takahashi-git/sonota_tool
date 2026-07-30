/*
 * 現品票 自動作成スクリプト（Google Slides版）
 *
 * 「内外エレクトロニクス生産計画」から注番・型式を自動取得し、
 * QRコード付きの現品票をGoogleスライドに自動レイアウトしてPDFを直接書き出す。
 * Sheets版（apps_script/Code.gs）との違い:
 *   ・文字要素はすべて Slides.Presentations.batchUpdate に積んで
 *     数回のAPI呼び出しにまとめて送るため、件数が増えても大きく遅くならない
 *   ・生成後、印刷ダイアログを毎回手で調整する必要がなく、印刷用PDFを直接書き出す
 * QRコード生成は外部通信を一切行わず、このスクリプト内で完結する
 * （qrcodegen.gs に同梱した QR Code generator library (Project Nayuki, MIT License) を使用）。
 *
 * ============ 事前準備（初回のみ） ============
 *   1. 空のGoogleスライドを1つ新規作成する
 *   2. ファイル > ページ設定 > カスタム を選び、単位を「cm」にして
 *      幅29.7 × 高さ21.0（A4横向き）を設定して保存する
 *   3. そのファイルのURLからファイルIDを取得し、下の TEMPLATE_PRESENTATION_ID に設定する
 *   4. このスクリプトを紐づけたいスプレッドシートを開き、拡張機能 > Apps Script を開く
 *   5. サービスの「+」から「Google Slides API」（Advanced Service）を追加する
 *
 * ============ 使い方 ============
 *   1. 「入力」シートのA列（2行目以降）に、現品票を作りたい注番を1行ずつ入力する
 *   2. スプレッドシート上部メニュー「現品票」>「入力した注番から作成」を実行する
 *   3. 完了メッセージに表示されるPDFのURLを開いて内容を確認し、そのまま印刷する
 */

/*==================== 設定 ====================*/

// コピー元「内外エレクトロニクス生産計画」のスプレッドシートID（URLの /d/ と /edit の間の文字列）
var SOURCE_SPREADSHEET_ID = '1VQoAyx76euNgH7DsQWljKPyp1oQbmQqjr-FWUL2-wEU';

// コピー元シートのgid（URLの #gid=... の数字）
var SOURCE_SHEET_GID = 332550892;

// コピー元シートの列番号（1=A, 2=B, 3=C, 4=D, 5=E ...）
var SOURCE_COL_ORDER_NO = 2;  // B列: 注番
var SOURCE_COL_MODEL = 4;     // D列: 型式
var SOURCE_COL_QUANTITY = 5;  // E列: 数量

// このスプレッドシート内のシート名
var INPUT_SHEET_NAME = '入力';

// 事前準備で作成した「A4横向き」テンプレートスライドのファイルID（事前準備の手順3を参照）
var TEMPLATE_PRESENTATION_ID = 'ここにテンプレートのファイルIDを入力';

// 生成したPDF/スライドの保存先フォルダID（空文字ならマイドライブ直下に保存）
var OUTPUT_FOLDER_ID = '';

// Slides.Presentations.batchUpdate 1回あたりの最大リクエスト数（安全マージン。多すぎる場合は分割送信する）
var REQUESTS_PER_BATCH = 300;

// 現品票レイアウト設定（単位はpt = 1/72インチ。すべて実寸で配置されるため、印刷ダイアログでの調整は不要）
// 1ブロック=7行の各行の高さ(pt)。内訳: 0=年月, 1=連番(+QR), 2=空白, 3=型式, 4=空白, 5=注番, 6=空白
var BLOCK_ROW_HEIGHTS_PT = [24, 130, 10, 20, 10, 65, 20];
var TAG_WIDTH_PT = 390;   // 現品票1枚分の幅(pt)
var GAP_PT = 15;          // 左右の現品票の間の余白(pt)
var QR_SIZE_PT = 90;      // スライド上でのQR表示サイズ(pt)
var QR_MODULE_PX = 8;     // QR 1モジュールあたりの生成解像度(px)。画質を上げたい場合は増やす
var SERIAL_FONT_SIZE = 80;    // 連番のフォントサイズ(pt)
var ORDER_NO_FONT_SIZE = 44;  // 注番のフォントサイズ(pt)
var MONTH_FONT_SIZE = 10;     // 年月のフォントサイズ(pt)
var MODEL_FONT_SIZE = 11;     // 型式のフォントサイズ(pt)
var FONT_FAMILY = 'MS Gothic';

/*==================== メニュー ====================*/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('現品票')
    .addItem('入力した注番から作成', 'generateTags')
    .addSeparator()
    .addItem('今月の連番をリセット', 'resetSerialForCurrentMonth')
    .addToUi();
}

// 今月分の連番カウンタを0に戻す（次回作成時に001から始まる）。
// テストで進んだカウンタを本番運用前に仕切り直したいときに使う。
function resetSerialForCurrentMonth() {
  var ui = SpreadsheetApp.getUi();
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var monthKey = Utilities.formatDate(new Date(), tz, 'yyyyMM');
  var props = PropertiesService.getScriptProperties();
  var key = 'serial_' + monthKey;
  var current = props.getProperty(key) || '0';

  var response = ui.alert(
    '連番リセット確認',
    '今月(' + monthKey + ')の連番カウンタを0にリセットします。\n現在の値: ' + current +
      '\n次回作成時は001から始まります。既に現品票を発行済みの場合は番号が重複する可能性があります。よろしいですか？',
    ui.ButtonSet.YES_NO
  );
  if (response == ui.Button.YES) {
    props.deleteProperty(key);
    ui.alert('リセットしました。次回作成時は001から始まります。');
  }
}

/*==================== メイン処理 ====================*/

function generateTags() {
  var ui = SpreadsheetApp.getUi();
  var orderNumbers = readInputOrderNumbers_();
  if (orderNumbers.length === 0) {
    ui.alert('「' + INPUT_SHEET_NAME + '」シートのA列（2行目以降）に注番を入力してください。');
    return;
  }

  var lookup = lookupOrdersFromSource_(orderNumbers);
  var now = new Date();
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var monthLabel = Utilities.formatDate(now, tz, 'yyyy-MM');
  var monthKey = Utilities.formatDate(now, tz, 'yyyyMM');

  var records = [];
  var notFound = [];
  for (var i = 0; i < orderNumbers.length; i++) {
    var orderNo = orderNumbers[i];
    var info = lookup[orderNo];
    if (info === undefined) {
      notFound.push(orderNo);
      continue;
    }
    // 数量が2以上の場合は、数量分だけ現品票を分けて作成し、
    // 注番に "-1/2", "-2/2" のような通し番号を付ける（数量1の場合は付けない）
    var qty = info.quantity > 0 ? info.quantity : 1;
    for (var unit = 1; unit <= qty; unit++) {
      var serial = getNextSerialForMonth_(monthKey);
      var displayOrderNo = qty > 1 ? (orderNo + '-' + unit + '/' + qty) : orderNo;
      records.push({
        serial: serial,
        model: info.model,
        month: monthLabel,
        orderNo: displayOrderNo
      });
    }
  }

  var presentationId = null;
  if (records.length > 0) {
    presentationId = buildSlidesDeck_(records);
  }

  var msg = records.length + ' 件の現品票を作成しました。';
  if (presentationId) {
    var pdfInfo = exportToPdf_(presentationId);
    msg += '\n\nPDF: ' + pdfInfo.pdfUrl;
    msg += '\nスライド（内容確認用）: ' + pdfInfo.slideUrl;
  }
  if (notFound.length > 0) {
    msg += '\n\n生産計画シート内に見つからなかった注番:\n' + notFound.join('\n');
  }
  ui.alert(msg);
}

/*==================== データ取得 ====================*/

function readInputOrderNumbers_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(INPUT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(INPUT_SHEET_NAME);
    sheet.getRange('A1').setValue('注番');
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var result = [];
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0]).trim();
    if (v !== '') result.push(v);
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
  // 注番・型式・数量の列をまとめて1回で読み込む（列がB,D,Eで連続していないため、
  // B〜Eをまとめて読み、使わないC列は無視する）
  var firstCol = Math.min(SOURCE_COL_ORDER_NO, SOURCE_COL_MODEL, SOURCE_COL_QUANTITY);
  var lastCol = Math.max(SOURCE_COL_ORDER_NO, SOURCE_COL_MODEL, SOURCE_COL_QUANTITY);
  var block = sheet.getRange(1, firstCol, lastRow, lastCol - firstCol + 1).getValues();
  var orderOffset = SOURCE_COL_ORDER_NO - firstCol;
  var modelOffset = SOURCE_COL_MODEL - firstCol;
  var qtyOffset = SOURCE_COL_QUANTITY - firstCol;

  var result = {};
  for (var r = 0; r < lastRow; r++) {
    var orderVal = String(block[r][orderOffset]).trim();
    if (orderVal !== '' && wanted[orderVal] && result[orderVal] === undefined) {
      var qtyRaw = block[r][qtyOffset];
      var qty = parseInt(qtyRaw, 10);
      result[orderVal] = {
        model: String(block[r][modelOffset]).trim(),
        quantity: isNaN(qty) ? 1 : qty
      };
    }
  }
  return result;
}

function getNextSerialForMonth_(monthKey) {
  var props = PropertiesService.getScriptProperties();
  var key = 'serial_' + monthKey;
  var current = parseInt(props.getProperty(key) || '0', 10);
  current += 1;
  props.setProperty(key, String(current));
  var padded = ('000' + current).slice(-3);
  return padded;
}

/*==================== スライド構築 ====================*/

function buildSlidesDeck_(records) {
  if (!TEMPLATE_PRESENTATION_ID || TEMPLATE_PRESENTATION_ID.indexOf('ここに') === 0) {
    throw new Error(
      'TEMPLATE_PRESENTATION_ID が未設定です。事前準備の手順に従ってA4横向きのテンプレートを作成し、\n' +
      'そのファイルIDをスクリプト先頭の TEMPLATE_PRESENTATION_ID に設定してください。'
    );
  }

  var tz = Session.getScriptTimeZone();
  var copyName = '現品票_' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
  var templateFile = DriveApp.getFileById(TEMPLATE_PRESENTATION_ID);
  var newFile = templateFile.makeCopy(copyName);
  if (OUTPUT_FOLDER_ID) {
    var folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    folder.addFile(newFile);
    DriveApp.getRootFolder().removeFile(newFile);
  }
  var presentationId = newFile.getId();

  var presentation = SlidesApp.openById(presentationId);
  var pageWidthPt = presentation.getPageWidth();
  var pageHeightPt = presentation.getPageHeight();
  var existingSlides = presentation.getSlides();

  var blockHeightPt = sumArray_(BLOCK_ROW_HEIGHTS_PT);
  var blocksPerPage = Math.max(1, Math.floor(pageHeightPt / blockHeightPt));
  var totalBlockHeightPt = blocksPerPage * blockHeightPt;
  var topMarginPt = Math.max((pageHeightPt - totalBlockHeightPt) / 2, 0);

  var totalWidthPt = TAG_WIDTH_PT * 2 + GAP_PT;
  var leftMarginPt = Math.max((pageWidthPt - totalWidthPt) / 2, 0);
  var leftX = leftMarginPt;
  var rightX = leftMarginPt + TAG_WIDTH_PT + GAP_PT;

  var numPages = Math.ceil(records.length / blocksPerPage);
  var rowOffsets = computeRowOffsets_(BLOCK_ROW_HEIGHTS_PT);

  // --- スライドページを必要枚数ぶん用意する ---
  // テンプレートに最初から入っている1枚目はそのまま1ページ目として使い、
  // 2ページ目以降だけ新規作成する（IDは自己採番せず、batchUpdateの返信(replies)から
  // 実際に作成されたobjectIdを取得する。これにより、あとから参照するIDが必ず実在するIDと一致する）
  var slideIds = [existingSlides[0].getObjectId()];

  var createSlideRequests = [];
  for (var p = 1; p < numPages; p++) {
    createSlideRequests.push({
      createSlide: { slideLayoutReference: { predefinedLayout: 'BLANK' } }
    });
  }
  // テンプレートに1枚目以外の余分なスライドが残っていた場合は削除する
  var deleteRequests = [];
  for (var e = 1; e < existingSlides.length; e++) {
    deleteRequests.push({ deleteObject: { objectId: existingSlides[e].getObjectId() } });
  }
  if (createSlideRequests.length > 0 || deleteRequests.length > 0) {
    var slideResponse = Slides.Presentations.batchUpdate(
      { requests: createSlideRequests.concat(deleteRequests) },
      presentationId
    );
    for (var r = 0; r < createSlideRequests.length; r++) {
      slideIds.push(slideResponse.replies[r].createSlide.objectId);
    }
  }

  // --- テキスト要素をまとめてbatchUpdateに積む ---
  var requests = [];
  for (var i = 0; i < records.length; i++) {
    var record = records[i];
    var pageIndex = Math.floor(i / blocksPerPage);
    var blockIndex = i % blocksPerPage;
    var slideId = slideIds[pageIndex];
    var blockTopY = topMarginPt + blockIndex * blockHeightPt;

    appendTagTextRequests_(requests, slideId, leftX, blockTopY, rowOffsets, i, 'L', record);
    appendTagTextRequests_(requests, slideId, rightX, blockTopY, rowOffsets, i, 'R', record);
  }

  runBatchUpdateChunked_(presentationId, requests);

  // --- QRコード画像はSlidesAppで直接挿入（左右で同じ内容のため生成は1回に抑える） ---
  presentation = SlidesApp.openById(presentationId);
  var slidesById = {};
  var slides = presentation.getSlides();
  for (var s = 0; s < slides.length; s++) {
    slidesById[slides[s].getObjectId()] = slides[s];
  }

  var qrOffsetXWithinTag = Math.max(TAG_WIDTH_PT - QR_SIZE_PT - 5, 0);
  for (var i2 = 0; i2 < records.length; i2++) {
    var record2 = records[i2];
    var pageIndex2 = Math.floor(i2 / blocksPerPage);
    var blockIndex2 = i2 % blocksPerPage;
    var slide = slidesById[slideIds[pageIndex2]];
    var blockTopY2 = topMarginPt + blockIndex2 * blockHeightPt;
    var qrY = blockTopY2 + rowOffsets[1];

    var qrText = record2.serial + '\n' + record2.model + '\n' + record2.month + '\n' + record2.orderNo;
    var blob = generateQrPngBlob(qrText, QR_MODULE_PX);

    slide.insertImage(blob, leftX + qrOffsetXWithinTag, qrY, QR_SIZE_PT, QR_SIZE_PT);
    slide.insertImage(blob, rightX + qrOffsetXWithinTag, qrY, QR_SIZE_PT, QR_SIZE_PT);
  }

  return presentationId;
}

// 1ブロック内の各行(0〜6)の「ブロック先頭からの相対Y座標(pt)」を前方累積で求める
// 例: BLOCK_ROW_HEIGHTS_PT=[24,130,...] なら rowOffsets=[0,24,154,...]
function computeRowOffsets_(heights) {
  var offsets = [0];
  var acc = 0;
  for (var i = 0; i < heights.length; i++) {
    acc += heights[i];
    offsets.push(acc);
  }
  return offsets;
}

function sumArray_(arr) {
  var total = 0;
  for (var i = 0; i < arr.length; i++) total += arr[i];
  return total;
}

// 1件分（月・連番・型式・注番の4テキストボックス）のcreateShape/insertText/updateTextStyle
// リクエストをrequests配列に積む。実際のAPI呼び出しはbuildSlidesDeck_側でまとめて行う。
function appendTagTextRequests_(requests, slideId, x, blockTopY, rowOffsets, recordIndex, side, record) {
  var fields = [
    { key: 'month', rowIndex: 0, fontSize: MONTH_FONT_SIZE, value: record.month },
    { key: 'serial', rowIndex: 1, fontSize: SERIAL_FONT_SIZE, value: record.serial },
    { key: 'model', rowIndex: 3, fontSize: MODEL_FONT_SIZE, value: record.model },
    { key: 'orderNo', rowIndex: 5, fontSize: ORDER_NO_FONT_SIZE, value: record.orderNo }
  ];

  for (var f = 0; f < fields.length; f++) {
    var field = fields[f];
    var objectId = 'tb_' + recordIndex + '_' + side + '_' + field.key;
    var y = blockTopY + rowOffsets[field.rowIndex];
    var heightPt = rowOffsets[field.rowIndex + 1] - rowOffsets[field.rowIndex];

    requests.push({
      createShape: {
        objectId: objectId,
        shapeType: 'TEXT_BOX',
        elementProperties: {
          pageObjectId: slideId,
          size: {
            width: { magnitude: TAG_WIDTH_PT, unit: 'PT' },
            height: { magnitude: heightPt, unit: 'PT' }
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: x,
            translateY: y,
            unit: 'PT'
          }
        }
      }
    });
    requests.push({
      insertText: {
        objectId: objectId,
        text: String(field.value),
        insertionIndex: 0
      }
    });
    requests.push({
      updateTextStyle: {
        objectId: objectId,
        style: {
          fontSize: { magnitude: field.fontSize, unit: 'PT' },
          fontFamily: FONT_FAMILY,
          foregroundColor: { opaqueColor: { rgbColor: { red: 0, green: 0, blue: 0 } } }
        },
        textRange: { type: 'ALL' },
        fields: 'fontSize,fontFamily,foregroundColor'
      }
    });
  }
}

// Slides.Presentations.batchUpdate は1回の呼び出しに全リクエストをまとめるのが理想だが、
// 件数が非常に多い場合の安全マージンとしてREQUESTS_PER_BATCH件ずつに分割して送信する。
function runBatchUpdateChunked_(presentationId, requests) {
  for (var i = 0; i < requests.length; i += REQUESTS_PER_BATCH) {
    var chunk = requests.slice(i, i + REQUESTS_PER_BATCH);
    Slides.Presentations.batchUpdate({ requests: chunk }, presentationId);
  }
}

/*==================== PDF書き出し ====================*/

function exportToPdf_(presentationId) {
  var file = DriveApp.getFileById(presentationId);
  var pdfBlob = file.getAs('application/pdf');
  var pdfFile = DriveApp.createFile(pdfBlob).setName(file.getName() + '.pdf');
  if (OUTPUT_FOLDER_ID) {
    var folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    folder.addFile(pdfFile);
    DriveApp.getRootFolder().removeFile(pdfFile);
  }
  return {
    pdfUrl: pdfFile.getUrl(),
    slideUrl: file.getUrl()
  };
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
