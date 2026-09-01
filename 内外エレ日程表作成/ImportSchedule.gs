/**
 * ============================================================================
 * 内外エレクトロニクス日程表 取込スクリプト
 * ============================================================================
 *
 * 【使い方】
 * 1. このスプレッドシート(検証用コピー)の「拡張機能 > Apps Script」を開く
 * 2. デフォルトの Code.gs の中身をすべて削除し、このファイルの内容を貼り付けて保存
 * 3. スプレッドシートをリロード(再読み込み)すると、メニューバーに
 *    「日程表取込」というメニューが追加される
 * 4. 得意先から届いた「タカハタ電子日程表YYYYMMDD.xlsx」を、このスプレッドシートに
 *    ファイル > インポート > アップロード > 「新しいシートとして挿入」で取り込む
 *    (シート名が自動で「タカハタ電子様」等になる)
 * 5. メニュー「日程表取込 > 取り込み実行」をクリック
 * 6. 追加件数の確認ダイアログが出るので、内容を確認してから「OK」
 *
 * 【前提・仕様】(2026/08/24 打ち合わせに基づく)
 * ・新規判定キー: 「注番」+「NEI注文番号」の組み合わせが生産計画シートに
 *   まだ1件も無ければ「新規注文」とみなす
 * ・Excelの「部材支給日」(G列)・「納期」(H列)の両方に記載がある行だけを転記対象にする
 *   (どちらかが空欄の行はスキップする)
 * ・新規注文は Excel の「数量」の数だけ、同じ内容の行を複製して1行=1台で追加する
 * ・型式はExcelの表記をそのまま転記する(変換ロジックは入れていない)
 * ・工数・LT・想定納期など、生産計画シート側で数式になっている列は
 *   直近行の数式をコピーして自動計算に任せる(値は書き込まない)
 * ・管理No.(L列)は新規行では空欄にする
 * ・型式は転記時に、半角スペースが2つ以上連続していたら1つにまとめる
 * ・ガントチャート部分(PO催促列より右の日付ごとのマス目)は、新規行では
 *   テンプレート行の値(「1」等)をクリアし(罫線は残す)、背景色は
 *   見出し行(月火水木金土日が書かれている行)の列ごとの色をそのまま反映する
 *   (土日だけ灰色、という決め打ちはせず、見出し行の色に合わせる)。
 *   部材支給/製造納期列(F・G列)のオレンジ・青の色はこの範囲の外なので触らない
 * ・注文日(A列)は「スプレッドシートに登録した日」= このスクリプトを実行した日
 *   (今日の日付)を入れる。Excel側の日付列は使わない
 * ・新規注文は「製造納期」を見て、既存データの中から同じ製造納期の行(無ければ
 *   直近で近い製造納期の行)を探し、そのかたまりの一番下に挿入する
 *   (末尾に一括で追加するわけではない。既存データはシート全体で製造納期の
 *   昇順とは限らない前提で、シート全体を探索して該当箇所を探す)
 * ・同じNEI注文番号の行が複数(数量分)並ぶグループには、NEI注文番号・型式の列に
 *   太い罫線で囲みを付ける
 * ・A列(注文日)の色付けは実行のたびにリセットする:
 *   まず生産計画シート全体のA列の色をクリアしてから、今回追加した行のA列だけ
 *   赤色で塗る(「今回の追加分がひと目でわかる」ようにするマーキング)
 *
 * 【まだ人の目で確認してほしい点】
 * ・型式の表記がExcelと生産計画シートで食い違うケースがあると、工数の自動計算式
 *   (型式で工数単価表を参照している場合)が空欄/エラーになることがあります。
 *   取込後、工数列が空欄の行がないか一度目視確認してください。
 * ・赤色のコード(CONFIG.NEW_ROW_HIGHLIGHT_COLOR)が実際の運用の色と違ったら
 *   直してください。
 * ・注文日がずれる場合は、スプレッドシートのタイムゾーン設定
 *   (ファイル > 設定 > 地域と表示形式の設定)を確認してください。
 *   このスクリプトはスプレッドシート自身のタイムゾーン設定に合わせて
 *   「今日の日付」を計算しています。
 * ・ガントチャート部分の色がうまく引き継がれない場合は、CONFIG.WEEKDAY_LABEL_ROW
 *   (「月火水木金土日」が書かれている行番号)が実際のシートと合っているか
 *   確認してください。
 * ============================================================================
 */

const CONFIG = {
  // 生産計画シート(本体)のシート名
  MAIN_SHEET_NAME: '生産計画',
  // 本体シートのヘッダー行番号(この行にある文字列で列を特定する)
  MAIN_HEADER_ROW: 4,
  // 本体シートで実データが始まる行
  MAIN_DATA_START_ROW: 5,
  // ガントチャート部分の「月火水木金土日」が書かれている行(土日の列を特定するため)
  WEEKDAY_LABEL_ROW: 3,

  // 取り込み元(Excelをインポートしてできるシート)の名前の候補
  // 実際にインポートしてできたシート名がこの中になければ、
  // 実行時にシートを選ぶダイアログが出る
  SOURCE_SHEET_CANDIDATES: ['タカハタ電子様'],
  SOURCE_HEADER_ROW: 1,
  SOURCE_DATA_START_ROW: 2,

  // ---- 本体シート側のヘッダー名(生産計画シートのD4セル等の実際の文字列と
  //      一致させること。見出しが変わったらここを直す) ----
  MAIN_HEADERS: {
    orderDate: '注文日',
    orderNo: '注番',
    neiOrderNo: 'NEI注文番号',
    model: '型式',
    materialSupply: '部材支給',
    dueDate: '製造納期',
    managementNo: '管理No.',
    // ガントチャート(日付ごとのマス目)の直前の見出し。この列より右は
    // 新規行では何もコピーせず、まるごと空欄にクリアする
    poReminder: 'PO催促',
  },

  // ---- Excel側のヘッダー名 ----
  SOURCE_HEADERS: {
    orderNo: '注文No',
    neiOrderNo: 'NEI注文番号',
    model: '型式',
    quantity: '数量',
    materialSupplyDate: '部材支給日',
    dueDate: '納期',
  },

  // 今回追加した行のA列(注文日)に付ける色
  NEW_ROW_HIGHLIGHT_COLOR: '#ff0000',
  // NEI注文番号ごとのグループを囲む罫線の色
  GROUP_BORDER_COLOR: '#000000',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日程表取込')
    .addItem('取り込み実行', 'importFromExcelSheet')
    .addToUi();
}

/**
 * メイン処理
 */
function importFromExcelSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const mainSheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  if (!mainSheet) {
    ui.alert(`シート「${CONFIG.MAIN_SHEET_NAME}」が見つかりません。`);
    return;
  }

  const sourceSheet = pickSourceSheet(ss, ui);
  if (!sourceSheet) return; // キャンセル

  // --- 列位置の特定 ---
  const mainCols = getHeaderColumnMap(mainSheet, CONFIG.MAIN_HEADER_ROW, CONFIG.MAIN_HEADERS);
  const srcCols = getHeaderColumnMap(sourceSheet, CONFIG.SOURCE_HEADER_ROW, CONFIG.SOURCE_HEADERS);

  const missingMain = Object.entries(mainCols).filter(([, v]) => v === -1).map(([k]) => k);
  const missingSrc = Object.entries(srcCols).filter(([, v]) => v === -1).map(([k]) => k);
  if (missingMain.length || missingSrc.length) {
    ui.alert(
      '見出し列が見つかりませんでした。CONFIGの見出し名を確認してください。\n' +
      (missingMain.length ? `本体シート側: ${missingMain.join(', ')}\n` : '') +
      (missingSrc.length ? `取込元シート側: ${missingSrc.join(', ')}` : '')
    );
    return;
  }

  // --- 既存キー(注番+NEI注文番号)の集合を作る ---
  const existingKeys = buildExistingKeySet(mainSheet, mainCols);

  // --- 取込元シートを読み込み、新規注文を抽出 ---
  const srcLastRow = sourceSheet.getLastRow();
  if (srcLastRow < CONFIG.SOURCE_DATA_START_ROW) {
    ui.alert('取込元シートにデータがありません。');
    return;
  }
  const srcValues = sourceSheet
    .getRange(CONFIG.SOURCE_DATA_START_ROW, 1, srcLastRow - CONFIG.SOURCE_DATA_START_ROW + 1, sourceSheet.getLastColumn())
    .getValues();

  const newOrders = []; // { orderNo, neiOrderNo, model, materialSupplyDate, dueDate, orderDate, quantity }
  let skippedExisting = 0;
  let skippedBlank = 0;
  let skippedIncompleteDates = 0;

  srcValues.forEach((row) => {
    const orderNo = row[srcCols.orderNo];
    const neiOrderNo = row[srcCols.neiOrderNo];
    if (!orderNo || !neiOrderNo) {
      skippedBlank++;
      return;
    }
    // 部材支給日(G列)・納期(H列)の両方に記載がある行だけを転記対象にする
    const materialSupplyDate = row[srcCols.materialSupplyDate];
    const dueDate = row[srcCols.dueDate];
    if (!materialSupplyDate || !dueDate) {
      skippedIncompleteDates++;
      return;
    }
    const key = makeKey(orderNo, neiOrderNo);
    if (existingKeys.has(key)) {
      skippedExisting++;
      return;
    }
    const quantity = Number(row[srcCols.quantity]) || 1;
    newOrders.push({
      orderNo: orderNo,
      neiOrderNo: neiOrderNo,
      model: row[srcCols.model],
      materialSupplyDate: materialSupplyDate,
      dueDate: dueDate,
      quantity: quantity,
    });
    // 同じキーが取込元シート内で複数回登場しないように既存キー集合にも足しておく
    existingKeys.add(key);
  });

  const totalRowsToAdd = newOrders.reduce((sum, o) => sum + o.quantity, 0);

  if (newOrders.length === 0) {
    ui.alert(
      `新規に追加する注文はありませんでした。\n` +
      `(既存スキップ: ${skippedExisting}件 / 注番かNEI注文番号が空欄でスキップ: ${skippedBlank}件 / ` +
      `部材支給日か納期が空欄でスキップ: ${skippedIncompleteDates}件)`
    );
    return;
  }

  const confirmMsg =
    `新規注文: ${newOrders.length}件\n` +
    `追加される行数(数量分の合計): ${totalRowsToAdd}行\n` +
    `既存につきスキップ: ${skippedExisting}件\n` +
    `注番/NEI注文番号が空欄でスキップ: ${skippedBlank}件\n` +
    `部材支給日/納期のどちらかが空欄でスキップ: ${skippedIncompleteDates}件\n\n` +
    `新規注文は、シート内で同じ製造納期の行(無ければ直近の製造納期の行)を探し、\n` +
    `そのかたまりの一番下に挿入します(末尾への一括追加ではありません)。\n` +
    `注文日(A列)には今日の日付を入れ、今回追加した行だけ赤色で塗ります\n` +
    `(既存の赤色は一度クリアされます)。\n\n` +
    `この内容で「${CONFIG.MAIN_SHEET_NAME}」シートに追加してよろしいですか?`;

  const resp = ui.alert('取り込み確認', confirmMsg, ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  appendOrders(mainSheet, mainCols, newOrders);

  ui.alert(`完了しました。${newOrders.length}件(${totalRowsToAdd}行)を追加しました。`);
}

/** ヘッダー名 -> 列インデックス(0始まり)のマップを作る */
function getHeaderColumnMap(sheet, headerRow, headerDefs) {
  const lastCol = sheet.getLastColumn();
  const headerValues = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const normalized = headerValues.map((h) => String(h).replace(/\s|\n/g, ''));

  const result = {};
  Object.entries(headerDefs).forEach(([key, headerText]) => {
    const target = String(headerText).replace(/\s|\n/g, '');
    result[key] = normalized.indexOf(target);
  });
  return result;
}

/** 本体シートに既にある「注番+NEI注文番号」の組み合わせをSetで返す */
function buildExistingKeySet(mainSheet, mainCols) {
  const lastRow = mainSheet.getLastRow();
  const set = new Set();
  if (lastRow < CONFIG.MAIN_DATA_START_ROW) return set;

  const numRows = lastRow - CONFIG.MAIN_DATA_START_ROW + 1;
  const orderNoCol = mainCols.orderNo + 1; // 1始まりに変換
  const neiCol = mainCols.neiOrderNo + 1;

  const orderNos = mainSheet.getRange(CONFIG.MAIN_DATA_START_ROW, orderNoCol, numRows, 1).getValues();
  const neiNos = mainSheet.getRange(CONFIG.MAIN_DATA_START_ROW, neiCol, numRows, 1).getValues();

  for (let i = 0; i < numRows; i++) {
    const o = orderNos[i][0];
    const n = neiNos[i][0];
    if (o && n) set.add(makeKey(o, n));
  }
  return set;
}

function makeKey(orderNo, neiOrderNo) {
  return `${String(orderNo).trim()}|${String(neiOrderNo).trim()}`;
}

/**
 * 新規注文を、既存データの製造納期を見て正しい位置に挿入する。
 * ・既存データ全体は必ずしも製造納期の昇順に並んでいるとは限らない
 *   (顧客/出荷ロットなどでブロック分けされている)ため、各新規注文ごとに
 *   シート全体を独立して探索し、「同じ製造納期の既存行のかたまりの一番下」
 *   (完全一致が無ければ、それ以下で一番近い製造納期の行の直後)に挿入する
 * ・数量分の行を複製し、挿入位置の直前行から書式・数式をコピーする
 * ・NEI注文番号ごとのグループ(数量分の行のかたまり)に太い罫線で囲みを付ける
 * ・注文日(A列)には実行日を入れ、A列の色は一旦全クリアしてから
 *   今回追加した行だけ赤色にする
 */
function appendOrders(mainSheet, mainCols, newOrders) {
  const lastCol = mainSheet.getLastColumn();
  const today = getTodayInSpreadsheetTimeZone(mainSheet);
  const dueDateCol = mainCols.dueDate + 1;

  // ガントチャート部分(PO催促より右)の背景色は、見出し行(「月火水木金土日」の行)の
  // 列ごとの色をそのまま引き継ぐ(土日だけ灰色、という決め打ちはせず、見出し行の
  // 色をそのまま反映する)
  const calendarStartCol = mainCols.poReminder + 2; // 1始まり、PO催促の次の列
  const calendarWidth = Math.max(0, lastCol - calendarStartCol + 1);
  let headerRowBackgrounds = [];
  if (calendarWidth > 0) {
    headerRowBackgrounds = mainSheet
      .getRange(CONFIG.WEEKDAY_LABEL_ROW, calendarStartCol, 1, calendarWidth)
      .getBackgrounds()[0];
  }

  // 既存データの製造納期を(挿入前の状態で)読み込んでおく
  const initialLastRow = mainSheet.getLastRow();
  const existing = [];
  if (initialLastRow >= CONFIG.MAIN_DATA_START_ROW) {
    const values = mainSheet
      .getRange(CONFIG.MAIN_DATA_START_ROW, dueDateCol, initialLastRow - CONFIG.MAIN_DATA_START_ROW + 1, 1)
      .getValues();
    values.forEach((v, i) => existing.push({ row: CONFIG.MAIN_DATA_START_ROW + i, time: toTime(v[0]) }));
  }

  // 新規注文ごとに、挿入先(既存データ上の何番目の要素の直後か)を独立して探す。
  // -1 は「該当なし=データの先頭に挿入」を意味する。
  const sortedOrders = newOrders.slice().sort((a, b) => toTime(a.dueDate) - toTime(b.dueDate));
  const targets = sortedOrders.map((order) => ({
    order: order,
    anchorIndex: findInsertAnchorIndex(existing, toTime(order.dueDate)),
  }));

  // シート上の物理的な位置(anchorIndex)の昇順に処理する。
  // 同じ位置に挿入される注文同士は、製造納期昇順(sortedOrdersの並び)を保ったまま連続して入る。
  targets.sort((a, b) => a.anchorIndex - b.anchorIndex);

  let rowOffset = 0; // これまでに挿入した行数(既存行の実際の行番号への補正値)
  const highlightRanges = []; // 赤色にする範囲 {start, end}

  targets.forEach(({ order, anchorIndex }) => {
    let templateRow; // 書式・数式のコピー元にする行
    let insertAfterRow; // この行の直後に挿入する
    if (anchorIndex === -1) {
      // 該当する製造納期(以下)の既存行が無い → データの先頭に挿入する
      templateRow = CONFIG.MAIN_DATA_START_ROW + rowOffset;
      insertAfterRow = templateRow - 1;
    } else {
      templateRow = existing[anchorIndex].row + rowOffset;
      insertAfterRow = templateRow;
    }

    const groupStartRow = insertAfterRow + 1;

    for (let i = 0; i < order.quantity; i++) {
      mainSheet.insertRowAfter(insertAfterRow);
      const newRow = insertAfterRow + 1;

      // 書式・数式をまるごとコピー(工数・LT等の数式列を含む)
      const srcRange = mainSheet.getRange(templateRow, 1, 1, lastCol);
      const destRange = mainSheet.getRange(newRow, 1, 1, lastCol);
      srcRange.copyTo(destRange);

      // ガントチャート部分(PO催促より右)は、テンプレート行の「1」などの値や
      // 赤・薄紫などの塗りつぶし色をそのまま引き継いでしまうと無関係なマークが
      // 混入するため、値をクリアしたうえで、背景色は見出し行(月火水木金土日)と
      // 同じ色に塗り直す(土日だけ、という決め打ちはしない)。
      // 部材支給/製造納期列(F・G列)のオレンジ・青はこの範囲の外なので影響しない。
      if (calendarWidth > 0) {
        const calendarRange = mainSheet.getRange(newRow, calendarStartCol, 1, calendarWidth);
        calendarRange.clearContent();
        calendarRange.setBackgrounds([headerRowBackgrounds]);
      }

      // 転記対象の列だけ値を上書き
      setCellValue(mainSheet, newRow, mainCols.orderDate, today);
      setCellValue(mainSheet, newRow, mainCols.orderNo, order.orderNo);
      setCellValue(mainSheet, newRow, mainCols.neiOrderNo, order.neiOrderNo);
      setCellValue(mainSheet, newRow, mainCols.model, normalizeModel(order.model));
      setCellValue(mainSheet, newRow, mainCols.materialSupply, order.materialSupplyDate);
      setCellValue(mainSheet, newRow, mainCols.dueDate, order.dueDate);
      // 管理No.は空欄にする(直前行の書式コピーで値が引き継がれてしまうため明示的にクリア)
      setCellValue(mainSheet, newRow, mainCols.managementNo, '');

      insertAfterRow = newRow;
      rowOffset++;
    }

    const groupEndRow = insertAfterRow;
    drawGroupBorder(mainSheet, mainCols, groupStartRow, groupEndRow);
    highlightRanges.push({ start: groupStartRow, end: groupEndRow });
  });

  // --- A列(注文日)の色を一旦全クリアしてから、今回追加分だけ赤色にする ---
  const orderDateCol = mainCols.orderDate + 1;
  const finalLastRow = mainSheet.getLastRow();
  mainSheet
    .getRange(CONFIG.MAIN_DATA_START_ROW, orderDateCol, finalLastRow - CONFIG.MAIN_DATA_START_ROW + 1, 1)
    .setBackground(null);
  highlightRanges.forEach((r) => {
    mainSheet.getRange(r.start, orderDateCol, r.end - r.start + 1, 1).setBackground(CONFIG.NEW_ROW_HIGHLIGHT_COLOR);
  });
}

/** NEI注文番号の列に、グループ(数量分の行のかたまり)を囲む太い罫線を引く */
function drawGroupBorder(mainSheet, mainCols, startRow, endRow) {
  const col = mainCols.neiOrderNo + 1;
  const range = mainSheet.getRange(startRow, col, endRow - startRow + 1, 1);
  range.setBorder(true, true, true, true, false, false, CONFIG.GROUP_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THICK);
}

/**
 * 日付らしき値をタイムスタンプ(数値)に変換する。空欄・不正値は最後尾扱い(Infinity)
 * ・Dateオブジェクト、シリアル値(数値)、"26/08/17"や"2026/08/17"のような文字列に対応
 */
function toTime(value) {
  if (value === null || value === undefined || value === '') return Infinity;

  if (value instanceof Date) {
    const t = value.getTime();
    return isNaN(t) ? Infinity : t;
  }

  if (typeof value === 'number') {
    // Google/Excelの日付シリアル値(1899/12/30基準)とみなす
    const epoch = new Date(Date.UTC(1899, 11, 30));
    return epoch.getTime() + value * 24 * 60 * 60 * 1000;
  }

  const str = String(value).trim();

  // "26/08/17" や "2026/08/17" のような YY(YY)/MM/DD 形式
  const m = str.match(/^(\d{2,4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m) {
    let year = Number(m[1]);
    if (year < 100) year += 2000;
    const month = Number(m[2]);
    const day = Number(m[3]);
    const d = new Date(year, month - 1, day);
    return isNaN(d.getTime()) ? Infinity : d.getTime();
  }

  const d = new Date(str);
  return isNaN(d.getTime()) ? Infinity : d.getTime();
}

/**
 * existing(既存データの製造納期の配列。行の並び順=シート上の並び順)の中から、
 * 新規注文を挿入すべき位置(直後に挿入する要素のインデックス)を探す。
 * ・targetTime と完全に同じ製造納期の行があれば、そのうち一番下(行番号が一番大きい)のもの
 * ・完全一致が無ければ、targetTime 以下の製造納期を持つ行のうち一番下のもの
 * ・該当が1つも無ければ -1 (データの先頭に挿入する)
 * シート全体が製造納期の昇順とは限らない(顧客/ロットでブロック分けされている)前提で、
 * 先頭から末尾まで全件を見て判定する。
 */
function findInsertAnchorIndex(existing, targetTime) {
  let lastExact = -1;
  let lastLte = -1;
  for (let i = 0; i < existing.length; i++) {
    const t = existing[i].time;
    if (t === Infinity) continue; // 製造納期が空欄/不明な行は無視する
    if (t === targetTime) lastExact = i;
    if (t <= targetTime) lastLte = i;
  }
  return lastExact !== -1 ? lastExact : lastLte;
}

/**
 * 「今日の日付」を、このスプレッドシートに設定されているタイムゾーンで求める。
 * 単純に new Date() を使うと、スプレッドシートのタイムゾーン設定によっては
 * (例: 日本時間ではなく米国時間になっている場合など)セルに表示される日付が
 * 実際の日本時間とズレることがあるため、スプレッドシート自身のタイムゾーンを
 * 明示的に使って日付を組み立てる。
 */
function getTodayInSpreadsheetTimeZone(sheet) {
  const tz = sheet.getParent().getSpreadsheetTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd');
  return new Date(todayStr);
}

function setCellValue(sheet, row, zeroBasedCol, value) {
  if (zeroBasedCol < 0) return;
  sheet.getRange(row, zeroBasedCol + 1).setValue(value);
}

/** 型式の文字列中で、半角スペースが2つ以上連続していたら1つにまとめる */
function normalizeModel(model) {
  if (model === null || model === undefined) return model;
  return String(model).replace(/ {2,}/g, ' ');
}

/** 取込元シートを選ぶ。候補名があればそれを使い、無ければ選択ダイアログを出す */
function pickSourceSheet(ss, ui) {
  for (const name of CONFIG.SOURCE_SHEET_CANDIDATES) {
    const sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }

  const sheetNames = ss.getSheets().map((s) => s.getName());
  const prompt = ui.prompt(
    '取込元シートの指定',
    `シート「${CONFIG.SOURCE_SHEET_CANDIDATES.join('/')}」が見つかりませんでした。\n` +
    `Excelを インポート済みですか?\n\n` +
    `現在のシート一覧:\n${sheetNames.join(', ')}\n\n` +
    `取り込み元にするシート名を入力してください。`,
    ui.ButtonSet.OK_CANCEL
  );
  if (prompt.getSelectedButton() !== ui.Button.OK) return null;

  const name = prompt.getResponseText().trim();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    ui.alert(`シート「${name}」が見つかりません。処理を中止します。`);
    return null;
  }
  return sheet;
}
