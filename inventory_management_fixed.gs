/**
 * 在庫管理表 GAS（修正版）
 *
 * 【元のスクリプトからの主な修正点】
 *  1. 棚卸登録：N列が空の行のO列（当月在庫）の計算式を消さないようにした
 *  2. 仕入登録・使用量登録・棚卸登録：全角数字は半角に直して受け付け、
 *     数値として読めない入力（例：「10袋」）があれば何も登録せずエラー表示にした
 *  3. 翌月繰越：L列・P列に未登録の入力が残っていれば繰越を止めるようにした
 *  4. R列「使用量計算」の式を =J+M-O（前月在庫＋仕入−在庫＝実使用量）に変更した
 *     （元の式 =O-J-M は使用量のマイナスになっていた）
 *  5. 翌月繰越：並べ替えのまとめ方を「仕入先（C列）ごと」と明記し、
 *     繰越後のシートでは「新規」表示を消すようにした（設定で切替可）
 *  6. 最新月の在庫表シートでのみ動くようにした（過去月の誤操作防止）
 *  7. その他：連番は品名のある行だけに振る／225行目以降に品があれば繰越を止める／
 *     数式の書き戻しを列単位にまとめて高速化／繰越に失敗したら作りかけのシートを削除／
 *     新シートを現在のシートの左隣に配置／完了メッセージに件数を表示
 */


// ===== 共通設定 =====

const START_ROW = 5;
const END_ROW = 224;
const NUM_ROWS = END_ROW - START_ROW + 1;

const COL = {
  A: 1,   // 繰越（新規・廃版）
  C: 3,   // 仕入先
  D: 4,   // 品名
  J: 10,  // 前月在庫
  L: 12,  // 当月仕入（入力）
  M: 13,  // 当月仕入計
  N: 14,  // 当月棚卸（入力）
  O: 15,  // 当月在庫
  P: 16,  // 当月使用（入力）
  Q: 17,  // 当月使用計
  R: 18   // 使用量計算
};

// 翌月繰越後の新シートで「新規」表示を消すかどうか
const CLEAR_NEW_MARK_ON_CARRYOVER = true;

// 月別在庫表シート名（例：R810在庫表 → 令和8年10月）
const STOCK_SHEET_PATTERN = /^R(\d+)(\d{2})在庫表$/;

// R列「使用量計算」の式：前月在庫＋当月仕入計−当月在庫（＝実使用量）
function usageFormula_(row) {
  return `=J${row}+M${row}-O${row}`;
}

// O列「当月在庫」の式：前月在庫＋当月仕入計−当月使用計
function currentStockFormula_(row) {
  return `=J${row}+M${row}-Q${row}`;
}


/**
 * 初回だけ手動実行
 * チェックボックス起動用のインストール型 onEdit トリガーを作成
 *
 * ※スマホ対応のため、シンプルトリガーではなくインストール型トリガーを使用します。
 */
function createOnEditTrigger() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既に同じトリガーがある場合は削除して、重複起動を防ぐ
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "onEditCheckboxTrigger") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("onEditCheckboxTrigger")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  console.log("チェックボックス起動用トリガーを作成しました。");
}


/**
 * チェックボックス起動用
 *
 * A2：翌月繰越
 * L2：仕入登録
 * N2：棚卸登録
 * P2：使用量登録
 *
 * スマホ対応のため、alert は使わず、処理状況はチェックボックス直下のセルに表示します。
 */
function onEditCheckboxTrigger(e) {

  if (!e || !e.range) return;

  // チェックされた時だけ実行
  if (e.value !== "TRUE") return;

  const range = e.range;
  const sheet = range.getSheet();
  const cell = range.getA1Notation();

  const SUCCESS_MESSAGE_TIME = 2000; // 完了メッセージ表示時間：2秒
  const ERROR_MESSAGE_TIME = 5000;   // エラーメッセージ表示時間：5秒

  const settings = {
    "A2": { name: "翌月繰越",   statusCell: "A3", action: () => PushCarryover(sheet) },
    "L2": { name: "仕入登録",   statusCell: "L3", action: () => PurchaseRegistration(sheet) },
    "N2": { name: "棚卸登録",   statusCell: "N3", action: () => InventoryRegistration(sheet) },
    "P2": { name: "使用量登録", statusCell: "P3", action: () => UsageRegistration(sheet) }
  };

  // 対象セル以外は無視
  if (!settings[cell]) return;

  // 月別在庫表シート以外（参照リスト・原本など）は何もしない
  if (stockSheetKey_(sheet.getName()) === null) return;

  const setting = settings[cell];
  const statusRange = sheet.getRange(setting.statusCell);
  const lock = LockService.getScriptLock();
  let hasLock = false;

  try {
    // 二重実行防止
    hasLock = lock.tryLock(30000);
    if (!hasLock) {
      throw new Error("他の処理が実行中です。少し待ってから再度チェックしてください。");
    }

    // 過去月のシートでの誤操作を防ぐ
    assertLatestStockSheet_(sheet);

    statusRange.setValue(setting.name + "を実行中...");
    SpreadsheetApp.flush();

    const result = setting.action();

    SpreadsheetApp.flush();

    let message = setting.name + "が完了しました。";
    if (typeof result === "number") {
      message = `${setting.name}が完了しました（${result}件）。`;
    } else if (typeof result === "string") {
      message = `${setting.name}が完了しました（${result}）。`;
    }

    statusRange.setValue(message);
    SpreadsheetApp.flush();

    Utilities.sleep(SUCCESS_MESSAGE_TIME);
    statusRange.clearContent();
    SpreadsheetApp.flush();

  } catch (error) {

    statusRange.setValue(setting.name + "でエラーが発生しました：" + error.message);
    SpreadsheetApp.flush();

    console.error(error);

    Utilities.sleep(ERROR_MESSAGE_TIME);
    statusRange.clearContent();
    SpreadsheetApp.flush();

  } finally {

    Utilities.sleep(500);

    // 元シート側のチェックボックスをOFFに戻す
    range.setValue(false);
    SpreadsheetApp.flush();

    if (hasLock) {
      lock.releaseLock();
    }
  }
}


/**
 * １．翌月繰越スクリプト
 *
 * チェックボックスを押したシートを対象に、翌月シートを作成します。
 */
function PushCarryover(targetSheet) {

  const currentSheet = targetSheet;
  const ss = currentSheet.getParent();

  // 事前チェック（ここで止まった場合はシートを作らない）
  assertNoPendingInput_(currentSheet);
  assertNoItemsBeyondRange_(currentSheet);

  const a1Value = currentSheet.getRange("A1").getValue();
  const baseDate = a1Value instanceof Date ? new Date(a1Value.getTime()) : new Date(a1Value);

  if (isNaN(baseDate.getTime())) {
    throw new Error("セルA1に有効な日付が入っていません。");
  }

  // 翌月1日にする（先に1日にしてから月を進めるので、月末日でもずれない）
  baseDate.setDate(1);
  baseDate.setMonth(baseDate.getMonth() + 1);

  const year = baseDate.getFullYear();
  const month = baseDate.getMonth() + 1;
  const reiwaYear = year - 2018;

  if (reiwaYear <= 0) {
    throw new Error("2019年、令和元年以降の日付を対象にしてください。");
  }

  const newSheetName = "R" + reiwaYear + ("0" + month).slice(-2) + "在庫表";

  if (ss.getSheetByName(newSheetName)) {
    throw new Error("同じ名前のシートが既に存在します：" + newSheetName);
  }

  const newSheet = currentSheet.copyTo(ss);

  try {
    newSheet.setName(newSheetName);

    // コピー直後はチェック状態や「実行中...」表示もコピーされているため初期化
    resetControlCells_(newSheet);

    // A1に翌月日付をセット
    newSheet.getRange("A1").setValue(baseDate).setNumberFormat("yyyy/m/d");

    // 当月在庫 O列 → 翌月の前月在庫 J列
    const sourceOValues = currentSheet.getRange(START_ROW, COL.O, NUM_ROWS, 1).getValues();
    newSheet.getRange(START_ROW, COL.J, NUM_ROWS, 1).setValues(sourceOValues);

    // 翌月シートの入力欄・集計欄をクリア
    [COL.L, COL.M, COL.N, COL.O, COL.P, COL.Q].forEach(col => {
      newSheet.getRange(START_ROW, col, NUM_ROWS, 1).clearContent();
    });

    // 行整理（廃版の削除・新規の仕入先を上へ）
    organizeCarryoverRows_(newSheet);

    if (CLEAR_NEW_MARK_ON_CARRYOVER) {
      clearNewMarks_(newSheet);
    }

    // 連番振り直し・計算式の再設定
    resetSerialNumbers_(newSheet);
    setCurrentStockFormulas_(newSheet);
    setUsageFormulas_(newSheet);

    resetControlCells_(newSheet);

    // 新しいシートを現在のシートの左隣に置いて表示
    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(currentSheet.getIndex());

    SpreadsheetApp.flush();

  } catch (error) {
    // 途中で失敗したら作りかけのシートを消して、やり直せる状態に戻す
    ss.setActiveSheet(currentSheet);
    ss.deleteSheet(newSheet);
    throw error;
  }

  return newSheetName + "を作成";
}


/**
 * 繰越前チェック：L列（仕入）・P列（使用）に未登録の入力が残っていないか
 */
function assertNoPendingInput_(sheet) {

  const targets = [
    { col: COL.L, letter: "L", name: "仕入登録" },
    { col: COL.P, letter: "P", name: "使用量登録" }
  ];

  targets.forEach(target => {
    const values = sheet.getRange(START_ROW, target.col, NUM_ROWS, 1).getDisplayValues();
    const rows = [];

    values.forEach((row, i) => {
      if (String(row[0]).trim() !== "") rows.push(START_ROW + i);
    });

    if (rows.length > 0) {
      throw new Error(
        `${target.letter}列に未登録の入力があります（${formatRowList_(rows)}）。` +
        `先に「${target.name}」をしてください。`
      );
    }
  });
}


/**
 * 繰越前チェック：処理対象（5～224行目）より下に品が入っていないか
 */
function assertNoItemsBeyondRange_(sheet) {

  const lastRow = sheet.getLastRow();
  if (lastRow <= END_ROW) return;

  const values = sheet.getRange(END_ROW + 1, COL.C, lastRow - END_ROW, 2).getDisplayValues();
  const rows = [];

  values.forEach((row, i) => {
    if (String(row[0]).trim() !== "" || String(row[1]).trim() !== "") rows.push(END_ROW + 1 + i);
  });

  if (rows.length > 0) {
    throw new Error(
      `${END_ROW}行目より下に品があります（${formatRowList_(rows)}）。` +
      `${START_ROW}～${END_ROW}行目の中に移してください。`
    );
  }
}


/**
 * 新しく作成されたシート側の操作セルを初期化する
 */
function resetControlCells_(sheet) {

  ["A2", "L2", "N2", "P2"].forEach(a1 => sheet.getRange(a1).setValue(false));
  ["A3", "L3", "N3", "P3"].forEach(a1 => sheet.getRange(a1).clearContent());

  SpreadsheetApp.flush();
}


/**
 * 翌月繰越時の行整理
 *
 * ・A列が「廃版」の行は除外
 * ・A列が「新規」の行がある仕入先（C列）は、その仕入先の行をまとめて上へ集める
 * ・それ以外の行は元の順番のまま後ろへ
 * ・除外した行の分は下に空行を補う（金額などの数式はその行に残す）
 */
function organizeCarryoverRows_(sheet) {

  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(START_ROW, 1, NUM_ROWS, lastCol);

  const values = range.getValues();
  const formulas = range.getFormulasR1C1();
  const displays = range.getDisplayValues();

  const rows = values.map((rowValues, i) => ({
    index: i,
    mark: String(displays[i][COL.A - 1]).trim(),
    supplier: String(displays[i][COL.C - 1]).trim(),
    values: rowValues,
    formulas: formulas[i]
  }));

  const discontinuedRows = rows.filter(row => row.mark === "廃版");
  const remainingRows = rows.filter(row => row.mark !== "廃版");

  // 「新規」の行がある仕入先（出てきた順）
  const newSuppliers = [...new Set(
    remainingRows
      .filter(row => row.mark === "新規" && row.supplier !== "")
      .map(row => row.supplier)
  )];

  const sortedRows = [];
  const used = new Set();

  newSuppliers.forEach(supplier => {
    remainingRows.forEach(row => {
      if (row.supplier === supplier && !used.has(row.index)) {
        sortedRows.push(row);
        used.add(row.index);
      }
    });
  });

  remainingRows.forEach(row => {
    if (!used.has(row.index)) {
      sortedRows.push(row);
      used.add(row.index);
    }
  });

  // 除外した行の分を補う。数式（相対参照）はそのまま残し、値だけ空にする
  const emptyValues = new Array(lastCol).fill("");
  discontinuedRows.forEach(row => {
    sortedRows.push({ values: emptyValues, formulas: row.formulas });
  });

  range.clearContent();
  range.setValues(sortedRows.map(row => row.values));

  // 数式は列ごとにまとめて書き戻す
  for (let c = 0; c < lastCol; c++) {

    const columnFormulas = sortedRows.map(row => row.formulas[c]);
    if (columnFormulas.every(f => f === "")) continue;

    // その列が「数式か空欄」だけなら一括で書き込める
    const formulaOrBlankOnly = sortedRows.every(row => row.formulas[c] !== "" || row.values[c] === "");

    if (formulaOrBlankOnly) {
      sheet.getRange(START_ROW, c + 1, NUM_ROWS, 1)
        .setFormulasR1C1(columnFormulas.map(f => [f]));
    } else {
      columnFormulas.forEach((f, r) => {
        if (f !== "") sheet.getRange(START_ROW + r, c + 1).setFormulaR1C1(f);
      });
    }
  }
}


/**
 * A列の「新規」表示を消す（「廃版」などは対象外）
 */
function clearNewMarks_(sheet) {

  const range = sheet.getRange(START_ROW, COL.A, NUM_ROWS, 1);
  const values = range.getValues();

  if (!values.some(row => String(row[0]).trim() === "新規")) return;

  range.setValues(values.map(row => [String(row[0]).trim() === "新規" ? "" : row[0]]));
}


/**
 * 品のある行（C列またはD列に入力あり）を返す
 */
function getItemRowFlags_(sheet) {

  return sheet.getRange(START_ROW, COL.C, NUM_ROWS, 2).getDisplayValues()
    .map(row => String(row[0]).trim() !== "" || String(row[1]).trim() !== "");
}


/**
 * B列の連番を、品のある行だけに 1 から振り直す
 */
function resetSerialNumbers_(sheet) {

  let serial = 0;
  const serialValues = getItemRowFlags_(sheet).map(hasItem => [hasItem ? ++serial : ""]);

  sheet.getRange(START_ROW, 2, NUM_ROWS, 1).setValues(serialValues);
}


/**
 * O列に当月在庫の計算式を設定（全行。月の途中で品を追加しても計算されるように）
 */
function setCurrentStockFormulas_(sheet) {

  const formulas = [];
  for (let row = START_ROW; row <= END_ROW; row++) {
    formulas.push([currentStockFormula_(row)]);
  }

  sheet.getRange(START_ROW, COL.O, NUM_ROWS, 1).setFormulas(formulas);
}


/**
 * R列に使用量計算の式を設定（品のある行だけ）
 */
function setUsageFormulas_(sheet) {

  const formulas = getItemRowFlags_(sheet)
    .map((hasItem, i) => [hasItem ? usageFormula_(START_ROW + i) : ""]);

  sheet.getRange(START_ROW, COL.R, NUM_ROWS, 1).setFormulas(formulas);
}


/**
 * ２．仕入登録スクリプト
 *
 * L列の仕入数をM列へ加算し、L列をクリアします。
 */
function PurchaseRegistration(targetSheet) {
  return addInputToTotal_(targetSheet, COL.L, COL.M, "L");
}


/**
 * ３．使用量登録スクリプト
 *
 * P列の使用量をQ列へ加算し、P列をクリアします。
 */
function UsageRegistration(targetSheet) {
  return addInputToTotal_(targetSheet, COL.P, COL.Q, "P");
}


/**
 * 入力列の数を集計列へ加算し、入力列をクリアする（仕入登録・使用量登録の共通処理）
 *
 * 数値でない入力が1つでもあれば、何も書き換えずにエラーにする。
 */
function addInputToTotal_(sheet, inputCol, totalCol, inputLetter) {

  const inputRange = sheet.getRange(START_ROW, inputCol, NUM_ROWS, 1);
  const totalRange = sheet.getRange(START_ROW, totalCol, NUM_ROWS, 1);

  const inputs = parseNumberColumn_(inputRange.getValues(), inputLetter);
  const totalValues = totalRange.getValues();
  const totalFormulas = totalRange.getFormulas();

  let count = 0;

  const newTotals = inputs.map((input, i) => {
    // 入力のない行は元のまま（数式があれば数式のまま）
    if (input === null) return [totalFormulas[i][0] || totalValues[i][0]];

    count++;
    return [(Number(totalValues[i][0]) || 0) + input];
  });

  if (count === 0) {
    throw new Error(inputLetter + "列に入力がありません。");
  }

  totalRange.setValues(newTotals);
  inputRange.clearContent();

  SpreadsheetApp.flush();

  return count;
}


/**
 * ４．棚卸登録スクリプト
 *
 * N列に棚卸数がある行だけ、O列へ棚卸数を反映し、R列へ使用量計算の式を設定します。
 * N列が空の行は、O列・R列を一切変更しません（計算式を残します）。
 * N列は消さないので、棚卸数を直して再実行することもできます。
 */
function InventoryRegistration(targetSheet) {

  const sheet = targetSheet;

  const nRange = sheet.getRange(START_ROW, COL.N, NUM_ROWS, 1);
  const oRange = sheet.getRange(START_ROW, COL.O, NUM_ROWS, 1);
  const rRange = sheet.getRange(START_ROW, COL.R, NUM_ROWS, 1);

  const counts = parseNumberColumn_(nRange.getValues(), "N");

  const oValues = oRange.getValues();
  const oFormulas = oRange.getFormulas();
  const rValues = rRange.getValues();
  const rFormulas = rRange.getFormulas();

  const newOValues = [];
  const newRValues = [];
  let count = 0;

  counts.forEach((countValue, i) => {

    if (countValue === null) {
      // 棚卸数のない行は、数式があれば数式のまま戻す
      newOValues.push([oFormulas[i][0] || oValues[i][0]]);
      newRValues.push([rFormulas[i][0] || rValues[i][0]]);
      return;
    }

    count++;
    newOValues.push([countValue]);
    newRValues.push([usageFormula_(START_ROW + i)]);
  });

  if (count === 0) {
    throw new Error("N列に棚卸数が入力されていません。");
  }

  // 全角で入力された棚卸数は半角の数値に直しておく
  nRange.setValues(counts.map(countValue => [countValue === null ? "" : countValue]));

  oRange.setValues(newOValues);
  rRange.setValues(newRValues);

  SpreadsheetApp.flush();

  return count;
}


/**
 * 入力列（1列分）の値を数値に変換する
 *
 * 空欄は null。数値でない入力があれば、行番号を添えてエラーにする。
 */
function parseNumberColumn_(values, columnLetter) {

  const errors = [];

  const numbers = values.map((row, i) => {
    const number = toNumber_(row[0]);
    if (Number.isNaN(number)) errors.push(`${START_ROW + i}行目「${row[0]}」`);
    return number;
  });

  if (errors.length > 0) {
    const shown = errors.slice(0, 5).join("、") + (errors.length > 5 ? " ほか" + (errors.length - 5) + "件" : "");
    throw new Error(`${columnLetter}列に数値でない入力があります：${shown}。直してから再度チェックしてください（何も登録していません）。`);
  }

  return numbers;
}


/**
 * セルの値を数値に変換する
 *
 * 空欄 → null、数値にできない → NaN
 * 全角数字・全角記号（．－＋，）や長音「ー」・「−」のマイナスも受け付ける
 */
function toNumber_(value) {

  if (value === "" || value === null || value === undefined) return null;
  if (typeof value === "number") return value;

  const text = String(value)
    .replace(/[０-９．－＋，]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/[ー−]/g, "-")
    .replace(/,/g, "")
    .replace(/[\s　]/g, "");

  if (text === "") return null;
  if (!/^[+-]?(\d+\.?\d*|\.\d+)$/.test(text)) return NaN;

  return Number(text);
}


/**
 * 月別在庫表シート名を比較用の数値にする（R810在庫表 → 810、R1001在庫表 → 1001）
 * 月別在庫表でなければ null
 */
function stockSheetKey_(sheetName) {

  const match = String(sheetName).match(STOCK_SHEET_PATTERN);
  if (!match) return null;

  return Number(match[1]) * 100 + Number(match[2]);
}


/**
 * 最新月の在庫表シートでなければエラーにする
 */
function assertLatestStockSheet_(sheet) {

  const keys = sheet.getParent().getSheets()
    .map(s => stockSheetKey_(s.getName()))
    .filter(key => key !== null);

  if (stockSheetKey_(sheet.getName()) !== Math.max(...keys)) {
    throw new Error("最新月の在庫表シートでのみ実行できます。");
  }
}


/**
 * 行番号の一覧を短く表示する（例：5行目、8行目 ほか3件）
 */
function formatRowList_(rows) {
  return rows.slice(0, 5).map(r => r + "行目").join("、") +
    (rows.length > 5 ? " ほか" + (rows.length - 5) + "件" : "");
}
