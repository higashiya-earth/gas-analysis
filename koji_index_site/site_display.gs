/**
 * 工事インデックス_v08 用：シート「サイト表示」を工事Index と同じ見た目で作る GAS
 *
 * 工事インデックス_v08 の「拡張機能 ＞ Apps Script」にだけ貼ります（転記用の GAS とは別）。
 * 工事Index の A列に ★ がある行から下を、書式・セル色・文字色・リンク・列の幅・行の高さ・
 * 条件付き書式・固定行・枠線の非表示ごと「サイト表示」へ写します。A列が × の行は写しません。
 * 見出し（1〜3行目）はいつも写します。★ の印はサイト表示には出しません。
 *
 * 更新されるとき
 *  ・A列の ★・× を入力・削除したとき（すぐ）
 *  ・1時間ごと（転記用 GAS が書き足した行を反映）※「自動更新を設定」を1回実行しておく
 *  ・メニュー「サイト表示 ＞ 今すぐ更新」
 *
 * シートは作り直さず中身だけ入れ替えるので、「ウェブに公開」の埋め込みはそのまま使えます。
 *
 * 【フラグ】工事Index の S列（フラグ）をプルダウンで選ぶと、その行の B〜P列の書式を変えます（S列のセルの色も同じにする）。
 *  未施工：元に戻す（文字は黒、セルの色は外す＝しま模様に戻る）
 *  継続-請無・継続-請有・工事完了：文字をピンク
 *  完工：セルをグレー　／　中止・重複：セルを濃いグレー
 */

const SOURCE_SHEET = '工事Index';   // 元のシート
const SITE_SHEET = 'サイト表示';     // 表示用のシート（無ければ作る）
const HEADER_ROWS = 3;              // いつも写す見出しの行数（1〜3行目）
const MARK_COL = 1;                 // 目印を入れる列（A列）
const START_MARK = '★';             // この行から下を表示
const HIDE_MARK = '×';              // この行は表示しない
const SPARE_ROWS = 10;              // サイト表示の下に残す空行
const REFRESH_HOURS = 1;            // 自動更新の間隔（時間）
const DEFAULT_ROW_HEIGHT = 21;

// フラグ（S列）と行の書式
const FLAG_COL = 19;                 // S列
const FLAG_FIRST_ROW = 4;            // データは4行目から
const FORMAT_FIRST_COL = 2;          // B列
const FORMAT_LAST_COL = 16;          // P列
const TEXT_BLACK = '#000000';
const TEXT_PINK = '#FF4FA3';
const FLAG_STYLES = {
  '未施工':   { font: TEXT_BLACK, background: null },        // 元に戻す（しま模様は条件付き書式で付く）
  '継続-請無': { font: TEXT_PINK, background: null },
  '継続-請有': { font: TEXT_PINK, background: null },
  '工事完了': { font: TEXT_PINK, background: null },
  '完工':     { font: TEXT_BLACK, background: '#D9D9D9' },   // グレーアウト
  '中止':     { font: TEXT_BLACK, background: '#808080' },   // ダークアウト
  '重複':     { font: TEXT_BLACK, background: '#808080' }
};


// ===== メニュー・きっかけ =====

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('サイト表示')
    .addItem('今すぐ更新', 'menuUpdateSite')
    .addItem('自動更新を設定（最初に1回）', 'setupSiteTrigger')
    .addSeparator()
    .addItem('フラグの書式をすべて付け直す', 'menuApplyAllFlags')
    .addToUi();
}

/**
 * A列の ★・× を入力・削除したら、すぐ更新する
 */
function onEdit(e) {
  const range = e && e.range;
  if (!range || range.getSheet().getName() !== SOURCE_SHEET) return;
  // S列（フラグ）を選んだら、その行の書式を変える
  if (range.getColumn() <= FLAG_COL && range.getLastColumn() >= FLAG_COL) applyFlagRows_(range.getSheet(), range.getRow(), range.getLastRow(), true);
  // A列の ★・× を変えたら、サイト表示を更新
  if (range.getColumn() <= MARK_COL) updateSiteWithLock_();
}

/**
 * メニュー「フラグの書式をすべて付け直す」：コピー・貼り付けや並べ替えのあとに使う
 */
function menuApplyAllFlags() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_SHEET);
  const n = applyFlagRows_(sheet, FLAG_FIRST_ROW, sheet.getLastRow(), false);
  SpreadsheetApp.getUi().alert(`フラグの書式を付け直しました（${n}行）。`);
}

/**
 * first〜last 行の S列を読み、B〜P列の文字色・セルの色を付ける
 *  一覧に無い値の行は触らない
 * @param {boolean} blankResets true：空欄も「未施工」と同じく元に戻す（フラグを消したとき）
 *                              false：空欄の行は触らない（まとめて付け直すとき。手で付けた色を消さない）
 * @return {number} 書式を付けた行数
 */
function applyFlagRows_(sheet, first, last, blankResets) {
  const from = Math.max(first, FLAG_FIRST_ROW);
  if (last < from) return 0;
  const flags = sheet.getRange(from, FLAG_COL, last - from + 1, 1).getDisplayValues();
  const width = FORMAT_LAST_COL - FORMAT_FIRST_COL + 1;
  let count = 0;
  flags.forEach((r, i) => {
    const flag = String(r[0]).trim();
    const style = flag === '' ? (blankResets ? FLAG_STYLES['未施工'] : null) : FLAG_STYLES[flag];
    if (!style) return;
    const range = sheet.getRange(from + i, FORMAT_FIRST_COL, 1, width);
    range.setFontColor(style.font);
    range.setBackground(style.background);
    sheet.getRange(from + i, FLAG_COL).setBackground(style.background);   // フラグのセルも同じ行の色に（文字色は変えない）
    count++;
  });
  return count;
}

function menuUpdateSite() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = updateSiteWithLock_();
    ui.alert(r ? `サイト表示を更新しました。\n${r.from}行目から ${r.rows}行を表示` + (r.hidden ? `（× の ${r.hidden}行は除外）` : '')
      : '別の更新が動いていたので、少し待ってからもう一度押してください。');
  } catch (err) {
    ui.alert(`サイト表示を更新できませんでした。\n${err.message}`);
  }
}

/**
 * 1時間ごとの自動更新を作る（何回実行しても1つだけ）
 */
function setupSiteTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshSite')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshSite').timeBased().everyHours(REFRESH_HOURS).create();
  updateSiteWithLock_();
  try {
    SpreadsheetApp.getUi().alert(`${REFRESH_HOURS}時間ごとにサイト表示を更新するように設定しました。`);
  } catch (e) { /* エディタから実行したとき */ }
}

/**
 * 自動更新を止める
 */
function removeSiteTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshSite')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function refreshSite() {
  updateSiteWithLock_();
}

function updateSiteWithLock_() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return null;
  try {
    return updateSite_();
  } finally {
    lock.releaseLock();
  }
}


// ===== 本体 =====

/**
 * 工事Index を見た目ごと「サイト表示」へ写す
 * @return {{from: number, rows: number, hidden: number}} 表示を始めた行・表示した行数・除いた行数
 */
function updateSite_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SOURCE_SHEET);
  if (!src) throw new Error(`シート「${SOURCE_SHEET}」がありません`);
  const dst = ss.getSheetByName(SITE_SHEET) || ss.insertSheet(SITE_SHEET, ss.getSheets().length);

  const cols = src.getMaxColumns();
  const last = Math.max(src.getLastRow(), HEADER_ROWS);

  // 写す行を決める
  const body = last > HEADER_ROWS ? src.getRange(HEADER_ROWS + 1, 1, last - HEADER_ROWS, cols).getDisplayValues() : [];
  const marks = body.map(r => String(r[MARK_COL - 1]).trim());
  const star = marks.indexOf(START_MARK);
  const fromIdx = Math.max(star, 0);
  const rows = [];
  for (let r = 1; r <= HEADER_ROWS; r++) rows.push(r);
  let hidden = 0;
  for (let i = fromIdx; i < body.length; i++) {
    if (marks[i] === HIDE_MARK) { hidden++; continue; }
    rows.push(HEADER_ROWS + 1 + i);
  }

  // サイト表示を空にして、行数・列数を合わせる
  const all = dst.getRange(1, 1, dst.getMaxRows(), dst.getMaxColumns());
  all.breakApart();
  all.clearDataValidations();
  dst.clear();
  dst.clearConditionalFormatRules();
  dst.setFrozenRows(0);
  dst.setFrozenColumns(0);
  const need = rows.length + SPARE_ROWS;
  if (dst.getMaxRows() > need) dst.deleteRows(need + 1, dst.getMaxRows() - need);
  else if (dst.getMaxRows() < need) dst.insertRowsAfter(dst.getMaxRows(), need - dst.getMaxRows());
  if (dst.getMaxColumns() < cols) dst.insertColumnsAfter(dst.getMaxColumns(), cols - dst.getMaxColumns());

  // 続いている行はまとめて写す（値・書式・セル色・文字色・リンク・結合・条件付き書式）
  let destRow = 1;
  rowRuns_(rows).forEach(run => {
    src.getRange(run.start, 1, run.length, cols).copyTo(dst.getRange(destRow, 1));
    destRow += run.length;
  });

  // 列の幅・行の高さ・固定・枠線・隠した列
  src.getRange(1, 1, 1, cols).copyTo(dst.getRange(1, 1), SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
  sameRuns_(rows.map(r => src.getRowHeight(r))).forEach(run => dst.setRowHeights(run.index + 1, run.length, run.value));
  dst.setRowHeights(rows.length + 1, SPARE_ROWS, DEFAULT_ROW_HEIGHT);
  for (let c = 1; c <= cols; c++) {
    if (src.isColumnHiddenByUser(c)) dst.hideColumns(c); else dst.showColumns(c);
  }
  dst.setFrozenRows(Math.min(src.getFrozenRows(), HEADER_ROWS));
  dst.setFrozenColumns(src.getFrozenColumns());
  dst.setHiddenGridlines(src.hasHiddenGridlines());

  // ★ の印は見せない
  if (star >= 0) dst.getRange(HEADER_ROWS + 1, MARK_COL).clearContent();

  return { from: HEADER_ROWS + 1 + fromIdx, rows: rows.length - HEADER_ROWS, hidden };
}

/**
 * 行番号の並びを、続いている区間に分ける → [{start, length}]
 */
function rowRuns_(rows) {
  const runs = [];
  rows.forEach(r => {
    const prev = runs[runs.length - 1];
    if (prev && r === prev.start + prev.length) prev.length++;
    else runs.push({ start: r, length: 1 });
  });
  return runs;
}

/**
 * 値の並びを、同じ値が続く区間に分ける → [{value, index, length}]
 */
function sameRuns_(values) {
  const runs = [];
  values.forEach((v, i) => {
    const prev = runs[runs.length - 1];
    if (prev && v === prev.value) prev.length++;
    else runs.push({ value: v, index: i, length: 1 });
  });
  return runs;
}
