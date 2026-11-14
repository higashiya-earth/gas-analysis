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


// ===== メニュー・きっかけ =====

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('サイト表示')
    .addItem('今すぐ更新', 'menuUpdateSite')
    .addItem('自動更新を設定（最初に1回）', 'setupSiteTrigger')
    .addToUi();
}

/**
 * A列の ★・× を入力・削除したら、すぐ更新する
 */
function onEdit(e) {
  const range = e && e.range;
  if (!range || range.getSheet().getName() !== SOURCE_SHEET) return;
  if (range.getColumn() > MARK_COL) return;
  updateSiteWithLock_();
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
