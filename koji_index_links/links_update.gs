/**
 * 工事インデックス リンク一覧：URL一覧を自動で書き直す GAS
 *
 * 「工事インデックス リンク一覧」の「拡張機能 ＞ Apps Script」に貼ります。
 * シート「設定」に書いた年ごとのフォルダーから「工事インデックス　YYYY年M月」のファイルを探し、
 * シート「URL一覧」の B〜H列（年・月・ファイル名・区分・URL・フォルダー・更新日）を書き直します。
 * 表示用シート（2024〜2026年・2027〜2029年）は数式なので、URL一覧が変わると自動で表示が変わります。
 *
 * 更新されるとき
 *  ・毎日 朝5時台（「自動更新を設定」を1回実行しておく）
 *  ・メニュー「リンク一覧 ＞ 今すぐ更新」
 */

const SETTINGS_SHEET = '設定';
const LIST_SHEET = 'URL一覧';
const LIST_FIRST_ROW = 2;          // 見出しの下
const LIST_COLS = 7;               // B〜H列
const REFRESH_HOUR = 5;            // 毎日この時台に更新
const TIMEZONE = 'Asia/Tokyo';
const TITLE_PATTERN = /(\d{4})\s*年\s*(\d{1,2})\s*月/;   // 「2026年11月」の部分
const SKIP_PATTERN = /コピー|copy/i;                     // 「〜 のコピー」は載せない


// ===== メニュー・きっかけ =====

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('リンク一覧')
    .addItem('今すぐ更新', 'menuUpdateLinks')
    .addItem('自動更新を設定（最初に1回）', 'setupLinksTrigger')
    .addToUi();
}

function menuUpdateLinks() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = updateLinksWithLock_();
    ui.alert(r ? resultText_(r) : '別の更新が動いていたので、少し待ってからもう一度押してください。');
  } catch (err) {
    ui.alert(`URL一覧を更新できませんでした。\n${err.message}`);
  }
}

/**
 * 毎日の自動更新を作る（何回実行しても1つだけ）。あわせて1回更新する
 */
function setupLinksTrigger() {
  removeLinksTrigger();
  ScriptApp.newTrigger('refreshLinks').timeBased().atHour(REFRESH_HOUR).everyDays(1).inTimezone(TIMEZONE).create();
  const r = updateLinksWithLock_();
  try {
    SpreadsheetApp.getUi().alert(`毎日 ${REFRESH_HOUR}時台に URL一覧を更新するように設定しました。\n\n${r ? resultText_(r) : ''}`);
  } catch (e) { /* エディタから実行したとき */ }
}

function removeLinksTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshLinks')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function refreshLinks() {
  const r = updateLinksWithLock_();
  if (r && r.errors.length) throw new Error(r.errors.join('\n'));   // 実行ログ・失敗メールで分かるように
}

function updateLinksWithLock_() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return null;
  try {
    return updateLinks_();
  } finally {
    lock.releaseLock();
  }
}

function resultText_(r) {
  const lines = [`URL一覧を ${r.rows}件で書き直しました。`];
  r.folders.forEach(f => lines.push(`・${f.year}${f.kubun ? ' ' + f.kubun : ''}：${f.count}件（${f.name}）`));
  if (r.skipped.length) lines.push('', '載せなかったファイル：', ...r.skipped.map(s => '・' + s));
  if (r.errors.length) lines.push('', '開けなかったフォルダー：', ...r.errors.map(s => '・' + s));
  return lines.join('\n');
}


// ===== 本体 =====

/**
 * 設定のフォルダーを読み、URL一覧の B〜H列を書き直す
 */
function updateLinks_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const list = ss.getSheetByName(LIST_SHEET);
  if (!list) throw new Error(`シート「${LIST_SHEET}」がありません`);

  const rows = [], folders = [], skipped = [], errors = [];
  readSettings_(ss).forEach(s => {
    let folder;
    try {
      folder = DriveApp.getFolderById(s.folderId);
    } catch (e) {
      errors.push(`${s.year}${s.kubun ? ' ' + s.kubun : ''}：${s.url}`);
      return;
    }
    const found = collectFiles_(folder, s, skipped);
    folders.push({ year: s.year, kubun: s.kubun, name: folder.getName(), count: found.length });
    rows.push(...found);
  });

  // 年 → 区分（空欄が先、R が後）→ 月 の順。表示用シートは同じ年月なら下の行（R）を使う
  rows.sort((a, b) => a[0] - b[0] || (a[3] > b[3] ? 1 : a[3] < b[3] ? -1 : 0) || a[1] - b[1]);

  // 書き直し（A列のキーは数式なので触らない）
  const old = Math.max(list.getLastRow() - LIST_FIRST_ROW + 1, 0);
  if (old > 0) list.getRange(LIST_FIRST_ROW, 2, old, LIST_COLS).clearContent();
  if (rows.length) list.getRange(LIST_FIRST_ROW, 2, rows.length, LIST_COLS).setValues(rows);

  return { rows: rows.length, folders, skipped, errors };
}

/**
 * 設定シート：A＝年、B＝区分、C＝フォルダーのURL（2行目から、年が空の行で終わり）
 */
function readSettings_(ss) {
  const sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) throw new Error(`シート「${SETTINGS_SHEET}」がありません`);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const out = [];
  for (const [y, k, url] of sheet.getRange(2, 1, last - 1, 3).getDisplayValues()) {
    const year = Number(toHalfWidth_(y).replace(/[^\d]/g, ''));
    if (!year) break;
    const folderId = folderIdOf_(url);
    if (!folderId) continue;
    out.push({ year, kubun: String(k).trim(), url, folderId });
  }
  return out;
}

/**
 * フォルダーの中から、設定の年と同じ年のファイルを集める
 * 同じ年月が2つあるときは、更新日の新しい方を使う
 * @return {Array[]} [年, 月, ファイル名, 区分, URL, フォルダー名, 更新日]
 */
function collectFiles_(folder, s, skipped) {
  const byMonth = {};
  const it = folder.getFiles();
  while (it.hasNext()) {
    const file = it.next();
    const name = file.getName();
    const title = toHalfWidth_(name);
    const m = title.match(TITLE_PATTERN);
    if (!m || title.indexOf('工事インデックス') < 0) continue;
    const year = Number(m[1]), month = Number(m[2]);
    if (year !== s.year || month < 1 || month > 12) continue;
    if (SKIP_PATTERN.test(title)) { skipped.push(name); continue; }
    const updated = file.getLastUpdated();
    const prev = byMonth[month];
    if (prev && prev.updated >= updated) { skipped.push(name); continue; }
    if (prev) skipped.push(prev.name);
    byMonth[month] = { name, updated, url: file.getUrl() };
  }
  return Object.keys(byMonth).map(Number).map(month => {
    const f = byMonth[month];
    return [s.year, month, f.name.replace(/\.xlsx?$/i, ''), s.kubun, f.url, folder.getName(), f.updated];
  });
}

function folderIdOf_(url) {
  const text = String(url).trim();
  const m = text.match(/folders\/([-\w]{10,})/) || text.match(/[?&]id=([-\w]{10,})/) || text.match(/^([-\w]{10,})$/);
  return m ? m[1] : '';
}

function toHalfWidth_(s) {
  return String(s).replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/　/g, ' ');
}
