/**
 * 工事受注入力 転記 GAS
 *
 * 設定ファイル（設定ファイル_YYMM）の内容に従って動きます。列の対応・ファイル名は
 * コードに書かず、毎回設定ファイルから読み込みます。
 *
 * 【GAS_1】工事受注入力 → 事前確認用（設定 5〜8）
 *  フォームから ❶工事受注入力 に行が追加されたら、その行を
 *  ❷確認_受注一覧・❸確認_工事インデックス の「列番号」の列へ書き込み、
 *  両方に書けたら ❶ の copyflag に日時を入れる。
 *
 * 【GAS_2】事前確認用 → 本番（設定 9）
 *  設定ファイルのメニュー「カスタムメニュー ＞ コピー実行（…）」、または毎日深夜1時に、
 *  copyflag が空の行を本番「受注一覧」「工事インデックス」へ書き込み、
 *  書けた行の copyflag に日時を入れる。
 *  受注一覧は「27期11月」のような期と月のシートへ書く（期は7月始まり）。
 *
 * 【設置】このコードを 設定ファイル の Apps Script（拡張機能 ＞ Apps Script）に貼り、
 *  setupTriggers を1回実行する（README.md 参照）。確認用ファイルには貼らなくてよい。
 */


// ===== 共通設定 =====

// 設定ファイルのID（空なら、setupTriggers を実行した設定ファイルを使う）
const SETTINGS_FILE_ID = '';
const SETTINGS_ID_PROPERTY = 'SETTINGS_FILE_ID';

const SETTINGS_SHEET_NAME = '設定ファイル';
const STAFF_TABLE_TITLE = '担当名変換テーブル';
const COPYFLAG_NAME = 'copyflag';
const NO_STAFF = '無し';

// 5) 書込み行：A〜H 列が空の行で、その下2行も空の最初の行
const BLANK_CHECK_LAST_COL = 8;
const BLANK_ROWS_BELOW = 2;
const SEARCH_START_ROW = 2;

const NIGHTLY_HOUR = 1;
const TIMEZONE = 'Asia/Tokyo';
const FLAG_FORMAT = 'yyyy/MM/dd HH:mm:ss';
const LOCK_WAIT_MS = 30000;


// ===== 設置・確認用（手動実行） =====

/**
 * 初回と月替わりに1回だけ手動実行（設定ファイルの Apps Script で）
 * この設定ファイルを覚え、フォーム送信時トリガー（GAS_1）と深夜1時トリガー（GAS_2）を作成
 */
function setupTriggers() {

  const cfg = loadConfig_();
  const inputSs = openBlock_(cfg, cfg.input);

  // トリガーから動くときも同じ設定ファイルを読むように覚えておく
  PropertiesService.getScriptProperties().setProperty(SETTINGS_ID_PROPERTY, cfg.settingsId);

  // 既に同じトリガーがある場合は削除して、重複起動を防ぐ
  removeTriggers();

  ScriptApp.newTrigger('onFormSubmitHandler')
    .forSpreadsheet(inputSs)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger('nightlyCopyToProduction')
    .timeBased()
    .atHour(NIGHTLY_HOUR)
    .everyDays(1)
    .inTimezone(TIMEZONE)
    .create();

  console.log(`トリガーを作成しました。設定：${cfg.settingsName}／フォーム送信：${inputSs.getName()}／毎日 ${NIGHTLY_HOUR}時台：本番へコピー`);
}

/**
 * このスクリプトのトリガーを削除（前の月の設定ファイルで実行して止める）
 */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    const fn = trigger.getHandlerFunction();
    if (fn === 'onFormSubmitHandler' || fn === 'nightlyCopyToProduction') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * 設定ファイルの読み込み結果を確認する（データは書き込まない）
 * 設置後や設定ファイルを直したあとに実行し、ログを確認する
 */
function checkSettings() {

  const cfg = loadConfig_();
  const lines = [`設定ファイル：${cfg.settingsName}`];

  [cfg.input, cfg.stageJuchu, cfg.stageIndex, cfg.prodJuchu, cfg.prodIndex].forEach(block => {
    const ss = openBlock_(cfg, block);
    lines.push(`■ ${block.label}：${block.file} → ${ss.getUrl()}`);
    if (block !== cfg.input && block !== cfg.prodJuchu) {
      const sheet = sheetOf_(ss, block);
      lines.push(`   シート「${sheet.getName()}」次の書込み行：${findWriteRow_(sheet)}`);
    }
    if (block.fields.length) {
      lines.push('   ' + block.fields.map(f =>
        `${columnLetter_(f.dest)}:${f.name}←${f.sources.map(columnLetter_).join(',')}`).join('  '));
    }
  });
  lines.push(`担当名変換テーブル：${Object.keys(cfg.staff).length} 件`);
  lines.push(`今日の受注一覧シート：${juchuSheetName_(new Date().getMonth() + 1, new Date())}`);

  console.log(lines.join('\n'));
}

/**
 * ❶ で copyflag が空の行をまとめて確認用へ書き込む（手動実行）
 * トリガーが止まっていた間の回答を取り込むときに使う
 */
function processPendingFormRows() {

  withLock_(() => {
    const cfg = loadConfig_();
    const sheet = sheetOf_(openBlock_(cfg, cfg.input), cfg.input);
    const flagCol = cfg.input.flagCol;
    const last = sheet.getLastRow();
    let count = 0;

    for (let row = 2; row <= last; row++) {
      const values = sheet.getRange(row, 1, 1, flagCol).getValues()[0];
      if (isBlank_(values[0]) || !isBlank_(values[flagCol - 1])) continue;
      processInputRow_(cfg, sheet, row);
      count++;
    }
    console.log(`確認用へ ${count} 件書き込みました。`);
  });
}


// ===== GAS_1：工事受注入力 → 事前確認用 =====

/**
 * フォーム送信時トリガーから呼ばれる
 */
function onFormSubmitHandler(e) {

  try {
    withLock_(() => {
      const cfg = loadConfig_();
      const sheet = e.range.getSheet();

      // 設定ファイルのシート名以外（古い回答シートなど）は対象外
      if (cfg.input.sheet && sheet.getName() !== cfg.input.sheet) return;

      processInputRow_(cfg, sheet, e.range.getRow());
    });
  } catch (err) {
    notifyError_('GAS_1（フォーム → 確認用）でエラー', err);
    throw err;
  }
}

/**
 * ❶ の1行を ❷・❸ に書き込み、❶ の copyflag に日時を入れる
 */
function processInputRow_(cfg, inputSheet, row) {

  const flagCol = cfg.input.flagCol;
  const values = inputSheet.getRange(row, 1, 1, flagCol).getValues()[0];

  // 書込み済み（copyflag あり）の行は二重に書かない
  if (!isBlank_(values[flagCol - 1])) return;

  [cfg.stageJuchu, cfg.stageIndex].forEach(stage => {
    const sheet = sheetOf_(openBlock_(cfg, stage), stage);
    const cells = stage.fields.map(f => ({ col: f.dest, value: convertValue_(f, values, cfg) }));
    writeCells_(sheet, findWriteRow_(sheet), cells);
  });

  inputSheet.getRange(row, flagCol).setValue(new Date()).setNumberFormat(FLAG_FORMAT);
  SpreadsheetApp.flush();
}

/**
 * 設定ファイルのルールに従って、確認用に書く値を作る
 *  1) タイムスタンプ → 月・日・日付
 *  2) メールアドレス → 担当名（無ければ「無し」）
 *  3) 受注区分 E〜H → 元下（回答の値）・受注先（回答した列の区分名）
 *  6) 金額は数値にする（カンマ付きの文字でも可）
 */
function convertValue_(field, rowValues, cfg) {

  const srcValues = field.sources.map(c => rowValues[c - 1]);
  const srcNames = field.sources.map(c => cfg.input.namesByCol[c] || '');

  // 3) 受注区分：左の列から最初に回答がある列を使う
  if (field.sources.length > 1) {
    const k = srcValues.findIndex(v => !isBlank_(v));
    if (k < 0) return '';
    return field.name === '受注先' ? categoryName_(srcNames[k]) : String(srcValues[k]).trim();
  }

  const value = srcValues[0];
  const srcName = srcNames[0];

  if (srcName.indexOf('タイムスタンプ') === 0) {
    const d = toDate_(value);
    if (!d) return value;
    if (field.name === '月') return d.getMonth() + 1;
    if (field.name === '日') return d.getDate();
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  if (srcName.indexOf('メールアドレス') === 0) {
    return cfg.staff[String(value).trim().toLowerCase()] || NO_STAFF;
  }

  if (srcName.indexOf('金額') >= 0) return toNumber_(value);

  return value;
}


// ===== GAS_2：事前確認用 → 本番 =====

/**
 * 設定ファイルを開いたときにメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カスタムメニュー')
    .addItem('コピー実行（確認_受注一覧 → 受注一覧）', 'menuCopyJuchu')
    .addItem('コピー実行（確認_工事インデックス → 工事インデックス）', 'menuCopyIndex')
    .addItem('コピー実行（両方）', 'menuCopyBoth')
    .addToUi();
}

function menuCopyJuchu() { runMenuCopy_(cfg => [cfg.stageJuchu]); }
function menuCopyIndex() { runMenuCopy_(cfg => [cfg.stageIndex]); }
function menuCopyBoth() { runMenuCopy_(cfg => [cfg.stageJuchu, cfg.stageIndex]); }

/**
 * メニューの「コピー実行」：選んだ確認用ファイルの分を本番へコピーし、結果を表示
 */
function runMenuCopy_(pickStages) {

  const ui = SpreadsheetApp.getUi();

  try {
    const message = withLock_(() => {
      const cfg = loadConfig_();
      return pickStages(cfg).map(stage => resultMessage_(stage, copyToProduction_(cfg, stage))).join('\n\n');
    });
    ui.alert(message);
  } catch (err) {
    ui.alert(`コピーできませんでした。\n${err.message}`);
  }
}

/**
 * 深夜1時トリガー：確認用2ファイルとも本番へコピー
 */
function nightlyCopyToProduction() {

  try {
    const messages = withLock_(() => {
      const cfg = loadConfig_();
      return [cfg.stageJuchu, cfg.stageIndex].map(stage => {
        const result = copyToProduction_(cfg, stage);
        return result.errors.length ? resultMessage_(stage, result) : '';
      });
    }).filter(m => m);

    if (messages.length) notifyError_('GAS_2（深夜の本番コピー）で一部コピーできない行がありました', new Error(messages.join('\n\n')));
  } catch (err) {
    notifyError_('GAS_2（深夜の本番コピー）でエラー', err);
    throw err;
  }
}

/**
 * 確認用1ファイルの copyflag が空の行を本番へ書き込む
 * @return {{copied: number, errors: string[]}}
 */
function copyToProduction_(cfg, stage) {

  const prod = stage === cfg.stageJuchu ? cfg.prodJuchu : cfg.prodIndex;
  const stageSheet = sheetOf_(openBlock_(cfg, stage), stage);
  const prodSs = openBlock_(cfg, prod);
  const flagCol = stage.flagCol;

  const headerRow = findHeaderRow_(stageSheet, flagCol);
  const last = stageSheet.getLastRow();
  const result = { copied: 0, errors: [] };
  if (last <= headerRow) return result;

  const width = Math.max(flagCol, ...prod.fields.map(f => f.sources[0]));
  const rows = stageSheet.getRange(headerRow + 1, 1, last - headerRow, width).getValues();
  const monthCol = stage === cfg.stageJuchu ? stage.colByName['月'] : null;

  rows.forEach((values, i) => {
    const row = headerRow + 1 + i;
    if (!isBlank_(values[flagCol - 1])) return;
    if (prod.fields.every(f => isBlank_(values[f.sources[0] - 1]))) return;

    try {
      // 受注一覧は期と月のシート、工事インデックスは設定のシート（空なら先頭）
      const sheet = monthCol
        ? requireSheet_(prodSs, juchuSheetName_(values[monthCol - 1], new Date()))
        : sheetOf_(prodSs, prod);
      const cells = prod.fields.map(f => ({ col: f.dest, value: values[f.sources[0] - 1] }));

      writeCells_(sheet, findWriteRow_(sheet), cells);
      stageSheet.getRange(row, flagCol).setValue(new Date()).setNumberFormat(FLAG_FORMAT);
      SpreadsheetApp.flush();
      result.copied++;
    } catch (err) {
      result.errors.push(`${row}行目：${err.message}`);
    }
  });

  return result;
}

/**
 * 受注一覧のシート名（例：27期11月）
 * 確認用には年が無いため、月が今月より後なら前年の受注とみなす
 */
function juchuSheetName_(month, now) {

  const m = Number(month);
  if (!(m >= 1 && m <= 12)) throw new Error(`月の値「${month}」からシートを決められません`);

  const year = m > now.getMonth() + 1 ? now.getFullYear() - 1 : now.getFullYear();
  const ki = (m >= 7 ? year + 1 : year) - 2000;   // 27期＝2026年7月〜2027年6月
  return `${ki}期${m}月`;
}

function resultMessage_(stage, result) {
  const lines = [`${stage.file}：本番へ ${result.copied} 件コピーしました。`];
  if (result.errors.length) lines.push('コピーできなかった行：', ...result.errors);
  return lines.join('\n');
}


// ===== 書込み =====

/**
 * 5) 書込み行を探す：A〜H 列が空（値も数式も無い）の行で、その下2行も空の最初の行
 */
function findWriteRow_(sheet) {

  const last = Math.max(sheet.getLastRow(), SEARCH_START_ROW);
  const height = last - SEARCH_START_ROW + 2 + BLANK_ROWS_BELOW;
  ensureRows_(sheet, SEARCH_START_ROW + height - 1);

  const range = sheet.getRange(SEARCH_START_ROW, 1, height, BLANK_CHECK_LAST_COL);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const blank = values.map((r, i) => r.every((v, j) => isBlank_(v) && formulas[i][j] === ''));

  return SEARCH_START_ROW + findBlankRun_(blank, BLANK_ROWS_BELOW + 1);
}

/**
 * 連続して空の行が need 行続く最初の位置（0始まり）
 */
function findBlankRun_(blank, need) {
  let run = 0;
  for (let i = 0; i < blank.length; i++) {
    run = blank[i] ? run + 1 : 0;
    if (run === need) return i - need + 1;
  }
  return blank.length;
}

/**
 * 指定の列だけに書き込む（間の列は触らないので、数式や手入力の列を消さない）
 */
function writeCells_(sheet, row, cells) {
  groupRuns_(cells).forEach(run => {
    sheet.getRange(row, run[0].col, 1, run.length).setValues([run.map(c => c.value)]);
  });
}

/**
 * 列番号が連続する単位にまとめる（書込み回数を減らす）
 */
function groupRuns_(cells) {
  const sorted = cells.slice().sort((a, b) => a.col - b.col);
  const runs = [];
  sorted.forEach(c => {
    const run = runs[runs.length - 1];
    if (run && run[run.length - 1].col + 1 === c.col) run.push(c);
    else runs.push([c]);
  });
  return runs;
}

function ensureRows_(sheet, needRows) {
  const max = sheet.getMaxRows();
  if (max < needRows) sheet.insertRowsAfter(max, needRows - max);
}

/**
 * copyflag 列に「copyflag」と書かれた見出し行を探す（見つからなければ1行目）
 */
function findHeaderRow_(sheet, flagCol) {
  const scan = Math.min(sheet.getLastRow(), 10);
  if (scan < 1) return 1;
  const values = sheet.getRange(1, flagCol, scan, 1).getDisplayValues();
  const i = values.findIndex(r => String(r[0]).trim() === COPYFLAG_NAME);
  return i >= 0 ? i + 1 : 1;
}


// ===== 設定ファイルの読み込み =====

/**
 * 設定ファイルを読み込み、ファイル情報・列の対応・担当名変換テーブルをまとめる
 */
function loadConfig_() {

  const ss = openSettingsSpreadsheet_();
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME) || ss.getSheets()[0];
  const cfg = parseSettings_(sheet.getDataRange().getDisplayValues());
  cfg.settingsName = ss.getName();
  cfg.settingsId = ss.getId();
  cfg.cache = {};
  return cfg;
}

/**
 * 設定シートの値（2次元配列）を解釈する
 *  ファイル情報❶❷❸ と、❷❸ の直後にある本番の「ファイル情報」を読み分ける
 */
function parseSettings_(rows) {

  const blocks = [];
  rows.forEach((row, i) => {
    const label = String(row[0]).trim();
    if (label.indexOf('ファイル情報') !== 0) return;

    const block = {
      label: label,
      mark: label.replace('ファイル情報', '').trim(),
      drive: String(row[1] || '').trim(),
      file: String(row[3] || '').trim(),
      sheet: String(row[5] || '').trim(),
      url: String(row[6] || '').trim(),   // G列にURLがあればそれを優先（任意）
      names: [], sources: [], dests: []
    };

    for (let j = i + 1; j < rows.length; j++) {
      const head = String(rows[j][0]).trim();
      if (head === '' || head.indexOf('ファイル情報') === 0) break;
      const cells = rows[j].slice(1).map(v => String(v).trim());
      if (head.indexOf('列名') === 0) block.names = cells;
      else if (head.indexOf('元列番号') === 0) block.sources = cells;
      else if (head.indexOf('列番号') === 0) block.dests = cells;
    }
    blocks.push(block);
  });

  const find = mark => {
    const i = blocks.findIndex(b => b.mark === mark);
    if (i < 0) throw new Error(`設定ファイルに「ファイル情報${mark}」がありません`);
    return i;
  };
  const prodAfter = i => {
    const b = blocks[i + 1];
    if (!b || b.mark !== '') throw new Error(`設定ファイルで「${blocks[i].label}」の次に本番の「ファイル情報」がありません`);
    return b;
  };

  const iJ = find('❷');
  const iI = find('❸');
  const cfg = {
    input: blocks[find('❶')],
    stageJuchu: blocks[iJ],
    stageIndex: blocks[iI],
    prodJuchu: prodAfter(iJ),
    prodIndex: prodAfter(iI),
    staff: parseStaffTable_(rows)
  };

  // ❶：列番号ごとの列名と copyflag 列
  cfg.input.namesByCol = {};
  cfg.input.dests.forEach((d, k) => {
    const col = columnNumber_(d);
    if (col) cfg.input.namesByCol[col] = cfg.input.names[k] || '';
  });
  cfg.input.fields = [];
  cfg.input.flagCol = flagColumn_(cfg.input, true);

  [cfg.stageJuchu, cfg.stageIndex, cfg.prodJuchu, cfg.prodIndex].forEach(b => {
    b.fields = buildFields_(b);
  });

  // 確認用：列名の並び＝A列からの並び（copyflag は列番号の行に無いため列名の位置で決める）
  [cfg.stageJuchu, cfg.stageIndex].forEach(b => {
    b.flagCol = flagColumn_(b, false);
    b.colByName = {};
    b.fields.forEach(f => { b.colByName[f.name] = f.dest; });
  });

  return cfg;
}

/**
 * 「列番号」がある位置だけを書込み対象にする（元列番号が空・--- の位置は除く）
 */
function buildFields_(block) {
  const fields = [];
  block.dests.forEach((d, k) => {
    const dest = columnNumber_(d);
    const sources = String(block.sources[k] || '').split(',').map(s => columnNumber_(s.trim())).filter(c => c);
    if (!dest || !sources.length) return;
    fields.push({ name: block.names[k] || '', dest: dest, sources: sources });
  });
  return fields;
}

function flagColumn_(block, useDests) {
  const k = block.names.indexOf(COPYFLAG_NAME);
  if (k < 0) throw new Error(`設定ファイルの「${block.label}」の列名に ${COPYFLAG_NAME} がありません`);
  return useDests ? columnNumber_(block.dests[k]) : k + 1;
}

/**
 * 担当名変換テーブル（A列：名前、B列：メールアドレス）
 */
function parseStaffTable_(rows) {
  const staff = {};
  const start = rows.findIndex(r => String(r[0]).trim() === STAFF_TABLE_TITLE);
  if (start < 0) throw new Error(`設定ファイルに「${STAFF_TABLE_TITLE}」がありません`);

  for (let i = start + 2; i < rows.length; i++) {
    const name = String(rows[i][0]).trim();
    const mail = String(rows[i][1]).trim().toLowerCase();
    if (name === '' && mail === '') break;
    if (mail) staff[mail] = name;
  }
  return staff;
}


// ===== ファイルを開く =====

/**
 * 設定ファイルを開く：コードの SETTINGS_FILE_ID → setupTriggers で覚えたID → 開いている設定ファイル の順
 */
function openSettingsSpreadsheet_() {
  const id = SETTINGS_FILE_ID || PropertiesService.getScriptProperties().getProperty(SETTINGS_ID_PROPERTY);
  if (id) return SpreadsheetApp.openById(id);

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active && active.getSheetByName(SETTINGS_SHEET_NAME)) return active;
  throw new Error('設定ファイルが分かりません。設定ファイルの Apps Script で setupTriggers を実行してください');
}

/**
 * 4) ドライブ名・ファイル名からスプレッドシートを開く（1回の実行の中では使い回す）
 */
function openBlock_(cfg, block) {
  if (!cfg.cache[block.label + block.file]) {
    const id = block.url ? idFromUrl_(block.url) : findSpreadsheetIdByName_(block.file, block.drive);
    cfg.cache[block.label + block.file] = SpreadsheetApp.openById(id);
  }
  return cfg.cache[block.label + block.file];
}

/**
 * 同じ名前のスプレッドシートが複数あるときは、ドライブ名の最後の部分
 * （例：010_G共通フォルダー¥010_営業管理部 → 010_営業管理部）と親フォルダー名が合うものを選ぶ
 */
function findSpreadsheetIdByName_(name, drive) {

  const q = `title = '${name.replace(/'/g, "\\'")}' and mimeType = '${MimeType.GOOGLE_SHEETS}' and trashed = false`;
  const it = DriveApp.searchFiles(q);
  const found = [];
  while (it.hasNext()) found.push(it.next());

  if (found.length === 1) return found[0].getId();
  if (found.length === 0) {
    throw new Error(`「${name}」が見つかりません。Google スプレッドシート形式で、${drive || 'ドライブ'} にあるか確認してください`);
  }

  const lastPart = String(drive).split(/[¥\\\/]/).pop().trim();
  const matched = found.filter(f => {
    const parents = f.getParents();
    while (parents.hasNext()) if (parents.next().getName() === lastPart) return true;
    return false;
  });
  if (matched.length === 1) return matched[0].getId();

  throw new Error(`「${name}」という名前のスプレッドシートが ${found.length} 件あり、どれか決められません。設定ファイルのG列にURLを書いてください`);
}

function idFromUrl_(url) {
  const m = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : url;
}

/**
 * 設定のシート名があればそのシート、空なら先頭のシート
 */
function sheetOf_(ss, block) {
  return block.sheet ? requireSheet_(ss, block.sheet) : ss.getSheets()[0];
}

function requireSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`「${ss.getName()}」にシート「${name}」がありません`);
  return sheet;
}


// ===== 共通の小物 =====

/**
 * 処理が重ならないようにロックを掛けて実行（フォーム同時送信・深夜とボタンの重なり対策）
 */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_WAIT_MS);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function notifyError_(subject, err) {
  console.error(`${subject}\n${err && err.stack || err}`);
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), `【工事受注転記】${subject}`, String(err && err.message || err));
  } catch (mailErr) {
    console.error(`通知メールを送れませんでした：${mailErr}`);
  }
}

function isBlank_(v) {
  return v === '' || v === null || v === undefined;
}

/**
 * 「受注区分 [官公庁]」→「官公庁」
 */
function categoryName_(header) {
  const m = String(header).match(/\[(.+?)\]/);
  return m ? m[1].trim() : String(header).trim();
}

function toDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v) ? null : v;
  const m = String(v).match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

/**
 * 「270,000」「２７００００」なども数値にする。数値にできなければ元の値のまま
 */
function toNumber_(v) {
  if (typeof v === 'number') return v;
  const s = String(v)
    .replace(/[０-９．，－]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[,\s円]/g, '');
  if (s === '') return '';
  const n = Number(s);
  return isNaN(n) ? v : n;
}

function columnNumber_(letter) {
  const s = String(letter || '').trim().toUpperCase();
  if (!/^[A-Z]{1,2}$/.test(s)) return 0;
  return s.split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
}

function columnLetter_(col) {
  let s = '';
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}
