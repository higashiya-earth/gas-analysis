/**
 * 受注データ連携 GAS（設定ファイル_v08 用）
 *
 * 設定ファイル_v08 の内容に従って動きます。ファイル名・シート名・列の対応・担当名は
 * コードに書かず、毎回設定ファイルから読み込みます。
 *
 * 【GAS_Pri】File1 → File2・File4（設定の規則 1)〜5)）
 *  フォームから File1 に行が追加されたら、その行を File2（確認_受注一覧）と
 *  File4（確認_工事インデックス）の空いた行へ書き込み、両方に書けたら Copyflag1 に日時を入れる。
 *
 * 【GAS_Sec】File2 → File3、File4 → File5・File6・File7
 *  メニュー「カスタムメニュー ＞ コピー実行」、または毎日深夜1時台に、
 *  Copyflag2・Copyflag4 が空の行を本番へ書き込み、書けた行の Copyflag に日時を入れる。
 *  File3 は「27期11月」のような期と月のシートへ（期は7月始まり）。
 *  当初受注金額が 1,000,000 以上の行は File6（注文書・注文請書チェックリスト）へも書く。
 *  File7（工事現場詳細）へは毎回書き、現場詳細の列は折り返して行の高さを内容に合わせる。
 *  Copyflag4 は File5・File6（対象の行のみ）・File7 がすべて書けたときに付ける。
 *
 * 【設置】このコードを設定ファイルの Apps Script に貼り、setupTriggers を1回実行する（README.md 参照）。
 */


// ===== 共通設定 =====

// 設定ファイル_v08 のID（URL の /d/ と /edit の間）。別の設定ファイルを使うときはここを変える
// 空にすると、setupTriggers で覚えたID、またはこのスクリプトを入れた設定ファイルを使う
const SETTINGS_FILE_ID = '1NJEMrTL5u0Zs29uAnvP-1I7R_Z3Qn8MTWxIFHonu-Jc';

// この GAS の版。メニューの結果表示と「設定の確認」に出るので、貼り替えたコードが動いているか確かめられる
const SCRIPT_VERSION = 'v08-File7対応 2026-11-13b';
const SETTINGS_ID_PROPERTY = 'SETTINGS_FILE_ID_V08';

const SETTINGS_SHEET_NAME = '設定ファイル';
const STAFF_TABLE_TITLE = '担当名変換テーブル';
const NO_STAFF = '無し';

// File4 → File6 の条件：当初受注金額がこの額以上
const FILE6_MIN_AMOUNT = 1000000;

// 書込み行：判定する列（A列から）が空の行で、その下2行も空の最初の行
const BLANK_COLS_DEFAULT = 8;   // H列まで
const BLANK_COLS_FILE6 = 6;     // File6 は F列まで
const BLANK_COLS_FILE7 = 7;     // File7 は G列まで

// File7 で折り返して行の高さを合わせる列の名前
const WRAP_FIELD_FILE7 = '現場詳細';
const BLANK_ROWS_BELOW = 2;
const SEARCH_START_ROW = 2;

const NIGHTLY_HOUR = 1;
// 日付・Copyflag は日本時間の文字列で書く（各ファイルのタイムゾーン設定が違っても日付がずれないように）
const TIMEZONE = 'Asia/Tokyo';
const FLAG_FORMAT = 'yyyy/MM/dd HH:mm:ss';
const LOCK_WAIT_MS = 30000;
const TRIGGER_FUNCTIONS = ['onFormSubmitV08', 'nightlyCopyV08'];


// ===== 設置・確認用（手動実行） =====

/**
 * 初回に1回だけ手動実行（設定ファイルの Apps Script で）
 * この設定ファイルを覚え、フォーム送信時トリガー（GAS_Pri）と深夜1時トリガー（GAS_Sec）を作成
 */
function setupTriggers() {

  const cfg = loadConfig_();
  const file1 = openFile_(cfg, 1);

  // トリガーから動くときも同じ設定ファイルを読むように覚えておく
  PropertiesService.getScriptProperties().setProperty(SETTINGS_ID_PROPERTY, cfg.settingsId);

  // 既に同じトリガーがある場合は削除して、重複起動を防ぐ
  removeTriggers();

  ScriptApp.newTrigger('onFormSubmitV08')
    .forSpreadsheet(file1)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger('nightlyCopyV08')
    .timeBased()
    .atHour(NIGHTLY_HOUR)
    .everyDays(1)
    .inTimezone(TIMEZONE)
    .create();

  console.log(`トリガーを作成しました。設定：${cfg.settingsName}／フォーム送信：${file1.getName()}／毎日 ${NIGHTLY_HOUR}時台：コピー実行`);
}

/**
 * このスクリプトのトリガーを削除（使うのをやめるとき、前の版を止めるとき）
 */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (TRIGGER_FUNCTIONS.indexOf(trigger.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(trigger);
  });
}

/**
 * 設定ファイルの読み込み結果を確認する（データは書き込まない）
 */
function checkSettings() {
  console.log(settingsReport_().join('\n'));
}

function settingsReport_() {

  const cfg = loadConfig_();
  const lines = [`GAS：${SCRIPT_VERSION}`, `設定ファイル：${cfg.settingsName}`];
  const jst = Utilities.formatDate(new Date(), TIMEZONE, 'Z');

  fileNumbers_(cfg).forEach(n => {
    const f = cfg.files[n];
    const ss = openFile_(cfg, n);
    lines.push(`■ File${n}：${f.file}（${f.drive}）→ ${ss.getUrl()}`);
    const tz = ss.getSpreadsheetTimeZone();
    if (Utilities.formatDate(new Date(), tz, 'Z') !== jst) {
      lines.push(`   ※タイムゾーンが ${tz}。表示がずれるので「ファイル ＞ 設定 ＞ タイムゾーン」を (GMT+09:00) 東京 にしてください`);
    }
    if (n !== 3) {
      const sheet = sheetOf_(ss, f);
      const head = findNamesRow_(sheet, f);
      lines.push(`   シート：${sheet.getName()}／見出し：${head}行目` +
        (n === 1 ? '' : `／次の書込み行：${findWriteRow_(sheet, blankColsOf_(n), head + 1)}行目`));
    }
    if (f.fields.length) {
      lines.push('   ' + f.fields.map(x =>
        `${columnLetter_(x.dest)}:${x.name}←${x.sources.length ? x.sources.map(columnLetter_).join(',') : '（規則）'}`).join('  '));
    }
  });
  lines.push(`担当名変換テーブル：${Object.keys(cfg.staff).length} 件`);
  const today = dateParts_(dateToText_(new Date(), TIMEZONE));
  lines.push(`今日の File3 シート：${periodSheetName_(today.y, today.m)}`);
  return lines;
}

/**
 * File1 で Copyflag1 が空の行をまとめて File2・File4 へ書き込む（手動実行）
 * トリガーが止まっていた間の回答を取り込むときに使う
 */
function processPendingFormRows() {

  withLock_(() => {
    const cfg = loadConfig_();
    const sheet = sheetOf_(openFile_(cfg, 1), cfg.files[1]);
    const flagCol = cfg.files[1].flagCol;
    const last = sheet.getLastRow();
    let count = 0;

    for (let row = 2; row <= last; row++) {
      const values = sheet.getRange(row, 1, 1, flagCol).getValues()[0];
      if (isBlank_(values[0]) || !isBlank_(values[flagCol - 1])) continue;
      processFormRow_(cfg, sheet, row);
      count++;
    }
    console.log(`File2・File4 へ ${count} 件書き込みました。`);
  });
}


// ===== GAS_Pri：File1 → File2・File4 =====

/**
 * フォーム送信時トリガーから呼ばれる
 */
function onFormSubmitV08(e) {

  try {
    withLock_(() => {
      const cfg = loadConfig_();
      const sheet = e.range.getSheet();

      // 設定のシート名以外（古い回答シートなど）は対象外
      if (cfg.files[1].sheet && sheet.getName() !== cfg.files[1].sheet) return;

      processFormRow_(cfg, sheet, e.range.getRow());
    });
  } catch (err) {
    notifyError_('GAS_Pri（File1 → File2・File4）でエラー', err);
    throw err;
  }
}

/**
 * File1 の1行を File2・File4 に書き込み、Copyflag1 に日時を入れる
 */
function processFormRow_(cfg, file1Sheet, row) {

  const f1 = cfg.files[1];
  const tz1 = file1Sheet.getParent().getSpreadsheetTimeZone();
  const values = file1Sheet.getRange(row, 1, 1, f1.lastCol).getValues()[0].map(v => dateToText_(v, tz1));

  // 書込み済み（Copyflag1 あり）の行は二重に書かない
  if (!isBlank_(values[f1.flagCol - 1])) return;

  [2, 4].forEach(n => {
    const f = cfg.files[n];
    const sheet = sheetOf_(openFile_(cfg, n), f);
    const cells = f.fields
      .filter(x => x.name !== f.flagName)
      .map(x => ({ col: x.dest, value: convertFormValue_(x, values, cfg) }))
      .filter(c => c.value !== undefined);
    writeCells_(sheet, findWriteRow_(sheet, BLANK_COLS_DEFAULT, findNamesRow_(sheet, f) + 1), cells);
  });

  file1Sheet.getRange(row, f1.flagCol).setValue(stamp_());
  SpreadsheetApp.flush();
}

/**
 * 規則 1)〜5) に従って、File1 の値から確認用に書く値を作る
 * 書かない項目（転記元が無い列）は undefined を返す
 */
function convertFormValue_(field, row, cfg) {

  const f1 = cfg.files[1];
  const name = field.name;
  const srcNames = field.sources.map(c => f1.namesByCol[c] || '');
  const srcValues = field.sources.map(c => row[c - 1]);

  // 1) 年・月・日・日付はタイムスタンプから（転記元の欄が空でも列名で決める）
  if (['年', '月', '日', '日付'].indexOf(name) >= 0) {
    const d = dateParts_(row[f1.colByName['タイムスタンプ'] - 1]);
    if (!d) return '';
    if (name === '年') return d.y;
    if (name === '月') return d.m;
    if (name === '日') return d.d;
    return `${d.y}/${pad2_(d.m)}/${pad2_(d.d)}`;
  }

  if (!field.sources.length) return undefined;

  // 5) 入力確認：「C)」があれば同一工事に〇、「D)」があれば撮影に〇
  if (srcNames[0] === '入力確認') {
    const text = String(srcValues[0] || '');
    if (name === '同一工事') return hasChoice_(text, 'C') ? '○' : '';
    if (name === '撮影') return hasChoice_(text, 'D') ? '○' : '';
    return text;
  }

  // 2) 担当：File1 の担当（C列）を優先、無ければメールアドレスを担当名に変換、無ければ「無し」
  const mailIdx = srcNames.indexOf('メールアドレス');
  if (mailIdx >= 0) {
    const direct = srcValues.find((v, i) => i !== mailIdx && !isBlank_(v));
    if (direct !== undefined) return String(direct).trim();
    return cfg.staff[String(srcValues[mailIdx]).trim().toLowerCase()] || NO_STAFF;
  }

  // 4) 受注区分：回答が1つだけのとき、元下＝その値、受注先＝その列の区分名。0個・複数は空
  if (srcNames.every(n => n.indexOf('受注区分') === 0) && field.sources.length > 1) {
    const answered = srcValues.map((v, i) => ({ v, i })).filter(x => !isBlank_(x.v));
    if (answered.length !== 1) return '';
    const k = answered[0].i;
    return name === '受注先' ? categoryName_(srcNames[k]) : String(srcValues[k]).trim();
  }

  // 金額は数値にする
  if (srcNames[0].indexOf('金額') >= 0) return toNumber_(srcValues[0]);

  return srcValues[0];
}

/**
 * 入力確認の文字列に「C)」のような選択肢記号があるか
 */
function hasChoice_(text, letter) {
  return new RegExp('(^|[\\s,、，])' + letter + '[)）]').test(text);
}


// ===== GAS_Sec：File2 → File3、File4 → File5・File6 =====

/**
 * 設定ファイル（またはこのスクリプトを入れたファイル）を開いたときにメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カスタムメニュー')
    .addItem('コピー実行', 'menuCopyV08')
    .addItem('設定の確認', 'menuCheckSettingsV08')
    .addToUi();
}

/**
 * メニュー「設定の確認」：checkSettings の結果を画面に出す（データは書き込まない）
 */
function menuCheckSettingsV08() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert(settingsReport_().join('\n'));
  } catch (err) {
    ui.alert(`設定を読めませんでした。\n${err.message}`);
  }
}

/**
 * メニュー「コピー実行」：File2→File3、File4→File5・File6 をまとめて実行し、結果を表示
 */
function menuCopyV08() {

  const ui = SpreadsheetApp.getUi();
  try {
    const message = withLock_(() => {
      const cfg = loadConfig_();
      return [`（GAS ${SCRIPT_VERSION}）`].concat([copyFile2_(cfg), copyFile4_(cfg)].map(resultMessage_)).join('\n\n');
    });
    ui.alert(message);
  } catch (err) {
    ui.alert(`コピーできませんでした。\n${err.message}`);
  }
}

/**
 * 深夜1時トリガー：メニューの「コピー実行」と同じ処理。コピーできない行があればメールで知らせる
 */
function nightlyCopyV08() {

  try {
    const messages = withLock_(() => {
      const cfg = loadConfig_();
      return [copyFile2_(cfg), copyFile4_(cfg)].filter(r => r.errors.length).map(resultMessage_);
    });
    if (messages.length) notifyError_('GAS_Sec（深夜のコピー実行）でコピーできない行がありました', new Error(messages.join('\n\n')));
  } catch (err) {
    notifyError_('GAS_Sec（深夜のコピー実行）でエラー', err);
    throw err;
  }
}

/**
 * File2 の Copyflag2 が空の行を File3 の期・月のシートへ書く
 */
function copyFile2_(cfg) {

  const f2 = cfg.files[2], f3 = cfg.files[3];
  const file3 = openFile_(cfg, 3);
  const yearCol = f2.colByName['年'], monthCol = f2.colByName['月'];

  return eachPendingRow_(cfg, 2, f3.fields, (values) => {
    const sheet = requireSheet_(file3, periodSheetName_(values[yearCol - 1], values[monthCol - 1]));
    const cells = f3.fields.map(x => ({ col: x.dest, value: values[x.sources[0] - 1] }));
    writeCells_(sheet, findWriteRow_(sheet, BLANK_COLS_DEFAULT, findNamesRow_(sheet, f3) + 1), cells);
  }, 'File2 → File3');
}

/**
 * File4 の Copyflag4 が空の行を File5・File7 へ、当初受注金額が基準以上なら File6 へも書く
 * すべて（File6 は対象の行だけ）書けたときに Copyflag4 を付ける。
 * 途中で止まった行や Copyflag4 を消した行をやり直しても二重にならないよう、
 * File5・File6・File7 とも同じ行がすでにあれば書かない
 */
function copyFile4_(cfg) {

  const f4 = cfg.files[4], f5 = cfg.files[5], f6 = cfg.files[6], f7 = cfg.files[7];
  const sheet5 = sheetOf_(openFile_(cfg, 5), f5);
  const sheet6 = sheetOf_(openFile_(cfg, 6), f6);
  const sheet7 = f7 ? sheetOf_(openFile_(cfg, 7), f7) : null;
  const head5 = findNamesRow_(sheet5, f5), head6 = findNamesRow_(sheet6, f6);
  const head7 = sheet7 ? findNamesRow_(sheet7, f7) : 0;
  const amountCol = f4.colByName['当初受注金額'];
  const targets = f5.fields.concat(f6.fields, f7 ? f7.fields : []);

  return eachPendingRow_(cfg, 4, targets, (values) => {
    if (toNumber_(values[amountCol - 1]) >= FILE6_MIN_AMOUNT) {
      const cells6 = f6.fields.map(x => ({ col: x.dest, value: amountOrValue_(f4, x.sources[0], values) }));
      if (!rowExists_(sheet6, head6, cells6)) writeCells_(sheet6, findWriteRow_(sheet6, BLANK_COLS_FILE6, head6 + 1), cells6);
    }
    if (sheet7) {
      const cells7 = f7.fields.map(x => ({ col: x.dest, value: values[x.sources[0] - 1] }));
      if (!rowExists_(sheet7, head7, cells7)) {
        const row7 = findWriteRow_(sheet7, BLANK_COLS_FILE7, head7 + 1);
        writeCells_(sheet7, row7, cells7);
        fitWrappedRow_(sheet7, row7, f7.colByName[WRAP_FIELD_FILE7]);
      }
    }
    const cells5 = f5.fields.map(x => ({ col: x.dest, value: values[x.sources[0] - 1] }));
    if (!rowExists_(sheet5, head5, cells5)) writeCells_(sheet5, findWriteRow_(sheet5, BLANK_COLS_DEFAULT, head5 + 1), cells5);
  }, f7 ? 'File4 → File5・File6・File7' : 'File4 → File5・File6');
}

/**
 * 現場詳細の列を折り返し表示にし、行の高さを内容に合わせる
 */
function fitWrappedRow_(sheet, row, col) {
  if (!col) return;
  sheet.getRange(row, col).setWrap(true).setVerticalAlignment('top');
  SpreadsheetApp.flush();
  sheet.autoResizeRows(row, 1);
}

/**
 * 確認用ファイル（File2・File4）の Copyflag が空でデータのある行ごとに write を実行し、
 * 成功した行の Copyflag に日時を入れる
 * @return {{label: string, copied: number, errors: string[]}}
 */
function eachPendingRow_(cfg, n, targetFields, write, label) {

  const f = cfg.files[n];
  const sheet = sheetOf_(openFile_(cfg, n), f);
  const flagCol = f.flagCol;
  const headerRow = findHeaderRow_(sheet, flagCol, f.flagName);
  const last = sheet.getLastRow();
  const result = { label, copied: 0, errors: [] };
  if (last <= headerRow) return result;

  const width = Math.max(flagCol, ...targetFields.map(x => x.sources[0] || 1));
  const tz = sheet.getParent().getSpreadsheetTimeZone();
  const rows = sheet.getRange(headerRow + 1, 1, last - headerRow, width).getValues()
    .map(r => r.map(v => dateToText_(v, tz)));

  rows.forEach((values, i) => {
    const row = headerRow + 1 + i;
    if (!isBlank_(values[flagCol - 1])) return;
    if (targetFields.every(x => !x.sources[0] || isBlank_(values[x.sources[0] - 1]))) return;

    try {
      write(values);
      sheet.getRange(row, flagCol).setValue(stamp_());
      SpreadsheetApp.flush();
      result.copied++;
    } catch (err) {
      result.errors.push(`${row}行目：${err.message}`);
    }
  });
  return result;
}

function amountOrValue_(f, col, values) {
  const v = values[col - 1];
  return (f.namesByCol[col] || '').indexOf('金額') >= 0 ? toNumber_(v) : v;
}

/**
 * 書こうとしている行が、見出しより下にすでにあるか
 * 指示書No（A列）があれば 指示書No・連番（A・B列）で、空なら書く列すべての値で比べる
 */
function rowExists_(sheet, headRow, cells) {
  const last = sheet.getLastRow();
  if (last <= headRow) return false;
  const width = Math.max(...cells.map(c => c.col));
  const tz = sheet.getParent().getSpreadsheetTimeZone();
  const rows = sheet.getRange(headRow + 1, 1, last - headRow, width).getValues();
  const keyOf = raw => {
    const v = dateToText_(raw, tz);
    const n = toNumber_(v);
    return typeof n === 'number' ? String(n) : normalize_(v);
  };
  const byNo = cells.filter(c => c.col <= 2);
  const useNo = byNo.length === 2 && keyOf(cells.find(c => c.col === 1).value) !== '';
  const compare = useNo ? byNo : cells;
  return rows.some(r => compare.every(c => keyOf(r[c.col - 1]) === keyOf(c.value)));
}

/**
 * File3 のシート名（例：27期11月）。27期＝2026年7月〜2027年6月
 */
function periodSheetName_(year, month) {
  const y = Number(year), m = Number(month);
  if (!(y >= 2000) || !(m >= 1 && m <= 12)) throw new Error(`年「${year}」・月「${month}」からシートを決められません`);
  const ki = (m >= 7 ? y + 1 : y) - 2000;
  return `${ki}期${m}月`;
}

function resultMessage_(result) {
  const lines = [`${result.label}：${result.copied} 件コピーしました。`];
  if (result.errors.length) lines.push('コピーできなかった行：', ...result.errors);
  return lines.join('\n');
}


// ===== 書込み =====

/**
 * 書込み行を探す：A列から blankCols 列が空（値も数式も無い）の行で、その下2行も空の最初の行
 */
function findWriteRow_(sheet, blankCols, startRow) {

  const start = Math.max(startRow || SEARCH_START_ROW, SEARCH_START_ROW);
  const last = Math.max(sheet.getLastRow(), start);
  const height = last - start + 2 + BLANK_ROWS_BELOW;
  ensureRows_(sheet, start + height - 1);

  const range = sheet.getRange(start, 1, height, blankCols);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const blank = values.map((r, i) => r.every((v, j) => isBlank_(v) && formulas[i][j] === ''));

  return start + findBlankRun_(blank, BLANK_ROWS_BELOW + 1);
}

/**
 * 見出し行を探す：上から15行のうち、設定ファイルの列名が半分以上（2つ以上）並んでいる最初の行
 * 見つからなければ 1（1行目を見出しとみなす）
 */
function findNamesRow_(sheet, f) {
  const fields = f.fields.length ? f.fields : Object.keys(f.namesByCol).map(c => ({ dest: Number(c), name: f.namesByCol[c] }));
  const named = fields.filter(x => x.name && x.name !== '---');
  if (!named.length) return 1;
  const scan = Math.min(sheet.getLastRow(), 15);
  if (scan < 1) return 1;
  const width = Math.max(...named.map(x => x.dest));
  const rows = sheet.getRange(1, 1, scan, width).getDisplayValues();
  const key = v => normalize_(v).replace(/[.．\s]/g, '');
  const need = Math.max(2, Math.ceil(named.length / 2));
  const i = rows.findIndex(r => named.filter(x => key(r[x.dest - 1]) === key(x.name)).length >= need);
  return i >= 0 ? i + 1 : 1;
}

function blankColsOf_(n) {
  return n === 6 ? BLANK_COLS_FILE6 : n === 7 ? BLANK_COLS_FILE7 : BLANK_COLS_DEFAULT;
}

function fileNumbers_(cfg) {
  return Object.keys(cfg.files).map(Number).sort((a, b) => a - b);
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
 * Copyflag 列にその名前が書かれた見出し行を探す（見つからなければ1行目）
 */
function findHeaderRow_(sheet, flagCol, flagName) {
  const scan = Math.min(sheet.getLastRow(), 10);
  if (scan < 1) return 1;
  const values = sheet.getRange(1, flagCol, scan, 1).getDisplayValues();
  const i = values.findIndex(r => normalize_(r[0]).toLowerCase() === normalize_(flagName).toLowerCase());
  return i >= 0 ? i + 1 : 1;
}


// ===== 設定ファイルの読み込み =====

/**
 * 設定ファイルを読み込み、File1〜6 の情報・列の対応・担当名変換テーブルをまとめる
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
 *  A列「File1」〜「File6」の行：B列＝ドライブ名、D列＝ファイル名、F列＝シート名（G列にURLがあれば優先）
 *  その下の「列名N」＝項目名、「File列N」（N＝自分の番号）＝書込み先の列、
 *  「File列M」（M＝別の番号）＝File M の転記元の列
 */
function parseSettings_(rows) {

  const files = {};
  rows.forEach((row, i) => {
    const m = String(row[0]).trim().match(/^file\s*(\d)$/i);
    if (!m) return;
    const n = Number(m[1]);
    const f = {
      no: n,
      drive: String(row[1] || '').trim(),
      file: String(row[3] || '').trim(),
      sheet: String(row[5] || '').trim(),
      url: String(row[6] || '').trim(),
      names: [], dests: [], sources: [], sourceFile: null
    };

    for (let j = i + 1; j < rows.length; j++) {
      const head = normalize_(rows[j][0]);
      if (head === '' || /^file\s*\d$/i.test(head)) break;
      const cells = rows[j].slice(1).map(v => normalize_(v));
      const lm = head.match(/^file列\s*(\d)$/i);
      if (/^列名/.test(head)) f.names = cells;
      else if (lm && Number(lm[1]) === n) f.dests = cells;
      else if (lm) { f.sources = cells; f.sourceFile = Number(lm[1]); }
    }
    files[n] = f;
  });

  // File1〜6 は必須、File7（工事現場詳細）は書いてあれば使う
  [1, 2, 3, 4, 5, 6].forEach(n => {
    if (!files[n]) throw new Error(`設定ファイルに「File${n}」の行がありません`);
  });

  Object.keys(files).forEach(k => {
    const f = files[k];
    f.namesByCol = {};
    f.colByName = {};
    f.dests.forEach((d, i) => {
      const col = columnNumber_(d);
      if (!col) return;
      f.namesByCol[col] = f.names[i] || '';
      if (f.names[i]) f.colByName[f.names[i]] = col;
    });
    f.fields = buildFields_(f);
    f.lastCol = Math.max(0, ...Object.keys(f.namesByCol).map(Number));
    const flagName = f.names.find(x => /^copyflag/i.test(x));
    f.flagName = flagName || '';
    f.flagCol = flagName ? f.colByName[flagName] : 0;
  });

  [1, 2, 4].forEach(n => {
    if (!files[n].flagCol) throw new Error(`設定ファイルの File${n} の列名に Copyflag がありません`);
  });

  return { files, staff: parseStaffTable_(rows) };
}

/**
 * 書込み先の列があるものを項目にする。転記元が空・「---」の項目は sources が空になる
 */
function buildFields_(f) {
  const fields = [];
  f.dests.forEach((d, i) => {
    const dest = columnNumber_(d);
    if (!dest) return;
    const sources = String(f.sources[i] || '').split(/[,、，]/).map(s => columnNumber_(s.trim())).filter(c => c);
    fields.push({ name: f.names[i] || '', dest, sources });
  });
  // 自分の列だけの File1 以外は、転記元が無く規則でも決まらない項目を除く
  if (f.no === 1) return [];
  return fields.filter(x => x.sources.length || ['年', '月', '日', '日付'].indexOf(x.name) >= 0);
}

/**
 * 担当名変換テーブル：「担当名変換テーブル」と書かれたセルの1行下が見出し、2行下からデータ
 * （その列＝担当名、右隣の列＝メールアドレス）
 */
function parseStaffTable_(rows) {
  let r0 = -1, c0 = -1;
  rows.some((row, i) => {
    const j = row.findIndex(v => normalize_(v) === STAFF_TABLE_TITLE);
    if (j >= 0) { r0 = i; c0 = j; return true; }
    return false;
  });
  if (r0 < 0) throw new Error(`設定ファイルに「${STAFF_TABLE_TITLE}」がありません`);

  const staff = {};
  for (let i = r0 + 2; i < rows.length; i++) {
    const name = normalize_(rows[i][c0]);
    const mail = normalize_(rows[i][c0 + 1]).toLowerCase();
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
  throw new Error('設定ファイルが分かりません。コード先頭の SETTINGS_FILE_ID に設定ファイルのIDを書くか、設定ファイルの Apps Script で setupTriggers を実行してください');
}

/**
 * File n を開く（1回の実行の中では使い回す）
 */
function openFile_(cfg, n) {
  if (!cfg.cache[n]) {
    const f = cfg.files[n];
    const id = f.url ? idFromUrl_(f.url) : findSpreadsheetIdByName_(f.file, f.drive);
    cfg.cache[n] = SpreadsheetApp.openById(id);
  }
  return cfg.cache[n];
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
function sheetOf_(ss, f) {
  return f.sheet ? requireSheet_(ss, f.sheet) : ss.getSheets()[0];
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
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), `【受注データ連携】${subject}`, String(err && err.message || err));
  } catch (mailErr) {
    console.error(`通知メールを送れませんでした：${mailErr}`);
  }
}

function isBlank_(v) {
  return v === '' || v === null || v === undefined;
}

/**
 * 改行・前後の空白・タブを除く（設定ファイルの「受注区分 \n[元請]」などに対応）
 */
function normalize_(v) {
  return String(v === null || v === undefined ? '' : v).replace(/[\r\n\t]+/g, '').trim();
}

/**
 * 「受注区分 [元請]」→「元請」
 */
function categoryName_(header) {
  const m = String(header).match(/\[(.+?)\]/);
  return m ? m[1].trim() : String(header).trim();
}

/**
 * 「2026/11/13 18:12:46」などの文字列から年・月・日を取り出す
 */
function dateParts_(v) {
  const m = String(v === null || v === undefined ? '' : v).match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/);
  return m ? { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) } : null;
}

/**
 * 日時の値を、そのファイルのタイムゾーンで見たとおりの文字列にする（日付だけなら yyyy/MM/dd）
 * 書き込み先のタイムゾーンが違っても、同じ日付・時刻として読まれる
 */
function dateToText_(v, tz) {
  if (Object.prototype.toString.call(v) !== '[object Date]' || isNaN(v)) return v;
  const t = Utilities.formatDate(v, tz, 'HH:mm:ss');
  return Utilities.formatDate(v, tz, t === '00:00:00' ? 'yyyy/MM/dd' : FLAG_FORMAT);
}

/**
 * Copyflag に入れる日時（日本時間）
 */
function stamp_() {
  return Utilities.formatDate(new Date(), TIMEZONE, FLAG_FORMAT);
}

function pad2_(n) {
  return ('0' + n).slice(-2);
}

/**
 * 「270,000」「２７００００」なども数値にする。数値にできなければ元の値のまま
 */
function toNumber_(v) {
  if (typeof v === 'number') return v;
  const s = String(v === null || v === undefined ? '' : v)
    .replace(/[０-９．，－]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[,\s円¥￥]/g, '');
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
