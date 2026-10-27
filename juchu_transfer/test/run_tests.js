/**
 * juchu_transfer.gs の動作確認（Node で実行：node juchu_transfer/test/run_tests.js）
 *
 * SpreadsheetApp などを簡易的な模擬オブジェクトに置き換え、
 * 実際の設定ファイル_2611 と同じ内容で GAS_1・GAS_2 を通しで動かす。
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

// vm の中で作った配列も比べられるように JSON にしてから比べる
const isDate = v => Object.prototype.toString.call(v) === '[object Date]';
const same = (actual, expected, msg) => assert.deepStrictEqual(JSON.parse(JSON.stringify(actual)), expected, msg);

// ===== 模擬シート =====

class Sheet {
  constructor(name, rows, formulas) {
    this.name = name;
    this.cells = rows.map(r => r.slice());
    this.formulas = formulas || {};          // "行,列" → 式
    this.maxRows = Math.max(this.cells.length, 20);
    this.formats = {};
  }
  getName() { return this.name; }
  getLastRow() {
    let last = 0;
    this.cells.forEach((r, i) => { if (r.some(v => v !== '' && v !== undefined)) last = i + 1; });
    Object.keys(this.formulas).forEach(k => { last = Math.max(last, Number(k.split(',')[0])); });
    return last;
  }
  getMaxRows() { return this.maxRows; }
  getDataRange() {
    const nc = Math.max(1, ...this.cells.map(r => r.length));
    return this.getRange(1, 1, Math.max(1, this.getLastRow()), nc);
  }
  insertRowsAfter(after, n) { this.maxRows += n; }
  get(r, c) { return (this.cells[r - 1] || [])[c - 1] ?? ''; }
  set(r, c, v) {
    while (this.cells.length < r) this.cells.push([]);
    this.cells[r - 1][c - 1] = v;
  }
  getRange(r, c, nr = 1, nc = 1) {
    const sheet = this;
    if (r + nr - 1 > sheet.maxRows) throw new Error(`範囲外：${sheet.name} ${r + nr - 1}行`);
    const grid = fn => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => fn(r + i, c + j)));
    return {
      getValues: () => grid((y, x) => sheet.get(y, x)),
      getDisplayValues: () => grid((y, x) => String(sheet.get(y, x))),
      getFormulas: () => grid((y, x) => sheet.formulas[`${y},${x}`] || ''),
      setValues(vals) {
        assert.strictEqual(vals.length, nr); assert.strictEqual(vals[0].length, nc);
        vals.forEach((row, i) => row.forEach((v, j) => sheet.set(r + i, c + j, v)));
        return this;
      },
      setValue(v) { sheet.set(r, c, v); return this; },
      setNumberFormat(f) { sheet.formats[`${r},${c}`] = f; return this; },
      getSheet: () => sheet,
      getRow: () => r
    };
  }
}

class Spreadsheet {
  constructor(id, name, sheets) { this.id = id; this.name = name; this.sheets = sheets; }
  getId() { return this.id; }
  getName() { return this.name; }
  getUrl() { return `https://docs.google.com/spreadsheets/d/${this.id}`; }
  getSheets() { return this.sheets; }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
}

// ===== 実際の設定ファイル_2611 の内容（2026-10-27 時点） =====

const settingsRows = [
  ['', 'ドライブ名', '', 'ファイル名', '', 'シート名'],
  ['ファイル情報❶', '310_工事受注入力', '', '工事受注入力_2611', '', 'フォームの回答 2', '', '', '❶Googleフォームから入力されたデータ配列'],
  ['列名', 'タイムスタンプ', 'メールアドレス', '客先名', '客先担当者名', '受注区分 [官公庁]', '受注区分 [民間]', '受注区分 [ゼネコン]', '受注区分 [同業]', '種別', '工事名', '工事場所（住所）', '工事内容（説明）', '受注金額（税抜）', '工事時期', 'その他（補足事項）', '重複入力の確認', 'copyflag'],
  ['列番号', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q'],
  [],
  ['', '1) タイムスタンプ …'],
  ['', '2) メールアドレス …'],
  ['', '3) セルF4～I4 …'],
  ['', '4) ドライブ名 …'],
  ['', '5) ❶の行データ …'],
  ['', '　 ❶の行データ …'],
  ['', '6) 受注金額 → 請求金額へ入れる'],
  ['', '7) GASは、2種類 …'],
  ['', '8) GAS_1)起動 …'],
  ['', '9) GAS_2)起動 …'],
  ['', '  GAS_2) コピーが起動したとき …'],
  ['', '　本番に書込むときも上記5) のルールを基本とする'],
  ['', '  本番「受注一覧」シート名 …'],
  ['', '  本番「工事インデックス」は、そのままシートへコピーする'],
  [],
  ['ファイル情報❷', '310_工事受注入力', '', '確認_受注一覧_2611', '', '', '', '', '❶ → ❷受注一覧 事前確認用スプレッド'],
  ['列名', '月', '日', '元請業者', '元下', '受注先', '種別', '場所', '担当', '当初受注金額', 'copyflag'],
  ['元列番号❶', 'A', 'A', 'C', 'E,F,G,H', 'E,F,G,H', 'I', 'K', 'B', 'M', 'J'],
  ['列番号❷', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'],
  [], [],
  ['ファイル情報', '010_G共通フォルダー¥010_営業管理部', '', '受注一覧_2611', '', '', '', '', '❷同じ列番号から書込む'],
  ['列名', '月', '日', '元請業者', '元下', '受注先', '種別', '場所', '担当', '当初受注金額'],
  ['元列番号❷', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'],
  ['列番号', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'],
  [], [],
  ['ファイル情報❸', '310_工事受注入力', '', '確認_工事インデックス_2611', '', '', '', '', '❶ → ❸工事インデックス 事前確認用スプレッド'],
  ['列名', '---', '日付', '指示書No', '指示書No_連番', '担当者', '元請業者', '担当者', '締日', 'リンク先', '工事個所', '工事名', '工事予定日', '完了日', '請求日', '請求金額', '外注先', 'copyflag'],
  ['元列番号❶', '---', 'A', '', '', 'B', 'C', 'D', '', '', 'K', 'J', 'N', '', '', 'M', '', 'Q', '元データ'],
  ['列番号❸', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'],
  [], [],
  ['ファイル情報', '100_作業指示書_共有', '', '工事インデックス_2611', '', '', '', '', '❸同じ列番号から書込む'],
  ['列名', '---', '日付', '指示書No', '指示書No_連番', '担当者', '元請業者', '担当者', '締日', 'リンク先', '工事個所', '工事名', '工事予定日', '完了日', '請求日', '請求金額', '外注先'],
  ['元列番号❸', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'],
  ['列番号', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'],
  [], [], [],
  ['担当名変換テーブル'],
  ['担当・担当者名', 'メールアドレス'],
  ['加藤', 'kato@nlg.co.jp'],
  ['東谷', 'higashiya@nlg.co.jp'],
  ['東谷', 'hhiro17@gmail.com'],
  ['横井', 'yokoi@nlg.co.jp']
];
const width = Math.max(...settingsRows.map(r => r.length));
const settingsGrid = settingsRows.map(r => Array.from({ length: width }, (_, i) => r[i] ?? ''));

// ===== 模擬ファイル一式 =====

function buildWorld() {
  const formHeader = ['タイムスタンプ', 'メールアドレス', '客先名', '客先担当者名', '受注区分 [官公庁]', '受注区分 [民間]', '受注区分 [ゼネコン]', '受注区分 [同業]', '種別', '工事名', '工事場所（住所）', '工事内容（説明）', '受注金額（税抜）', '工事時期（西暦）', 'その他（補足事項）', '入力の確認', 'copyflag'];
  const input = new Sheet('フォームの回答 2', [
    formHeader,
    [new Date(2026, 10, 5, 9, 30), 'hhiro17@gmail.com', '岐阜市契約課', '岐阜 花子', '元請', '', '', '', '塗装', '金華橋通り区画線工事', '岐阜市金華橋通り1', '4車線区画線100', '270,000', '2026年11月', '歩道橋下を注意', '重複入力ではありません', '']
  ]);
  const oldInput = new Sheet('フォームの回答 1', [formHeader.slice(0, 16)]);

  const stageJ = new Sheet('確認_受注一覧', [
    ['※必ず記入　'], ['工事受注日'],
    ['月', '日', '元請業者', '元下', '受注先', '種別', '場所', '担当', '当初受注金額', 'copyflag']
  ]);

  const idxHeader = ['w', '日付', '指示書No', '', '担当者', '元請業者', '担当者', '締日', 'リンク先', '工事個所', '工事名', '工事予定日', '完了日', '請求日', '請求金額', '外注先', 'copyflag'];
  const old = n => ['', '2026/10/01', '2612510', String(n).padStart(3, '0'), '加藤', '既存', '竹内', '末', '作業指示書', '', `既存工事${n}`, '', '', '', '', ''];
  const stageI = new Sheet('工事Index', [idxHeader, [], [], ...[1, 2, 3].map(n => [...old(n), '2026/10/01'])]);

  // 本番 受注一覧：3行目が見出し、空き行にも K・L 列の数式がある
  const prodJRows = [
    ['※必ず記入　'], ['工事受注日'],
    ['月', '日', '元請業者', '元下', '受注先', '種別', '場所', '担当', '当初受注金額', '変更受注差額金額', '受注金額合計'],
    ['R7.10', 2, 'ｷｸﾃｯｸ愛知', '下請', '同業者', '塗装', 'R302', '加藤', 4696134, 15000000, 19696134],
    [6, 8, 'ｷｸﾃｯｸ愛知', '下請', '同業者', '塗装', '名高東片端', '加藤', 6817000, '', 6817000]
  ];
  const prodJFormulas = {};
  for (let r = 4; r <= 15; r++) { prodJFormulas[`${r},11`] = `=I${r}+J${r}`; prodJFormulas[`${r},12`] = `=K${r}-R${r}`; }
  const prod11 = new Sheet('27期11月', prodJRows, prodJFormulas);
  const prod10 = new Sheet('27期10月', prodJRows.slice(0, 3));

  const prodI = new Sheet('工事Index', [idxHeader.slice(0, 16), [], [], ...[1, 2, 3].map(old)]);

  const files = [
    new Spreadsheet('SET', '設定ファイル_2611', [new Sheet('設定ファイル', settingsGrid)]),
    new Spreadsheet('IN', '工事受注入力_2611', [input, oldInput]),
    new Spreadsheet('SJ', '確認_受注一覧_2611', [stageJ]),
    new Spreadsheet('SI', '確認_工事インデックス_2611', [stageI]),
    new Spreadsheet('PJ', '受注一覧_2611', [new Sheet('雛形', []), prod10, prod11]),
    new Spreadsheet('PI', '工事インデックス_2611', [prodI])
  ];
  return { files, input, oldInput, stageJ, stageI, prod11, prodI };
}

function loadGas(world, activeId) {
  const alerts = [];
  const mails = [];
  const byId = id => {
    const f = world.files.find(x => x.id === id);
    if (!f) throw new Error(`no file ${id}`);
    return f;
  };
  const ctx = {
    console: { log: () => {}, error: () => {} },
    SpreadsheetApp: {
      openById: byId,
      flush: () => {},
      getActiveSpreadsheet: () => byId(activeId),
      getUi: () => ({ alert: m => alerts.push(m), createMenu: () => ({ addItem() { return this; }, addToUi() {} }) })
    },
    DriveApp: {
      searchFiles(q) {
        const name = q.match(/title = '(.+?)'/)[1];
        const list = world.files.filter(f => f.name === name).map(f => ({ getId: () => f.id, getParents: () => ({ hasNext: () => false }) }));
        let i = 0;
        return { hasNext: () => i < list.length, next: () => list[i++] };
      }
    },
    MimeType: { GOOGLE_SHEETS: 'application/vnd.google-apps.spreadsheet' },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    MailApp: { sendEmail: (to, s, b) => mails.push(s + '\n' + b) },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'test@example.com' }) }
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'juchu_transfer.gs'), 'utf8'), ctx);
  return { gas: ctx, alerts, mails };
}

// ===== テスト =====

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

test('設定ファイルを読み分けられる', () => {
  const { gas } = loadGas(buildWorld(), 'SJ');
  const cfg = gas.parseSettings_(settingsGrid);
  assert.strictEqual(cfg.input.flagCol, 17);
  assert.strictEqual(cfg.stageJuchu.flagCol, 10);
  assert.strictEqual(cfg.stageIndex.flagCol, 17);
  assert.strictEqual(cfg.prodJuchu.file, '受注一覧_2611');
  assert.strictEqual(cfg.prodIndex.file, '工事インデックス_2611');
  // ❷ は列番号 A〜I の9項目（copyflag の元列「J」は書かない）
  same(cfg.stageJuchu.fields.map(f => f.dest), [1, 2, 3, 4, 5, 6, 7, 8, 9]);
  // ❸ は B,E,F,G,J,K,L,O の8項目（copyflag の元列「Q」は書かない）
  same(cfg.stageIndex.fields.map(f => f.dest), [2, 5, 6, 7, 10, 11, 12, 15]);
  assert.strictEqual(cfg.prodIndex.fields.length, 16);
  assert.strictEqual(cfg.staff['hhiro17@gmail.com'], '東谷');
});

test('期と月のシート名', () => {
  const { gas } = loadGas(buildWorld(), 'SJ');
  const nov = new Date(2026, 10, 5);
  assert.strictEqual(gas.juchuSheetName_(11, nov), '27期11月');
  assert.strictEqual(gas.juchuSheetName_(7, nov), '27期7月');
  assert.strictEqual(gas.juchuSheetName_(6, new Date(2027, 5, 1)), '27期6月');
  assert.strictEqual(gas.juchuSheetName_(7, new Date(2027, 6, 1)), '28期7月');
  assert.strictEqual(gas.juchuSheetName_(12, new Date(2027, 0, 3)), '27期12月');   // 1月に前年12月分
  assert.throws(() => gas.juchuSheetName_('R7.10', nov));
});

test('書込み行：下に2行空いた最初の行', () => {
  const { gas } = loadGas(buildWorld(), 'SJ');
  const b = [false, true, false, true, true, true, true];
  assert.strictEqual(gas.findBlankRun_(b, 3), 3);
  assert.strictEqual(gas.findBlankRun_([false, false], 3), 2);
});

test('数値と区分名の変換', () => {
  const { gas } = loadGas(buildWorld(), 'SJ');
  assert.strictEqual(gas.toNumber_('270,000'), 270000);
  assert.strictEqual(gas.toNumber_('２７００００'), 270000);
  assert.strictEqual(gas.toNumber_('未定'), '未定');
  assert.strictEqual(gas.categoryName_('受注区分 [ゼネコン]'), 'ゼネコン');
});

test('GAS_1：フォーム送信で ❷❸ に書き、❶ に copyflag', () => {
  const w = buildWorld();
  const { gas } = loadGas(w, 'SJ');
  gas.onFormSubmitHandler({ range: w.input.getRange(2, 1) });

  // ❷：見出し3行の下の4行目
  same(w.stageJ.cells[3].slice(0, 9), [11, 5, '岐阜市契約課', '元請', '官公庁', '塗装', '岐阜市金華橋通り1', '東谷', 270000]);
  assert.strictEqual(w.stageJ.get(4, 10), '', 'copyflag は書かない');

  // ❸：既存3行（4〜6行目）の下の7行目。書かない列は空のまま
  const r = w.stageI.cells[6];
  assert.strictEqual(r[1].getTime(), new Date(2026, 10, 5).getTime());
  same([r[4], r[5], r[6], r[9], r[10], r[11], r[14]],
    ['東谷', '岐阜市契約課', '岐阜 花子', '岐阜市金華橋通り1', '金華橋通り区画線工事', '2026年11月', 270000]);
  assert.ok(r[2] === undefined && r[3] === undefined && r[16] === undefined);

  // ❶ の copyflag（Q列）に日時
  assert.ok(isDate(w.input.get(2, 17)));

  // 同じ行で再度呼ばれても二重に書かない
  gas.onFormSubmitHandler({ range: w.input.getRange(2, 1) });
  assert.strictEqual(w.stageJ.getLastRow(), 4);
});

test('GAS_1：古い回答シートからの送信は無視', () => {
  const w = buildWorld();
  const { gas } = loadGas(w, 'SJ');
  w.oldInput.cells.push([new Date(), 'kato@nlg.co.jp']);
  gas.onFormSubmitHandler({ range: w.oldInput.getRange(2, 1) });
  assert.strictEqual(w.stageJ.getLastRow(), 3);
});

test('GAS_1：未登録アドレスは「無し」、受注区分は左を優先', () => {
  const w = buildWorld();
  const { gas } = loadGas(w, 'SJ');
  w.input.cells[1][1] = 'unknown@example.com';
  w.input.cells[1][4] = '';
  w.input.cells[1][5] = '下請';
  w.input.cells[1][6] = '元請';
  gas.onFormSubmitHandler({ range: w.input.getRange(2, 1) });
  same([w.stageJ.get(4, 4), w.stageJ.get(4, 5), w.stageJ.get(4, 8)], ['下請', '民間', '無し']);
});

test('GAS_2：コピー実行で本番へ書き、copyflag を入れる（既存行は飛ばす）', () => {
  const w = buildWorld();
  const realNow = Date;
  const { gas, alerts } = loadGas(w, 'SJ');
  gas.onFormSubmitHandler({ range: w.input.getRange(2, 1) });
  // テストは 2026年11月 として実行
  gas.Date = class extends realNow { constructor(...a) { super(...(a.length ? a : [2026, 10, 6, 1, 0])); } };

  gas.menuCopyToProduction();
  assert.match(alerts[0], /1 件コピー/);
  // 本番 受注一覧 27期11月：データ2行の下、K・L列に数式がある空き行（6行目）へ
  same(w.prod11.cells[5].slice(0, 9), [11, 5, '岐阜市契約課', '元請', '官公庁', '塗装', '岐阜市金華橋通り1', '東谷', 270000]);
  assert.strictEqual(w.prod11.formulas['6,11'], '=I6+J6', '数式の列は触らない');
  assert.ok(isDate(w.stageJ.get(4, 10)));

  // 2回目は何もしない
  gas.menuCopyToProduction();
  assert.match(alerts[1], /0 件コピー/);
});

test('GAS_2：深夜実行は ❷❸ 両方。❸ は copyflag 済みの既存行を飛ばす', () => {
  const w = buildWorld();
  const { gas, mails } = loadGas(w, 'SJ');
  gas.onFormSubmitHandler({ range: w.input.getRange(2, 1) });
  const realNow = Date;
  gas.Date = class extends realNow { constructor(...a) { super(...(a.length ? a : [2026, 10, 6, 1, 0])); } };
  gas.nightlyCopyToProduction();

  assert.strictEqual(mails.length, 0, mails.join('\n'));
  assert.strictEqual(w.prodI.getLastRow(), 7, '本番 工事Index は既存3行の下に1行だけ追加');
  assert.strictEqual(w.prodI.get(7, 11), '金華橋通り区画線工事');
  assert.strictEqual(w.prodI.get(7, 15), 270000);
  assert.ok(isDate(w.stageI.get(7, 17)));
  assert.strictEqual(w.stageI.get(4, 17), '2026/10/01');
});

test('GAS_2：該当する期・月のシートが無ければ、その行は残してエラー通知', () => {
  const w = buildWorld();
  const { gas, mails } = loadGas(w, 'SJ');
  gas.onFormSubmitHandler({ range: w.input.getRange(2, 1) });
  const realNow = Date;
  gas.Date = class extends realNow { constructor(...a) { super(...(a.length ? a : [2026, 10, 6, 1, 0])); } };
  w.files.find(f => f.id === 'PJ').sheets = [];
  gas.nightlyCopyToProduction();
  assert.strictEqual(mails.length, 1);
  assert.match(mails[0], /27期11月/);
  assert.strictEqual(w.stageJ.get(4, 10), '', 'copyflag は空のまま（次回また試す）');
});

test('GAS_2：確認用以外のファイルでは何もしない', () => {
  const w = buildWorld();
  const { gas, alerts } = loadGas(w, 'PJ');
  gas.menuCopyToProduction();
  assert.match(alerts[0], /確認用ファイルではありません/);
});

let failed = 0;
tests.forEach(t => {
  try { t.fn(); console.log(`ok   ${t.name}`); }
  catch (e) { failed++; console.log(`FAIL ${t.name}\n     ${e.stack.split('\n').slice(0, 3).join('\n     ')}`); }
});
console.log(`\n${tests.length - failed}/${tests.length} passed`);
process.exit(failed ? 1 : 0);
