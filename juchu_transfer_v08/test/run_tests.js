/**
 * juchu_transfer_v08.gs の動作確認（Node で実行：node juchu_transfer_v08/test/run_tests.js）
 *
 * SpreadsheetApp などを簡易的な模擬オブジェクトに置き換え、
 * 実際の設定ファイル_v08 と同じ内容で GAS_Pri・GAS_Sec を通しで動かす。
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const isDate = v => Object.prototype.toString.call(v) === '[object Date]';
const same = (actual, expected, msg) => assert.deepStrictEqual(JSON.parse(JSON.stringify(actual)), expected, msg);

// ===== 模擬シート =====

class Sheet {
  constructor(name, rows, formulas) {
    this.name = name;
    this.cells = rows.map(r => r.slice());
    this.formulas = formulas || {};
    this.maxRows = Math.max(this.cells.length, 30);
    this.formats = {};
    this.wraps = {};
    this.resized = [];
  }
  autoResizeRows(start, n) { this.resized.push([start, n]); }
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
      setWrap(w) { sheet.wraps[`${r},${c}`] = w; return this; },
      setVerticalAlignment() { return this; },
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

// ===== 実際の設定ファイル_v08 の内容（2026-11-13 時点、File1〜File7） =====

const settingsRows = JSON.parse(fs.readFileSync(path.join(__dirname, 'settings_v08.json'), 'utf8'));
const width = Math.max(...settingsRows.map(r => r.length));
const settingsGrid = settingsRows.map(r => Array.from({ length: width }, (_, i) => r[i] ?? ''));

// ===== 模擬ファイル一式（実ファイルの見出し・先頭データを再現） =====

// File1：A〜S列（S＝Copyflag1）。File4：T＝現場詳細、U＝Copyflag4
const FORM_HEADER = ['タイムスタンプ', 'メールアドレス', '担当', '元請業者', '元請担当者', '受注区分 [元請]', '受注区分 [民間]', '受注区分 [ゼネコン]', '受注区分 [同業者]', '種別', '契約工事名', '工事場所', '受注金額', '工事時期', '締日', '現場詳細', '外注先', '入力確認', 'Copyflag1'];
const INDEX_HEADER = ['', '日付', '指示書No', '', '担当者', '元請業者', '担当者', '締日', 'リンク先', '工事個所', '工事名', '工事予定日', '完了日', '請求日', '請求金額', '外注先', '同一工事', '撮影', '当初受注金額', '現場詳細', 'Copyflag4'];
const C1 = { mail: 1, tanto: 2, kubun1: 5, kubun2: 6, amount: 12, detail: 15, check: 17, flag: 18 };   // File1 の列（0始まり）

function formRow(over) {
  const r = [new Date(2026, 10, 8, 10, 49, 27), 'hhiro17@gmail.com', '', '岐阜市', '岐阜 花子', '', '', '', '元請', 'とび', '美江寺交差点歩道', '市役所前町', '100000', '2026/12/12', '随時', '横断歩道\n歩道路側帯', '無し',
    'A) 重複した工事受注入力ではありません, B) 入力内容を確認しました, C)「過去1年以内・同一客先・同一工事場所」の場合 → チェック, D) 現場撮影が可能な場合 → チェック', ''];
  Object.keys(over || {}).forEach(k => { r[k] = over[k]; });
  return r;
}

function buildWorld() {
  const file1 = new Sheet('フォームの回答 1', [FORM_HEADER, formRow()]);
  const file2 = new Sheet('確認_受注一覧', [[], ['工事受注日'],
    ['月', '日', '元請業者', '元下', '受注先', '種別', '場所', '担当', '当初受注金額', '年', 'Copyflag2'],
    [10, 27, '岐阜市役所契約課', '元請', '官公庁', 'とび', '市役所敷地内', '東谷', 150000, 2026, '2026/10/27 21:33:46']]);
  const old4 = n => ['', '2026/11/02', '2661511', String(n).padStart(3, '0'), '加藤', 'メイラ', '林', '随時', '作業指示書', '', `既存工事${n}`, '', '', '', '', '', '', '', '', '', '2026/11/02'];
  const file4 = new Sheet('工事Index', [INDEX_HEADER, [], [], old4(1), old4(2)]);

  // File3：3行目が見出し。空き行にも K・L 列の数式がある
  const f3Rows = [['※必ず記入　'], ['工事受注日'], ['月', '日', '元請業者', '元下', '受注先', '種別', '場所', '担当', '当初受注金額', '変更受注差額金額', '受注金額合計'],
    [11, 2, 'メイラ', '下請', '同業者', '塗装', '関', '加藤', 300000]];
  const f3Formulas = {};
  for (let r = 4; r <= 20; r++) f3Formulas[`${r},11`] = `=I${r}+J${r}`;
  const file3Nov = new Sheet('27期11月', f3Rows, f3Formulas);
  const file3Dec = new Sheet('27期12月', f3Rows.slice(0, 3));

  const file5 = new Sheet('工事Index', [INDEX_HEADER.slice(0, 18), [], [], old4(1).slice(0, 18)]);

  // File6：7行目が見出し、8行目からデータ。空き行にも G〜K 列に「作成」「経理保管」が入っている
  const f6Rows = [['《27期　注文書・注文請書　作成及び保管チェックリスト》'], [], ['※…'], ['※…'], ['※…'], ['前期から繰越し'],
    ['指示書No.', '', '担当者', '取引先名', '工事名', '金額', '注文書（官公庁は契約書）', '', '', '', '注文請書'],
    ['2500509', '076', '小島', '青葉工業', '平尾２号ため池地区', 3800000, '作成', '〇', '経理保管', '〇', '経理保管'],
    ['\t2616610', '037', '幅', '小林建設工業', '川崎重工南工場正門前', 2170000, '作成', '', '経理保管', '', '経理保管']];
  for (let i = 0; i < 6; i++) f6Rows.push(['', '', '', '', '', '', '作成', '', '経理保管', '', '経理保管']);
  const file6 = new Sheet('27期', f6Rows);

  // File7：1行目がタイトル、2行目が見出し、3行目から見本データ（A〜E列）
  const file7 = new Sheet('現場詳細', [['工事現場詳細_v08'], ['指示書No.', '', '担当者', '元請業者', '工事名', '工事予定日', '現場詳細'],
    ['2500509', '076', '小島', '青葉工業', '平尾２号ため池地区'], ['2612507', '011', '加藤', 'キクテック愛知', '伊勢湾岸自動車道']]);

  const files = [
    new Spreadsheet('SET', '設定ファイル_v08', [new Sheet('設定ファイル', settingsGrid)]),
    new Spreadsheet('F1', '工事受注入力_v08（回答）', [file1]),
    new Spreadsheet('F2', '確認_受注一覧_v08', [file2]),
    new Spreadsheet('F3', '受注一覧_v08', [new Sheet('雛形', []), file3Nov, file3Dec]),
    new Spreadsheet('F4', '確認_工事インデックス_v08', [file4]),
    new Spreadsheet('F5', '工事インデックス_v08', [file5]),
    new Spreadsheet('F6', '注文書・注文請書チェックリスト', [file6, new Sheet('キクテック愛知', [])]),
    new Spreadsheet('F7', '工事現場詳細_v08', [file7])
  ];
  return { files, file1, file2, file3Nov, file4, file5, file6, file7 };
}

function loadGas(world, activeId, props = {}) {
  const alerts = [], mails = [], triggers = [];
  const REAL_SETTINGS_ID = '1NJEMrTL5u0Zs29uAnvP-1I7R_Z3Qn8MTWxIFHonu-Jc';
  const byId = id => {
    const f = world.files.find(x => x.id === (id === REAL_SETTINGS_ID ? 'SET' : id));
    if (!f) throw new Error(`no file ${id}`);
    return f;
  };
  const ctx = {
    console: { log: () => {}, error: () => {} },
    SpreadsheetApp: {
      openById: byId,
      flush: () => {},
      getActiveSpreadsheet: () => (activeId ? byId(activeId) : null),
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
    Session: { getEffectiveUser: () => ({ getEmail: () => 'test@example.com' }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    ScriptApp: {
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: t => triggers.splice(triggers.indexOf(t), 1),
      newTrigger: fn => {
        const b = { forSpreadsheet: () => b, onFormSubmit: () => b, timeBased: () => b, atHour: () => b, everyDays: () => b, inTimezone: () => b,
          create: () => triggers.push({ getHandlerFunction: () => fn }) };
        return b;
      }
    }
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'juchu_transfer_v08.gs'), 'utf8'), ctx);
  return { gas: ctx, alerts, mails, props, triggers };
}

// ===== テスト =====

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

test('設定ファイル_v08 を読み分けられる', () => {
  const { gas } = loadGas(buildWorld(), 'SET');
  const cfg = gas.parseSettings_(settingsGrid);
  const f = cfg.files;
  same([f[1].flagCol, f[2].flagCol, f[4].flagCol], [19, 11, 21]);
  same([f[1].sheet, f[2].sheet, f[3].sheet, f[4].sheet, f[5].sheet, f[6].sheet, f[7].sheet], ['フォームの回答 1', '確認_受注一覧', '', '工事Index', '', '27期', '現場詳細']);
  same(f[2].fields.map(x => x.name), ['月', '日', '元請業者', '元下', '受注先', '種別', '場所', '担当', '当初受注金額', '年']);
  same(f[2].fields.find(x => x.name === '担当').sources, [3, 2]);
  same(f[4].fields.map(x => gas.columnLetter_(x.dest)).join(''), 'BEFGHJKLPQRST');
  same(f[2].fields.find(x => x.name === '当初受注金額').sources, [13]);
  same(f[5].fields.map(x => gas.columnLetter_(x.dest)).join(''), 'BCDEFGHIJKLMNOPQR');
  same(f[6].fields.map(x => [gas.columnLetter_(x.dest), gas.columnLetter_(x.sources[0])]), [['A', 'C'], ['B', 'D'], ['C', 'E'], ['D', 'F'], ['E', 'K'], ['F', 'S']]);
  same(f[7].fields.map(x => [gas.columnLetter_(x.dest), gas.columnLetter_(x.sources[0])]), [['A', 'C'], ['B', 'D'], ['C', 'E'], ['D', 'F'], ['E', 'K'], ['F', 'L'], ['G', 'T']]);
  assert.strictEqual(Object.keys(cfg.staff).length, 14);
  assert.strictEqual(cfg.staff['hhiro17@gmail.com'], '東谷');
});

test('期と月のシート名', () => {
  const { gas } = loadGas(buildWorld(), 'SET');
  assert.strictEqual(gas.periodSheetName_(2026, 11), '27期11月');
  assert.strictEqual(gas.periodSheetName_(2027, 6), '27期6月');
  assert.strictEqual(gas.periodSheetName_('2027', '7'), '28期7月');
  assert.throws(() => gas.periodSheetName_('', 'R7.10'));
});

test('入力確認の選択肢記号', () => {
  const { gas } = loadGas(buildWorld(), 'SET');
  const t = 'A) 重複…, B) 入力内容…, D) 現場撮影は可能です → チェック';
  assert.ok(!gas.hasChoice_(t, 'C'));
  assert.ok(gas.hasChoice_(t, 'D'));
  assert.ok(gas.hasChoice_('C)「過去1年以内」', 'C'));
});

test('GAS_Pri：File2・File4 に書き、Copyflag1 に日時', () => {
  const w = buildWorld();
  const { gas } = loadGas(w, 'SET');
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });

  // File2：既存1行（4行目）の下の5行目
  same(w.file2.cells[4].slice(0, 10), [11, 8, '岐阜市', '元請', '同業者', 'とび', '市役所前町', '東谷', 100000, 2026]);
  assert.strictEqual(w.file2.get(5, 11), '', 'Copyflag2 は書かない');

  // File4：既存2行（4〜5行目）の下の6行目
  const r = w.file4.cells[5];
  assert.ok(isDate(r[1]));
  same([r[4], r[5], r[6], r[7], r[9], r[10], r[11], r[15], r[16], r[17], r[18], r[19]],
    ['東谷', '岐阜市', '岐阜 花子', '随時', '市役所前町', '美江寺交差点歩道', '2026/12/12', '無し', '○', '○', 100000, '横断歩道\n歩道路側帯']);
  assert.ok(r[2] === undefined && r[3] === undefined && r[8] === undefined && r[20] === undefined, '指示書No・リンク先・Copyflag4 は書かない');

  assert.ok(isDate(w.file1.get(2, 19)));
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  assert.strictEqual(w.file2.getLastRow(), 5, '同じ行は二重に書かない');
});

test('GAS_Pri：担当はC列を優先、受注区分が複数なら空、Dだけなら撮影のみ〇', () => {
  const w = buildWorld();
  const { gas } = loadGas(w, 'SET');
  w.file1.cells[1] = formRow({ [C1.tanto]: '山田', [C1.kubun1]: '元請', [C1.kubun2]: '下請', [C1.check]: 'A) 重複…, D) 現場撮影は可能です → チェック' });
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  same([w.file2.get(5, 4), w.file2.get(5, 5), w.file2.get(5, 8)], ['', '', '山田']);
  same([w.file4.get(6, 5), w.file4.get(6, 17), w.file4.get(6, 18)], ['山田', '', '○']);
});

test('GAS_Pri：担当が空で未登録アドレスなら「無し」、金額のカンマも数値に', () => {
  const w = buildWorld();
  const { gas } = loadGas(w, 'SET');
  w.file1.cells[1] = formRow({ [C1.mail]: 'nobody@example.com', [C1.amount]: '1,500,000' });
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  same([w.file2.get(5, 8), w.file2.get(5, 9), w.file4.get(6, 5), w.file4.get(6, 19)], ['無し', 1500000, '無し', 1500000]);
});

test('GAS_Sec：File2→File3（期・月のシート、数式の列は触らない）', () => {
  const w = buildWorld();
  const { gas, alerts } = loadGas(w, 'SET');
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  gas.menuCopyV08();
  assert.match(alerts[0], /File2 → File3：1 件/);
  same(w.file3Nov.cells[4].slice(0, 9), [11, 8, '岐阜市', '元請', '同業者', 'とび', '市役所前町', '東谷', 100000]);
  assert.strictEqual(w.file3Nov.formulas['5,11'], '=I5+J5');
  assert.ok(isDate(w.file2.get(5, 11)));
  assert.strictEqual(w.file2.get(4, 11), '2026/10/27 21:33:46', '既存のコピー済み行はそのまま');
});

test('GAS_Sec：100万円未満は File5 だけ、Copyflag4 を付ける', () => {
  const w = buildWorld();
  const { gas, alerts } = loadGas(w, 'SET');
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  w.file4.set(6, 3, '2616311'); w.file4.set(6, 4, '015');   // 確認時に指示書No を手入力
  gas.menuCopyV08();
  assert.match(alerts[0], /File4 → File5・File6・File7：1 件/);
  assert.strictEqual(w.file5.getLastRow(), 5);
  same([w.file5.get(5, 3), w.file5.get(5, 4), w.file5.get(5, 11), w.file5.get(5, 17), w.file5.get(5, 18)], ['2616311', '015', '美江寺交差点歩道', '○', '○']);
  assert.strictEqual(w.file5.get(5, 19), '', 'File5 は R列まで');
  assert.strictEqual(w.file6.getLastRow(), 15, 'File6 には書かない');
  assert.ok(isDate(w.file4.get(6, 21)));
});

test('GAS_Sec：100万円以上は File6 の A〜F 列にも書く（G列以降は触らない）', () => {
  const w = buildWorld();
  const { gas } = loadGas(w, 'SET');
  w.file1.cells[1] = formRow({ [C1.amount]: '2500000' });
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  w.file4.set(6, 3, '2616311'); w.file4.set(6, 4, '015');
  gas.nightlyCopyV08();
  // File6：データ2行（8〜9行目）の下の10行目
  same(w.file6.cells[9].slice(0, 11), ['2616311', '015', '東谷', '岐阜市', '美江寺交差点歩道', 2500000, '作成', '', '経理保管', '', '経理保管']);
  assert.ok(isDate(w.file4.get(6, 21)));
});

test('GAS_Sec：File6 に同じ指示書No・連番があれば二重に書かない', () => {
  const w = buildWorld();
  const { gas } = loadGas(w, 'SET');
  w.file1.cells[1] = formRow({ [C1.amount]: '2170000' });
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  w.file4.set(6, 3, '2616610'); w.file4.set(6, 4, '037');   // File6 9行目と同じ（A列にタブ入り）
  gas.menuCopyV08();
  assert.strictEqual(w.file6.get(10, 1), '');
  assert.strictEqual(w.file5.getLastRow(), 5, 'File5 には書く');
});

test('GAS_Sec：期・月のシートが無ければ行を残してメール通知', () => {
  const w = buildWorld();
  const { gas, mails } = loadGas(w, 'SET');
  w.file1.cells[1] = formRow({ 0: new Date(2027, 7, 1, 9, 0) });   // 28期8月（未作成）
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  gas.nightlyCopyV08();
  assert.strictEqual(mails.length, 1);
  assert.match(mails[0], /28期8月/);
  assert.strictEqual(w.file2.get(5, 11), '');
});

test('setupTriggers：設定ファイルを覚え、トリガーは1組だけ', () => {
  const w = buildWorld();
  const props = {};
  const first = loadGas(w, 'SET', props);
  first.gas.setupTriggers();
  first.gas.setupTriggers();
  same(first.triggers.map(t => t.getHandlerFunction()), ['onFormSubmitV08', 'nightlyCopyV08']);
  assert.strictEqual(props.SETTINGS_FILE_ID_V08, 'SET');   // コードのIDで開いた設定ファイル
  const { gas } = loadGas(w, null, props);
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  assert.strictEqual(w.file2.getLastRow(), 5);
});

test('スクリプトが設定ファイル以外（File1 など）に入っていても、コードのIDで設定を読める', () => {
  const w = buildWorld();
  const { gas } = loadGas(w, 'F1');
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  assert.strictEqual(w.file2.getLastRow(), 5);
  gas.setupTriggers();
});

test('GAS_Sec：File7 の見出し（2行目）より下に書き、現場詳細を折り返して行の高さを合わせる', () => {
  const w = buildWorld();
  const { gas } = loadGas(w, 'SET');
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  w.file4.set(6, 3, '2616311'); w.file4.set(6, 4, '015');
  gas.menuCopyV08();
  // 見本2行（3〜4行目）の下の5行目
  same(w.file7.cells[4].slice(0, 7), ['2616311', '015', '東谷', '岐阜市', '美江寺交差点歩道', '2026/12/12', '横断歩道\n歩道路側帯']);
  assert.strictEqual(w.file7.wraps['5,7'], true);
  same(w.file7.resized, [[5, 1]]);
  assert.ok(isDate(w.file4.get(6, 21)));
});

test('GAS_Sec：File7 のシートが無ければ Copyflag4 を付けず、File5 にも書かない', () => {
  const w = buildWorld();
  const { gas, alerts } = loadGas(w, 'SET');
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  w.files.find(f => f.id === 'F7').sheets[0].name = '別名';
  gas.menuCopyV08();
  assert.match(alerts[0], /現場詳細/);
  assert.strictEqual(w.file4.get(6, 21), '');
  assert.strictEqual(w.file5.getLastRow(), 4);
});

test('GAS_Sec：途中で止まった行をやり直しても File6・File7 は二重にならない（指示書No が空でも）', () => {
  const w = buildWorld();
  const { gas } = loadGas(w, 'SET');
  w.file1.cells[1] = formRow({ [C1.amount]: '1200000' });
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  // 1回目：File5 への書込みだけを失敗させる（File6・File7 は書けた状態で止まる）
  const orig = w.file5.getRange.bind(w.file5);
  w.file5.getRange = (...a) => { const r = orig(...a); r.setValues = () => { throw new Error('書込み失敗'); }; return r; };
  gas.menuCopyV08();
  assert.strictEqual(w.file4.get(6, 21), '');
  const rows6 = w.file6.getLastRow(), rows7 = w.file7.getLastRow();
  // 2回目：File5 を戻してやり直す
  assert.ok(w.file6.getLastRow() > 9 && w.file7.getLastRow() > 4, '1回目で File6・File7 には書けている');
  delete w.file5.getRange;
  gas.menuCopyV08();
  assert.strictEqual(w.file6.getLastRow(), rows6, 'File6 が二重');
  assert.strictEqual(w.file7.getLastRow(), rows7, 'File7 が二重');
  assert.strictEqual(w.file5.getLastRow(), 5);
  assert.ok(isDate(w.file4.get(6, 21)));
});

test('書込み行は見出し行より下だけを探す（上に空行がある新しいシート）', () => {
  const w = buildWorld();
  const { gas } = loadGas(w, 'SET');
  w.file6.cells = [[], [], [], [], [], [], ['指示書No.', '', '担当者', '取引先名', '工事名', '金額']];
  w.file1.cells[1] = formRow({ [C1.amount]: '2500000' });
  gas.onFormSubmitV08({ range: w.file1.getRange(2, 1) });
  gas.menuCopyV08();
  assert.strictEqual(w.file6.get(8, 3), '東谷', '見出し（7行目）の次の8行目');
  assert.strictEqual(w.file6.get(2, 3), '');
});

let failed = 0;
tests.forEach(t => {
  try { t.fn(); console.log(`ok   ${t.name}`); }
  catch (e) { failed++; console.log(`FAIL ${t.name}\n     ${e.stack.split('\n').slice(0, 3).join('\n     ')}`); }
});
console.log(`\n${tests.length - failed}/${tests.length} passed`);
process.exit(failed ? 1 : 0);
