/**
 * links_update.gs の動作確認（node koji_index_links/test/run_tests.js）
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

class Sheet {
  constructor(rows) { this.cells = rows.map(r => r.slice()); }
  get(r, c) { return (this.cells[r - 1] || [])[c - 1] ?? ''; }
  set(r, c, v) { while (this.cells.length < r) this.cells.push([]); this.cells[r - 1][c - 1] = v; }
  getLastRow() { let l = 0; this.cells.forEach((r, i) => { if (r.some(v => v !== '' && v !== undefined && v !== null)) l = i + 1; }); return l; }
  getRange(r, c, nr = 1, nc = 1) {
    const s = this;
    const grid = f => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => f(r + i, c + j)));
    return {
      getDisplayValues: () => grid((y, x) => String(s.get(y, x))),
      clearContent() { grid((y, x) => s.set(y, x, '')); return this; },
      setValues(v) { assert.strictEqual(v.length, nr); assert.strictEqual(v[0].length, nc); v.forEach((row, i) => row.forEach((x, j) => s.set(r + i, c + j, x))); return this; }
    };
  }
}

const file = (name, id, d) => ({ getName: () => name, getUrl: () => `https://drive.google.com/file/d/${id}/view`, getLastUpdated: () => new Date(d) });
function folder(name, files) { return { getName: () => name, getFiles: () => { let i = 0; return { hasNext: () => i < files.length, next: () => files[i++] }; } }; }

function world() {
  const folders = {
    F2025: folder('2025年 工事インデックス', [file('工事インデックス　2025年1月.xlsx', 'a1', '2025-05-09'), file('工事インデックス　2025年6月.xlsx', 'a6', '2025-05-09'),
      file('工事インデックス　2024年12月.xlsx', 'old', '2025-05-09'), file('メモ.txt', 'x', '2025-01-01')]),
    F2025R: folder('2025年 R', [file('工事インデックス　2025年6月.xlsx', 'r6', '2026-08-18')]),
    F2026: folder('2026年 工事インデックス', [file('工事インデックス　2026年11月', 'b11', '2026-11-13'), file('工事インデックス　２０２６年１２月', 'b12', '2026-01-05'),
      file('工事インデックス　2026年10月 のコピー', 'cp', '2026-10-26'), file('工事インデックス　2026年10月', 'b10', '2026-11-11'),
      file('工事インデックス　2026年9月_.xlsx', 'b9old', '2026-01-01'), file('工事インデックス　2026年9月.xlsx', 'b9', '2026-11-06')])
  };
  const settings = new Sheet([['年', '区分', 'フォルダーのURL', 'メモ'],
    [2025, '', 'https://drive.google.com/drive/folders/F2025ABCDEFG'], [2025, 'R', 'https://drive.google.com/drive/folders/F2025RABCDEF'],
    [2026, '', 'https://drive.google.com/drive/folders/F2026ABCDEFG'], [2027, '', 'https://drive.google.com/drive/folders/NOTFOUND12345'],
    [], ['使い方'], ['年ごとのフォルダーを1行ずつ書く…']]);
  const list = new Sheet([['キー', '年', '月', 'ファイル名', '区分', 'URL', 'フォルダー', '更新日'],
    ['=ARRAYFORMULA(…)', 2099, 1, '古い行'], ['', 2099, 2, '古い行'], ['', 2099, 3, '古い行']]);
  const ss = { getSheetByName: n => ({ 設定: settings, URL一覧: list })[n] || null };
  return { folders, settings, list, ss };
}

function load(w) {
  const alerts = [], triggers = [];
  const ctx = {
    SpreadsheetApp: { getActiveSpreadsheet: () => w.ss, getUi: () => ({ alert: m => alerts.push(m), createMenu: () => ({ addItem() { return this; }, addToUi() {} }) }) },
    DriveApp: { getFolderById: id => { const f = w.folders[id.replace(/ABCDEF?G?$/, '')]; if (!f) throw new Error('not found'); return f; } },
    LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    ScriptApp: { getProjectTriggers: () => triggers.slice(), deleteTrigger: t => triggers.splice(triggers.indexOf(t), 1),
      newTrigger: fn => { const b = { timeBased: () => b, atHour: () => b, everyDays: () => b, inTimezone: () => b, create: () => triggers.push({ getHandlerFunction: () => fn }) }; return b; } }
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'links_update.gs'), 'utf8'), ctx);
  return { gas: ctx, alerts, triggers };
}

const tests = [];
const test = (n, f) => tests.push([n, f]);
const rowsOf = l => l.cells.slice(1).filter(r => r.slice(1).some(v => v !== '')).map(r => [r[1], r[2], r[3], r[4], r[5].replace('https://drive.google.com/file/d/', '').replace('/view', '')]);

test('設定のフォルダーから年月のファイルを集め、年→区分→月の順に書く（Rは後ろ）', () => {
  const w = world(); const { gas, alerts } = load(w);
  gas.menuUpdateLinks();
  assert.deepStrictEqual(rowsOf(w.list), [
    [2025, 1, '工事インデックス　2025年1月', '', 'a1'], [2025, 6, '工事インデックス　2025年6月', '', 'a6'],
    [2025, 6, '工事インデックス　2025年6月', 'R', 'r6'],
    [2026, 9, '工事インデックス　2026年9月', '', 'b9'], [2026, 10, '工事インデックス　2026年10月', '', 'b10'],
    [2026, 11, '工事インデックス　2026年11月', '', 'b11'], [2026, 12, '工事インデックス　２０２６年１２月', '', 'b12']]);
  assert.strictEqual(w.list.get(2, 1), '=ARRAYFORMULA(…)', 'A列の数式は残す');
  assert.ok(Object.prototype.toString.call(w.list.get(2, 8)) === '[object Date]', '更新日は日付');
  assert.strictEqual(w.list.get(9, 2), '', '古い行は消す');
  assert.match(alerts[0], /7件/);
  assert.match(alerts[0], /のコピー/);
  assert.match(alerts[0], /開けなかったフォルダー：\n・2027/);
});

test('自動更新のトリガーは1つだけ。トリガー実行でフォルダーが開けなければエラーにする', () => {
  const w = world(); const { gas, triggers } = load(w);
  gas.setupLinksTrigger(); gas.setupLinksTrigger();
  assert.deepStrictEqual(triggers.map(t => t.getHandlerFunction()), ['refreshLinks']);
  assert.throws(() => gas.refreshLinks(), /2027/);
  assert.strictEqual(rowsOf(w.list).length, 7, 'エラーでも見つかった分は書く');
});

test('フォルダーのURLの書き方（URL・?id=・IDだけ）', () => {
  const { gas } = load(world());
  assert.strictEqual(gas.folderIdOf_('https://drive.google.com/drive/folders/18mmdk0n_qoRRO7R2Rh4iaPOp6Va1J6kF?usp=sharing'), '18mmdk0n_qoRRO7R2Rh4iaPOp6Va1J6kF');
  assert.strictEqual(gas.folderIdOf_('https://drive.google.com/open?id=1X77Ixz14HKpQxR4MgmazZ0CmaV1GWjRJ'), '1X77Ixz14HKpQxR4MgmazZ0CmaV1GWjRJ');
  assert.strictEqual(gas.folderIdOf_('1UM4MRBgf2oU6Wq1auE9HKjvicviHiZag'), '1UM4MRBgf2oU6Wq1auE9HKjvicviHiZag');
  assert.strictEqual(gas.folderIdOf_(''), '');
});

let failed = 0;
tests.forEach(([n, f]) => { try { f(); console.log(`ok   ${n}`); } catch (e) { failed++; console.log(`FAIL ${n}\n     ${e.message}`); } });
console.log(`\n${tests.length - failed}/${tests.length} passed`);
process.exit(failed ? 1 : 0);
