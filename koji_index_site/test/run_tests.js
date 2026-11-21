/**
 * site_display.gs の動作確認（node koji_index_site/test/run_tests.js）
 * SpreadsheetApp を簡易的な模擬オブジェクトに置き換えて動かす
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');
const same = (a, e, m) => assert.deepStrictEqual(JSON.parse(JSON.stringify(a)), e, m);

class Sheet {
  constructor(name, rows) {
    this.name = name; this.cells = rows.map(r => r.slice()); this.maxRows = Math.max(rows.length, 30); this.maxCols = 24;
    this.formats = {}; this.heights = {}; this.widths = {}; this.hiddenCols = {}; this.frozen = [0, 0]; this.gridHidden = false; this.cf = 0; this.fonts = {}; this.bgs = {};
  }
  getName() { return this.name; }
  get(r, c) { return (this.cells[r - 1] || [])[c - 1] ?? ''; }
  set(r, c, v) { while (this.cells.length < r) this.cells.push([]); this.cells[r - 1][c - 1] = v; }
  getLastRow() { let l = 0; this.cells.forEach((r, i) => { if (r.some(v => v !== '' && v !== undefined)) l = i + 1; }); return l; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }
  insertRowsAfter(a, n) { this.maxRows += n; }
  insertColumnsAfter(a, n) { this.maxCols += n; }
  deleteRows(s, n) {
    this.cells.splice(s - 1, n); this.maxRows -= n;
    const h = {}; Object.keys(this.heights).map(Number).forEach(k => { if (k < s) h[k] = this.heights[k]; else if (k >= s + n) h[k - n] = this.heights[k]; }); this.heights = h;
  }
  clear() { this.cells = []; this.formats = {}; }
  clearConditionalFormatRules() { this.cf = 0; }
  setFrozenRows(n) { this.frozen[0] = n; } setFrozenColumns(n) { this.frozen[1] = n; }
  getFrozenRows() { return this.frozen[0]; } getFrozenColumns() { return this.frozen[1]; }
  hasHiddenGridlines() { return this.gridHidden; } setHiddenGridlines(b) { this.gridHidden = b; }
  getRowHeight(r) { return this.heights[r] || 21; }
  setRowHeights(s, n, h) { if (s + n - 1 > this.maxRows) throw new Error('行の高さが範囲外'); for (let i = 0; i < n; i++) this.heights[s + i] = h; }
  isColumnHiddenByUser(c) { return !!this.hiddenCols[c]; }
  hideColumns(c) { this.hiddenCols[c] = true; } showColumns(c) { delete this.hiddenCols[c]; }
  getRange(r, c, nr = 1, nc = 1) {
    const s = this;
    if (r + nr - 1 > s.maxRows) throw new Error(`範囲外 ${s.name} ${r + nr - 1}行`);
    return {
      getDisplayValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => String(s.get(r + i, c + j)))),
      breakApart() { return this; }, clearDataValidations() { return this; },
      clearContent() { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) s.set(r + i, c + j, ''); return this; },
      copyTo(dest, type) {
        const d = dest.getSheet(), dr = dest.getRow(), dc = dest.getColumn();
        if (type === 'PASTE_COLUMN_WIDTHS') { for (let j = 0; j < nc; j++) d.widths[dc + j] = s.widths[c + j] || 100; return; }
        if (dr + nr - 1 > d.maxRows) throw new Error('貼り付けが範囲外');
        for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) {
          d.set(dr + i, dc + j, s.get(r + i, c + j));
          const f = s.formats[`${r + i},${c + j}`]; if (f) d.formats[`${dr + i},${dc + j}`] = f;
        }
        d.cf += s.cf;
      },
      getSheet: () => s, getRow: () => r, getColumn: () => c, getLastRow: () => r + nr - 1, getLastColumn: () => c + nc - 1,
      setFontColor(v) { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) s.fonts[`${r + i},${c + j}`] = v; return this; },
      setBackground(v) { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) { if (v === null) delete s.bgs[`${r + i},${c + j}`]; else s.bgs[`${r + i},${c + j}`] = v; } return this; }
    };
  }
}

function world() {
  const row = (no, seq, name) => ['', '2026/11/02', no, seq, '加藤', 'メイラ', '林', '随時', '作業指示書', '', name, '', '', '', '', '', '', ''];
  const head = ['w', '日付', '指示書No', '', '担当者', '元請業者', '担当者', '締日', 'リンク先', '工事個所', '工事名'];
  const src = new Sheet('工事Index', [head, [], [], row('2612510', '001', '10月の工事1'), row('2663410', '002', '10月の工事2'),
    row('2661511', '001', '11月の工事1'), row('', '', 'テスト'), row('2600011', '002', '11月の工事2')]);
  src.formats['1,2'] = '青地・白太字'; src.formats['6,9'] = 'リンク'; src.formats['6,11'] = '赤文字';
  src.widths = { 1: 20, 2: 90, 11: 300 }; src.heights = { 1: 40, 6: 30 }; src.frozen = [1, 0]; src.gridHidden = true; src.hiddenCols = { 13: true }; src.cf = 1;
  const sheets = [src];
  const ss = { getSheetByName: n => sheets.find(s => s.name === n) || null, getSheets: () => sheets,
    insertSheet: (n, i) => { const s = new Sheet(n, []); sheets.splice(i, 0, s); return s; } };
  return { src, ss, sheets };
}

function load(w) {
  const alerts = [], triggers = [];
  const ctx = {
    SpreadsheetApp: { getActiveSpreadsheet: () => w.ss, CopyPasteType: { PASTE_COLUMN_WIDTHS: 'PASTE_COLUMN_WIDTHS' },
      getUi: () => ({ alert: m => alerts.push(m), createMenu: () => ({ addItem() { return this; }, addToUi() {} }) }) },
    LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    ScriptApp: { getProjectTriggers: () => triggers.slice(), deleteTrigger: t => triggers.splice(triggers.indexOf(t), 1),
      newTrigger: fn => { const b = { timeBased: () => b, everyHours: () => b, create: () => triggers.push({ getHandlerFunction: () => fn }) }; return b; } }
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'site_display.gs'), 'utf8'), ctx);
  return { gas: ctx, alerts, triggers };
}

const tests = [];
const test = (n, f) => tests.push([n, f]);
const col = (s, c) => s.cells.map(r => r[c - 1] || '');

test('★の行から下を、見出し・書式・色・幅・高さごと写し、×の行は除く', () => {
  const w = world(); w.src.set(6, 1, '★'); w.src.set(7, 1, '×');
  const { gas, alerts } = load(w);
  gas.menuUpdateSite();
  const d = w.ss.getSheetByName('サイト表示');
  assert.strictEqual(w.sheets[1], d, '工事Index の右に作る');
  same(col(d, 11), ['工事名', '', '', '11月の工事1', '11月の工事2']);
  same([d.get(1, 1), d.get(4, 1), w.src.get(6, 1)], ['w', '', '★']);
  same([d.formats['1,2'], d.formats['4,9'], d.formats['4,11']], ['青地・白太字', 'リンク', '赤文字']);
  same([d.widths[1], d.widths[11], d.heights[1], d.heights[4], d.heights[5]], [20, 300, 40, 30, 21]);
  same([d.frozen[0], d.gridHidden, d.isColumnHiddenByUser(13), d.cf > 0, d.getMaxRows()], [1, true, true, true, 15]);
  assert.match(alerts[0], /6行目から 2行を表示（× の 1行は除外）/);
});

test('★が無ければ4行目から。行が減っても前の内容・高さが残らない', () => {
  const w = world(); const { gas } = load(w);
  gas.refreshSite();
  const d = w.ss.getSheetByName('サイト表示');
  same(col(d, 11), ['工事名', '', '', '10月の工事1', '10月の工事2', '11月の工事1', 'テスト', '11月の工事2']);
  w.src.set(8, 1, '★');
  gas.refreshSite();
  same(col(d, 11), ['工事名', '', '', '11月の工事2']);
  same([d.heights[4], d.heights[6], d.getMaxRows()], [21, 21, 14]);
  assert.strictEqual(d.formats['6,11'], undefined, '前の書式が残っている');
});

test('A列を編集したときだけすぐ更新し、自動更新のトリガーは1つだけ', () => {
  const w = world(); const { gas, triggers } = load(w);
  gas.onEdit({ range: w.src.getRange(5, 11) });
  assert.strictEqual(w.ss.getSheetByName('サイト表示'), null, 'A列以外では更新しない');
  w.src.set(6, 1, '★');
  gas.onEdit({ range: w.src.getRange(6, 1) });
  same(col(w.ss.getSheetByName('サイト表示'), 11).slice(3), ['11月の工事1', 'テスト', '11月の工事2']);
  gas.setupSiteTrigger(); gas.setupSiteTrigger();
  same(triggers.map(t => t.getHandlerFunction()), ['refreshSite']);
});

// ===== フラグ（S列） =====
const setFlag = (w, row, v) => w.src.set(row, 19, v);
const rowStyle = (sh, row) => ({ font: [2, 9, 16].map(c => sh.fonts[`${row},${c}`] || null), bg: [2, 9, 16].map(c => sh.bgs[`${row},${c}`] || null), out: [sh.fonts[`${row},1`] || null, sh.fonts[`${row},17`] || null, sh.bgs[`${row},19`] || null] });

test('フラグを選ぶと B〜P列だけ書式が変わる（ピンク・グレー・濃いグレー）', () => {
  const w = world(); const { gas } = load(w);
  [['継続-請無', 4], ['継続-請有', 5], ['工事完了', 6], ['完工', 7], ['中止', 8]].forEach(([f, row]) => { setFlag(w, row, f); gas.onEdit({ range: w.src.getRange(row, 19) }); });
  const P = '#FF4FA3';
  assert.deepStrictEqual(rowStyle(w.src, 4), { font: [P, P, P], bg: [null, null, null], out: [null, null, null] });
  assert.deepStrictEqual(rowStyle(w.src, 6).font, [P, P, P]);
  assert.deepStrictEqual(rowStyle(w.src, 7), { font: ['#000000', '#000000', '#000000'], bg: ['#D9D9D9', '#D9D9D9', '#D9D9D9'], out: [null, null, '#D9D9D9'] });
  assert.strictEqual(w.src.fonts['7,19'], undefined, 'フラグの文字色を変えた');
  assert.deepStrictEqual(rowStyle(w.src, 8).bg, ['#808080', '#808080', '#808080']);
  assert.strictEqual(w.ss.getSheetByName('サイト表示'), null, 'フラグの変更ではサイト表示を更新しない');
});

test('「未施工」・空欄にすると元に戻す（文字は黒、セルの色を外す）。一覧に無い値は触らない', () => {
  const w = world(); const { gas } = load(w);
  setFlag(w, 5, '完工'); gas.onEdit({ range: w.src.getRange(5, 19) });
  setFlag(w, 5, '未施工'); gas.onEdit({ range: w.src.getRange(5, 19) });
  assert.deepStrictEqual(rowStyle(w.src, 5), { font: ['#000000', '#000000', '#000000'], bg: [null, null, null], out: [null, null, null] });
  setFlag(w, 6, '工事完了'); gas.onEdit({ range: w.src.getRange(6, 19) });
  setFlag(w, 6, ''); gas.onEdit({ range: w.src.getRange(6, 19) });
  assert.deepStrictEqual(rowStyle(w.src, 6).font, ['#000000', '#000000', '#000000']);
  setFlag(w, 7, 'その他'); gas.onEdit({ range: w.src.getRange(7, 19) });
  assert.deepStrictEqual(rowStyle(w.src, 7).font, [null, null, null]);
});

test('複数行の貼り付け・メニューの付け直しでもまとめて書式を付ける。見出し行は触らない', () => {
  const w = world(); const { gas, alerts } = load(w);
  setFlag(w, 4, '重複'); setFlag(w, 5, '継続-請有');
  gas.onEdit({ range: w.src.getRange(1, 17, 5, 3) });   // Q〜S列の1〜5行目に貼り付け
  assert.deepStrictEqual(rowStyle(w.src, 4).bg, ['#808080', '#808080', '#808080']);
  assert.deepStrictEqual(rowStyle(w.src, 5).font, ['#FF4FA3', '#FF4FA3', '#FF4FA3']);
  assert.deepStrictEqual(rowStyle(w.src, 1).font, [null, null, null], '見出しを変えた');
  setFlag(w, 8, '中止');
  gas.menuApplyAllFlags();
  assert.deepStrictEqual(rowStyle(w.src, 8).bg, ['#808080', '#808080', '#808080']);
  assert.match(alerts[0], /3行/);   // フラグのある 4・5・8行目だけ
  assert.deepStrictEqual(rowStyle(w.src, 6).font, [null, null, null], 'フラグの無い行を触った');
});

let failed = 0;
tests.forEach(([n, f]) => { try { f(); console.log(`ok   ${n}`); } catch (e) { failed++; console.log(`FAIL ${n}\n     ${e.message}`); } });
console.log(`\n${tests.length - failed}/${tests.length} passed`);
process.exit(failed ? 1 : 0);
