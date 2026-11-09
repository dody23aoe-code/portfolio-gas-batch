'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const core = require('../src/Core.js');

// GAS の境界をメモリ上で再現し、本番の Main.js をそのまま実行する。
function environment() {
  const sheets = new Map();
  const properties = new Map();
  let triggers = [];
  let fault = null;
  let locked = false;
  const hit = event => {
    if (fault && fault === event) { fault = null; throw new Error('simulated interruption: ' + event); }
  };
  class Sheet {
    constructor(name, source) {
      this.name = name;
      this.cells = source ? structuredClone(source.cells) : [];
      this.maxRows = 1000;
    }
    getRange(row, col, height = 1, width = 1) {
      let a1 = '';
      if (typeof row === 'string') {
        a1 = row;
        const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(row);
        const number = s => [...s].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
        row = Number(m[2]); col = number(m[1]);
        height = m[4] ? Number(m[4]) - row + 1 : 1;
        width = m[3] ? number(m[3]) - col + 1 : 1;
      }
      const sheet = this;
      const range = {
        setValues(values) {
          assert.equal(values.length, height);
          values.forEach((line, r) => {
            assert.equal(line.length, width);
            line.forEach((v, c) => {
              sheet.cells[row + r - 1] ||= [];
              sheet.cells[row + r - 1][col + c - 1] = v;
            });
          });
          if (sheet.name.startsWith('PO_')) hit(row === 9 ? 'detail-written' : 'header-written');
          if (sheet.name === '処理記録' && row >= 5) hit('audit-written');
          if (sheet.name === '処理記録' && a1 === 'A2:H2') hit('progress-written');
          return range;
        },
        setValue(value) { return range.setValues(Array.from({ length: height }, () => Array(width).fill(value))); }
      };
      ['merge', 'setFontSize', 'setFontWeight', 'setBackground', 'setFontColor', 'setNumberFormat'].forEach(method => {
        range[method] = () => range;
      });
      return range;
    }
    getDataRange() {
      return { getValues: () => this.cells.map(row => Array.from({ length: 6 }, (_, c) => row[c] ?? '')) };
    }
    getMaxRows() { return this.maxRows; }
    insertRowsAfter(start, n) { this.maxRows += n; return this; }
    clearContents() { this.cells = []; return this; }
  }
  ['hideSheet', 'showSheet', 'setFrozenRows', 'autoResizeColumns', 'setColumnWidth', 'setColumnWidths'].forEach(method => {
    Sheet.prototype[method] = function () { return this; };
  });
  const ss = {
    getId: () => 'test-spreadsheet',
    getSheetByName: name => sheets.get(name) || null,
    insertSheet(name, options) {
      assert.ok(!sheets.has(name), 'duplicate sheet creation');
      const sheet = new Sheet(name, options?.template);
      sheets.set(name, sheet);
      if (name.startsWith('PO_')) hit('sheet-created');
      return sheet;
    }
  };
  const ctx = vm.createContext({
    console,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss, flush() {} },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: key => properties.get(key) || null,
      setProperty(key, value) {
        if (JSON.parse(value).next > 0) hit('checkpoint-before');
        properties.set(key, value);
        if (JSON.parse(value).next > 0) hit('checkpoint-after');
      }
    }) },
    LockService: { getScriptLock: () => ({
      tryLock: () => { if (locked) return false; locked = true; return true; },
      releaseLock: () => { locked = false; }
    }) },
    Utilities: { getUuid: () => '00000000-0000-4000-8000-000000000000' },
    ScriptApp: {
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger(t) { hit('trigger-delete'); triggers = triggers.filter(x => x !== t); },
      newTrigger(handler) {
        const builder = { timeBased: () => builder, everyMinutes: () => builder,
          create: () => { const t = { getHandlerFunction: () => handler }; triggers.push(t); return t; } };
        return builder;
      }
    }
  });
  function reloadScripts() {
    for (const name of ['Core.js', 'Main.js']) {
      vm.runInContext(fs.readFileSync(path.join(__dirname, '../src', name), 'utf8'), ctx, { filename: name });
    }
  }
  reloadScripts();
  ctx.setupDemo();
  return {
    ctx, sheets, ss,
    fault: event => { fault = event; },
    lock: value => { locked = value; },
    state: () => JSON.parse(properties.get('PURCHASE_BATCH_STATE') || 'null'),
    triggers: () => triggers.length,
    finish() {
      for (let i = 0; i < 50 && (this.state()?.status !== 'COMPLETED' || triggers.length); i++) {
        reloadScripts();
        ctx.resumeBatch();
      }
      assert.equal(this.state().status, 'COMPLETED');
      assert.equal(triggers.length, 0);
    },
    outputs: () => JSON.parse(JSON.stringify([...sheets].filter(([name]) => name.startsWith('PO_')).map(([name, sheet]) => [name, sheet.cells])))
  };
}

test('CSV は600行の決定的なサンプルと一致する', () => {
  const csv = fs.readFileSync(path.join(__dirname, '../data/orders.csv'), 'utf8');
  const expected = '\uFEFF' + core.sampleRows().map(r => r.map(v => '"' + v + '"').join(',')).join('\r\n') + '\r\n';
  assert.equal(csv, expected);
  assert.equal(core.tasksFromRows(core.sampleRows(), 100).length, 30);
});

test('一括実行と5帳票ずつの再開は全セルが一致し、600注文が各1回だけ出力される', () => {
  const one = environment();
  one.ctx.BATCH_CONFIG = { ...one.ctx.BATCH_CONFIG, maxTasks: 10000 };
  one.ctx.startBatch();
  assert.equal(one.state().status, 'COMPLETED');
  const resumed = environment();
  resumed.ctx.startBatch();
  assert.equal(resumed.state().next, 5);
  assert.equal(resumed.triggers(), 1);
  resumed.finish();
  assert.deepEqual(resumed.outputs(), one.outputs());
  const ids = resumed.outputs().flatMap(([, rows]) => rows.slice(8, 108).map(r => r[0]).filter(Boolean));
  assert.equal(ids.length, 600);
  assert.equal(new Set(ids).size, 600);
  assert.equal(resumed.state().processedRows, 600);
  assert.equal(resumed.sheets.get('処理記録').cells[1][0], 'COMPLETED');
  assert.equal(resumed.sheets.get('処理記録').cells.slice(4).length, 30);
  resumed.ctx.startBatch();
  resumed.ctx.resumeBatch();
  assert.deepEqual(resumed.outputs(), one.outputs());
});

for (const event of ['sheet-created', 'header-written', 'detail-written', 'audit-written',
  'checkpoint-before', 'checkpoint-after', 'progress-written', 'trigger-delete']) {
  test('障害からの回復で重複しない: ' + event, () => {
    const reference = environment();
    reference.ctx.startBatch(); reference.finish();
    const interrupted = environment();
    interrupted.fault(event);
    if (event === 'trigger-delete') {
      interrupted.ctx.startBatch();
      assert.throws(() => interrupted.finish(), /simulated interruption/);
    } else {
      assert.throws(() => interrupted.ctx.startBatch(), /simulated interruption/);
    }
    assert.equal(interrupted.triggers(), 1);
    interrupted.finish();
    assert.deepEqual(interrupted.outputs(), reference.outputs());
    assert.equal(interrupted.sheets.get('処理記録').cells.slice(4).length, 30);
  });
}

test('途中で入力を変更しても開始時のスナップショットを使う', () => {
  const reference = environment(); reference.ctx.startBatch(); reference.finish();
  const changed = environment(); changed.ctx.startBatch();
  changed.sheets.get('注文').cells[1][4] = 999;
  changed.sheets.get('発注書雛形').cells[0][0] = '変更';
  changed.finish();
  assert.deepEqual(changed.outputs(), reference.outputs());
});

test('ロック競合時は処理せず、既存の再開トリガーを維持する', () => {
  const e = environment(); e.ctx.startBatch();
  const before = e.state(); e.lock(true); e.ctx.resumeBatch();
  assert.deepEqual(e.state(), before); assert.equal(e.triggers(), 1);
  e.lock(false); e.finish();
});

test('時間予算で処理を中断し、合計6分超の仮想作業も最後まで再開する', () => {
  const tasks = core.tasksFromRows(core.sampleRows(), 100);
  let clock = 0, next = 0, runs = 0;
  const outputs = new Map();
  while (next < tasks.length) {
    const deadline = clock + 240000;
    next = core.run(tasks, next, {
      now: () => clock,
      render: (t, i) => { outputs.set(i, t); clock += 70000; },
      record() {}, checkpoint() {}
    }, { maxTasks: 1000, deadline });
    runs++;
  }
  assert.ok(clock > 360000); assert.ok(runs > 1);
  assert.deepEqual([...outputs.values()], tasks);
  assert.equal(core.run(tasks, 0, { now: () => 100, render: () => assert.fail('予算超過') },
    { maxTasks: 5, deadline: 100 }), 0);
});

test('同一仕入れ先の100行超はページ分割し、順序も固定される', () => {
  const rows = core.sampleRows();
  rows.slice(1).forEach(r => { r[1] = 'SUP-001'; r[2] = '架空商事001'; });
  const tasks = core.tasksFromRows(rows, 100);
  assert.equal(tasks.length, 6);
  assert.deepEqual(tasks.map(t => t.page), [1, 2, 3, 4, 5, 6]);
  assert.deepEqual(core.tasksFromRows([rows[0], ...rows.slice(1).reverse()], 100), tasks);
});

test('空の注文も完了し、不正入力は出力前に拒否する', () => {
  const empty = environment();
  empty.sheets.get('注文').cells = [core.HEADERS.slice()];
  empty.ctx.startBatch(); assert.equal(empty.state().status, 'COMPLETED');
  assert.equal(empty.outputs().length, 0); assert.equal(empty.triggers(), 0);
  for (const mutate of [r => { r[2][0] = r[1][0]; }, r => { r[1][4] = -1; },
    r => { r[1][5] = '100'; }, r => { r[1][3] = '=HYPERLINK("example")'; },
    r => { r[31][2] = '不一致'; }, r => { r[0][0] = 'unknown'; }]) {
    const rows = core.sampleRows(); mutate(rows);
    assert.throws(() => core.tasksFromRows(rows, 100));
  }
});

test('初期化エラー後も入力修正で再開できる', () => {
  const e = environment(); e.sheets.get('注文').cells[1][4] = -1;
  assert.throws(() => e.ctx.startBatch(), /不正な注文/);
  assert.equal(e.state().status, 'PREPARING'); assert.equal(e.triggers(), 1);
  e.sheets.get('注文').cells[1][4] = core.sampleRows()[1][4];
  e.finish(); assert.equal(e.state().processedRows, 600);
});
