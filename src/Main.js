/* 1つのスクリプトにつき1つのスプレッドシート・1ジョブを扱うサンプル。 */
var BATCH_CONFIG = Object.freeze({
  stateKey: 'PURCHASE_BATCH_STATE', orders: '注文', template: '発注書雛形', log: '処理記録',
  snapshot: '_注文スナップショット', templateSnapshot: '_雛形スナップショット',
  pageSize: 100, maxTasks: 5, budgetMs: 240000, handler: 'resumeBatch'
});

function onOpen() {
  SpreadsheetApp.getUi().createMenu('発注書サンプル')
    .addItem('サンプルを準備', 'setupDemo').addItem('開始 / 再開', 'startBatch')
    .addItem('状態を表示', 'showBatchStatus').addToUi();
}

function withBatchLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try { return fn(); } finally { lock.releaseLock(); }
}
function loadState_() {
  var raw = PropertiesService.getScriptProperties().getProperty(BATCH_CONFIG.stateKey);
  return raw ? JSON.parse(raw) : null;
}
function saveState_(state) {
  PropertiesService.getScriptProperties().setProperty(BATCH_CONFIG.stateKey, JSON.stringify(state));
}

function setupDemo() {
  withBatchLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('スプレッドシートに紐づくスクリプトで実行してください');
    if (loadState_() || [BATCH_CONFIG.orders, BATCH_CONFIG.template, BATCH_CONFIG.log,
      BATCH_CONFIG.snapshot, BATCH_CONFIG.templateSnapshot].some(function (name) { return ss.getSheetByName(name); })) {
      throw new Error('既存データを保護するため準備を中止しました。新しい空のスプレッドシートを使用してください');
    }
    var orders = ss.insertSheet(BATCH_CONFIG.orders);
    orders.getRange(1, 1, 601, 6).setValues(BatchCore.sampleRows());
    orders.setFrozenRows(1);
    orders.autoResizeColumns(1, 6);
    var template = ss.insertSheet(BATCH_CONFIG.template);
    template.getRange('A1:F1').merge().setValue('発 注 書').setFontSize(20).setFontWeight('bold');
    template.getRange('A3:A6').setValues([['発注書番号'], ['仕入れ先ID'], ['仕入れ先名'], ['ページ']]);
    template.getRange('E7').setValue('合計（円）');
    template.getRange('A8:F8').setValues([['注文ID', '商品名', '数量', '単価（円）', '金額（円）', '備考']])
      .setBackground('#16324f').setFontColor('#ffffff').setFontWeight('bold');
    template.getRange(9, 3, 100, 3).setNumberFormat('#,##0');
    template.getRange('F7').setNumberFormat('#,##0');
    template.setColumnWidth(1, 150).setColumnWidth(2, 220);
    template.setColumnWidths(3, 4, 110);
    template.setFrozenRows(8);
    var log = ss.insertSheet(BATCH_CONFIG.log);
    log.getRange('A1:H1').setValues([['状態', '完了帳票数', '総帳票数', '処理済行数', '総行数', '更新日時', 'ジョブID', 'メッセージ']]);
    log.getRange('A4:H4').setValues([['帳票番号', '仕入れ先ID', 'ページ', '行数', '合計（円）', '状態', '更新日時', '出力シート']]);
    log.setFrozenRows(4);
    SpreadsheetApp.flush();
  });
}

function ensureTrigger_() {
  var matches = ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === BATCH_CONFIG.handler; });
  if (!matches.length) ScriptApp.newTrigger(BATCH_CONFIG.handler).timeBased().everyMinutes(1).create();
  matches.slice(1).forEach(function (t) { ScriptApp.deleteTrigger(t); });
}
function removeTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === BATCH_CONFIG.handler) ScriptApp.deleteTrigger(t);
  });
}

function startBatch() {
  withBatchLock_(function () {
    var state = loadState_();
    if (!state) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss || !ss.getSheetByName(BATCH_CONFIG.orders) || !ss.getSheetByName(BATCH_CONFIG.template) ||
          !ss.getSheetByName(BATCH_CONFIG.log)) throw new Error('先に setupDemo を実行してください');
      // 事前に巡回トリガーを設け、初期化中の強制終了にも備える。
      ensureTrigger_();
      state = { spreadsheetId: ss.getId(), jobId: Utilities.getUuid().replace(/-/g, ''),
        status: 'PREPARING', next: 0, total: 0, totalRows: 0, processedRows: 0 };
      saveState_(state);
    }
  });
  resumeBatch();
}

function prepareBatch_(ss, state) {
  var source = ss.getSheetByName(BATCH_CONFIG.orders).getDataRange().getValues();
  var tasks = BatchCore.tasksFromRows(source, BATCH_CONFIG.pageSize);
  var snapshot = ss.getSheetByName(BATCH_CONFIG.snapshot) || ss.insertSheet(BATCH_CONFIG.snapshot);
  if (snapshot.getMaxRows() < source.length) snapshot.insertRowsAfter(snapshot.getMaxRows(), source.length - snapshot.getMaxRows());
  snapshot.clearContents();
  snapshot.getRange(1, 1, source.length, 6).setValues(source);
  snapshot.hideSheet();
  if (!ss.getSheetByName(BATCH_CONFIG.templateSnapshot)) {
    ss.insertSheet(BATCH_CONFIG.templateSnapshot, { template: ss.getSheetByName(BATCH_CONFIG.template) }).hideSheet();
  }
  SpreadsheetApp.flush();
  state.total = tasks.length;
  state.totalRows = tasks.reduce(function (n, t) { return n + t.lines.length; }, 0);
  state.status = 'RUNNING';
  saveState_(state);
  return tasks;
}

function renderPurchaseOrder_(ss, state, task, index) {
  var name = BatchCore.outputName(state.jobId, index);
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name, { template: ss.getSheetByName(BATCH_CONFIG.templateSnapshot) });
  sheet.showSheet();
  sheet.getRange('B3:B6').setValues([[name], [task.supplierId], [task.supplierName], [task.page]]);
  sheet.getRange('F7').setValue(task.total);
  var detail = task.lines.map(function (r) { return [r[0], r[3], r[4], r[5], r[4] * r[5], '']; });
  while (detail.length < BATCH_CONFIG.pageSize) detail.push(['', '', '', '', '', '']);
  sheet.getRange(9, 1, detail.length, 6).setValues(detail);
  SpreadsheetApp.flush();
}
function recordTask_(ss, state, task, index) {
  var log = ss.getSheetByName(BATCH_CONFIG.log);
  var row = index + 5;
  if (log.getMaxRows() < row) log.insertRowsAfter(log.getMaxRows(), row - log.getMaxRows());
  log.getRange(row, 1, 1, 8).setValues([[index + 1, task.supplierId, task.page, task.lines.length,
    task.total, 'DONE', new Date(), BatchCore.outputName(state.jobId, index)]]);
  SpreadsheetApp.flush();
}
function recordProgress_(ss, state, message) {
  ss.getSheetByName(BATCH_CONFIG.log).getRange('A2:H2').setValues([[state.status, state.next, state.total,
    state.processedRows, state.totalRows, new Date(), state.jobId, String(message || '').slice(0, 500)]]);
  SpreadsheetApp.flush();
}

function resumeBatch() {
  var started = Date.now();
  withBatchLock_(function () {
    var state = loadState_();
    if (!state) { removeTriggers_(); return; }
    var ss = SpreadsheetApp.openById(state.spreadsheetId);
    if (state.status === 'COMPLETED') {
      recordProgress_(ss, state, '全件完了');
      removeTriggers_();
      return;
    }
    // 次回を先に予約。実行時間超過で catch が動かなくても再開できる。
    ensureTrigger_();
    try {
      var tasks = state.status === 'PREPARING' ? prepareBatch_(ss, state) :
        BatchCore.tasksFromRows(ss.getSheetByName(BATCH_CONFIG.snapshot).getDataRange().getValues(), BATCH_CONFIG.pageSize);
      var cumulative = [0];
      tasks.forEach(function (t, i) { cumulative.push(cumulative[i] + t.lines.length); });
      BatchCore.run(tasks, state.next, {
        now: Date.now,
        render: function (task, index) { renderPurchaseOrder_(ss, state, task, index); },
        record: function (task, index) { recordTask_(ss, state, task, index); },
        checkpoint: function (next) {
          state.next = next;
          state.processedRows = cumulative[next];
          saveState_(state);
          recordProgress_(ss, state, '処理中');
        }
      }, { maxTasks: BATCH_CONFIG.maxTasks, deadline: started + BATCH_CONFIG.budgetMs });
      if (state.next === tasks.length) {
        state.status = 'COMPLETED';
        saveState_(state);
        recordProgress_(ss, state, '全件完了');
        removeTriggers_();
      } else {
        recordProgress_(ss, state, '次回の時間トリガーで再開');
      }
    } catch (error) {
      // 保存の失敗後にメモリ上の next を誤って永続化しない。
      var persisted = loadState_();
      recordProgress_(ss, persisted, '再試行待ち: ' + error.message);
      throw error;
    }
  });
}

function showBatchStatus() {
  SpreadsheetApp.getUi().alert(JSON.stringify(loadState_(), null, 2) || '未開始');
}
