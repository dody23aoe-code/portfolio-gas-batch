/* Google Apps Script と Node.js の両方で実行する、副作用のない処理。 */
var BatchCore = (function () {
  'use strict';
  var HEADERS = ['order_id', 'supplier_id', 'supplier_name', 'item_name', 'quantity', 'unit_price'];
  function sampleRows() {
    var rows = [HEADERS.slice()];
    for (var i = 1; i <= 600; i++) {
      var supplier = String((i - 1) % 30 + 1).padStart(3, '0');
      rows.push(['ORD-' + String(i).padStart(4, '0'), 'SUP-' + supplier,
        '架空商事' + supplier, 'サンプル商品' + String((i - 1) % 40 + 1).padStart(2, '0'),
        i % 9 + 1, (i % 20 + 1) * 100]);
    }
    return rows;
  }
  function tasksFromRows(rows, pageSize) {
    if (!Number.isInteger(pageSize) || pageSize < 1) throw new Error('ページサイズが不正です');
    if (!rows.length || JSON.stringify(rows[0]) !== JSON.stringify(HEADERS)) {
      throw new Error('注文シートの列名・列順を確認してください');
    }
    var ids = new Set();
    var groups = new Map();
    rows.slice(1).forEach(function (r, index) {
      if (r.every(function (v) { return v === ''; })) return;
      if (r.length !== 6 || !/^[A-Z0-9-]+$/.test(r[0]) || !/^[A-Z0-9-]+$/.test(r[1]) ||
          typeof r[2] !== 'string' || !r[2].trim() || typeof r[3] !== 'string' || !r[3].trim() ||
          /^[=+@-]/.test(r[2]) || /^[=+@-]/.test(r[3]) ||
          !Number.isSafeInteger(r[4]) || r[4] <= 0 || !Number.isSafeInteger(r[5]) || r[5] < 0 ||
          !Number.isSafeInteger(r[4] * r[5])) throw new Error('不正な注文: 行 ' + (index + 2));
      if (ids.has(r[0])) throw new Error('注文ID重複: ' + r[0]);
      ids.add(r[0]);
      if (!groups.has(r[1])) groups.set(r[1], { id: r[1], name: r[2], rows: [] });
      var group = groups.get(r[1]);
      if (group.name !== r[2]) throw new Error('仕入れ先名が一致しません: ' + r[1]);
      group.rows.push(r.slice());
    });
    var tasks = [];
    Array.from(groups.keys()).sort().forEach(function (id) {
      var group = groups.get(id);
      group.rows.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; });
      for (var offset = 0; offset < group.rows.length; offset += pageSize) {
        var lines = group.rows.slice(offset, offset + pageSize);
        var total = lines.reduce(function (sum, r) { return sum + r[4] * r[5]; }, 0);
        if (!Number.isSafeInteger(total)) throw new Error('合計金額が整数の安全範囲を超えています');
        tasks.push({ supplierId: id, supplierName: group.name, page: offset / pageSize + 1,
          lines: lines, total: total });
      }
    });
    return tasks;
  }
  function outputName(jobId, index) {
    return 'PO_' + jobId + '_' + String(index + 1).padStart(5, '0');
  }
  // render と record が完了した後だけ checkpoint を進める。
  // render/record は、同じキーへの再実行で同じ結果になる必要がある。
  function run(tasks, next, io, limits) {
    var count = 0;
    while (next < tasks.length && count < limits.maxTasks && io.now() < limits.deadline) {
      io.render(tasks[next], next);
      io.record(tasks[next], next);
      io.checkpoint(next + 1);
      next++;
      count++;
    }
    return next;
  }
  return { HEADERS: HEADERS, sampleRows: sampleRows, tasksFromRows: tasksFromRows,
    outputName: outputName, run: run };
}());
if (typeof module !== 'undefined') module.exports = BatchCore;
