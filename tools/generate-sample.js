'use strict';
const fs = require('node:fs');
const path = require('node:path');
const core = require('../src/Core.js');
const csv = core.sampleRows().map(row => row.map(value => '"' + String(value).replace(/"/g, '""') + '"').join(',')).join('\r\n') + '\r\n';
fs.mkdirSync(path.join(__dirname, '../data'), { recursive: true });
fs.writeFileSync(path.join(__dirname, '../data/orders.csv'), '\uFEFF' + csv, 'utf8');
console.log('data/orders.csv: 600 rows generated');
