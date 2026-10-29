/* 用真实 build() 产出 data.json（离线版）——喂 build_cockpit2.py 做端到端验证
   用法：node dump_json.js <content.md> <输出路径> */
var fs = require('fs'), path = require('path');
var ARG = process.argv.slice(2);
var SRC = path.resolve(ARG[0] || '../_kdocs/content_new.md');
var DST = path.resolve(ARG[1] || '../_kdocs/data_air.json');
var SHEETS = {};
fs.readFileSync(SRC, 'utf8').split(/^## /m).slice(1).forEach(function (p) {
  var i = p.indexOf('\n'), name = p.slice(0, i).trim(), rows = [];
  p.slice(i + 1).split('\n').forEach(function (l) {
    if (l.charAt(0) !== '|') return;
    rows.push(l.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(function (c) { return c.trim(); }));
  });
  SHEETS[name] = rows;
});
function ci(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
function colIndex(ref) { var m = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/); if (!m) return null;
  return { r1: +m[2], c1: ci(m[1]), r2: +m[4], c2: ci(m[3]) }; }
function mkSheet(name) { var g = SHEETS[name] || [];
  return { Name: name, Range: function (ref) { var r = colIndex(ref), o = [];
    for (var i = r.r1; i <= r.r2; i++) { var s = g[i - 1] || [], row = [];
      for (var j = r.c1; j <= r.c2; j++) row.push(s[j - 1] === undefined ? '' : s[j - 1]);
      o.push(row); }
    return { Value2: o }; } }; }
var names = Object.keys(SHEETS);
global.Application = { Sheets: { Count: names.length, Item: function (k) {
  if (typeof k === 'number') return mkSheet(names[k - 1]);
  if (SHEETS[k]) return mkSheet(k); throw new Error('no sheet ' + k); } } };
var LOGS = [];
global.console = { log: function (x) { LOGS.push(String(x)); }, error: function () {} };
eval(fs.readFileSync(path.join(__dirname, 'read_daily.js'), 'utf8').replace(/main\(\);\s*$/, ''));
var D = build();
fs.writeFileSync(DST, JSON.stringify(D), 'utf8');
process.stdout.write('已写出 ' + DST + '（' + JSON.stringify(D).length + ' 字节）\n');
LOGS.forEach(function (x) { process.stdout.write(x + '\n'); });
