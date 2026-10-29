/* 本地离线验证：用 content.md 还原成二维表，桩掉 Application/console，
   跑 AirScript 的 build()，与 extract.py 产出的 _kdocs/data.json 逐字段对比。
   用法：node verify_mock.js [content文件] [--diag]
        不带参数默认 _kdocs/content.md；--diag 只打印逐日诊断不做对比 */
var fs = require('fs');
var path = require('path');
var OUT = function (m) { process.stdout.write(m + '\n'); };
var ARG = process.argv.slice(2);
var DIAG = ARG.indexOf('--diag') >= 0;
ARG = ARG.filter(function (a) { return a !== '--diag'; });

/* ---- 1. content.md -> {sheetName: 二维数组(0基)} ---- */
var BASE = path.join(__dirname, '..', '_kdocs');
var raw = fs.readFileSync(ARG[0] ? path.resolve(ARG[0]) : path.join(BASE, 'content.md'), 'utf8');
var SHEETS = {};
raw.split(/^## /m).slice(1).forEach(function (p) {
  var i = p.indexOf('\n');
  var name = p.slice(0, i).trim();
  var rows = [];
  p.slice(i + 1).split('\n').forEach(function (line) {
    if (line.charAt(0) !== '|') return;
    var cells = line.trim().replace(/^\|/, '').replace(/\|$/, '').split('|')
      .map(function (c) { return c.trim(); });
    rows.push(cells);
  });
  SHEETS[name] = rows;
});

function colIndex(ref) {
  var m = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!m) return null;
  function ci(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
  return { r1: +m[2], c1: ci(m[1]), r2: +m[4], c2: ci(m[3]) };
}

function mkSheet(name) {
  var g = SHEETS[name] || [];
  return {
    Name: name,
    Range: function (ref) {
      var r = colIndex(ref);
      if (!r) throw new Error('bad ref ' + ref);
      var out = [];
      for (var i = r.r1; i <= r.r2; i++) {
        var src = g[i - 1] || [];
        var row = [];
        for (var j = r.c1; j <= r.c2; j++) row.push(src[j - 1] === undefined ? '' : src[j - 1]);
        out.push(row);
      }
      return { Value2: out };
    }
  };
}

var names = Object.keys(SHEETS);
global.Application = {
  Sheets: {
    Count: names.length,
    Item: function (k) {
      if (typeof k === 'number') return mkSheet(names[k - 1]);
      if (SHEETS[k]) return mkSheet(k);
      throw new Error('no sheet ' + k);
    }
  }
};
global.console = { log: function () {} };

/* ---- 2. 载入脚本（去掉末尾 main() 调用） ---- */
var src = fs.readFileSync(path.join(__dirname, 'read_daily.js'), 'utf8').replace(/main\(\);\s*$/, '');
eval(src);

/* ---- 3. 跑 build() 并与真值对比 ---- */
/* 真值优先取 data_true_ref.json（extract.py 独立实现产出的冻结基线），
   没有才退回 data.json（当前在跑的那份，可能是本脚本自己产的，属循环对照）。 */
var REF = path.join(BASE, 'data_true_ref.json');
if (!fs.existsSync(REF)) REF = path.join(BASE, 'data.json');
OUT('真值来源 ' + path.basename(REF));
var T = JSON.parse(fs.readFileSync(REF, 'utf8'));
var D;
try { D = build(); } catch (e) { OUT('构建失败：' + e.message); process.exit(1); }

/* ---- 诊断模式：逐日列出捕获到的站点与合计 ---- */
if (DIAG) {
  OUT('会计月 ' + D.meta.accounting_month + '   天数 ' + D.meta.days_with_data.length);
  D.meta.days_with_data.forEach(function (d) {
    var v = D.sale[d], nm = Object.keys(v.stations);
    OUT('  日' + d + '  合计=' + v.t + '  站数=' + nm.length + '  小额=' + v.xiao.t);
    OUT('       站点: ' + nm.join(' '));
  });
  process.exit(0);
}

var pass = 0, fail = 0;
function eq(label, a, b, tol) {
  var ok;
  if (typeof a === 'number' && typeof b === 'number') ok = Math.abs(a - b) <= (tol === undefined ? 0.02 : tol);
  else ok = JSON.stringify(a) === JSON.stringify(b);
  if (ok) { pass++; } else { fail++; OUT('  ✗ ' + label + '  期望=' + JSON.stringify(b) + '  实际=' + JSON.stringify(a)); }
}

OUT('真值会计月 ' + T.meta.accounting_month + ' / 脚本 ' + D.meta.accounting_month);
eq('天数', D.meta.days_with_data.join(','), T.meta.days_with_data.join(','));
eq('站数', D.meta.station_count, T.meta.station_count);
eq('渠道张数', D.meta.channels.length, T.meta.channels.length);

var days = T.meta.days_with_data;
days.forEach(function (d) {
  var a = D.sale[d], b = T.sale[d];
  if (!a) { fail++; OUT('  ✗ 缺日 ' + d); return; }
  eq('日' + d + ' 总吨', a.t, b.t);
  eq('日' + d + ' 柴油', a.diesel, b.diesel);
  eq('日' + d + ' 汽油', a.gas, b.gas);
  eq('日' + d + ' 含税金额', a.amt, b.amt, 0.05);
  eq('日' + d + ' 不含税万元', a.wan, b.wan);
  eq('日' + d + ' 小额吨', a.xiao.t, b.xiao.t);
  eq('日' + d + ' 站数', a.n, b.n);
});

/* 逐站逐列对比（任一天） */
var d0 = days[0];
var sn = Object.keys(T.sale[d0].stations);
var bad = 0;
sn.forEach(function (name) {
  var a = D.sale[d0].stations[name], b = T.sale[d0].stations[name];
  if (!a) { bad++; return; }
  ['t10', 't0', 't92', 't95', 't', 'l', 'amt', 'wan', 'diesel', 'gas'].forEach(function (k) {
    if (Math.abs(a[k] - b[k]) > 0.02) { bad++; if (bad < 6) OUT('  ✗ 站' + name + '.' + k + ' ' + a[k] + ' vs ' + b[k]); }
  });
  if (a.road !== b.road) { bad++; if (bad < 6) OUT('  ✗ 站' + name + ' 路段 ' + a.road + ' vs ' + b.road); }
});
if (bad === 0) { pass++; OUT('  逐站 30×10 列全部一致'); } else { fail++; OUT('  逐站不一致项 ' + bad); }

/* 渠道对比 */
var chs = Object.keys(T.channels);
var chBad = 0;
chs.forEach(function (ch) {
  days.forEach(function (d) {
    var A = (D.channels[ch] || {})[d] || {}, B = T.channels[ch][d] || {};
    Object.keys(B).forEach(function (st) {
      if (!A[st] || Math.abs(A[st].t - B[st].t) > 0.02) chBad++;
    });
  });
});
if (chBad === 0) { pass++; OUT('  渠道 ' + chs.length + ' 张逐站吨位一致'); } else { fail++; OUT('  渠道不一致项 ' + chBad); }

/* 大客户对比 */
var bA = (D.bigcust['大客户明细'] || {})[d0] || {}, bB = (T.bigcust['大客户明细'] || {})[d0] || {};
eq('大客户站数', Object.keys(bA).length, Object.keys(bB).length);

OUT('\n结果：PASS ' + pass + ' / FAIL ' + fail);
OUT('JSON 体积 ' + JSON.stringify(D).length + ' 字节（真值 ' + JSON.stringify(T).length + '）');
process.exit(fail ? 1 : 0);
