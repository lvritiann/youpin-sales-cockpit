/* ============================================================================
 * 成品油经营数据中转簿 · 同步脚本  transit_sync.js   v1.0   2026-10-29
 * ============================================================================
 * 挂载位置：★ 本脚本挂在「成品油经营数据中转簿」自己身上 ★
 *          （不是填报簿。填报簿那边一格 AirScript 都不用再留）
 *
 * 职责边界（刻意做窄）：
 *   只负责「量」的三层落表 —— 日粒度 / 月汇总 / 站级明细 / 运行日志。
 *   渠道、大客户、车流、data.json 推送仍由填报簿上的 read_daily.js 负责。
 *   两边职责不重叠，任何一边出错都不影响另一边。
 *
 * 为什么挂在「中转簿」而不是「填报簿」：
 *   1) 写自己的子表 = 本地写，read_daily.js 的 archiveToSheet 已验证多年的写法；
 *      跨文档「写」在 WPS 上没有验证过，不做没把握的事。
 *   2) 跨文档「读」已被 flow_push.js 验证可行（KSDrive.openFile + Value2）。
 *   3) 填报簿因此可以彻底变成「纯填报簿」—— 改名、另存、换位置，只要在这里的
 *      「配置」表把 URL 改掉即可，脚本一行都不用动。
 *   4) 灾备：填报簿哪天丢了 / 损坏了，本簿的日粒度全矩阵还能把「量」重算出来，
 *      不必回本地那 34 个 xlsx。
 *
 * 前置条件：
 *   ★ 脚本编辑器左侧「服务」→ 添加「云文档API」（KSDrive 依赖它），
 *     否则 KSDrive 是 undefined，第 1 步就报错。
 *   ★ 「配置」表 A2（键 = 填报簿URL）必须填填报簿的分享链接 https://www.kdocs.cn/l/xxxx
 *
 * 幂等性：
 *   日粒度  键 = 日期 + 站名     （同一天同一站只留一行，重跑覆盖）
 *   月汇总  键 = 会计月          （同一个月只留一行）
 *   站级明细 键 = 会计月 + 站名
 *   跑 10 次和跑 1 次结果完全一样。
 * ==========================================================================*/

var VER = 'v1.0';

var CFG = {
  /* 填报簿 —— URL 优先从「配置」表读，读不到才用这里的兜底值 */
  SRC_URL   : '',

  /* 填报簿子表（只需要「销售日报」一张；渠道/大客户不归本脚本管） */
  SALE_NAME : '销售日报',
  SALE_ROWS : 1601,
  SALE_COLS : 28,

  /* 本簿子表名 */
  S_DAY     : '日粒度',
  S_MON     : '月汇总',
  S_ST      : '站级明细',
  S_CFG     : '配置',
  S_LOG     : '运行日志',

  /* 日粒度最多扫多少行做去重索引。30,502 行历史 + 每月约 900 行增量，
     留到 40000 够用 10 年。★ 不能按「当前行数」动态取 ——
     取小了会把已有历史判成「没见过」，然后从末尾重写一遍，历史行就重复了。 */
  DAY_ROWS  : 40000
};

/* 销售日报列位（0 基）：
   C=-10 D=0# E=92# F=95# G=小计 / 吨 2-6
   升 7-11（H=-10 I=0# J=92# K=95# L=小计）
   金额 12-16（M=-10 N=0# O=92# P=95# Q=小计）
   万元 21（V）                                                                  */
var C = { t10: 2, t0: 3, t92: 4, t95: 5, t: 6,
          l: 11, amt: 16, wan: 21,
          dL1: 7, dL0: 8, gL92: 9, gL95: 10,
          dA1: 12, dA0: 13, gA92: 14, gA95: 15 };

/* 日粒度表 19 列表头 —— 顺序必须与 _kdocs/scan_daymatrix.py 的输出严格一致，
   否则本地生成的 xlsx 与脚本每日写入的行会错位。 */
var DAY_HEAD = ['日期', '会计月', '日号', '路段', '站名',
                '零售吨', '柴油吨', '汽油吨', '-10#吨', '0#吨', '92#吨', '95#吨',
                '升', '柴油升', '汽油升',
                '含税金额元', '柴油金额', '汽油金额', '万元'];
var DAY_COLS = 19;

var AGG = ['小计', '合计', '零售小计', '小额配送', '总计', '累计'];
var XIAO = '小额配送';
var XIAO_ROAD = '公司级';

/* 日号兜底序列：源表恒 31 个日块，标签丢人时按块序取 */
var SEQ = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14',
           '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25',
           '26', '27', '28', '29', '30', '31'];

/* ---------------------------------------------------------------- 日志 */
var LOG = [];
function say(x) {
  var t = String(x);
  LOG.push(t);
  try { console.log(t); } catch (e) {}
}

/* ---------------------------------------------------------------- 基础工具 */
function s(x) { return x === null || x === undefined ? '' : String(x).trim(); }
function num(x) {
  if (x === null || x === undefined) return 0;
  if (typeof x === 'number') return isFinite(x) ? x : 0;
  var t = String(x).replace(/,/g, '').replace(/\s| /g, '').trim();
  if (t === '' || t === '-' || t === '—' || t === '/') return 0;
  var f = parseFloat(t);
  return isNaN(f) ? 0 : f;
}
function r2(v) { return Math.round(v * 100) / 100; }
function r3(v) { return Math.round(v * 1000) / 1000; }
function p2(n) { return (n < 10 ? '0' : '') + n; }

function colName(n) {
  var t = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    t = String.fromCharCode(65 + m) + t;
    n = Math.floor((n - 1) / 26);
  }
  return t;
}

/* ---------------------------------------------------------------- 表访问 */

/* ★ 这里的 Application 指「本脚本宿主 = 中转簿」。
   跨文档读填报簿时不要用这个函数，要用 sheetsOf(srcApp, name)。 */
function sheet(name) { return sheetsOf(Application, name); }

function sheetsOf(app, name) {
  var tries = [
    function () { return app.Sheets.Item(name); },
    function () { return app.Sheets(name); },
    function () { return app.Worksheets.Item(name); },
    function () { return app.Worksheets(name); },
    function () { return app.ActiveWorkbook.Sheets.Item(name); },
    function () { return app.ActiveWorkbook.Worksheets.Item(name); }
  ];
  for (var i = 0; i < tries.length; i++) {
    try { var sh = tries[i](); if (sh) return sh; } catch (e) {}
  }
  return null;
}

function sheetNamesOf(app) {
  var out = [];
  try {
    var S = app.Sheets;
    var n = S.Count || S.length || 0;
    for (var i = 1; i <= n; i++) {
      try { var sh = S.Item(i); if (sh && sh.Name) out.push(String(sh.Name)); } catch (e) {}
    }
  } catch (e) {}
  return out;
}

function getCell(sh, ref) {
  try { var v = sh.Range(ref).Value2; if (v !== undefined && v !== null) return v; } catch (e) {}
  try { var v2 = sh.Range(ref).Value; if (v2 !== undefined && v2 !== null) return v2; } catch (e) {}
  try { var v3 = sh.Range(ref).Text; if (v3 !== undefined && v3 !== null) return v3; } catch (e) {}
  return '';
}

function setCell(sh, ref, v) {
  try { sh.Range(ref).Value2 = v; return true; } catch (e) {}
  try { sh.Range(ref).Value = v; return true; } catch (e) {}
  try { sh.Range(ref).setValue(v); return true; } catch (e) {}
  return false;
}

function readRange(sh, ref) {
  try { var v = sh.Range(ref).Value2; if (v !== undefined && v !== null) return v; } catch (e) {}
  try { var v2 = sh.Range(ref).Value; if (v2 !== undefined && v2 !== null) return v2; } catch (e) {}
  try { var v3 = sh.getRange(ref).getValues(); if (v3) return v3; } catch (e) {}
  try { var v4 = sh.Range(ref).getValues(); if (v4) return v4; } catch (e) {}
  throw new Error('无法读取区域 ' + ref + '（Range/Value2/Value/getValues 全部失败）');
}

/* Value2 可能返回 0 基或 1 基、也可能行是被裁过的锯齿数组。
   不靠 `undefined` 猜，直接拿「请求的行列数」反推基序。 */
function grid(v, nrows, ncols) {
  if (!v || !v.length) return [];
  var L = v.length;
  var oneRow = (L === nrows + 1) ? true : (L === nrows ? false : (v[0] === undefined));

  var sample = null;
  for (var i = oneRow ? 1 : 0; i < L; i++) {
    if (v[i] && v[i].length) { sample = v[i]; break; }
  }
  var cl = sample ? sample.length : ncols;
  var oneCol = (cl === ncols + 1) ? true : (cl === ncols ? false : (sample && sample[0] === undefined));

  var out = [];
  for (var r = oneRow ? 1 : 0; r < L; r++) {
    var src = v[r] || [];
    var row = [];
    for (var c = 0; c < ncols; c++) row.push(src[oneCol ? c + 1 : c]);
    out.push(row);
  }
  return out;
}

/* 只取单列（去重索引专用）—— 比整块 Value2 省一半以上时间。
   ★ Value2 对单列区域可能返回一维 [v1,v2,…]，也可能返回二维 [[v1],[v2],…]，
     不要用 `instanceof Array` 判（AirScript 环境里不一定靠得住），按"是不是对象"判最稳。 */
function colValues(sh, letter, rows) {
  var raw = readRange(sh, letter + '1:' + letter + rows);
  var out = [], i, x;
  if (!raw) return out;
  for (i = 0; i < raw.length; i++) {
    x = raw[i];
    if (x === null || x === undefined) { out.push(''); continue; }
    if (typeof x === 'object') out.push(x.length ? x[0] : '');
    else out.push(x);
  }
  return out;
}

/* ---------------------------------------------------------------- 日期 */

/* 今天是几号（本地时区），返回 'YYYY-MM-DD' */
function today0() {
  var d = new Date();
  return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate());
}

/* 会计月：日号 >= 26 归起月（本月），否则归上月
   ★ 与 sync.py / read_daily.js 同规则，改一处必须三处一起改 */
function acctFromToday() {
  var d = new Date();
  var y = d.getFullYear(), m = d.getMonth() + 1, dd = d.getDate();
  if (dd >= 26) { m += 1; if (m > 12) { m = 1; y += 1; } }
  else { m -= 1; if (m < 1) { m = 12; y -= 1; } }
  return y + '-' + p2(m);
}

/* 会计月标签 'YYYY-MM' → { y, m, start:'YYYY-MM-DD', end:'YYYY-MM-DD' }
   2026-10 → 2026-09-26 ~ 2026-10-25 */
function acctSpan(lab) {
  var p = String(lab).split('-');
  var y = parseInt(p[0], 10), m = parseInt(p[1], 10);
  if (!y || !m) return null;
  var sy = m > 1 ? y : y - 1, sm = m > 1 ? m - 1 : 12;
  return {
    y: y, m: m,
    start: sy + '-' + p2(sm) + '-26',
    end: y + '-' + p2(m) + '-25'
  };
}

/* 日号 → 真实日期。>=26 归上月，<=25 归本月。
   ★ 源表恒 31 个日块、日号硬排 1..31，不随实际天数收缩。
     2 月区间会遇到 30/31 —— 目标月没这天，返回 null，调用方跳过。 */
function realDate(span, day) {
  var d = parseInt(day, 10);
  if (!d || d < 1 || d > 31) return null;
  var y, m;
  if (d >= 26) { y = span.m > 1 ? span.y : span.y - 1; m = span.m > 1 ? span.m - 1 : 12; }
  else { y = span.y; m = span.m; }
  var dt = new Date(y, m - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return null;  // 溢出
  return y + '-' + p2(m) + '-' + p2(d);
}

/* ---------------------------------------------------------------- 日块切分 */

function startsAt(rows, key, col) {
  var out = [], i;
  for (i = 0; i < rows.length; i++) {
    if (s((rows[i] || [])[col]).indexOf(key) >= 0) out.push(i);
  }
  return out;
}

/* ★ 块首只认 C 列（下标 2）。日报每张子表最右侧还挂着一张「当日销量（吨）」小表
   （Y/Z 列），若按「任意列含关键词」判块首，那块小表会把一个日块从中间切成两段。 */
function dayBlocks(rows, key) {
  var c = 2;
  var starts = startsAt(rows, key, c);
  if (!starts.length) {
    for (var j = 0; j < 26 && !starts.length; j++) {
      if (j !== c) starts = startsAt(rows, key, j);
    }
  }
  var bl = [], k;
  for (k = 0; k < starts.length; k++) {
    bl.push({
      label: s((rows[starts[k]] || [])[0]),
      rows: rows.slice(starts[k], k + 1 < starts.length ? starts[k + 1] : rows.length)
    });
  }
  var res = {}, order = [], idx = 0;
  for (k = 0; k < bl.length; k++) {
    if (bl[k].label.indexOf('-') >= 0) continue;        // 会计月累计块（'26-25'），丢弃
    var lab = bl[k].label, day = null;
    if (/^\d{1,2}$/.test(lab) && +lab >= 1 && +lab <= 31) day = String(+lab);
    if (!day) {
      day = SEQ[idx] || null;
      say('  [warn] 第 ' + (k + 1) + ' 个日块标签「' + lab + '」非日号，按块序取 ' + day);
    }
    idx++;
    if (!day) continue;
    if (!res[day]) order.push(day);
    res[day] = bl[k].rows;
  }
  return { map: res, order: order };
}

var NOTNAME = ['柴油', '汽油', '小计', '合计', '零售小计', '小额配送', '总计', '累计',
               '-10', '- 10', '0', '92', '95', '26-25'];
function isName(x) {
  var t = s(x);
  if (!t) return false;
  if (NOTNAME.indexOf(t) >= 0) return false;
  if (t.indexOf('当日') >= 0 || t.indexOf('销量') >= 0) return false;
  if (/^[\d\-－—.]+$/.test(t)) return false;
  return true;
}
function isRoad(x) {
  var t = s(x);
  return !!t && NOTNAME.indexOf(t) < 0 && !/^[\d\-－—.]+$/.test(t);
}

/* ---------------------------------------------------------------- 读配置 */

function readCfg() {
  var sh = sheet(CFG.S_CFG);
  if (!sh) { say('[warn] 本簿没有「' + CFG.S_CFG + '」表，配置全部走代码兜底'); return {}; }
  var raw = readRange(sh, 'A1:B30');
  var out = {}, i, row;
  for (i = 0; i < (raw || []).length; i++) {
    row = raw[i] || [];
    var k = s(row[0]);
    if (k && k !== '键') out[k] = row[1];
  }
  return out;
}

/* ---------------------------------------------------------------- 主流程 */

function main() {
  say('=== 中转簿同步 ' + VER + '  ' + today0() + ' ===');

  /* --- 1) 配置 --- */
  var cfg = readCfg();
  var srcUrl = s(cfg['填报簿URL']) || CFG.SRC_URL;
  if (!srcUrl) {
    say('❌ 「配置」表 A2（键 = 填报簿URL）没填，无法跨文档读填报簿。');
    say('   去「配置」表把 B2 填成 https://www.kdocs.cn/l/xxxx 再跑一次。');
    return;
  }
  var acLab = s(cfg['会计月']) || acctFromToday();
  var span = acctSpan(acLab);
  if (!span) { say('❌ 会计月「' + acLab + '」不是 YYYY-MM 格式'); return; }
  say('会计月 ' + acLab + '（' + span.start + ' ~ ' + span.end + '）'
      + (s(cfg['会计月']) ? '（取自配置表）' : '（按今天自动推导）'));
  say('填报簿 ' + srcUrl);

  /* --- 2) 跨文档读填报簿 --- */
  if (typeof KSDrive === 'undefined') {
    say('❌ KSDrive 未定义 —— 左侧「服务」里没添加「云文档API」，加上再跑。');
    return;
  }
  var f = null, saleRows = null;
  try {
    f = KSDrive.openFile(srcUrl);
    if (!f) throw new Error('openFile 返回空');
    var saleSh = sheetsOf(f.Application, CFG.SALE_NAME);
    if (!saleSh) {
      throw new Error('填报簿里找不到「' + CFG.SALE_NAME + '」子表（现有：'
                      + sheetNamesOf(f.Application).join(' / ') + '）');
    }
    var ref = 'A1:' + colName(CFG.SALE_COLS) + CFG.SALE_ROWS;
    saleRows = grid(readRange(saleSh, ref), CFG.SALE_ROWS, CFG.SALE_COLS);
    say('已读填报簿「' + CFG.SALE_NAME + '」' + saleRows.length + ' 行');
  } catch (e) {
    say('❌ 读填报簿失败：' + (e && e.message ? e.message : e));
    say('   常见原因：① 链接不是「分享链接」而是浏览器地址栏 URL；'
        + '② 填报簿改名/另存导致分享失效；③ 当前账号无该文件权限。');
    return;
  } finally {
    /* 只读用，读完立刻释放，别占着文件锁（catch 分支也走这里，不重复 close） */
    if (f) { try { f.close(); } catch (_) {} }
  }

  /* --- 3) 切日块 --- */
  var db = dayBlocks(saleRows, '当日销量（吨）');
  say('识别到 ' + db.order.length + ' 个日块');

  var T = today0();
  var days = [], per = {};      // per[date] = { stations:{}, xiao:{} }
  var totRet = 0, totXiao = 0;
  var i, k;

  for (i = 0; i < db.order.length; i++) {
    var day = db.order[i];
    var dt = realDate(span, day);
    if (!dt) continue;                       // 该月没这天（2 月的 30/31 块）
    if (dt > T) continue;                    // 未来日块不入库
    var blk = db.map[day];
    var rec = { stations: {}, xiao: null };
    var road = null, j, r, xr = null;

    for (j = 0; j < blk.length; j++) {
      r = blk[j] || [];
      if (r.length < 7) continue;
      var a = r[0], b = r[1];
      if (s(a) === XIAO) { xr = r; continue; }
      if (s(a) === '零售小计' || s(a) === '合计' || s(a) === '小计') continue;
      if (isRoad(a)) road = s(a);
      if (!isName(b)) continue;
      if (/^\d+$/.test(s(b))) continue;
      var t = num(r[C.t]);
      if (t <= 0) continue;
      rec.stations[s(b)] = {
        road: road || '',
        t: t,
        diesel: num(r[C.t10]) + num(r[C.t0]),
        gas: num(r[C.t92]) + num(r[C.t95]),
        t10: num(r[C.t10]), t0: num(r[C.t0]),
        t92: num(r[C.t92]), t95: num(r[C.t95]),
        l: num(r[C.l]),
        dL: num(r[C.dL1]) + num(r[C.dL0]),
        gL: num(r[C.gL92]) + num(r[C.gL95]),
        amt: num(r[C.amt]),
        dA: num(r[C.dA1]) + num(r[C.dA0]),
        gA: num(r[C.gA92]) + num(r[C.gA95]),
        wan: num(r[C.wan])
      };
    }
    if (xr) {
      rec.xiao = {
        t: num(xr[C.t]),
        diesel: num(xr[C.t10]) + num(xr[C.t0]),
        gas: num(xr[C.t92]) + num(xr[C.t95]),
        t10: num(xr[C.t10]), t0: num(xr[C.t0]),
        t92: num(xr[C.t92]), t95: num(xr[C.t95]),
        l: num(xr[C.l]),
        dL: num(xr[C.dL1]) + num(xr[C.dL0]),
        gL: num(xr[C.gL92]) + num(xr[C.gL95]),
        amt: num(xr[C.amt]),
        dA: num(xr[C.dA1]) + num(xr[C.dA0]),
        gA: num(xr[C.gA92]) + num(xr[C.gA95]),
        wan: num(xr[C.wan])
      };
    }

    var names = Object.keys(rec.stations);
    if (!names.length && (!rec.xiao || rec.xiao.t <= 0)) continue;   // 空日块

    per[dt] = rec;
    days.push(dt);
    var dRet = 0, dX = 0;
    for (k = 0; k < names.length; k++) dRet += rec.stations[names[k]].t;
    if (rec.xiao) dX = rec.xiao.t;
    totRet += dRet; totXiao += dX;
    say('  ' + dt + '：' + names.length + ' 站 / 零售 ' + r2(dRet)
        + ' 吨 + 小额 ' + r2(dX) + ' 吨');
  }

  if (!days.length) {
    say('❌ 当前会计月 ' + acLab + ' 一个有效日块都没有 —— 填报簿还没填这个月？');
    return;
  }
  days.sort();
  say('有效天数 ' + days.length + '，零售 ' + r2(totRet) + ' 吨 + 小额 ' + r2(totXiao)
      + ' 吨 = 全口径 ' + r2(totRet + totXiao) + ' 吨');

  /* --- 4) 写日粒度 --- */
  var nDay = writeDays(acLab, days, per);

  /* --- 5) 写站级明细 + 月汇总 --- */
  var stMap = {}, stOrder = [];
  for (i = 0; i < days.length; i++) {
    var dd = days[i], ss = per[dd].stations;
    for (var nm in ss) {
      if (!ss.hasOwnProperty(nm)) continue;
      if (!stMap[nm]) {
        stMap[nm] = { road: ss[nm].road || '', t: 0, d: 0, g: 0, l: 0, a: 0, w: 0, nd: 0 };
        stOrder.push(nm);
      }
      var o = stMap[nm];
      o.t += ss[nm].t; o.d += ss[nm].diesel; o.g += ss[nm].gas;
      o.l += ss[nm].l; o.a += ss[nm].amt; o.w += ss[nm].wan; o.nd++;
    }
  }
  stOrder.sort(function (x, y) { return stMap[y].t - stMap[x].t; });
  var nSt = writeStations(acLab, stOrder, stMap);
  var nMon = writeMonth(acLab, span, totRet, totXiao, stOrder.length, days.length);

  /* --- 6) 运行日志 --- */
  writeLog(acLab, days.length, totRet, totXiao, nSt, nDay);

  say('=== 完成：日粒度 ' + nDay + ' 格系 / 站级 ' + nSt + ' 行 / 月汇总 ' + nMon + ' 行 ===');
}

/* ---------------------------------------------------------- 日粒度 upsert */

/* 组装日粒度一行的 19 个格。站点行与小额行共用，避免两处各写一遍后漂移。 */
function dayCells(dt, acLab, road, name, v) {
  return [dt, acLab, parseInt(dt.substring(8), 10), road, name,
          r2(v.t), r2(v.diesel), r2(v.gas),
          r2(v.t10), r2(v.t0), r2(v.t92), r2(v.t95),
          Math.round(v.l), Math.round(v.dL), Math.round(v.gL),
          r2(v.amt), r2(v.dA), r2(v.gA), r3(v.wan)];
}

function writeDays(acLab, days, per) {
  var sh = sheet(CFG.S_DAY);
  if (!sh) { say('[warn] 本簿没有「' + CFG.S_DAY + '」表，跳过日粒度'); return 0; }

  /* ★ 版本护栏：先数一下表头实际有几列。
     旧版是 11 列（…升/含税金额元/万元），新版 19 列。若拿 19 列的写法往 11 列的表里灌，
     每一行的「柴油升」会覆盖掉「含税金额元」、「万元」会写到根本不存在的列上，
     整张表从此错位且**不会报错** —— 这种静默损坏比直接崩掉危险得多。
     所以列数不符时一律停下，让人去换表，不在错误的结构上继续写。 */
  var curCols = 0;
  while (curCols < 40 && s(getCell(sh, colName(curCols + 1) + '1'))) curCols++;
  if (curCols > 0 && curCols !== DAY_COLS) {
    say('❌ 「' + CFG.S_DAY + '」表是旧版：实测 ' + curCols + ' 列，本脚本需要 '
        + DAY_COLS + ' 列。');
    say('   请重新上传新版「成品油经营数据中转簿.xlsx」（3.2 MB / 日粒度 19 列），');
    say('   或在表里手工补齐到 19 列后再跑。本次跳过日粒度，其余表照常。');
    return 0;
  }

  var needHead = false, h;
  for (h = 0; h < DAY_HEAD.length; h++) {
    if (s(getCell(sh, colName(h + 1) + '1')) !== DAY_HEAD[h]) { needHead = true; break; }
  }
  if (needHead) {
    for (h = 0; h < DAY_HEAD.length; h++) setCell(sh, colName(h + 1) + '1', DAY_HEAD[h]);
    say('已补「' + CFG.S_DAY + '」表头（' + DAY_COLS + ' 列）');
  }

  /* 去重索引：日期(A) + 站名(E) 两列分开读，比整块 Value2 省时间 */
  var dCol = colValues(sh, 'A', CFG.DAY_ROWS);
  var nCol = colValues(sh, 'E', CFG.DAY_ROWS);
  var used = {}, last = 1, gi;
  for (gi = 1; gi < dCol.length; gi++) {
    var dk = s(dCol[gi]);
    var nk = s(nCol[gi]);
    if (!dk || !nk) continue;
    used[dk + '|' + nk] = gi + 1;
    if (gi + 1 > last) last = gi + 1;
  }
  say('日粒度索引 ' + Object.keys(used).length + ' 条，末行 ' + last);

  var wrote = 0, added = 0, c;
  for (var i = 0; i < days.length; i++) {
    var dt = days[i], rec = per[dt];
    var names = Object.keys(rec.stations);
    var j;
    /* 站点行 */
    for (j = 0; j < names.length; j++) {
      var sn = names[j];
      var key = dt + '|' + sn;
      var row = used[key] || (++last);
      if (!used[key]) added++;
      var cells = dayCells(dt, acLab, rec.stations[sn].road, sn, rec.stations[sn]);
      for (c = 0; c < cells.length; c++) setCell(sh, colName(c + 1) + row, cells[c]);
      used[key] = row;
      wrote++;
    }
    /* 小额：公司级虚拟实体，单独一行，与 15 对站平权 */
    if (rec.xiao && (rec.xiao.t > 0 || rec.xiao.amt > 0)) {
      var xk = dt + '|' + XIAO;
      var xr = used[xk] || (++last);
      if (!used[xk]) added++;
      var xc = dayCells(dt, acLab, XIAO_ROAD, XIAO, rec.xiao);
      for (c = 0; c < xc.length; c++) setCell(sh, colName(c + 1) + xr, xc[c]);
      used[xk] = xr;
      wrote++;
    }
  }
  say('日粒度：覆盖 ' + wrote + ' 行（其中新增 ' + added + ' 行），末行 ' + last);
  return wrote;
}

/* -------------------------------------------------------- 站级明细 upsert */

function writeStations(acLab, order, map) {
  var sh = sheet(CFG.S_ST);
  if (!sh) { say('[warn] 本簿没有「' + CFG.S_ST + '」表，跳过站级明细'); return 0; }

  var HEAD = ['会计月', '站名', '路段', '零售吨', '柴油吨', '汽油吨',
              '升', '含税金额元', '万元'];
  if (s(getCell(sh, 'A1')) !== HEAD[0]) {
    for (var h = 0; h < HEAD.length; h++) setCell(sh, colName(h + 1) + '1', HEAD[h]);
    say('已补「' + CFG.S_ST + '」表头');
  }

  var aCol = colValues(sh, 'A', 6000);
  var bCol = colValues(sh, 'B', 6000);
  var used = {}, last = 1, gi;
  for (gi = 1; gi < aCol.length; gi++) {
    var ak = s(aCol[gi]), bk = s(bCol[gi]);
    if (!ak || !bk) continue;
    used[ak + '|' + bk] = gi + 1;
    if (gi + 1 > last) last = gi + 1;
  }

  var wrote = 0;
  for (var i = 0; i < order.length; i++) {
    var sn = order[i], v = map[sn];
    var key = acLab + '|' + sn;
    var row = used[key] || (++last);
    var cells = [acLab, sn, v.road, r2(v.t), r2(v.d), r2(v.g),
                 '', r2(v.a), r3(v.w)];      // 升：归档口径无逐站加注升数，留空不造数
    for (var c = 0; c < cells.length; c++) setCell(sh, colName(c + 1) + row, cells[c]);
    used[key] = row;
    wrote++;
  }
  say('站级明细：' + acLab + ' 共 ' + wrote + ' 行（末行 ' + last + '）');
  return wrote;
}

/* ---------------------------------------------------------- 月汇总 upsert */

function writeMonth(acLab, span, totRet, totXiao, nSt, nDay) {
  var sh = sheet(CFG.S_MON);
  if (!sh) { say('[warn] 本簿没有「' + CFG.S_MON + '」表，跳过月汇总'); return 0; }

  var HEAD = ['会计月', '区间', '零售吨', '柴油吨', '汽油吨', '小额吨',
              '全口径吨', '含税金额元', '万元', '站数', '天数', '写入时间'];
  if (s(getCell(sh, 'A1')) !== HEAD[0]) {
    for (var h = 0; h < HEAD.length; h++) setCell(sh, colName(h + 1) + '1', HEAD[h]);
    say('已补「' + CFG.S_MON + '」表头');
  }

  /* 月汇总只有 30 多行，直接扫 A 列 */
  var aCol = colValues(sh, 'A', 400);
  var target = 0, firstEmpty = 0, i;
  for (i = 1; i < aCol.length; i++) {
    var cv = s(aCol[i]);
    if (cv === acLab) { target = i + 1; break; }
    if (!cv && !firstEmpty) firstEmpty = i + 1;
  }
  if (!target) target = firstEmpty || (aCol.length + 1);

  /* 柴油 / 汽油 / 金额 从站级明细当月行回算 ——
     日粒度和站级明细是同一批数，回算就不会出现「月汇总与明细对不上」 */
  var stSh = sheet(CFG.S_ST);
  var die = 0, gas = 0, amt = 0, wan = 0;
  if (stSh) {
    var gridSt = grid(readRange(stSh, 'A1:I6000'), 6000, 9);
    for (i = 1; i < gridSt.length; i++) {
      if (s(gridSt[i][0]) !== acLab) continue;
      die += num(gridSt[i][4]); gas += num(gridSt[i][5]);
      amt += num(gridSt[i][7]); wan += num(gridSt[i][8]);
    }
  }

  var line = [acLab, span.start + ' ~ ' + span.end,
              r2(totRet), r2(die), r2(gas), r2(totXiao), r2(totRet + totXiao),
              r2(amt), r3(wan), nSt, nDay,
              today0() + ' ' + VER];
  for (var c = 0; c < line.length; c++) setCell(sh, colName(c + 1) + target, line[c]);
  say('月汇总：' + acLab + ' → 第 ' + target + ' 行  零售 ' + r2(totRet)
      + ' + 小额 ' + r2(totXiao) + ' = ' + r2(totRet + totXiao) + ' 吨');
  return 1;
}

/* ---------------------------------------------------------------- 运行日志 */

function writeLog(acLab, nDay, totRet, totXiao, nSt, nDayRow) {
  var sh = sheet(CFG.S_LOG);
  if (!sh) return;
  var aCol = colValues(sh, 'A', 2000);
  var row = 1, i;
  for (i = 1; i < aCol.length; i++) { if (s(aCol[i])) row = i + 2; else break; }
  if (row < 2) row = 2;
  var d = new Date();
  var line = [d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate())
              + ' ' + p2(d.getHours()) + ':' + p2(d.getMinutes()),
              VER, acLab, nDay, r2(totRet), r2(totXiao), r2(totRet + totXiao),
              nSt, nDayRow, '落表成功', ''];
  for (var c = 0; c < line.length; c++) setCell(sh, colName(c + 1) + row, line[c]);
}

/* AirScript「运行」只执行顶层代码，不主动调 main() —— 这行不能删 */
main();
