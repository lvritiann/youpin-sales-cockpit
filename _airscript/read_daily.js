/* ==========================================================================
 * 成品油日报 → GitHub _incoming/data.json   v2.4
 * 运行位置：金山文档「202611销售及库存日报（新).xlsx」内的 AirScript
 * 触发方式：定时任务（8:30 / 9:00 / 9:30 / 10:00 / 11:30 各挂一条，跑同一个 main）
 *
 * v2.4 变更（2026-10-29，修 401 Requires authentication）
 *  1) ★ HTTP 调用改为优先走 HTTP.fetch(url,{method,headers,body,timeout})。
 *     原先优先调 HTTP.get / HTTP.put 短接口，而它们【不吃 opts.headers】——
 *     传进去不报错，Authorization 却悄悄丢了，GitHub 回 401 "Requires authentication"，
 *     现象像令牌失效，其实令牌根本没发出去。HTTP.fetch 是 flow_push.js 已验证可用的形式。
 *  2) ghGet 不再静默吞错：401/404/其它状态各打一条日志，并把远端现有体积与 sha 打出来。
 *  3) 新增体积护栏：新内容小于远端 85% 时中止推送（防某天只读到 1 个日块把 3 天冲掉）。
 *
 * v2.3 变更（2026-10-29，修 GET 响应解析崩溃）
 *  1) ★ ghGet 原先写 JSON.parse(r.body || r.text || '{}')。HTTP.fetch 的返回对象
 *     【没有 body 属性】，只有 text() 方法 —— r.body 为 undefined 就落到 r.text，
 *     而那是函数对象，JSON.parse 隐式 String(函数) 得到 "text(){return ...}"，
 *     抛 SyntaxError: Unexpected token 'e', "text(){retu"... is not valid JSON。
 *     新增 respText() 三级兜底（函数 → 字符串 → 其它），两种返回形态都能取正文。
 *  2) PUT 结果判断加 Number() 归一 —— 有的实现 status 给字符串 '200'，
 *     原先用 === 比数字会误判成失败，明明推上去了却报失败。
 *  3) GET 返回非 JSON 时不再崩溃，按「远端无文件」处理并继续推送。
 *
 * v2.2 变更（2026-10-29，「驾驶舱配置」表不存在导致推送中断）
 *  1) ★ 令牌改为两处来源，取到其一即可：①「驾驶舱配置」B1  ② 代码顶部 CFG.TOKEN。
 *     实测该簿 63 张子表里并没有「驾驶舱配置」，v2.1 会直接抛错停在推送前一步。
 *     现在没有这张表也照常推送，只是少了「C 列日志」和「A1 会计月回写」两项便利。
 *  2) 日志会打印「令牌来源：…」，一眼看出用的是哪一处，不用猜。
 *  3) 两处都空才报错，错误信息直接给出两种补法。
 * v2.1 变更（2026-10-29 开启自动推送）
 *  1) ★ CFG.PUSH 出厂值改为 true —— 从本版起，定时任务跑完会真的推送到 GitHub。
 *     想临时只看不推，把 PUSH 改回 false 即可（其余逻辑不变）。
 *  2) 令牌读取加清洗：从单元格粘贴过来的值常带引号、空格、换行，
 *     或被人连 "Bearer " 一起粘进来 —— 任一情况都会让 Authorization 头非法。
 *     新函数 cleanTok() 统一剪掉，只留令牌本体。
 *  3) 令牌 B1 缺失 / 清洗后为空 → 明确报错，不静默跳过。
 *
 * v2 变更（针对首跑报错「驾驶舱配置!A1 需写会计月」）
 *  1) 会计月改为「A1 有就认 A1，没有就按今天自动推导」——不再强制手工维护；
 *     推导结果会回写 A1，让你在表里看得见。
 *  2) 日志同时打到 console 与「驾驶舱配置」C 列；没有该表也能在控制台看全部结果。
 *  3) 取表 / 取格 / 取区域 全部加多重兼容（Range·Value2·Value·getValues）。
 *  4) 启动先打印所有子表名 —— 取表失败时一眼看出是名字不对还是取法不对。
 *
 * ★ 运行注意（2026-10-29 实测确认）
 *   AirScript 2.0 的「运行」按钮只执行顶层代码，不会主动调用 main()。
 *   本文件末尾第 581 行的 main(); 不能删 —— 删了就是一跑 100ms、一行日志都没有。
 *
 * 设计原则
 *  1) 只读源表，绝不写回任何业务单元格；唯一写入是「驾驶舱配置」的状态格。
 *  2) 幂等：先 GET 远端 data.json，内容相同直接退出，不产生提交。
 *     → 「补发」不是独立机制，而是幂等的副产品：上次失败则远端仍是旧内容，下一档自动重推。
 *  3) 产出结构与本地 extract.py 的 data.json 完全同构，下游 build_cockpit2.py 零改动。
 *  4) 全零日块 = 未填报，跳过（不可补零，会把日均拉低）。
 * ========================================================================== */

var VER = 'v2.4';

var CFG = {
  OWNER : 'lvritiann',
  REPO  : 'youpin-sales-cockpit',
  BRANCH: 'main',
  PATH  : '_incoming/data.json',

  SHEET_CFG : '驾驶舱配置',   // 可选表：A1=会计月(留空自动推导), B1=GitHub令牌, C1..=运行日志
  TOKEN     : '',            // ★ 令牌兜底位（v2.2 新增）
                             //   取值优先级：①「驾驶舱配置」B1  →  ② 这里的 TOKEN
                             //   没有这张表、或 B1 空着，就用这行；两边都空才报错。
                             //   不想在代码里放令牌，就建一张「驾驶舱配置」表把令牌填 B1。
  PUSH      : true,          // ★ v2.1 起默认 true：跑完会真的推送到 GitHub
                             //   想只看不推（例如核对口径）临时改回 false，其余逻辑不变

  SALE : { name: '销售日报', rows: 1601, cols: 28 },
  CH   : { rows: 1537, cols: 22 },
  CUST : { name: '大客户明细', rows: 1547, cols: 162 }
};

/* 销售日报列（0 基）：吨 2-6 / 升 7-11 / 销售额元 12-16 / 万元 17-21
   ★ 常驻字段只留汇总值；_ 开头的分油品中间量折完即删，不进 data.json */
var C_SALE = { t10: 2, t0: 3, t92: 4, t95: 5, t: 6, l: 11, amt: 16, wan: 21,
               _dL1: 7, _dL0: 8, _gL92: 9, _gL95: 10,
               _dA1: 12, _dA0: 13, _gA92: 14, _gA95: 15 };
var C_AGG  = { t: 6, l: 11, amt: 16, wan: 21 };
var CHANNELS = ['现金', 'POS', '微信', '支付宝', '工贸卡', '壳牌卡', '大客户',
                '通四海', '定期结算', '回罐', '内结'];
var AGG = ['小计', '合计', '零售小计', '小额配送', '总计', '累计'];
var CLIENT_COLS = [[12, '中融致远'], [22, '锦航石化'], [32, '大连润途'], [42, '陕西卡一车'],
                   [52, '能链石化'], [62, '中润国泰'], [72, '潍坊星油'], [102, '运达通'],
                   [142, '通四海']];
var SEQ = ['26', '27', '28', '29', '30', '31', '1', '2', '3', '4', '5', '6', '7', '8',
           '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20',
           '21', '22', '23', '24', '25'];

/* ---------------------------------------------------------------- 日志 */
var LOG = [];
function say(x) {
  var t = String(x);
  LOG.push(t);
  try { console.log(t); } catch (e) {}
}
function flush() {
  var sh = sheet(CFG.SHEET_CFG);
  if (!sh) return;
  for (var i = 0; i < LOG.length && i < 60; i++) {
    try { setCell(sh, 'C' + (i + 1), LOG[i]); } catch (e) {}
  }
}

/* ---------------------------------------------------------------- 基础工具 */

function s(x) { return x === null || x === undefined ? '' : String(x).trim(); }

/* GitHub 令牌清洗。
   从「驾驶舱配置」B1 粘过来的值经常带这些杂质：
     · 前后空格 / 中间换行（单元格里按了 Alt+Enter）
     · 英文或中文引号： "ghp_xxx"  'ghp_xxx'  “ghp_xxx”
     · 连 "Bearer " 或 "token " 一起粘进来
   任何一种都会让 Authorization 头变成非法值，GitHub 直接 401。
   这里统一剥到只剩令牌本体，减少一类"看起来填了却推不动"的故障。 */
function cleanTok(v) {
  var t = s(v);
  if (!t) return '';
  t = t.replace(/[\s\u00a0\u3000]/g, '');            // 半角/全角空格、换行、制表符
  t = t.replace(/^Bearer/i, '').replace(/^token/i, '');
  t = t.replace(/^['"“”‘’]+/, '').replace(/['"“”‘’]+$/, '');
  return t;
}

function num(x) {
  if (x === null || x === undefined) return 0;
  if (typeof x === 'number') return isFinite(x) ? x : 0;
  var t = String(x).replace(/,/g, '').replace(/\s| /g, '').trim();
  if (t === '' || t === '-' || t === '—' || t === '/') return 0;
  var f = parseFloat(t);
  return isNaN(f) ? 0 : f;
}

function r3(v) { return Math.round(v * 1000) / 1000; }
function r2(v) { return Math.round(v * 100) / 100; }
function p2(n) { return (n < 10 ? '0' : '') + n; }

/* ---------------------------------------------------------------- 表/格访问 */

function sheet(name) {
  var tries = [
    function () { return Application.Sheets.Item(name); },
    function () { return Application.Sheets(name); },
    function () { return Application.Worksheets.Item(name); },
    function () { return Application.Worksheets(name); },
    function () { return Application.ActiveWorkbook.Sheets.Item(name); },
    function () { return Application.ActiveWorkbook.Worksheets.Item(name); }
  ];
  for (var i = 0; i < tries.length; i++) {
    try { var sh = tries[i](); if (sh) return sh; } catch (e) {}
  }
  return null;
}

function sheetNames() {
  var out = [];
  try {
    var S = Application.Sheets;
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
   不再靠 `undefined` 猜 —— 直接拿「请求的行列数」反推基序，最稳。 */
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

function colName(n) {
  var t = '';
  while (n > 0) { var m = (n - 1) % 26; t = String.fromCharCode(65 + m) + t; n = Math.floor((n - 1) / 26); }
  return t;
}

function readSheet(name, rows, cols) {
  var sh = sheet(name);
  if (!sh) throw new Error('找不到子表：' + name + '（现有：' + sheetNames().join(' / ') + '）');
  var ref = 'A1:' + colName(cols) + rows;
  var raw = readRange(sh, ref);
  return { rows: grid(raw, rows, cols), sh: sh };
}

/* ---------------------------------------------------------------- 日块切分 */

/* ★ 关键：块首只认第 3 列（C 列，下标 2）。
   日报每张子表最右侧还挂着一张「当日销量（吨）」小表（Y/Z 列），
   若按「任意列含关键词」判块首，那块小表会把一个日块从中间切成两段，
   第二段丢掉 A 列日号，只能靠 SEQ 兜底 → 日期整体错位、站点数缺一半。 */
function blocksOf(rows, key, col) {
  var c = (col === undefined || col === null) ? 2 : col;
  var starts = startsAt(rows, key, c);
  if (!starts.length) {
    /* 某些子表排版不同（如渠道表），C 列认不到再退化为全列扫描 */
    for (var j = 0; j < 26 && !starts.length; j++) {
      if (j !== c) starts = startsAt(rows, key, j);
    }
  }
  var out = [];
  for (var k = 0; k < starts.length; k++) {
    out.push({
      label: s((rows[starts[k]] || [])[0]),
      row  : starts[k],
      rows : rows.slice(starts[k], k + 1 < starts.length ? starts[k + 1] : rows.length)
    });
  }
  return out;
}

function startsAt(rows, key, col) {
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    if (s((rows[i] || [])[col]).indexOf(key) >= 0) out.push(i);
  }
  return out;
}

function dayBlocks(rows, key) {
  var bl = blocksOf(rows, key), res = {}, order = [];
  var idx = 0;
  for (var k = 0; k < bl.length; k++) {
    if (bl[k].label.indexOf('-') >= 0) continue;      // 会计月累计块（'26-25'），丢弃
    var lab = bl[k].label, day = null;
    if (/^\d{1,2}$/.test(lab) && +lab >= 1 && +lab <= 31) day = String(+lab);
    if (!day) {
      day = SEQ[idx] || null;
      say('[warn] 第 ' + (k + 1) + ' 个日块标签「' + lab + '」非日号，按块序取 ' + day);
    }
    idx++;
    if (!day) continue;
    res[day] = bl[k].rows;
    if (order.indexOf(day) < 0) order.push(day);
  }
  return { map: res, order: order };
}

/* 子表头行里那些不是站名的东西：-10 / 0 / 92 / 95 / 柴油 / 汽油 / 26-25 / 小计 … */
var NOTNAME = ['柴油', '汽油', '小计', '合计', '零售小计', '小额配送', '总计', '累计',
               '-10', '- 10', '0', '92', '95', '26-25'];

function isName(x) {
  var t = s(x);
  if (!t) return false;
  if (NOTNAME.indexOf(t) >= 0) return false;
  if (t.indexOf('当日') >= 0 || t.indexOf('销量') >= 0) return false;
  if (/^[-－—:]+$/.test(t)) return false;                       // '---' 之类的分隔线
  if (/^[-－—+]?\s*\d+(\.\d+)?$/.test(t)) return false;          // 纯数字 / 负数
  if (/^\d{1,2}\s*[-－—~～]\s*\d{1,2}$/.test(t)) return false;   // 26-25
  return true;
}

function stations(blk, cols) {
  var out = {}, road = null;
  for (var i = 0; i < blk.length; i++) {
    var r = blk[i];
    if (!r || r.length < 2) continue;
    var c0 = s(r[0]), c1 = s(r[1]);
    if (c0 && AGG.indexOf(c0) < 0 && !/^[\d\-－—]/.test(c0)) road = c0;
    if (!isName(c1)) continue;
    var o = { road: road || c0 };
    for (var k in cols) { if (cols.hasOwnProperty(k)) o[k] = num(r[cols[k]]); }
    out[c1] = o;
  }
  return out;
}

/* ---------------------------------------------------------------- 会计月 */

function acctMonth() {
  var txt = '';
  var sh = sheet(CFG.SHEET_CFG);
  if (sh) { try { txt = s(getCell(sh, 'A1')); } catch (e) { txt = ''; } }

  var m = txt.match(/(\d{4})-(\d{1,2})-(\d{1,2})\s*[~～]\s*(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) {
    return { text: txt, sY: +m[1], sM: +m[2], sD: +m[3], eY: +m[4], eM: +m[5], eD: +m[6], auto: false };
  }

  /* A1 没填或格式不对：按今天自动推导（与 sync.py 同规则，日号 >= 26 归起月） */
  var t = new Date(), y = t.getFullYear(), mo = t.getMonth() + 1, d = t.getDate();
  var sY, sM, eY, eM;
  if (d >= 26) {
    sY = y; sM = mo;
    eY = y + (mo === 12 ? 1 : 0);
    eM = (mo % 12) + 1;
  } else {
    eY = y; eM = mo;
    sM = (mo === 1 ? 12 : mo - 1);
    sY = (mo === 1 ? y - 1 : y);
  }
  var auto = sY + '-' + p2(sM) + '-26 ~ ' + eY + '-' + p2(eM) + '-25';

  if (txt) say('A1 内容「' + txt + '」不是合法会计月，已改为按今天自动推导');
  else say('「' + CFG.SHEET_CFG + '」A1 未填会计月，已按今天自动推导');

  if (sh) { try { if (setCell(sh, 'A1', auto)) say('已把推导结果回写 A1：' + auto); } catch (e) {} }

  return { text: auto, sY: sY, sM: sM, sD: 26, eY: eY, eM: eM, eD: 25, auto: true };
}

function realDate(ac, day) {
  var d = +day;
  var y = d >= 26 ? ac.sY : ac.eY;
  var m = d >= 26 ? ac.sM : ac.eM;
  return new Date(y, m - 1, d);
}

function today0() { var n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); }

function fmtDate(d) {
  return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate());
}

/* ---------------------------------------------------------------- 组装 */

function build() {
  var ac = acctMonth();
  say('会计月 ' + ac.text + (ac.auto ? '（自动推导）' : '（取自 A1）'));
  var T = today0();

  /* --- 销售日报 --- */
  var saleRes = readSheet(CFG.SALE.name, CFG.SALE.rows, CFG.SALE.cols);
  var db = dayBlocks(saleRes.rows, '当日销量（吨）');
  say('「' + CFG.SALE.name + '」识别到 ' + db.order.length + ' 个日块');

  var sale = {}, days = [];

  for (var di = 0; di < db.order.length; di++) {
    var day = db.order[di];
    if (realDate(ac, day) > T) continue;
    var blk = db.map[day];
    var st = stations(blk, C_SALE);

    var xe = null;
    for (var i = 0; i < blk.length; i++) { if (s((blk[i] || [])[0]) === '小额配送') { xe = blk[i]; break; } }
    var xiao = {
      t:   xe ? r2(num(xe[6]))  : 0,
      amt: xe ? r2(num(xe[16])) : 0,
      wan: xe ? r3(num(xe[21])) : 0,
      /* 小额行同样有 -10/0/92/95 分列；全 0 视为「未拆柴汽」 */
      diesel: xe ? r3(num(xe[2]) + num(xe[3])) : 0,
      gas:    xe ? r3(num(xe[4]) + num(xe[5])) : 0
    };

    var names = Object.keys(st), any = false;
    var sumT = 0, sumD = 0, sumG = 0, sumL = 0, sumA = 0, sumW = 0;
    for (var k = 0; k < names.length; k++) {
      var v = st[names[k]];
      v.diesel = r3(v.t10 + v.t0);
      v.gas    = r3(v.t92 + v.t95);
      /* 柴汽单价：用「升」与「元」两列的 小计 折出，中间量不进 JSON */
      v.dL = r2(v._dL1 + v._dL0);
      v.gL = r2(v._gL92 + v._gL95);
      v.dA = r2(v._dA1 + v._dA0);
      v.gA = r2(v._gA92 + v._gA95);
      delete v._dL1; delete v._dL0; delete v._gL92; delete v._gL95;
      delete v._dA1; delete v._dA0; delete v._gA92; delete v._gA95;
      if (v.t) any = true;
      sumT += v.t; sumD += v.diesel; sumG += v.gas;
      sumL += v.l; sumA += v.amt;   sumW += v.wan;
    }
    if (!any && xiao.t <= 0) continue;

    sale[day] = {
      stations: st,
      t: r3(sumT), diesel: r3(sumD), gas: r3(sumG),
      l: r2(sumL), amt: r2(sumA), wan: r3(sumW),
      n: names.length, xiao: xiao
    };
    days.push(day);
    say('  ' + day + ' 日：' + names.length + ' 站 / 站级 ' + r2(sumT)
        + ' 吨 + 小额 ' + r2(xiao.t) + ' 吨（块 ' + blk.length + ' 行）');
  }
  if (days.length) {
    var cnt = {}, mm = 0;
    for (var q = 0; q < days.length; q++) {
      var cq = sale[days[q]].n; cnt[cq] = (cnt[cq] || 0) + 1; if (cq > mm) mm = cq;
    }
    if (Object.keys(cnt).length > 1) {
      say('[warn] 各日站点数不一致：' + JSON.stringify(cnt) + '，请核对是否有日块被切断');
    }
  }

  /* --- 渠道表 --- */
  var channels = {};
  for (var ci = 0; ci < CHANNELS.length; ci++) {
    var ch = CHANNELS[ci];
    if (!sheet(ch)) continue;
    var crow = readSheet(ch, CFG.CH.rows, CFG.CH.cols);
    var cdb = dayBlocks(crow.rows, '当日销量（吨）');
    var acc = {};
    for (var cj = 0; cj < cdb.order.length; cj++) {
      var d2 = cdb.order[cj];
      if (realDate(ac, d2) > T) continue;
      var cst = stations(cdb.map[d2], C_AGG);
      var per = {}, has = false, ns = Object.keys(cst);
      for (var cn = 0; cn < ns.length; cn++) {
        var y2 = cst[ns[cn]];
        if (y2.t) has = true;
        per[ns[cn]] = { t: Math.round(y2.t * 10000) / 10000, l: r2(y2.l), amt: r2(y2.amt), wan: r3(y2.wan) };
      }
      if (has) acc[d2] = per;
    }
    if (Object.keys(acc).length) channels[ch] = acc;
  }
  say('渠道表 ' + Object.keys(channels).length + ' 张：' + (Object.keys(channels).join('/') || '无'));

  /* --- 大客户明细（表头是「当日销量（升）」） --- */
  var bigcust = {};
  if (sheet(CFG.CUST.name)) {
    var urow = readSheet(CFG.CUST.name, CFG.CUST.rows, CFG.CUST.cols);
    var udb = dayBlocks(urow.rows, '当日销量（升）');
    var uacc = {};
    for (var ui = 0; ui < udb.order.length; ui++) {
      var d3 = udb.order[ui];
      if (realDate(ac, d3) > T) continue;
      var ublk = udb.map[d3], per2 = {}, road2 = null, anyU = false;
      for (var ur = 0; ur < ublk.length; ur++) {
        var r = ublk[ur];
        if (!r || r.length < 2) continue;
        var a0 = s(r[0]), a1 = s(r[1]);
        if (a0 && AGG.indexOf(a0) < 0 && !/^[\d\-－—]/.test(a0)) road2 = a0;
        if (!isName(a1)) continue;
        var rec = { road: road2 || a0, l: num(r[6]), amt: num(r[11]), c: {} };
        for (var ck = 0; ck < CLIENT_COLS.length; ck++) {
          var base = CLIENT_COLS[ck][0], cname = CLIENT_COLS[ck][1];
          if (base + 9 < r.length) {
            var lv = num(r[base + 4]), av = num(r[base + 9]);
            if (lv || av) { rec.c[cname] = { l: r2(lv), amt: r2(av) }; anyU = true; }
          }
        }
        if (rec.l || rec.amt) anyU = true;
        per2[a1] = rec;
      }
      if (anyU) uacc[d3] = per2;
    }
    if (Object.keys(uacc).length) bigcust[CFG.CUST.name] = uacc;
  }

  days.sort(function (a, b) { return (a < 26) - (b < 26) || (+a - +b); });

  var D = {
    meta: {
      file: '202611销售及库存日报（新).xlsx',
      file_id: 'hLsU8RQ4Q9MPTbvtSaRtxxQdQoC38LbED',
      drive_id: '682951598',
      accounting_month: ac.text,
      days_with_data: days,
      station_count: days.length ? Object.keys(sale[days[0]].stations).length : 0,
      channels: Object.keys(channels),
      extracted_at: fmtDate(T),
      source: 'airscript ' + VER
    },
    sale: sale,
    channels: channels,
    bigcust: bigcust,
    trade: {}
  };
  return D;
}

/* ---------------------------------------------------------------- Base64 */

var B64C = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

function utf8Bytes(str) {
  var out = [];
  for (var i = 0; i < str.length; i++) {
    var c = str.charCodeAt(i);
    if (c < 0x80) out.push(c);
    else if (c < 0x800) { out.push(0xC0 | (c >> 6), 0x80 | (c & 63)); }
    else if (c >= 0xD800 && c <= 0xDBFF) {
      var c2 = str.charCodeAt(++i);
      var cp = 0x10000 + ((c - 0xD800) << 10) + (c2 - 0xDC00);
      out.push(0xF0 | (cp >> 18), 0x80 | ((cp >> 12) & 63), 0x80 | ((cp >> 6) & 63), 0x80 | (cp & 63));
    } else { out.push(0xE0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63)); }
  }
  return out;
}

function b64(str) {
  var b = utf8Bytes(str), o = '', i = 0;
  while (i < b.length) {
    var b0 = b[i++], b1 = i < b.length ? b[i++] : -1, b2 = i < b.length ? b[i++] : -1;
    o += B64C.charAt(b0 >> 2);
    o += B64C.charAt(((b0 & 3) << 4) | (b1 < 0 ? 0 : b1 >> 4));
    o += b1 < 0 ? '=' : B64C.charAt(((b1 & 15) << 2) | (b2 < 0 ? 0 : b2 >> 6));
    o += b2 < 0 ? '=' : B64C.charAt(b2 & 63);
  }
  return o;
}

/* ---------------------------------------------------------------- GitHub */

function apiUrl() {
  return 'https://api.github.com/repos/' + CFG.OWNER + '/' + CFG.REPO + '/contents/' + CFG.PATH;
}

function http(method, url, opts) {
  if (typeof HTTP === 'undefined' || !HTTP) {
    throw new Error('HTTP 对象不存在（请先在编辑器「服务」里添加「网络API」）');
  }
  /* ★ 必须优先走 HTTP.fetch（2026-10-29 实测）
     HTTP.get / HTTP.put 这两个短接口【不吃 opts.headers】—— 传进去也不报错，
     只是 Authorization 悄悄丢了，GitHub 回 401 "Requires authentication"，
     看起来像令牌失效，其实令牌根本没发出去。
     HTTP.fetch(url, {method, headers, body, timeout}) 是 flow_push.js 已验证可用的形式。 */
  var o = opts || {};
  o.method = method;
  if (typeof HTTP.fetch === 'function') return HTTP.fetch(url, o);
  var name = method.toLowerCase();
  if (typeof HTTP[name] === 'function') return HTTP[name](url, o);
  if (typeof HTTP.request === 'function') return HTTP.request(method, url, o);
  throw new Error('HTTP 不支持 ' + method + ' 方法');
}

/* 从 HTTP.fetch 的返回里安全地取出响应正文。
   ★ 2026-10-29 踩过的坑：该返回对象【没有 body 属性】，只有 text() 方法。
     老代码写成 JSON.parse(r.body || r.text || '{}')，r.body 是 undefined 就落到 r.text，
     而 r.text 是【函数对象】，JSON.parse 会隐式 String(函数) 成 "text(){return ...}"，
     直接抛 SyntaxError: Unexpected token 'e', "text(){retu"... is not valid JSON。
   这里按「函数优先 → 字符串 → 其它」三级兜底，两种返回形态都能取到正文。 */
function respText(r) {
  if (!r) return '';
  if (typeof r.text === 'function') { try { return String(r.text()); } catch (e) {} }
  if (typeof r.body === 'string') return r.body;
  if (r.body !== undefined && r.body !== null && typeof r.body !== 'function') return String(r.body);
  return '';
}

function ghGet(tok) {
  var r = http('GET', apiUrl() + '?ref=' + CFG.BRANCH, {
    headers: { 'Authorization': 'Bearer ' + tok, 'Accept': 'application/vnd.github+json',
               'User-Agent': 'kdocs-airscript' },
    timeout: 20000
  });
  var st = r ? Number(r.status) : 0;
  if (st !== 200) {
    /* 不再静默返回 null —— 401 是最常见的坑（令牌没发出去 / 令牌失效），
       静默吞掉会让人以为「远端没文件」，然后去 PUT，再撞一次 401，白跑一轮。 */
    if (st === 401) say('⚠️ GET 401 —— 令牌没生效（先查 HTTP.fetch 是否带上了 Authorization 头）');
    else if (st === 404) say('（GET 404：远端还没有 ' + CFG.PATH + '，按新建处理）');
    else if (st) say('（GET HTTP ' + r.status + '：' + respText(r).substring(0, 140) + '）');
    else say('（GET 无响应）');
    return null;
  }
  var txt = respText(r);
  if (!txt) return null;
  var j = {};
  try { j = JSON.parse(txt) || {}; }
  catch (e) { say('（GET 返回不是 JSON，按无远端文件处理：' + txt.substring(0, 80) + '）'); return null; }
  say('远端现有 data.json：' + (j.size || '?') + ' 字节，sha ' + String(j.sha || '').substring(0, 8));
  return { sha: j.sha, size: j.size, content: j.content };
}

function ghPut(tok, text, sha) {
  var body = {
    message: 'auto: 日报同步 ' + fmtDate(today0()),
    content: b64(text),
    branch : CFG.BRANCH
  };
  if (sha) body.sha = sha;
  return http('PUT', apiUrl(), {
    headers: { 'Authorization': 'Bearer ' + tok, 'Accept': 'application/vnd.github+json',
               'Content-Type': 'application/json', 'User-Agent': 'kdocs-airscript' },
    body: JSON.stringify(body),
    timeout: 30000
  });
}

/* ---------------------------------------------------------------- 主流程 */

function main() {
  say('=== 成品油驾驶舱同步 ' + VER + ' ===');
  var names = sheetNames();
  if (names.length) say('子表（' + names.length + '）：' + names.join(' / '));
  else say('未能枚举子表名（不影响解析，继续尝试按名取表）');

  try {
    var D = build();
    var json = JSON.stringify(D);
    var days = D.meta.days_with_data;
    var last = days.length ? days[days.length - 1] : null;

    say('有数据天数 ' + days.length + '：' + (days.join(',') || '无'));
    if (last) {
      var L = D.sale[last], tot = 0, xs = L.xiao ? L.xiao.t : 0;
      var nm = Object.keys(L.stations);
      for (var i = 0; i < nm.length; i++) tot += L.stations[nm[i]].t;
      say('最新 ' + last + ' 日：站级 ' + r2(tot) + ' 吨 + 小额 ' + r2(xs) + ' 吨 = ' + r2(tot + xs) + ' 吨（' + nm.length + ' 站）');
    }
    say('JSON ' + json.length + ' 字节');

    if (!CFG.PUSH) {
      say('PUSH=false，仅解析未推送。核对上面数字无误后把 CFG.PUSH 改成 true');
      flush();
      return json;
    }

    /* 令牌来源：①「驾驶舱配置」B1  ② CFG.TOKEN（代码里）
       二者取其一即可。没有这张表不算错误 —— 只是少了「C 列运行日志」和「A1 会计月回写」
       这两项便利，控制台日志照常完整。 */
    var cfgSh = sheet(CFG.SHEET_CFG);
    var tok = cfgSh ? cleanTok(getCell(cfgSh, 'B1')) : '';
    var tokFrom = '「' + CFG.SHEET_CFG + '」B1';
    if (!cfgSh) {
      say('（本簿没有「' + CFG.SHEET_CFG + '」表，跳过 B1，改用代码里的 CFG.TOKEN）');
      tokFrom = '代码 CFG.TOKEN';
      tok = cleanTok(CFG.TOKEN);
    } else if (!tok) {
      say('（「' + CFG.SHEET_CFG + '」B1 是空的，改用代码里的 CFG.TOKEN）');
      tokFrom = '代码 CFG.TOKEN';
      tok = cleanTok(CFG.TOKEN);
    }
    if (!tok) {
      throw new Error('找不到 GitHub 令牌 —— 二选一：'
        + '① 新建「' + CFG.SHEET_CFG + '」表，把令牌填进 B1；'
        + '② 把令牌填进代码顶部的 CFG.TOKEN');
    }
    if (!/^(gh[pousr]_|github_pat_)/.test(tok)) {
      // 不打印令牌本身，只提示前 4 位，避免日志外泄
      say('⚠️ ' + tokFrom + ' 的值不像 GitHub 令牌（应以 ghp_ / github_pat_ 开头，实际前 4 位「'
        + tok.substring(0, 4) + '」），仍继续尝试');
    } else {
      say('令牌来源：' + tokFrom);
    }

    var cur = ghGet(tok);
    if (cur && cur.content) {
      var same = String(cur.content).replace(/\s/g, '');
      if (same === b64(json).replace(/\s/g, '')) {
        say('远端内容一致，跳过推送');
        flush();
        return json;
      }
    }

    /* 体积护栏：新内容比远端小超过 15% 就中止推送。
       典型事故是某天只读到 1 个日块（权限变了 / 表结构改了），
       照样推上去会把已有 3 天的数据冲成 1 天，页面直接掉数。
       这种情况下宁可这次不推，让远端保持上一版完整数据，等人工查。 */
    if (cur && cur.size > 1000) {
      var ratio = json.length / cur.size;
      if (ratio < 0.85) {
        say('⚠️ 中止推送：新内容 ' + json.length + ' 字节，只有远端 ' + cur.size
          + ' 字节的 ' + (ratio * 100).toFixed(1) + '%（阈值 85%）');
        say('   → 大概率本次读表不全。远端保持上一版不动，请检查日块/权限后重跑。');
        flush();
        return json;
      }
      say('体积对比：新 ' + json.length + ' 字节 / 远端 ' + cur.size
        + ' 字节（' + (ratio * 100).toFixed(1) + '%）');
    }

    var res = ghPut(tok, json, cur ? cur.sha : null);
    var raw = res && res.status !== undefined && res.status !== null ? res.status : '无响应';
    var st = res ? Number(res.status) : 0;          // 有的实现给字符串 '200'，统一转数字再比
    var ok = st === 200 || st === 201;
    say('推送 HTTP ' + raw + (ok ? ' 成功' : ' 失败'));
    if (!ok) say('响应 ' + respText(res).substring(0, 180));
    flush();
    return json;
  } catch (e) {
    say('异常：' + (e && e.message ? e.message : e));
    flush();
    throw e;
  }
}

main();
