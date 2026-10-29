/* ==========================================================================
 * 成品油日报 → GitHub _incoming/data.json   v2
 * 运行位置：金山文档「202611销售及库存日报（新).xlsx」内的 AirScript
 * 触发方式：定时任务（8:30 / 9:00 / 9:30 / 10:00 / 11:30 各挂一条，跑同一个 main）
 *
 * v2 变更（针对首跑报错「驾驶舱配置!A1 需写会计月」）
 *  1) 会计月改为「A1 有就认 A1，没有就按今天自动推导」——不再强制手工维护；
 *     推导结果会回写 A1，让你在表里看得见。
 *  2) 日志同时打到 console 与「驾驶舱配置」C 列；没有该表也能在控制台看全部结果。
 *  3) 取表 / 取格 / 取区域 全部加多重兼容（Range·Value2·Value·getValues）。
 *  4) 启动先打印所有子表名 —— 取表失败时一眼看出是名字不对还是取法不对。
 *
 * 设计原则
 *  1) 只读源表，绝不写回任何业务单元格；唯一写入是「驾驶舱配置」的状态格。
 *  2) 幂等：先 GET 远端 data.json，内容相同直接退出，不产生提交。
 *     → 「补发」不是独立机制，而是幂等的副产品：上次失败则远端仍是旧内容，下一档自动重推。
 *  3) 产出结构与本地 extract.py 的 data.json 完全同构，下游 build_cockpit2.py 零改动。
 *  4) 全零日块 = 未填报，跳过（不可补零，会把日均拉低）。
 * ========================================================================== */

var VER = 'v2';

var CFG = {
  OWNER : 'lvritiann',
  REPO  : 'youpin-sales-cockpit',
  BRANCH: 'main',
  PATH  : '_incoming/data.json',

  SHEET_CFG : '驾驶舱配置',   // 可选表：A1=会计月(留空自动推导), B1=GitHub令牌, C1..=运行日志
  PUSH      : false,         // ★ 出厂 false：只解析。核对无误后改 true 才会真推

  SALE : { name: '销售日报', rows: 1601, cols: 28 },
  CH   : { rows: 1537, cols: 22 },
  CUST : { name: '大客户明细', rows: 1547, cols: 162 }
};

var C_SALE = { t10: 2, t0: 3, t92: 4, t95: 5, t: 6, l: 11, amt: 16, wan: 21 };
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

/* Value2 可能返回 0 基或 1 基二维数组，统一压成 0 基并补齐列宽 */
function grid(v, cols) {
  if (!v || !v.length) return [];
  var one = (v[0] === undefined);
  var out = [];
  for (var i = one ? 1 : 0; i < v.length; i++) {
    var src = v[i];
    if (!src) { out.push(new Array(cols)); continue; }
    var inner = (src[0] === undefined && src.length > 1) ? 1 : 0;
    var row = [];
    for (var j = inner; j < src.length; j++) row.push(src[j]);
    while (row.length < cols) row.push('');
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
  return { rows: grid(raw, cols), sh: sh };
}

/* ---------------------------------------------------------------- 日块切分 */

function blocksOf(rows, key) {
  var starts = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i] || [];
    for (var j = 0; j < r.length; j++) {
      if (s(r[j]).indexOf(key) >= 0) { starts.push(i); break; }
    }
  }
  var out = [];
  for (var k = 0; k < starts.length; k++) {
    out.push({
      label: s((rows[starts[k]] || [])[0]),
      rows : rows.slice(starts[k], k + 1 < starts.length ? starts[k + 1] : rows.length)
    });
  }
  return out;
}

function dayBlocks(rows, key) {
  var bl = blocksOf(rows, key), res = {}, order = [];
  var idx = 0;
  for (var k = 0; k < bl.length; k++) {
    if (bl[k].label.indexOf('-') >= 0) continue;      // 会计月累计块，丢弃
    var lab = bl[k].label, day = null;
    if (/^\d{1,2}$/.test(lab) && +lab >= 1 && +lab <= 31) day = String(+lab);
    if (!day) day = SEQ[idx] || null;
    idx++;
    if (!day) continue;
    res[day] = bl[k].rows;
    if (order.indexOf(day) < 0) order.push(day);
  }
  return { map: res, order: order };
}

function stations(blk, cols) {
  var out = {}, road = null;
  for (var i = 0; i < blk.length; i++) {
    var r = blk[i];
    if (!r || r.length < 2) continue;
    var c0 = s(r[0]), c1 = s(r[1]);
    if (c0 && AGG.indexOf(c0) < 0 && !/^\d/.test(c0)) road = c0;
    if (!c1 || AGG.indexOf(c1) >= 0 || c1.indexOf('当日') >= 0 || /^\d/.test(c1)) continue;
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
      wan: xe ? r3(num(xe[21])) : 0
    };

    var names = Object.keys(st), any = false;
    var sumT = 0, sumD = 0, sumG = 0, sumL = 0, sumA = 0, sumW = 0;
    for (var k = 0; k < names.length; k++) {
      var v = st[names[k]];
      v.diesel = r3(v.t10 + v.t0);
      v.gas    = r3(v.t92 + v.t95);
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
        if (a0 && AGG.indexOf(a0) < 0 && !/^\d/.test(a0)) road2 = a0;
        if (!a1 || AGG.indexOf(a1) >= 0 || /^\d/.test(a1)) continue;
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
  var name = method.toLowerCase();
  if (typeof HTTP[name] === 'function') return HTTP[name](url, opts);
  if (typeof HTTP.request === 'function') return HTTP.request(method, url, opts);
  throw new Error('HTTP 不支持 ' + method + ' 方法');
}

function ghGet(tok) {
  var r = http('GET', apiUrl() + '?ref=' + CFG.BRANCH, {
    headers: { 'Authorization': 'Bearer ' + tok, 'Accept': 'application/vnd.github+json',
               'User-Agent': 'kdocs-airscript' },
    timeout: 20000
  });
  if (!r || r.status !== 200) return null;
  var j = JSON.parse(r.body || r.text || '{}');
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

    var cfgSh = sheet(CFG.SHEET_CFG);
    var tok = cfgSh ? s(getCell(cfgSh, 'B1')) : '';
    if (!tok) throw new Error('「' + CFG.SHEET_CFG + '」B1 缺少 GitHub 令牌');

    var cur = ghGet(tok);
    if (cur && cur.content) {
      var same = String(cur.content).replace(/\s/g, '');
      if (same === b64(json).replace(/\s/g, '')) {
        say('远端内容一致，跳过推送');
        flush();
        return json;
      }
    }

    var res = ghPut(tok, json, cur ? cur.sha : null);
    var st = res ? res.status : '无响应';
    say('推送 HTTP ' + st + (st === 200 || st === 201 ? ' 成功' : ' 失败'));
    if (!(st === 200 || st === 201)) say('响应 ' + s(res && (res.body || res.text)).substring(0, 180));
    flush();
    return json;
  } catch (e) {
    say('异常：' + (e && e.message ? e.message : e));
    flush();
    throw e;
  }
}

main();
