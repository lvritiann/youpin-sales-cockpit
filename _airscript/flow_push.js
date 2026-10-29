/* ============================================================================
 * 成品油驾驶舱「前一日车流」推送 —— 独立 AirScript
 * ----------------------------------------------------------------------------
 * 运行位置：金山文档「非油销售汇总驾驶舱V1.xlsx」→ 效率 → 高级开发 → 脚本编辑器
 * 触发方式：定时任务，每日 09:40（错开既有 08:30 / 09:00 的驾驶舱刷新）
 *
 * 做什么
 *   逐站读 15 个非油 V1 报表的「今日填报」，取最近一个有车次的会计日，
 *   把每站的 柴油车次 / 柴油升数 / 汽油车次 / 汽油升数 汇总成一对一份，
 *   组装成 _incoming/flow.json 推到 GitHub 仓库 lvritiann/youpin-sales-cockpit。
 *   推上去后 rebuild-cockpit 工作流会自动重建页面，驾驶舱窄条卡的「车流 / 单加」即更新。
 *
 * ★ 为什么是独立脚本，而不是改 cockpit_refresh_full_v3_6
 *   那份主脚本 1700+ 行、每天在正常出数，任何改动都可能连带把「按日汇总 / 图表数据源」
 *   弄坏。本脚本只做一件事，出问题直接删掉即可，主链路零风险。
 *   它读的文件与主脚本完全相同，所以数据口径天然一致。
 *
 * 前置条件（缺一不可）
 *   ① 脚本编辑器左侧「服务」→ 添加「云文档API」（KSDrive.openFile 依赖）
 *   ② 脚本编辑器左侧「服务」→ 添加「网络API」（HTTP.fetch 依赖）
 *   ③ 15 站 V1 报表已分享给本账号（≥ 可阅读；建议可编辑）
 *   ④ CFG.TOKEN 填入 GitHub PAT（需 repo 权限，与主脚本用的是同一个即可）
 *
 * 首次上线：CFG.PUSH 保持 false 先干跑一次，看控制台打印的数字对不对，
 *           确认无误后再改 true 并保存。
 * ========================================================================== */

/* ---------------------------------------------------------------- 配置 */
var CFG = {
  PUSH  : false,                          // ★ 干跑 false → 核对无误后改 true
  OWNER : 'lvritiann',
  REPO  : 'youpin-sales-cockpit',
  PATH  : '_incoming/flow.json',
  BRANCH: 'main',
  TOKEN : '',                             // ★ 填 GitHub PAT（ghp_ 开头，repo 权限）
  MIN_COVER: 12,                          // 至少这么多站有车次才认这一天，防半填状态被推
  XIAO  : '小额配送'                       // 公司级虚拟实体，无车流，不参与
};

/* 与主脚本 STATIONS 一一对应（服务区 = 一对，站内两行南/东 + 北/西按日相加） */
var STATIONS = [
  ['秦岭',     'https://www.kdocs.cn/l/cehkKd3fI0Ty'],
  ['宁陕',     'https://www.kdocs.cn/l/cmvi9xbaRq9n'],
  ['洋县',     'https://www.kdocs.cn/l/cmgfqvzCHw6q'],
  ['汉中',     'https://www.kdocs.cn/l/crrc5h7pc230'],
  ['富平',     'https://www.kdocs.cn/l/csNF5NlRQiu8'],
  ['韩城',     'https://www.kdocs.cn/l/cbb6o9PPzllF'],
  ['富县',     'https://www.kdocs.cn/l/cqt6hlBkoxQ4'],
  ['南沙',     'https://www.kdocs.cn/l/ctqpCVv1OXK9'],
  ['华山',     'https://www.kdocs.cn/l/ccREyNEiVWlM'],
  ['白河',     'https://www.kdocs.cn/l/caDuqWdGlgQl'],
  ['旬阳',     'https://www.kdocs.cn/l/co9cZ2Ud8tgQ'],
  ['略阳',     'https://www.kdocs.cn/l/cbx5Jwnou1u8'],
  ['玉华宫',   'https://www.kdocs.cn/l/cflD3dlLedoK'],
  ['照金',     'https://www.kdocs.cn/l/cetm8edGFWMl'],
  ['天汉水城', 'https://www.kdocs.cn/l/cgrklENA3ssO']
];

var YEAR_START = '2025-12-26';            // 会计年首日 → 今日填报第 2 行（0 基 1）
var LOG = [];

/* ---------------------------------------------------------------- 工具 */
function say(x) { var t = String(x); LOG.push(t); try { console.log(t); } catch (e) {} }

function num(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var t = String(v).replace(/,/g, '').replace(/\s/g, '').trim();
  if (t === '' || t === '-' || t === '—' || t === '/') return 0;
  var f = parseFloat(t);
  return isNaN(f) ? 0 : f;
}

function pad2(n) { return (n < 10 ? '0' : '') + n; }

/* 日期归一：AirScript 的 Range.Value2 读日期返回的是 Excel 序列号（不是 Date），
   不转就会拿 46323 这种数字当 key，导致「读到了但 0 站匹配」。 */
function normDate(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    return v.getUTCFullYear() + '-' + pad2(v.getUTCMonth() + 1) + '-' + pad2(v.getUTCDate());
  }
  if (typeof v === 'number') {
    var s = Math.floor(v);
    if (s >= 1 && s <= 73000) {
      var d = new Date((s - 25569) * 86400000);
      return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
    }
    return String(v);
  }
  var str = String(v).trim().replace(/\//g, '-');
  var m = str.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + pad2(parseInt(m[2], 10)) + '-' + pad2(parseInt(m[3], 10));
  return str.slice(0, 10);
}

/* 会计日 → 「今日填报」里该日 南/东 所在的行号（1 基，Excel 显示行号）
   ★ 不用 new Date('2026-10-28T00:00:00Z') —— 部分 JS 引擎不认 ISO 字符串，
     会返回 Invalid Date 且不报错，结果 rowOfDate 全是 NaN。改用 Date.UTC 数值构造。 */
function rowOfDate(ds) {
  var p = String(ds).split('-');
  var q = YEAR_START.split('-');
  var a = Date.UTC(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
  var b = Date.UTC(parseInt(q[0], 10), parseInt(q[1], 10) - 1, parseInt(q[2], 10));
  if (isNaN(a) || isNaN(b)) return 0;
  return 1 + 2 * Math.round((a - b) / 86400000);
}

function r2(x) { return Math.round(x * 100) / 100; }
function r1(x) { return Math.round(x * 10) / 10; }

/* UTF-8 → Base64（AirScript 没有 btoa/TextEncoder，中文必须自己转） */
function utf8ToBase64(str) {
  var bytes = [], i, c, c2, cp;
  for (i = 0; i < str.length; i++) {
    c = str.charCodeAt(i);
    if (c < 0x80) bytes.push(c);
    else if (c < 0x800) bytes.push(0xc0 | (c >> 6), 0x80 | (c & 0x3f));
    else if (c >= 0xd800 && c <= 0xdbff) {
      c2 = str.charCodeAt(++i);
      cp = 0x10000 + ((c & 0x3ff) << 10) + (c2 & 0x3ff);
      bytes.push(0xf0 | (cp >> 18), 0x80 | ((cp >> 12) & 0x3f),
                 0x80 | ((cp >> 6) & 0x3f), 0x80 | (cp & 0x3f));
    } else bytes.push(0xe0 | (c >> 12), 0x80 | ((c >> 6) & 0x3f), 0x80 | (c & 0x3f));
  }
  var tbl = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/', out = '';
  for (var j = 0; j < bytes.length; j += 3) {
    var e1 = bytes[j], e2 = bytes[j + 1], e3 = bytes[j + 2];
    var n = (e1 << 16) | ((e2 || 0) << 8) | (e3 || 0);
    out += tbl[(n >> 18) & 63] + tbl[(n >> 12) & 63]
         + (e2 != null ? tbl[(n >> 6) & 63] : '=') + (e3 != null ? tbl[n & 63] : '=');
  }
  return out;
}

/* ------------------------------------------------- 读一个站的逐日车流 */
function readStationFlow(url) {
  var f = null, out = null;
  try {
    f = KSDrive.openFile(url);
    var data = f.Application.Sheets('今日填报').Range('A1:T731').Value2;
    f.close(); f = null;
    if (!data || !data.length) return null;

    var h = data[0], ix = {}, j;
    for (j = 0; j < h.length; j++) ix[String(h[j]).trim()] = j;
    var need = ['日期', '柴油单数', '汽油单数', '柴油升数', '汽油升数'];
    var miss = [];
    for (j = 0; j < need.length; j++) if (!(need[j] in ix)) miss.push(need[j]);
    if (miss.length) { say('  ⚠️ 缺列 ' + miss.join('/') + ' @ ' + url); return null; }

    out = {};
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var ds = normDate(row[ix['日期']]);
      if (!ds || ds === '0000-00-00') continue;
      var dv = num(row[ix['柴油单数']]), gv = num(row[ix['汽油单数']]);
      var dl = num(row[ix['柴油升数']]), gl = num(row[ix['汽油升数']]);
      if (dv + gv <= 0) continue;                     // 未填报，跳过（绝不补零）
      if (!out[ds]) out[ds] = { dv: 0, gv: 0, dl: 0, gl: 0 };
      out[ds].dv += dv; out[ds].gv += gv;
      out[ds].dl += dl; out[ds].gl += gl;
    }
  } catch (e) {
    say('  ⚠️ 读取失败 ' + url + ' : ' + e);
    if (f) { try { f.close(); } catch (_) {} }
    return null;
  }
  return out;
}

/* --------------------------------- 从「单数统计」取公司级当日值做对账 */
function readCountsCheck(asof) {
  try {
    var sh = Application.Sheets('单数统计');
    if (!sh) return null;
    var v = sh.Range('A1:K500').Value2;
    if (!v || !v.length) return null;
    var hdrRow = -1, i, j;
    for (i = 0; i < v.length; i++) {
      if (String(v[i][0]).trim() === '日期') { hdrRow = i; break; }
    }
    if (hdrRow < 0) return null;
    var ix = {}, h = v[hdrRow];
    for (j = 0; j < h.length; j++) ix[String(h[j]).trim()] = j;
    for (i = hdrRow + 1; i < v.length; i++) {
      if (normDate(v[i][0]) !== asof) continue;
      return {
        gasVeh: num(v[i][ix['汽油车次']]),
        dieVeh: num(v[i][ix['柴油车次']]),
        veh   : num(v[i][ix['总车次']])
      };
    }
  } catch (e) { say('  （单数统计对账跳过：' + e + '）'); }
  return null;
}

/* ---------------------------------------------------------------- 主流程 */
function main() {
  say('===== 前一日车流推送 START ' + new Date() + ' =====');

  /* 1) 逐站读取 */
  var flow = {}, okN = 0, i;
  for (i = 0; i < STATIONS.length; i++) {
    var nm = STATIONS[i][0], url = STATIONS[i][1];
    var d = readStationFlow(url);
    if (d && Object.keys(d).length) { flow[nm] = d; okN++; }
    else say('  ⚠️ ' + nm + ' 无可用车流数据');
  }
  say('读取完成：' + okN + '/' + STATIONS.length + ' 站有数据');
  if (okN < CFG.MIN_COVER) {
    say('❌ 有车次的站只有 ' + okN + ' 个（阈值 ' + CFG.MIN_COVER + '），视为数据未就绪，本次不推送');
    return;
  }

  /* 2) 找最近一个「够站数」的会计日 */
  var cover = {};                                   // 日期 → 有车次的站数
  Object.keys(flow).forEach(function (nm) {
    Object.keys(flow[nm]).forEach(function (ds) {
      cover[ds] = (cover[ds] || 0) + 1;
    });
  });
  var asof = '';
  Object.keys(cover).forEach(function (ds) {
    if (cover[ds] >= CFG.MIN_COVER && ds > asof) asof = ds;
  });
  if (!asof) { say('❌ 没有任何一天达到 ' + CFG.MIN_COVER + ' 站覆盖，放弃'); return; }
  say('选定会计日 asof=' + asof + '（覆盖 ' + cover[asof] + ' 站）');

  /* 3) 组装 payload —— 结构与 _kdocs/_flow.json 完全同构，下游 zero 改动 */
  var st = {}, tv = 0, tg = 0, td = 0;
  var sumGL = 0, sumDL = 0;                          // 分油品升数合计，用来算公司级单加
  Object.keys(flow).forEach(function (nm) {
    var rec = flow[nm][asof];
    if (!rec) return;
    var iv = Math.round(rec.dv + rec.gv);
    if (iv <= 0) return;
    st[nm] = {
      iv : iv,
      ig : Math.round(rec.gv),
      id : Math.round(rec.dv),
      isp: rec.gv > 0 ? r2(rec.gl / rec.gv) : null,
      idp: rec.dv > 0 ? r2(rec.dl / rec.dv) : null,
      il : r1(rec.dl + rec.gl),
      gl : r1(rec.gl),
      dl : r1(rec.dl)
    };
    tv += st[nm].iv; tg += st[nm].ig; td += st[nm].id;
    sumGL += rec.gl; sumDL += rec.dl;
  });

  var r0 = rowOfDate(asof);
  var payload = {
    asof : asof,
    basis: '前一日实报',
    note : '车流/单加逐站取各站非油 V1 报表「今日填报」' + asof
         + ' 两行（南/东 + 北/西）实报值，非日均分摊、非估值（AirScript 自动生成）',
    source: 'airscript flow_push.js',
    rows : [r0, r0 + 1],
    check: {
      veh   : tv,
      gasVeh: tg,
      dieVeh: td,
      /* 公司级单加 = 该油品总升数 ÷ 该油品总车次（不是各站单加的算术平均） */
      gasPer: tg > 0 ? r2(sumGL / tg) : null,
      diePer: td > 0 ? r2(sumDL / td) : null
    },
    pairs   : st,
    stations: st
  };

  /* 4) 对账：15 对加总 vs 驾驶舱「单数统计」公司级当日值 */
  var ck = readCountsCheck(asof);
  var line = '15 对加总：总车次 ' + tv + '（汽 ' + tg + ' / 柴 ' + td + '）';
  if (ck) {
    var d1 = ck.veh - tv, d2 = ck.gasVeh - tg, d3 = ck.dieVeh - td;
    var okAll = (Math.abs(d1) <= 2 && Math.abs(d2) <= 2 && Math.abs(d3) <= 2);
    say(line);
    say('单数统计：总车次 ' + ck.veh + '（汽 ' + ck.gasVeh + ' / 柴 ' + ck.dieVeh + '）');
    say('差值：总 ' + d1 + ' / 汽 ' + d2 + ' / 柴 ' + d3 + (okAll ? '  ✅ 吻合' : '  ⚠️ 有偏差，请人工确认'));
    payload.check.company = ck;
    payload.check.delta = { veh: d1, gas: d2, diesel: d3 };
  } else {
    say(line);
    say('（未找到「单数统计」中 ' + asof + ' 的行，跳过公司级对账 —— 可能是该表还没刷到今天）');
  }

  say('站数 ' + Object.keys(st).length + ' · payload ' + JSON.stringify(payload).length + ' 字节');

  /* 5) 干跑就直接结束 */
  if (!CFG.PUSH) {
    say('【干跑】CFG.PUSH=false，不推送。核对上面数字无误后把 CFG.PUSH 改成 true。');
    say('样例（前 3 站）：' + JSON.stringify(
      Object.keys(st).slice(0, 3).map(function (k) { return k + '=' + JSON.stringify(st[k]); })));
    say('===== END（未推送）=====');
    return;
  }

  /* 6) 推送 */
  if (!CFG.TOKEN || String(CFG.TOKEN).indexOf('ghp_') !== 0) {
    say('❌ CFG.TOKEN 未填或格式不对（应以 ghp_ 开头）。到 GitHub → Settings → Developer settings'
      + ' → Personal access tokens 生成一个带 repo 权限的令牌，填进 CFG.TOKEN。');
    return;
  }

  var api = 'https://api.github.com/repos/' + CFG.OWNER + '/' + CFG.REPO + '/contents/' + CFG.PATH;
  var headers = {
    'Authorization': 'Bearer ' + CFG.TOKEN,
    'Accept': 'application/vnd.github+json',
    'Content-Type': 'application/json'
  };

  var sha = null;
  var getR = HTTP.fetch(api + '?ref=' + CFG.BRANCH, { method: 'GET', headers: headers, timeout: 10000 });
  if (getR && ((getR.status >= 200 && getR.status < 300) || getR.ok)) {
    try { sha = JSON.parse(getR.text()).sha; } catch (e) { sha = null; }
  } else {
    say('[GitHub] GET 未取到 sha（status=' + (getR && getR.status) + '），按新建处理');
  }

  var body = {
    message: 'chore(flow): 车流 ' + asof + ' 自动推送（AirScript）',
    content: utf8ToBase64(JSON.stringify(payload, null, 1)),
    branch : CFG.BRANCH
  };
  if (sha) body.sha = sha;

  var putR = HTTP.fetch(api, { method: 'PUT', headers: headers, body: JSON.stringify(body), timeout: 15000 });
  if (putR && ((putR.status >= 200 && putR.status < 300) || putR.ok)) {
    say('✅ 已推送 _incoming/flow.json（asof=' + asof + '，' + Object.keys(st).length + ' 站）'
      + ' → rebuild-cockpit 会自动重建页面');
  } else {
    say('❌ 推送失败 status=' + (putR && putR.status) + ' ' + (putR && putR.text && putR.text()));
  }
  say('===== END =====');
}
