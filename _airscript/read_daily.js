/* ==========================================================================
 * 成品油日报 → GitHub _incoming/data.json   v2.9
 * 运行位置：金山文档「202611销售及库存日报（新).xlsx」内的 AirScript
 * 触发方式：定时任务（8:30 / 9:00 / 9:30 / 10:00 / 11:30 各挂一条，跑同一个 main）
 *
 * v2.9 变更（2026-10-29，主表瘦身：归档迁到独立的中转簿）
 *  1) ★ 本簿不再写「历史归档」表 —— CFG.ARCH_SHEET 置空，archiveToSheet 首句返回。
 *     那 1001 行站级 + 34 行月汇总从填报簿搬到了独立工作簿
 *     「成品油经营数据中转簿」，由挂在那本簿上的 transit_sync.js 负责落表。
 *  2) 为什么搬：填报簿 13.4 MB / 62 张子表，历史却和每日填报挤在一起；
 *     而驾驶舱要的「任意区间 × 任意站」切片，光有月汇总 + 月站级根本做不出来，
 *     中转簿多了 30,502 行 date × station 日粒度矩阵。
 *  3) 为什么脚本不跟着搬：本脚本要读三张大表（销售日报 1601×28、渠道 1537×22、
 *     大客户 1547×162），读源留在本簿是本地读、零跨文档成本；
 *     跨文档「写」在 WPS 上没验证过，不做没把握的事。
 *     transit_sync.js 反着来：跨文档「读」销售日报（flow_push.js 已验证可行），
 *     写自己簿的表（本地写）。
 *  4) 本脚本职责不变：解析全量 → 推 data.json。渠道 / 大客户 / 车流仍只看本簿。
 *  5) 大客户槽位改为表头驱动（clientSlots）：从硬编码 9 个客户改为扫 15 个槽位，
 *     表头空着的统一归入「其他」。详见下方 CLIENT_* 注释。
 *
 * v2.8 变更（2026-10-29，新增「历史归档」中转表落表）
 *  1) ★ 每次解析成功后，把「当前会计月汇总 + 站级明细」写入本簿的「历史归档」表。
 *     该表与「销售日报」物理隔离 —— 每月 26 号清表不会再连带抹掉历史。
 *     版式：A~L 月汇总（一月一行，幂等覆盖），N~V 站级明细（按「会计月+站名」定位）。
 *  2) 与 PUSH 解耦：不推 GitHub 也照样落表；落表失败不阻断推送。
 *  3) 没有该表时只提示一句，不报错、不影响原有流程。
 *
 * v2.7 变更（2026-10-29，修「每月 26 号必然误拦」）
 *  ★ 体积护栏在新会计月起步时会连拦 8 天：26 号清表后只剩 1 天数据，
 *    UTF-8 体积约为月末的 1/30，被判成「读表不全」，而页面会一直停在旧会计月。
 *    现在加了「会计月起步放行」：本次最早日块 = 会计月首日时跳过体积护栏。
 *
 * v2.6 变更（2026-10-29，修体积护栏误伤）
 *  1) ★ 护栏原先拿 json.length（JS 字符数）去比 GitHub 返回的 size（字节数）。
 *     中文站名一个占 3 字节，两者差约 11%：实测 70,555 字符 ↔ 78,687 字节。
 *     结果一次正常的「长小数截到 4 位」（体积本就该降）被算成 84.7%，
 *     误判成读表不全而拦下 —— 数据本身完全正常（3 天 / 30 站）。
 *     新增 utf8Len() 换算，护栏改比 UTF-8 字节。
 *  2) 新增「0 天保护」：一个日块都没读到就中止推送，绝不把远端清空。
 *
 * v2.5 变更（2026-10-29，源字段精度归一）
 *  1) ★ 站级 t10 / t0 / t92 / t95 / t / l / amt / wan 统一取 4 位小数。
 *     源表吨位列是「升 × 密度」折出来的长小数（实测 8.8945041），
 *     不取整会带浮点噪声（9.196847100000001、11123.250000000002）进 JSON；
 *     砍到 2 位又会失真 —— 略阳汽油 0.771→0.76 吨，平均单价从 11,716 跳到 11,886（差 1.4%）。
 *     取 4 位（0.1 公斤）两头兼顾。
 *  2) 派生字段 diesel / gas / dL / gL / dA / gA 仍按 r3 / r2，保持与旧版一致。
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

var VER = 'v2.9';

var CFG = {
  OWNER : 'lvritiann',
  REPO  : 'youpin-sales-cockpit',
  BRANCH: 'main',
  PATH  : '_incoming/data.json',

  SHEET_CFG : '驾驶舱配置',   // 可选表：A1=会计月(留空自动推导), B1=GitHub令牌, C1..=运行日志
  ARCH_SHEET: '',            // ★ 2026-10-29 起停用（留空 = 跳过归档落表）
                             //   归档已迁到独立工作簿「成品油经营数据中转簿」，
                             //   由挂在那本簿上的 transit_sync.js 负责 —— 它跨文档读本簿的
                             //   「销售日报」，把日粒度/月汇总/站级明细写进中转簿自己的表。
                             //
                             //   留空的收益：填报簿不再累积 1001 行站级 + 34 行月汇总。
                             //   每月清空日报表时，历史也不再和填报数据挤在同一本簿里。
                             //
                             //   真要临时回到「写在本簿」：把这里改回 '历史归档'，
                             //   并在本簿手工建一张同名空工作表即可（archiveToSheet 会自己补表头）。
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

/* ★ 客户块布局（0 基）：每块 10 列（5 列升 + 5 列金额），步长 10，首块是「XX合计」。
   2026-10-29 之前的写法是硬编码 9 个客户 + 固定列位：
       [[12,'中融致远'],[22,'锦航石化'],…,[102,'运达通'],[142,'通四海']]
   但源表实际有 15 个槽位（C/M/W/AG/AQ/BA/BK/BU/CE/CO/CY/DI/DS/EC/EM），
   其中 CE(82)、CO(92)、DI(112)、DS(122)、EC(132) 五块代码从来没读过。
   实测 CO 与 DI 天天都有真实量（占大客户总量约 0.8%）——
   这些量在「大客户构成」里既看不到、也不会报错，属于静默丢失。
   现在改为扫第 1 行表头：有名字用名字，没名字但确有量的记「未署名客户 <列号>」，
   并逐行拿各槽位之和去对「XX合计」列，对不上立刻告警。 */
var CLIENT_STEP  = 10;
var CLIENT_AGGC  = 2;          // 「XX合计」块的起始列（0 基）

function clientSlots(hdr, width) {
  var out = [], W = Math.max(hdr ? hdr.length : 0, width || 0);
  for (var c = CLIENT_AGGC + CLIENT_STEP; c + 9 < W; c += CLIENT_STEP) {
    out.push([c, (hdr && hdr[c] !== undefined && hdr[c] !== null) ? s(hdr[c]) : '']);
  }
  return out;
}

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
function r4(v) { return Math.round(v * 10000) / 10000; }

/* UTF-8 字节长度。
   ★ 不能拿 String.length 当字节数：JS 的 length 是 UTF-16 码元数，
     一个中文站名在 UTF-8 里占 3 字节，实测 70,555 字符 ↔ 78,687 字节（差 11%）。
     体积护栏要和 GitHub 返回的 size（字节）比，就必须先换算。 */
function utf8Len(str) {
  var n = 0, i, c;
  for (i = 0; i < str.length; i++) {
    c = str.charCodeAt(i);
    if (c < 0x80) n += 1;
    else if (c < 0x800) n += 2;
    else if (c >= 0xD800 && c <= 0xDBFF) { n += 4; i++; }   // 代理对（表情等）
    else n += 3;
  }
  return n;
}
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
      /* ★ 源字段精度归一（v2.5）
         源表吨位列本身是「升 × 密度」折出来的长小数（实测 8.8945041），两种极端都不可取：
         · 完全不取整 → JSON 带浮点噪声（9.196847100000001、11123.250000000002），
           体积虚增、逐位核对时刺眼；
         · 砍到 2 位 → 小站失真：略阳汽油 0.771 吨被砍成 0.76，平均单价从 11,716 跳到 11,886
           （差 1.4%）。基数越小，末位误差放大得越厉害。
         取 4 位（0.1 公斤）是折中：噪声清干净，失真压到 0.1% 以内。 */
      v.t10 = r4(v.t10); v.t0 = r4(v.t0); v.t92 = r4(v.t92); v.t95 = r4(v.t95);
      v.t   = r4(v.t);   v.l  = r4(v.l);  v.amt = r4(v.amt); v.wan = r4(v.wan);
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
    /* 客户名只在第 1 行（31 个日块共用这一行表头），不在任何日块内部，
       所以必须单独取 rows[0]，不能指望在日块里扫到名字。 */
    var uSlots = clientSlots(urow.rows[0] || [], CFG.CUST.cols);
    var uNamed = 0, uBlank = 0;
    for (var si = 0; si < uSlots.length; si++) { if (uSlots[si][1]) uNamed++; else uBlank++; }
    var udb = dayBlocks(urow.rows, '当日销量（升）');
    var uacc = {}, uGap = 0, uTot = 0, uBlankL = 0, uUsed = [], uBlankCols = [];
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
        var sumL = 0;
        for (var ck = 0; ck < uSlots.length; ck++) {
          var base = uSlots[ck][0], cnm = uSlots[ck][1];
          if (base + 9 >= r.length) continue;
          var lv = num(r[base + 4]), av = num(r[base + 9]);
          if (!lv && !av) continue;
          /* 表头空着 = 该客户没登记名字，统一归入「其他」；
             具体是哪一列仍然打日志，方便回源表补名字。 */
          var ckey = cnm || '其他';
          if (!cnm) { uBlankL += lv; if (uBlankCols.indexOf(colName(base + 1)) < 0) uBlankCols.push(colName(base + 1)); }
          if (uUsed.indexOf(ckey) < 0) uUsed.push(ckey);
          var o = rec.c[ckey];
          if (!o) { o = rec.c[ckey] = { l: 0, amt: 0 }; }
          o.l += lv; o.amt += av;
          sumL += lv;
        }
        for (var kk in rec.c) {
          if (!rec.c.hasOwnProperty(kk)) continue;
          rec.c[kk].l = r2(rec.c[kk].l); rec.c[kk].amt = r2(rec.c[kk].amt);
        }
        /* 逐行对账：各槽位之和必须等于「XX合计」块的升数，否则是列位漂移 */
        var aggL = num(r[CLIENT_AGGC + 4]);
        if (aggL > 0) { uTot += aggL; uGap += (aggL - sumL); }
        if (rec.l || rec.amt) anyU = true;
        per2[a1] = rec;
      }
      if (anyU) uacc[d3] = per2;
    }
    if (Object.keys(uacc).length) bigcust[CFG.CUST.name] = uacc;

    say('大客户槽位 ' + uSlots.length + ' 个（具名 ' + uNamed + ' / 空表头 ' + uBlank + '）'
        + '，有量客户 ' + uUsed.length + ' 个');
    if (uBlankL > 0) {
      say('空表头列（' + uBlankCols.join('、') + '）贡献 ' + r2(uBlankL) + ' 升'
          + (uTot > 0 ? '（占大客户 ' + (uBlankL / uTot * 100).toFixed(2) + '%）' : '')
          + ' → 已统一归入「其他」；要单列成客户，去源表第 1 行把这些格子填上名字');
    }
    if (uTot > 0 && Math.abs(uGap) / uTot > 0.005) {
      say('⚠️ 大客户对账不平：合计 ' + r2(uTot) + ' 升，各槽位之和差 ' + r2(uGap) + ' 升'
          + '（可能是客户块步长变了，请核对第 1 行表头）');
    }
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


/* ---------------------------------------------------------------- 归档落表 */
/* 把本次解析结果写入「历史归档」中转表 —— 解决「每月 26 号清空日报表导致数据断档」。
   背景：日报表每会计月清空重写，data.json 只保留当前会计月；
        一旦月底没来得及另存 xlsx、也没重建 _archive，那个月的数据就断了。
   做法：每次解析成功就把「会计月汇总 + 站级明细」落到一张**独立的表**里，
        与日报表物理隔离 —— 清表不影响它，逐月累积，可追溯。
   幂等：同一会计月重复跑只覆盖该月那几行，不会产生重复记录。
   版式：A~L 月汇总（一行一个月），N~V 站级明细（从第 2 行起逐站追加）。 */

var ARCH_HEAD = ['会计月', '区间', '零售吨', '柴油吨', '汽油吨', '小额吨',
                 '全口径吨', '含税金额元', '万元', '站数', '天数', '写入时间'];
var ARCH_ST_HEAD = ['会计月', '站名', '路段', '零售吨', '柴油吨', '汽油吨',
                    '升', '含税金额元', '万元'];
var ARCH_ST_COL = 14;                     // 站级明细起始列 = N

function colName(n) {
  var out = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    out = String.fromCharCode(65 + m) + out;
    n = (n - m - 1) / 26;
  }
  return out;
}

function archiveToSheet(D) {
  /* ★ v2.9 起 CFG.ARCH_SHEET 默认为空 —— 归档搬到独立的中转簿了（见 CFG 注释）。
     这里第一句就返回，连 sheet() 都不调，避免拿空表名去 Sheets.Item('') 试探。 */
  if (!CFG.ARCH_SHEET) {
    say('（归档落表已迁移到「成品油经营数据中转簿」，本簿不再写历史 —— 见 transit_sync.js）');
    return;
  }
  var sh = sheet(CFG.ARCH_SHEET);
  if (!sh) {
    say('（本簿没有「' + CFG.ARCH_SHEET + '」表，跳过归档落表 —— 新建一张同名空表即可启用）');
    return;
  }
  /* ★ D 是 build() 的返回值，meta 里**没有** month_order（那是 build 端拼 payload 时才加的）。
     会计月标签从 accounting_month 推：'2026-10-26 ~ 2026-11-25' → 取右半段前 7 位 '2026-11'，
     与页面 months 的键完全一致，便于日后对账。 */
  var range = (D.meta && D.meta.accounting_month) || '';
  var _rParts = String(range).split('~');
  var acm = _rParts.length > 1 ? s(_rParts[1]).substring(0, 7) : '';
  var days = (D.meta && D.meta.days_with_data) || [];
  if (!acm || !days.length) { say('（归档跳过：会计月或数据天数为空）'); return; }

  /* ---- 1) 汇总 ---- */
  var ret = 0, die = 0, gas = 0, xt = 0, amt = 0, wan = 0, stMap = {}, stOrder = [];
  for (var i = 0; i < days.length; i++) {
    var rs = D.sale[days[i]];
    if (!rs) continue;
    var x = rs.xiao || {};
    ret += rs.t || 0; die += rs.diesel || 0; gas += rs.gas || 0;
    xt += x.t || 0; amt += rs.amt || 0; wan += rs.wan || 0;
    var ss = rs.stations || {};
    for (var nm in ss) {
      if (!ss.hasOwnProperty(nm)) continue;
      var v = ss[nm];
      if (!stMap[nm]) { stMap[nm] = { road: v.road || '', t: 0, d: 0, g: 0, l: 0, a: 0, w: 0 }; stOrder.push(nm); }
      var o = stMap[nm];
      o.t += v.t || 0; o.d += v.diesel || 0; o.g += v.gas || 0;
      o.l += v.l || 0; o.a += v.amt || 0; o.w += v.wan || 0;
    }
  }
  stOrder.sort(function (a, b) { return stMap[b].t - stMap[a].t; });

  /* ---- 2) 月汇总 upsert：A 列找会计月，命中覆盖；否则写第一个空行 ---- */
  if (s(getCell(sh, 'A1')) !== ARCH_HEAD[0]) {
    for (var h = 0; h < ARCH_HEAD.length; h++) setCell(sh, colName(h + 1) + '1', ARCH_HEAD[h]);
  }
  var target = 0, firstEmpty = 0;
  for (var r = 2; r <= 200; r++) {
    var cv = s(getCell(sh, 'A' + r));
    if (cv === acm) { target = r; break; }
    if (!cv && !firstEmpty) firstEmpty = r;
  }
  if (!target) target = firstEmpty || (200 + 1);

  var line = [acm, range, r2(ret), r2(die), r2(gas), r2(xt), r2(ret + xt),
              r2(amt), r3(wan), stOrder.length, days.length,
              fmtDate(today0()) + ' ' + VER];
  for (var c = 0; c < line.length; c++) setCell(sh, colName(c + 1) + target, line[c]);

  /* ---- 3) 站级明细 upsert：一次性读 N~V 区，内存定位 ---- */
  if (s(getCell(sh, colName(ARCH_ST_COL) + '1')) !== ARCH_ST_HEAD[0]) {
    for (var k = 0; k < ARCH_ST_HEAD.length; k++) {
      setCell(sh, colName(ARCH_ST_COL + k) + '1', ARCH_ST_HEAD[k]);
    }
  }
  /* ★ 600 是「还没补历史」时的安全值，补完 2024-01～2026-10 历史后
     N~V 区已有 1002 行（34 月 × 28~30 站）。若仍按 600 扫，
     当前月的站级行会被判成「没见过」→ 从第 601 行起覆盖写，
     把已有历史行整段冲掉。所以必须留足余量。 */
  var ROWS = 3000;
  var grid = null;
  try { grid = sh.Range('N1:V' + ROWS).Value2; } catch (e) { grid = null; }
  var used = {};                                   // "会计月|站名" → 行号
  if (grid && grid.length) {
    for (var gi = 1; gi < grid.length; gi++) {
      var gk = s(grid[gi][0]) + '|' + s(grid[gi][1]);
      if (s(grid[gi][0]) && s(grid[gi][1])) used[gk] = gi + 1;
    }
  }
  var lastRow = 1;
  for (var t in used) { if (used[t] > lastRow) lastRow = used[t]; }

  var wrote = 0;
  for (var q = 0; q < stOrder.length; q++) {
    var sn = stOrder[q], sv = stMap[sn];
    var key = acm + '|' + sn;
    var row = used[key] || (++lastRow);
    var cells = [acm, sn, sv.road, r2(sv.t), r2(sv.d), r2(sv.g), r2(sv.l), r2(sv.a), r3(sv.w)];
    for (var m = 0; m < cells.length; m++) {
      setCell(sh, colName(ARCH_ST_COL + m) + row, cells[m]);
    }
    wrote++;
  }
  say('✅ 已落表「' + CFG.ARCH_SHEET + '」：会计月 ' + acm
      + ' 汇总行 ' + target + '，站级 ' + wrote + ' 行'
      + '（零售 ' + r2(ret) + ' 吨 + 小额 ' + r2(xt) + ' 吨 = ' + r2(ret + xt) + ' 吨）');
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
    say('JSON ' + utf8Len(json) + ' 字节（' + json.length + ' 字符）');

    /* 归档落表：与 PUSH 无关，即使不推 GitHub 也先把数据存进中转表 */
    try { archiveToSheet(D); } catch (eA) { say('归档落表失败（不影响推送）：' + (eA && eA.message ? eA.message : eA)); }

    if (!CFG.PUSH) {
      say('PUSH=false，仅解析未推送。核对上面数字无误后把 CFG.PUSH 改成 true');
      flush();
      return json;
    }

    /* 0 天保护：一个日块都没读到就别推了。
       这种情况几乎都是源表被改名 / 权限丢失 / 日块标题变动，
       推上去等于把页面清空，比不推危害大得多。 */
    if (!days.length) {
      say('⚠️ 本次没读到任何有数据的日块，中止推送（远端保持上一版不动）');
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

    /* ★ 会计月起步放行（v2.7）
       每月 26 号清表后，新会计月只有 1 天数据，UTF-8 体积天然只有月末的 1/30。
       体积护栏会把它当成「读表不全」拦下，而且会**连拦 8 天**（直到累计到 85%），
       那 8 天页面上一直是旧会计月的数字 —— 比不推更糟。
       判据：本次最早的一个日块就是会计月第一天（26 号）→ 处于月初起步阶段，放行。
       月中某天读表不全时 days[0] 不是 26 号，护栏照旧生效。 */
    var _acStart = String((D.meta && D.meta.accounting_month) || '').split('~')[0].trim();
    var _acDay = _acStart.indexOf('-') >= 0 ? _acStart.split('-')[2] : '';
    var _isMonthStart = !!(days.length && _acDay
      && parseInt(days[0], 10) === parseInt(_acDay, 10));
    if (_isMonthStart) {
      say('本会计月刚起步（首日 ' + _acStart + '，共 ' + days.length + ' 天），体积护栏放行');
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
       这种情况下宁可这次不推，让远端保持上一版完整数据，等人工查。

       ★ 2026-10-29 自己踩的坑：这里必须比 **UTF-8 字节数**，不能比 json.length。
         json.length 是 JS 字符数，中文站名一字占 3 字节，两者差约 11%：
         实测 70,555 字符 ↔ 78,687 字节。拿字符数比字节数，
         一次正常的「长小数截到 4 位」（体积本就该降）被算成 84.7%，
         差点误判成读表不全。 */
    if (cur && cur.size > 1000 && !_isMonthStart) {
      var nBytes = utf8Len(json);
      var ratio = nBytes / cur.size;
      if (ratio < 0.85) {
        say('⚠️ 中止推送：新内容 ' + nBytes + ' 字节，只有远端 ' + cur.size
          + ' 字节的 ' + (ratio * 100).toFixed(1) + '%（阈值 85%）');
        say('   → 大概率本次读表不全。远端保持上一版不动，请检查日块/权限后重跑。');
        flush();
        return json;
      }
      say('体积对比：新 ' + nBytes + ' 字节 / 远端 ' + cur.size
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
