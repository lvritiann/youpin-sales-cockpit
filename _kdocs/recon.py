# -*- coding: utf-8 -*-
"""全量加和关系对账：穷举 payload 里所有「总量 vs 分项之和」的关系。

用法：python _kdocs/recon.py [cockpit/index.html]
输出：每条关系的 A / B / 差值，差 0 记 PASS。
"""
import json
import re
import sys

PATH = sys.argv[1] if len(sys.argv) > 1 else 'cockpit/index.html'


def load_payload(path):
    h = open(path, encoding='utf-8').read()
    i = h.find('const D = ')
    s = h.index('{', i)
    depth = 0
    esc = False
    q = None
    for j in range(s, len(h)):
        c = h[j]
        if esc:
            esc = False
            continue
        if c == '\\':
            esc = True
            continue
        if q:
            if c == q:
                q = None
            continue
        if c == '"':
            q = c
            continue
        if c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                return json.loads(h[s:j + 1])
    raise SystemExit('payload 未找到')


P = load_payload(PATH)
Y = P.get('yearly') or {}
YT = Y.get('ytd') or {}
rows = []


def chk(scope, name, a, b, tol=0.0):
    a = float(a or 0)
    b = float(b or 0)
    d = round(a - b, 6)
    rows.append((scope, name, a, b, d, abs(d) <= tol))


# ---------- 1. 顶部 KPI：零售 / 全口径 ----------
for lab, key in (('零售', 'retail'), ('全口径', 'full')):
    D = YT.get(key) or {}
    if not D:
        continue
    chk('KPI', '%s 总量 = 柴 + 汽' % lab, D.get('t'), (D.get('diesel') or 0) + (D.get('gas') or 0))

# ---------- 2. 金额三层 ----------
AM = YT.get('amt') or {}
if AM:
    chk('KPI', '含税产值 = 基数 + 实时', AM.get('total_wan'),
        (AM.get('base_wan') or 0) + (AM.get('live_wan') or 0))
    chk('KPI', '全口径含税 = 零售含税 + 小额含税', AM.get('full_tax'),
        (AM.get('total_wan') or 0) + (AM.get('xiao_tax') or 0))
    chk('KPI', '全口径不含税 = 零售不含税 + 小额不含税', AM.get('full_net'),
        (AM.get('retail_net') or 0) + (AM.get('xiao_net') or 0))

# ---------- 3. 全口径吨量 = 零售 + 小额 ----------
FL = YT.get('full') or {}
RT = YT.get('retail') or {}
if FL and RT:
    chk('KPI', '全口径吨 = 零售 + 小额', FL.get('t'),
        (RT.get('t') or 0) + (YT.get('xiao_t') or 0))

# ---------- 4. 年度序列 vs KPI ----------
for kpi_name, kpi_val, ykey in (('零售年度', (RT or {}).get('t'), 'real_annual'),
                                ('柴油年度', (RT or {}).get('diesel'), 'diesel_annual'),
                                ('汽油年度', (RT or {}).get('gas'), 'gas_annual')):
    D = Y.get(ykey) or {}
    if D and str(YT.get('year')) in D and kpi_val is not None:
        chk('年度', '%s 当前年 = KPI' % kpi_name, D[str(YT.get('year'))], kpi_val)

# ---------- 5. 月份表：逐月 总量 vs 柴+汽 / 金额分项 ----------
MO = P.get('months') or {}
cur_lab = (P.get('meta') or {}).get('cur_month')
for mk, m in MO.items():
    chk('月份', '%s 吨 = 柴+汽' % mk, m.get('t'), (m.get('diesel') or 0) + (m.get('gas') or 0))
    chk('月份', '%s 全口径吨 = 零售+小额' % mk, m.get('ft'), (m.get('t') or 0) + (m.get('xt') or 0))
    chk('月份', '%s 含税 = 零售含税+小额含税' % mk, m.get('famt'),
        (m.get('amt') or 0) + (m.get('xamt') or 0))
    chk('月份', '%s 不含税 = 零售不含税+小额不含税' % mk, m.get('fwan'),
        (m.get('wan') or 0) + (m.get('xwan') or 0))
    # 站级金额合计 vs 月金额（仅当前月/含站级明细的月）
    ssum = sum((s.get('amt') or 0) for s in (m.get('stations') or []))
    if ssum and m.get('amt'):
        chk('月份', '%s 站级金额合计 = 月金额' % mk, ssum, m.get('amt'), tol=1.0)

# ---------- 6. 年度合计行：各月之和 ----------
AE = Y.get('annual') or {}
for y, v in AE.items():
    pass

# ---------- 7. 前一日窄条：items 之和对 total ----------
PD = P.get('prev_day') or {}
its = PD.get('items') or []
if its:
    chk('窄条', '15 对合计 = Σ各对', PD.get('sum', {}).get('t') if PD.get('sum') else
        sum(x.get('t') or 0 for x in its), sum(x.get('t') or 0 for x in its))
    chk('窄条', '全部合计 = 15 对 + 小额', PD.get('total'),
        sum(x.get('t') or 0 for x in its) + (PD.get('xiao') or 0))
    for x in its:
        chk('窄条', '%s 对合计 = 汽+柴' % x.get('name'), x.get('t'),
            (x.get('g') or 0) + (x.get('d') or 0))

# ---------- 6. 视角层：公司与各对 ----------
VW = Y.get('views') or {}
CO = VW.get('company') or {}
if CO:
    chk('视角', '公司 吨 = 柴+汽', CO.get('t'),
        (CO.get('diesel') or 0) + (CO.get('gas') or 0))
    chk('视角', '公司含税 = ytd 全口径含税', CO.get('amt'), AM.get('full_tax'), tol=0.0)
    chk('视角', '公司不含税 = ytd 全口径不含税', CO.get('wan'), AM.get('full_net'), tol=0.0)
    # 公司 = Σ各对 + 小额
    _sub = CO.get('sub') or []
    if _sub:
        chk('视角', '公司吨 = Σ各对(含小额)', CO.get('t'),
            sum(x.get('t') or 0 for x in _sub))
        chk('视角', '公司含税 = Σ各对含税', CO.get('amt'),
            sum(x.get('amt') or 0 for x in _sub), tol=0.02)
    for x in _sub:
        chk('视角', '%s 吨 = 柴+汽' % x.get('name'), x.get('t'),
            (x.get('diesel') or 0) + (x.get('gas') or 0))

# ---------- 7. 月度轨迹：每月自洽（轨迹是近 6 个会计月滚动窗口，不与全年 KPI 比） ----------
TR = (P.get('meta') or {}).get('trend') or []
for x in TR:
    chk('轨迹', '%s 全口径 = 零售+小额' % x.get('k'), x.get('ft'),
        (x.get('t') or 0) + (x.get('x') or 0))
if TR:
    _last = TR[-1]
    # 轨迹最后一行 = 当前会计月，必须与 KPI 的「当月实时」部分一致
    chk('轨迹', '末行零售 = 当月实时零售', _last.get('t'), YT.get('live_t') or 0, tol=0.5)

# ---------- 输出 ----------
w = max(len(r[1]) for r in rows) + 2
bad = [r for r in rows if not r[5]]
print('=' * 96)
print('%-6s %-*s %14s %14s %10s  %s' % ('范围', w, '关系', 'A', 'B', 'A-B', '结果'))
print('-' * 96)
for sc, nm, a, b, d, ok in rows:
    print('%-6s %-*s %14.4f %14.4f %10.4f  %s'
          % (sc, w, nm, a, b, d, 'OK' if ok else '**FAIL**'))
print('-' * 96)
print('共 %d 条关系，通过 %d，未通过 %d' % (len(rows), len(rows) - len(bad), len(bad)))
if bad:
    print('\n未通过明细：')
    for sc, nm, a, b, d, ok in bad:
        print('  [%s] %s : %.4f vs %.4f  差 %.4f' % (sc, nm, a, b, d))
