# -*- coding: utf-8 -*-
"""驾驶舱 v2：会计月切换 + 多月趋势 + 历年同比"""
import json, os, datetime

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
ARCH = os.path.join(ROOT, '_archive')

HM = json.load(open(os.path.join(ARCH, 'hist_months.json'), encoding='utf-8'))['months']
YE = json.load(open(os.path.join(ARCH, 'history_yearly.json'), encoding='utf-8'))
CUR = json.load(open(os.path.join(BASE, 'data.json'), encoding='utf-8'))

# 三年（2024/2025/2026）日报全量：逐月零售·小额·柴·汽，以及逐站逐年
# 来源 D:\成品油销量\日报20xx\，由 scan_folder.py / months_cum.py / annual_roll.py 产出
MC = json.load(open(os.path.join(ARCH, 'months_cum.json'), encoding='utf-8'))
SA = json.load(open(os.path.join(ARCH, 'stations_annual.json'), encoding='utf-8'))
FS = {r['file'][:6]: r for r in
      json.load(open(os.path.join(ARCH, 'folder_scan.json'), encoding='utf-8'))}
# 逐站 × 逐月（含分品类金额）、逐站近月逐日、逐站渠道结构
# 来源 D:\成品油销量\日报20xx\，由 scan_station.py 产出
SER = json.load(open(os.path.join(ARCH, 'station_series.json'), encoding='utf-8'))

# 「小额配送」在源表是公司级独立一行（无站级归属），scan_station.py 已把它收成
# 虚拟站 XIAO，字段与普通站完全一致。驾驶舱里它与 15 对平权：独立成行、可下钻、
# 可拆柴汽、可比同比。这样「零售」与「全口径」的差异由一行数据直接体现，
# 不再依赖一句口头口径说明。
XIAO = '小额配送'


def _acc(ym_list, sts, key):
    """若干个「年月」月表里，指定站集合的某字段合计（单位沿用源字段）"""
    tot = 0.0
    for _ym in ym_list:
        _mm = SER['months'].get(_ym) or {}
        for _st in sts:
            tot += float((_mm.get(_st) or {}).get(key) or 0)
    return tot


# 30 个真实站（零售口径的站集合），与 XIAO 严格互斥
ALL_REAL_ST = sorted(set(SA['stations'].keys()))


# 按年归拢日报口径月度序列：YRB['2026'] = {'1': {...}, ...}
# 零售/柴/汽取「日块逐站累加」而非「月销售累计相减」——
# 202402 原表日块 15,768.39 与其月累计 15,975.38 差 207 吨，若混用会导致
# 年度条形图、月份表、各站年度表三处数字各不相同。统一到日块后三者完全自洽。
YRB = {}
for _tag, _v in MC.items():
    _day = FS.get(_tag, {}).get('daily', {})
    if _day:
        _t = round(sum(d['t'] for d in _day.values()), 2)
        _d = round(sum(d['diesel'] for d in _day.values()), 2)
        _g = round(sum(d['gas'] for d in _day.values()), 2)
    else:
        _t, _d, _g = _v['m_retail'], _v['m_diesel'], _v['m_gas']
    YRB.setdefault(_v['year'], {})[str(_v['mon'])] = {
        'm_retail': _t, 'm_diesel': _d, 'm_gas': _g,
        'm_xiao': _v['m_xiao'],                       # 小额只能取「月销售累计」，日块常年记 0
        'm_total': round(_t + _v['m_xiao'], 2),
        'cum_retail': _v['m_retail'],                 # 月累计原值，仅作对账
    }


def rb_sum(y, months_=None, key='m_retail'):
    d = YRB.get(str(y), {})
    ks = months_ if months_ is not None else d.keys()
    return round(sum(d[k][key] for k in ks if k in d), 2)

# 站名 → 路段映射：data.json 解析自最新日报，路段字段完整；
# 历史归档 hist_months.json 里部分站的 road 是字符串 "None"（表格合并单元格被当文本），
# 导致「路段结构」和「15 对」按 road 分组时被拆散，这里用它反向修复。
ROADMAP = {}
for _d, _v in CUR.get('sale', {}).items():
    for _n, _x in _v.get('stations', {}).items():
        _r = str(_x.get('road') or '').strip()
        if _r and _r not in ('None', 'none', 'null'):
            ROADMAP[_n] = _r


def fix_road(sts):
    for s in sts:
        r = str(s.get('road') or '').strip()
        s['road'] = r if (r and r not in ('None', 'none', 'null')) else ROADMAP.get(s.get('name'), '')
    return sts

RANGE = {'2026-09': '2026-08-26 ~ 09-25', '2026-10': '2026-09-26 ~ 10-25',
         '2026-11': '2026-10-26 ~ 11-25'}
NAME = {'2026-09': '9月会计月', '2026-10': '10月会计月', '2026-11': '11月会计月'}


def pack_month(mkey, raw):
    T = raw.get('total', raw)
    t = T.get('t', 0)
    st = sorted(fix_road(raw['stations']), key=lambda x: -x['t'])
    # ★ 单位归一：归档月（hist_months.json）站级 amt 是「元」，实时月（data.json）是「万元」。
    #   用「站级合计 ÷ 月级合计」的量级判定（1e4 倍差），统一折成万元，
    #   否则明细表金额列与吨均价在切到 9/10 月时会差 10000 倍。
    _samt = sum(s.get('amt', 0) or 0 for s in st)
    _matm = T.get('amt', 0) or 0
    if _samt > 0 and _matm > 0 and _samt / _matm > 100:
        for s in st:
            s['amt'] = round((s['amt'] or 0) / 1e4, 2)
    for i, s in enumerate(st):
        s['share'] = round(s['t'] / t * 100, 1) if t else 0
        s['price'] = round(s['amt'] * 1e4 / s['t']) if s['t'] else 0
        s['price0'] = round(s['wan'] * 1e4 / s['t']) if s['t'] else 0
    ch = sorted([{'name': k, **v, 'share': round(v['t'] / t * 100, 1) if t else 0}
                 for k, v in raw['channels'].items()], key=lambda x: -x['t'])
    cl = raw.get('clients', [])
    bt = sum(c['t'] for c in cl)
    for c in cl:
        c['share'] = round(c['l'] / sum(x['l'] for x in cl) * 100, 1) if cl else 0
    ba = sum(c['amt'] for c in cl)
    return {
        'key': mkey, 'name': NAME.get(mkey, mkey), 'range': RANGE.get(mkey, ''),
        'partial': mkey == '2026-11',
        't': round(t, 1), 'diesel': round(T.get('diesel', 0), 1),
        'gas': round(T.get('gas', 0), 1), 'amt': round(T.get('amt', 0), 1),
        'wan': round(T.get('wan', 0), 1), 'days': T.get('days', len(raw.get('daily', []))),
        'price': round(T.get('amt', 0) * 1e4 / t) if t else 0,
        'stations': st, 'channels': ch, 'clients': cl,
        'big_t': round(bt, 1), 'big_a': round(ba, 1),
        'big_share': round(bt / t * 100, 1) if t else 0,
        'big_ashare': round(ba / T.get('amt', 1) * 100, 1) if T.get('amt') else 0,
        'daily': raw.get('daily', []),
        'src': raw.get('sources', []),
        'xt': round(raw.get('xiao', {}).get('t', 0), 1),
        'xamt': round(raw.get('xiao', {}).get('amt', 0), 1),
        'xwan': round(raw.get('xiao', {}).get('wan', 0), 1),
        'ft': round(t + raw.get('xiao', {}).get('t', 0), 1),
        'famt': round(T.get('amt', 0) + raw.get('xiao', {}).get('amt', 0), 1),
        'fwan': round(T.get('wan', 0) + raw.get('xiao', {}).get('wan', 0), 1),
    }


months = {}
for mk, v in HM.items():
    months[mk] = pack_month(mk, v)

# 当前月（kdocs 实时）
cur_days = [d for d in CUR['sale'] if CUR['sale'][d]['t'] > 1]
cur_t = sum(CUR['sale'][d]['t'] for d in cur_days)
cur_amt = sum(CUR['sale'][d]['amt'] for d in cur_days) / 1e4
cur_wan = sum(CUR['sale'][d]['wan'] for d in cur_days)
stc, chc, clc = {}, {}, {}
for d in cur_days:
    for n, v in CUR['sale'][d]['stations'].items():
        a = stc.setdefault(n, dict(name=n, road=v['road'], t=0, diesel=0, gas=0, amt=0, wan=0))
        a['t'] += v['t']; a['amt'] += v['amt'] / 1e4; a['wan'] += v['wan']
        a['diesel'] += v['diesel']; a['gas'] += v['gas']
    for c, dd in CUR['channels'].items():
        if d in dd:
            x = chc.setdefault(c, dict(t=0, amt=0, wan=0))
            x['t'] += sum(y['t'] for y in dd[d].values())
            x['amt'] += sum(y['amt'] for y in dd[d].values()) / 1e4
            x['wan'] += sum(y.get('wan', 0) for y in dd[d].values())
    for stn, rec in CUR['bigcust'].get('大客户明细', {}).get(d, {}).items():
        for cn, y in rec['c'].items():
            x = clc.setdefault(cn, dict(name=cn, l=0, amt=0))
            x['l'] += y['l']; x['amt'] += y['amt']
for x in clc.values():
    x['t'] = round(x['l'] * 0.83 / 1000, 1)
    x['amt'] = round(x['amt'] / 1e4, 2)
months['2026-11'] = pack_month('2026-11', {
    't': cur_t, 'diesel': sum(CUR['sale'][d]['diesel'] for d in cur_days),
    'gas': sum(CUR['sale'][d]['gas'] for d in cur_days),
    'amt': cur_amt, 'wan': cur_wan, 'days': len(cur_days),
    'stations': list(stc.values()), 'channels': chc, 'clients': list(clc.values()),
    'xiao': {'t': round(sum(CUR['sale'][d].get('xiao', {}).get('t', 0) for d in cur_days), 2),
             'amt': round(sum(CUR['sale'][d].get('xiao', {}).get('amt', 0) for d in cur_days) / 1e4, 2),
             'wan': round(sum(CUR['sale'][d].get('xiao', {}).get('wan', 0) for d in cur_days), 2)},
    'daily': [{'d': d, 't': round(CUR['sale'][d]['t'], 1),
               'diesel': round(CUR['sale'][d]['diesel'], 1),
               'gas': round(CUR['sale'][d]['gas'], 1),
               'amt': round(CUR['sale'][d]['amt'] / 1e4, 1),
               'wan': round(CUR['sale'][d]['wan'], 1),
               'xt': round(CUR['sale'][d].get('xiao', {}).get('t', 0), 1),
               'lab': ('10-%02d' % int(d)) if int(d) >= 26 else ('11-%02d' % int(d))}
              for d in sorted(cur_days, key=lambda x: int(x))],
    'sources': ['kdocs'],
})

# ---- 每月补小额配送的分品类吨（pack_month 的 xiao 只有总量，表格要拆柴汽） ----
# 数据源同 SER，与「小额配送」实体的口径完全一致；11 月源表未拆 → 留 0 不臆造。
for _mk, _mob in months.items():
    _xr = (SER['months'].get(_mk.replace('-', '')) or {}).get(XIAO) or {}
    _mob['xd'] = round(float(_xr.get('d') or 0), 1)
    _mob['xg'] = round(float(_xr.get('g') or 0), 1)

# ---- 日销量滚动窗口：按真实日期倒推最近 N 天，跨会计月拼接 ----
def acct_date(mkey, dayno):
    """会计月 key + 日号 -> 真实日期（>=26 落上月，<26 落本月）"""
    y, m = (int(x) for x in mkey.split('-'))
    if dayno >= 26:
        return datetime.date(y - 1 if m == 1 else y, (m - 1) or 12, dayno)
    return datetime.date(y, m, dayno)


BYDATE = {}
for _mk, _m in months.items():
    for _d in _m['daily']:
        _dn = int(_d['d'])
        if (_d.get('t') or 0) <= 0:
            continue          # 未填报的日块不占位，计入缺失
        BYDATE[acct_date(_mk, _dn)] = _d

LAST_DT = max(BYDATE) if BYDATE else None
WINDOWS = {}
if LAST_DT:
    for _n in (30, 45, 60):
        _start = LAST_DT - datetime.timedelta(days=_n - 1)
        _seq, _miss, _cur = [], [], _start
        while _cur <= LAST_DT:
            _d = BYDATE.get(_cur)
            if _d is None:
                _miss.append(_cur.isoformat())
            else:
                _seq.append({'date': _cur.isoformat(), 'lab': _cur.strftime('%m-%d'),
                             't': round(_d.get('t', 0) or 0, 1),
                             'diesel': round(_d.get('diesel', 0) or 0, 1),
                             'gas': round(_d.get('gas', 0) or 0, 1),
                             'amt': round(_d.get('amt', 0) or 0, 1),
                             'wan': round(_d.get('wan', 0) or 0, 1),
                             'xt': round(_d.get('xt', 0) or 0, 1)})
            _cur += datetime.timedelta(days=1)
        WINDOWS[str(_n)] = {
            'days': _seq, 'missing': _miss,
            'start': _start.isoformat(), 'end': LAST_DT.isoformat(),
            'total': round(sum(x['t'] for x in _seq), 1),
            'avg': round(sum(x['t'] for x in _seq) / len(_seq), 1) if _seq else 0,
        }

# ---- 历年 ----
yrs = YE['years']
yearly = {
    'years': yrs,
    'annual': {str(y): round(YE['annual'].get('合计', {}).get(str(y), {}).get('total', 0), 1)
               for y in yrs},
    'months': {str(m): {str(y): round(YE['months'][str(m)].get('合计', {}).get(str(y), {}).get('total', 0), 1)
                        for y in yrs} for m in range(1, 13)},
    'xiao': {str(m): {str(y): round(YE['months'][str(m)].get('小额', {}).get(str(y), {}).get('total', 0), 1)
                      for y in yrs} for m in range(1, 13)},
    'annual_xiao': {str(y): round(YE['annual'].get('小额', {}).get(str(y), {}).get('total', 0), 1)
                    for y in yrs},
    'st_ytd': {},
}
# 站级 2026 累计 vs 2025 同期（仅有数据月份）
done = [m for m in range(1, 13) if YE['months'][str(m)].get('合计', {}).get('2026', {}).get('total', 0) > 0]
for name, rec in YE['annual'].items():
    if name in ('合计',):
        continue
    y26 = sum(YE['months'][str(m)].get(name, {}).get('2026', {}).get('total', 0) for m in done)
    y25 = sum(YE['months'][str(m)].get(name, {}).get('2025', {}).get('total', 0) for m in done)
    if y26 or y25:
        yearly['st_ytd'][name] = {'y26': round(y26, 1), 'y25': round(y25, 1),
                                  'yoy': round((y26 - y25) / y25 * 100, 1) if y25 else None}

# ---- 年度累计（年初 → 最新有数据自然月），分零售/全口径，含柴汽油与同期 ----
def agg(key, y, mons):
    d = {'t': 0.0, 'diesel': 0.0, 'gas': 0.0}
    for m in mons:
        r = YE['months'][str(m)].get(key, {}).get(str(y), {}) or {}
        d['t'] += r.get('total', 0) or 0
        d['diesel'] += r.get('diesel', 0) or 0
        d['gas'] += r.get('gas', 0) or 0
    return {k: round(v, 1) for k, v in d.items()}


def ret_agg(y, mons):
    """零售 = 合计 − 小额（月表里没有单独'零售'键）"""
    f, x = agg('合计', y, mons), agg('小额', y, mons)
    return {k: round(max(0.0, f[k] - x[k]), 1) for k in f}

CUR_Y, PRE_Y = 2026, 2025

# ---- 年度累计：基线 =「月销售累计·零售」(1~10 月，含分品类)，叠加当前会计月实时天数 ----
YP = json.load(open(os.path.join(ARCH, 'yearly_progress.json'), encoding='utf-8'))
LIVE_KEY, LIVE_NO = '2026-11', 11
live = months[LIVE_KEY]
live_days = [d for d in live['daily'] if (d.get('t') or 0) > 0]
live_t = round(sum(d['t'] for d in live_days), 2)
live_d = round(sum(d.get('diesel', 0) for d in live_days), 2)
live_g = round(sum(d.get('gas', 0) for d in live_days), 2)
live_xt = round(live.get('xt') or 0, 2)

_mx = YP['ytd_mix']
_br, _bx = _mx['retail'], (_mx.get('xiao_mix') or {})

# 口径统一切到「日报逐月」（D:\成品油销量），不再用月报链，理由见 memory：
# 日报口径可拆到每一站、每一天，且 2024–2026 三年同源，年度表才不会有缺口。
_base_mons = sorted([k for k in YRB.get(str(CUR_Y), {})], key=int)   # 已完成会计月 1..10

# 视角层与金额层共用的月份键 / 同期日号（提前定义，避免下面两段各算一套）
BASE_YM26 = ['%d%02d' % (CUR_Y, int(x)) for x in _base_mons]
BASE_YM25 = ['%d%02d' % (PRE_Y, int(x)) for x in _base_mons]
LIVE_DN = [int(d['d']) for d in live_days]
PREV_YM_LIVE = '%d%02d' % (PRE_Y, LIVE_NO)

_br_t = rb_sum(CUR_Y, _base_mons, 'm_retail')
_br_d = rb_sum(CUR_Y, _base_mons, 'm_diesel')
_br_g = rb_sum(CUR_Y, _base_mons, 'm_gas')
# ★ 小额改取「日块逐站累加」而不是「月销售累计·小额」：后者是月报链口径，
#   与零售不同源，2026 年两处金额列差 1.5 亿元（月累计列含批发口径）。
#   统一到日块后，15 对零售 + 小额 严格等于全口径，三处数字不再各说各话。
_br_x = round(_acc(BASE_YM26, [XIAO], 't'), 2)

disp_t = round(_br_t + live_t, 1)
disp_d = round(_br_d + live_d, 1)
disp_g = round(_br_g + live_g, 1)
# 全口径 = 零售 + 小额（含当月实时那几天的小额，11 月可能为 0 但不能漏项）
full_t = round(_br_t + _br_x + live_t + live_xt, 1)
full_d = round(_br_d + live_d, 1)
full_g = round(_br_g + live_g, 1)


def acct_md(dayno):
    """会计月日号 -> 真实月-日：>=26 落在上月，<26 落在本月"""
    return '%02d-%02d' % ((LIVE_NO - 1) or 12, dayno) if dayno >= 26 else '%02d-%02d' % (LIVE_NO, dayno)


prev_same, as_of = None, ''
if live_days:
    _last = max(int(d['d']) for d in live_days)
    as_of = acct_md(_last)

# 同期基数 = 去年 1..N 月 + 去年同一会计月已过天数（从去年同月日报的日块取），保证分子分母同期
_pd = FS.get('%d%02d' % (PRE_Y, LIVE_NO), {}).get('daily', {})
_ps = {'t': 0.0, 'diesel': 0.0, 'gas': 0.0}
for _d in live_days:
    _cell = _pd.get(str(int(_d['d'])))
    if _cell:
        _ps['t'] += _cell['t']; _ps['diesel'] += _cell['diesel']; _ps['gas'] += _cell['gas']
prev_same = round(rb_sum(PRE_Y, _base_mons, 'm_retail') + _ps['t'], 2)
prev_d = round(rb_sum(PRE_Y, _base_mons, 'm_diesel') + _ps['diesel'], 2)
prev_g = round(rb_sum(PRE_Y, _base_mons, 'm_gas') + _ps['gas'], 2)

flow_now = round(_br_t + live_t, 2)
yoy_all = round((flow_now / prev_same - 1) * 100, 1) if prev_same else None
yoy_d = round((disp_d / prev_d - 1) * 100, 1) if prev_d else None
yoy_g = round((disp_g / prev_g - 1) * 100, 1) if prev_g else None

# ---- 产值金额：含税销售额，日块逐站累加（元→万元），口径与销量同源 ----
def _amt_of(tags):
    return round(sum(sum(d.get('amt', 0) for d in FS[t]['daily'].values()) / 1e4
                     for t in tags if t in FS), 2)


_amt_base = _amt_of([_t for _t in MC if MC[_t]['year'] == str(CUR_Y)
                     and str(MC[_t]['mon']) in _base_mons])
_amt_prev_base = _amt_of([_t for _t in MC if MC[_t]['year'] == str(PRE_Y)
                          and str(MC[_t]['mon']) in _base_mons])
_amt_prev_live = round(sum(_pd.get(str(int(_d['d'])), {}).get('amt', 0)
                           for _d in live_days) / 1e4, 2)
_amt_all = round(_amt_base + cur_amt, 2)
_amt_prev = round(_amt_prev_base + _amt_prev_live, 2)
_amt_yoy = round((_amt_all / _amt_prev - 1) * 100, 1) if _amt_prev else None

# ---- 小额配送金额：切到「日块」口径，与零售同源 ----
# 旧实现取「月销售累计」· 小额行 · 累计金额列（不含税）×1.13 还原含税，
# 该列是月报链口径、含批发，2026 年比日块多出 15,571 万元，
# 导致「全口径产值」虚高。现改为逐站日块累加，含税/不含税分列取，三处金额自洽。
_xiao_wan = round(_acc(BASE_YM26, [XIAO], 'amt') / 1e4, 2)     # 含税万元
_xiao_wan0 = round(_acc(BASE_YM26, [XIAO], 'wan'), 2)          # 不含税万元
_retail_wan0 = round(_acc(BASE_YM26, ALL_REAL_ST, 'wan'), 2)   # 零售不含税万元
_cur_wan = round(sum(CUR['sale'][d].get('wan', 0) for d in cur_days), 2)
_retail_net_all = round(_retail_wan0 + _cur_wan, 2)
_xiao_net = _xiao_wan0
_xiao_tax = _xiao_wan

yearly['ytd'] = {
    'year': CUR_Y, 'prev': PRE_Y,
    'base_months': [int(x) for x in _base_mons],
    'live_month': LIVE_NO, 'live_days': len(live_days), 'as_of': as_of,
    'retail': {'t': disp_t, 'diesel': disp_d, 'gas': disp_g},
    'full': {'t': full_t, 'diesel': full_d, 'gas': full_g},
    'prev_same': prev_same, 'yoy': yoy_all,
    'prev_diesel': prev_d, 'prev_gas': prev_g, 'yoy_d': yoy_d, 'yoy_g': yoy_g,
    'xiao_t': _br_x,
    'xiao_share': round(_br_x / full_t * 100, 1) if full_t else 0,
    'amt': {'total_wan': _amt_all, 'base_wan': _amt_base,
            'live_wan': round(cur_amt, 2), 'prev_wan': _amt_prev, 'yoy': _amt_yoy,
            # 零售不含税 / 小额（不含税·含税）/ 全口径四套值
            'retail_net': _retail_net_all,
            'xiao_net': _xiao_net, 'xiao_tax': _xiao_tax,
            'full_tax': round(_amt_all + _xiao_tax, 2),
            'full_net': round(_retail_net_all + _xiao_net, 2),
            'xiao_src': '日报日块·小额配送行·当日销售额/销售收入列'},
    'flow_now': flow_now,
    'gap_vs_flow': round(_br_t - YP['ytd_retail'][str(CUR_Y)], 2),
    'src': '2024–2026 日报（D:\\成品油销量）',
}

yearly['done_months'] = [int(x) for x in _base_mons]
yearly['yp_months'] = YP['months_retail']
yearly['yp_annual'] = YP['ytd_retail']   # 月报链官方累计，仅作对照，不再参与主展示

# ---- 2023：只有公司台账（自然月），零售 = 合计 − 小额；无日报文件 ----
# 2024 起一律用日报，台账只作对照
hist_m, hist_a = {}, {}
for _y in (2023,):
    _tot = 0.0
    for _m in range(1, 13):
        _f = YE['months'][str(_m)].get('合计', {}).get(str(_y), {}) or {}
        _x = YE['months'][str(_m)].get('小额', {}).get(str(_y), {}) or {}
        _v = max(0.0, (_f.get('total', 0) or 0) - (_x.get('total', 0) or 0))
        hist_m.setdefault(str(_y), {})[str(_m)] = round(_v, 2)
        _tot += _v
    hist_a[str(_y)] = round(_tot, 2)

yearly['hist_retail_months'] = hist_m
yearly['hist_retail_annual'] = hist_a

# ---- 2024–2026：日报口径逐月（真实来源 D:\成品油销量） ----
yearly['rb_months'] = {_y: {_k: round(_v['m_retail'], 2) for _k, _v in _d.items()}
                       for _y, _d in YRB.items()}
# ★ 小额逐月序列改用 SER（日块）而不是 MC（月销售累计），与「小额配送」实体同源，
#   否则历年对比勾上小额后，与实体页的月数字会差几十吨。
_rbx = {}
for _k, _v in SER['months'].items():
    _rbx.setdefault(_k[:4], {})[str(int(_k[4:]))] = round(
        float((_v.get(XIAO) or {}).get('t') or 0), 2)
_rbx.setdefault(str(CUR_Y), {})[str(LIVE_NO)] = round(live_xt, 2)
yearly['rb_xiao_months'] = _rbx
# 2023 无日报，取公司台账「小额」科目，与台账「合计」= 零售 + 小额 自洽
yearly['hist_xiao_months'] = {'2023': {
    str(_m): round(float(((YE['months'][str(_m)].get('小额', {}) or {})
                          .get('2023', {}) or {}).get('total', 0) or 0), 2)
    for _m in range(1, 13)}}
yearly['xiao_annual'] = {
    '2023': round(sum(yearly['hist_xiao_months']['2023'].values()), 2)}
for _y in ('2024', '2025', '2026'):
    yearly['xiao_annual'][_y] = round(sum(_rbx.get(_y, {}).values()), 2)
yearly['real_annual'] = {
    '2023': hist_a.get('2023', 0),
    '2024': rb_sum(2024),
    '2025': rb_sum(2025),
    '2026': round(disp_t, 2),
}
yearly['real_note'] = ('2023 为公司台账（自然月，无日报文件）；2024–2026 为各月日报逐站累计，'
                       '会计月 26 日→次月 25 日。2026 含 11 月会计月已填 %d 天。' % len(live_days))

# ---- 各对年度销量：2024–2026 走日报逐站聚合，2023 只能取台账 ----
_SUMKEYS = ('合计', '零售', '小额')
PAIR_ANN, PAIR_ROAD = {}, {}
for _pn, _rec in SA['pair'].items():
    PAIR_ANN[_pn] = {'2023': 0.0,
                     '2024': float(_rec.get('2024', 0) or 0),
                     '2025': float(_rec.get('2025', 0) or 0),
                     '2026': float(_rec.get('2026', 0) or 0)}
for _pn, _rec in YE['annual'].items():      # 2023 仅有台账
    if _pn in _SUMKEYS or _pn not in PAIR_ANN:
        continue
    PAIR_ANN[_pn]['2023'] = float((_rec.get('2023', {}) or {}).get('total', 0) or 0)

for _b, _rec in SA['pair'].items():
    _r = str(_rec.get('road') or '').strip()
    if _r and _r not in ('None', 'none', 'null'):
        PAIR_ROAD[_b] = _r
for _s in months['2026-10']['stations']:      # 缺路段时回落到当前月
    _b = _s['name'][:-1] if _s['name'][-1] in '东西南北' else _s['name']
    if _b not in PAIR_ROAD:
        PAIR_ROAD[_b] = _s['road']

yearly['pair_annual'] = PAIR_ANN
yearly['pair_road'] = PAIR_ROAD
# 公司台账（历年销量同比表）零售口径，仅用于对账说明
yearly['ledger_retail'] = {str(_y): round(
    ((YE['annual'].get('零售', {}).get(str(_y), {}) or {}).get('total', 0) or 0), 1)
    for _y in (2024, 2025)}
yearly['pair_annual_note'] = (
    '2023 为公司台账（自然月，无日报文件）；2024–2026 为各月日报逐站累计（会计月 26 日→次月 25 日）。'
    '2026 已含 8 月会计月（202608 报表）与 11 月已填 %d 天。' % len(live_days))
yearly['yp_live'] = {'month': LIVE_NO, 'days': len(live_days), 't': live_t, 'x': live_xt}
yearly['pair_stations'] = SA.get('pair_stations', {})

# ---- 小额配送作为第 16 个实体并入「对」级结构 ----------------
# 它没有站级归属，但有完整分品类吨与金额，因此按虚拟对处理：
#   列表里独立成行 → 排行榜里单列 → 点进去就是一个「单站视角」的完整体。
# 合计口径随之变成：15 对（零售）+ 小额配送 = 全口径，与顶部 KPI 严格对齐。
_PAIRS_ALL = dict(SA.get('pair_stations', {}))
_PAIRS_ALL[XIAO] = [XIAO]
PAIR_ROAD[XIAO] = '公司级'
_XA = {}
for _y in (2023, 2024, 2025, 2026):
    if _y == 2023:
        # 2023 无日报，取公司台账「小额」科目（与 15 对的台账零售口径同源，可相加）
        _XA['2023'] = round(float(
            ((YE['annual'].get('小额', {}) or {}).get('2023', {}) or {}).get('total', 0) or 0), 1)
    else:
        _XA[str(_y)] = round(sum(v.get(XIAO, {}).get('t', 0) or 0
                                 for k, v in SER['months'].items() if k.startswith(str(_y))), 1)
# 当年补上 11 月实时（小额当月为 0 也照加，保证与全口径一致）
_XA[str(CUR_Y)] = round(_XA.get(str(CUR_Y), 0) + live_xt, 1)
PAIR_ANN[XIAO] = _XA
yearly['pair_annual'] = PAIR_ANN
yearly['pair_stations_all'] = _PAIRS_ALL
yearly['xiao'] = {
    'name': XIAO, 'road': '公司级', 'stations': [XIAO],
    't26': round(_XA.get('2026', 0), 1),
    'share': round(_XA.get('2026', 0) / full_t * 100, 1) if full_t else 0,
}

# ---- 会计月轨迹：完整 6 个月跨度（含进行中当月） ----
# 原先只列 3 个月且靠页顶「月份长条」承担切换，两处信息重复；
# 现改为轨迹图直接铺 6 个月，月份切换收进控制条的紧凑 seg。
NDAYS = {k: v.get('n_days') or 30 for k, v in FS.items()}


def _acct_range(m_):
    """会计月 m = (m-1)月26日 → m月25日"""
    return '%d/%d–%d/%d' % ((CUR_Y - 1) if m_ == 1 else CUR_Y, (m_ - 1) or 12, CUR_Y, m_)


TREND = []
for _m in range(LIVE_NO - 5, LIVE_NO + 1):
    # 小额配送单列一段：轨迹柱要堆叠「柴油 + 汽油 + 小额」，才看得出结构而不是一根总高
    if _m == LIVE_NO:
        _t = round(live_t, 2)
        _dseg, _gseg = round(live_d, 2), round(live_g, 2)
        _x = round(live_xt, 2)
        _days = len(live_days)
        _part = True
        # 进行中月份不能拿去年整月比，否则同比虚低；取去年同一会计月「同样这几天」
        _prev = round(_ps['t'], 2) if _ps['t'] else None
    else:
        _rec = YRB.get(str(CUR_Y), {}).get(str(_m))
        if not _rec:
            continue
        _t = round(_rec['m_retail'], 2)
        _dseg, _gseg = round(_rec['m_diesel'], 2), round(_rec['m_gas'], 2)
        _x = round(_acc(['%d%02d' % (CUR_Y, _m)], [XIAO], 't'), 2)
        _days = NDAYS.get('%d%02d' % (CUR_Y, _m), 30)
        _part = False
        _p = YRB.get(str(PRE_Y), {}).get(str(_m))
        _prev = round(_p['m_retail'], 2) if (_p and _p['m_retail'] > 0) else None
    TREND.append({'k': '%d-%02d' % (CUR_Y, _m), 'm': _m, 'name': '%d月' % _m,
                  'range': _acct_range(_m), 't': _t, 'x': _x, 'ft': round(_t + _x, 2),
                  'd': _dseg, 'g': _gseg, 'days': _days,
                  'partial': _part, 'avg': round(_t / _days, 1) if _days else 0,
                  'prev': _prev,
                  'yoy': round((_t / _prev - 1) * 100, 1) if _prev else None})
# 分站级年度（按站切换用）：2023 只有公司台账、无分站明细，故不造数
yearly['station_annual'] = {
    _k: {'pair': _v.get('pair', ''), 'road': _v.get('road', ''),
         '2024': float(_v.get('2024') or 0),
         '2025': float(_v.get('2025') or 0),
         '2026': float(_v.get('2026') or 0)}
    for _k, _v in SA['stations'].items()}

# ============================================================================
# 视角数据层：L1 全公司（company） / 单站（= 一对，pair）
# ----------------------------------------------------------------------------
# 与公司级完全同源，不引入第二套算法：
#   1–10 会计月 = 「销售日报」当日销量（吨）日块逐站累加（SER['months']）
#   11 月       = kdocs 实时逐日（CUR['sale'][d]['stations']）
#   去年同期    = 去年 1–10 会计月整月 + 去年同会计月「同样已填的这几天」
# 金额：SER 里 amt/damt/gamt 是「元（含税）」，wan/dwan/gwan 是「万元（不含税）」，
#       统一折成万元后再汇总，避免再出现 1e4 倍的口径错位。
# 11 月分品类金额源表未拆，按当月柴/汽吨量占比折算，并在页脚注明。
# ============================================================================
PAIR_ST = dict(_PAIRS_ALL)          # 15 对 + 小额配送（虚拟对）

# 逐站逐日（吨/柴/汽）：支撑「近 30 天滚动窗口」与「去年同期同日」两条口径
ST_BYDATE = {}
for _ym, _dst in SER['daily'].items():
    _mk = '%s-%s' % (_ym[:4], _ym[4:])
    for _st, _rec in _dst.items():
        _box = ST_BYDATE.setdefault(_st, {})
        for _dn, _cell in _rec.items():
            if not _cell or not (_cell[0] or 0):
                continue
            try:
                _box[acct_date(_mk, int(_dn))] = _cell
            except (ValueError, TypeError):
                pass
for _d in cur_days:                               # 11 月实时并入同一张日历
    _dt = acct_date(LIVE_KEY, int(_d))
    for _st, _v in CUR['sale'][_d]['stations'].items():
        if (_v.get('t') or 0) <= 0:
            continue
        ST_BYDATE.setdefault(_st, {})[_dt] = [_v['t'], _v.get('diesel', 0), _v.get('gas', 0)]
    # 小额配送同样并入，否则「小额」这一实体的近 30 天曲线会缺当月
    _xr = CUR['sale'][_d].get('xiao') or {}
    if (_xr.get('t') or 0) > 0:
        ST_BYDATE.setdefault(XIAO, {})[_dt] = [_xr['t'], 0.0, 0.0]


def _acc(ym_list, sts, key):
    """若干个「年月」月表里，指定站集合的某字段合计（单位沿用源字段）"""
    tot = 0.0
    for _ym in ym_list:
        _mm = SER['months'].get(_ym) or {}
        for _st in sts:
            tot += float((_mm.get(_st) or {}).get(key) or 0)
    return tot


def _prev_days(sts, idx):
    """去年同会计月、与今年同样这几个日号的吨/柴/汽（idx: 0/1/2）"""
    _rec = SER['daily'].get(PREV_YM_LIVE) or {}
    tot = 0.0
    for _st in sts:
        _r = _rec.get(_st) or {}
        for _dn in LIVE_DN:
            _c = _r.get(str(_dn))
            if _c:
                tot += float(_c[idx] or 0)
    return tot


def _live(sts, key):
    """11 月实时：按日 × 站累加（key 取 t/diesel/gas/amt(元)/wan(万元)）

    小额配送不在 stations 里，走当日记录上的 xiao 小对象；
    它当月没有分品类拆分，柴/汽按 0 计，不臆造。
    """
    tot = 0.0
    for _d in cur_days:
        _rec = CUR['sale'][_d]
        _ss = _rec['stations']
        for _st in sts:
            if _st == XIAO:
                _x = _rec.get('xiao') or {}
                if key in _x:
                    tot += float(_x.get(key) or 0)
                continue
            _v = _ss.get(_st)
            if _v and _v.get(key):
                tot += float(_v[key])
    return tot


def _yoy(cur, prev):
    return round((cur / prev - 1) * 100, 1) if prev else None


def build_view(name, road, sts):
    """一个视角实体：公司级与单站级共用同一套字段与算法"""
    t = _acc(BASE_YM26, sts, 't') + _live(sts, 't')
    d = _acc(BASE_YM26, sts, 'd') + _live(sts, 'diesel')
    g = _acc(BASE_YM26, sts, 'g') + _live(sts, 'gas')
    a = (_acc(BASE_YM26, sts, 'amt') + _live(sts, 'amt')) / 1e4      # 含税万元
    w = _acc(BASE_YM26, sts, 'wan') + _live(sts, 'wan')              # 不含税万元
    # 11 月分品类金额源表未拆：按当月柴/汽吨量占比折算
    _lt = _live(sts, 't')
    _ld, _lg = _live(sts, 'diesel'), _live(sts, 'gas')
    _la = _live(sts, 'amt') / 1e4
    _rw = _live(sts, 'wan')
    da = _acc(BASE_YM26, sts, 'damt') / 1e4 + (_la * _ld / _lt if _lt else 0)
    ga = _acc(BASE_YM26, sts, 'gamt') / 1e4 + (_la * _lg / _lt if _lt else 0)
    dw = _acc(BASE_YM26, sts, 'dwan') + (_rw * _ld / _lt if _lt else 0)
    gw = _acc(BASE_YM26, sts, 'gwan') + (_rw * _lg / _lt if _lt else 0)

    pt = _acc(BASE_YM25, sts, 't') + _prev_days(sts, 0)
    pd_ = _acc(BASE_YM25, sts, 'd') + _prev_days(sts, 1)
    pg = _acc(BASE_YM25, sts, 'g') + _prev_days(sts, 2)
    # 去年同月只有整月金额，按「同期吨量占整月吨量」的比例折出同期金额（与分子分母同期限）
    _pm = SER['months'].get(PREV_YM_LIVE) or {}
    _pm_t = sum(float((_pm.get(s) or {}).get('t') or 0) for s in sts)
    _pm_a = sum(float((_pm.get(s) or {}).get('amt') or 0) for s in sts) / 1e4
    _pm_w = sum(float((_pm.get(s) or {}).get('wan') or 0) for s in sts)
    _rt = (_prev_days(sts, 0) / _pm_t) if _pm_t else 0
    pa = _acc(BASE_YM25, sts, 'amt') / 1e4 + _pm_a * _rt
    pw = _acc(BASE_YM25, sts, 'wan') + _pm_w * _rt

    # 近 60 天滚动窗口：跨会计月按真实日期拼接（前端按 WIN 30/45/60 现切）
    _bd, _start = {}, (LAST_DT - datetime.timedelta(days=59)) if LAST_DT else None
    if _start:
        for _st in sts:
            for _dt, _c in (ST_BYDATE.get(_st) or {}).items():
                if _start <= _dt <= LAST_DT:
                    _x = _bd.setdefault(_dt, [0.0, 0.0, 0.0])
                    _x[0] += _c[0] or 0; _x[1] += _c[1] or 0; _x[2] += _c[2] or 0
    d30 = [{'date': _dt.isoformat(), 'lab': _dt.strftime('%m-%d'),
            't': round(_bd[_dt][0], 1), 'diesel': round(_bd[_dt][1], 1),
            'gas': round(_bd[_dt][2], 1)} for _dt in sorted(_bd)]
    tot30 = sum(x['t'] for x in d30[-30:])
    on_days = len([x for x in d30[-30:] if x['t'] > 0])

    # 渠道结构（2026 年累计，11 张站级渠道表逐站累加；缺失的站跳过不造数）
    _cy = SER['chan'].get(str(CUR_Y)) or {}
    _cht = {}
    for _st in sts:
        for _c, _v in (_cy.get(_st) or {}).items():
            _cht[_c] = _cht.get(_c, 0.0) + float(_v or 0)
    _ctot = sum(_cht.values()) or 0
    chan = sorted([{'name': _k, 't': round(_v, 1),
                    'share': round(_v / _ctot * 100, 1) if _ctot else 0}
                   for _k, _v in _cht.items()], key=lambda x: -x['t'])

    # 逐会计月（1..11）：供轨迹图与「本月信息带」在单站视角下取数
    mo = {}
    for _m_ in range(1, LIVE_NO + 1):
        if _m_ == LIVE_NO:
            _t2, _d2, _g2 = _lt, _ld, _lg
            _a2, _w2, _days2 = _la, _rw, len(live_days)
            _p2 = _prev_days(sts, 0)
        else:
            _ym = '%d%02d' % (CUR_Y, _m_)
            _rec = (SER['months'].get(_ym) or {})
            _one = [_rec.get(s) or {} for s in sts]
            _t2 = sum(float(x.get('t') or 0) for x in _one)
            _d2 = sum(float(x.get('d') or 0) for x in _one)
            _g2 = sum(float(x.get('g') or 0) for x in _one)
            _a2 = sum(float(x.get('amt') or 0) for x in _one) / 1e4
            _w2 = sum(float(x.get('wan') or 0) for x in _one)
            _days2 = FS.get('%d%02d' % (CUR_Y, _m_), {}).get('n_days') or 30
            _p2 = _acc(['%d%02d' % (PRE_Y, _m_)], sts, 't')
        mo[str(_m_)] = {'t': round(_t2, 1), 'diesel': round(_d2, 1), 'gas': round(_g2, 1),
                        'amt': round(_a2, 1), 'wan': round(_w2, 1), 'days': _days2,
                        'price': round(_a2 * 1e4 / _t2) if _t2 else 0,
                        'price0': round(_w2 * 1e4 / _t2) if _t2 else 0,
                        'prev': round(_p2, 1) if _p2 else None,
                        'yoy': _yoy(_t2, _p2) if _p2 else None}

    return {
        'name': name, 'road': road, 'n_st': len(sts), 'stations': list(sts),
        't': round(t, 1), 'diesel': round(d, 1), 'gas': round(g, 1),
        'amt': round(a, 1), 'wan': round(w, 1),
        'damt': round(da, 1), 'gamt': round(ga, 1),
        'dwan': round(dw, 1), 'gwan': round(gw, 1),
        'price': round(a * 1e4 / t) if t else 0,
        'price0': round(w * 1e4 / t) if t else 0,
        'prev': {'t': round(pt, 1), 'diesel': round(pd_, 1), 'gas': round(pg, 1),
                 'amt': round(pa, 1), 'wan': round(pw, 1)},
        'yoy': {'t': _yoy(t, pt), 'diesel': _yoy(d, pd_), 'gas': _yoy(g, pg),
                'amt': _yoy(a, pa), 'wan': _yoy(w, pw)},
        'avg30': round(tot30 / 30, 1) if d30 else 0,
        'peak30': round(max([x['t'] for x in d30[-30:]]), 1) if d30 else 0,
        'on30': on_days, 'n30': min(30, len(d30)), 'daily': d30,
        'chan': chan, 'chan_t': round(_ctot, 1),
        'months': mo,
    }


ALL_ST = sorted({s for v in PAIR_ST.values() for s in v} | set(SA['stations'].keys()))
VIEW_CO = build_view('全公司', '', ALL_ST)
VIEW_PAIR = {}
for _pn, _sts in PAIR_ST.items():
    VIEW_PAIR[_pn] = build_view(_pn, PAIR_ROAD.get(_pn, ''), _sts)

# 占比与站内拆分：单站视角下「该对在公司中的位置」与「东西/南北两站对比」都要用
for _pn, _v in VIEW_PAIR.items():
    _v['share'] = round(_v['t'] / VIEW_CO['t'] * 100, 2) if VIEW_CO['t'] else 0
    _sub = []
    for _sn in _v['stations']:
        _one = SER['months']
        _st_t = _acc(BASE_YM26, [_sn], 't') + _live([_sn], 't')
        _st_d = _acc(BASE_YM26, [_sn], 'd') + _live([_sn], 'diesel')
        _st_g = _acc(BASE_YM26, [_sn], 'g') + _live([_sn], 'gas')
        _st_a = (_acc(BASE_YM26, [_sn], 'amt') + _live([_sn], 'amt')) / 1e4
        _st_w = _acc(BASE_YM26, [_sn], 'wan') + _live([_sn], 'wan')
        _pt = _acc(BASE_YM25, [_sn], 't') + _prev_days([_sn], 0)
        _sub.append({'name': _sn, 't': round(_st_t, 1), 'diesel': round(_st_d, 1),
                     'gas': round(_st_g, 1), 'amt': round(_st_a, 1), 'wan': round(_st_w, 1),
                     'share': round(_st_t / _v['t'] * 100, 1) if _v['t'] else 0,
                     'yoy': _yoy(_st_t, _pt)})
    _sub.sort(key=lambda x: -x['t'])
    _v['sub'] = _sub

VIEW_CO['share'] = 100.0
VIEW_CO['sub'] = sorted([{'name': _pn, 't': _v['t'], 'diesel': _v['diesel'],
                          'gas': _v['gas'], 'amt': _v['amt'], 'wan': _v['wan'],
                          'share': _v['share'], 'yoy': _v['yoy']['t']}
                         for _pn, _v in VIEW_PAIR.items()], key=lambda x: -x['t'])

# 渠道结构的公司基准线：单站视角下与该均值对照，才看得出「结构差异」
_co_chan = {c['name']: c['share'] for c in VIEW_CO['chan']}
VIEW_CO['chan_base'] = _co_chan
for _pn, _v in VIEW_PAIR.items():
    _v['chan_base'] = _co_chan

yearly['views'] = {
    'company': VIEW_CO,
    'pairs': VIEW_PAIR,
    'note': ('单站视角 = 一个服务区（东西 / 南北两站合计）。与全公司视角同源同算法：'
             '1–%d 会计月取日报日块逐站累加，%d 月取实时已填 %d 天。'
             '%d 月分品类产值按当月柴/汽吨量占比拆分；去年同期金额按同期吨量占整月比例折算。'
             % (max(int(x) for x in _base_mons), LIVE_NO, len(live_days), LIVE_NO)),
}

# ---- 路段结构：全年累计 + 逐会计月（供图表在「全年 / 各月」之间切换）----
# 口径与「站点排行」一致：只算 15 对实体站，小额配送没有路段归属，天然不含。
# 1–10 月取日报日块逐站累加，11 月取实时已填天数，与月度表同源。
_rd_months, _rd_days = {}, {}
for _m_ in range(1, LIVE_NO + 1):
    _per = {}
    for _pn, _rd in PAIR_ROAD.items():
        if _pn == XIAO or not _rd:
            continue
        _psts = _PAIRS_ALL.get(_pn) or []
        if _m_ == LIVE_NO:
            _v = _live(_psts, 't')
        else:
            _rec = SER['months'].get('%d%02d' % (CUR_Y, _m_)) or {}
            _v = sum(float((_rec.get(_s) or {}).get('t') or 0) for _s in _psts)
        if _v:
            _per[_rd] = _per.get(_rd, 0.0) + _v
    _rd_months[str(_m_)] = {k: round(v, 1) for k, v in _per.items()}
    _rd_days[str(_m_)] = (len(live_days) if _m_ == LIVE_NO
                          else (FS.get('%d%02d' % (CUR_Y, _m_), {}).get('n_days') or 30))

_rd_annual = {}
for _mo in _rd_months.values():
    for _k, _v in _mo.items():
        _rd_annual[_k] = _rd_annual.get(_k, 0.0) + _v
_rd_annual = {k: round(v, 1) for k, v in _rd_annual.items()}

yearly['roads'] = {
    'annual': _rd_annual,
    'months': _rd_months,
    'days': _rd_days,
    'order': sorted(_rd_annual, key=lambda k: -_rd_annual[k]),
    'note': '8 条高速路段 = 15 对实体站按所属路段归集，不含小额配送（小额无路段归属）',
}

payload = {'meta': {'generated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),                    'month_order': ['2026-11', '2026-10', '2026-09'],
                    'trend': TREND},
           'months': months, 'yearly': yearly,
           'daily_win': WINDOWS, 'last_date': LAST_DT.isoformat() if LAST_DT else ''}

# ---- 前一日各站销量（日走势图上方那条极窄 KPI 条） ----
# 源：实时月逐日逐站（CUR['sale'][日]['stations']），按「对」合并成 15 项，小额单列第 16 项。
# 取有数据的最后一天；小额在源表里是公司级、无站点，故独立成项（与站点排行口径一致）。
_pd_day = max(cur_days, key=lambda x: int(x)) if cur_days else None
_pd_items, _pd_x = [], 0.0
if _pd_day:
    _cell = CUR['sale'][_pd_day]['stations']
    for _pn, _sts in PAIR_ST.items():
        if _pn == XIAO:
            continue
        _v = sum(float((_cell.get(s) or {}).get('t') or 0) for s in _sts)
        if _v > 0:
            _pd_items.append({'name': _pn, 't': round(_v, 1)})
    _pd_items.sort(key=lambda x: -x['t'])
    _pd_x = round(float(CUR['sale'][_pd_day].get('xiao', {}).get('t', 0)), 1)

prev_day = {
    'day': _pd_day,
    'date': LAST_DT.isoformat() if LAST_DT else '',
    'label': LAST_DT.strftime('%m-%d') if LAST_DT else '',
    'items': _pd_items,
    'xiao': _pd_x,
    'total': round(sum(x['t'] for x in _pd_items) + _pd_x, 1),
}
payload['prev_day'] = prev_day

tpl = open(os.path.join(BASE, 'cockpit2_template.html'), encoding='utf-8').read()
html = tpl.replace('__PAYLOAD__', json.dumps(payload, ensure_ascii=False))
out = os.path.join(ROOT, 'cockpit', 'index.html')
open(out, 'w', encoding='utf-8').write(html)
print('written', out, len(html))
for k in ['2026-11', '2026-10', '2026-09']:
    m = months[k]
    print(' %s %-8s %d天 %8.1f吨 含税%7.0f万 站%d 渠道%d 客户%d'
          % (k, m['range'], m['days'], m['t'], m['amt'], len(m['stations']),
             len(m['channels']), len(m['clients'])))
print('历年年度:', yearly['annual'], '| 有数据月:', yearly['done_months'])
_V = yearly['views']
print('\n-- 视角对账 --')
print('公司视角: %.1f 吨 · 柴 %.1f / 汽 %.1f · 含税 %.1f 万元 · 同比 %s%%'
      % (_V['company']['t'], _V['company']['diesel'], _V['company']['gas'],
         _V['company']['amt'], _V['company']['yoy']['t']))
print('顶部 KPI : %.1f 吨 · 柴 %.1f / 汽 %.1f · 含税 %.1f 万元 · 同比 %s%%'
      % (disp_t, disp_d, disp_g, _amt_all, yoy_all))
print('差异      : 吨 %.2f | 柴 %.2f | 汽 %.2f | 金额 %.2f'
      % (_V['company']['t'] - disp_t, _V['company']['diesel'] - disp_d,
         _V['company']['gas'] - disp_g, _V['company']['amt'] - _amt_all))
_s = sum(v['t'] for v in _V['pairs'].values())
print('各实体合计: %.1f 吨（含小额，与公司差 %.2f）' % (_s, _s - _V['company']['t']))
print('含税金额  : 公司 %.1f vs 各对合计 %.1f'
      % (_V['company']['amt'], sum(v['amt'] for v in _V['pairs'].values())))
print('\n-- 单站视角 Top5 --')
for _pn, _v in sorted(_V['pairs'].items(), key=lambda x: -x[1]['t'])[:5]:
    print(' %-6s %-4s %8.1f 吨 · 柴 %6.1f / 汽 %6.1f 吨 · 含税 %7.1f 万 '
          '(柴 %6.1f / 汽 %6.1f) · 同比 %6s%% · 近30天日均 %5.1f 峰值 %5.1f 在营 %d/%d · 渠道 %d 项'
          % (_pn, _v['road'], _v['t'], _v['diesel'], _v['gas'], _v['amt'],
             _v['damt'], _v['gamt'], _v['yoy']['t'], _v['avg30'], _v['peak30'],
             _v['on30'], _v['n30'], len(_v['chan'])))

