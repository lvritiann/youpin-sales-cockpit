# -*- coding: utf-8 -*-
"""三年日报「大客户明细」全量抽取 → _archive/cust_hist.json

为什么要有这份脚本
------------------
线上 WPS AirScript 只推「当前会计月」，驾驶舱 payload 里因此只有 9/10/11 三个月
带客户明细。用户明确指出三年日报原始文件都已提供，要求「全年 / 每个月」筛选能
看到真实历史客户构成 —— 不能拿三个月冒充全年。

源表结构（2024-01 ~ 2026-10 共 34 份，逐份实测完全一致）
------------------------------------------------------
「大客户明细」1546 行 × 162 列：
  · 第 0 行是客户名合并表头，每 10 列一个「客户槽位」，起点 col2，共 16 槽
  · 槽 0 (col2)  = 「大客户合计」，其后 15 槽是具体客户
  · 每槽前 5 列 = 当日销量（升）的 -10 / 0 / 92 / 95 / 小计  → 小计在 j+4
  · 每槽后 5 列 = 当日销售额（元）的同四栏 / 小计          → 小计在 j+9
  · 第 1 行起每 48 行是一个日块（A 列 = 日号 26,27,...,25），
    最后一块 A 列是「26-25」累计块，必须丢弃
  · B 列是站名；A 列是路段（合并单元格，只有首行有值）

★ 空表头槽位不能丢：实测 202601 有 2 个空表头槽位共 237,702.7 升、
  202511 有 82,383.9 升、202403 有 135,093.9 升。WPS 里跑的 v2.6 用硬编码
  CLIENT_COLS 客户名，这部分量整个漏掉，正是「大客户构成」分类不对的根因。
  这里按 v2.9 的口径统一归入「其他」。

★ 逐月对账：合计槽升数 − Σ15 个客户槽升数 必须 ≈ 0（实测 34 份全部为 0），
  任何一份不为 0 都会在输出里标 bad，build 侧据此在页面上打警示。

用法
----
  python _kdocs/cust_hist_extract.py                 # 全量 34 份
  ONLY=202601,202511 python _kdocs/cust_hist_extract.py   # 只跑指定月
  LIMIT=2 python _kdocs/cust_hist_extract.py        # 试跑前 2 份
只读原始 xlsx，不改动任何日报文件。
"""
import os
import re
import io
import json
import datetime
import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
ARCH = os.path.join(ROOT, '_archive')
SRC_ROOT = r'D:\成品油销量'

# 与 build_cockpit2.py / read_daily.js 完全相同的折算系数（升 → 吨）
COEF = 0.83 / 1000.0
STEP = 10          # 每个客户槽位 10 列
AGGC = 2           # 「大客户合计」槽位起始列
CHANNELS = ['现金', 'POS', '微信', '支付宝', '工贸卡', '壳牌卡',
            '大客户', '定期结算', '回罐', '内结']
AGG_WORDS = ('小计', '合计', '零售小计', '小额配送', '总计', '大客户合计')


def num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(',', '').strip()
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def is_station(v):
    if v is None:
        return False
    s = str(v).strip()
    if not s or s in AGG_WORDS or '当日' in s or re.fullmatch(r'\d+(\.\d+)?', s):
        return False
    return True


def day_blocks(rs, key):
    """日块起点行号。key = '当日销量（升）'（客户明细）/ '当日销量（吨）'（渠道表）"""
    return [i for i, r in enumerate(rs)
            if any(isinstance(c, str) and key in c for c in r)]


def iter_station_rows(rs, starts):
    """产出所有「有效日块 × 站行」，累计块（A 列含 '-'）整块丢弃。"""
    for k, i in enumerate(starts):
        e = starts[k + 1] if k + 1 < len(starts) else len(rs)
        lab = str(rs[i][0]).strip() if rs[i] and rs[i][0] is not None else ''
        if '-' in lab or not re.fullmatch(r'\d{1,2}', lab):
            continue                       # 26-25 累计块 / 异常表头
        yield lab, rs[i:e]


def sheet_rows(wb, name):
    if name not in wb.sheetnames:
        return None
    return [list(r) for r in wb[name].iter_rows(values_only=True)]


def extract(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {'sheets': len(wb.sheetnames), 'warnings': []}
    try:
        # ---------- 大客户明细 ----------
        rs = sheet_rows(wb, '大客户明细')
        if rs is None:
            out['warnings'].append('无「大客户明细」表')
            return out
        W = max(len(r) for r in rs)
        slots = []
        for j in range(AGGC, W - 9, STEP):
            c = rs[0][j] if j < len(rs[0]) else None
            slots.append([j, str(c).strip() if c is not None else ''])
        if not slots or slots[0][1] != '大客户合计':
            out['warnings'].append('槽位表头异常：%s' % (slots[:2],))
            return out
        starts = day_blocks(rs, '当日销量（升）')
        agg_l = agg_a = 0.0
        per = {}
        unnamed = {'l': 0.0, 'amt': 0.0, 'slots': set()}
        days = set()
        for lab, blk in iter_station_rows(rs, starts):
            for r in blk:
                if len(r) < 12 or not is_station(r[1]):
                    continue
                agg_l += num(r[AGGC + 4])
                agg_a += num(r[AGGC + 9])
                days.add(lab)
                for j, cn in slots[1:]:
                    if j + 9 >= len(r):
                        continue
                    l_, a_ = num(r[j + 4]), num(r[j + 9])
                    if not (l_ or a_):
                        continue
                    if not cn:                       # ★ 空表头但有量 → 其他
                        unnamed['l'] += l_
                        unnamed['amt'] += a_
                        # 记「哪几个槽位」而不是累加次数：同一槽位跨 31 天 × 30 站
                        # 会出现上千次，报成「2333 个槽位」纯属误导
                        unnamed['slots'].add(j)
                        cn = '其他'
                    x = per.setdefault(cn, {'l': 0.0, 'amt': 0.0})
                    x['l'] += l_
                    x['amt'] += a_
        sum_l = sum(v['l'] for v in per.values())
        sum_a = sum(v['amt'] for v in per.values())
        # ★ 出数以「各客户槽之和」为准，不用源表「大客户合计」槽。
        #   2026-02 的 29 日块「华山南」合计槽漏填 0，客户槽实有 59,725.9 升，
        #   合计槽比客户槽少 49.6 吨 —— 拿合计槽出数会凭空少一个站。
        #   gap 只作「源表合计槽是否可信」的信号，前端据此打警示，不参与出数。
        gap = agg_l - sum_l
        clients = []
        for cn, v in per.items():
            t = round(v['l'] * COEF, 1)
            clients.append({'name': cn, 'l': round(v['l'], 1),
                            'amt': round(v['amt'] / 1e4, 2), 't': t})
        clients.sort(key=lambda x: -x['l'])
        tot_l = sum(x['l'] for x in clients) or 1
        for c in clients:
            c['share'] = round(c['l'] / tot_l * 100, 1)
        out.update({
            'slots': [c for _, c in slots],
            'days': len(days),
            'clients': clients,
            'big_l': round(sum_l, 1),                # 客户槽之和（出数用）
            'big_a': round(sum_a / 1e4, 2),          # 万元（含税）
            'big_t': round(sum_l * COEF, 1),         # 吨
            'src_agg_l': round(agg_l, 1),            # 源表「大客户合计」槽，仅供对账
            'src_agg_a': round(agg_a / 1e4, 2),
            'unnamed_l': round(unnamed['l'], 1),
            'unnamed_t': round(unnamed['l'] * COEF, 1),
            'unnamed_n': len(unnamed['slots']),
            'check_gap_l': round(gap, 2),
            'bad': abs(gap) > 1.0,
        })

        # ---------- 渠道表（副图「大客户渠道在全口径中的位置」）----------
        ch = {}
        for c in CHANNELS:
            crs = sheet_rows(wb, c)
            if crs is None:
                continue
            cst = day_blocks(crs, '当日销量（吨）')
            t = 0.0
            for _lab, blk in iter_station_rows(crs, cst):
                for r in blk:
                    if len(r) < 7 or not is_station(r[1]):
                        continue
                    t += num(r[6])                   # 当日销量（吨）·小计
            if t > 0:
                ch[c] = round(t, 2)
        out['channels'] = ch
    finally:
        wb.close()
    return out


def main():
    files = []
    for dirpath, _dirs, names in os.walk(SRC_ROOT):
        for fn in names:
            if fn.startswith('~$') or not fn.lower().endswith('.xlsx'):
                continue
            m = re.match(r'^(\d{6})', fn)
            if not m or '销售及库存日报' not in fn:
                continue
            tag = m.group(1)
            lab = '%s-%s' % (tag[:4], tag[4:])
            files.append((lab, tag, os.path.join(dirpath, fn)))
    files.sort()
    only = [x.strip() for x in os.environ.get('ONLY', '').split(',') if x.strip()]
    if only:
        files = [f for f in files if f[1] in only]
    lim = int(os.environ.get('LIMIT', '0') or 0)
    if lim:
        files = files[:lim]
    print('待抽取 %d 份：%s ... %s' % (len(files), files[0][0] if files else '-',
                                      files[-1][0] if files else '-'))

    prev = {}
    if os.environ.get('MERGE'):
        _p = os.path.join(ARCH, 'cust_hist.json')
        if os.path.exists(_p):
            prev = (json.load(io.open(_p, encoding='utf-8')).get('months') or {})
            print('合并模式：已有 %d 个月，本次只替换 %d 个' % (len(prev), len(files)))
    months, bad = {}, []
    for idx, (lab, tag, path) in enumerate(files, 1):
        try:
            r = extract(path)
        except Exception as e:
            print('[%02d/%02d] %s 读取失败：%s' % (idx, len(files), lab, e))
            bad.append((lab, str(e)))
            continue
        r.update({'key': lab, 'tag': tag, 'file': os.path.basename(path),
                  'src': os.path.basename(os.path.dirname(path))})
        months[lab] = r
        print('[%02d/%02d] %s %d天 槽%d 客户%d 大客户 %.1f吨 未署名 %.1f吨 渠道%d 对账差 %.2f升%s'
              % (idx, len(files), lab, r.get('days', 0), len(r.get('slots', [])),
                 len(r.get('clients', [])), r.get('big_t', 0), r.get('unnamed_t', 0),
                 len(r.get('channels', {})), r.get('check_gap_l', 0),
                 '  ★BAD' if r.get('bad') else ''))
        if r.get('bad') or r.get('warnings'):
            bad.append((lab, 'gap=%.2f %s' % (r.get('check_gap_l', 0), r.get('warnings'))))

    if prev:
        months = dict(prev, **months)
        bad = [(k, 'gap=%.2f %s' % (v.get('check_gap_l', 0), v.get('warnings') or ''))
               for k, v in sorted(months.items()) if v.get('bad') or v.get('warnings')]
    doc = {
        'generated': datetime.datetime.now(
            datetime.timezone(datetime.timedelta(hours=8))).strftime('%Y-%m-%d %H:%M'),
        'source': r'D:\成品油销量\日报20xx\YYYYMM销售及库存日报（新).xlsx · 大客户明细',
        'coef': COEF,
        'note': ('会计月标签 = 文件名 YYYYMM（26 日→次月 25 日，标签取止月）。'
                 '客户吨 = 升 × 0.83 / 1000，与实时月 pack_from_payload 同一系数。'
                 '空表头但有用的槽位归入「其他」。'),
        'months': months,
        'bad_months': bad,
    }
    os.makedirs(ARCH, exist_ok=True)
    dst = os.path.join(ARCH, 'cust_hist.json')
    with io.open(dst, 'w', encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False, indent=1)
    print('\n写出 %s（%d 个月，异常 %d）' % (dst, len(months), len(bad)))
    if bad:
        for x in bad:
            print('  ⚠', x)


if __name__ == '__main__':
    main()
