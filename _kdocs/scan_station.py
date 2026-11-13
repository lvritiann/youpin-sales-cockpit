# -*- coding: utf-8 -*-
"""逐站 × 逐月 × 分品类（含金额）· 逐站逐日 · 逐站渠道结构

为什么单独写一个：`scan_folder.py` 只累到「站 × 月」的总量，没有分品类金额、
没有逐站逐日、也没有渠道拆分，单站视角下「柴油/汽油产值」「本站日销量」
「消费结构」都拿不到。这里重扫 34 个日报，一次遍历产出三层：

  months[ym][站] = {t,d,g, amt,wan, damt,gamt, dwan,gwan}
  daily [ym][站][日号] = [t,d,g]            （2025-11 起，支撑近 30 天 + 去年同期）
  chan  [年][站][渠道] = 吨                 （消费结构，按量）

★ 「小额配送」不参与站级归属（源表里它是 A 列独立一行、B 列为空），
  但自带 -10/0#/92#/95# 分品类吨与金额，所以按**虚拟站**收进同一套结构，
  命名固定为 XIAO，让它在驾驶舱里与 15 对平权：可下钻、可比同比、可拆柴汽。
  唯一缺项是渠道：11 张渠道表里没有小额行，故 chan 里不会出现它（不造数）。

列位（0 基，2024/2025/2026 三年表头已实测一致）：
  「销售日报」2..5 分品类吨 6 小计 | 7..10 升 11 小计
              12..15 分品类销售额(元) 16 小计 | 17..20 分品类不含税(万元) 21 小计
  柴油 = -10 + 0#，汽油 = 92# + 95#
  「现金/POS/微信/…」11 张渠道表同构：0 路段 1 站名 2..5 分品类吨 6 小计

只读，绝不写回原表。
"""
import openpyxl, os, glob, json, re

ROOT = r'D:\成品油销量'
BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(os.path.dirname(BASE), '_archive', 'station_series.json')

XIAO = '小额配送'          # 虚拟站名，全站唯一，勿与真实站名冲突
AGG = ('小计', '合计', '零售小计', '小额配送', '总计', '累计')

# ---- 站名归一：同一实体在源表里被写成两个名字 ------------------------------
# ★ 键 = 渠道表的写法，值 = 「销售日报」的写法。
#   为什么以销售日报为准：「销售日报」是 months / daily 的键，也是驾驶舱里
#   站级下拉与「对」级配对（stations_annual.json 的 pair_stations）用的名字。
#   渠道表若保留自己的叫法，单站视角按「销售日报」名去查 chan 就查不到 ——
#   页面上「消费结构」整块空白，而数据其实一直都在。
#   2026-11-13 定位：天汉水城同一服务区，销售日报写「西 / 东」、10 张渠道表写
#   「南 / 北」，两边站名集合互不相交。量级配对已核实（2026 累计）：
#       渠道表 南 5,036.5 吨 ≡ 销售日报 西 5,032.8 吨
#       渠道表 北 4,837.8 吨 ≡ 销售日报 东 4,834.5 吨
#   差 0.07% 是两表吨量折算差，不是两个站。
# ★ 只放已用量级对上的等价关系。拿不准的宁可留着让下面的审计报出来，也不猜。
STN_ALIAS = {
    '天汉水城南': '天汉水城西',
    '天汉水城北': '天汉水城东',
}


def stn_name(s):
    """把站名收敛到「销售日报」的写法。两张表都过一遍，改名方向改了也不怕。"""
    s = s.strip()
    return STN_ALIAS.get(s, s)
# 逐站逐日要支撑两条曲线：本站近 30 天、去年同一会计月的同期天数
DAILY_KEEP_FROM = '202511'
# 渠道结构只回溯到 2025（2024 年结构对比价值低，省一半解析时间）
CHAN_FROM = '2025'
# 11 张零售渠道表（「工贸/贸易/成品油贸易/LNG贸易」是公司级贸易口径，不是站级，排除）
CHAN_SHEETS = ('现金', 'POS', '微信', '支付宝', '工贸卡', '壳牌卡', '大客户',
               '通四海', '定期结算', '回罐', '内结')

C_T10, C_T0, C_T92, C_T95, C_T = 2, 3, 4, 5, 6
C_A10, C_A0, C_A92, C_A95, C_AMT = 12, 13, 14, 15, 16
C_N10, C_N0, C_N92, C_N95, C_WAN = 17, 18, 19, 20, 21


def num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', '').replace(' ', '').strip())
    except Exception:
        return 0.0


def is_road(v):
    return isinstance(v, str) and v.strip() and v.strip() not in AGG \
        and not re.fullmatch(r'\d+', v.strip())


def day_blocks(rows, col=0, key='当日销量（吨）'):
    """按「当日销量（吨）」表头切日块，返回 [(日号, [行...]), ...]

    ★ 渠道表的日号是文本 '26'、而末块是 '26-25' 累计块，两者都是 str，
      只认 1~2 位纯数字，把累计块排除掉，否则渠道吨数会被重复计一遍。
    """
    starts = [i for i, r in enumerate(rows)
              if any(isinstance(c, str) and key in c for c in r)]
    out = []
    for k, i in enumerate(starts):
        e = starts[k + 1] if k + 1 < len(starts) else len(rows)
        lab = rows[i][col]
        if isinstance(lab, (int, float)):
            d = int(lab)
        elif isinstance(lab, str) and re.fullmatch(r'\d{1,2}', lab.strip()):
            d = int(lab.strip())
        else:
            continue                          # '26-25' 累计块
        out.append((d, rows[i + 1:e]))
    return out


def scan_one(path):
    ym = os.path.basename(path)[:6]
    year = ym[:4]
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    mon, day_map = {}, {}

    # ---- 销售日报：站 × 月（分品类金额）+ 站 × 日 ----
    if '销售日报' in wb.sheetnames:
        rows = [list(r) for r in wb['销售日报'].iter_rows(max_col=22, values_only=True)]
        for d, blk in day_blocks(rows):
            for r in blk:
                if len(r) < 22:
                    continue
                a, b = r[0], r[1]
                alab = a.strip() if isinstance(a, str) else ''
                if alab == XIAO:
                    # 公司级独立一行：A 列标记、B 列为空 → 收成虚拟站，走同一套字段
                    name = XIAO
                else:
                    if alab in ('零售小计', '合计'):
                        continue
                    if not isinstance(b, str) or not b.strip() or b.strip() in AGG:
                        continue
                    if re.fullmatch(r'\d+', b.strip()):
                        continue
                    name = stn_name(b)
                t = num(r[C_T])
                if t <= 0:
                    continue
                dt = num(r[C_T10]) + num(r[C_T0])
                dg = num(r[C_T92]) + num(r[C_T95])
                m = mon.setdefault(name, {'t': 0.0, 'd': 0.0, 'g': 0.0,
                                          'amt': 0.0, 'wan': 0.0,
                                          'damt': 0.0, 'gamt': 0.0,
                                          'dwan': 0.0, 'gwan': 0.0, 'days': 0})
                m['t'] += t; m['d'] += dt; m['g'] += dg
                m['amt'] += num(r[C_AMT]); m['wan'] += num(r[C_WAN])
                m['damt'] += num(r[C_A10]) + num(r[C_A0])
                m['gamt'] += num(r[C_A92]) + num(r[C_A95])
                m['dwan'] += num(r[C_N10]) + num(r[C_N0])
                m['gwan'] += num(r[C_N92]) + num(r[C_N95])
                m['days'] += 1
                if ym >= DAILY_KEEP_FROM:
                    day_map.setdefault(name, {})[str(d)] = [round(t, 2),
                                                            round(dt, 2), round(dg, 2)]

    # ---- 11 张渠道表：站 × 年 × 渠道（吨） ----
    chan = {}
    if year >= CHAN_FROM:
        for sn in CHAN_SHEETS:
            if sn not in wb.sheetnames:
                continue
            rows = [list(r) for r in wb[sn].iter_rows(max_col=7, values_only=True)]
            for d, blk in day_blocks(rows):
                for r in blk:
                    if len(r) < 7:
                        continue
                    b = r[1]
                    if not isinstance(b, str) or not b.strip() or b.strip() in AGG:
                        continue
                    if re.fullmatch(r'\d+', b.strip()):
                        continue
                    t = num(r[C_T])
                    if t <= 0:
                        continue
                    c = chan.setdefault(stn_name(b), {})
                    c[sn] = round(c.get(sn, 0.0) + t, 3)
    wb.close()

    def rnd(x):
        return {k: (round(v, 3) if isinstance(v, float) else v) for k, v in x.items()}
    return ym, {k: rnd(v) for k, v in mon.items()}, day_map, chan


if __name__ == '__main__':
    files = []
    for y in ('日报2024', '日报2025', '日报2026'):
        files += sorted(glob.glob(os.path.join(ROOT, y, '20*销售及库存日报*.xlsx')))
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    print('待处理 %d 个文件' % len(files))

    MON, DAY, CHAN = {}, {}, {}
    for i, f in enumerate(files, 1):
        try:
            ym, mon, day, chan = scan_one(f)
        except Exception as ex:
            print('  !! %s 失败: %s' % (os.path.basename(f), ex))
            continue
        MON[ym] = mon
        if day:
            DAY[ym] = day
        for st, rec in chan.items():
            tgt = CHAN.setdefault(ym[:4], {}).setdefault(st, {})
            for k, v in rec.items():
                tgt[k] = round(tgt.get(k, 0.0) + v, 3)
        print('[%2d/%d] %s 站%2d 零售%9.1f 吨 柴油额%9.0f 万 汽油额%9.0f 万 渠道站%2d'
              % (i, len(files), ym, len(mon),
                 sum(v['t'] for v in mon.values()),
                 sum(v['damt'] for v in mon.values()) / 1e4,
                 sum(v['gamt'] for v in mon.values()) / 1e4, len(chan)))

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump({'months': MON, 'daily': DAY, 'chan': CHAN},
              open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
    print('\n写出 %s  (%.1f KB)' % (os.path.abspath(OUT), os.path.getsize(OUT) / 1024))

    t26 = sum(v['t'] for ym, d in MON.items() if ym.startswith('2026') for v in d.values())
    a26 = sum(v['amt'] for ym, d in MON.items() if ym.startswith('2026') for v in d.values())
    dd = sum(v['damt'] for ym, d in MON.items() if ym.startswith('2026') for v in d.values())
    gg = sum(v['gamt'] for ym, d in MON.items() if ym.startswith('2026') for v in d.values())
    print('2026 1-10月：全口径 %.1f 吨 | 金额 %.0f 元 = 柴 %.0f + 汽 %.0f（差 %.2f）'
          % (t26, a26, dd, gg, a26 - dd - gg))
    # 小额配送虚拟站对账：与「月销售累计」的零售/小额比例应能互证
    for _y in ('2026', '2025'):
        _x = sum(v['t'] for ym, d in MON.items() if ym.startswith(_y)
                 for k, v in d.items() if k == XIAO)
        _r = sum(v['t'] for ym, d in MON.items() if ym.startswith(_y)
                 for k, v in d.items() if k != XIAO)
        print('  %s：零售 %9.1f + 小额 %8.1f = 全口径 %9.1f 吨（小额占 %.1f%%）'
              % (_y, _r, _x, _r + _x, _x / (_r + _x) * 100 if (_r + _x) else 0))
    for y in sorted(CHAN):
        tot = sum(sum(v.values()) for v in CHAN[y].values())
        ret = sum(v['t'] for ym, d in MON.items() if ym.startswith(y)
                  for k, v in d.items() if k != XIAO)
        print('  %s 渠道合计 %.1f 吨（对比零售 %s，小额 %s 不计入渠道）'
              % (y, tot, round(ret, 1),
                 round(sum(v['t'] for ym, d in MON.items() if ym.startswith(y)
                           for k, v in d.items() if k == XIAO), 1)))

    # ---- 站名审计：渠道表有、销售日报没有的站名 → 该站消费结构必然空白 ----
    # 这类问题在页面上表现为「某个站没有消费结构」，很容易被当成算错。
    # 一旦报出来，就往上面 STN_ALIAS 里补一条（先用吨量级配对确认是同一实体）。
    print('\n---- 站名审计 ----')
    for y in sorted(CHAN):
        ms = {k for ym, d in MON.items() if ym.startswith(y) for k in d}
        extra = sorted(set(CHAN[y]) - ms)
        if extra:
            print('  [!!] %s 渠道表有 %d 个站名在销售日报里不存在：%s'
                  % (y, len(extra), extra))
            print('       → 单站视角下这些站的「消费结构」会是空的')
        else:
            print('  [ok] %s 渠道表站名 %d 个，全部能在销售日报里找到'
                  % (y, len(CHAN[y])))
