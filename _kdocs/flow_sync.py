# -*- coding: utf-8 -*-
"""车流量 / 单加率 → `_flow.json`（供驾驶舱「前一日」窄条使用）

★ 口径（2026-10-29 按用户纠正重写，别改回去）
  车流与单加**不是** YTD 日均分摊，而是逐站取各站「非油销售日报V1.xlsx / 今日填报」
  (sheetId=2) 里**前一日那两行**的实报值（南/东 + 北/西 两行相加）。

  读表列位（0 基，表头在 row 0，数据从 row 1 起）：
    A=日期  B=区域  C=柴油车次  D=柴油升数  E=汽油车次  F=汽油升数
    行号公式：row = 1 + 2 * (日期 - 2025-12-26)
      2026-10-28 偏移 306 天 → 613(南/东) / 614(北/西)

  派生：
    iv  = ΣC + ΣE          前一日总车次
    ig  = ΣE               汽油车次
    id  = ΣC               柴油车次
    isp = ΣF / ΣE          汽油单加（升/次）
    idp = ΣD / ΣC          柴油单加（升/次）

  原始两行数值落在 `_flow_raw.json`（含 file_id、行号、列位、公司级对账值）。
  本脚本只做「原始 → 派生 + 自校验」，不含任何分摊、不写 ≈。

用法：
  python flow_sync.py                  # 用 _flow_raw.json 里的 asof
  python flow_sync.py --asof 2026-10-28
"""
import json
import os
import sys
import datetime as dt

BASE = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(BASE, '_flow_raw.json')
OUT = os.path.join(BASE, '_flow.json')

YEAR_START = dt.date(2025, 12, 26)   # 会计年起点（非自然年）


def row_of(d):
    """日期 → 「今日填报」里该日 南/东 所在行号（0 基）；北/西 为 +1"""
    return 1 + 2 * (d - YEAR_START).days


def main():
    if not os.path.exists(RAW):
        raise SystemExit('缺少 _flow_raw.json（车流真值快照），先取数再跑本脚本')

    raw = json.load(open(RAW, encoding='utf-8'))
    asof_s = raw.get('asof') or ''
    if '--asof' in sys.argv:
        asof_s = sys.argv[sys.argv.index('--asof') + 1]
    try:
        asof = dt.date.fromisoformat(asof_s)
    except ValueError:
        raise SystemExit('_flow_raw.json 的 asof 非法：%r' % asof_s)

    # 行号自校验：快照里记录的行号必须和公式算出来的一致，否则快照与日期脱节
    r0 = row_of(asof)
    used = raw.get('rows_used') or [r0, r0 + 1]
    if list(used) != [r0, r0 + 1]:
        print('[warn] rows_used=%s 与公式算出的 [%d,%d] 不一致，以 asof 为准'
              % (used, r0, r0 + 1))

    st, tot = {}, {'veh': 0, 'gv': 0, 'dv': 0, 'gl': 0.0, 'dl': 0.0}
    for nm, rec in (raw.get('pairs') or {}).items():
        rows = rec.get('rows') or []
        if len(rows) != 2:
            print('[warn] %s 应有 2 行（南/东 + 北/西），实际 %d 行，跳过' % (nm, len(rows)))
            continue
        dv = sum(float(r[0] or 0) for r in rows)      # 柴油车次
        dl = sum(float(r[1] or 0) for r in rows)      # 柴油升
        gv = sum(float(r[2] or 0) for r in rows)      # 汽油车次
        gl = sum(float(r[3] or 0) for r in rows)      # 汽油升
        st[nm] = {
            'iv': int(round(dv + gv)),                 # 前一日总车次
            'ig': int(round(gv)),                      # 汽油车次
            'id': int(round(dv)),                      # 柴油车次
            'isp': round(gl / gv, 2) if gv else None,  # 汽油单加 升/次
            'idp': round(dl / dv, 2) if dv else None,  # 柴油单加 升/次
            'il': round(dl + gl, 1),                   # 总升
            'gl': round(gl, 1),
            'dl': round(dl, 1),
        }
        tot['veh'] += st[nm]['iv']; tot['gv'] += st[nm]['ig']; tot['dv'] += st[nm]['id']
        tot['gl'] += gl; tot['dl'] += dl

    chk = {'veh': tot['veh'], 'gasVeh': tot['gv'], 'dieVeh': tot['dv'],
           'gasPer': round(tot['gl'] / tot['gv'], 2) if tot['gv'] else None,
           'diePer': round(tot['dl'] / tot['dv'], 2) if tot['dv'] else None}

    out = {
        'asof': asof.isoformat(),
        'basis': '前一日实报',
        'note': '车流/单加逐站取各站非油 V1 报表「今日填报」%s 两行（南/东 + 北/西）实报值，'
                '非日均分摊、非估值' % asof.isoformat(),
        'source': '非油各站 V1 报表 / 今日填报 sheetId=2（列 C 柴油车次 D 柴油升 E 汽油车次 F 汽油升）',
        'rows': [r0, r0 + 1],
        'check': chk,
        'pairs': len(st),
        'stations': st,
    }
    open(OUT, 'w', encoding='utf-8').write(json.dumps(out, ensure_ascii=False, indent=1))
    print('written', OUT)
    print(' asof %s  行 %d-%d  %d 对' % (asof, r0, r0 + 1, len(st)))
    print(' 合计 %d 车次（汽 %d / 柴 %d） · 单加 汽 %.2f / 柴 %.2f 升'
          % (chk['veh'], chk['gasVeh'], chk['dieVeh'], chk['gasPer'] or 0, chk['diePer'] or 0))

    # ---- 公司级对账：与非油「单数统计」(sheetId=12) 按日区块比 ----
    exp = ((raw.get('_verify') or {}).get('公司级单数统计(非油 sheetId=12 按日)') or {}).get(asof_s)
    if exp:
        pairs = [('veh', '总车次'), ('gasVeh', '汽油车次'), ('dieVeh', '柴油车次'),
                 ('gasPer', '汽油单加'), ('diePer', '柴油单加')]
        bad = []
        print(' -- 对账 vs 非油单数统计（公司级按日） --')
        for k, lab in pairs:
            a, b = chk.get(k), exp.get(lab)
            ok = (a is not None and b is not None and abs(float(a) - float(b)) < 0.005)
            print('   %-8s 逐站合计 %-10s 公司级 %-10s %s'
                  % (lab, a, b, 'OK' if ok else '★不一致'))
            if not ok:
                bad.append(lab)
        if bad:
            print(' ★ 有 %d 项不一致：%s（先查快照是否漏站/串行）' % (len(bad), '/'.join(bad)))
        else:
            print(' 全部 %d 项吻合 —— 15 对逐站加总 = 公司级当日总量' % len(pairs))
    else:
        print(' （无公司级对账值，跳过校验）')


if __name__ == '__main__':
    main()
