# -*- coding: utf-8 -*-
"""校验 AirScript 推来的 `_incoming/flow.json` 能不能用。

★ 为什么要有这一道
  车流是云端自动推的，没人看着。万一哪天 AirScript 只读到一半站（比如某站文件权限变了、
  或者当天还没填报完），推上去的「前一日车流」就会缺站、合计变小，页面上的车次和单加
  直接被改错 —— 而且没人会发现，因为它看起来只是个数字。

  所以重建之前先验一遍：
    硬错误 → 不采用，沿用仓库里上一次的快照（页面照常出，只是车流停在旧日期）
    软警告 → 照用，但打印出来，方便事后追

  设计原则同全项目：**宁可停在旧值，也不把错数据摆上台面。**

用法：
  python verify_flow.py <候选 flow.json> <回退用 flow.json>
  退出码 0 = 可用（已就地写回候选文件，供后续 cp）；1 = 不可用（不修改任何文件）
"""
import datetime as dt
import json
import os
import sys

# 逐站自洽允许的浮点/取整误差（车次是整数，这里只兜底）
SUM_TOL = 1

# asof 相对今天允许落后的天数；超过只警告不拦（节假日、月初补填都会拖）
LAG_WARN = 3


def load(path):
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def check(d, today=None):
    """返回 (硬错误列表, 警告列表)"""
    errs, warns = [], []
    today = today or dt.date.today()

    # ---- 1. 顶层字段 ----
    asof = (d.get('asof') or '').strip()
    try:
        d_asof = dt.date.fromisoformat(asof)
    except Exception:
        errs.append('asof 缺失或不是 YYYY-MM-DD：%r' % (d.get('asof'),))
        d_asof = None

    st = d.get('stations') or {}
    if not isinstance(st, dict) or not st:
        errs.append('stations 为空或不是对象')
        return errs, warns

    if len(st) < 12:
        errs.append('站数只有 %d，少于 12 —— 大概率有站没读到（权限/未填报）' % len(st))

    # ---- 2. 逐站自洽 ----
    for nm, r in st.items():
        for k in ('iv', 'ig', 'id'):
            if not isinstance(r.get(k), (int, float)):
                errs.append('%s 的 %s 不是数字：%r' % (nm, k, r.get(k)))
        iv, ig, idv = r.get('iv'), r.get('ig'), r.get('id')
        if isinstance(iv, (int, float)) and isinstance(ig, (int, float)) and isinstance(idv, (int, float)):
            if abs(iv - (ig + idv)) > SUM_TOL:
                errs.append('%s 总车次 %s ≠ 汽油 %s + 柴油 %s' % (nm, iv, ig, idv))
            if iv <= 0:
                errs.append('%s 总车次为 0，不该出现在 payload 里' % nm)
        # 单加合理区间：汽油 10~200 升/次，柴油 30~1000 升/次
        for k, lo, hi, label in (('isp', 10, 200, '汽油'), ('idp', 30, 1000, '柴油')):
            v = r.get(k)
            if v is None:
                warns.append('%s 缺 %s（%s单加）' % (nm, k, label))
            elif not (lo <= v <= hi):
                warns.append('%s %s单加 %s 超出常见区间 %d~%d，留意是否源表异常'
                             % (nm, label, v, lo, hi))

    # ---- 3. 合计对账 ----
    ck = d.get('check') or {}
    vs = sum(r.get('iv') or 0 for r in st.values())
    gs = sum(r.get('ig') or 0 for r in st.values())
    ds = sum(r.get('id') or 0 for r in st.values())
    for key, mine, label in (('veh', vs, '总车次'), ('gasVeh', gs, '汽油车次'), ('dieVeh', ds, '柴油车次')):
        if key not in ck:
            warns.append('check 缺 %s' % key)
            continue
        if abs((ck.get(key) or 0) - mine) > SUM_TOL:
            errs.append('check.%s = %s 与逐站加总 %s 不符（%s）' % (key, ck.get(key), mine, label))

    # ---- 4. 时效 ----
    if d_asof:
        lag = (today - d_asof).days
        if lag < 0:
            errs.append('asof %s 在未来（今天是 %s）' % (asof, today))
        elif lag > LAG_WARN:
            warns.append('asof %s 距今已 %d 天（超过 %d 天），源表可能停止填报' % (asof, lag, LAG_WARN))

    return errs, warns


def main():
    if len(sys.argv) < 2:
        print('用法: python verify_flow.py <候选 flow.json> [回退用 flow.json]')
        return 2
    cand = sys.argv[1]
    fallback = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(cand):
        print('[车流校验] 没有 %s，跳过（沿用现有快照）' % cand)
        return 0

    try:
        d = load(cand)
    except Exception as e:
        print('[车流校验] ❌ 不是合法 JSON：%s —— 不采用' % e)
        if fallback and os.path.exists(fallback):
            print('           沿用 %s' % fallback)
        return 1

    errs, warns = check(d)

    print('[车流校验] asof=%s · 站数=%s · 总车次=%s'
          % (d.get('asof'), len(d.get('stations') or {}),
             (d.get('check') or {}).get('veh')))

    for w in warns:
        print('  ⚠️  %s' % w)

    if errs:
        print('[车流校验] ❌ 发现 %d 个硬错误，本次不采用：' % len(errs))
        for e in errs:
            print('     - %s' % e)
        if fallback and os.path.exists(fallback):
            fb = load(fallback)
            print('           沿用上一版快照 asof=%s（页面照常重建，车流停在旧日期）'
                  % fb.get('asof'))
        return 1

    print('[车流校验] ✅ 通过' + ('（有 %d 条警告）' % len(warns) if warns else ''))
    return 0


if __name__ == '__main__':
    sys.exit(main())
