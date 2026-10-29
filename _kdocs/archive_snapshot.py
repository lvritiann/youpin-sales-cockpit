# -*- coding: utf-8 -*-
"""data.json 快照归档 —— 让「月底清表」永久不再需要人工下载 xlsx

## 为什么需要它

填报簿每月 26 号清空重填。清空之后，上个月的原始数据在金山文档里就没了，
过去唯一的保命手段是人工把 xlsx 下载到 D:\\成品油销量\\日报20xx\\ ——
一旦某个月忘了下载，那个月就永久断档，驾驶舱的历史直接缺口。

但事实是：`data.json` 是**当前会计月的完整快照**，含
    sale（逐日逐站：吨/柴汽四段/升/金额/柴汽升金额）
    channels（10 类渠道）
    bigcust（15 个客户槽位）
    trade
每个月 25 号那一次的 data.json，就已经包含了整月的全部内容。
而 AirScript 用 Contents API 推送，**每一次推送都在 git 历史里留下一个 commit** ——
也就是说这些快照本来就是全的，只是从来没被固化成文件。

本脚本做三件事：
  1) 每月一份 `_archive/monthly/<会计月>.json` —— **每次覆盖写**。
     翻月后这个文件自然停止更新，永久冻结在 25 号那次的完整态。
  2) `_archive/last_good.json` —— 只保留**一份**「天数最多」的快照。
     作用是防坏数据覆盖：某次推送若读表不全（天数变少），monthly 会被挡住，
     last_good 也不会退。真出了事，这一份就是最后的好数据。
  3) 维护 `_archive/months_all.json`（月索引）与 `_archive/HEALTH.json`（健康状态）。

## 为什么不存「每日快照」

算过账：一份整月 data.json 约 730 KB，一天一份存 400 天就是 290 MB，
把仓库撑成仓库瘤。而**每天的完整快照本来就在 git 历史里** ——
AirScript 用 Contents API 每次推送都产生一个 commit，`_incoming/data.json` 的每个版本都在。

要回溯任意一天，用 git 就够了，不占额外空间：
    git log --oneline -- _incoming/data.json
    git show <commit>:_incoming/data.json > day.json

所以这里只固化「整月」（核心资产）和「最后一份好数据」（保险），不重复存 git 已有的东西。

## 为什么用「每次覆盖」而不是「翻月时搬一次」

翻月搬一次要判断"现在是不是刚翻月"，一旦那天 Actions 挂了判断就错、搬移就漏。
"每次覆盖当前月"是幂等的：跑 1 次和跑 100 次结果一样，翻月后旧文件自然不再被碰。
**幂等 > 聪明**，这是无人值守方案的第一原则。

## 完整性判定（三态）

    complete   实际天数 >= 应有天数           → 这个月齐了
    running    未到期，天数不足是正常的        → 月中本来就没走完
    incomplete 已过期末，天数仍然不足          → ★ 真丢数据了，要告警

第三态是关键：它会在**翻月之后**回头把「上个月没走完」的那些月标出来，
不必等人去发现。

## 退出码策略

**永远 exit 0**。归档失败绝不能阻断驾驶舱重建 —— 页面能出比归档完整更重要。
失败信息写进 HEALTH.json 的 alert 字段，下次跑会再试（因为覆盖式，天然可重试）。
"""
import json
import os
import sys
import hashlib
import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent        # 仓库根
INCOMING = ROOT / '_incoming' / 'data.json'
ARCH = ROOT / '_archive'
MONTHLY = ARCH / 'monthly'
LAST_GOOD = ARCH / 'last_good.json'
MONTHS_ALL = ARCH / 'months_all.json'
HEALTH = ARCH / 'HEALTH.json'

LOG = []


def say(x):
    t = str(x)
    LOG.append(t)
    print(t)


def fail(msg):
    """归档失败：写 HEALTH 的 alert，但绝不阻断重建"""
    say('⚠️ ' + msg)
    try:
        write_health({
            'checked_at': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'alert': msg,
            'ok': False,
        })
    except Exception as e:
        say('   （连 HEALTH.json 都没写成：%s）' % e)
    sys.exit(0)


def parse_span(text):
    """'2026-10-26 ~ 2026-11-25' → (date(2026,10,26), date(2026,11,25))

    ★ 分隔符是 ' ~ '，但历史上手填过全角空格、'～'、'-'，全部兼容。
    """
    if not text:
        return None, None
    t = str(text).replace('～', '~').replace('—', '~').replace('–', '~').replace('至', '~')
    if '~' in t:
        a, b = t.split('~', 1)
    elif ' - ' in t:
        a, b = t.split(' - ', 1)
    else:
        return None, None
    try:
        d1 = datetime.date(*[int(x) for x in a.strip().split('-')[:3]])
        d2 = datetime.date(*[int(x) for x in b.strip().split('-')[:3]])
        return d1, d2
    except Exception:
        return None, None


def acct_key_of(d1, d2):
    """会计月标签取**止月**：2026-10-26 ~ 2026-11-25 → '2026-11'"""
    return '%04d-%02d' % (d2.year, d2.month)


def read_months_all():
    if MONTHS_ALL.exists():
        try:
            return json.loads(MONTHS_ALL.read_text(encoding='utf-8'))
        except Exception:
            pass
    return {}


def write_health(h):
    """写健康状态。★ 只在「实质字段」变化时才落盘。

    checked_at 每次都变，若无条件写盘，一天 5 档定时就会产生 5 次 git diff，
    仓库历史会被这些噪音淹没。这里拿掉时间戳比对，其余字段一致就跳过。
    """
    KEY = ('ok', 'current_month', 'current_days', 'current_expected',
           'current_status', 'months_total', 'incomplete_months', 'alert')
    if HEALTH.exists():
        try:
            old = json.loads(HEALTH.read_text(encoding='utf-8'))
            if all(old.get(k) == h.get(k) for k in KEY if k in h):
                say('   · 健康状态无变化，跳过写入')
                return
        except Exception:
            pass
    HEALTH.parent.mkdir(parents=True, exist_ok=True)
    HEALTH.write_text(json.dumps(h, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


def main():
    say('=== data.json 快照归档 %s ===' % datetime.date.today().isoformat())

    # ---------- 1) 读入 ----------
    if not INCOMING.exists():
        fail('_incoming/data.json 不存在，跳过归档')
        return
    try:
        raw = json.loads(INCOMING.read_text(encoding='utf-8'))
    except Exception as e:
        fail('data.json 解析失败：%s' % e)
        return

    meta = raw.get('meta') or {}
    span_txt = meta.get('accounting_month') or ''
    days = [str(x) for x in (meta.get('days_with_data') or [])]
    sale = raw.get('sale') or {}

    if not span_txt or not days or not sale:
        fail('data.json 关键字段缺失（accounting_month=%r / days=%d / sale=%d 天）'
             % (span_txt, len(days), len(sale)))
        return

    d1, d2 = parse_span(span_txt)
    if not d1:
        fail('accounting_month 格式无法识别：%r' % span_txt)
        return
    acct = acct_key_of(d1, d2)
    expected = (d2 - d1).days + 1
    n = len(days)

    today = datetime.date.today()
    if n >= expected:
        status = 'complete'
    elif today > d2:
        status = 'incomplete'
    else:
        status = 'running'

    # ---------- 2) 月归档（每次覆盖，幂等） ----------
    MONTHLY.mkdir(parents=True, exist_ok=True)
    mfile = MONTHLY / ('%s.json' % acct)

    # 汇总数字（写进索引，供 build 端直接读，不必再解整包）
    ret = sum((sale[d].get('t') or 0) for d in days if d in sale)
    die = sum((sale[d].get('diesel') or 0) for d in days if d in sale)
    gas = sum((sale[d].get('gas') or 0) for d in days if d in sale)
    xt = sum(((sale[d].get('xiao') or {}).get('t') or 0) for d in days if d in sale)
    amt = sum((sale[d].get('amt') or 0) for d in days if d in sale)
    wan = sum((sale[d].get('wan') or 0) for d in days if d in sale)

    months_all = read_months_all()

    # ★ 回滚保护的状态一律以**月归档文件自身**为准，不读 months_all。
    #   months_all 是每次从文件重新生成的派生品，拿派生品做判断会不自洽：
    #   文件被写成 3 天，索引也跟着变 3 天，保护条件下次就失效了。
    old_status, old_days, old_hash = None, 0, None
    if mfile.exists():
        try:
            _o = json.loads(mfile.read_text(encoding='utf-8'))
            old_status = _o.get('status')
            old_days = _o.get('days_count') or 0
            old_hash = _o.get('hash')
        except Exception:
            pass

    # ★ 内容指纹：相同就整段跳过写入。
    #   不加这一层的话，stamp 用的是 datetime.now()，一天 5 档定时、数据一字未变，
    #   也会产生 5 次 git diff → 5 个毫无意义的 commit，把仓库历史冲成一团噪音。
    payload_hash = hashlib.sha1(
        json.dumps(raw, ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest()

    # ★ 回滚保护：绝不因为一次「读表不全」的推送，把已经 complete 的月覆盖成 running。
    #   read_daily.js 有体积护栏，但护栏本身也可能漏；这里再兜一道。
    if old_status == 'complete' and n < expected:
        say('   ⚠️ %s 已是 complete（%d 天），本次只有 %d 天 —— 判为读表不全，'
            '不覆盖月归档，保留原文件' % (acct, old_days, n))
        status = 'complete'
    elif old_hash == payload_hash:
        say('   · 月归档内容未变（hash %s），跳过写入' % payload_hash[:8])
    else:
        rec = {
            'acct': acct,
            'span': span_txt,
            'start': d1.isoformat(),
            'end': d2.isoformat(),
            'expected_days': expected,
            'days_count': n,
            'days': days,
            'status': status,
            'retail_t': round(ret, 2),
            'diesel_t': round(die, 2),
            'gas_t': round(gas, 2),
            'xiao_t': round(xt, 2),
            'full_t': round(ret + xt, 2),
            'amt': round(amt, 2),
            'wan': round(wan, 3),
            'stamp': datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%S'),
            'hash': payload_hash,
            'source': 'data.json',
            'payload': raw,
        }
        mfile.write_text(json.dumps(rec, ensure_ascii=False, separators=(',', ':')),
                         encoding='utf-8')
        say('   ✅ 月归档 _archive/monthly/%s.json  %d/%d 天  %s  '
            '零售 %.2f + 小额 %.2f = %.2f 吨'
            % (acct, n, expected, status, ret, xt, ret + xt))

    # ---------- 3) last_good：只留「天数最多」的那一份，防坏数据覆盖 ----------
    #   ★ 跨月要重置比较基准：11 月 3 天当然比 10 月的 31 天少，
    #     若不看会计月就直接比天数，翻月后这份保险就永远不再更新了。
    lg_days, lg_acct = -1, None
    if LAST_GOOD.exists():
        try:
            _lg = json.loads(LAST_GOOD.read_text(encoding='utf-8'))
            lg_days = _lg.get('days_count') or -1
            lg_acct = _lg.get('acct')
        except Exception:
            pass
    if lg_acct != acct:
        lg_days = -1                                     # 换月了，重新开始攒
    if n > lg_days:
        LAST_GOOD.write_text(json.dumps({
            'acct': acct, 'span': span_txt, 'days_count': n,
            'expected_days': expected, 'days': days, 'status': status,
            'stamp': datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%S'),
            'payload': raw,
        }, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
        say('   ✅ 最后好快照 _archive/last_good.json  %s %d 天  %.1f KB'
            % (acct, n, LAST_GOOD.stat().st_size / 1024))
    else:
        say('   · last_good 已是 %s 的 %d 天，本次 %d 天不更（防止坏数据盖掉好数据）'
            % (lg_acct, lg_days, n))

    # ---------- 4) 回溯校验：把「已过期但仍 running」的月标成 incomplete ----------
    #   ★ 这是这套方案里唯一会主动发现问题的机制：
    #     10 月 25 号那天若脚本没跑成，月归档会停在 running；
    #     等到 11 月 26 号以后，这里会把它翻成 incomplete 并写进告警。
    changed = []
    for f in sorted(MONTHLY.glob('*.json')):
        key = f.stem
        try:
            r = json.loads(f.read_text(encoding='utf-8'))
        except Exception:
            continue
        end_s = r.get('end') or ''
        try:
            e = datetime.date(*[int(x) for x in end_s.split('-')[:3]])
        except Exception:
            continue
        if e >= today:
            continue                                      # 还没到期
        if r.get('status') == 'running':
            r['status'] = 'incomplete'
            r['note'] = ('期末 %s 已过，实际只收到 %s/%s 天 —— 该月数据不完整'
                         % (end_s, r.get('days_count'), r.get('expected_days')))
            f.write_text(json.dumps(r, ensure_ascii=False, separators=(',', ':')),
                         encoding='utf-8')
            changed.append(key)
            say('   ⚠️ %s 已过期但只收到 %s/%s 天 → 标为 incomplete'
                % (key, r.get('days_count'), r.get('expected_days')))

    # ---------- 5) 月索引 ----------
    idx = {}
    for f in sorted(MONTHLY.glob('*.json')):
        try:
            r = json.loads(f.read_text(encoding='utf-8'))
        except Exception:
            continue
        idx[f.stem] = {k: r.get(k) for k in
                       ('span', 'start', 'end', 'expected_days', 'days_count',
                        'status', 'retail_t', 'diesel_t', 'gas_t', 'xiao_t',
                        'full_t', 'amt', 'wan', 'stamp')}
        if r.get('note'):
            idx[f.stem]['note'] = r['note']
    MONTHS_ALL.write_text(json.dumps(idx, ensure_ascii=False, indent=1) + '\n',
                          encoding='utf-8')
    say('   ✅ 月索引 _archive/months_all.json  %d 个月' % len(idx))

    # ---------- 6) 健康状态 ----------
    bad = [k for k, v in idx.items() if v.get('status') == 'incomplete']
    alert = None
    if bad:
        alert = ('以下会计月数据不完整：%s —— 请回填报簿补录对应日期，'
                 '或接受该月缺口（不要重跑脚本，重跑补不回来）' % '、'.join(bad))

    write_health({
        'ok': not bad,
        'checked_at': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'current_month': acct,
        'current_span': span_txt,
        'current_days': n,
        'current_expected': expected,
        'current_status': status,
        'months_total': len(idx),
        'incomplete_months': bad,
        'last_good_acct': lg_acct or acct,
        'last_good_days': max(lg_days, n),
        'alert': alert,
    })
    say('   ✅ 健康状态 _archive/HEALTH.json  ok=%s' % (not bad))

    say('=== 归档完成 ===')
    sys.exit(0)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        # 兜底：任何未预料的异常都不许让 Actions 挂掉（页面重建优先）
        fail('归档脚本异常：%s' % e)
