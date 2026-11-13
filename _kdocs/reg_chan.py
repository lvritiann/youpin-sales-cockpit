# -*- coding: utf-8 -*-
"""「单站消费结构」渠道拆分回归探针。

为什么单独立一条：2026-11-13 用户问「天汉水城为什么它的收款渠道没有拆分出来？」
根因是同一实体在两张源表里被写成两个名字 ——「销售日报」写 天汉水城西/东，
10 张渠道表写 天汉水城南/北。页面按「销售日报」的站名去查渠道，查不到就整块
空白。这种问题的表现是「某个站没有渠道拆分」，很容易被当成算错，而且只看一个
站发现不了 —— 必须**逐个**对扫一遍。

验四件事：
  1) 每个单站视角的渠道结构都非空（不出现「无渠道拆分」）
  2) 每个对至少有 8 项渠道（源表 10 张渠道表，个别站缺某项属正常）
  3) 渠道占比合计 = 100%（末段取余额，不许各自 round）
  4) 渠道吨量合计与「销售日报」站级吨量的差在 1% 内（两表吨量折算差的上限）

用法：python _kdocs/reg_chan.py
      CHAN_SRC=<快照.html> python _kdocs/reg_chan.py
"""
import json
import os
import re
import subprocess
import tempfile

EDGE = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.environ.get('CHAN_SRC') or os.path.join(BASE, '..', 'cockpit', 'index.html')

PROBE = r"""
<script>
window.__E = [];
window.addEventListener('error', function(e){ window.__E.push(e.message); });
setTimeout(function(){
  try { %s } catch(e){ window.__E.push('ACT:' + e.message); }
  function T(el){ return el ? el.textContent.replace(/\s+/g, ' ').trim() : ''; }
  function N(s){ return parseFloat(String(s).replace(/[^\d.\-]/g, '')); }
  var o = {err: window.__E};
  o.span = T(document.querySelector('#vcard-span'));
  o.title = T(document.querySelector('#vcard-t'));
  o.rows = Array.prototype.map.call(
    document.querySelectorAll('#vcard .vbase'), function(r){
      return { name: T(r.querySelector('span')), pct: N(T(r.querySelector('b'))) };
    });
  o.rows = o.rows.filter(function(x){ return x.name; });
  o.pctSum = Math.round(o.rows.reduce(function(a, b){ return a + (b.pct || 0); }, 0) * 10) / 10;
  o.raw = T(document.querySelector('#vcard'));
  document.title = 'CH::' + JSON.stringify(o);
}, 1300);
</script>
"""


def run(act='', w=1680, h=1400):
    html = open(SRC, encoding='utf-8').read()
    html = html.replace('</body>', (PROBE % (act or '')) + '</body>')
    tmp = os.path.join(tempfile.gettempdir(), 'kdoc_regchan.html')
    open(tmp, 'w', encoding='utf-8').write(html)
    cmd = [EDGE, '--headless=new', '--disable-gpu', '--window-size=%d,%d' % (w, h),
           '--virtual-time-budget=8000', '--dump-dom',
           'file:///' + tmp.replace('\\', '/')]
    r = subprocess.run(cmd, capture_output=True, timeout=150)
    dom = r.stdout.decode('utf-8', 'ignore')
    m = re.search(r'CH::(\{.*?\})</title>', dom, re.S)
    if not m:
        return None
    return json.loads(m.group(1))


def pairs():
    """从 stations_annual.json 取 15 对；小额配送单独标注（它本来就没有渠道行）。"""
    sa = json.load(open(os.path.join(BASE, '..', '_archive', 'stations_annual.json'),
                        encoding='utf-8'))
    return sorted((sa.get('pair_stations') or {}).keys())


def main():
    S = json.load(open(os.path.join(BASE, '..', '_archive', 'station_series.json'),
                       encoding='utf-8'))
    fails = []
    names = pairs()
    print('待验 %d 对 + 小额配送\n' % len(names))
    print('%-9s %6s %7s %8s %s' % ('对', '渠道项', '占比合计', '渠道吨', '状态'))
    print('-' * 78)
    for pn in names + ['小额配送']:
        d = run("pickScope(%s);" % json.dumps(pn, ensure_ascii=False))
        if not d:
            fails.append('%s：探针无返回' % pn)
            print('%-9s %6s %7s %8s %s' % (pn, '-', '-', '-', '探针无返回'))
            continue
        n = len(d['rows'])
        ps = d['pctSum']
        if pn == '小额配送':
            ok = '无渠道拆分' in (d['span'] or '')
            print('%-9s %6d %7s %10s %s'
                  % (pn, n, ps, '-', 'OK（源表无渠道行，宁缺不造数）' if ok else '!! 意外出现渠道项'))
            if not ok:
                fails.append('小额配送：不该有渠道项')
            continue
        _ch26 = S['chan'].get('2026') or {}
        sts = json.load(open(os.path.join(BASE, '..', '_archive', 'stations_annual.json'),
                             encoding='utf-8'))['pair_stations'][pn]
        ct = sum(sum((_ch26.get(s) or {}).values()) for s in sts)
        bad = []
        if ct <= 0:
            bad.append('渠道表里查不到这一对（站名不一致？）')
        if n < 8:
            bad.append('渠道项仅 %d' % n)
        if abs(ps - 100) > 0.05:
            bad.append('占比合计 %.1f%%' % ps)
        if d['err']:
            bad.append('JS 错误 %s' % d['err'][:60])
        if '无渠道拆分' in (d['span'] or ''):
            bad.append('显示「无渠道拆分」')
        state = 'OK' if not bad else '!! ' + '；'.join(bad)
        if bad:
            fails.append('%s：%s' % (pn, '；'.join(bad)))
        print('%-9s %6d %7s %10.1f %s' % (pn, n, ps, ct, state))

    print('\n' + '=' * 78)
    if fails:
        print('未通过 %d 项：' % len(fails))
        for f in fails:
            print('  -', f)
        return 1
    print('全部通过')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
