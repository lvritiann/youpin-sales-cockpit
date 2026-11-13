# -*- coding: utf-8 -*-
"""「大客户构成」下拉 + 口径脚注回归探针。

验五件事：
  1) 公司视角显示这张卡、单站视角整块让位给「消费结构」
  2) 下拉含「跟随当前月 + 3 个会计年 + 34 个月」且默认跟随
  3) 每个视图下 环形图中心数 ≡ 图例吨数之和（精度铁律）
  4) 单月显示「已填 N 天」、含进行中月份的会计年写「不计算同比」、
     完整会计年必须给出「上年同期 + 同比」
  5) 副图恒含「大客户」行（按销量取前四时它可能跌出前四）

用法：python _kdocs/reg_cust.py
      CUST_SRC=<线上快照.html> python _kdocs/reg_cust.py
"""
import json
import os
import re
import subprocess
import sys
import tempfile

EDGE = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
SRC = os.environ.get('CUST_SRC') or os.path.join(
    os.path.dirname(os.path.abspath(__file__)), '..', 'cockpit', 'index.html')

PROBE = r"""
<script>
window.__E = [];
window.addEventListener('error', function(e){ window.__E.push(e.message); });
setTimeout(function(){
  try { %s } catch(e){ window.__E.push('ACT:' + e.message); }
  function T(el){ return el ? el.textContent.replace(/\s+/g, ' ').trim() : ''; }
  function N(s){ return parseFloat(String(s).replace(/[^\d.\-]/g, '')); }
  var o = {err: window.__E};
  var g = document.querySelector('#custgrid');
  o.custDisp = g ? getComputedStyle(g).display : 'NO_GRID';
  var vg = document.querySelector('#vgrid');
  o.vgDisp = vg ? getComputedStyle(vg).display : 'NO_VG';
  var sel = document.querySelector('#custsel');
  o.opts = sel ? sel.options.length : 0;
  o.selVal = sel ? sel.value : '';
  o.selTxt = (sel && sel.selectedIndex >= 0) ? sel.options[sel.selectedIndex].text : '';
  o.span = T(document.querySelector('#cust-span'));
  o.ctr = N(T(document.querySelector('#cd-v')));
  o.legend = Array.prototype.map.call(document.querySelectorAll('#cls .ci'), function(r){
    return { name: T(r.querySelector('span')), t: N(T(r.querySelector('b'))),
             pct: T(r.querySelector('u')) };
  });
  o.cmp = Array.prototype.map.call(document.querySelectorAll('#cmp .row'), function(r){
    return { name: T(r.querySelector('span')), pct: T(r.querySelector('b')),
             w: (r.querySelector('.fl') || {}).style ?
                r.querySelector('.fl').style.width : '' };
  });
  o.foot = T(document.querySelector('#cmp-t'));
  o.note = Array.prototype.map.call(document.querySelectorAll('#cust-note > div'), T);
  var s = o.legend.map(function(x){ return x.t; }).filter(function(v){ return !isNaN(v); })
                   .reduce(function(a, b){ return a + b; }, 0);
  o.legSum = Math.round(s * 10) / 10;
  o.legN = o.legend.length;
  o.cmpSum = Math.round(o.cmp.map(function(x){ return N(x.pct); })
    .filter(function(v){ return !isNaN(v); })
    .reduce(function(a, b){ return a + b; }, 0) * 10) / 10;
  document.title = 'REG::' + JSON.stringify(o);
}, 1300);
</script>
"""


def run(act='', w=1680, h=1400):
    html = open(SRC, encoding='utf-8').read()
    html = html.replace('</body>', (PROBE % (act or '')) + '</body>')
    tmp = os.path.join(tempfile.gettempdir(), 'kdoc_regcust.html')
    open(tmp, 'w', encoding='utf-8').write(html)
    cmd = [EDGE, '--headless=new', '--disable-gpu', '--window-size=%d,%d' % (w, h),
           '--virtual-time-budget=8000', '--dump-dom',
           'file:///' + tmp.replace('\\', '/')]
    r = subprocess.run(cmd, capture_output=True, timeout=150)
    dom = r.stdout.decode('utf-8', 'ignore')
    m = re.search(r'REG::(\{.*?\})</title>', dom, re.S)
    if not m:
        print('!! 未拿到探针结果')
        print(dom[:800])
        return None
    return json.loads(m.group(1))


CASES = [
    ('默认（跟随当前会计月）', ''),
    ('2026 会计年累计', "CUST_SEL='y:2026'; render();"),
    ('2025 会计年累计', "CUST_SEL='y:2025'; render();"),
    ('锁定 2024-03', "CUST_SEL='2024-03'; render();"),
    ('单站视角（华山）', "pickScope('华山');"),
]


def main():
    fails = []
    for name, act in CASES:
        d = run(act)
        if not d:
            fails.append('%s：探针无返回' % name)
            continue
        print('\n===== %s =====' % name)
        print('  卡片 display=%s | 消费结构 display=%s | 下拉项=%d value=%r'
              % (d['custDisp'], d['vgDisp'], d['opts'], d['selVal']))
        print('  选中项  ', d['selTxt'])
        print('  口径条  ', d['span'])
        print('  中心数  %s 吨 | 图例 %d 段合计 %s 吨 | 差 %.2f'
              % (d['ctr'], d['legN'], d['legSum'], (d['ctr'] or 0) - d['legSum']))
        print('  图例    ', ' / '.join('%s %s吨 %s' % (x['name'], x['t'], x['pct']) for x in d['legend']))
        print('  副图    ', ' / '.join('%s %s' % (x['name'], x['pct']) for x in d['cmp']))
        print('  副图合计 %.1f%%' % d['cmpSum'])
        print('  脚注1行 ', d['foot'])
        for i, x in enumerate(d['note'], 1):
            print('  脚注+%d   %s' % (i, x))

        if d['err']:
            fails.append('%s：JS 报错 %s' % (name, d['err']))
        if name.startswith('单站'):
            if d['custDisp'] != 'none':
                fails.append('%s：大客户卡未隐藏（%s）' % (name, d['custDisp']))
            if d['vgDisp'] == 'none':
                fails.append('%s：消费结构未显示' % name)
            continue
        if d['custDisp'] == 'none':
            fails.append('%s：大客户卡被隐藏' % name)
        if d['opts'] < 38:
            fails.append('%s：下拉项只有 %d（应 ≥38）' % (name, d['opts']))
        if abs((d['ctr'] or 0) - d['legSum']) > 0.3:
            fails.append('%s：中心数 %.1f ≠ 图例合计 %.1f'
                         % (name, d['ctr'], d['legSum']))
        names = [x['name'] for x in d['cmp']]
        if '大客户' not in names:
            fails.append('%s：副图缺「大客户」行' % name)

    # 场景专属断言
    d0 = run('')
    if d0 and d0['selVal'] != '':
        fails.append('默认视图的 CUST_SEL 不是空（%r）' % d0['selVal'])
    d1 = run("CUST_SEL='y:2026'; render();")
    if d1 and not any('不计算同比' in x for x in d1['note']):
        fails.append('2026 会计年（含进行中 11 月）未标注「不计算同比」')
    d2 = run("CUST_SEL='y:2025'; render();")
    if d2 and not any('同比' in x and '上年同期' in x for x in d2['note']):
        fails.append('2025 完整会计年缺「上年同期 + 同比」')
    d3 = run("CUST_SEL='2024-03'; render();")
    if d3 and not any('已填' in x for x in d3['note']):
        fails.append('单月视图未标「已填 N 天」')
    d4 = run("CUST_SEL='y:2026'; render();")
    if d4 and not any('未署名客户槽位' in x for x in d4['note']):
        fails.append('2026 会计年未提示「未署名客户槽位」（应有）')

    print('\n' + '=' * 60)
    if fails:
        print('未通过 %d 项：' % len(fails))
        for f in fails:
            print('  ·', f)
        sys.exit(1)
    print('全部通过')


if __name__ == '__main__':
    main()
