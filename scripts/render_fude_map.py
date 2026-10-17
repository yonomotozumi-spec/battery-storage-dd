# -*- coding: utf-8 -*-
import json, re
import numpy as np
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon as MPoly
from matplotlib.font_manager import FontProperties
from matplotlib.lines import Line2D

FP = FontProperties(fname='/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf')
ps = json.load(open('higashinagareishi.json', encoding='utf-8'))['parcels']

HOAN = set('1-8 11 11-1 11-2 22-1 22-2 22-3 22-4 22-8 26-2 27-1 27-2 27-3 27-4 27-5'.split())
TA   = set('12 13 18 19 22-5 23'.split())
SAN  = set('19-1 21 24 25 26-1 26-4'.split())
SITE = set(('1-1 1-3 1-4 1-5 1-6 1-7 1-8 1-9 1-10 1-11 1-12 1-13 1-14 1-15 1-16 1-22 1-27 1-28 '
            '1-29 1-30 1-31 1-32 1-33 1-37 2 3 4 5 6-1 11 11-1 11-2 12 13 14 15 16 17 18 19 19-1 '
            '20 21 22-1 22-2 22-3 22-4 22-5 22-8 23 24 25 26-1 26-2 26-4 27-1 27-2 27-3 27-4 27-5').split())

C = {'保安林':('#D9534F','#8C1D18'), '田':('#7FB3D5','#2E5F86'), '山林':('#A9CF7B','#4F7A3A'),
     '案件地・地目未確認':('#F0AD4E','#A66A10'), '案件地外':('#E8E8E3','#B9BFB2')}

norm = lambda t: re.sub(r'[VW]\d+$', '', t)
def area(r):
    a = 0.0
    for i in range(len(r)-1): a += r[i][1]*r[i+1][0] - r[i+1][1]*r[i][0]
    return abs(a)/2

def cat_of(k):
    if k in HOAN: return '保安林'
    if k in TA:   return '田'
    if k in SAN:  return '山林'
    return '案件地・地目未確認' if k in SITE else '案件地外'

fig, ax = plt.subplots(figsize=(13.5,11), dpi=200)
stat = {}
for p in ps:
    k = norm(p['地番']); c = cat_of(k)
    xy = np.array([(pt[1], pt[0]) for pt in p['ring']])
    fc, ec = C[c]
    z = {'保安林':5,'案件地・地目未確認':4,'田':4,'山林':4,'案件地外':2}[c]
    ax.add_patch(MPoly(xy, closed=True, facecolor=fc, edgecolor=ec,
                       lw=1.5 if c=='保安林' else (0.9 if c!='案件地外' else 0.5),
                       alpha=0.9 if c!='案件地外' else 0.5, zorder=z))
    a = area(p['ring'])
    st = stat.setdefault(c, {'fude': set(), 'area': 0.0}); st['fude'].add(k); st['area'] += a
    if a > 300 and c != '案件地外':
        ax.text(xy[:,0].mean(), xy[:,1].mean(), p['地番'], fontproperties=FP,
                fontsize=6.6 if a < 1100 else 8.4, ha='center', va='center',
                color='#111', zorder=8)

allp = np.array([pt for q in ps for pt in q['ring']])
x0,x1 = allp[:,1].min(), allp[:,1].max(); y0,y1 = allp[:,0].min(), allp[:,0].max()
mx,my = (x1-x0)*0.03, (y1-y0)*0.03
ax.set_xlim(x0-mx, x1+mx); ax.set_ylim(y0-my, y1+my)
ax.set_aspect('equal', adjustable='box'); ax.set_xticks([]); ax.set_yticks([])
for s in ax.spines.values(): s.set_edgecolor('#C6CEC0')
ax.set_title('岡崎市板田町 字東流石　筆界図と地目（法務省 登記所備付地図データ 2026年）',
             fontproperties=FP, fontsize=17, color='#1F4620', pad=14)

order = ['保安林','田','山林','案件地・地目未確認','案件地外']
labels = {'保安林':'地目「保安林」','田':'地目「田」','山林':'地目「山林」',
          '案件地・地目未確認':'案件地だが登記簿未取得','案件地外':'案件地外（字東流石の他の筆）'}
handles = [Line2D([],[],marker='s',ls='',ms=13,mfc=C[c][0],mec=C[c][1],
                  label='%s　%d筆・約%s㎡' % (labels[c], len(stat[c]['fude']), format(int(stat[c]['area']), ',')))
           for c in order if c in stat]
ax.legend(handles=handles, loc='lower right', prop=FP, fontsize=11, framealpha=0.96)

# scale bar (座標単位=m)
bx, by = x0 + (x1-x0)*0.04, y0 + (y1-y0)*0.055
ax.plot([bx,bx+100],[by,by], color='#222', lw=3.5, zorder=9, solid_capstyle='butt')
ax.text(bx+50, by+(y1-y0)*0.012, '100m', ha='center', fontproperties=FP, fontsize=10, color='#222', zorder=9)
# north arrow
nx = x1-mx*1.2
ax.annotate('', xy=(nx, y1-my*1.5), xytext=(nx, y1-my*5.5),
            arrowprops=dict(arrowstyle='-|>', color='#1F4620', lw=2), zorder=9)
ax.text(nx, y1-my*1.2, 'N', fontproperties=FP, fontsize=13, color='#1F4620',
        ha='center', va='bottom', zorder=9)
ax.text(0.01, 0.015,
        '※任意座標系のため実際の緯度経度には配置できない（形状・縮尺・相対位置は正確）。求積図の実測面積と概ね一致。',
        transform=ax.transAxes, fontproperties=FP, fontsize=9.5, color='#4A5546',
        bbox=dict(fc='white', ec='none', alpha=0.85, pad=3))
fig.subplots_adjust(left=0.02, right=0.98, top=0.94, bottom=0.02)
fig.savefig('fude_map.png')

for c in order:
    if c in stat:
        print('%-20s %3d筆  公図面積 %9s ㎡' % (c, len(stat[c]['fude']), format(int(stat[c]['area']), ',')))
site_area = sum(stat[c]['area'] for c in order[:4] if c in stat)
site_fude = sum(len(stat[c]['fude']) for c in order[:4] if c in stat)
print('案件地 計 %d筆  %s ㎡' % (site_fude, format(int(site_area), ',')))
