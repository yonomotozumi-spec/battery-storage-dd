# -*- coding: utf-8 -*-
"""公図(任意座標)の26番4を、配置図の設備位置(=27番1と仮定)を基準に実座標へ平行移動して推定"""
import json, math, os, urllib.request
import numpy as np, shapefile
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon as MPoly
from matplotlib.font_manager import FontProperties
from matplotlib.lines import Line2D
from PIL import Image
FP=FontProperties(fname='/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf')
ps=json.load(open('higashinagareishi.json',encoding='utf-8'))['parcels']
P={p['地番']:np.array([(pt[1],pt[0]) for pt in p['ring']]) for p in ps}   # (E,N) m
EQ_LAT,EQ_LON=34.962328,137.231498          # 配置図の設備重心（実座標）
ANCHOR='27-1'
KX=111320*math.cos(math.radians(EQ_LAT)); KY=110940.0
c0=P[ANCHOR][:-1].mean(0)
def to_ll(xy):   # 公図(E,N) → (lat,lon)  回転0と仮定
    d=xy-c0; return EQ_LAT+d[:,1]/KY, EQ_LON+d[:,0]/KX
# 26番4 の可用部分（北端から10〜65m）中心
a=P['26-4']; lat,lon=to_ll(a)
th=math.radians(90-166); u=np.array([math.cos(th),math.sin(th)])
s=(a-a.mean(0))@u; L0=s.min()
use=a[(s-L0>10)&(s-L0<65)] if ((s-L0>10)&(s-L0<65)).sum()>=3 else a
ulat,ulon=to_ll(use); clat,clon=ulat.mean(),ulon.mean()
print('26番4 全体重心      %.6f, %.6f'%(lat.mean(),lon.mean()))
print('26番4 可用部分の中心 %.6f, %.6f'%(clat,clon))
d=math.hypot((clon-EQ_LON)*KX,(clat-EQ_LAT)*KY); brg=(math.degrees(math.atan2((clon-EQ_LON)*KX,(clat-EQ_LAT)*KY))+360)%360
print('配置図の設備から %.0f m 方位 %.0f°'%(d,brg))
print('Googleマップ: https://www.google.com/maps/search/?api=1&query=%.6f,%.6f'%(clat,clon))
print('地理院地図:  https://maps.gsi.go.jp/#17/%.6f/%.6f/'%(clat,clon))
# --- 図 ---
TILE=256; Z=18
def ll2px(lo,la,z):
    n=TILE*2**z; return (lo+180)/360*n,(1-math.asinh(math.tan(math.radians(la)))/math.pi)/2*n
CLAT,CLON=(clat+EQ_LAT)/2,(clon+EQ_LON)/2+0.0002
mpp=156543.03392*math.cos(math.radians(CLAT))/2**Z; HW,HH=150/mpp,120/mpp
cx,cy=ll2px(CLON,CLAT,Z); x0,x1,y0,y1=cx-HW,cx+HW,cy-HH,cy+HH
tx0,tx1=int(x0//TILE),int(x1//TILE); ty0,ty1=int(y0//TILE),int(y1//TILE)
mos=Image.new('RGB',((tx1-tx0+1)*TILE,(ty1-ty0+1)*TILE),(230,230,230))
for tx in range(tx0,tx1+1):
    for ty in range(ty0,ty1+1):
        fn='seamlessphoto_%d_%d_%d.jpg'%(Z,tx,ty)
        if not os.path.exists(fn):
            try:
                rq=urllib.request.Request('https://cyberjapandata.gsi.go.jp/xyz/seamlessphoto/%d/%d/%d.jpg'%(Z,tx,ty),headers={'User-Agent':'battery-storage-dd/1.0'})
                open(fn,'wb').write(urllib.request.urlopen(rq,timeout=30).read())
            except Exception: open(fn,'wb').write(b'')
        if os.path.getsize(fn)>0: mos.paste(Image.open(fn).convert('RGB'),((tx-tx0)*TILE,(ty-ty0)*TILE))
OX,OY=tx0*TILE,ty0*TILE; img=np.asarray(mos)
fig,ax=plt.subplots(figsize=(11,8.8),dpi=190)
ax.imshow(img,extent=[OX,OX+img.shape[1],OY+img.shape[0],OY],interpolation='bilinear')
def PX(lo,la): return ll2px(lo,la,Z)
# A13保安林
rd=shapefile.Reader('a001230020160210',encoding='cp932')
for sr in rd.iterShapeRecords():
    sh=sr.shape
    if not sh.points: continue
    bb=sh.bbox
    if not(bb[0]<CLON+0.01 and bb[2]>CLON-0.01 and bb[1]<CLAT+0.01 and bb[3]>CLAT-0.01): continue
    pts=sh.points; parts=list(sh.parts)+[len(pts)]
    for i in range(len(parts)-1):
        rg=pts[parts[i]:parts[i+1]]; arr=np.array([PX(p[0],p[1]) for p in rg])
        if arr[:,0].max()<x0-500 or arr[:,0].min()>x1+500 or arr[:,1].max()<y0-500 or arr[:,1].min()>y1+500: continue
        ax.add_patch(MPoly(arr,closed=True,facecolor='#2C5F2D',edgecolor='#FFE24D',lw=2.5,alpha=0.25,zorder=3))
# 公図の筆を平行移動して重ねる（推定）
style={'26-4':('#F9D648','#8A6D00',0.55,3.0),'26-1':('#D9534F','#8C1D18',0.30,1.5),'27-4':('#D9534F','#8C1D18',0.30,1.5),
       '27-1':('#D9534F','#8C1D18',0.30,1.5),'27-5':('#D9534F','#8C1D18',0.30,1.5),'11':('#D9534F','#8C1D18',0.25,1.5),
       '26-3':('#DDDDDD','#888888',0.30,1.2),'29-2':('#DDDDDD','#888888',0.30,1.2),'22-2':('#D9534F','#8C1D18',0.25,1.5),'23':('#7FB3D5','#2E5F86',0.35,1.2)}
for k,(fc,ec,al,lw) in style.items():
    la,lo=to_ll(P[k]); arr=np.array([PX(o,l) for l,o in zip(la,lo)])
    ax.add_patch(MPoly(arr,closed=True,facecolor=fc,edgecolor=ec,lw=lw,alpha=al,ls='--',zorder=4))
    c=arr[:-1].mean(0); ax.text(c[0],c[1],k,fontproperties=FP,fontsize=9,ha='center',va='center',color='#111',zorder=8,
                                bbox=dict(fc='white',ec='none',alpha=0.6,pad=1))
# 設備（現行配置図）と推定26番4中心
eq=json.load(open('equip_ll.json',encoding='utf-8'))
for e in eq:
    arr=np.array([PX(c[0],c[1]) for c in e['hull']]); ax.add_patch(MPoly(arr,closed=True,facecolor='#E8352B',edgecolor='#7A0F0A',lw=1.2,zorder=6))
p=PX(clon,clat); ax.plot([p[0]],[p[1]],marker='*',ms=18,mfc='#FFD400',mec='#7A0F0A',mew=1.5,zorder=9)
ax.annotate('26番4 推定中心\n%.5f, %.5f'%(clat,clon),xy=p,xytext=(18,14),textcoords='offset points',fontproperties=FP,fontsize=10,color='white',zorder=10,
            bbox=dict(fc='#7A0F0A',ec='none',alpha=0.9,boxstyle='round,pad=0.3'))
p2=PX(EQ_LON,EQ_LAT); ax.annotate('現行配置図の設備',xy=p2,xytext=(14,-22),textcoords='offset points',fontproperties=FP,fontsize=10,color='white',zorder=10,
            bbox=dict(fc='#E8352B',ec='none',alpha=0.9,boxstyle='round,pad=0.3'))
ax.set_xlim(x0,x1); ax.set_ylim(y1,y0); ax.set_xticks([]); ax.set_yticks([])
ax.set_title('26番4の推定位置（公図を27番1＝現行設備位置に合わせて平行移動・誤差±40m程度）',fontproperties=FP,fontsize=13,color='#1F4620',pad=10)
h=[Line2D([],[],marker='s',ls='',ms=12,mfc='#F9D648',mec='#8A6D00',label='26番4（推定・破線）'),
   Line2D([],[],marker='s',ls='',ms=12,mfc='#D9534F',mec='#8C1D18',label='保安林の筆（推定・破線）'),
   Line2D([],[],marker='s',ls='',ms=12,mfc='#7A9B78',mec='#FFE24D',label='保安林（国土数値情報A13）'),
   Line2D([],[],marker='s',ls='',ms=12,mfc='#E8352B',mec='#7A0F0A',label='現行配置図の設備（実測ベース）')]
ax.legend(handles=h,loc='lower left',prop=FP,fontsize=9.5,framealpha=0.94)
bx=x0+(x1-x0)*0.7; by=y1-(y1-y0)*0.06; L=50/mpp
ax.plot([bx,bx+L],[by,by],color='white',lw=5,zorder=9); ax.plot([bx,bx+L],[by,by],color='black',lw=2.5,zorder=10)
ax.text(bx+L/2,by-(y1-y0)*0.012,'50m',ha='center',fontproperties=FP,fontsize=10,color='white',zorder=11,bbox=dict(fc='black',ec='none',alpha=0.5,pad=1.5))
fig.subplots_adjust(left=0.02,right=0.98,top=0.94,bottom=0.02); fig.savefig('fig_264_est.png')
# KML 追記
la,lo=to_ll(a)
co=' '.join('%.7f,%.7f,0'%(o,l) for l,o in zip(la,lo))
kml=('<Style id="est"><LineStyle><color>ff0090ff</color><width>2.5</width></LineStyle><PolyStyle><color>5548d6f9</color></PolyStyle></Style>'
     '<Folder><name>26番4 推定位置（誤差±40m・法的位置ではない）</name>'
     '<Placemark><name>26番4（推定）</name><styleUrl>#est</styleUrl><Polygon><outerBoundaryIs><LinearRing><coordinates>%s</coordinates></LinearRing></outerBoundaryIs></Polygon></Placemark>'
     '<Placemark><name>26番4 推定中心</name><Point><coordinates>%.7f,%.7f,0</coordinates></Point></Placemark></Folder>'%(co,clon,clat))
s=open('保安林_周辺_A13.kml',encoding='utf-8').read().replace('</Document></kml>',kml+'</Document></kml>')
open('保安林_周辺_A13.kml','w',encoding='utf-8').write(s)
json.dump({'center':[clat,clon],'whole':[lat.mean(),lon.mean()]},open('est264.json','w'))
print('saved')
