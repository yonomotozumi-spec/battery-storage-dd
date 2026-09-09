# -*- coding: utf-8 -*-
"""配置図の設備ポリゴンを実緯度経度に変換して出力"""
import math, json
import numpy as np
from PIL import Image
from scipy import ndimage
from scipy.spatial import ConvexHull
LAT0,LON0=34.962114039819326,137.23105862800367
PIN=(465.0,1507.0); MPP=0.047; THETA=-11.6
KX=111320*math.cos(math.radians(LAT0)); KY=110940.0
bu=math.radians(-THETA); br=math.radians(-THETA+90)
def to_ll(u,v):
    right=(u-PIN[0])*MPP; up=(PIN[1]-v)*MPP
    N=up*math.cos(bu)+right*math.cos(br); E=up*math.sin(bu)+right*math.sin(br)
    return LAT0+N/KY, LON0+E/KX

A=np.asarray(Image.open('haichi-1.png').convert('RGB')).astype(int)
r,g,b=A[:,:,0],A[:,:,1],A[:,:,2]
red=(r>190)&(g<90)&(b<90)
lab,n=ndimage.label(red)
sizes=ndimage.sum(red,lab,range(1,n+1))
names={}
out=[]
for i in np.argsort(sizes)[::-1][:5]:
    m=(lab==i+1); yy,xx=np.nonzero(m)
    if sizes[i]<5000: continue
    cxm,cym=xx.mean(),yy.mean()
    if abs(cxm-465)<120 and abs(cym-1507)<120:   # Googleピンは除外
        print('pin cluster center (%.0f,%.0f) size %d'%(cxm,cym,sizes[i])); continue
    pts=np.stack([xx,yy],1).astype(float)
    hull=pts[ConvexHull(pts).vertices]
    lls=[to_ll(p[0],p[1]) for p in hull]
    la,lo=to_ll(cxm,cym)
    out.append({'center':[round(la,6),round(lo,6)],'hull':[[round(x[1],7),round(x[0],7)] for x in lls],
                'px_center':[round(cxm,1),round(cym,1)],'px_area':int(sizes[i])})
    print('cluster px(%.0f,%.0f) -> %.6f, %.6f'%(cxm,cym,la,lo))
json.dump(out,open('equip_ll.json','w',encoding='utf-8'),ensure_ascii=False,indent=1)

allpx=np.vstack([np.array(o['px_center']) for o in out])
cu,cv=allpx[:,0].mean(),allpx[:,1].mean()
la,lo=to_ll(cu,cv)
print('\n設備全体の重心 -> %.6f, %.6f'%(la,lo))
d=math.hypot((lo-LON0)*KX,(la-LAT0)*KY)
brg=(math.degrees(math.atan2((lo-LON0)*KX,(la-LAT0)*KY))+360)%360
print('与点からの距離 %.0f m 方位 %.0f 度'%(d,brg))
# 構内柱（シアン丸）と01キ401
print('\n参考: 構内柱 px(1490,616)->', '%.6f, %.6f'%to_ll(1490,616))
print('参考: 01キ401 px(1922,552)->', '%.6f, %.6f'%to_ll(1922,552))
