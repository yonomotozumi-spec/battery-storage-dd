# -*- coding: utf-8 -*-
"""求積図の道路分岐(=01キ401)を基準に 26番4 の位置を推定。求積図は北上・0.527m/px と仮定"""
import json, math, numpy as np
KX=111320*math.cos(math.radians(34.9624)); KY=110940.0; MPP=0.527
A_PX=(1060.0,240.0); A_LL=(34.962386,137.231893)   # 求積図px(道路西端@分岐) ≈ 01キ401
def k2ll(u,v):
    return A_LL[0]+(A_PX[1]-v)*MPP/KY, A_LL[1]+(u-A_PX[0])*MPP/KX
pts={'26番4 北端付近':(1035,96),'26番4 ラベル位置':(1026,138),'26番4 可用部分の中心':(1040,168),'26番4 南端付近（幅4m）':(1075,300),
     '27番4 ラベル':(1003,278),'27番1 ラベル':(1046,332),'26番1 ラベル':(958,152)}
out={}
for k,(u,v) in pts.items():
    la,lo=k2ll(u,v); out[k]=(la,lo); print('%-16s → %.6f, %.6f'%(k,la,lo))
# 配置図設備(実座標)は求積図のどこか
eq=(34.962328,137.231498)
u=A_PX[0]+(eq[1]-A_LL[1])*KX/MPP; v=A_PX[1]-(eq[0]-A_LL[0])*KY/MPP
print('配置図設備 → 求積図px (%.0f, %.0f)  ※27-4ラベル(1003,278)/27-1ラベル(1046,332) との関係'%(u,v))
c=out['26番4 可用部分の中心']; d=math.hypot((c[1]-eq[1])*KX,(c[0]-eq[0])*KY)
print('可用部分中心は配置図設備から %.0f m'%d)
prev=(34.963167,137.231627); print('前回推定との差 %.0f m'%math.hypot((c[1]-prev[1])*KX,(c[0]-prev[0])*KY))
json.dump({k:list(v) for k,v in out.items()},open('est264b.json','w'),ensure_ascii=False,indent=1)
print('Googleマップ: https://www.google.com/maps/search/?api=1&query=%.6f,%.6f'%c)
