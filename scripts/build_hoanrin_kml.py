# -*- coding: utf-8 -*-
"""国土数値情報A13の保安林・地域森林計画対象民有林と対象地点をKML化（Googleマップ/Earth用）"""
import math, shapefile
LAT, LON = 34.962114039819326, 137.23105862800367
R = 1500.0  # m

KX=111320*math.cos(math.radians(LAT)); KY=110574
BOX=(LON-R/KX, LAT-R/KY, LON+R/KX, LAT+R/KY)

def clip(poly, box):
    """Sutherland-Hodgman: 矩形でポリゴンをクリップ"""
    x0,y0,x1,y1=box
    def half(pts, inside, inter):
        out=[]
        for i in range(len(pts)):
            a=pts[i-1]; b=pts[i]
            ia, ib = inside(a), inside(b)
            if ib:
                if not ia: out.append(inter(a,b))
                out.append(b)
            elif ia:
                out.append(inter(a,b))
        return out
    p=list(poly)
    for inside, inter in (
        (lambda q: q[0]>=x0, lambda a,b:(x0, a[1]+(b[1]-a[1])*(x0-a[0])/(b[0]-a[0]+1e-15))),
        (lambda q: q[0]<=x1, lambda a,b:(x1, a[1]+(b[1]-a[1])*(x1-a[0])/(b[0]-a[0]+1e-15))),
        (lambda q: q[1]>=y0, lambda a,b:(a[0]+(b[0]-a[0])*(y0-a[1])/(b[1]-a[1]+1e-15), y0)),
        (lambda q: q[1]<=y1, lambda a,b:(a[0]+(b[0]-a[0])*(y1-a[1])/(b[1]-a[1]+1e-15), y1)),
    ):
        if not p: return []
        p=half(p, inside, inter)
    return p

def thin(pts, tol=1.5):
    """tol(m)未満の連続点を間引く"""
    out=[]
    for q in pts:
        if not out or abs((q[0]-out[-1][0])*KX)+abs((q[1]-out[-1][1])*KY) > tol:
            out.append(q)
    if out and out[0]!=out[-1]: out.append(out[0])
    return out

def rings(path):
    out=[]
    r=shapefile.Reader(path, encoding='cp932')
    for sr in r.iterShapeRecords():
        sh=sr.shape
        if not sh.points: continue
        bb=sh.bbox
        if not (bb[0]<LON+0.03 and bb[2]>LON-0.03 and bb[1]<LAT+0.03 and bb[3]>LAT-0.03): continue
        pts=sh.points; parts=list(sh.parts)+[len(pts)]
        for i in range(len(parts)-1):
            rg=pts[parts[i]:parts[i+1]]
            if len(rg)<4: continue
            cl=thin(clip(rg, BOX))
            if len(cl)>3: out.append(cl)
    return out

def poly(rg):
    co=' '.join('%.7f,%.7f,0'%(x,y) for x,y in rg)
    return ('<Placemark><styleUrl>#s</styleUrl><Polygon><outerBoundaryIs><LinearRing>'
            '<coordinates>%s</coordinates></LinearRing></outerBoundaryIs></Polygon></Placemark>'%co)

def folder(name, rgs, style):
    body=''.join(poly(r).replace('#s','#'+style) for r in rgs)
    return '<Folder><name>%s</name>%s</Folder>'%(name,body)

styles = ('<Style id="hoan"><LineStyle><color>ff1d1d8c</color><width>2.4</width></LineStyle>'
          '<PolyStyle><color>553b3bd9</color></PolyStyle></Style>'
          '<Style id="minyu"><LineStyle><color>ff3a7a4f</color><width>1.2</width></LineStyle>'
          '<PolyStyle><color>2262bc97</color></PolyStyle></Style>'
          '<Style id="pt"><IconStyle><color>ff1f1fd9</color><scale>1.2</scale>'
          '<Icon><href>http://maps.google.com/mapfiles/kml/paddle/red-circle.png</href></Icon></IconStyle></Style>')

pt = ('<Placemark><name>対象地点 34.962114, 137.231059</name><styleUrl>#pt</styleUrl>'
      '<description>愛知県岡崎市板田町 字東流石（BESS候補地）</description>'
      '<Point><coordinates>%.7f,%.7f,0</coordinates></Point></Placemark>'%(LON,LAT))

kml = ('<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2"><Document>'
       '<name>岡崎市板田町 保安林確認</name>'
       '<description>国土数値情報 森林地域(A13-15 愛知県, 保安林は2002年度整備)。参考表示データであり精度は保証されない。</description>'
       + styles + pt
       + folder('地域森林計画対象民有林', rings('a001230020160209'), 'minyu')
       + folder('保安林（国土数値情報A13）', rings('a001230020160210'), 'hoan')
       + '</Document></kml>')
open('保安林_周辺_A13.kml','w',encoding='utf-8').write(kml)
print('wrote KML', len(kml), 'bytes')
