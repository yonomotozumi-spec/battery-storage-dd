# -*- coding: utf-8 -*-
"""法務省 登記所備付地図XML を読み、指定小字の筆ポリゴンを GeoJSON 風 dict で返す"""
import re, json, sys
from xml.etree import ElementTree as ET

NS = {'m':'http://www.moj.go.jp/MINJI/tizuxml','z':'http://www.moj.go.jp/MINJI/tizuzumen'}

def parse(path, koaza=None):
    tree = ET.parse(path); root = tree.getroot()
    zahyo = root.findtext('m:座標系', namespaces=NS)
    # points (indirect refs)
    pts_by_id = {}
    for pt in root.iter('{%s}GM_Point' % NS['z']):
        d = pt.find('.//z:DirectPosition', NS)
        if d is None: continue
        pts_by_id[pt.get('id')] = (float(d.findtext('z:X', namespaces=NS)), float(d.findtext('z:Y', namespaces=NS)))
    # curves
    curves = {}
    for c in root.iter('{%s}GM_Curve' % NS['z']):
        cid = c.get('id')
        if cid in curves: continue
        pts = []
        for col in c.iter('{%s}GM_PointArray.column' % NS['z']):
            d = col.find('z:GM_Position.direct', NS)
            if d is not None:
                pts.append((float(d.findtext('z:X', namespaces=NS)), float(d.findtext('z:Y', namespaces=NS))))
                continue
            ind = col.find('z:GM_Position.indirect/z:GM_PointRef.point', NS)
            if ind is not None:
                q = pts_by_id.get(ind.get('idref'))
                if q: pts.append(q)
        curves[cid] = pts
    # surfaces
    surfaces = {}
    for s in root.iter('{%s}GM_Surface' % NS['z']):
        sid = s.get('id')
        ext = s.find('.//z:GM_SurfaceBoundary.exterior', NS)
        if ext is None: continue
        ring = []
        for g in ext.iter('{%s}GM_CompositeCurve.generator' % NS['z']):
            seg = curves.get(g.get('idref'))
            if not seg: continue
            if ring and ring[-1] == seg[0]:
                ring.extend(seg[1:])
            elif ring and ring[-1] == seg[-1]:
                ring.extend(list(reversed(seg))[1:])
            elif not ring:
                ring.extend(seg)
            else:
                ring.extend(seg)
        if ring and ring[0] != ring[-1]:
            ring.append(ring[0])
        surfaces[sid] = ring
    # parcels
    out = []
    for h in root.iter('{%s}筆' % NS['m']):
        ko = h.findtext('m:小字名', namespaces=NS) or ''
        if koaza and ko != koaza: continue
        shape = h.find('m:形状', NS)
        sid = shape.get('idref') if shape is not None else None
        ring = surfaces.get(sid)
        if not ring or len(ring) < 4: continue
        out.append({'大字': h.findtext('m:大字名', namespaces=NS),
                    '小字': ko,
                    '地番': h.findtext('m:地番', namespaces=NS),
                    '種別': h.findtext('m:座標値種別', namespaces=NS),
                    'ring': ring})
    return zahyo, out

if __name__ == '__main__':
    z, ps = parse(sys.argv[1], sys.argv[2] if len(sys.argv)>2 else None)
    print('座標系:', z, '/ 筆数:', len(ps))
    json.dump({'座標系': z, 'parcels': ps}, open(sys.argv[3],'w',encoding='utf-8'), ensure_ascii=False)
