#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
緯度経度が「保安林」「国有林」「地域森林計画対象民有林」「森林地域」に該当するかを判定する。

データ: 国土数値情報「森林地域(A13)」(国土交通省) のシェープファイル。
        レイヤ番号 07=森林地域 / 08=国有林 / 09=地域森林計画対象民有林 / 10=保安林

使い方:
    python3 scripts/check_hoanrin.py --lat 34.962114 --lon 137.231059
    python3 scripts/check_hoanrin.py --lat 34.96 --lon 137.23 --pref 23   # 都道府県コード指定
    python3 scripts/check_hoanrin.py --lat 34.96 --lon 137.23 --json out.json

依存: pyshp   (pip install pyshp)

注意: A13は土地利用基本計画の参考表示であり、国土交通省も精度を保証していない。
      「該当なし」でも境界至近の場合があるため、最終判断は必ず県の農林水産事務所
      (保安林台帳・保安林指定図) で確認すること。
"""
import argparse
import io
import json
import math
import os
import sys
import urllib.request
import zipfile

CACHE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "forest_cache")
A13_URL = "https://nlftp.mlit.go.jp/ksj/gml/data/A13/A13-{ver}/A13-{ver}_{pref:02d}_GML.zip"
# 都道府県ごとに最新版が異なるため、新しい順に試す
A13_VERSIONS = ["15", "11", "06"]

LAYERS = {
    "07": "森林地域",
    "08": "国有林",
    "09": "地域森林計画対象民有林",
    "10": "保安林",
}


# ---------- 幾何 ----------
def pip_ring(x, y, ring):
    inside = False
    n = len(ring)
    j = n - 1
    for i in range(n):
        xi, yi = ring[i][0], ring[i][1]
        xj, yj = ring[j][0], ring[j][1]
        if ((yi > y) != (yj > y)) and (x < (xj - xi) * (y - yi) / (yj - yi) + xi):
            inside = not inside
        j = i
    return inside


def ring_area(ring):
    s = 0.0
    for i in range(len(ring) - 1):
        s += ring[i][0] * ring[i + 1][1] - ring[i + 1][0] * ring[i][1]
    return s / 2.0


def rings_of(shape):
    pts = shape.points
    parts = list(shape.parts) + [len(pts)]
    return [pts[parts[i]:parts[i + 1]] for i in range(len(parts) - 1)]


def seg_dist_m(px, py, a, b):
    """点(px,py)と線分ab の距離[m](簡易平面近似)"""
    kx = 111320.0 * math.cos(math.radians(py))
    ky = 110574.0
    ax = (a[0] - px) * kx
    ay = (a[1] - py) * ky
    bx = (b[0] - px) * kx
    by = (b[1] - py) * ky
    dx, dy = bx - ax, by - ay
    if dx == 0 and dy == 0:
        return math.hypot(ax, ay)
    t = max(0.0, min(1.0, -(ax * dx + ay * dy) / (dx * dx + dy * dy)))
    return math.hypot(ax + t * dx, ay + t * dy)


# ---------- 取得 ----------
def reverse_geocode(lat, lon):
    url = ("https://mreversegeocoder.gsi.go.jp/reverse-geocoder/LonLatToAddress"
           "?lat=%s&lon=%s" % (lat, lon))
    with urllib.request.urlopen(url, timeout=30) as r:
        res = json.load(r).get("results") or {}
    return res.get("muniCd"), res.get("lv01Nm")


def download_a13(pref):
    os.makedirs(CACHE_DIR, exist_ok=True)
    for ver in A13_VERSIONS:
        path = os.path.join(CACHE_DIR, "A13-%s_%02d.zip" % (ver, pref))
        if os.path.exists(path) and os.path.getsize(path) > 100000:
            return path, ver
        url = A13_URL.format(ver=ver, pref=pref)
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "battery-storage-dd/1.0"})
            with urllib.request.urlopen(req, timeout=600) as r:
                data = r.read()
        except Exception:
            continue
        if len(data) < 100000:
            continue
        with open(path, "wb") as f:
            f.write(data)
        return path, ver
    raise SystemExit("A13データを取得できませんでした (pref=%s)" % pref)


def extract_layers(zip_path):
    """zip内のshpをレイヤ番号(ファイル名末尾2桁)ごとに展開し、basename を返す"""
    out = {}
    with zipfile.ZipFile(zip_path) as z:
        names = z.namelist()
        bases = sorted({n[:-4] for n in names if n.lower().endswith(".shp")})
        dest = zip_path[:-4] + "_shp"
        os.makedirs(dest, exist_ok=True)
        for base in bases:
            key = base[-2:]
            if key not in LAYERS:
                continue
            for ext in (".shp", ".shx", ".dbf", ".prj"):
                member = base + ext
                if member in names:
                    target = os.path.join(dest, os.path.basename(member))
                    if not os.path.exists(target):
                        with z.open(member) as src, open(target, "wb") as dst:
                            dst.write(src.read())
            out[key] = os.path.join(dest, os.path.basename(base))
    return out


# ---------- 判定 ----------
def judge_layer(path, lat, lon, radius_m=1000.0):
    import shapefile  # pyshp
    r = shapefile.Reader(path, encoding="cp932")
    fields = [f[0] for f in r.fields[1:]]
    hits, near = [], []
    for sr in r.iterShapeRecords():
        sh = sr.shape
        if not sh.points:
            continue
        bb = sh.bbox
        if not (bb[0] - 0.02 <= lon <= bb[2] + 0.02 and bb[1] - 0.02 <= lat <= bb[3] + 0.02):
            continue
        rings = rings_of(sh)
        outers = [g for g in rings if ring_area(g) < 0]   # shapefile: 外周=時計回り
        inners = [g for g in rings if ring_area(g) >= 0]  # 穴=反時計回り
        inside = any(pip_ring(lon, lat, g) for g in outers)
        if inside and any(pip_ring(lon, lat, g) for g in inners):
            inside = False
        dist = min(seg_dist_m(lon, lat, g[i], g[i + 1])
                   for g in rings for i in range(len(g) - 1))
        rec = dict(zip(fields, [str(v) if not isinstance(v, (int, float)) else v
                                for v in list(sr.record)]))
        rec["_distance_m"] = round(dist, 1)
        if inside:
            hits.append(rec)
        elif dist <= radius_m:
            near.append(rec)
    near.sort(key=lambda d: d["_distance_m"])
    return {"inside": bool(hits), "hits": hits, "near": near[:5]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--lat", type=float, required=True)
    ap.add_argument("--lon", type=float, required=True)
    ap.add_argument("--pref", type=int, default=None, help="都道府県コード(1-47)")
    ap.add_argument("--radius", type=float, default=1000.0, help="近傍として報告する距離[m]")
    ap.add_argument("--json", default=None, help="結果をJSONで保存")
    a = ap.parse_args()

    muni_cd, aza = reverse_geocode(a.lat, a.lon)
    pref = a.pref or (int(muni_cd[:2]) if muni_cd else None)
    if not pref:
        raise SystemExit("都道府県コードを判定できません。--pref を指定してください。")

    print("対象: %.6f, %.6f  (市区町村コード %s / %s)" % (a.lat, a.lon, muni_cd, aza or "-"))
    zip_path, ver = download_a13(pref)
    print("使用データ: 国土数値情報 森林地域 A13-%s (都道府県コード %02d)" % (ver, pref))
    paths = extract_layers(zip_path)

    result = {"lat": a.lat, "lon": a.lon, "muni_cd": muni_cd, "aza": aza,
              "source": "A13-%s_%02d" % (ver, pref), "layers": {}}
    for key in ("10", "08", "09", "07"):
        if key not in paths:
            continue
        name = LAYERS[key]
        res = judge_layer(paths[key], a.lat, a.lon, a.radius)
        result["layers"][name] = res
        mark = "★該当" if res["inside"] else "  非該当"
        print("%s  %s" % (mark, name))
        if res["inside"]:
            for h in res["hits"]:
                print("        面積 %s ha / 年度 %s / 境界まで %.1f m"
                      % (h.get("AREA_SIZE"), h.get("FIS_YEAR"), h["_distance_m"]))
        elif res["near"]:
            n = res["near"][0]
            print("        最寄り %s m (面積 %s ha / 年度 %s)"
                  % (n["_distance_m"], n.get("AREA_SIZE"), n.get("FIS_YEAR")))

    if a.json:
        with open(a.json, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print("JSON出力: %s" % a.json)

    print("\n※A13は参考表示データ(精度保証なし)。境界至近の場合は必ず県の農林水産事務所で確認すること。")


if __name__ == "__main__":
    main()
