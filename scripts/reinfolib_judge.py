#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
不動産情報ライブラリ(reinfolib)API の GeoJSON から、候補地の「該当/非該当」と詳細を判定する。
(1)取得URL生成、(2)APIからのGeoJSON取得（fetch / run のみネットワークを使う）、
(3)範囲判定→値・許認可の組み立て を担う。judge だけなら取得済みGeoJSONでネットワーク無しで動く。

■ v2（2026-11-20）判定の単位を「点」から「敷地（範囲）」に変更
  従来は候補地の緯度経度1点が入る1タイルだけを取得し、その点がポリゴンの内側かだけを見ていた。
  reinfolibのタイルは z=14 で約2km角・z=15 で約1km角なので、候補地がタイルの端にあると
  数十m先のハザードが手元のデータに存在せず、同じ地番の別の端がハザードにかかっていても
  「該当なし」と出てしまう（例：土岐市妻木町は土砂災害警戒区域の境界まで2mだが点では該当なし）。
  v2では
    - 敷地を ①中心＋半径（既定 高圧2MW 50m／特高 150m）または ②筆ポリゴン(GeoJSON) で表す
    - 敷地＋近接マージン（既定100m）の矩形を覆う全タイルを取得する
    - 敷地を格子状に標本化して面積比を出し、該当なしのときは敷地境界からの最短距離を出す
    - 判定は 敷地全体／一部該当（面積比）／近接（最短距離）／該当なし の4段階。
      一部該当も「該当」として扱う（安全側。面積比と最短距離を併記して人が配置で避けられるか判断する）
  --radius 0 を渡すと従来どおりの点判定になる。

使い方:
  # 取得＋判定を一度に行う（APIキーは環境変数 REINFOLIB_API_KEY か --key で渡す。キーはリポジトリに置かない）
  REINFOLIB_API_KEY=... python reinfolib_judge.py run --lat L --lon N --out values.json
  REINFOLIB_API_KEY=... python reinfolib_judge.py run --lat L --lon N --pattern ehv --out values.json
  REINFOLIB_API_KEY=... python reinfolib_judge.py run --lat L --lon N --parcel fude.geojson --out values.json
  # 個別に行う場合
  python reinfolib_judge.py urls  --lat L --lon N --key <APIKEY>
  python reinfolib_judge.py fetch --lat L --lon N --dir <reinfolibdir> [--key <APIKEY>]
  python reinfolib_judge.py judge --lat L --lon N --dir <reinfolibdir> --out values.json
出力 values.json:
  {"site": {...敷地の条件...},
   "values": {DD項目No: {"value","comment","level"}},
   "permits": {法令No: {"req","note"}},
   "coverage": {レイヤコード: {"desc","status","ratio_pct","dist_m","point","tiles"}},
   "raw": {...}, "missing": [...]}
"""
import argparse, glob, json, math, os, sys, tempfile, time, urllib.error, urllib.request

BASE = "https://www.reinfolib.mlit.go.jp/ex-api/external/{code}?response_format=geojson&z={z}&x={x}&y={y}"

# code -> (zoom, 説明)
LAYERS = {
    "XKT001": (14, "都市計画区域/区域区分"),
    "XKT002": (14, "用途地域"),
    "XKT025": (14, "液状化の発生傾向"),
    "XKT026": (15, "洪水浸水想定区域(想定最大)"),
    "XKT027": (15, "高潮浸水想定区域"),
    "XKT028": (15, "津波浸水想定"),
    "XKT029": (14, "土砂災害警戒区域"),
    "XKT019": (14, "自然公園地域"),
    "XKT021": (14, "地すべり防止地区"),
    "XKT022": (14, "急傾斜地崩壊危険区域"),
    "XKT016": (14, "災害危険区域"),
    "XKT020": (14, "大規模盛土造成地"),
}

FLOOD_DEPTH = {1: "0.5m未満", 2: "0.5〜3m", 3: "3〜5m", 4: "5〜10m", 5: "10〜20m", 6: "20m〜"}
SEDIMENT_PHENOM = {1: "急傾斜地の崩壊", 2: "土石流", 3: "地すべり"}
SEDIMENT_KUBUN = {1: "土砂災害警戒区域", 2: "土砂災害特別警戒区域"}

# 敷地の既定の広さ（案件パターン別）。高圧2MWは約1,500〜2,000㎡、特高は2〜3ha級
PATTERN_RADIUS = {"hv": 50.0, "ehv": 150.0}
DEFAULT_MARGIN_M = 100.0   # 敷地の外でもこの距離以内なら「近接」として警告する
SAMPLE_MIN = 800           # 敷地を標本化する点の最低数


# ───────────────────────── タイル・幾何の基本 ─────────────────────────

def deg2num(lat, lon, z):
    n = 2 ** z
    x = int((lon + 180.0) / 360.0 * n)
    y = int((1.0 - math.asinh(math.tan(math.radians(lat))) / math.pi) / 2.0 * n)
    return x, y


def m_per_deg(lat):
    """この緯度での 経度1度・緯度1度 のメートル換算（局所平面近似）"""
    return 111320.0 * math.cos(math.radians(lat)), 110540.0


def pip_ring(lon, lat, ring):
    inside = False; n = len(ring); j = n - 1
    for i in range(n):
        xi, yi = ring[i][0], ring[i][1]; xj, yj = ring[j][0], ring[j][1]
        if ((yi > lat) != (yj > lat)) and (lon < (xj - xi) * (lat - yi) / (yj - yi) + xi):
            inside = not inside
        j = i
    return inside


def pip_polygon(lon, lat, poly):
    if not poly or not pip_ring(lon, lat, poly[0]):
        return False
    return not any(pip_ring(lon, lat, h) for h in poly[1:])


def polys_of(geom):
    """Polygon/MultiPolygon を [polygon, ...] に揃える（polygon = [外周ring, 穴ring, ...]）"""
    if not geom:
        return []
    t = geom.get("type"); c = geom.get("coordinates")
    if t == "Polygon":
        return [c]
    if t == "MultiPolygon":
        return list(c)
    if t == "GeometryCollection":
        out = []
        for g in geom.get("geometries") or []:
            out.extend(polys_of(g))
        return out
    return []


def feature_contains(lon, lat, geom):
    return any(pip_polygon(lon, lat, p) for p in polys_of(geom))


def bbox_of_polys(polys):
    xs = []; ys = []
    for poly in polys:
        for ring in poly:
            for pt in ring:
                xs.append(pt[0]); ys.append(pt[1])
    if not xs:
        return None
    return (min(xs), min(ys), max(xs), max(ys))


def seg_dist_m(lon, lat, a, b, kx, ky):
    """点(lon,lat) から線分 a-b までの距離（m）"""
    ax, ay = (a[0] - lon) * kx, (a[1] - lat) * ky
    bx, by = (b[0] - lon) * kx, (b[1] - lat) * ky
    dx, dy = bx - ax, by - ay
    L2 = dx * dx + dy * dy
    if L2 == 0.0:
        return math.hypot(ax, ay)
    t = -(ax * dx + ay * dy) / L2
    if t < 0.0:
        t = 0.0
    elif t > 1.0:
        t = 1.0
    return math.hypot(ax + t * dx, ay + t * dy)


def dist_points_to_polys(points, polys, kx, ky, limit):
    """敷地境界の点群から ポリゴン群の外周線 までの最短距離（m）。limit を超えたら limit+1 を返す"""
    best = limit + 1.0
    for poly in polys:
        for ring in poly:
            for i in range(len(ring) - 1):
                a, b = ring[i], ring[i + 1]
                for (lon, lat) in points:
                    d = seg_dist_m(lon, lat, a, b, kx, ky)
                    if d < best:
                        best = d
                        if best <= 0.5:
                            return best
    return best


# ───────────────────────── 敷地（範囲）の組み立て ─────────────────────────

def circle_site(lat, lon, radius):
    return {"mode": ("point" if radius <= 0 else "circle"), "lat": lat, "lon": lon,
            "radius": max(0.0, float(radius)), "polys": None}


def load_parcel(path, lat=None, lon=None):
    """筆ポリゴン（GeoJSON）を読む。法務省地図XMLから作ったポリゴンをそのまま渡せる"""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    polys = []
    t = data.get("type")
    if t == "FeatureCollection":
        for ft in data.get("features") or []:
            polys.extend(polys_of(ft.get("geometry")))
    elif t == "Feature":
        polys.extend(polys_of(data.get("geometry")))
    else:
        polys.extend(polys_of(data))
    if not polys:
        sys.exit(f"{path}: ポリゴンが見つかりません（Polygon/MultiPolygon のGeoJSONを渡してください）")
    bb = bbox_of_polys(polys)
    clat = lat if lat is not None else (bb[1] + bb[3]) / 2.0
    clon = lon if lon is not None else (bb[0] + bb[2]) / 2.0
    return {"mode": "parcel", "lat": clat, "lon": clon, "radius": 0.0, "polys": polys}


def site_bbox(site):
    if site["mode"] == "parcel":
        return bbox_of_polys(site["polys"])
    kx, ky = m_per_deg(site["lat"]); r = site["radius"]
    return (site["lon"] - r / kx, site["lat"] - r / ky, site["lon"] + r / kx, site["lat"] + r / ky)


def site_contains(lon, lat, site):
    if site["mode"] == "parcel":
        return any(pip_polygon(lon, lat, p) for p in site["polys"])
    if site["mode"] == "point":
        return False
    kx, ky = m_per_deg(site["lat"])
    return math.hypot((lon - site["lon"]) * kx, (lat - site["lat"]) * ky) <= site["radius"]


def site_samples(site):
    """敷地を格子状に標本化した点（面積比の計算用）"""
    if site["mode"] == "point":
        return [(site["lon"], site["lat"])], 0.0
    x0, y0, x1, y1 = site_bbox(site)
    kx, ky = m_per_deg(site["lat"])
    n = 60
    pts = []
    while True:
        pts = []
        for i in range(n):
            lon = x0 + (x1 - x0) * (i + 0.5) / n
            for j in range(n):
                lat = y0 + (y1 - y0) * (j + 0.5) / n
                if site_contains(lon, lat, site):
                    pts.append((lon, lat))
        if len(pts) >= SAMPLE_MIN or n >= 240:
            break
        n *= 2
    cell = ((x1 - x0) * kx / n) * ((y1 - y0) * ky / n)
    area = (math.pi * site["radius"] ** 2) if site["mode"] == "circle" else cell * len(pts)
    if not pts:   # 極端に細い筆で格子に乗らない場合は中心点だけ
        pts = [(site["lon"], site["lat"])]
    return pts, area


def site_boundary(site, step=5.0):
    """敷地の外周上の点（最短距離の計算用）"""
    if site["mode"] == "point":
        return [(site["lon"], site["lat"])]
    kx, ky = m_per_deg(site["lat"])
    if site["mode"] == "circle":
        r = site["radius"]
        n = max(36, int(2 * math.pi * r / step))
        return [(site["lon"] + r * math.cos(2 * math.pi * i / n) / kx,
                 site["lat"] + r * math.sin(2 * math.pi * i / n) / ky) for i in range(n)]
    out = []
    for poly in site["polys"]:
        for ring in poly:
            for i in range(len(ring) - 1):
                a, b = ring[i], ring[i + 1]
                dm = math.hypot((b[0] - a[0]) * kx, (b[1] - a[1]) * ky)
                k = max(1, int(dm / step))
                for t in range(k):
                    out.append((a[0] + (b[0] - a[0]) * t / k, a[1] + (b[1] - a[1]) * t / k))
    return out or [(site["lon"], site["lat"])]


def site_tiles(site, z, margin):
    """敷地＋近接マージンの矩形を覆う全タイル"""
    x0, y0, x1, y1 = site_bbox(site)
    kx, ky = m_per_deg(site["lat"])
    x0 -= margin / kx; x1 += margin / kx
    y0 -= margin / ky; y1 += margin / ky
    xa, ya = deg2num(y1, x0, z)   # 北西
    xb, yb = deg2num(y0, x1, z)   # 南東
    return [(x, y) for x in range(min(xa, xb), max(xa, xb) + 1)
            for y in range(min(ya, yb), max(ya, yb) + 1)]


def site_label(site, margin):
    if site["mode"] == "point":
        return "点判定（従来互換・半径0m）"
    if site["mode"] == "circle":
        return f"中心＋半径{site['radius']:.0f}m の円（近接マージン{margin:.0f}m）"
    return f"筆ポリゴン{len(site['polys'])}面（近接マージン{margin:.0f}m）"


# ───────────────────────── 取得 ─────────────────────────

def build_urls(site, key, margin):
    out = []
    for code, (z, _d) in LAYERS.items():
        for (x, y) in site_tiles(site, z, margin):
            url = BASE.format(code=code, z=z, x=x, y=y)
            if key:
                url += "&subscription-key=" + key
            out.append(f"{code}\t{url}")
    return out


def resolve_key(key):
    """--key が空なら環境変数 REINFOLIB_API_KEY を使う。どちらも無ければ None。"""
    return key or os.environ.get("REINFOLIB_API_KEY") or None


def tile_path(directory, code, z, x, y):
    return os.path.join(directory, f"{code}__z{z}_x{x}_y{y}.geojson")


def fetch(site, directory, key, margin=DEFAULT_MARGIN_M, retries=3, timeout=30):
    """敷地＋マージンを覆う全タイルのGeoJSONを <directory>/<code>__z<z>_x<x>_y<y>.geojson に保存する。
    1タイルも取れなかったレイヤは judge 側で「未取得」として扱われる。戻り値は {code: "ok 4/4" ...}。"""
    key = resolve_key(key)
    if not key:
        sys.exit("APIキーがありません。環境変数 REINFOLIB_API_KEY を設定するか --key を指定してください。")
    os.makedirs(directory, exist_ok=True)
    status = {}
    for code, (z, desc) in LAYERS.items():
        tiles = site_tiles(site, z, margin)
        ok = 0; errs = []
        for (x, y) in tiles:
            url = BASE.format(code=code, z=z, x=x, y=y)
            req = urllib.request.Request(url, headers={"Ocp-Apim-Subscription-Key": key,
                                                       "Accept": "application/json"})
            last = None
            for i in range(retries):
                try:
                    with urllib.request.urlopen(req, timeout=timeout) as r:
                        body = r.read().decode("utf-8")
                    json.loads(body)  # JSONとして読めることだけ確認
                    with open(tile_path(directory, code, z, x, y), "w", encoding="utf-8") as f:
                        f.write(body)
                    ok += 1; last = None
                    break
                except urllib.error.HTTPError as e:
                    last = f"HTTP {e.code}"
                    if e.code in (401, 403):
                        sys.exit(f"{code}: {last} — APIキーが無効か期限切れです（キーは本文に出しません）。")
                    if e.code == 404:
                        # 範囲外などで404を返すことがあるため、空のFeatureCollectionとして扱う
                        with open(tile_path(directory, code, z, x, y), "w", encoding="utf-8") as f:
                            f.write('{"type":"FeatureCollection","features":[]}')
                        ok += 1; last = None
                        break
                    if e.code == 429:
                        time.sleep(2 * (i + 1)); continue
                except Exception as e:  # ネットワーク断・JSON不正など
                    last = f"{type(e).__name__}: {e}"
                    time.sleep(1 + i)
            if last:
                errs.append(f"({x},{y}) {last}")
        status[code] = f"ok {ok}/{len(tiles)}タイル" + (" ／ 失敗: " + "; ".join(errs[:3]) if errs else "")
        print(f"  {code} {desc}: {status[code]}")
    return status


def load_layer(directory, code):
    """そのレイヤの全タイルを1つのFeatureCollectionにまとめて返す。(fc, タイル数)。無ければ (None, 0)"""
    paths = sorted(glob.glob(os.path.join(directory, code + "__z*.geojson")))
    legacy = os.path.join(directory, code + ".geojson")   # v1 で取得した単一タイルも読む
    if os.path.exists(legacy):
        paths.append(legacy)
    if not paths:
        return None, 0
    feats = []
    for p in paths:
        try:
            t = open(p, encoding="utf-8").read().strip()
            fc = json.loads(t) if t else {}
        except Exception:
            continue
        feats.extend(fc.get("features") or [])
    return {"type": "FeatureCollection", "features": feats}, len(paths)


# ───────────────────────── 範囲判定 ─────────────────────────

def layer_eval(site, samples, bpts, fc, margin):
    """レイヤ1つを敷地と重ね合わせる。
    返り値: {"ratio": 面積比0-1, "props": 該当した地物の属性, "dist": 最短距離m or None,
             "near": [(距離, 属性)...], "point": 中心点が該当するか}"""
    sb = site_bbox(site)
    kx, ky = m_per_deg(site["lat"])
    mx, my = margin / kx, margin / ky
    hit = [False] * len(samples)
    props_hit = []
    near = []
    point_hit = False
    for f in (fc.get("features") or []):
        g = f.get("geometry"); p = f.get("properties") or {}
        polys = polys_of(g)
        bb = bbox_of_polys(polys)
        if bb is None:
            continue
        covered = 0
        if not (bb[2] < sb[0] or bb[0] > sb[2] or bb[3] < sb[1] or bb[1] > sb[3]):
            for i, (lon, lat) in enumerate(samples):
                if lon < bb[0] or lon > bb[2] or lat < bb[1] or lat > bb[3]:
                    continue
                if any(pip_polygon(lon, lat, pl) for pl in polys):
                    hit[i] = True; covered += 1
            if covered:
                q = dict(p); q["_ratio"] = covered / len(samples) if samples else 0.0
                props_hit.append(q)
            if (bb[0] <= site["lon"] <= bb[2] and bb[1] <= site["lat"] <= bb[3]
                    and any(pip_polygon(site["lon"], site["lat"], pl) for pl in polys)):
                point_hit = True
        if not covered:
            if not (bb[2] < sb[0] - mx or bb[0] > sb[2] + mx or bb[3] < sb[1] - my or bb[1] > sb[3] + my):
                d = dist_points_to_polys(bpts, polys, kx, ky, margin)
                if d <= margin:
                    near.append((d, p))
    ratio = (sum(1 for h in hit if h) / len(hit)) if hit else 0.0
    props_hit.sort(key=lambda q: -q.get("_ratio", 0.0))
    near.sort(key=lambda t: t[0])
    return {"ratio": ratio, "props": props_hit, "near": near,
            "dist": (near[0][0] if near else None), "point": point_hit}


def ev_status(ev):
    """敷地全体 / 一部 / 近接 / なし"""
    if ev["ratio"] >= 0.995:
        return "敷地全体"
    if ev["ratio"] > 0:
        return "一部該当"
    if ev["dist"] is not None:
        return "近接"
    return "該当なし"


def hit(ev):
    return ev["ratio"] > 0


def area_text(ev):
    st = ev_status(ev)
    if st == "敷地全体":
        return "敷地全体"
    if st == "一部該当":
        return f"敷地の約{max(1, round(ev['ratio'] * 100))}%"
    if st == "近接":
        return f"敷地外・最短{round(ev['dist'])}m"
    return ""


def share_text(q, ev):
    """レイヤ内の1区分の面積比（敷地全体を覆っていれば「敷地全体」）"""
    r = q.get("_ratio", 0.0)
    if r >= 0.995:
        return "敷地全体"
    return f"敷地の約{max(1, round(r * 100))}%"


def near_text(ev, label):
    return f"{label}が{round(ev['dist'])}m先" if ev["dist"] is not None else ""


def best_text(props_list):
    if not props_list:
        return ""
    p = props_list[0]; skip = ("code", "id", "date", "number", "year", "size", "leng", "area", "no", "div")
    picks = [v.strip() for k, v in p.items()
             if isinstance(v, str) and v.strip() and not any(s in k.lower() for s in skip)]
    return " / ".join(dict.fromkeys(picks))[:60]


def judge(site, directory, margin=DEFAULT_MARGIN_M):
    values, permits, raw, coverage = {}, {}, {}, {}
    samples, area = site_samples(site)
    bpts = site_boundary(site)

    fcs = {}; tiles = {}
    for code in LAYERS:
        fc, n = load_layer(directory, code)
        fcs[code] = fc; tiles[code] = n

    ev = {}
    for code in LAYERS:
        if fcs[code] is None:
            ev[code] = None
            continue
        e = layer_eval(site, samples, bpts, fcs[code], margin)
        ev[code] = e
        raw[code] = e["props"]
        coverage[code] = {"desc": LAYERS[code][1], "status": ev_status(e),
                          "ratio_pct": round(e["ratio"] * 100, 1),
                          "dist_m": (round(e["dist"], 1) if (e["dist"] is not None and e["ratio"] == 0) else None),
                          "point": e["point"], "tiles": tiles[code],
                          "parts": [{"label": best_text([q])[:40], "ratio_pct": round(q.get("_ratio", 0) * 100, 1)}
                                    for q in e["props"][:6]]}

    def part_note(code):
        """点では該当しないが敷地の一部が該当する場合の注記（点判定との違いを明示）"""
        e = ev[code]
        if e and e["ratio"] > 0 and not e["point"]:
            return "（候補地の点そのものは区域外。敷地の範囲で該当）"
        return ""

    # 用途地域(10)
    if ev["XKT002"] is not None:
        e = ev["XKT002"]
        if hit(e):
            names = []
            for p in e["props"]:
                nm = p.get("use_area_ja") or "用途地域"
                bits = []
                if p.get("u_building_coverage_ratio_ja"): bits.append("建蔽率" + str(p["u_building_coverage_ratio_ja"]))
                if p.get("u_floor_area_ratio_ja"): bits.append("容積率" + str(p["u_floor_area_ratio_ja"]))
                names.append(nm + "（" + "・".join([x for x in [" ".join(bits), share_text(p, e)] if x] + []) + "）")
            val = " / ".join(dict.fromkeys(names))
            if len(dict.fromkeys(names)) > 1:
                val += "　※敷地が複数の用途地域にまたがる"
            lvl = 1
        elif e["dist"] is not None:
            nm = (e["near"][0][1].get("use_area_ja") or "用途地域")
            val = f"指定なし（{nm}が{round(e['dist'])}m先）"
            lvl = 1
        else:
            val = "指定なし"; lvl = 2
        values["10"] = {"value": val, "level": lvl,
                        "comment": "reinfolib XKT002より範囲判定" + part_note("XKT002")}

    # 区域区分(8,9)
    if ev["XKT001"] is not None:
        e = ev["XKT001"]
        vals = [str(p.get("area_classification_ja") or "") for p in e["props"]]
        joined = " / ".join(dict.fromkeys([v for v in vals if v]))
        kasenka = any("市街化区域" in v for v in vals)
        chosei = any("市街化調整区域" in v for v in vals)
        shares = " ＋ ".join(dict.fromkeys(
            [f"{str(p.get('area_classification_ja') or '区域区分なし')}（{share_text(p, e)}）" for p in e["props"]]))
        cmt = "reinfolib XKT001より範囲判定" + (f"（{shares}）" if shares
                                            else "（都市計画区域外の可能性）") + part_note("XKT001")
        if kasenka and chosei:
            cmt += "　※敷地が市街化区域と市街化調整区域にまたがる（開発許可の扱いが分かれるため要確認）"
        values["8"] = {"value": "指定あり" if kasenka else "指定なし", "level": 2, "comment": cmt}
        values["9"] = {"value": "指定あり" if chosei else "指定なし",
                       "level": 1 if chosei else 2, "comment": cmt}

    # 浸水(4): 洪水/高潮/津波
    FLOODS = (("XKT026", "洪水浸水想定区域"), ("XKT027", "高潮浸水想定区域"), ("XKT028", "津波浸水想定"))
    if any(ev[c] is not None for c, _l in FLOODS):
        parts, nears, notes = [], [], []
        for code, lab in FLOODS:
            e = ev[code]
            if e is None:
                continue
            if hit(e):
                d = []
                if code == "XKT026":
                    ranks = [p.get("A31a_205") for p in e["props"] if isinstance(p.get("A31a_205"), int)]
                    rivers = sorted({str(p.get("A31a_202")) for p in e["props"] if p.get("A31a_202")})
                    if rivers: d.append("・".join(rivers))
                    if ranks: d.append("浸水深" + FLOOD_DEPTH.get(max(ranks), ""))
                d.append(area_text(e))
                parts.append(lab + "（" + "／".join([x for x in d if x]) + "）")
                notes.append(part_note(code))
            elif e["dist"] is not None:
                nears.append(near_text(e, lab))
        if parts:
            val = " / ".join(parts); lvl = 0
        elif nears:
            val = "該当なし（" + "／".join(nears) + "）"; lvl = 1
        else:
            val = "該当なし"; lvl = 2
        values["4"] = {"value": val, "level": lvl,
                       "comment": "reinfolib 洪水/高潮/津波 浸水想定区域より範囲判定"
                                  + "".join(dict.fromkeys([n for n in notes if n]))}

    # 土砂(5)
    sed_extra = []
    if ev["XKT029"] is not None:
        e = ev["XKT029"]
        if hit(e):
            labels = []; tokubetsu = False
            for p in e["props"]:
                k = SEDIMENT_KUBUN.get(p.get("A33_002"), "土砂災害（警戒区域）")
                if p.get("A33_002") == 2:
                    tokubetsu = True
                ph = SEDIMENT_PHENOM.get(p.get("A33_001"), "")
                labels.append(k + (f"（{ph}）" if ph else ""))
            uniq = sorted(set(labels), key=lambda s: ("特別警戒" not in s, s))
            values["5"] = {"value": " / ".join(uniq) + f"（{area_text(e)}）"
                                    + (f"・区域{len(e['props'])}件" if len(e["props"]) > 1 else ""), "level": 0,
                           "comment": "reinfolib XKT029より範囲判定" + part_note("XKT029")}
            permits["16"] = {"req": "要", "note": f"土砂災害危険箇所/警戒区域に該当（{area_text(e)}・reinfolib XKT029）"}
            if tokubetsu:
                permits["15"] = {"req": "要", "note": f"土砂災害特別警戒区域に該当（{area_text(e)}・reinfolib XKT029）"}
        elif e["dist"] is not None:
            lab = SEDIMENT_KUBUN.get(e["near"][0][1].get("A33_002"), "土砂災害警戒区域")
            values["5"] = {"value": f"該当なし（{lab}が{round(e['dist'])}m先）", "level": 1,
                           "comment": "reinfolib XKT029より範囲判定。敷地の外だが近接しているため配置と区域図を確認"}
            permits["16"] = {"req": "要確認", "note": f"土砂災害警戒区域が敷地から{round(e['dist'])}m（reinfolib XKT029）"}
        else:
            values["5"] = {"value": "該当なし", "level": 2, "comment": "reinfolib XKT029より範囲判定"}

    # 液状化(6)
    if ev["XKT025"] is not None:
        e = ev["XKT025"]
        if hit(e):
            p = e["props"][0]; note = p.get("note") or ""
            micro = p.get("topographic_classification_name_ja") or ""
            lv = p.get("liquefaction_tendency_level")
            levels = [x.get("liquefaction_tendency_level") for x in e["props"]
                      if isinstance(x.get("liquefaction_tendency_level"), int)]
            extra = [x for x in [micro, (f"傾向{max(levels)}/6" if levels else
                                         (f"傾向{lv}/6" if isinstance(lv, int) else ""))] if x]
            val = (note + ("（" + "・".join(extra) + "）" if extra else "")) or "該当（傾向図あり）"
            if ev_status(e) == "一部該当" and len(e["props"]) > 1:
                val += f"（敷地内に{len(e['props'])}区分・最も強い傾向を表示）"
            lvl = 1
        else:
            val = "該当なし（対象範囲外）"; lvl = 2
        values["6"] = {"value": val, "level": lvl,
                       "comment": "reinfolib XKT025（地形区分に基づく液状化発生傾向）より範囲判定"}

    # 公園(13) 自然公園
    if ev["XKT019"] is not None:
        e = ev["XKT019"]
        if hit(e):
            names = sorted({str(p.get("OBJ_NAME_ja")) for p in e["props"] if p.get("OBJ_NAME_ja")})
            base = "・".join(names) if names else "自然公園地域に該当"
            values["13"] = {"value": f"{base}（{area_text(e)}）", "level": 0,
                            "comment": "reinfolib XKT019（自然公園地域）より範囲判定。特別地域/普通地域の区別は"
                                       "国土数値情報A10か現行の区域図で確認" + part_note("XKT019")}
        elif e["dist"] is not None:
            values["13"] = {"value": f"該当なし（自然公園地域が{round(e['dist'])}m先）", "level": 1,
                            "comment": "reinfolib XKT019（自然公園地域）より範囲判定"}
        else:
            values["13"] = {"value": "該当なし", "level": 2,
                            "comment": "reinfolib XKT019（自然公園地域）より範囲判定"}

    # 防災補足 → 許認可＋item5コメント
    EXTRA = (("XKT021", "地すべり防止地区", "13"), ("XKT022", "急傾斜地崩壊危険区域", "14"),
             ("XKT016", "災害危険区域", None), ("XKT020", "大規模盛土造成地", "7"))
    for code, lab, permit_no in EXTRA:
        e = ev[code]
        if e is None:
            continue
        if hit(e):
            sed_extra.append(f"{lab}に該当（{area_text(e)}）"
                             + ("（建築制限の可能性）" if code == "XKT016" else ""))
            if permit_no:
                permits[permit_no] = {"req": "要", "note": f"{lab}に該当（{area_text(e)}・reinfolib {code}）"}
        elif e["dist"] is not None:
            sed_extra.append(f"{lab}が{round(e['dist'])}m先")
            if permit_no:
                permits[permit_no] = {"req": "要確認",
                                      "note": f"{lab}が敷地から{round(e['dist'])}m（reinfolib {code}）"}
    if sed_extra and "5" in values:
        values["5"]["comment"] += " ／ " + "、".join(sed_extra)
        worst = 0 if any("に該当" in x for x in sed_extra) else 1
        values["5"]["level"] = min(values["5"]["level"], worst)

    missing = [c for c in LAYERS if fcs[c] is None]
    site_out = {"mode": site["mode"], "lat": site["lat"], "lon": site["lon"],
                "radius_m": site["radius"] if site["mode"] == "circle" else None,
                "parcel_polygons": (len(site["polys"]) if site["mode"] == "parcel" else None),
                "margin_m": margin, "area_m2": round(area), "sample_points": len(samples),
                "label": site_label(site, margin),
                "tiles_total": sum(tiles.values())}
    return {"site": site_out, "values": values, "permits": permits,
            "coverage": coverage, "raw": raw, "missing": missing}


# ───────────────────────── CLI ─────────────────────────

def add_site_args(p, with_dir=True):
    p.add_argument("--lat", type=float)
    p.add_argument("--lon", type=float)
    p.add_argument("--parcel", default="", help="筆ポリゴンのGeoJSON（指定すると半径より優先）")
    p.add_argument("--radius", type=float, default=None,
                   help="敷地の半径m（既定: 高圧2MW 50m／--pattern ehv で150m。0で従来の点判定）")
    p.add_argument("--pattern", choices=["hv", "ehv"], default="hv",
                   help="案件パターン hv=高圧2MW（既定）／ehv=特高")
    p.add_argument("--margin", type=float, default=DEFAULT_MARGIN_M,
                   help=f"敷地の外で「近接」とみなす距離m（既定{DEFAULT_MARGIN_M:.0f}）")


def site_from_args(a):
    if a.parcel:
        return load_parcel(a.parcel, a.lat, a.lon)
    if a.lat is None or a.lon is None:
        sys.exit("--lat / --lon（または --parcel）が必要です。")
    r = a.radius if a.radius is not None else PATTERN_RADIUS[a.pattern]
    return circle_site(a.lat, a.lon, r)


def main():
    ap = argparse.ArgumentParser(); sub = ap.add_subparsers(dest="cmd", required=True)
    pu = sub.add_parser("urls"); add_site_args(pu); pu.add_argument("--key", default="")
    pj = sub.add_parser("judge"); add_site_args(pj); pj.add_argument("--dir", required=True); pj.add_argument("--out", required=True)
    pf = sub.add_parser("fetch", help="APIから敷地を覆う全タイルのGeoJSONを取得して --dir に保存")
    add_site_args(pf); pf.add_argument("--dir", required=True); pf.add_argument("--key", default="")
    pr = sub.add_parser("run", help="fetch と judge を一度に実行")
    add_site_args(pr); pr.add_argument("--out", required=True); pr.add_argument("--dir", default=""); pr.add_argument("--key", default="")
    a = ap.parse_args()
    site = site_from_args(a)
    if a.cmd == "urls":
        for ln in build_urls(site, a.key, a.margin):
            print(ln)
        return
    print("判定の範囲:", site_label(site, a.margin))
    if a.cmd == "fetch":
        fetch(site, a.dir, a.key, a.margin)
        return
    if a.cmd == "run":
        a.dir = a.dir or tempfile.mkdtemp(prefix="reinfolib_")
        print("取得先:", a.dir)
        fetch(site, a.dir, a.key, a.margin)
    res = judge(site, a.dir, a.margin)
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=2)
    print("OK:", a.out)
    s = res["site"]
    print(f"  敷地: {s['label']} ／ 面積 約{s['area_m2']:,}㎡ ／ 標本{s['sample_points']:,}点 ／ 取得タイル{s['tiles_total']}枚")
    for k in sorted(res["values"], key=lambda s: int(s)):
        v = res["values"][k]
        mark = {0: "✕", 1: "△", 2: "○"}.get(v.get("level"), "?")
        print(f"  {mark} DD項目{k}: {v['value']}")
    for k in sorted(res["permits"], key=lambda s: int(s)):
        print(f"  許認可No.{k}: {res['permits'][k]['req']} - {res['permits'][k]['note']}")
    diff = [f"{c}（{d['status']}）" for c, d in res["coverage"].items()
            if d["status"] == "一部該当" and not d["point"]]
    if diff:
        print("  ※点判定では見落とす該当:", "、".join(diff))
    if res["missing"]:
        print("  未取得:", ", ".join(res["missing"]))


if __name__ == "__main__":
    main()
