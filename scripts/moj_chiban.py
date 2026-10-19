# -*- coding: utf-8 -*-
"""
緯度経度 → 地番 を推定し、eMAFF農地ナビで地番検索するためのリンクとメモを作る。

データ源
  法務省 登記所備付地図データ（G空間情報センター、市区町村ごとのzip）
  農林水産省 筆ポリゴン（Pacific Spatial Solutions のGeoParquet。任意座標系の地図の位置合わせに使用）

手順
  1. 国土地理院の逆ジオコーダで市区町村コード・大字名を取得
  2. G空間情報センターから、その市区町村の最新の登記所備付地図zipを取得（--cache に保存）
  3. 大字に該当する地図XMLを読み込む
  4. 公共座標系の地図 → 緯度経度を平面直角座標に変換し、点を含む筆の地番を返す（確定に近い）
     任意座標系の地図 → 筆ポリゴン（田・畑）に形を合わせて位置を推定し、点の近くの筆を返す（推定）
  5. eMAFF農地ナビの検索語・メモ・確認用の重ね合わせ画像を出力

使い方
  python scripts/moj_chiban.py --lat 35.38001 --lon 137.06507 --cache <作業dir>/moj --out chiban.json [--png overlay.png]
  出来た chiban.json は build_px_checksheet.py --chiban chiban.json でチェックシートに反映できる。

必要なもの: pip install pyproj shapely scipy duckdb pillow
"""
import argparse, csv, glob, io, json, math, os, re, sys, unicodedata, urllib.parse, urllib.request, zipfile
from xml.etree import ElementTree as ET

EMAFF_URL = "https://map.maff.go.jp/"
GSI_REVERSE = "https://mreversegeocoder.gsi.go.jp/reverse-geocoder/LonLatToAddress?lat={lat}&lon={lon}"
CKAN_SEARCH = "https://www.geospatial.jp/ckan/api/3/action/package_search?q={q}&fq=title:%E7%99%BB%E8%A8%98%E6%89%80%E5%82%99%E4%BB%98%E5%9C%B0%E5%9B%B3&rows=20"
FUDE_PARQUET = "https://data.source.coop/pacificspatial/field-polygon-jp/parquet/jp_field_polygons_2024.parquet"
NS = {'m': 'http://www.moj.go.jp/MINJI/tizuxml', 'z': 'http://www.moj.go.jp/MINJI/tizuzumen'}
KANJI_NUM = "〇一二三四五六七八九"
# 「公共座標7系」→ JGD2011 平面直角座標系 第VII系 (EPSG:6675)。I系=6669。
EPSG_BASE = 6668

# 筆ポリゴンへの位置合わせの合格ライン（地図の農地のうち筆ポリゴンに入る割合）
FIT_ACCEPT = 0.45


def fetch(url, timeout=120):
    req = urllib.request.Request(url, headers={"User-Agent": "battery-storage-dd/moj_chiban"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def norm_name(s):
    """表記ゆれ（薮/藪、ケ/ヶ、全角数字）を吸収した比較用の文字列"""
    s = unicodedata.normalize("NFKC", s or "")
    for a, b in (("藪", "薮"), ("ヶ", "ケ"), ("ヵ", "ケ"), ("之", "の"), (" ", ""), ("　", "")):
        s = s.replace(a, b)
    return s


def reverse_geocode(lat, lon):
    d = json.loads(fetch(GSI_REVERSE.format(lat=lat, lon=lon)))["results"]
    return d["muniCd"], d["lv01Nm"]


# ── 登記所備付地図の取得 ──
def find_dataset_zips(muni_cd):
    """市区町村コードで始まるzipのうち、データセットごとに最新年のものを返す"""
    res = json.loads(fetch(CKAN_SEARCH.format(q=muni_cd)))["result"]["results"]
    latest = {}
    for p in res:
        for r in p["resources"]:
            m = re.match(rf"({muni_cd}-\d+)-(\d{{4}})\.zip$", r.get("name", ""))
            if m:
                key, year = m.group(1), int(m.group(2))
                if key not in latest or year > latest[key][0]:
                    latest[key] = (year, r["url"], p["title"])
    return [(k, v[0], v[1], v[2]) for k, v in sorted(latest.items())]


def download_zip(url, path):
    if not os.path.exists(path):
        data = fetch(url, timeout=600)
        if not data.startswith(b"PK"):
            # CKANはS3への転送ページを返すことがある
            m = re.search(rb'href="([^"]+)"', data)
            if not m:
                raise RuntimeError("zipを取得できませんでした: " + url)
            data = fetch(m.group(1).decode().replace("&amp;", "&"), timeout=600)
        open(path, "wb").write(data)
    return path


def xmls_for_oaza(zpath, oaza, workdir):
    """search-list.csv で大字に該当する内側zipを特定し、XMLを展開して返す"""
    out = []
    with zipfile.ZipFile(zpath) as z:
        lst = [n for n in z.namelist() if n.endswith("search-list.csv")]
        if not lst:
            return out
        text = z.read(lst[0]).decode("cp932", errors="replace")
        target = norm_name(oaza)
        inner = sorted({row[0] for row in csv.reader(io.StringIO(text))
                        if len(row) > 2 and norm_name(row[2]).endswith(target)})
        for name in inner:
            dst = os.path.join(workdir, name[:-4])
            if not os.path.isdir(dst):
                with zipfile.ZipFile(io.BytesIO(z.read(name))) as iz:
                    iz.extractall(dst)
            out += glob.glob(os.path.join(dst, "**", "*.xml"), recursive=True)
    return out


def parse_moj_xml(path, oaza):
    """登記所備付地図XMLを読み、(座標系, メタ情報, 筆のリスト) を返す。筆の ring は (X=北, Y=東)"""
    root = ET.parse(path).getroot()
    zahyo = root.findtext("m:座標系", namespaces=NS) or ""
    meta = {k: root.findtext(f".//m:{k}", namespaces=NS) for k in ("地図名", "地図種類", "地図分類", "縮尺分母")}
    pts = {}
    for pt in root.iter("{%s}GM_Point" % NS["z"]):
        d = pt.find(".//z:DirectPosition", NS)
        if d is not None:
            pts[pt.get("id")] = (float(d.findtext("z:X", namespaces=NS)), float(d.findtext("z:Y", namespaces=NS)))
    curves = {}
    for c in root.iter("{%s}GM_Curve" % NS["z"]):
        seq = []
        for col in c.iter("{%s}GM_PointArray.column" % NS["z"]):
            d = col.find("z:GM_Position.direct", NS)
            if d is not None:
                seq.append((float(d.findtext("z:X", namespaces=NS)), float(d.findtext("z:Y", namespaces=NS))))
                continue
            ind = col.find("z:GM_Position.indirect/z:GM_PointRef.point", NS)
            if ind is not None and ind.get("idref") in pts:
                seq.append(pts[ind.get("idref")])
        curves[c.get("id")] = seq
    surfaces = {}
    for s in root.iter("{%s}GM_Surface" % NS["z"]):
        ext = s.find(".//z:GM_SurfaceBoundary.exterior", NS)
        if ext is None:
            continue
        ring = []
        for g in ext.iter("{%s}GM_CompositeCurve.generator" % NS["z"]):
            seg = curves.get(g.get("idref"))
            if not seg:
                continue
            if ring and ring[-1] == seg[-1]:
                seg = list(reversed(seg))
            ring.extend(seg[1:] if ring and ring[-1] == seg[0] else seg)
        if ring and ring[0] != ring[-1]:
            ring.append(ring[0])
        surfaces[s.get("id")] = ring
    target = norm_name(oaza)
    parcels = []
    for h in root.iter("{%s}筆" % NS["m"]):
        if not norm_name(h.findtext("m:大字名", namespaces=NS)).endswith(target):
            continue
        shape = h.find("m:形状", NS)
        ring = surfaces.get(shape.get("idref")) if shape is not None else None
        if ring and len(ring) >= 4:
            parcels.append({"大字": h.findtext("m:大字名", namespaces=NS),
                            "地番": h.findtext("m:地番", namespaces=NS) or "",
                            "ring": ring})
    return zahyo, meta, parcels


def chiban_label(c):
    """XML上の地番（例: 1430-2W2）を、登記上の表記（1430番2）に直す"""
    c = re.sub(r"[VW]\d+$", "", c)
    if not re.match(r"^\d", c):
        return None
    head, _, tail = c.partition("-")
    return f"{head}番{tail}" if tail else f"{head}番"


def is_land(c):
    return bool(re.match(r"^\d", c))


# ── 位置の推定 ──
def public_system_no(zahyo):
    m = re.search(r"公共座標(\d+)系", unicodedata.normalize("NFKC", zahyo))
    return int(m.group(1)) if m else None


def local_crs(lat, lon):
    # 筆ポリゴンとの位置合わせ用。任意座標系はメートル単位なので、地点中心の正距方位図法で十分
    return f"+proj=aeqd +lat_0={lat} +lon_0={lon} +ellps=GRS80 +units=m"


def fetch_fude(lat, lon, r_deg=0.02):
    import duckdb
    con = duckdb.connect()
    con.execute("INSTALL httpfs; LOAD httpfs; INSTALL spatial; LOAD spatial;")
    rows = con.execute(f"""select land_type_jp, ST_AsGeoJSON(GEOM) from read_parquet('{FUDE_PARQUET}')
        where point_lat::double between {lat - r_deg} and {lat + r_deg}
          and point_lng::double between {lon - r_deg * 1.25} and {lon + r_deg * 1.25}""").fetchall()
    return [(lt, json.loads(g)) for lt, g in rows]


def fit_to_fude(parcels, fude_xy, pxy, search_m=1800):
    """任意座標系の農地の筆を、筆ポリゴン（平面直角座標）に回転・平行移動・拡縮で重ねる。
    戻り値: (一致率, 変換パラメータ, 中心)。一致率=地図の農地のうち筆ポリゴンに入る割合"""
    import numpy as np
    from PIL import Image, ImageDraw
    from shapely.geometry import Polygon, Point
    from scipy.optimize import minimize
    R, half = 2.0, 3000
    x0, y0 = pxy[0] - half, pxy[1] - half
    W = int(2 * half / R)
    im = Image.new("L", (W, W), 0)
    d = ImageDraw.Draw(im)
    for poly in fude_xy:
        d.polygon([((x - x0) / R, (y0 + 2 * half - y) / R) for x, y in poly], fill=1)
    M = np.array(im, bool)
    samples = []
    for q in parcels:
        if not is_land(q["地番"]):
            continue
        P = Polygon([(p[1], p[0]) for p in q["ring"]]).buffer(0)
        if P.area < 50:
            continue
        b = P.bounds
        for x in np.arange(b[0], b[2], 6):
            for y in np.arange(b[1], b[3], 6):
                if P.contains(Point(x, y)):
                    samples.append((x, y))
    if len(samples) < 50:
        return 0.0, None, None
    S = np.array(samples)
    c = S.mean(0)
    Q = S - c

    def score(p):
        s, th, tx, ty = p
        if not 0.8 < s < 1.2:  # 縮めて1点に潰すと一致率が見かけ上1になるため、縮尺は±20%に制限
            return 0.0
        ct, st = math.cos(math.radians(th)), math.sin(math.radians(th))
        X = s * (ct * Q[:, 0] - st * Q[:, 1]) + tx
        Y = s * (st * Q[:, 0] + ct * Q[:, 1]) + ty
        i = ((X - x0) / R).astype(int)
        j = ((y0 + 2 * half - Y) / R).astype(int)
        ok = (i >= 0) & (i < W) & (j >= 0) & (j < W)
        return M[j[ok], i[ok]].sum() / len(S)

    cand = []
    step = 100
    for th in range(0, 360, 6):
        for tx in np.arange(pxy[0] - search_m, pxy[0] + search_m + 1, step):
            for ty in np.arange(pxy[1] - search_m, pxy[1] + search_m + 1, step):
                cand.append((score([1, th, tx, ty]), th, tx, ty))
    cand.sort(reverse=True)
    best = (0.0, None)
    for sc, th, tx, ty in cand[:12]:
        simplex = [[1, th, tx, ty], [1.05, th, tx, ty], [1, th + 3, tx, ty], [1, th, tx + 40, ty], [1, th, tx, ty + 40]]
        r = minimize(lambda p: -score(p), [1, th, tx, ty], method="Nelder-Mead",
                     options={"xatol": 0.2, "fatol": 1e-4, "maxiter": 800, "initial_simplex": simplex})
        if -r.fun > best[0]:
            best = (-r.fun, list(r.x))
    return best[0], best[1], list(c)


def transform_ring(ring, p, c):
    s, th, tx, ty = p
    ct, st = math.cos(math.radians(th)), math.sin(math.radians(th))
    out = []
    for n, e in ring:
        q0, q1 = e - c[0], n - c[1]
        out.append((s * (ct * q0 - st * q1) + tx, s * (st * q0 + ct * q1) + ty))
    return out


def locate(lat, lon, files, log):
    """各地図から、点を含む（または近い）筆を探す。戻り値: 候補のリスト"""
    from pyproj import Transformer
    from shapely.geometry import Polygon, Point
    cands, fude_cache, fitted = [], {}, []
    for path, zahyo, meta, parcels in files:
        sysno = public_system_no(zahyo)
        if sysno:
            epsg = EPSG_BASE + sysno
            tr = Transformer.from_crs(4326, epsg, always_xy=True)
            pxy = tr.transform(lon, lat)
            pt = Point(pxy)
            for q in parcels:
                P = Polygon([(p[1], p[0]) for p in q["ring"]]).buffer(0)
                d = P.distance(pt)
                if d < 15:
                    cands.append({"地番": q["地番"], "距離m": round(d, 1), "面積m2": round(P.area),
                                  "地図": meta.get("地図名"), "方法": f"公共座標{sysno}系の地図で判定", "確度": "高",
                                  "_xy": [list(xy) for xy in P.exterior.coords], "_epsg": epsg})
            log.append(f"{meta.get('地図名')}: 公共座標{sysno}系 → 座標で直接判定")
            continue
        # 任意座標系 → 筆ポリゴンに位置合わせ
        epsg = local_crs(lat, lon)
        tr = Transformer.from_crs(4326, epsg, always_xy=True)
        pxy = tr.transform(lon, lat)
        if epsg not in fude_cache:
            log.append("筆ポリゴンを取得中（農水省 2024年度）")
            fx = []
            for _lt, g in fetch_fude(lat, lon):
                polys = g["coordinates"] if g["type"] == "MultiPolygon" else [g["coordinates"]]
                for poly in polys:
                    fx.append([tr.transform(x, y) for x, y in poly[0]])
            fude_cache[epsg] = fx
        score, p, c = fit_to_fude(parcels, fude_cache[epsg], pxy)
        name = meta.get("地図名")
        if not p or score < FIT_ACCEPT:
            log.append(f"{name}: 任意座標系（{meta.get('地図種類')}）。筆ポリゴンと形が合わず位置を推定できず（一致率{score:.2f}）")
            continue
        log.append(f"{name}: 任意座標系（{meta.get('地図種類')}）→ 筆ポリゴンに位置合わせ（一致率{score:.2f}、回転{p[1]:.1f}°、縮尺{p[0]:.3f}）")
        fitted.append((name, score))
        pt = Point(pxy)
        for q in parcels:
            P = Polygon(transform_ring(q["ring"], p, c)).buffer(0)
            d = P.distance(pt)
            if d < 15:
                cands.append({"地番": q["地番"], "距離m": round(d, 1), "面積m2": round(P.area),
                              "地図": name, "方法": "筆ポリゴンへの位置合わせで推定", "確度": "中",
                              "_xy": [list(xy) for xy in P.exterior.coords], "_epsg": epsg})
    cands.sort(key=lambda x: x["距離m"])
    return cands


def draw_overlay(lat, lon, cands, path):
    """候補の筆を地理院の航空写真に重ねた確認画像（z18・5×5タイル）"""
    from pyproj import Transformer
    from PIL import Image, ImageDraw, ImageFont
    z, h = 18, 2
    N = 2 ** z

    def tile_xy(lo, la):
        r = math.radians(la)
        return (lo + 180) / 360 * N, (1 - math.log(math.tan(r) + 1 / math.cos(r)) / math.pi) / 2 * N
    fx, fy = tile_xy(lon, lat)
    cx, cy = int(fx), int(fy)
    img = Image.new("RGB", (256 * (2 * h + 1),) * 2)
    for dx in range(-h, h + 1):
        for dy in range(-h, h + 1):
            t = fetch(f"https://cyberjapandata.gsi.go.jp/xyz/seamlessphoto/{z}/{cx + dx}/{cy + dy}.jpg")
            img.paste(Image.open(io.BytesIO(t)), ((dx + h) * 256, (dy + h) * 256))
    d = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf", 11)
    except OSError:
        font = ImageFont.load_default()
    to_px = lambda lo, la: ((tile_xy(lo, la)[0] - cx + h) * 256, (tile_xy(lo, la)[1] - cy + h) * 256)
    for c in cands:
        inv = Transformer.from_crs(c["_epsg"], 4326, always_xy=True)
        pts = [to_px(*inv.transform(x, y)) for x, y in c["_xy"]]
        d.line(pts, fill=(255, 70, 70), width=3)
        if is_land(c["地番"]):
            mx, my = sum(p[0] for p in pts) / len(pts), sum(p[1] for p in pts) / len(pts)
            d.text((mx, my), chiban_label(c["地番"]), fill=(255, 255, 0), font=font, anchor="mm",
                   stroke_width=2, stroke_fill=(0, 0, 0))
    X, Y = to_px(lon, lat)
    d.ellipse((X - 9, Y - 9, X + 9, Y + 9), outline=(0, 255, 255), width=3)
    # 地点の周り約200m四方に切り出して拡大（地番の文字が重ならないように）
    img.crop((int(X - 200), int(Y - 200), int(X + 200), int(Y + 200))).resize((900, 900)).save(path)


def run(lat, lon, cache, png=None):
    os.makedirs(cache, exist_ok=True)
    log = []
    muni_cd, oaza = reverse_geocode(lat, lon)
    log.append(f"逆ジオコーダ: 市区町村コード{muni_cd}・大字「{oaza}」")
    files = []
    for key, year, url, title in find_dataset_zips(muni_cd):
        zpath = download_zip(url, os.path.join(cache, f"{key}-{year}.zip"))
        log.append(f"登記所備付地図 {title} {year}年版")
        for xp in xmls_for_oaza(zpath, oaza, os.path.join(cache, f"{key}-{year}")):
            zahyo, meta, parcels = parse_moj_xml(xp, oaza)
            if parcels:
                files.append((xp, zahyo, meta, parcels))
    if not files:
        log.append("この大字の地図データが見つかりませんでした")
    cands = locate(lat, lon, files, log) if files else []
    inside = [c for c in cands if c["距離m"] == 0 and is_land(c["地番"])]
    near = [c for c in cands if is_land(c["地番"])][:6]
    if inside:
        status, main = "地番を特定", inside
    elif near:
        status, main = "筆の境目付近（候補あり）", near
    else:
        status, main = "特定できず", []
    labels = []
    for c in main:
        lb = chiban_label(c["地番"])
        if lb and lb not in labels:
            labels.append(lb)
    method = main[0]["方法"] if main else ""
    conf = main[0]["確度"] if main else ""
    search_text = f"{oaza}{labels[0].replace('番', '-').rstrip('-')}" if labels else oaza
    memo = (f"eMAFF農地ナビ（{EMAFF_URL}）の「住所・地番で探す」に『{search_text}』と入力（都道府県・市区町村を選んでから）。"
            "表示された筆で、地目（田・畑）と農振区分（農用地区域=青地／それ以外=白地）を確認する。")
    if status != "地番を特定":
        memo += f" 地点は筆の境目付近のため、候補の{('・'.join(labels[:4]) or '地番')}を順に確認する。"
    if conf == "中":
        memo += " 地番は公図の形を農地の区画（筆ポリゴン）に重ねて推定したもの。登記簿・公図で必ず確認する。"
    unfit = [l.split(":")[0] for l in log if "位置を推定できず" in l]
    if unfit:
        memo += f" 位置合わせできなかった地図（{'・'.join(unfit)}。山林・宅地など農地以外の筆）にまたがる可能性もある。"
    result = {
        "lat": lat, "lon": lon, "muni_cd": muni_cd, "oaza": oaza,
        "status": status, "chiban": labels, "method": method, "confidence": conf,
        "candidates": [{k: v for k, v in c.items() if not k.startswith("_")} for c in cands[:10]],
        "emaff": {"url": EMAFF_URL, "search_text": search_text, "memo": memo},
        "log": log,
    }
    if png and cands:
        draw_overlay(lat, lon, [c for c in cands if c["距離m"] < 15], png)
        result["overlay_png"] = png
    return result


def main():
    ap = argparse.ArgumentParser(description="緯度経度から地番を推定し、eMAFF農地ナビの検索語を作る")
    ap.add_argument("--lat", type=float, required=True)
    ap.add_argument("--lon", type=float, required=True)
    ap.add_argument("--cache", default="moj_cache", help="登記所備付地図zipの保存先（再利用される）")
    ap.add_argument("--out", required=True)
    ap.add_argument("--png", default="", help="候補の筆を航空写真に重ねた確認画像の保存先（任意）")
    a = ap.parse_args()
    res = run(a.lat, a.lon, a.cache, a.png or None)
    json.dump(res, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"{res['oaza']}: {res['status']} → {'・'.join(res['chiban']) or '―'}（{res['method']}）")
    for line in res["log"]:
        print("  " + line)
    print("  eMAFF検索語: " + res["emaff"]["search_text"])


if __name__ == "__main__":
    main()
