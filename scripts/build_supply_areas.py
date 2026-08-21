#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一般送配電事業者の供給区域ポリゴン（data/supply_areas.js）を生成するスクリプト。

substation.html は「候補地の緯度経度」から最寄りの変電所を採点するが、
変電所リストのエリア（北海道〜沖縄）は *変電所側* の所属でしかない。
候補地がどの一般送配電事業者の供給区域にあるかを判定しないと、
供給区域をまたいだ変電所（実際には連系できない）を最有力候補として出してしまう。

例: 愛媛県今治市吉海町（大島）は県は愛媛だが供給区域は**中国電力ネットワーク**。
    9.5km先の波止浜変電所（四国電力送配電）は系統ランクSだが、管轄が違うため連系できない。

そこで市区町村界（国土数値情報 N03 を smartnews-smri/japan-topography が
GeoJSON 化したもの）を各社の供給区域に割り当て、エリアごとに融合して
ブラウザで点内判定できる軽量ポリゴンとして書き出す。

供給区域の定義は各社の託送供給等約款・Wikipedia記載の区域による（README参照）。
市区町村より細かい例外（富士市の旧富士川町など）は分割できないため
「境界市町村」として別に持ち、画面では断定せず注意喚起する。

使い方:
  # ① 取得（要ネットワーク）
  python scripts/build_supply_areas.py fetch --cache-dir muni_cache
  # ② 生成（ネットワーク不要）
  python scripts/build_supply_areas.py build --cache-dir muni_cache --out data/supply_areas.js
"""
import argparse
import datetime
import json
import os
import sys
import urllib.request

from shapely.geometry import shape, Point
from shapely.ops import unary_union
from shapely.strtree import STRtree

SRC_URL = ("https://raw.githubusercontent.com/smartnews-smri/japan-topography/"
           "main/data/municipality/geojson/s0010/N03-21_210101.json")
SRC_NAME = "N03-21_210101.json"
SRC_LABEL = "国土数値情報 行政区域(N03-21) / smartnews-smri/japan-topography s0010"

AREAS = ["北海道", "東北", "東京", "中部", "北陸", "関西", "中国", "四国", "九州", "沖縄"]

# ── 都道府県 → 供給区域（既定値）──
PREF_AREA = {}
for _p in ["北海道"]:
    PREF_AREA[_p] = "北海道"
for _p in ["青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県", "新潟県"]:
    PREF_AREA[_p] = "東北"
for _p in ["茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県", "山梨県"]:
    PREF_AREA[_p] = "東京"
for _p in ["長野県", "岐阜県", "静岡県", "愛知県", "三重県"]:
    PREF_AREA[_p] = "中部"
for _p in ["富山県", "石川県", "福井県"]:
    PREF_AREA[_p] = "北陸"
for _p in ["滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県"]:
    PREF_AREA[_p] = "関西"
for _p in ["鳥取県", "島根県", "岡山県", "広島県", "山口県"]:
    PREF_AREA[_p] = "中国"
for _p in ["徳島県", "香川県", "愛媛県", "高知県"]:
    PREF_AREA[_p] = "四国"
for _p in ["福岡県", "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県"]:
    PREF_AREA[_p] = "九州"
PREF_AREA["沖縄県"] = "沖縄"

# ── 市区町村まるごと他社管内になる例外（都道府県, 市区町村名）→ 供給区域 ──
# 静岡県は富士川で東西に分かれるため、東京電力PG側の市町を列挙する。
MUNI_AREA = {
    # 静岡県 富士川以東 = 東京電力パワーグリッド
    ("静岡県", "熱海市"): "東京", ("静岡県", "沼津市"): "東京", ("静岡県", "三島市"): "東京",
    ("静岡県", "伊東市"): "東京", ("静岡県", "御殿場市"): "東京", ("静岡県", "下田市"): "東京",
    ("静岡県", "裾野市"): "東京", ("静岡県", "伊豆市"): "東京", ("静岡県", "伊豆の国市"): "東京",
    ("静岡県", "富士宮市"): "東京", ("静岡県", "富士市"): "東京",
    ("静岡県", "函南町"): "東京", ("静岡県", "清水町"): "東京", ("静岡県", "長泉町"): "東京",
    ("静岡県", "小山町"): "東京", ("静岡県", "東伊豆町"): "東京", ("静岡県", "河津町"): "東京",
    ("静岡県", "南伊豆町"): "東京", ("静岡県", "松崎町"): "東京", ("静岡県", "西伊豆町"): "東京",
    # 三重県 熊野市以南 = 関西電力送配電
    ("三重県", "熊野市"): "関西", ("三重県", "御浜町"): "関西", ("三重県", "紀宝町"): "関西",
    # 福井県 若狭（嶺南）= 関西電力送配電。敦賀市は北陸電力送配電のまま。
    ("福井県", "小浜市"): "関西", ("福井県", "美浜町"): "関西", ("福井県", "高浜町"): "関西",
    ("福井県", "おおい町"): "関西", ("福井県", "若狭町"): "関西",
    # 香川県 小豆島・直島諸島 = 中国電力ネットワーク
    ("香川県", "土庄町"): "中国", ("香川県", "小豆島町"): "中国", ("香川県", "直島町"): "中国",
    # 愛媛県 芸予諸島 = 中国電力ネットワーク
    ("愛媛県", "上島町"): "中国",
}

# ── 市区町村より細かい例外（分割できないため境界として注意喚起する）──
# (都道府県, 市区町村名): (既定の供給区域, 注記)
BOUNDARY_MUNI = {
    ("静岡県", "富士市"): ("東京", "旧庵原郡富士川町の区域のみ中部電力パワーグリッド"),
    ("静岡県", "富士宮市"): ("東京", "旧庵原郡内房村の区域のみ中部電力パワーグリッド"),
    ("三重県", "熊野市"): ("関西", "東部の海沿い8地区のみ中部電力パワーグリッド"),
    ("岐阜県", "関ケ原町"): ("中部", "旧今須村の区域のみ関西電力送配電"),
    ("岐阜県", "飛騨市"): ("中部", "旧吉城郡神岡町・坂下村の区域のみ北陸電力送配電"),
    ("岐阜県", "郡上市"): ("中部", "旧郡上郡白鳥町石徹白の区域のみ北陸電力送配電"),
    ("兵庫県", "赤穂市"): ("関西", "旧岡山県和気郡日生町から編入した福浦地区のみ中国電力ネットワーク"),
    ("愛媛県", "新居浜市"): ("四国", "別子山地区は住友共同電力（特定送配電事業者）の供給区域"),
}

# ── 今治市は島ごとに供給区域が違う（旧越智郡の島嶼部＝中国電力ネットワーク）──
# 四国電力送配電の供給区域から除かれるのは旧町村単位（伯方町・上浦町・大三島町・
# 宮窪町・吉海町・関前村）なので、市区町村界では割れない。ただし対象はすべて島で
# あり N03 では別ポリゴンになるため、島ごとに振り分ける。
# 各点は地理院リバースジオコーダで大字を確認済み（コメントの地名がその結果）。
# 列挙外のポリゴン（本土側と旧今治市の島）は四国のまま。
IMABARI_CHUGOKU_POINTS = [
    (34.24497, 133.02424),   # 大三島町宮浦（大三島）
    (34.15219, 133.05936),   # 吉海町仁江（大島）
    (34.21644, 133.09323),   # 伯方町北浦（伯方島）
    (34.18819, 132.88014),   # 関前岡村（岡村島）
    (34.19239, 132.92510),   # 関前大下（大下島）
    (34.19082, 132.90283),   # 関前小大下（小大下島）
    (34.15138, 133.00083),   # 吉海町津島（津島）
    (34.18959, 133.08973),   # 宮窪町宮窪（鵜島）
    (34.10739, 133.18319),   # 宮窪町四阪島（四阪島）
    (34.11788, 133.18403),   # 宮窪町四阪島（家ノ島）
    (34.24583, 132.95236),   # 大三島町口総（大三島 西部の別ポリゴン）
    (34.20295, 132.93584),   # 大三島町宗方（大三島 南西部の別ポリゴン）
]

SCALE = 10000  # 座標を度×10000（≒11m）の整数に量子化する
SIMPLIFY_TOL = 0.0008  # 約80m。供給区域の判定は境界付近では断定しないため十分

# リングは Google の encoded polyline と同じ可変長6bitで文字列化する。
# JSON文字列にそのまま入れられるよう、英数字と - _ の64文字だけを使う
# （バックスラッシュや引用符を含む標準の 63..126 の並びはエスケープでかえって太る）。
ALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"


def enc_val(v, out):
    v = (v << 1) ^ (v >> 63) if v < 0 else (v << 1)  # zigzag
    while True:
        c = v & 0x1F
        v >>= 5
        out.append(ALPHA[c | 0x20] if v else ALPHA[c])
        if not v:
            return


def muni_name(props):
    """N03の市区町村名。

    N03_003 は政令市では市名、郡部では郡名が入る。政令市の区は同じ市に属するので
    市名にまとめ、郡部は N03_004（町村名）を使う。
    """
    a, b = props.get("N03_003"), props.get("N03_004")
    if a and a.endswith("市"):
        return a
    return b or a or ""


def area_of(props):
    pref = props.get("N03_001") or ""
    name = muni_name(props)
    if (pref, name) in MUNI_AREA:
        return MUNI_AREA[(pref, name)]
    if (pref, name) in BOUNDARY_MUNI:
        return BOUNDARY_MUNI[(pref, name)][0]
    return PREF_AREA.get(pref)


def cmd_fetch(args):
    os.makedirs(args.cache_dir, exist_ok=True)
    dst = os.path.join(args.cache_dir, SRC_NAME)
    print(f"取得: {SRC_URL}")
    with urllib.request.urlopen(SRC_URL, timeout=300) as r, open(dst, "wb") as f:
        f.write(r.read())
    print(f"保存: {dst}  {os.path.getsize(dst)/1024/1024:.1f}MB")


def load_features(cache_dir):
    src = os.path.join(cache_dir, SRC_NAME)
    if not os.path.exists(src):
        sys.exit(f"キャッシュがありません: {src}\n先に `fetch` を実行してください。")
    with open(src, encoding="utf-8") as f:
        return json.load(f)["features"]


def polygons(geom):
    return list(geom.geoms) if geom.geom_type == "MultiPolygon" else [geom]


def enc_ring(coords):
    """[[lon,lat],...] → 量子化＋差分をpolyline風に詰めた文字列。

    先頭は絶対値、以降は前の点との差分。閉じたリングの最終点は先頭と同じなので
    落とす（JS側で閉じる）。
    """
    out, px, py, n = [], 0, 0, 0
    for i, (lo, la) in enumerate(coords):
        x, y = round(lo * SCALE), round(la * SCALE)
        if i and x == px and y == py:
            continue
        enc_val(x - px if i else x, out)
        enc_val(y - py if i else y, out)
        px, py = x, y
        n += 1
    return "".join(out), n


def enc_poly(poly):
    ring, n = enc_ring(list(poly.exterior.coords)[:-1])
    if n < 3:
        return None
    rings = [ring]
    for r in poly.interiors:
        hole, hn = enc_ring(list(r.coords)[:-1])
        if hn >= 3:
            rings.append(hole)
    b = poly.bounds  # (minx, miny, maxx, maxy)
    bb = [round(b[0] * SCALE), round(b[1] * SCALE), round(b[2] * SCALE), round(b[3] * SCALE)]
    return {"b": bb, "r": rings}


def cmd_build(args):
    feats = load_features(args.cache_dir)

    by_area = {a: [] for a in AREAS}
    boundary = []   # [{"p":都道府県,"m":市区町村,"a":既定エリア,"n":注記,"g":[poly...]}]
    unknown = []
    imabari_hit = 0

    imabari_pts = [Point(lo, la) for la, lo in IMABARI_CHUGOKU_POINTS]

    for f in feats:
        props = f["properties"]
        pref, name = props.get("N03_001") or "", muni_name(props)
        ar = area_of(props)
        if ar is None:
            unknown.append((pref, name))
            continue
        try:
            geom = shape(f["geometry"])
        except Exception:
            continue
        if geom.is_empty:
            continue
        if not geom.is_valid:
            geom = geom.buffer(0)

        parts = polygons(geom)
        if (pref, name) == ("愛媛県", "今治市"):
            # 島ごとに振り分ける
            for p in parts:
                if any(p.covers(pt) for pt in imabari_pts):
                    by_area["中国"].append(p)
                    imabari_hit += 1
                else:
                    by_area["四国"].append(p)
            missed = [i for i, pt in enumerate(imabari_pts)
                      if not any(p.covers(pt) for p in parts)]
            if missed:
                print("警告: 今治市の島の代表点がどのポリゴンにも入りませんでした "
                      f"（IMABARI_CHUGOKU_POINTS の {missed}）。元データの更新で"
                      "島の形が変わった可能性があります。")
        else:
            by_area[ar].extend(parts)

        if (pref, name) in BOUNDARY_MUNI:
            boundary.append({"pref": pref, "muni": name, "area": ar,
                             "note": BOUNDARY_MUNI[(pref, name)][1], "geom": geom})

    if unknown:
        print("警告: 供給区域を割り当てられなかった市区町村:", unknown[:10], f"（計{len(unknown)}件）")
    print(f"今治市: 中国エリアに振り替えた島ポリゴン {imabari_hit}件")

    polys = []
    for ai, a in enumerate(AREAS):
        if not by_area[a]:
            continue
        merged = unary_union(by_area[a])
        merged = merged.simplify(SIMPLIFY_TOL, preserve_topology=True)
        n = 0
        for p in polygons(merged):
            if p.is_empty:
                continue
            e = enc_poly(p)
            if e is None:
                continue
            e["a"] = ai
            polys.append(e)
            n += 1
        print(f"  {a}: {n}ポリゴン")

    bnd = []
    for b in boundary:
        merged = b["geom"].simplify(SIMPLIFY_TOL, preserve_topology=True)
        gs = []
        for p in polygons(merged):
            if p.is_empty:
                continue
            e = enc_poly(p)
            if e is not None:
                gs.append(e)
        if gs:
            bnd.append({"p": b["pref"], "m": b["muni"], "a": AREAS.index(b["area"]),
                        "n": b["note"], "g": gs})

    payload = {
        "meta": {
            "generated": datetime.date.today().isoformat(),
            "source": SRC_LABEL,
            "scale": SCALE,
            "simplifyTol": SIMPLIFY_TOL,
            "alpha": ALPHA,
        },
        "areas": AREAS,
        "polys": polys,
        "boundary": bnd,
    }
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    header = ("// 自動生成ファイル - 直接編集しないこと\n"
              "// 生成: scripts/build_supply_areas.py\n"
              f"// 元データ: {SRC_LABEL}\n")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(header + "window.SUPPLY_AREA_DATA=" + body + ";\n")
    print(f"出力: {args.out}  {os.path.getsize(args.out)/1024:.0f}KB "
          f"（区域{len(polys)}ポリゴン／境界市町村{len(bnd)}件）")


def cmd_check(args):
    """変電所リストのエリアと、その座標の供給区域が食い違う行を洗い出す。

    リストの「エリア」は各社の公表資料の出所であって供給区域とは限らない
    （連系線・他社エリア内の自社設備などは食い違って当然）。座標の取り違えや
    リスト側の誤りを見つけるための点検用で、データは書き換えない。
    """
    feats = load_features(args.cache_dir)
    geoms, areas = [], []
    imabari_pts = [Point(lo, la) for la, lo in IMABARI_CHUGOKU_POINTS]
    for f in feats:
        props = f["properties"]
        pref, name = props.get("N03_001") or "", muni_name(props)
        ar = area_of(props)
        if ar is None:
            continue
        geom = shape(f["geometry"])
        if not geom.is_valid:
            geom = geom.buffer(0)
        for p in polygons(geom):
            a = ar
            if (pref, name) == ("愛媛県", "今治市"):
                a = "中国" if any(p.covers(pt) for pt in imabari_pts) else "四国"
            geoms.append(p)
            areas.append(a)
    tree = STRtree(geoms)

    with open(args.substations, encoding="utf-8") as f:
        js = f.read()
    head = "window.SUBSTATION_DATA="
    data = json.loads(js[js.index(head) + len(head):].strip().rstrip(";"))
    F = {n: i for i, n in enumerate(data["fields"])}
    d_area, d_rank, d_acc = data["dict"]["area"], data["dict"]["rank"], data["dict"]["acc"]

    rows, n_coord = [], 0
    for r in data["rows"]:
        la, lo = r[F["lat"]], r[F["lon"]]
        if la is None or lo is None:
            continue
        n_coord += 1
        listed = d_area[r[F["area"]]]
        if listed in ("J-POWER",):
            continue
        pt = Point(lo, la)
        got = None
        for i in tree.query(pt):
            if geoms[i].covers(pt):
                got = areas[i]
                break
        if got and got != listed:
            rows.append([r[F["no"]], listed, got, r[F["name"]], r[F["pref"]],
                         d_acc[r[F["acc"]]], d_rank[r[F["rank"]]], la, lo])

    print(f"座標あり {n_coord}件 / エリアと供給区域が不一致 {len(rows)}件")
    for x in sorted(rows, key=lambda x: (x[1], x[2], x[3])):
        print(f"  no={x[0]:<5} リスト={x[1]:4} 座標の供給区域={x[2]:4} {x[3]}"
              f"（{x[4] or '都道府県なし'}／{x[5]}／{x[6]}） {x[7]},{x[8]}")
    if args.report:
        import csv
        with open(args.report, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["No", "リストのエリア", "座標の供給区域", "変電所名",
                        "都道府県", "座標精度", "ランク", "緯度", "経度"])
            w.writerows(sorted(rows, key=lambda x: (x[1], x[2], x[3])))
        print(f"出力: {args.report}")


def main():
    ap = argparse.ArgumentParser(description="一般送配電事業者の供給区域ポリゴンを生成")
    sub = ap.add_subparsers(dest="cmd", required=True)
    f = sub.add_parser("fetch", help="市区町村界GeoJSONを取得してキャッシュ")
    f.add_argument("--cache-dir", default="muni_cache")
    f.set_defaults(func=cmd_fetch)
    b = sub.add_parser("build", help="キャッシュから供給区域JSを生成")
    b.add_argument("--cache-dir", default="muni_cache")
    b.add_argument("--out", default="data/supply_areas.js")
    b.set_defaults(func=cmd_build)
    c = sub.add_parser("check", help="変電所リストのエリアと座標の供給区域の食い違いを点検")
    c.add_argument("--cache-dir", default="muni_cache")
    c.add_argument("--substations", default="data/substations.js")
    c.add_argument("--report", help="CSVに出力する場合のパス")
    c.set_defaults(func=cmd_check)
    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
