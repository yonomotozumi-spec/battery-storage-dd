#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中部電力パワーグリッド「系統空容量マッピング」のGeoJSON（chuden_cache/KRSIH001.geojson.gz、
fetch_chuden_coords.py extract で保存したもの）から、substation.html の地図モードで
送電線・変電所を重ねて表示するための data/chuden_grid.js を作る。

中部は唯一、公表マッピングが地理座標（WGS84）を持つため、送電線の実形状を地図に出せる。
他エリアの公表マップは模式図で座標が無い（README「他エリアの座標が自動取得できない理由」参照）。

  python scripts/build_chuden_grid_web.py --geo chuden_cache/KRSIH001.geojson.gz --out data/chuden_grid.js

出力は <script src> で読める素のJS（window.CHUDEN_GRID）。地図モードで「中部電力PGの送電線」を
オンにしたときだけ読み込む（約1MB）。
"""
import argparse
import gzip
import json
import os
import re
import unicodedata


def kv_of(s):
    m = re.search(r"(\d+)", unicodedata.normalize("NFKC", s or ""))
    return int(m.group(1)) if m else None


def mw_of(s):
    m = re.search(r"(-?\d+(?:\.\d+)?)", unicodedata.normalize("NFKC", s or ""))
    return float(m.group(1)) if m else None


def main():
    ap = argparse.ArgumentParser(description="中部電力PG 系統空容量マッピングの地図用データ生成")
    ap.add_argument("--geo", default="chuden_cache/KRSIH001.geojson.gz")
    ap.add_argument("--out", default="data/chuden_grid.js")
    args = ap.parse_args()

    with gzip.open(args.geo, "rt", encoding="utf-8") as f:
        geo = json.load(f)
    lines, points = [], []
    for ft in geo["features"]:
        p, g = ft["properties"], ft["geometry"]
        name = p.get("設備名称", "")
        kv = kv_of(p.get("電圧階級", ""))
        n1 = mw_of(p.get("N-1電制適用可能量", ""))
        nonfirm = 1 if str(p.get("ノンファーム型接続", "")).startswith("適用") else 0
        limit = p.get("N-1故障時の電源制限・抑制対象送変電設備", "") or ""
        limit = "" if limit.strip() in ("-", "－") else limit
        if g["type"] == "LineString":
            coords = [[round(c[1], 4), round(c[0], 4)] for c in g["coordinates"]]
            lines.append([name, kv, p.get("送電線番号", ""), n1, nonfirm, limit, coords])
        elif g["type"] == "Point":
            c = g["coordinates"]
            points.append([name, kv, p.get("変電所番号", ""), mw_of(p.get("空容量（当該設備）", "")), n1, nonfirm,
                           round(c[1], 5), round(c[0], 5)])
    payload = {
        "meta": {"source": "中部電力パワーグリッド 系統空容量マッピング（gridmap.powergrid.chuden.co.jp）",
                 "asof": geo.get("_taishobi", ""), "lines": len(lines), "points": len(points)},
        "lineFields": ["name", "kv", "no", "n1amt", "nonfirm", "limitEq", "coords"],
        "pointFields": ["name", "kv", "no", "availCur", "n1amt", "nonfirm", "lat", "lon"],
        "lines": lines, "points": points,
    }
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write("// 自動生成ファイル - 直接編集しないこと\n// 生成: scripts/build_chuden_grid_web.py\n"
                "window.CHUDEN_GRID=" + body + ";\n")
    print(f"出力: {args.out} 送電線{len(lines)}本・地点{len(points)}件 {os.path.getsize(args.out) / 1024:.0f}KB")


if __name__ == "__main__":
    main()
