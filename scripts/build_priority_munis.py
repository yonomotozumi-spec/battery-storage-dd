#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""S/Aランク変電所の所在市区町村リスト（条例リスク優先監視対象）を生成する。

data/substations.js のランクS/A変電所（座標あり）を国土地理院の逆ジオコーダで
市区町村に変換し、data/priority_munis.json に集計を書き出す。
条例リスク監視（層③）で「実際に土地を探すエリアの自治体」を優先的に見るために使う。

  python scripts/build_priority_munis.py            # 生成（逆ジオコードはgeo_cacheに蓄積）
  python scripts/build_priority_munis.py --limit 50 # 動作確認用に件数を絞る

逆ジオコーダ: https://mreversegeocoder.gsi.go.jp/reverse-geocoder/LonLatToAddress
市区町村コード表: https://maps.gsi.go.jp/js/muni.js
"""

import argparse
import json
import re
import ssl
import time
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SUBS_JS = ROOT / "data" / "substations.js"
OUT_PATH = ROOT / "data" / "priority_munis.json"
CACHE_DIR = ROOT / "geo_cache"
REVGEO_CACHE = CACHE_DIR / "revgeo.json"
MUNI_CACHE = CACHE_DIR / "muni_codes.json"

REVGEO_URL = "https://mreversegeocoder.gsi.go.jp/reverse-geocoder/LonLatToAddress?lat={lat}&lon={lon}"
MUNI_JS_URL = "https://maps.gsi.go.jp/js/muni.js"
UA = "battery-storage-dd/1.0 (ordinance risk screening)"


def http_get(url, timeout=20, retries=3):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    ctx = ssl.create_default_context()
    last = None
    for i in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
                return r.read()
        except Exception as exc:
            last = exc
            time.sleep(2 ** i)
    raise last


def load_muni_table():
    """muniCd → (都道府県, 市区町村) の表。政令市の区は市名に集約する。"""
    if MUNI_CACHE.exists():
        return {k: tuple(v) for k, v in json.loads(MUNI_CACHE.read_text(encoding="utf-8")).items()}
    raw = http_get(MUNI_JS_URL).decode("utf-8", errors="replace")
    table = {}
    for m in re.finditer(r"MUNI_ARRAY\[\"(\d+)\"\]\s*=\s*'([^']+)'", raw):
        code, val = m.group(1), m.group(2)
        parts = val.split(",")
        if len(parts) >= 4:
            pref, muni = parts[1], parts[3]
            muni = muni.replace("　", " ")
            if " " in muni:  # 「札幌市 中央区」→「札幌市」
                muni = muni.split(" ")[0]
            table[code.lstrip("0") or "0"] = (pref, muni)
    CACHE_DIR.mkdir(exist_ok=True)
    MUNI_CACHE.write_text(json.dumps({k: list(v) for k, v in table.items()}, ensure_ascii=False), encoding="utf-8")
    return table


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--limit", type=int, help="処理する変電所数の上限（動作確認用）")
    ap.add_argument("--sleep", type=float, default=0.15, help="逆ジオコーダへのリクエスト間隔秒")
    args = ap.parse_args()

    raw = SUBS_JS.read_text(encoding="utf-8")
    data = json.loads(raw[raw.index("{"): raw.rindex("}") + 1])
    f = {n: i for i, n in enumerate(data["fields"])}
    rank_dict = data["dict"]["rank"]

    targets = [
        r for r in data["rows"]
        if rank_dict[r[f["rank"]]] in ("S", "A") and r[f["lat"]] is not None and r[f["lon"]] is not None
    ]
    if args.limit:
        targets = targets[: args.limit]

    CACHE_DIR.mkdir(exist_ok=True)
    cache = json.loads(REVGEO_CACHE.read_text(encoding="utf-8")) if REVGEO_CACHE.exists() else {}
    muni_table = load_muni_table()

    fetched = failed = 0
    for i, r in enumerate(targets):
        key = f"{r[f['lat']]:.4f},{r[f['lon']]:.4f}"
        if key in cache:
            continue
        try:
            body = json.loads(http_get(REVGEO_URL.format(lat=r[f["lat"]], lon=r[f["lon"]])))
            cache[key] = body.get("results", {}).get("muniCd") or ""
            fetched += 1
        except Exception:
            cache[key] = None  # 海上・国外・エラー
            failed += 1
        if fetched % 100 == 0 and fetched:
            REVGEO_CACHE.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
            print(f"  {i+1}/{len(targets)} 逆ジオコード済み…")
        time.sleep(args.sleep)
    REVGEO_CACHE.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")

    munis = {}
    unresolved = 0
    for r in targets:
        key = f"{r[f['lat']]:.4f},{r[f['lon']]:.4f}"
        code = cache.get(key)
        code = code.lstrip("0") if code else None
        if not code or code not in muni_table:
            unresolved += 1
            continue
        pref, muni = muni_table[code]
        rank = rank_dict[r[f["rank"]]]
        rec = munis.setdefault((pref, muni), {"pref": pref, "muni": muni, "S": 0, "A": 0, "substations": []})
        rec[rank] += 1
        rec["substations"].append(f"{r[f['name']]}({rank})")

    out = {
        "description": "ランクS/A変電所（系統スコア上位＝実際に土地を探すエリア）の所在市区町村。条例リスク監視の優先対象。build_priority_munis.py で再生成。",
        "generated": time.strftime("%Y-%m-%d"),
        "source_ranks": ["S", "A"],
        "substations_total": len(targets),
        "unresolved": unresolved,
        "munis": sorted(munis.values(), key=lambda x: (-x["S"], -x["A"], x["pref"])),
    }
    OUT_PATH.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    n = len(munis)
    print(f"生成: {OUT_PATH.relative_to(ROOT)}  変電所{len(targets)}件 → {n}市区町村（未解決{unresolved}件、新規取得{fetched}件、失敗{failed}件）")


if __name__ == "__main__":
    main()
