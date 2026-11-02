#!/usr/bin/env python3
"""公開リポジトリに接続検討回答の実績が戻っていないか確認する。入っていれば公開を止める。"""
import csv, glob, json, os, sys
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
bad = []
for n in ["data/connection_results.json", "data/接続検討回答_検証一覧.xlsx", "geo_cache/connection_sites.json"]:
    if os.path.exists(os.path.join(ROOT, n)):
        bad.append(f"{n} が存在します")
src = open(os.path.join(ROOT, "data", "substations.js"), encoding="utf-8").read()
d = json.loads(src[src.index("window.SUBSTATION_DATA=") + 23:].rstrip().rstrip(";"))
if d["results"]:
    bad.append(f"data/substations.js の実績が {len(d['results'])} 件あります")
if d["analysisNotes"]:
    bad.append("data/substations.js に分析メモがあります")
csv.field_size_limit(10**8)
for p in glob.glob(os.path.join(ROOT, "exports", "*.csv")):
    with open(p, encoding="utf-8-sig") as f:
        rows = csv.reader(f)
        h = next(rows)
        idx = [i for i, c in enumerate(h) if c.startswith("接続検討の実績") or c.startswith("実績_")]
        if any(r[i].strip() not in ("", "0", "－", "-", "なし") for r in rows for i in idx):
            bad.append(os.path.basename(p) + " に実績が入っています")
if bad:
    sys.exit("公開版に実績が含まれています:\n- " + "\n- ".join(bad))
print("OK: 公開版に実績は含まれていません")
