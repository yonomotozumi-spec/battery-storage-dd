#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
変電所座標の精度レポート。定期更新（substation-coords-refresh ワークフロー）のたびに
「どれだけ座標が揃い、どれだけ確かか」を測り、推移を残す。

座標精度の値を4段階にまとめて数える。

  高 … OSM実測 / 公表資料(手動)        位置そのものを測った・公表された座標
  中 … Google Places(名称一致) / OSM実測(名称一致)
                                       名称一致で取った座標。同名の別施設・1km前後のずれの余地あり
  低 … 推定                            市町村レベルの推定。S7（500m〜5km刻み）には粗すぎる
  無 … 未取得・空欄

あわせて「中・低」の行をOSMキャッシュと照合し、OSM側の座標と1km以上離れているものを
**要確認**として一覧にする（データは書き換えない。どちらが正しいかは地図で確認する）。

使い方:
  python scripts/coord_quality_report.py \
    --in "data/全国変電所リスト_66kV以上_座標追加.xlsx" --cache-dir osm_cache \
    --out "docs/変電所座標_精度レポート.md" --history data/coord_quality_history.csv
"""
import argparse
import csv
import datetime
import os
import urllib.parse

from fetch_substation_coords import (area_ranges, haversine_km, load_cache, match_row,
                                     read_rows)
from score_substations import read_substations
import substation_scoring as S

TIERS = ["高", "中", "低", "無"]
TIER_OF = {
    "OSM実測": "高", "公表資料(手動)": "高",
    "Google Places(名称一致)": "中", "OSM実測(名称一致)": "中",
    "推定": "低",
}
CHECK_KM = 1.0   # 要確認にするOSMとの食い違い（S7の刻み 500m/1km を跨ぎうる距離）
HISTORY_FIELDS = ["日付", "全件", "高", "中", "低", "無", "要確認"]


def tier(r):
    if not isinstance(r["lat"], (int, float)) or not isinstance(r["lon"], (int, float)):
        return "無"
    return TIER_OF.get(r["acc"] or "", "中")   # 未知の値は安全側に「中」


def gmap(lat, lon):
    return "https://www.google.com/maps?q=" + urllib.parse.quote(f"{lat},{lon}")


def ranks_by_no(path, sheet):
    out = {}
    for (no, area, name, pref, vmax, vsec, lat, lon, acc, cap, opcap, flow,
         avail, avail_up, n1, curtail, addr) in read_substations(path, sheet):
        out[no] = S.score_row(area, vsec, opcap, flow, avail, avail_up, n1, curtail)["rank"]
    return out


def pct(n, d):
    return f"{n / d * 100:.1f}%" if d else "-"


def update_history(path, today, counts):
    rows = []
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as f:
            rows = [r for r in csv.DictReader(f) if r["日付"] != today]
    rows.append({"日付": today, **{k: str(v) for k, v in counts.items()}})
    rows.sort(key=lambda r: r["日付"])
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=HISTORY_FIELDS)
        w.writeheader()
        w.writerows(rows)
    return rows


def main():
    ap = argparse.ArgumentParser(description="変電所座標の精度レポートを出力し、推移を記録する")
    ap.add_argument("--in", dest="src", required=True)
    ap.add_argument("--cache-dir", required=True)
    ap.add_argument("--out", required=True, help="レポートMarkdownの出力先")
    ap.add_argument("--history", help="推移CSV（同じ日付の行は上書き）")
    ap.add_argument("--list-sheet", default="全国変電所リスト")
    ap.add_argument("--max-check", type=int, default=50, help="要確認の一覧に載せる上限")
    args = ap.parse_args()

    rows, _, _ = read_rows(args.src, args.list_sheet)
    by_pref, japan = load_cache(args.cache_dir)
    box = area_ranges(rows)
    rank = ranks_by_no(args.src, args.list_sheet)

    areas = []
    by_area = {}
    total = {t: 0 for t in TIERS}
    missing_sa = {}
    checks = []
    for r in rows:
        t = tier(r)
        if r["area"] not in by_area:
            areas.append(r["area"])
            by_area[r["area"]] = {k: 0 for k in TIERS}
        by_area[r["area"]][t] += 1
        total[t] += 1
        if t in ("無", "低") and rank.get(r["no"]) in ("S", "A"):
            missing_sa[r["area"]] = missing_sa.get(r["area"], 0) + 1
        if t in ("中", "低") and r["acc"] != "OSM実測(名称一致)":
            best, how, scope, _ = match_row(r, by_pref, japan, box)
            if best is not None:
                d = haversine_km(r["lat"], r["lon"], best["lat"], best["lon"])
                if d >= CHECK_KM:
                    checks.append((d, r, best, scope))
    checks.sort(key=lambda c: -c[0])

    n = len(rows)
    counts = {"全件": n, **total, "要確認": len(checks)}
    today = datetime.date.today().isoformat()
    history = update_history(args.history, today, counts) if args.history else []

    L = []
    L.append("# 変電所座標 精度レポート")
    L.append("")
    L.append(f"生成: {today}（`scripts/coord_quality_report.py`。月1回の定期更新で自動生成）")
    L.append("")
    L.append("| 区分 | 意味 | 件数 | 割合 |")
    L.append("|---|---|---|---|")
    desc = {"高": "OSM実測・公表資料（位置を測った座標）",
            "中": "名称一致（Google Places・OSM）。同名の別施設・1km前後のずれの余地",
            "低": "市町村レベルの推定。S7の判定には粗い",
            "無": "座標なし（距離検索に出ない）"}
    for t in TIERS:
        L.append(f"| {t} | {desc[t]} | {total[t]:,} | {pct(total[t], n)} |")
    L.append(f"| 計 | | {n:,} | |")
    L.append("")
    L.append("## エリア別")
    L.append("")
    L.append("| エリア | 全件 | 高 | 中 | 低 | 無 | 座標あり | S/Aで低・無 |")
    L.append("|---|---|---|---|---|---|---|---|")
    for a in areas:
        c = by_area[a]
        m = sum(c.values())
        L.append(f"| {a} | {m:,} | {c['高']:,} | {c['中']:,} | {c['低']:,} | {c['無']:,} | "
                 f"{pct(m - c['無'], m)} | {missing_sa.get(a, 0):,} |")
    L.append("")
    L.append("「S/Aで低・無」は系統ランクS/Aなのに座標が推定か未取得の件数で、精度を上げる効果が最も大きい行です。"
             "一覧は `scripts/list_missing_coords.py` で出せます。")
    L.append("")
    if len(history) > 1:
        L.append("## 推移")
        L.append("")
        L.append("| 日付 | 高 | 中 | 低 | 無 | 要確認 |")
        L.append("|---|---|---|---|---|---|")
        for h in history[-12:]:
            L.append(f"| {h['日付']} | {h['高']} | {h['中']} | {h['低']} | {h['無']} | {h['要確認']} |")
        L.append("")
    L.append(f"## 要確認：OSMの座標と{CHECK_KM:g}km以上食い違う行（{len(checks)}件）")
    L.append("")
    if not checks:
        L.append("ありません。")
    else:
        L.append("座標精度が「中・低」の行のうち、OSMに同名の変電所が一意に見つかり、位置が食い違うものです。"
                 "データは書き換えていません。両方の地図を開き、正しい方を `merge_coords.py --force` で反映してください。")
        L.append("")
        L.append("| No. | エリア | 都道府県 | 変電所名 | 現在の精度 | 食い違い | 現在の座標 | OSMの座標 |")
        L.append("|---|---|---|---|---|---|---|---|")
        for d, r, best, scope in checks[:args.max_check]:
            L.append(f"| {r['no']} | {r['area']} | {r['pref'] or '-'} | {r['name']} | {r['acc']} | {d:.1f}km | "
                     f"[地図]({gmap(r['lat'], r['lon'])}) | [地図]({gmap(best['lat'], best['lon'])}) |")
        if len(checks) > args.max_check:
            L.append("")
            L.append(f"ほか{len(checks) - args.max_check}件")
    L.append("")

    with open(args.out, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    print(f"高{total['高']} 中{total['中']} 低{total['低']} 無{total['無']} / 要確認{len(checks)}件")
    print(f"出力: {args.out}")


if __name__ == "__main__":
    main()
