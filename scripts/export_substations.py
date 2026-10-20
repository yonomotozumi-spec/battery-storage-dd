#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
変電所を地域・条件で絞り込んでCSVに払い出す（システム側の出力）。

data/substations.js（build_substation_web_data.py の生成物）を読むので、スコア・ランク・
市区町村は Web版（substation.html）と同じ値になる。CSVの列も Web版の「CSVで保存」と同じ。
上位系統の状態・要接続検討・実績の分類・上位側の送電線・上位系の制約設備（地図モードと同じ判定）は
data/grid_topology.js を grid_status.py で読んで付ける（無ければ空欄）。
出力は BOM付きUTF-8 なので Excel でそのまま開ける。

使い方:
  # 地域で払い出す。--region はエリア・都道府県・市区町村・変電所名の部分一致（カンマ区切りでOR）
  python scripts/export_substations.py --region 秋田県 --out 秋田県_SA.csv
  python scripts/export_substations.py --region "横手市,大仙市" --ranks S,A,B --out 横手周辺.csv
  python scripts/export_substations.py --region 中部 --up ok --min-cap 6 --out 中部_上位系確認済.csv

  # 都道府県別CSVを一括生成（exports/ に47都道府県＋全国S/A＋目次README）
  python scripts/export_substations.py --by-pref exports
"""
import argparse
import csv
import datetime
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from grid_status import GridStatus, TOPO_JS  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_JS = os.path.join(ROOT, "data", "substations.js")
GS = None  # 上位系統の判定（main で読み込む）
ALL_RANKS = ["S", "A", "B", "C", "データ無", "除外"]


def load(path):
    raw = open(path, encoding="utf-8").read()
    d = json.loads(raw[raw.index("{"): raw.rindex("}") + 1])
    f = {n: i for i, n in enumerate(d["fields"])}
    memo = {}
    for r in d["results"]:
        memo.setdefault(r[0] + str(r[1]).strip(), []).append(r[8])

    out = []
    for row in d["rows"]:
        g = lambda k: row[f[k]] if k in f else None  # noqa: E731
        rec = {k: g(k) for k in d["fields"]}
        for k in ("area", "acc", "n1", "curtail", "rank"):
            rec[k] = d["dict"][k][row[f[k]]]
        # 所在地の都道府県は座標から引いたものを優先する。元リストの都道府県は管轄の区分らしく、
        # 実際の所在地と違う行がある（例: 上野変電所が「埼玉県」）。座標が無い行は元リストの値
        rec["prefAny"] = rec.get("gpref") or rec.get("pref") or ""
        rec["memo"] = " / ".join(memo.get(rec["area"] + str(rec["name"]).strip(), []))
        rec["_row"] = row
        out.append(rec)
    return d["meta"], out


def _gs(fn, default=""):
    """上位系統の判定列。grid_status が読めないときは空欄にする。"""
    def f(r):
        if GS is None:
            return default
        return fn(GS, r["_row"])
    return f


def has_data(r):
    return r["availUp"] is not None or r["availCur"] is not None


def up_label(r):
    if not has_data(r):
        return "空容量データ無し"
    return "確認済み" if r["availUp"] is not None else "未確認(当該設備で代用)"


def eff_cap(r):
    return r["availUp"] if r["availUp"] is not None else r["availCur"]


def blank(v):
    return "" if v is None or v == "—" else v


# Web版（substation.html の CSV_COLS）と同じ並び
COLUMNS = [
    ("No.", lambda r: r["no"]), ("エリア", lambda r: r["area"]), ("変電所名", lambda r: r["name"]),
    ("都道府県", lambda r: r["prefAny"]), ("市区町村", lambda r: r.get("muni")),
    ("元リストの都道府県", lambda r: r.get("pref")),
    ("最大電圧kV", lambda r: r["vmax"]), ("二次電圧kV", lambda r: r["vsec"]),
    ("系統スコア(80)", lambda r: r["grid"]), ("系統ランク", lambda r: r["rank"]),
    # 地図と同じカテゴリ（5色）と、接続検討の実績の件数・分類
    ("カテゴリ（地図の色）", lambda r: GS.category_text(r["_row"]) if GS else ""),
    ("接続検討の実績（件数・分類）", lambda r: GS.memo_text(r["_row"]) if GS else ""),
    ("上位系の確認", up_label),
    ("空容量_上位系MW", lambda r: r["availUp"]), ("空容量_当該MW", lambda r: r["availCur"]),
    ("N-1電制", lambda r: r["n1"]), ("出力制御", lambda r: r["curtail"]),
    ("運用容量MW", lambda r: r["opcap"]), ("予想潮流MW", lambda r: r["flow"]),
    ("S1空容量", lambda r: r["s1"]), ("S2 N-1", lambda r: r["s2"]), ("S3潮流", lambda r: r["s3"]),
    ("S4出制", lambda r: r["s4"]), ("S5配変", lambda r: r["s5"]), ("S6エリア", lambda r: r["s6"]),
    ("緯度", lambda r: r["lat"]), ("経度", lambda r: r["lon"]), ("座標精度", lambda r: r["acc"]),
    ("所在地(参考)", lambda r: r["addr"]), ("接続検討の実績", lambda r: r["memo"]),
    ("接続検討の実績_詳細", lambda r: GS.memo_detail_text(r["_row"]) if GS else ""),
    ("Googleマップ", lambda r: f"https://www.google.com/maps?q={r['lat']},{r['lon']}" if r["lat"] is not None else ""),
    # ── 地図モードと同じ判定（grid_status.py）──
    ("変電所単体の空き", _gs(lambda g, row: {"ok": "空きあり", "bad": "空き無し", "unk": "非公表"}[g.grid_status(row)["own"]["st"]])),
    ("上位系の空き", _gs(lambda g, row: {"ok": "空きあり", "bad": "空き無し", "unk": "不明"}[g.grid_status(row)["up"]])),
    ("上位系統の判定理由", _gs(lambda g, row: g.upper_status(row)["reason"])),
    ("要接続検討", _gs(lambda g, row: "○" if g.grid_status(row)["cls"] == "upbad" else "")),
    ("要接続検討の理由", _gs(lambda g, row: ("変電所は空きあり・上位系は空き無し" + ("。" + "、".join(g.need_check(row)) if g.need_check(row) else "")) if g.grid_status(row)["cls"] == "upbad" else "")),
    ("実績_対策なし", _gs(lambda g, row: g.memo_summary(row)["none"] or "")),
    ("実績_バンク等の対策待ち", _gs(lambda g, row: g.memo_summary(row)["bank"] or "")),
    ("実績_上位系対策あり", _gs(lambda g, row: g.memo_summary(row)["yes"] or "")),
    ("実績_記載なし", _gs(lambda g, row: g.memo_summary(row)["unk"] or "")),
    ("実績_うち変電所は推定", _gs(lambda g, row: g.memo_summary(row)["est"] or "")),
    ("つながりデータ", _gs(lambda g, row: g.topo_available(row))),
    ("つながる送電線の本数", _gs(lambda g, row: g.links_count(row))),
    ("上位側の送電線（相手側・上位側の空き）", _gs(lambda g, row: g.up_lines_text(row))),
    ("上位系の制約設備（▼＝空き0以下）", _gs(lambda g, row: g.upper_equipment_text(row))),
]


def matches_region(r, words):
    if not words:
        return True
    hay = [r["area"], r["prefAny"], r.get("muni") or "", str(r["name"]), r.get("addr") or ""]
    return any(w in h for w in words for h in hay)


def select(rows, args):
    words = [w.strip() for w in (args.region or "").replace("、", ",").split(",") if w.strip()]
    ranks = [x.strip() for x in args.ranks.split(",") if x.strip()]
    out = []
    for r in rows:
        if not matches_region(r, words):
            continue
        if r["rank"] not in ranks:
            continue
        if args.up == "ok" and r["availUp"] is None:
            continue
        if args.up == "ng" and (r["availUp"] is not None or not has_data(r)):
            continue
        if args.min_cap is not None and not (eff_cap(r) is not None and eff_cap(r) >= args.min_cap):
            continue
        if args.n1 and r["n1"] != "可":
            continue
        if args.no_curtail and r["curtail"] != "なし":
            continue
        if args.with_coords and r["lat"] is None:
            continue
        out.append(r)
    cap = lambda r: eff_cap(r) if eff_cap(r) is not None else float("-inf")  # noqa: E731
    if args.sort == "cap":
        out.sort(key=lambda r: (-cap(r), -r["grid"]))
    elif args.sort == "region":
        out.sort(key=lambda r: (r["prefAny"] or "~", r.get("muni") or "~", -r["grid"]))
    else:
        out.sort(key=lambda r: (-r["grid"], -cap(r)))
    return out


def write_csv(path, rows):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([c[0] for c in COLUMNS])
        for r in rows:
            w.writerow([blank(fn(r)) for _, fn in COLUMNS])


def counts(rows):
    sa = [r for r in rows if r["rank"] in ("S", "A")]
    return len(rows), len(sa), sum(1 for r in sa if r["availUp"] is not None)


def by_pref(rows, meta, outdir):
    os.makedirs(outdir, exist_ok=True)
    prefs = {}
    for r in rows:
        prefs.setdefault(r["prefAny"] or "都道府県不明", []).append(r)
    key = lambda r: (-r["grid"], r["rank"] not in ("S", "A"))  # noqa: E731
    lines = [
        "# 変電所 地域別CSV",
        "",
        f"自動生成: {datetime.date.today().isoformat()}（`scripts/export_substations.py --by-pref`。"
        f"元データ生成 {meta.get('generated', '')}）。毎月の座標更新のたびに作り直されます。",
        "",
        "各都道府県の全変電所（全ランク）を系統スコア順に並べたCSVです。Excelでそのまま開けます。"
        "「系統ランク」「上位系の確認」列でフィルタしてください。",
        "",
        "地図モード（substation.html#map）と同じ判定の列も入っています：「カテゴリ（地図の色）」（緑＝両方空きあり／黄＝変電所は空きあり・上位系は不明／"
        "紫＝変電所は空きあり・上位系は空き無し＝要接続検討／赤＝変電所単体で空き無し／灰＝データ無し）と「変電所単体の空き」「上位系の空き」「上位系統の判定理由」、"
        "「要接続検討」（○＝紫）、「接続検討の実績（件数・分類）」（対策なし／バンク等の対策待ち／"
        "上位系対策あり／記載なし、うち変電所は推定）、「上位側の送電線（相手側・上位側の空き）」、「上位系の制約設備（▼＝空き0以下）」。"
        "つながりの元は東北・北陸・関西・中国・四国・九州・北海道の公表物で、東京・中部・沖縄は「対象外エリア」になります。",
        "",
        "| ファイル | 全件 | S/A | うち上位系確認済み |",
        "|---|---|---|---|",
    ]
    sa_all = sorted([r for r in rows if r["rank"] in ("S", "A")], key=key)
    write_csv(os.path.join(outdir, "全国_SA.csv"), sa_all)
    n, s, u = counts(sa_all)
    lines.append(f"| [全国_SA.csv](全国_SA.csv) | {n:,} | {s:,} | {u:,} |")
    order = sorted(prefs, key=lambda p: (p == "都道府県不明", p))
    for p in order:
        rs = sorted(prefs[p], key=key)
        write_csv(os.path.join(outdir, f"{p}.csv"), rs)
        n, s, u = counts(rs)
        lines.append(f"| [{p}.csv]({p}.csv) | {n:,} | {s:,} | {u:,} |")
    lines.append("")
    with open(os.path.join(outdir, "README.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"出力: {outdir}/ に {len(order) + 1}ファイル（都道府県{len(order)}＋全国_SA）")


def main():
    ap = argparse.ArgumentParser(description="変電所を地域・条件で絞り込んでCSVに払い出す")
    ap.add_argument("--data", default=DEFAULT_JS, help="data/substations.js のパス")
    ap.add_argument("--region", help="エリア・都道府県・市区町村・変電所名の部分一致。カンマ区切りでOR（例: 秋田県 / 横手市,大仙市 / 中部）")
    ap.add_argument("--ranks", default="S,A", help=f"系統ランク（カンマ区切り）。既定 S,A。指定できる値: {','.join(ALL_RANKS)}")
    ap.add_argument("--up", choices=["all", "ok", "ng"], default="all",
                    help="上位系の確認: all=すべて / ok=上位系考慮で採点したものだけ / ng=当該設備で代用したものだけ")
    ap.add_argument("--min-cap", type=float, help="空容量（上位系優先、無ければ当該）の下限MW")
    ap.add_argument("--n1", action="store_true", help="N-1電制「可」だけ")
    ap.add_argument("--no-curtail", action="store_true", help="出力制御「なし」だけ")
    ap.add_argument("--with-coords", action="store_true", help="座標がある変電所だけ")
    ap.add_argument("--sort", choices=["grid", "cap", "region"], default="grid", help="並び順")
    ap.add_argument("--out", help="出力CSVのパス")
    ap.add_argument("--by-pref", metavar="DIR", help="都道府県別CSVを一括生成するディレクトリ")
    args = ap.parse_args()

    meta, rows = load(args.data)
    global GS
    if os.path.exists(TOPO_JS):
        GS = GridStatus(args.data, TOPO_JS)
    else:
        print("※ data/grid_topology.js が無いため、上位系統の判定列は空欄になります", file=sys.stderr)
    if args.by_pref:
        by_pref(rows, meta, args.by_pref)
        return
    if not args.out:
        ap.error("--out か --by-pref を指定してください")
    sel = select(rows, args)
    write_csv(args.out, sel)
    n, s, u = counts(sel)
    print(f"該当 {n}件（S/A {s}件・うち上位系確認済み {u}件） → {args.out}")
    if not sel:
        print("※ 条件に合う変電所がありません。地域名の表記（「県」「市」の有無）や条件を見直してください。", file=sys.stderr)


if __name__ == "__main__":
    main()
