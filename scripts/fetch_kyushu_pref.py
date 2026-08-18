#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
九州電力送配電の地区別「送電系統図」PDFから、変電所の所在都道府県を判定する。

九州エリアの行は元リストで都道府県が空のため、変電所名だけで地図を検索することに
なり手作業の当たりが付けられない。九電の系統図は地区ごと（〔２２熊本〕など31本）に
分かれており、どの地区図に載るかで都道府県まで絞り込める。

【都道府県は断定できない】当初は地区図から都道府県を決めて本体xlsxに書き込むつもりだったが、
検証した結果その方法は誤る。隣接地区の設備が図の端にも描かれるため、「地区図Xに載る＝
Xの県にある」が成り立たない。実際に次の誤りが出た。

  村所変電所   〔２４人吉〕(熊本)に載るが実際は宮崎県西米良村
  日南変電所   〔３０大隅〕(鹿児島)に載るが実際は宮崎県日南市
  諸富変電所   〔１４大牟田〕(福岡)に載るが実際は佐賀県佐賀市諸富町

地理院の地名検索で検算しようとしたが、全国の同名地名を拾うため検証器にならない
（111件中72件が九州外の地名にヒットした）。誤った都道府県は誤った座標を記録させる
原因になるので、書き込みは行わない。

【このスクリプトが出すもの】「どの地区図に載るか」という事実だけを一覧にする。
断定はしないが、全国から探すのと〔２２熊本〕周辺を見るのとでは手間が大きく違うため、
座標を手作業で調べるときの当たりを付ける材料になる。座標そのものは得られない
（系統図は模式図で実際の位置を示さない）。

使い方:
  # ① 地区図を取得して名称を抽出（要ネットワーク）
  python scripts/fetch_kyushu_pref.py extract \
    --index-url https://www.kyuden.co.jp/var/rev0/0900/4946/td_renkeimap_260730.pdf \
    --out kyushu_cache/regions.json

  # ② 九州行がどの地区図に載るかの対応表を出す（ネットワーク不要）
  python scripts/fetch_kyushu_pref.py report \
    --regions kyushu_cache/regions.json \
    --in "data/全国変電所リスト_66kV以上_座標追加.xlsx" \
    --out "docs/九州_地区図対応.csv"
"""
import argparse
import csv
import json
import os
import re
import urllib.request

from fetch_substation_coords import read_rows

UA = {"User-Agent": "Mozilla/5.0"}
TARGET_AREA = "九州"

# 地区名（PDFの〔〕見出し）に含まれる地名 → 都道府県。
# 見出しに複数県の地名が並ぶ地区（久留米・鳥栖 等）は候補が2県になり、
# 他の地区図と併せても1県に絞れなければ採用しない。
REGION_PREF = {
    "門司": "福岡県", "苅田": "福岡県", "小倉": "福岡県", "若松": "福岡県",
    "八幡": "福岡県", "遠賀": "福岡県", "古賀": "福岡県", "飯塚": "福岡県",
    "田川": "福岡県", "東福岡": "福岡県", "南福岡": "福岡県", "赤坂": "福岡県",
    "伊都": "福岡県", "豊前": "福岡県", "高田": "福岡県", "中央": "福岡県",
    "山家": "福岡県", "久留米": "福岡県", "大牟田": "福岡県",
    "宇佐": "大分県", "中津": "大分県", "大分": "大分県", "日田": "大分県",
    "佐伯": "大分県", "竹田": "大分県",
    "鳥栖": "佐賀県", "唐津": "佐賀県", "佐賀": "佐賀県", "武雄": "佐賀県",
    "東佐世保": "長崎県", "長崎": "長崎県", "諫早": "長崎県",
    "熊本": "熊本県", "八代": "熊本県", "水俣": "熊本県", "天草": "熊本県",
    "人吉": "熊本県",
    "日向": "宮崎県", "宮崎": "宮崎県", "都城": "宮崎県",
    "川内": "鹿児島県", "出水": "鹿児島県", "霧島": "鹿児島県",
    "鹿児島": "鹿児島県", "大隅": "鹿児島県",
}

# 離島図は〔離島①〜④〕に分かれ、本文に島名が入る。島名 → 都道府県。
ISLAND_PREF = {
    "小値賀": "長崎県", "五島列島": "長崎県", "福江": "長崎県",
    "対馬": "長崎県", "壱岐": "長崎県",
    "種子島": "鹿児島県", "徳之島": "鹿児島県", "奄美大島": "鹿児島県",
    "甑島": "鹿児島県", "沖永良部島": "鹿児島県", "与論島": "鹿児島県",
    "喜界島": "鹿児島県",
}


def region_prefs(header):
    """地区見出しから候補県の集合を返す。"""
    return {p for k, p in REGION_PREF.items() if k in header}


def tokens_of(text):
    """図中のラベルを取り出す。数字だけのもの（設備番号）は落とす。"""
    return {x for x in re.split(r"[\s\n]+", text)
            if x and not re.fullmatch(r"[\d\-\.,]+", x)}


def extract(args):
    from pdfminer.high_level import extract_text

    if args.index_url:
        idx = urllib.request.urlopen(
            urllib.request.Request(args.index_url, headers=UA), timeout=90).read()
    else:
        idx = open(args.index_pdf, "rb").read()
    urls = sorted({u.decode("latin1") for u in re.findall(rb"/URI\s*\(([^)]+)\)", idx)})
    urls = [u for u in urls if re.search(r"/td_\d{2}[a-z_-]*_map_\d+\.pdf$", u)]
    # 同じ地区に旧日付のURLが残っていることがあるので、地区番号ごとに1本へ寄せる
    latest = {}
    for u in urls:
        no = re.search(r"/td_(\d{2})", u).group(1)
        latest[no] = max(latest.get(no, ""), u)
    urls = [latest[k] for k in sorted(latest)]

    os.makedirs(args.cache_dir, exist_ok=True)
    regions = []
    for u in urls:
        fn = os.path.join(args.cache_dir, u.rsplit("/", 1)[1])
        if not (os.path.exists(fn) and os.path.getsize(fn) > 1000):
            data = urllib.request.urlopen(
                urllib.request.Request(u, headers=UA), timeout=120).read()
            with open(fn, "wb") as f:
                f.write(data)
        text = extract_text(fn)
        heads = re.findall(r"〔([^〕]+)〕", text)
        head = next((h for h in heads if h != "凡例"), "")
        is_island = "離島" in head
        if is_island:
            # 〔離島①〕〜〔離島④〕の各節に島名が入るので、節ごとに切って島名で県を決める
            parts = re.split(r"〔離島[^〕]*〕", text)[1:]
            for part in parts:
                toks = tokens_of(part)
                prefs = {p for k, p in ISLAND_PREF.items()
                         if k in re.sub(r"\s+", "", part)}
                regions.append({"name": "離島", "island": True,
                                "prefs": sorted(prefs), "tokens": sorted(toks)})
        else:
            regions.append({"name": head, "island": False,
                            "prefs": sorted(region_prefs(head)),
                            "tokens": sorted(tokens_of(text))})
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(regions, f, ensure_ascii=False)
    noprefs = [r["name"] for r in regions if not r["prefs"]]
    print(f"地区 {len(regions)}件 → {args.out}")
    if noprefs:
        print(f"※ 県を割り当てられなかった地区: {noprefs}")


def report(args):
    regions = json.load(open(args.regions, encoding="utf-8"))
    for r in regions:
        r["tokens"] = set(r["tokens"])
    rows, _, _ = read_rows(args.src, args.list_sheet)

    out, found = [], 0
    for row in rows:
        if row["area"] != TARGET_AREA or row.get("pref"):
            continue
        base = re.sub(r"(変電所|開閉所|変換所)$", "", str(row["name"] or ""))
        base = re.sub(r"\(\d+kV\)", "", base)
        if not base:
            continue
        hits = [r for r in regions if base in r["tokens"]]
        names = [r["name"] for r in hits]
        cand = sorted({p for r in hits for p in r["prefs"]})
        if hits:
            found += 1
        out.append([row["no"], row["name"], "/".join(names) or "（載っていない）",
                    "/".join(cand), "" if len(cand) == 1 else "要確認"])

    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["# 都道府県は断定していません。隣接地区の設備も図の端に描かれるため、"
                    "地区図に載ることは所在県の証明になりません（村所=宮崎が熊本の図に載る等）。"])
        w.writerow(["No.", "変電所名", "載っている地区図", "候補の都道府県", "備考"])
        w.writerows(out)
    print(f"九州で都道府県が空の行: {len(out)}件 / 地区図に載っていたもの {found}件")
    print(f"出力: {args.out}")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    e = sub.add_parser("extract", help="地区図PDFを取得して名称を抽出する")
    g = e.add_mutually_exclusive_group(required=True)
    g.add_argument("--index-url", help="地区図へのリンクを持つ連系マップPDFのURL")
    g.add_argument("--index-pdf", help="同PDFのローカルパス")
    e.add_argument("--cache-dir", default="kyushu_cache/pdf")
    e.add_argument("--out", default="kyushu_cache/regions.json")

    f = sub.add_parser("report", help="九州行がどの地区図に載るかの対応表を出す")
    f.add_argument("--regions", default="kyushu_cache/regions.json")
    f.add_argument("--in", dest="src", required=True)
    f.add_argument("--out", default="docs/九州_地区図対応.csv")
    f.add_argument("--list-sheet", default="全国変電所リスト")

    args = ap.parse_args()
    extract(args) if args.cmd == "extract" else report(args)


if __name__ == "__main__":
    main()
