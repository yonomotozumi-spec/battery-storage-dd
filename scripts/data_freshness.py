#!/usr/bin/env python3
"""変電所リストの「スコア元データ」が各エリアでいつ時点かを集計する。

本体リスト（xlsx）の「データソース」シートの「データ日付」欄から、
エリアごとの ①公表日（各社が公表した時点）と ②最終取込日（本ツールに取り込んだ日）を読み、
今日までの経過日数を出す。自動取得できないエリア（東京・中部・沖縄）は確認方法を添える。

毎月5日の capacity-refresh から呼び、結果を Issue に載せて手作業の確認を促す。

  python scripts/data_freshness.py --in data/全国変電所リスト_66kV以上_座標追加.xlsx --out freshness.md
"""
import argparse
import calendar
import datetime as dt
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from update_capacity import MANUAL_AREAS  # noqa: E402

# 公表日がこれより古いエリアは「要確認」として印を付ける（各社の公表はおおむね月次〜四半期）
STALE_DAYS = 100


def _parse_dates(text):
    """文字列中の日付を古い順ではなく出現順に返す。YYYY-MM-DD／YYYY年M月D日／YYYY年M月末 に対応。"""
    out = []
    for m in re.finditer(r"(\d{4})[-/年](\d{1,2})[-/月](?:(\d{1,2})日?|末)", text):
        y, mo, d = int(m.group(1)), int(m.group(2)), m.group(3)
        day = int(d) if d else calendar.monthrange(y, mo)[1]
        try:
            out.append((m.start(), dt.date(y, mo, day)))
        except ValueError:
            pass
    return out


def read_sources(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["データソース"]
    rows = []
    header_seen = False
    for r in ws.iter_rows(values_only=True):
        cells = ["" if c is None else str(c) for c in r]
        if not header_seen:
            header_seen = cells[:1] == ["エリア"]
            continue
        area = cells[0].strip()
        if not area or area == "全国共通":
            continue
        date_text = cells[3] if len(cells) > 3 else ""
        url = cells[4] if len(cells) > 4 else ""
        found = _parse_dates(date_text)
        auto = None
        m = re.search(r"自動更新\s*([^）)]+)", date_text)
        if m:
            a = _parse_dates(m.group(1))
            if a:
                auto = a[0][1]
        published = None
        for pos, d in found:
            if auto is None or d != auto:
                published = d
                break
        rows.append({"area": area, "published": published, "fetched": auto,
                     "date_text": date_text.strip(), "url": url.strip()})
    wb.close()
    return rows


def build(rows, today):
    def age(d):
        return None if d is None else (today - d).days

    lines = ["| エリア | 公表日 | 本ツールへの取込 | 公表からの経過 | 取得方法 |",
             "|---|---|---|---|---|"]
    stale = []
    for r in rows:
        a = age(r["published"])
        manual = r["area"] in MANUAL_AREAS
        warn = manual or (a is not None and a > STALE_DAYS) or r["published"] is None
        how = "手作業（自動取得できず）" if manual else "自動（毎月5日）"
        pub = r["published"].isoformat() if r["published"] else "不明"
        got = r["fetched"].isoformat() if r["fetched"] else "—"
        ago = f"{a}日" if a is not None else "不明"
        lines.append(f"| {'⚠ ' if warn else ''}{r['area']} | {pub} | {got} | {ago} | {how} |")
        if warn:
            stale.append(r)

    out = ["## エリアごとのデータの新しさ（" + today.isoformat() + "時点）", ""] + lines + [""]
    if stale:
        out += [f"### ⚠ 手作業での確認をお願いします（{len(stale)}エリア）", ""]
        for r in stale:
            a = age(r["published"])
            ago = f"{a}日前の公表値のまま" if a is not None else "公表日が不明"
            out.append(f"- **{r['area']}**（{ago}）: {MANUAL_AREAS.get(r['area'], '公表日が古いため最新版を確認')}")
            if r["url"]:
                out.append(f"  - 公表ページ: {r['url']}")
        out += ["",
                "落としたCSV（ZIPのままでも可）をClaudeに渡すか、次のコマンドで取り込みます。",
                "",
                "```bash",
                "python scripts/update_capacity.py add --area 東京 --cache-dir capacity_cache 落としたファイル...",
                "python scripts/update_capacity.py apply --in data/全国変電所リスト_66kV以上_座標追加.xlsx \\",
                "  --cache-dir capacity_cache --out data/全国変電所リスト_66kV以上_座標追加.xlsx",
                "```",
                ""]
    else:
        out += ["全エリアが最新です。手作業の確認は不要です。", ""]
    return "\n".join(out), stale


def main():
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--in", dest="inp", required=True, help="本体リスト（xlsx）")
    p.add_argument("--out", help="markdownの出力先（省略時は標準出力）")
    p.add_argument("--summary-json", help="要確認エリア数などをJSONで書き出す")
    p.add_argument("--today", help="基準日（YYYY-MM-DD。試験用）")
    args = p.parse_args()

    today = dt.date.fromisoformat(args.today) if args.today else dt.date.today()
    rows = read_sources(args.inp)
    md, stale = build(rows, today)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(md)
    else:
        print(md)
    if args.summary_json:
        with open(args.summary_json, "w", encoding="utf-8") as f:
            json.dump({"today": today.isoformat(), "stale_areas": [r["area"] for r in stale],
                       "areas": [{"area": r["area"],
                                  "published": r["published"].isoformat() if r["published"] else None,
                                  "fetched": r["fetched"].isoformat() if r["fetched"] else None}
                                 for r in rows]}, f, ensure_ascii=False, indent=2)
    print(f"要確認エリア: {len(stale)}件 ({', '.join(r['area'] for r in stale) or 'なし'})",
          file=sys.stderr)


if __name__ == "__main__":
    main()
