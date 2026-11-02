#!/usr/bin/env python3
"""社内版の変電所マップを、1ファイルのHTMLとして作る。

実績（接続検討回答）は案件情報なので公開リポジトリに置かない。公開リポジトリの外にある
実績JSONと回答書取込台帳CSVを読み、実績入りのHTMLをリポジトリの外へ書き出す。

  python scripts/build_internal_map.py --results 実績DB.json --ledger 回答書取込台帳.csv \
      --out ../変電所マップ_社内版.html

・台帳は「状態＝取込済み」かつ今回の取込分（取込日が日付のもの）だけを実績に足す
・出力先がリポジトリの中なら止める
"""
import argparse, csv, json, os, re, subprocess, sys, tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
LIST = os.path.join(ROOT, "data", "全国変電所リスト_66kV以上_座標追加.xlsx")


def known_substations():
    src = open(os.path.join(ROOT, "data", "substations.js"), encoding="utf-8").read()
    d = json.loads(src[src.index("window.SUBSTATION_DATA=") + 23:].rstrip().rstrip(";"))
    f = {k: i for i, k in enumerate(d["fields"])}
    return {d["dict"]["area"][r[f["area"]]] + str(r[f["name"]]).strip() for r in d["rows"]}


def man_yen(v):
    return f"{float(v) / 10000:,.0f}万円"


def ledger_records(path, known):
    recs = []
    for r in csv.DictReader(open(path, encoding="utf-8-sig")):
        if r["状態"] != "取込済み" or not re.match(r"\d{4}-\d\d-\d\d$", r["取込日"]):
            continue
        area = r["電力会社"]
        first = (r["変電所"].split() or [""])[0]
        name = first if area + first in known else "－(" + (r["変電所"] or r["案件名"]) + ")"
        cost = (man_yen(r["負担金(円)"]) + (f"({r['税区分']})" if r["税区分"] else "")) if r["負担金(円)"] else "－"
        term = f"入金後{r['工期(月)']}ヶ月" if r["工期(月)"] else "－"
        kw = f"{float(r['放電kW']):,.0f}kW" if r["放電kW"] else "－"
        upper = r["上位系統増強"] or "記載なし"
        note = f"{r['連系条件']} ／上位系: {upper}（AI読取・信頼度{r['読取信頼度']}・要原本確認）"
        recs.append([area, name, r["案件名"], kw, cost, term, "", note,
                     f"{r['回答種別']}: {cost}・{term}({r['回答日']}) {r['連系条件']}"])
    return recs


def main():
    ap = argparse.ArgumentParser(description="社内版の変電所マップ（1ファイルHTML）を作る")
    ap.add_argument("--results", required=True, help="実績JSON（公開リポジトリの外）")
    ap.add_argument("--ledger", help="回答書取込台帳CSV（公開リポジトリの外）")
    ap.add_argument("--out", required=True, help="出力HTML（公開リポジトリの外）")
    a = ap.parse_args()
    out = os.path.abspath(a.out)
    if out.startswith(ROOT + os.sep):
        sys.exit("出力先が公開リポジトリの中です。外に出してください: " + out)

    base = json.load(open(a.results, encoding="utf-8"))
    added = ledger_records(a.ledger, known_substations()) if a.ledger else []
    base["records"] = base["records"] + added
    with tempfile.TemporaryDirectory() as tmp:
        merged, js = os.path.join(tmp, "r.json"), os.path.join(tmp, "s.js")
        json.dump(base, open(merged, "w", encoding="utf-8"), ensure_ascii=False)
        env = dict(os.environ, BESS_PRIVATE_RESULTS=merged)
        subprocess.run([sys.executable, os.path.join(ROOT, "scripts", "build_substation_web_data.py"),
                        "--in", LIST, "--out", js], check=True, env=env, stdout=subprocess.DEVNULL)
        data = open(js, encoding="utf-8").read()

    def script(body):
        return "<script>\n" + body.replace("</script", "<\\/script") + "\n</script>"
    read = lambda n: open(os.path.join(ROOT, "data", n), encoding="utf-8").read()
    html = open(os.path.join(ROOT, "substation.html"), encoding="utf-8").read()
    html = html.replace('<script src="data/substations.js"></script>', script(data))
    html = html.replace('<script src="data/grid_topology.js"></script>', script(read("grid_topology.js")))
    html = html.replace("loadScript('data/chuden_grid.js')", "Promise.resolve()")
    html = html.replace("</head>", '<meta name="robots" content="noindex">\n' + script(read("chuden_grid.js")) + "\n</head>", 1)
    banner = ('<div style="background:#b00020;color:#fff;padding:8px 16px;font-weight:700">社内限定：接続検討回答の実績を含みます。'
              '社外へ渡さないでください（回答書は目的外使用・第三者開示が禁止されています）。</div>')
    html = re.sub(r"(<body[^>]*>)", lambda m: m.group(1) + banner, html, count=1)
    html = html.replace("<title>", "<title>【社内版】", 1)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    open(out, "w", encoding="utf-8").write(html)
    print(f"出力: {out}  実績{len(base['records'])}件（台帳から{len(added)}件）  {os.path.getsize(out) / 1e6:.1f}MB")


if __name__ == "__main__":
    main()
