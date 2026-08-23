#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""蓄電池条例リスクの定期監視ツール。

data/ordinance_registry.json（条例・開発許可運用の台帳）を使って3つのことを行う。

  list     台帳の一覧を表示する
  check    案件の所在地（都道府県・市区町村）に該当する規制を照合してレポートを出す
  monitor  台帳の各URLを取得して前回スナップショットと比較し、変更・消失を検知する
  report   台帳から docs/条例リスク台帳.md を生成する

いずれも標準ライブラリのみで動く。monitor はネットワークが必要（GitHub Actionsでの
週次実行を想定）。使い方の詳細は docs/ordinance_watch_運用手順.md を参照。
"""

import argparse
import hashlib
import json
import re
import ssl
import sys
import time
import unicodedata
import urllib.request
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
REGISTRY_PATH = ROOT / "data" / "ordinance_registry.json"
STATE_PATH = ROOT / "data" / "ordinance_page_state.json"

UA = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0 Safari/537.36 ordinance-watch/1.0"
)


def load_registry(path=REGISTRY_PATH):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------- fetch/normalize

def fetch(url, timeout=30, retries=3):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    ctx = ssl.create_default_context()
    last_exc = None
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
                return resp.read()
        except Exception as exc:  # 自治体サイトは一時的な接続リセットが多いため指数バックオフで再試行
            last_exc = exc
            time.sleep(2 ** attempt)
    raise last_exc


def normalize_page(raw: bytes, url: str) -> str:
    """比較用にページ本文を正規化する。

    PDFはバイト列のまま扱い、HTMLはタグ・スクリプト・空白を落として本文だけにする。
    アクセスカウンタ等の微小な揺れで誤検知しないため、数字のみの短いトークンは除く。
    """
    if url.lower().endswith(".pdf") or raw[:5] == b"%PDF-":
        return hashlib.sha256(raw).hexdigest()
    for enc in ("utf-8", "cp932", "euc-jp"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("utf-8", errors="replace")
    text = re.sub(r"(?is)<(script|style|noscript)[^>]*>.*?</\1>", " ", text)
    text = re.sub(r"(?s)<!--.*?-->", " ", text)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    text = unicodedata.normalize("NFKC", text)
    tokens = [t for t in text.split() if not re.fullmatch(r"\d{1,7}", t)]
    return " ".join(tokens)


def digest(url):
    raw = fetch(url)
    norm = normalize_page(raw, url)
    return hashlib.sha256(norm.encode("utf-8")).hexdigest(), len(raw)


# ---------------------------------------------------------------- commands

def cmd_list(args):
    reg = load_registry()
    print(f"台帳更新日: {reg['updated']}  件数: {len(reg['entries'])}")
    print()
    for e in reg["entries"]:
        where = e["muni"] or e["pref"] or "全国"
        eff = e["effective_date"] or "施行日不明"
        print(f"- [{e['level']}] {where}: {e['title']}（{e['reg_type']}／{eff}）")
        print(f"    {e['url']}")
    return 0


def _match(entry, pref, muni):
    if entry["level"] == "国":
        return True
    if entry["level"] == "都道府県":
        return pref and entry["pref"] == pref
    if entry["level"] == "市区町村":
        return bool(muni and entry["muni"] == muni and (not pref or entry["pref"] == pref))
    return False


CHECKLIST = """\
## 台帳ヒットの有無に関わらず確認すること（案件共通チェックリスト）

台帳は既知の公表事例のみを収録している。**未収録の自治体でも、国交省技術的助言
（国都計第7号・2025-04-08）を根拠に開発許可の運用がいつでも始まりうる。**

1. 開発許可担当課への事前照会（系統用蓄電池を第一種特定工作物として扱うか、運用基準の有無）
2. 所轄消防への事前協議（火災予防条例の蓄電池設備基準・届出）
3. 例規集で「蓄電池」「再生可能エネルギー」「開発」「環境保全」等の条例・要綱を検索
   （多くの自治体は https://www1.g-reiki.net/ 等で例規集を公開）
4. 議会会議録・パブリックコメントで条例制定の動きを検索（制定前の察知）
5. 市街化調整区域の場合は開発許可の要否と標準処理期間を必ず確認
"""


def cmd_check(args):
    reg = load_registry()
    hits = [e for e in reg["entries"] if _match(e, args.pref, args.muni)]
    local_hits = [e for e in hits if e["level"] != "国"]
    lines = []
    where = " ".join(x for x in (args.pref, args.muni) if x)
    lines.append(f"# 条例リスク照合結果: {where or '（所在地未指定）'}")
    lines.append("")
    lines.append(f"台帳更新日: {reg['updated']} ／ 照合日: {date.today().isoformat()}")
    lines.append("")
    if local_hits:
        lines.append(f"## ⚠ 所在自治体に該当する規制 {len(local_hits)}件")
        lines.append("")
        for e in local_hits:
            lines.append(f"### {e['title']}（{e['level']}・{e['reg_type']}）")
            lines.append(f"- 対象: {e['scope']}")
            lines.append(f"- 手続き: {e['requirement']}")
            lines.append(f"- 施行/運用開始: {e['effective_date'] or '不明（ページ参照）'}")
            lines.append(f"- DDへの影響: {e['dd_impact']}")
            lines.append(f"- 出典: {e['url']}（最終確認 {e['last_verified']}）")
            lines.append("")
    else:
        lines.append("## 所在自治体に該当する台帳エントリはなし")
        lines.append("")
        lines.append("ただし未収録＝規制なしではない。下のチェックリストを必ず実施すること。")
        lines.append("")
    lines.append("## 全国共通で適用される枠組み")
    lines.append("")
    for e in [e for e in hits if e["level"] == "国"]:
        lines.append(f"- {e['title']}: {e['dd_impact']}")
    lines.append("")
    lines.append(CHECKLIST)
    out = "\n".join(lines)
    if args.out:
        Path(args.out).write_text(out, encoding="utf-8")
        print(f"書き出し: {args.out}")
    else:
        print(out)
    return 0


def cmd_monitor(args):
    reg = load_registry()
    state = {}
    if STATE_PATH.exists():
        state = json.loads(STATE_PATH.read_text(encoding="utf-8"))
    prev = state.get("pages", {})
    today = date.today().isoformat()
    new_pages = {}
    changed, added, failed = [], [], []
    targets = [(e["url"], f"{e['muni'] or e['pref'] or '国'}: {e['title']}") for e in reg["entries"]]
    targets += [(w["url"], f"定点: {w['title']}") for w in reg.get("watchpoints", [])]
    for url, label in targets:
        try:
            h, size = digest(url)
        except Exception as exc:  # ネットワーク・404等はエラー扱いで報告して続行
            failed.append((label, url, str(exc)))
            if url in prev:
                new_pages[url] = prev[url]  # 取得失敗時は前回値を保持
            continue
        new_pages[url] = {"hash": h, "bytes": size, "checked": today}
        if url not in prev:
            added.append((label, url))
        elif prev[url]["hash"] != h:
            changed.append((label, url, prev[url].get("checked", "?")))

    print(f"監視対象 {len(targets)}件 ／ 変更 {len(changed)} ／ 初回登録 {len(added)} ／ 取得失敗 {len(failed)}")
    print()
    if changed:
        print("## 🔄 前回から内容が変わったページ（規制の改正・追加の可能性）")
        for label, url, last in changed:
            print(f"- {label}")
            print(f"  {url}（前回確認: {last}）")
        print()
    if added:
        print("## 🆕 今回から監視を開始したページ")
        for label, url in added:
            print(f"- {label}\n  {url}")
        print()
    if failed:
        print("## ⚠ 取得に失敗したページ（URL変更・削除の可能性。手動確認が必要）")
        for label, url, err in failed:
            print(f"- {label}\n  {url}\n  エラー: {err}")
        print()

    if args.update:
        STATE_PATH.write_text(
            json.dumps({"updated": today, "pages": new_pages}, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        print(f"スナップショット更新: {STATE_PATH.relative_to(ROOT)}")

    # CIで「要確認あり」を検知できるよう、変更または取得失敗があれば終了コード1
    return 1 if (changed or failed) else 0


def cmd_report(args):
    reg = load_registry()
    lines = [
        "# 蓄電池条例リスク台帳",
        "",
        "<!-- このファイルは scripts/ordinance_watch.py report で自動生成。直接編集せず "
        "data/ordinance_registry.json を更新して再生成すること -->",
        "",
        f"台帳更新日: {reg['updated']} ／ 収録 {len(reg['entries'])}件 ／ 生成日: {date.today().isoformat()}",
        "",
        "系統用蓄電池（蓄電所）の設置に影響する条例・開発許可運用・国の通知の一覧。",
        "国交省技術的助言（国都計第7号・2025-04-08）を受けて、危険物を含有する系統用蓄電池を",
        "都市計画法の第一種特定工作物として扱う自治体が増加中。**未収録の自治体でも同様の運用が",
        "いつでも始まりうる**ため、案件ごとに `python scripts/ordinance_watch.py check` と",
        "窓口照会を必ず行うこと。",
        "",
        "| レベル | 自治体 | 名称 | 種別 | 手続き | 施行/運用開始 | 最終確認 |",
        "|---|---|---|---|---|---|---|",
    ]
    order = {"国": 0, "都道府県": 1, "市区町村": 2}
    for e in sorted(reg["entries"], key=lambda x: (order.get(x["level"], 9), x["pref"] or "", x["muni"] or "")):
        where = e["muni"] or e["pref"] or "全国"
        req = e["requirement"].split("（")[0]
        lines.append(
            f"| {e['level']} | {where} | [{e['title']}]({e['url']}) | {e['reg_type']} "
            f"| {req} | {e['effective_date'] or '−'} | {e['last_verified']} |"
        )
    wps = reg.get("watchpoints", [])
    if wps:
        lines += ["", "## 定点監視ページ（規制そのものではないが新規制の早期検知に使う）", ""]
        for w in wps:
            lines.append(f"- [{w['title']}]({w['url']}) — {w['why']}（最終確認 {w['last_verified']}）")
    lines += ["", "## 各エントリの詳細", ""]
    for e in reg["entries"]:
        where = e["muni"] or e["pref"] or "全国"
        lines += [
            f"### {where}: {e['title']}",
            "",
            f"- 種別: {e['reg_type']}（{e['level']}）",
            f"- 対象: {e['scope']}",
            f"- 手続き: {e['requirement']}",
            f"- 公表日: {e['announced_date'] or '不明'} ／ 施行・運用開始: {e['effective_date'] or '不明'}",
            f"- 内容: {e['summary']}",
            f"- DDへの影響: {e['dd_impact']}",
            f"- 出典: {e['url']}（最終確認 {e['last_verified']}）",
        ]
        if e.get("notes"):
            lines.append(f"- 備考: {e['notes']}")
        lines.append("")
    out_path = Path(args.out)
    out_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"生成: {out_path}")
    return 0


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("list", help="台帳の一覧を表示")

    pc = sub.add_parser("check", help="案件所在地に該当する規制を照合")
    pc.add_argument("--pref", help="都道府県名（例: 千葉県）")
    pc.add_argument("--muni", help="市区町村名（例: 千葉市）")
    pc.add_argument("--out", help="結果Markdownの書き出し先（省略時は標準出力）")

    pm = sub.add_parser("monitor", help="台帳URLの変更・消失を検知")
    pm.add_argument("--update", action="store_true", help="スナップショット(data/ordinance_page_state.json)を更新")

    pr = sub.add_parser("report", help="台帳Markdownを生成")
    pr.add_argument("--out", default=str(ROOT / "docs" / "条例リスク台帳.md"))

    args = p.parse_args()
    return {"list": cmd_list, "check": cmd_check, "monitor": cmd_monitor, "report": cmd_report}[args.cmd](args)


if __name__ == "__main__":
    sys.exit(main())
