#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
チャット（Claude）で回答書を読み取って書き出したCSVを、「接続検討回答 検証一覧」xlsx の
「回答一覧」シートに追記する。

入力CSV（UTF-8、1行=回答書1件）の列:
  案件番号, 案件名, 回答種別, 電力会社, 回答日, 申込者, 設置場所, 放電kW, 変電所, 負担金(円), 税区分,
  工期(月), 上位系統増強, 連系条件, 有効期限, 読取信頼度, ファイル名, 原本リンク, 取込日, 状態
  （状態: 取込済み／要確認／未読。xlsx に既にある行はファイル名＋回答種別＋回答日＋設置場所で見分けて飛ばす）

やること:
  1. 新しい行を「回答一覧」の末尾に追記（書式は最終行をコピー。数式列＝負担金(税込・万円)・負担金帯・工期帯・
     自動チェック は同じ数式を入れ、検証ステータスは「未確認」、原本リンクはハイパーリンク）
  2. 追記した行と同じ案件（案件番号、無ければ設置場所）の「集計対象」を付け直す
     ＝ 接続検討回答／再接続検討／契約申込回答で負担金(円)のある最新の回答だけ ○、他は －（凡例シートの定義どおり）
  3. 状態が要確認・未読、信頼度low、負担金か工期が無い行を「要確認」シートにも載せる
  4. 入力規則・条件付き書式・オートフィルタ・集計シートの範囲を新しい最終行まで広げる
  5. LibreOffice（soffice）で開き直して数式の値を計算し直す（import_connection_results.py は計算済みの値を読むため）

使い方:
  python scripts/append_connection_csv.py --csv 読取結果.csv \
      --in "data/接続検討回答_検証一覧.xlsx" --out "data/接続検討回答_検証一覧.xlsx"
  続けて import_connection_results.py → score_substations.py → build_substation_web_data.py →
  export_substations.py --by-pref exports を実行する（README「接続検討回答の実績DBへの取込」）。
"""
import argparse
import copy
import csv
import datetime
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.worksheet.datavalidation import DataValidation

SHEET = "回答一覧"
REVIEW_SHEET = "要確認"
SUMMARY_SHEET = "集計"
COUNT_TYPES = ("接続検討回答", "再接続検討", "契約申込回答")
FORMULA_COLS = {
    "負担金(税込・万円)": '=IF(O{r}="","",IF(P{r}="税抜",O{r}*1.1,O{r})/10000)',
    "負担金帯": ('=IF(Q{r}="","不明",IF(Q{r}<=500,"〜500万",IF(Q{r}<=1000,"500〜1,000万",'
               'IF(Q{r}<=3000,"1,000〜3,000万","3,000万〜"))))'),
    "工期帯": '=IF(U{r}="","不明",IF(U{r}<=6,"〜6ヶ月",IF(U{r}<=12,"7〜12ヶ月",IF(U{r}<=24,"13〜24ヶ月","25ヶ月〜"))))',
    "自動チェック": ('=IF(AD{r}<>"","未読",IF(AND(D{r}<>"接続検討回答",D{r}<>"再接続検討",D{r}<>"契約申込回答"),"",'
                 'TRIM(IF(O{r}="","負担金なし ","")&IF(U{r}="","工期なし ","")&IF(AND(O{r}<>"",OR(P{r}="",P{r}="不明")),'
                 '"税区分不明 ","")&IF(AC{r}="low","信頼度low",""))))'),
}


def nfkc(s):
    return unicodedata.normalize("NFKC", str(s or "")).replace("　", " ").strip()


def num(v):
    """CSVの数値欄 → float（空なら None）。"""
    s = nfkc(v).replace(",", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def normalize_upper(text):
    """「上位系統増強」欄の書き方を、実績の分類（grid_status.py の memo_upper）が読める形に揃える。
    「表の上位系増強工事は「-」」「上位系統工事 0.0百万円」「上位系統対策工事 0.00万円」は
    どれも上位系の対策なしの意味なので、既存の行と同じ「なし（…）」の形にする。"""
    t = nfkc(text)
    if not t:
        return None
    if re.match(r"^(表の)?上位系(統)?(増強|対策)?工事(欄)?は?「?[-－ー―]」?$", t):
        return "なし（上位系増強工事 -）"
    if re.match(r"^上位系(統)?(増強|対策)?工事\s*0(\.0+)?(百)?万円$", t):
        return f"なし（{t}）"
    return t


def cell_key(d):
    """既存行と新規行の同一判定キー。"""
    return (nfkc(d.get("ファイル名")), nfkc(d.get("回答種別")), nfkc(d.get("回答日")), nfkc(d.get("設置場所")))


def case_key(d):
    """集計対象を付け直す単位（案件番号 → 設置場所 → 案件名）。"""
    no = nfkc(d.get("案件番号"))
    if no:
        return "no:" + no
    site = re.split(r"他|ほか|（|\(", nfkc(d.get("設置場所")))[0].strip()
    if site:
        return "site:" + site
    return "name:" + nfkc(d.get("案件名"))


def date_key(v):
    s = nfkc(v)
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime("%Y-%m-%d")
    return s or "0000-00-00"


def shift_range(ref, last):
    """"AJ2:AJ125" のような範囲の終端行を last に置き換える。"""
    return re.sub(r"(\$?[A-Z]+\$?)(\d+)$", lambda m: f"{m.group(1)}{last}", ref)


def recalc_with_soffice(path):
    """LibreOffice で開き直して数式の計算値を埋める（soffice が無ければ警告だけ）。"""
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        print("警告: soffice が見つからないため数式の再計算をしていません。Excel で開いて保存し直してください。",
              file=sys.stderr)
        return False
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, "in.xlsx")
        shutil.copy(path, src)
        subprocess.run([exe, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", os.path.join(td, "out"), src],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
        out = os.path.join(td, "out", "in.xlsx")
        shutil.copy(out, path)
    return True


def main():
    ap = argparse.ArgumentParser(description="チャット読取CSVを検証一覧xlsxに追記")
    ap.add_argument("--csv", required=True)
    ap.add_argument("--in", dest="src", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--no-recalc", action="store_true", help="LibreOfficeでの再計算を行わない")
    args = ap.parse_args()

    with open(args.csv, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))

    wb = openpyxl.load_workbook(args.src)
    ws = wb[SHEET]
    header = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(header)}
    last = max(r for r in range(2, ws.max_row + 1) if ws.cell(r, col["No"]).value is not None)
    existing = {}
    for r in range(2, last + 1):
        d = {h: ws.cell(r, col[h]).value for h in header}
        existing[r] = d
    seen = {cell_key(d) for d in existing.values()}

    new = [row for row in rows if cell_key(row) not in seen]
    if not new:
        print("追記する行はありません（すべて既存）。")
        return
    style_row = last
    added = []
    today = datetime.date.today().isoformat()
    for row in new:
        r = ws.max_row + 1 if ws.max_row >= last + 1 and ws.cell(last + 1, 1).value is not None else last + 1
        last = r
        no = max(int(float(d["No"])) for d in existing.values()) + 1 if existing else 1
        # 書式を最終行からコピー
        for c in range(1, len(header) + 1):
            src = ws.cell(style_row, c)
            dst = ws.cell(r, c)
            if src.has_style:
                dst._style = copy.copy(src._style)
        state = nfkc(row.get("状態"))
        cond = nfkc(row.get("連系条件"))
        kw = num(row.get("放電kW"))
        values = {
            "No": no,
            "案件番号": nfkc(row.get("案件番号")) or None,
            "案件名": nfkc(row.get("案件名")) or None,
            "回答種別": nfkc(row.get("回答種別")) or None,
            "電力会社": nfkc(row.get("電力会社")) or None,
            "回答日": nfkc(row.get("回答日")) or None,
            "申込者": nfkc(row.get("申込者")) or None,
            "設置場所": nfkc(row.get("設置場所")) or None,
            "放電kW": kw,
            "電圧": ("特高" if kw and kw > 10000 else None),
            "変電所": nfkc(row.get("変電所")) or None,
            "負担金(円)": num(row.get("負担金(円)")),
            "税区分": nfkc(row.get("税区分")) or None,
            "工期(月)": num(row.get("工期(月)")),
            "上位系統増強": normalize_upper(row.get("上位系統増強")),
            "連系条件": cond or None,
            "有効期限": nfkc(row.get("有効期限")) or None,
            "読取信頼度": nfkc(row.get("読取信頼度")) or None,
            "未読理由": cond if state == "未読" else None,
            "備考": f"状態: {state or '不明'}（{nfkc(row.get('取込日')) or today} 追加・チャット読取CSVから追記）",
            "律速・可否メモ": ("連系不可" if re.search(r"^連系(否|不可)|連系否", cond)
                         else "運開希望日に間に合わず" if re.search(r"希望日には連系不可|間に合わ", cond) else None),
            "検証ステータス": "未確認",
            "原本リンク": "開く" if nfkc(row.get("原本リンク")) else None,
            "ファイル名": nfkc(row.get("ファイル名")) or None,
        }
        for h, v in values.items():
            if h in col:
                ws.cell(r, col[h]).value = v
        for h, fml in FORMULA_COLS.items():
            ws.cell(r, col[h]).value = fml.format(r=r)
        link = nfkc(row.get("原本リンク"))
        if link:
            ws.cell(r, col["原本リンク"]).hyperlink = link
        d = {h: ws.cell(r, col[h]).value for h in header}
        existing[r] = d
        added.append((r, d, row))

    # 集計対象の付け直し（追記した行と同じ案件だけ）
    touched = {case_key(d) for _, d, _ in added}
    groups = {}
    for r, d in existing.items():
        k = case_key(d)
        if k in touched:
            groups.setdefault(k, []).append((r, d))
    flips = []
    for k, items in groups.items():
        cands = [(date_key(d["回答日"]), r) for r, d in items
                 if d["回答種別"] in COUNT_TYPES and isinstance(d["負担金(円)"], (int, float))]
        winner = max(cands)[1] if cands else None
        for r, d in items:
            mark = "○" if r == winner else "－"
            if d["集計対象"] != mark:
                flips.append((int(float(d["No"])), d["集計対象"], mark))
            ws.cell(r, col["集計対象"]).value = mark
            d["集計対象"] = mark

    # 入力規則・条件付き書式・オートフィルタの範囲を広げる
    for dv in ws.data_validations.dataValidation:
        dv.sqref = openpyxl.worksheet.cell_range.MultiCellRange(
            " ".join(shift_range(str(rng), last) for rng in str(dv.sqref).split()))
    new_cf = ConditionalFormattingList()
    for cf in ws.conditional_formatting:
        rng = " ".join(shift_range(str(x), last) for x in str(cf.sqref).split())
        for rule in cf.rules:
            new_cf.add(rng, rule)
    ws.conditional_formatting = new_cf
    if ws.auto_filter.ref:
        ws.auto_filter.ref = shift_range(ws.auto_filter.ref, last)

    # 集計シートの参照範囲
    if SUMMARY_SHEET in wb.sheetnames:
        for row_cells in wb[SUMMARY_SHEET].iter_rows():
            for c in row_cells:
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.value = re.sub(r"\$(\d+)(?=[,)])", lambda m: f"${last}" if m.group(1) != "2" else m.group(0), c.value)

    # 要確認シート
    review_added = 0
    if REVIEW_SHEET in wb.sheetnames:
        wr = wb[REVIEW_SHEET]
        rlast = max(r for r in range(1, wr.max_row + 1) if wr.cell(r, 1).value is not None)
        for r, d, row in added:
            state = nfkc(row.get("状態"))
            reasons = []
            if state and state != "取込済み":
                reasons.append(f"状態: {state}")
            if d["律速・可否メモ"]:
                reasons.append(d["律速・可否メモ"])
            if d["回答種別"] in COUNT_TYPES and d["負担金(円)"] is None:
                reasons.append("負担金が読めていない")
            if d["回答種別"] in COUNT_TYPES and d["工期(月)"] is None:
                reasons.append("工期(月)が読めていない")
            if d["読取信頼度"] in ("medium", "low"):
                reasons.append(f"読取信頼度{d['読取信頼度']}")
            if not reasons:
                continue
            cond = nfkc(row.get("連系条件"))
            if cond and state != "取込済み":
                reasons.append(cond)
            rlast += 1
            for c in range(1, 6):
                src = wr.cell(rlast - 1, c)
                if src.has_style:
                    wr.cell(rlast, c)._style = copy.copy(src._style)
            wr.cell(rlast, 1).value = d["No"]
            wr.cell(rlast, 2).value = d["案件名"]
            wr.cell(rlast, 3).value = d["ファイル名"]
            wr.cell(rlast, 4).value = " / ".join(reasons)
            link = nfkc(row.get("原本リンク"))
            wr.cell(rlast, 5).value = "開く" if link else None
            if link:
                wr.cell(rlast, 5).hyperlink = link
            review_added += 1

    wb.save(args.out)
    recalced = False if args.no_recalc else recalc_with_soffice(args.out)

    print(f"追記 {len(added)}件（No {added[0][1]['No']}〜{added[-1][1]['No']}）、要確認シートに {review_added}件、"
          f"集計対象の変更 {len(flips)}件、数式の再計算: {'済' if recalced else '未'} → {args.out}")
    for no, before, after in sorted(flips):
        print(f"  No{no}: 集計対象 {before} → {after}")
    marks = sum(1 for _, d, _ in added if d["集計対象"] == "○")
    print(f"  追記分のうち集計対象○: {marks}件（実績DBに取り込まれるのはこの行）")


if __name__ == "__main__":
    main()
