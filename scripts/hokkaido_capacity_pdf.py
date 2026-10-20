#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
北海道電力ネットワークの「混雑状況マップ・予想潮流等一覧表」PDFから、変電所（変圧器）の
行を読み取る。update_capacity.py の北海道分（他社のCSVに相当）として使う。

北海道はCSVを公表しておらず、系統ごとのPDF（map_forecast_tide_list_01〜24.pdf と
187kV以上の map_forecast_tide_list.pdf）に「送電線」と「変電所」の2つの表が入っている。
変電所の表には他社の共通形式と同じ項目（一次・二次電圧、台数、設備容量、運用容量値、
予想潮流、空容量〔当該設備／上位系等考慮〕、N-1電制適用可否・可能量、平常時出力制御の可能性）
がある。配電用バンク（二次6.6kV）には当該・上位系の両方の空容量が載る。

読み取りは pdfplumber（単語と罫線の座標）で行う。レイアウトの要点:
  - 1行＝1バンク。複数バンクの変電所は変電所名のセルが縦に結合され、名前は結合セルの中央
    （バンク行の間や、3バンクなら真ん中のバンク行）に置かれる。名前は「どの罫線の間に
    あるか」で決める（縦方向の距離の近さでは3バンクの変電所と隣の変電所を取り違える）
  - 「送電線名」のヘッダー以降は送電線の表、「変電所名」以降は変電所の表（ページをまたぐ）
  - 非公開の設備は値が「―」で備考に「◇」

  pip install pdfplumber
  python scripts/hokkaido_capacity_pdf.py map_forecast_tide_list_13.pdf   # 読み取り結果をCSVで表示
"""
import csv
import re
import sys
import unicodedata

DASHES = {"", "-", "－", "ー", "―", "—", "‐", "−", "–"}
CONSTRAINTS = ("熱容量", "電圧", "安定度", "短絡", "同期", "周波数")
Y_TOL = 1.5   # 同じ行とみなす縦方向のずれ（pt。行の間隔は約6.8pt）


def _nfkc(s):
    return unicodedata.normalize("NFKC", s).strip()


def _num(tok):
    t = _nfkc(tok).replace(",", "")
    if t in DASHES or not re.fullmatch(r"-?\d+(\.\d+)?", t):
        return None
    v = float(t)
    return int(v) if v == int(v) else v


def _is_num_or_dash(tok):
    t = _nfkc(tok).replace(",", "")
    return t in DASHES or re.fullmatch(r"-?\d+(\.\d+)?", t) is not None


def _lines(words):
    """単語を中心のy座標で行にまとめる。[(y, [word(x順)])]"""
    out, cur, cur_y = [], [], None
    for w in sorted(words, key=lambda w: (w["top"] + w["bottom"]) / 2):
        y = (w["top"] + w["bottom"]) / 2
        if cur and abs(y - cur_y) > Y_TOL:
            out.append((cur_y, sorted(cur, key=lambda w: w["x0"])))
            cur = []
        if not cur:
            cur_y = y
        cur.append(w)
    if cur:
        out.append((cur_y, sorted(cur, key=lambda w: w["x0"])))
    return out


def _parse_bank(tokens):
    """バンク行なら値のdict、違えば None。tokens は行の文字列（x順）。"""
    ci = next((i for i, t in enumerate(tokens) if _nfkc(t) in CONSTRAINTS), None)
    if ci is None or ci < 5:
        return None
    head = tokens[:ci]
    if not all(_is_num_or_dash(t) for t in head[-5:]):
        return None
    v1, v2, units, cap, opcap = head[-5:]
    if _num(v1) is None or _num(v2) is None:
        return None            # 送電線の行（電圧・回線数の2つしか数字が無い）はここで落ちる
    rest = tokens[ci + 1:]
    # 予想潮流・当該・上位系 の3つ（数値か「―」）
    if len(rest) < 4 or not all(_is_num_or_dash(t) for t in rest[:3]):
        return None
    flow, cur, up = (_num(t) for t in rest[:3])
    i = 3
    n1 = _nfkc(rest[i])
    i += 1
    if n1 == "不可":
        # 「不可 ♯３」「不可 ♯４、♯５」のように理由番号が続く
        while i < len(rest) and re.match(r"^[♯#、]", _nfkc(rest[i])):
            n1 += " " + _nfkc(rest[i]).replace("♯", "#")
            i += 1
    elif n1 in DASHES:
        n1 = None
    elif n1 != "可":
        return None
    n1_amt = _num(rest[i]) if i < len(rest) else None
    i += 1
    curtail = _nfkc(rest[i]) if i < len(rest) else ""
    if curtail.startswith("有"):
        curtail = "有り"
    elif curtail in DASHES:
        curtail = "なし"
    else:
        curtail = None
    return {
        "v1": _num(v1), "v2": _num(v2), "units": _num(units), "cap": _num(cap), "opcap": _num(opcap),
        "flow": flow, "cur": cur, "up": up, "n1": n1, "n1_amount": n1_amt, "curtail": curtail,
        "has_up": True,
    }


def _separators(page, xm):
    """変電所名の列（x=xm）を横切る罫線のy座標（セルの上下の境界）。"""
    ys = sorted(e["top"] for e in page.edges if e["orientation"] == "h" and e["x0"] - 1 <= xm <= e["x1"] + 1)
    seps = []
    for y in ys:
        if not seps or y - seps[-1] > 1.0:
            seps.append(y)
    return seps


def _name_in_cell(words, y, seps, x_lo, x_hi):
    """バンク行(y)を含む名前セルの中にある文字を名前として返す。"""
    # 結合セルの途中でページが変わると上下どちらかの罫線が無い。その場合は行の前後10pt（約1.5行）だけ見る
    lo = max((s for s in seps if s <= y), default=y - 10)
    hi = min((s for s in seps if s > y), default=y + 10)
    parts = []
    for w in words:
        xc, yc = (w["x0"] + w["x1"]) / 2, (w["top"] + w["bottom"]) / 2
        if x_lo <= xc <= x_hi and lo < yc < hi:
            t = _nfkc(w["text"])
            if _is_num_or_dash(t) or t.startswith(("※", "◇")):
                continue
            parts.append((yc, w["x0"], t))
    return re.sub(r"^\d+", "", "".join(t for _, _, t in sorted(parts))).strip()


def parse_pdf(path):
    """(更新日の文字列, バンク行のリスト) を返す。各行は update_capacity.parse_csv と同じキーを持つ。"""
    import pdfplumber

    date, mode, out, dropped = "", None, [], 0
    name_x = None   # 変電所名の列のx範囲（ヘッダーから決め、続きのページにも使う）
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            seps = None
            for y, ws in _lines(words):
                toks = [w["text"] for w in ws]
                joined = "".join(_nfkc(t) for t in toks)
                m = re.search(r"更新日[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)", joined)
                if m:
                    date = m.group(1)
                if "送電線名" in joined:
                    mode = "line"
                    continue
                hdr = next((w for w in ws if _nfkc(w["text"]).startswith("変電所名")), None)
                if hdr:
                    mode = "sub"
                    # 名前の列：「No」列の右端〜「一次」(電圧)列の左端。ヘッダー語の位置から決める
                    no = next((w for w in ws if _nfkc(w["text"]) == "変電所"), None)
                    v1 = next((w for w in words if _nfkc(w["text"]) == "一次" and w["top"] > hdr["top"] - 2), None)
                    x_lo = (no["x1"] if no else hdr["x0"] - 25) + 1
                    x_hi = (v1["x0"] if v1 else hdr["x1"] + 30) - 1
                    name_x = (x_lo, x_hi, (hdr["x0"] + hdr["x1"]) / 2)
                    continue
                if mode != "sub" or name_x is None:
                    continue
                row = _parse_bank(toks)
                if not row:
                    continue
                if seps is None:
                    seps = _separators(page, name_x[2])
                name = _name_in_cell(words, y, seps, name_x[0], name_x[1])
                if not name:
                    dropped += 1     # 名前が決められないバンクは捨てる（推測で付けない）
                    continue
                row["name"] = name
                out.append(row)
    if dropped:
        print(f"{path}: 名前を決められなかったバンク {dropped}行", file=sys.stderr)
    return date, out


def main():
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    w = csv.writer(sys.stdout)
    w.writerow(["file", "date", "name", "v1", "v2", "units", "cap", "opcap", "flow", "cur", "up",
                "n1", "n1_amount", "curtail"])
    for p in sys.argv[1:]:
        date, rows = parse_pdf(p)
        for r in rows:
            w.writerow([p, date, r["name"], r["v1"], r["v2"], r["units"], r["cap"], r["opcap"], r["flow"],
                        r["cur"], r["up"], r["n1"], r["n1_amount"], r["curtail"]])


if __name__ == "__main__":
    main()
