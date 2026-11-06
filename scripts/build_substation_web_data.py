#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Web版（substation.html）用の変電所データ生成スクリプト。

全国変電所リスト（66kV以上）のxlsxを読み込み、substation_scoring.py の配点で
全変電所の系統スコア(80点)を算出し、data/substations.js を出力する。
S7近接(20点)はブラウザ側で入力座標からの距離に応じて加算される。

xlsx版（score_substations.py）と同じ substation_scoring.py を参照するため、
配点定義を変えれば両方に同じように反映される。

出力は `<script src>` で読める素のJSファイル（window.SUBSTATION_DATA に代入）。
fetch()を使わないため、file:// でHTMLを直接開いても動作する。

使い方:
  python scripts/build_substation_web_data.py \
    --in 全国変電所リスト_66kV以上.xlsx --out data/substations.js
"""
import argparse
import datetime
import json
import os

from score_substations import read_substations
import substation_scoring as S

# rows の並び。substation.html の F.* と対応する
FIELDS = ["no", "area", "name", "pref", "lat", "lon", "acc", "vmax", "vsec",
          "opcap", "flow", "availCur", "availUp", "n1", "curtail", "addr",
          "s1", "s2", "s3", "s4", "s5", "s6", "grid", "rank", "muni", "gpref", "gate",
          # v2: S1の内訳（S1a 変電所単体＝当該設備・S1b 上位系）と接続検討の判定（dict.judge の添字。実績による格下げを反映した最終判定）
          "s1a", "s1b", "judge",
          # 実績による格下げ（2026-11-06）: 公表データだけの判定（dict.judge の添字）、実績NGの理由、上位系の実績警告の理由（無ければ ""）
          "judgePub", "resultNg", "upperWarn"]

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REVGEO_CACHE = os.path.join(ROOT, "geo_cache", "revgeo.json")
MUNI_CACHE = os.path.join(ROOT, "geo_cache", "muni_codes.json")


def load_geo():
    """座標→(都道府県, 市区町村) を引く関数を返す。

    逆ジオコードの結果は build_priority_munis.py --all が geo_cache/ に貯めたものを使い、
    ここではネットワークに出ない。キャッシュに無い座標は空欄になる。
    """
    try:
        with open(REVGEO_CACHE, encoding="utf-8") as f:
            cache = json.load(f)
        with open(MUNI_CACHE, encoding="utf-8") as f:
            table = json.load(f)
    except OSError:
        return lambda la, lo: ("", "")

    def lookup(la, lo):
        if la is None or lo is None:
            return "", ""
        code = cache.get(f"{la:.4f},{lo:.4f}")
        code = code.lstrip("0") if code else None
        pm = table.get(code) if code else None
        return (pm[0], pm[1]) if pm else ("", "")
    return lookup


class Dict_:
    """繰り返し出現する文字列を添字に置き換えて出力サイズを抑える。"""

    def __init__(self):
        self.values = []
        self._idx = {}

    def idx(self, v):
        v = "" if v is None else str(v)
        if v not in self._idx:
            self._idx[v] = len(self.values)
            self.values.append(v)
        return self._idx[v]


def num(v, digits=None):
    """数値ならJSON向けに整形（整数値はintに落とす）。それ以外は None。"""
    if not isinstance(v, (int, float)) or isinstance(v, bool):
        return None
    v = round(float(v), digits) if digits is not None else float(v)
    return int(v) if v == int(v) else v


def main():
    ap = argparse.ArgumentParser(description="Web版(substation.html)用の変電所データJSを生成")
    ap.add_argument("--in", dest="src", required=True, help="全国変電所リスト（66kV以上）のxlsxパス")
    ap.add_argument("--out", dest="out", required=True, help="出力JSパス（例: data/substations.js）")
    ap.add_argument("--list-sheet", default="全国変電所リスト", help="入力の変電所リストのシート名")
    ap.add_argument("--source-label", default="全国変電所リスト（66kV以上・系統接続検討用）",
                    help="画面に表示する元データ名")
    args = ap.parse_args()

    src_rows = read_substations(args.src, args.list_sheet)

    d_area, d_acc, d_n1, d_curtail, d_rank, d_judge = Dict_(), Dict_(), Dict_(), Dict_(), Dict_(), Dict_()
    rows = []
    with_coords = 0
    geo = load_geo()
    with_muni = 0
    adj = S.result_adjuster()
    n_result_ng = n_upper_warn = 0
    for (no, area, name, pref, vmax, vsec, lat, lon, acc, cap, opcap, flow,
         avail, avail_up, n1, curtail, addr) in src_rows:
        sc = S.score_row(area, vsec, opcap, flow, avail, avail_up, n1, curtail)
        a = adj.adjust(area, name, no, sc["judge"])   # 実績による格下げ（実績DBが無ければ何もしない）
        n_result_ng += 1 if a["resultNg"] else 0
        n_upper_warn += 1 if a["upperWarn"] else 0
        la, lo = num(lat, 5), num(lon, 5)
        if la is not None and lo is not None:
            with_coords += 1
        gpref, muni = geo(la, lo)
        if muni:
            with_muni += 1
        rows.append([
            num(no), d_area.idx(area), name or "", pref or "", la, lo, d_acc.idx(acc),
            num(vmax), num(vsec), num(opcap), num(flow), num(avail), num(avail_up),
            d_n1.idx(n1), d_curtail.idx(curtail), addr or "",
            sc["s1"], sc["s2"], sc["s3"], sc["s4"], sc["s5"], sc["s6"],
            sc["grid_score"], d_rank.idx(sc["rank"]),
            # 市区町村と、座標から引いた都道府県（元リストの都道府県が空の行の補完用）
            muni, gpref,
            # ゲート（上位系0かつN-1不可）の警告。1=警告あり
            1 if sc["gate"] else 0,
            sc["s1a"], sc["s1b"], d_judge.idx(a["judge"]),
            d_judge.idx(sc["judge"]), a["resultNg"], a["upperWarn"],
        ])

    payload = {
        "meta": {
            "generated": datetime.date.today().isoformat(),
            "source": args.source_label,
            "total": len(rows),
            "withCoords": with_coords,
            "withMuni": with_muni,
            "gridMax": S.GRID_MAX,
            "s7Max": S.S7_MAX,
            "totalMax": S.TOTAL_MAX,
            # 実績による格下げの基準日（この日から RESULT_NG_WINDOW_DAYS 以内の回答を見る）と件数
            "resultAsOf": adj.as_of.isoformat(),
            "resultNgCount": n_result_ng,
            "upperWarnCount": n_upper_warn,
        },
        "dict": {
            "area": d_area.values, "acc": d_acc.values,
            "n1": d_n1.values, "curtail": d_curtail.values, "rank": d_rank.values, "judge": d_judge.values,
        },
        "fields": FIELDS,
        "params": {
            "s7Steps": [list(s) for s in S.S7_STEPS],
            "totalRankSteps": [list(s) for s in S.TOTAL_RANK_STEPS],
            "gridRank": {"S": S.P["rank_s"]["thr"], "A": S.P["rank_a"]["thr"], "B": S.P["rank_b"]["thr"]},
            # S1 空容量の段階 [(MW以上, 点)]。画面の「要接続検討」（ゲート除外だが当該設備に空き）の閾値にも使う
            "s1Steps": [[S.P["s1_t1"]["thr"], S.P["s1_t1"]["pts"]], [S.P["s1_t2"]["thr"], S.P["s1_t2"]["pts"]]],
            # 最低必要容量（MW）。これ未満の空容量は画面でも「空き無し」と判定する
            "minMw": S.MIN_MW,
            # v2: S1a（変電所単体）の段階と S1b（上位系）の点、判定のラベル・意味、候補地の判定に使う距離（m）
            "s1aSteps": [[S.P[k]["thr"], S.P[k]["pts"]] for k in ("s1_t1", "s1_t2", "s1_t3")],
            "s1bPoints": {"ok": S.P["s1b_ok"]["pts"], "unk": S.P["s1b_unk"]["pts"], "bad": 0},
            "judgeLabels": {"ok": S.JUDGE_OK, "check": S.JUDGE_CHECK, "ng": S.JUDGE_NG, "nodata": S.JUDGE_NODATA},
            "judgeMeaning": S.JUDGE_MEANING,
            "siteOkM": S.SITE_OK_M, "siteRangeM": S.SITE_RANGE_M,
            "s5MaxKv": S.P["s5_1"]["thr"],
            # 実績による格下げの閾値とラベル（画面の凡例・理由の文言に使う）
            "resultNg": {"yenMan": S.RESULT_NG_YEN_MAN, "months": S.RESULT_NG_MONTHS, "windowDays": S.RESULT_NG_WINDOW_DAYS,
                         "tag": S.RESULT_NG_TAG, "upperTag": S.UPPER_WARN_TAG, "meaning": S.RESULT_NG_MEANING},
        },
        "areaPoints": [list(a) for a in S.AREA_POINTS],
        "scoreDefs": [list(s) for s in S.SCORE_DEFS],
        "results": [list(r) for r in S.RESULTS_DB],
        "resultsManual": S.RESULTS_MANUAL_COUNT,
        "notes": S.NOTES,
        "analysisNotes": S.ANALYSIS_NOTES,
        "rows": rows,
    }

    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    header = ("// 自動生成ファイル - 直接編集しないこと\n"
              "// 生成: scripts/build_substation_web_data.py\n"
              f"// 元データ: {args.source_label}\n")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(header + "window.SUBSTATION_DATA=" + body + ";\n")

    size_kb = os.path.getsize(args.out) / 1024
    print(f"出力: {args.out}  {len(rows)}件（座標あり {with_coords}件・市区町村あり {with_muni}件） {size_kb:.0f}KB")
    print(f"実績による格下げ: 実績NG {n_result_ng}件・⚠上位系に実績NG {n_upper_warn}件（基準日 {adj.as_of}・実績NGの回答 {len(adj.ngs)}件）")


if __name__ == "__main__":
    main()
