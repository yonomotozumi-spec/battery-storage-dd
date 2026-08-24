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
          "s1", "s2", "s3", "s4", "s5", "s6", "grid", "rank"]


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

    d_area, d_acc, d_n1, d_curtail, d_rank = Dict_(), Dict_(), Dict_(), Dict_(), Dict_()
    rows = []
    with_coords = 0
    for (no, area, name, pref, vmax, vsec, lat, lon, acc, cap, opcap, flow,
         avail, avail_up, n1, curtail, addr) in src_rows:
        sc = S.score_row(area, vsec, opcap, flow, avail, avail_up, n1, curtail, name=name)
        la, lo = num(lat, 5), num(lon, 5)
        if la is not None and lo is not None:
            with_coords += 1
        rows.append([
            num(no), d_area.idx(area), name or "", pref or "", la, lo, d_acc.idx(acc),
            num(vmax), num(vsec), num(opcap), num(flow), num(avail), num(avail_up),
            d_n1.idx(n1), d_curtail.idx(curtail), addr or "",
            sc["s1"], sc["s2"], sc["s3"], sc["s4"], sc["s5"], sc["s6"],
            sc["grid_score"], d_rank.idx(sc["rank"]),
        ])

    payload = {
        "meta": {
            "generated": datetime.date.today().isoformat(),
            "source": args.source_label,
            "total": len(rows),
            "withCoords": with_coords,
            "gridMax": S.GRID_MAX,
            "s7Max": S.S7_MAX,
            "totalMax": S.TOTAL_MAX,
        },
        "dict": {
            "area": d_area.values, "acc": d_acc.values,
            "n1": d_n1.values, "curtail": d_curtail.values, "rank": d_rank.values,
        },
        "fields": FIELDS,
        "params": {
            "s7Steps": [list(s) for s in S.S7_STEPS],
            "totalRankSteps": [list(s) for s in S.TOTAL_RANK_STEPS],
            "gridRank": {"S": S.P["rank_s"]["thr"], "A": S.P["rank_a"]["thr"], "B": S.P["rank_b"]["thr"]},
            # Web版が「なぜその点数になったか」を出すための閾値。
            # ここを渡さないと画面側に配点を書き写すことになり、
            # substation_scoring.py を直したときに食い違う。
            "thr": {
                "s1": [[S.P["s1_t1"]["thr"], S.P["s1_t1"]["pts"]],
                       [S.P["s1_t2"]["thr"], S.P["s1_t2"]["pts"]],
                       [0, S.P["s1_t3"]["pts"]]],
                "s2": S.P["s2"]["pts"],
                "s3": [[S.P["s3_t1"]["thr"], S.P["s3_t1"]["pts"]],
                       [S.P["s3_t2"]["thr"], S.P["s3_t2"]["pts"]],
                       [S.P["s3_t3"]["thr"], S.P["s3_t3"]["pts"]]],
                "s4": S.P["s4"]["pts"],
                "s5": [S.P["s5_1"]["thr"], S.P["s5_1"]["pts"], S.P["s5_2"]["pts"]],
            "s6Actual": [list(x) for x in S.S6_ACTUAL_STEPS],
            },
        },
        "areaPoints": [list(a) for a in S.AREA_POINTS],
        "scoreDefs": [list(s) for s in S.SCORE_DEFS],
        "results": [list(r) for r in S.RESULTS_DB],
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
    print(f"出力: {args.out}  {len(rows)}件（座標あり {with_coords}件） {size_kb:.0f}KB")


if __name__ == "__main__":
    main()
