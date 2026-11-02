#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
変電所スコアリングの配点定義（系統用蓄電池・高圧2MW前提 v1 / 2026-07-28）。

配点・閾値・エリア点・実績DBの単一ソース。
xlsx生成（score_substations.py）とWeb用データ生成（build_substation_web_data.py）が
ともにこのモジュールを参照するため、定義はここだけを直せば両方に反映される。

スコア構成:
  系統スコア80点 = S1空容量35 + S2 N-1電制10 + S3潮流15 + S4出力制御5 + S5配変10 + S6エリア5
  総合スコア100点 = 系統スコア80 + S7近接20（土地の座標が決まってから加算）
最低必要容量: 空容量が MIN_MW（2MW＝高圧の案件規模）未満は「空き無し」と同じ扱い（S1=0点・ゲート・地図の色）
ゲート: 空容量(上位系考慮)＜2MW かつ N-1電制≠可 → 警告（ランクは通常どおり算出。S1=0なので最高45点=B止まり）
  2026-10-20までは「除外」としてランク計算から外していた。上位系増強で長期化した例が根拠だったが、
  上位系0・N-1不可でも短工期で連系した実績があり、上位系0はノンファーム接続の混雑を
  示すだけで増強必須とは限らないため、隠さずに警告を付けて見せる方式に変えた。
"""

import json
import os

MIN_MW = 2       # 最低必要容量（高圧2MWの案件規模）。これ未満の空容量は「空き無し」と同じ扱い（2026-10-20 余野さん判断）
GRID_MAX = 80    # 系統スコア満点
S7_MAX = 20      # 近接スコア満点
TOTAL_MAX = 100  # 総合スコア満点

# ── 配点パラメータ（key, ラベル, 閾値, 点数）──
PARAM_ROWS = [
    ("s1_t1", "S1: 空容量 ≧(MW) で35点", 20, 35),
    ("s1_t2", "S1: 空容量 ≧(MW) で28点", 6, 28),
    ("s1_t3", "S1: 空容量 ≧(MW) で15点（最低必要容量）", MIN_MW, 15),
    ("s2", "S2: N-1電制「可」", None, 10),
    ("s3_t1", "S3: |潮流|/運用容量 ≦ で15点", 0.3, 15),
    ("s3_t2", "S3: |潮流|/運用容量 ≦ で9点", 0.6, 9),
    ("s3_t3", "S3: |潮流|/運用容量 ≦ で4点", 0.9, 4),
    ("s4", "S4: 出力制御「なし」", None, 5),
    ("s5_1", "S5: 二次電圧 ≦(kV) で10点", 7, 10),
    ("s5_2", "S5: 二次電圧 その他数値", None, 5),
    ("rank_s", "ランクS: 系統スコア ≧", 60, None),
    ("rank_a", "ランクA: 系統スコア ≧", 48, None),
    ("rank_b", "ランクB: 系統スコア ≧", 36, None),
]
P = {key: {"thr": thr, "pts": pts} for key, _, thr, pts in PARAM_ROWS}

# ── S7 近接（土地確定後）。(距離m以下, 点数) を近い順に。超過は0点 ──
S7_STEPS = [(500, 20), (1000, 15), (2000, 10), (5000, 4)]

# ── 総合スコア(100点)のランク閾値 ──
# 参考シートは系統スコア80点でのランク(S:60/A:48/B:36)しか定義していないため、
# 同じ比率(75%/60%/45%)を100点満点に換算した本ツール独自の閾値。系統ランクとは別物として表示すること。
TOTAL_RANK_STEPS = [
    (round(P["rank_s"]["thr"] / GRID_MAX * TOTAL_MAX), "S"),
    (round(P["rank_a"]["thr"] / GRID_MAX * TOTAL_MAX), "A"),
    (round(P["rank_b"]["thr"] / GRID_MAX * TOTAL_MAX), "B"),
]

# S6 エリア点（実績工期・出力制御環境からの仮置き。実績追加で更新）
AREA_POINTS = [
    ("北海道", 5, "計量のみで短工期の例あり"),
    ("東北", 4, "配電工事のみなら短工期の例あり"),
    ("東京", 4, "需要地・配変多い"),
    ("中部", 2, "増強系の工期が長い傾向"),
    ("北陸", 3, "実績なし・中位仮置き"),
    ("関西", 3, "工期は中位"),
    ("中国", 3, "中位仮置き"),
    ("四国", 3, "実績なし・中位仮置き"),
    ("九州", 5, "短工期の例あり・出力制御は多め"),
    ("沖縄", 2, "市場環境考慮・仮置き"),
    ("J-POWER", 0, "対象外"),
]
AREA_POINT_MAP = {area: pts for area, pts, _ in AREA_POINTS}

# 接続検討回答の実績は案件情報なので、公開リポジトリには置かない（読み込みは下の load 処理）
RESULTS_DB = []

# 分析メモ（2026-08-06追記・19件ベース）
# 実績に基づく分析メモは社内版のみ（公開版は空）
ANALYSIS_NOTES = []

# 変数・配点定義（実績12件で校正）— 基準シートの表示用
SCORE_DEFS = [
    ("G1", "ゲート: 上位系0（警告）", "警告", "空容量(上位系考慮)＜2MW かつ N-1電制≠可 → ⚠警告（ランクは通常算出。S1=0でB以下）", "上位系増強で長期化する例と、上位系0でも短工期で連系した例の両方がある（2026-10-20に除外→警告）"),
    ("G2", "ゲート: 募集プロセス/増強エリア", "除外(手動)", "各社公表資料で該当時に手動除外", "高額・長納期化リスク"),
    ("S1", "空容量(上位系考慮)", "35", "≧20MW(2MWの10倍):35 / ≧6MW:28 / ≧2MW:15 / ＜2MW:0（最低必要容量2MW未満は空き無し扱い）", "最重要。上位系増強の回避確率"),
    ("S2", "N-1電制 可否", "10", "可:10 / それ以外:0", "空容量ゼロでも接続余地"),
    ("S3", "予想潮流の余裕(|潮流|/運用容量)", "15", "≦30%:15 / ≦60%:9 / ≦90%:4 / >90%:0", "混雑度=出力制御・増強リスク"),
    ("S4", "出力制御可能性", "5", "なし:5 / 有り:0", "収益影響(九州で頻発)"),
    ("S5", "配電用変電所か(二次電圧)", "10", "二次≦7kV(6.6/6.0):10 / その他数値:5 / 不明:0", "高圧2MW連系の受け皿"),
    ("S6", "エリア", "5", "エリア点表参照", "実績工期のエリア差"),
    ("合計", "系統スコア", "80", "S1+S2+S3+S4+S5+S6", ""),
    ("S7", "(土地確定後)変電所・既設配電線への近接", "20", "≦500m:20 / ≦1km:15 / ≦2km:10 / ≦5km:4 / 超:0", "負担金≒400万+0.8〜1.0万円/m×新設延長"),
    ("", "実績補正", "±", "実績DBに同一変電所があれば優先参照(実績メモ列)", "実績は社内版のみ"),
]

NOTES = [
    "スコアリング v1 2026-07-28（接続検討の実回答12件の分析から逆算した配点）に準拠。",
    "使い方: ランク S/A の変電所を優先し、その周辺（目安2km以内・既設高圧配電線の近く）で土地を探す→接続検討を申込む。",
    "系統スコアは80点満点。土地確定後に近接性S7(20点)を加えて100点満点で最終評価する。",
    "負担金の事前推定式(実績回帰・2026-08-06更新): 推定負担金(万円) ≒ 固定費200〜400 + 高圧線新設・張替延長(m)×0.7〜1.0 + SVC/SVR台数×500〜 + バンク逆潮流対策500〜1,000。",
    "工期目安: 九州は計量・張替のみ3-6ヶ月 / 中部は最軽微工事でも18ヶ月・増強系18-26ヶ月 / SVC込16-17ヶ月 / 特高・上位系増強は2〜10年。",
    "接続検討→契約申込で +10〜15%・+半年 の増額リスクあり(2026-08-06更新)。",
    "公表空容量は特高側評価。高圧連系の最終確定は接続検討のみ。バンク・配電線の空きは実績DBタブで補正すること。",
    "【前提】S1・ゲートの空容量は「上位系考慮」を優先し、単一値のみ公表のエリア(北陸・中部・沖縄等)は「当該設備」で代用する。",
    "【前提】ランク「データ無」= 空容量が両列とも未公表(OSMのみ等)。ゲート該当行（上位系0かつN-1不可）は「⚠上位系0」の警告を付けて通常どおりランク付けする（2026-10-20までは「除外」として非表示）。",
    "【前提】最低必要容量は2MW（高圧の案件規模）。空容量が2MW未満の変電所は「空き無し」と同じ扱い（S1=0点・ゲート判定・地図の色）。2026-10-20に「0以下」から変更。",
    "【前提】S4は出力制御が「なし」(完全一致)の場合のみ加点。「有り」「あり」表記ゆれ・不明は0点。",
    "【参考】外部DB: 蓄電所ネット「系統連系診断」(https://bess-net.jp/tools/grid-connection-check) は10社8,346変電所の公表データで連系候補Top5を出す無料ツール(2026-10-12更新)。"
    "空容量は当該設備ベースで上位系のゲートが無いため本ツールより甘く出る(例: 旭川中央は上位系考慮0MW・N-1不可でも55点)。クロスチェック用に使い、採否は本ツールの上位系考慮と接続検討で判断する。",
]


def _num(v):
    """数値なら float、それ以外（None・文字列）は None。"""
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def effective_capacity(avail_cur, avail_up):
    """S1・ゲートで使う空容量。「上位系考慮」を優先し、無ければ「当該設備」で代用。

    どちらも未公表なら None（= 空容量データ無し）。
    """
    up, cur = _num(avail_up), _num(avail_cur)
    return up if up is not None else cur


def score_row(area, vsec, opcap, flow, avail_cur, avail_up, n1, curtail):
    """1変電所の系統スコアを算出する。xlsxの数式と同じ判定。

    戻り値: {"gate", "s1".."s6", "grid_score", "rank"}
    """
    eff = effective_capacity(avail_cur, avail_up)
    gate = "警告" if (eff is not None and eff < MIN_MW and n1 != "可") else ""

    if eff is None:
        s1 = 0
    elif eff >= P["s1_t1"]["thr"]:
        s1 = P["s1_t1"]["pts"]
    elif eff >= P["s1_t2"]["thr"]:
        s1 = P["s1_t2"]["pts"]
    elif eff >= P["s1_t3"]["thr"]:
        s1 = P["s1_t3"]["pts"]
    else:
        s1 = 0

    s2 = P["s2"]["pts"] if n1 == "可" else 0

    cap, fl = _num(opcap), _num(flow)
    if cap is None or fl is None or cap == 0:
        s3 = 0
    else:
        ratio = abs(fl) / cap
        if ratio <= P["s3_t1"]["thr"]:
            s3 = P["s3_t1"]["pts"]
        elif ratio <= P["s3_t2"]["thr"]:
            s3 = P["s3_t2"]["pts"]
        elif ratio <= P["s3_t3"]["thr"]:
            s3 = P["s3_t3"]["pts"]
        else:
            s3 = 0

    s4 = P["s4"]["pts"] if curtail == "なし" else 0

    v = _num(vsec)
    s5 = 0 if v is None else (P["s5_1"]["pts"] if v <= P["s5_1"]["thr"] else P["s5_2"]["pts"])

    s6 = AREA_POINT_MAP.get(area, 0)

    grid = s1 + s2 + s3 + s4 + s5 + s6
    if eff is None:
        rank = "データ無"
    elif grid >= P["rank_s"]["thr"]:
        rank = "S"
    elif grid >= P["rank_a"]["thr"]:
        rank = "A"
    elif grid >= P["rank_b"]["thr"]:
        rank = "B"
    else:
        rank = "C"

    return {"gate": gate, "s1": s1, "s2": s2, "s3": s3, "s4": s4, "s5": s5, "s6": s6,
            "grid_score": grid, "rank": rank}


def s7_points(distance_m):
    """S7 近接スコア（20点満点）。距離は変電所までの直線距離(m)。"""
    if distance_m is None:
        return 0
    for limit, pts in S7_STEPS:
        if distance_m <= limit:
            return pts
    return 0


def total_rank(total_score):
    """総合スコア(100点)のランク。系統ランクを100点満点に比例換算した本ツール独自の区分。"""
    for thr, rank in TOTAL_RANK_STEPS:
        if total_score >= thr:
            return rank
    return "C"


def results_key(area, name):
    """実績DBの引き当てキー（エリア＋変電所名）。"""
    return f"{area}{(name or '').strip()}"


# 社内版のビルド時だけ、公開リポジトリの外に置いた実績JSONを環境変数で渡す
# （scripts/build_internal_map.py が使う）。公開版のビルドでは実績は常に0件。
RESULTS_MANUAL_COUNT = 0
_private_path = os.environ.get("BESS_PRIVATE_RESULTS")
if _private_path:
    with open(_private_path, encoding="utf-8") as _f:
        _private = json.load(_f)
    RESULTS_DB = [tuple(r) for r in _private["records"]]
    RESULTS_MANUAL_COUNT = _private.get("manualCount", 0)
    ANALYSIS_NOTES = _private.get("analysisNotes", ANALYSIS_NOTES)
    NOTES = _private.get("notes", NOTES)
    SCORE_DEFS = [tuple(x) for x in _private.get("scoreDefs", SCORE_DEFS)]
    AREA_POINTS = [tuple(x) for x in _private.get("areaPoints", AREA_POINTS)]

# 同じ変電所に複数の回答があるときは要約を「／」でつなぐ
RESULTS_MEMO = {}
for _r in RESULTS_DB:
    _k = results_key(_r[0], _r[1])
    RESULTS_MEMO[_k] = f"{RESULTS_MEMO[_k]} ／ {_r[8]}" if _k in RESULTS_MEMO else _r[8]
