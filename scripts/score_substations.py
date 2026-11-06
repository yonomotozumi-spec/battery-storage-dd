#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
変電所スコアリング xlsx 生成スクリプト（系統用蓄電池・高圧2MW前提 v1）

全国変電所リスト（66kV以上・系統接続検討用）xlsx を読み込み、
接続検討の実績から校正した配点（変電所スコアリング v2 2026-11-06。v1は2026-07-28）で
全変電所を「接続検討を申し込む価値」の順に点数化した3タブのxlsxを出力する。

  スコアリング     : 全行のゲート判定・接続検討の判定（OK／要検討／NG）・S1a・S1b・S2〜S6・系統スコア(80点)・ランク・実績メモ。
                     スコア列は全て数式で、入力値や基準パラメータを変えると再計算される。
                     ランクS/Aの行は条件付き書式で黄色ハイライト。
  スコアリング基準 : 配点定義・エリア点・ランク閾値のパラメータ表（数式の参照元）と前提。
  実績DB           : 接続検討回答の実績19件(2026-08-06更新)と分析メモ。
                     エリア＋変電所名の一致で実績メモ列に自動表示。

スコア構成（系統スコア80点。土地確定後に近接性S7=20点を加えて100点満点）:
  S1a 変電所単体(当該設備)の空容量 25 / S1b 上位系の空き 10 / S2 N-1電制 10 / S3 潮流余裕 15 / S4 出力制御 5 / S5 配変 10 / S6 エリア 5
ゲート: 空容量(上位系考慮)＜2MW かつ N-1電制≠可 → ⚠警告（ランクは通常算出。2026-10-20までは「除外」）
判定: 変電所単体＜2MW＝接続検討NG ／ 単体は空きありだが上位系空き無し・二次電圧7kV超＝要検討 ／ それ以外＝接続検討OK ／ 空容量未公表＝データ無
実績による格下げ（2026-11-06〜）: 直近1年以内の接続検討の回答が 連系不可／負担金4,000万円以上／工期24ヶ月以上 なら判定NG（AE列）。
  実績NGが名指しした上位系の設備を制約設備に持つ変電所は「⚠上位系に実績NG」（AF列。OK→要検討）。公表データだけの判定はAD列
ランク: S:60点以上 / A:48-59 / B:36-47 / C:35以下 / データ無:空容量未公表(OSMのみ等)

配点・閾値・エリア点・実績DBの定義は substation_scoring.py に集約している
（Web版 substation.html 用データもそこを参照するため、定義の変更は同モジュールだけで済む）。

使い方:
  python scripts/score_substations.py --in 全国変電所リスト.xlsx --out 変電所スコアリング.xlsx
"""
import argparse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

from substation_scoring import (
    AREA_POINTS, ANALYSIS_NOTES, JUDGE_CHECK, JUDGE_MEANING, JUDGE_NG, JUDGE_NODATA, JUDGE_OK, NOTES, PARAM_ROWS,
    RESULTS_DB, RESULTS_MANUAL_COUNT, RESULTS_MEMO, SCORE_DEFS, results_key,
    RESULT_NG_MEANING, RESULT_NG_TAG, UPPER_WARN_TAG, judge_row, result_adjuster,
)

HEADER_FILL = PatternFill("solid", fgColor="FCE4D6")
PARAM_FILL = PatternFill("solid", fgColor="FFFF00")   # 変更可能なパラメータセル
RANK_FILL = PatternFill("solid", fgColor="FFF2A8")    # ランクS/Aハイライト（黄色）
GATE_FILL = PatternFill("solid", fgColor="FFE0CC")    # ゲート該当（上位系0の警告）行（薄い橙）
RESULT_NG_FILL = PatternFill("solid", fgColor="F8CBAD")  # 実績NG（直近1年の接続検討で連系不可・高額・長工期 → 判定NG）行（薄い赤）
HDR_FONT = Font(bold=True, size=11)
TITLE_FONT = Font(bold=True, size=14)
SMALL = Font(size=10)
THIN = Side(style="thin", color="D0D0D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP = Alignment(wrap_text=True, vertical="center")

BASE = "スコアリング基準"



def build_base_sheet(ws):
    """基準シートを作成し、数式が参照するパラメータセルの絶対参照を dict で返す。"""
    ws["A1"] = "変電所スコアリング基準（系統用蓄電池・高圧2MW前提）v2"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("v2: 2026-11-06 ｜ S1空容量35点を S1a 変電所単体（当該設備）25点＋S1b 上位系の空き10点に分け、接続検討の判定（OK／要検討／NG）を追加。"
                "高圧2MWは上位系が0でも配電用バンクに空きがあれば連系できることが多い（実績32件中31件）ため ｜ v1: 2026-07-28 接続検討の実回答12件の分析から逆算した配点")
    ws["A2"].font = SMALL

    r = 4
    ws.cell(r, 1, "変数・配点定義（v1は実績12件で校正、v2で S1 を当該設備と上位系に分割）").font = HDR_FONT
    r += 1
    for c, h in enumerate(["記号", "変数", "配点", "閾値・ルール", "根拠(実績)"], 1):
        cell = ws.cell(r, c, h)
        cell.fill = HEADER_FILL; cell.font = HDR_FONT; cell.alignment = CENTER; cell.border = BORDER
    for row in SCORE_DEFS:
        r += 1
        for c, v in enumerate(row, 1):
            if c == 9:  # スコア表は最初の一致行しか引けないため、同じ変電所の要約をまとめて載せる
                v = RESULTS_MEMO[results_key(row[0], row[1])]
            cell = ws.cell(r, c, v); cell.border = BORDER; cell.alignment = WRAP

    # 配点パラメータ（数式の参照元。黄色セルを変えるとスコアリングタブが再計算される）
    r += 2
    ws.cell(r, 1, "配点パラメータ（黄色セルを変更するとスコアが自動再計算されます）").font = HDR_FONT
    r += 1
    for c, h in enumerate(["項目", "閾値", "点数"], 1):
        cell = ws.cell(r, c, h)
        cell.fill = HEADER_FILL; cell.font = HDR_FONT; cell.alignment = CENTER; cell.border = BORDER
    params = {}
    for key, label, thr, pts in PARAM_ROWS:
        r += 1
        ws.cell(r, 1, label).border = BORDER
        tc = ws.cell(r, 2, thr); pc = ws.cell(r, 3, pts)
        for cell in (tc, pc):
            cell.border = BORDER; cell.alignment = CENTER
            if cell.value is not None:
                cell.fill = PARAM_FILL
        params[key] = {"thr": f"{BASE}!$B${r}", "pts": f"{BASE}!$C${r}"}

    # S6 エリア点表（VLOOKUP参照元）
    r += 2
    ws.cell(r, 1, "S6 エリア点（実績工期・出力制御環境からの仮置き。実績追加で更新）").font = HDR_FONT
    r += 1
    for c, h in enumerate(["エリア", "点数", "根拠(実績)"], 1):
        cell = ws.cell(r, c, h)
        cell.fill = HEADER_FILL; cell.font = HDR_FONT; cell.alignment = CENTER; cell.border = BORDER
    area_first = r + 1
    for area, pts, why in AREA_POINTS:
        r += 1
        ws.cell(r, 1, area).border = BORDER
        pc = ws.cell(r, 2, pts); pc.border = BORDER; pc.alignment = CENTER; pc.fill = PARAM_FILL
        ws.cell(r, 3, why).border = BORDER
    params["area_range"] = f"{BASE}!$A${area_first}:$B${r}"

    # ランク集計（スコアリングタブに連動）
    r += 2
    ws.cell(r, 1, "ランク・判定の集計（スコアリングタブに連動）").font = HDR_FONT
    r += 1
    for c, h in enumerate(["ランク／判定", "件数", "定義"], 1):
        cell = ws.cell(r, c, h)
        cell.fill = HEADER_FILL; cell.font = HDR_FONT; cell.alignment = CENTER; cell.border = BORDER
    params["rank_count_row"] = r + 1  # 呼び出し側で件数のCOUNTIF数式を入れる

    return params


RANK_COUNT_ROWS = 12  # add_rank_counts が使う行数（ランク5＋警告1＋判定4＋実績2）


def add_rank_counts(ws, start_row, rank_col_range, gate_col_range, judge_col_range, resng_col_range, upwarn_col_range):
    defs = [
        ("S", "60点以上（最優先で周辺の土地を探す）"),
        ("A", "48〜59点（優先候補）"),
        ("B", "36〜47点"),
        ("C", "35点以下"),
        ("データ無", "空容量未公表(OSMのみ等)。スコア比較不能"),
        ("警告", "ゲート該当（空容量(上位系考慮)＜2MW かつ N-1電制≠可）。ランクは通常算出（S1b=0）。判定は「要検討」。上位系の混雑・増強リスクに注意"),
        (JUDGE_OK, JUDGE_MEANING[JUDGE_OK]),
        (JUDGE_CHECK, JUDGE_MEANING[JUDGE_CHECK]),
        (JUDGE_NG, JUDGE_MEANING[JUDGE_NG]),
        (JUDGE_NODATA, JUDGE_MEANING[JUDGE_NODATA]),
        (RESULT_NG_TAG, RESULT_NG_MEANING[RESULT_NG_TAG]),
        (UPPER_WARN_TAG, RESULT_NG_MEANING[UPPER_WARN_TAG]),
    ]
    for i, (rank, note) in enumerate(defs):
        r = start_row + i
        ws.cell(r, 1, rank).border = BORDER
        if rank in (RESULT_NG_TAG, UPPER_WARN_TAG):
            rng = resng_col_range if rank == RESULT_NG_TAG else upwarn_col_range
            fc = ws.cell(r, 2, f'=COUNTIF({rng},"?*")')   # 文字が入っている行の数
        else:
            rng = gate_col_range if rank == "警告" else judge_col_range if rank in JUDGE_MEANING and i >= 6 else rank_col_range
            fc = ws.cell(r, 2, f'=COUNTIF({rng},"{rank}")')
        fc.border = BORDER; fc.alignment = CENTER
        ws.cell(r, 3, note).border = BORDER
    for col, w in [("A", 44), ("B", 14), ("C", 52), ("D", 52), ("E", 44)]:
        ws.column_dimensions[col].width = w


def build_results_sheet(ws):
    """実績DBシート。戻り値: (要約列レンジ, キー列レンジ)"""
    ws["A1"] = (f"接続検討回答 実績{len(RESULTS_DB)}件（手入力{RESULTS_MANUAL_COUNT}件: 2026-08-06更新／"
                f"残り{len(RESULTS_DB) - RESULTS_MANUAL_COUNT}件: 接続検討回答 検証一覧から取込・AI読取値のため要原本確認）")
    ws["A1"].font = TITLE_FONT
    headers = ["エリア", "変電所名(リスト表記)", "案件", "出力", "負担金", "連系工期",
               "主な対策工事", "教訓・律速", "要約(スコア表参照用・同じ変電所の回答はまとめて表示)", "キー(自動)"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(2, c, h)
        cell.fill = HEADER_FILL; cell.font = HDR_FONT; cell.alignment = CENTER; cell.border = BORDER
    for i, row in enumerate(RESULTS_DB):
        r = 3 + i
        for c, v in enumerate(row, 1):
            if c == 9:  # スコア表は最初の一致行しか引けないため、同じ変電所の要約をまとめて載せる
                v = RESULTS_MEMO[results_key(row[0], row[1])]
            cell = ws.cell(r, c, v); cell.border = BORDER; cell.alignment = WRAP
        kc = ws.cell(r, 10, f"=A{r}&TRIM(B{r})")
        kc.border = BORDER; kc.font = SMALL
    last = 2 + len(RESULTS_DB)
    r = last + 2
    ws.cell(r, 1, "■分析メモ(2026-08-06追記・19件ベース)").font = HDR_FONT
    for i, note in enumerate(ANALYSIS_NOTES):
        ws.cell(r + 1 + i, 1, note).font = SMALL
    widths = [8, 20, 24, 12, 30, 30, 34, 30, 40, 18]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A3"
    return f"実績DB!$I$3:$I${last}", f"実績DB!$J$3:$J${last}"


def find_header_row(ws):
    for row in ws.iter_rows(min_row=1, max_row=10):
        if row[0].value == "No.":
            return row[0].row
    raise SystemExit("入力シートに 'No.' ヘッダー行が見つかりません")


def read_substations(path, sheet_name):
    wb = load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise SystemExit(f"シート '{sheet_name}' がありません。存在するシート: {wb.sheetnames}")
    ws = wb[sheet_name]
    hdr = find_header_row(ws)
    headers = [c.value or "" for c in ws[hdr]]

    def col(token):
        for i, h in enumerate(headers):
            if token in str(h).replace("\n", ""):
                return i
        raise SystemExit(f"ヘッダー '{token}' が見つかりません: {headers}")

    idx = {
        "no": col("No."), "area": col("エリア"), "name": col("変電所名"),
        "pref": col("都道府県"), "vmax": col("最大電圧"), "vsec": col("二次電圧"),
        "lat": col("緯度"), "lon": col("経度"), "acc": col("座標精度"),
        "cap": col("設備容量"), "opcap": col("運用容量"), "flow": col("予想潮流"),
        "avail": col("当該設備"), "avail_up": col("上位系考慮"),
        "n1": col("N-1電制"), "curtail": col("出力制御"), "addr": col("所在地"),
    }
    rows = []
    for row in ws.iter_rows(min_row=hdr + 1, values_only=True):
        if row[idx["no"]] is None and row[idx["name"]] is None:
            continue
        rows.append([row[idx[k]] for k in
                     ["no", "area", "name", "pref", "vmax", "vsec", "lat", "lon", "acc",
                      "cap", "opcap", "flow", "avail", "avail_up", "n1", "curtail", "addr"]])
    return rows


def build_scoring_sheet(ws, rows, p, memo_range, key_range):
    ws["A1"] = ("全国変電所スコアリング v2（高圧2MW BESS）　黄色=ランクS/A ／ 橙=ゲート該当（上位系0かつN-1不可の警告。ランクは通常算出・判定は要検討）"
                " ／ 赤=実績NG（直近1年の接続検討で連系不可・負担金4,000万円以上・工期24ヶ月以上 → 判定NG）")
    ws["A1"].font = TITLE_FONT
    headers = ["No.", "エリア", "変電所名", "都道府県", "最大電圧kV", "二次電圧kV", "緯度", "経度",
               "座標精度", "設備容量MW", "運用容量MW", "予想潮流MW", "空容量_当該MW", "空容量_上位系MW",
               "N-1電制", "出力制御", "所在地(参考)", "ゲート", "判定(高圧2MW)", "S1a単体空容量(25)", "S1b上位系(10)", "S2 N-1(10)",
               "S3潮流(15)", "S4出制(5)", "S5配変(10)", "S6エリア(5)", "系統スコア(80)", "ランク", "実績メモ",
               # 実績による格下げ（AD〜AF）。判定(高圧2MW)＝S列はこの3列から数式で決まる
               "判定(公表データのみ)", "実績NG(1年以内)", "上位系の実績警告"]
    adj = result_adjuster()
    for c, h in enumerate(headers, 1):
        cell = ws.cell(2, c, h)
        cell.fill = HEADER_FILL; cell.font = HDR_FONT; cell.alignment = CENTER; cell.border = BORDER

    # 空容量: ゲート・データ有無は上位系考慮(N列)を優先、S1a・判定は当該設備(M列)を優先（基準シートの前提参照）
    eff = lambda r: f'IF($N{r}<>"",$N{r},$M{r})'
    own = lambda r: f'IF($M{r}<>"",$M{r},$N{r})'

    for i, data in enumerate(rows):
        r = 3 + i
        for c, v in enumerate(data, 1):
            ws.cell(r, c, v)
        # 実績による格下げ（Python側で算出して文字で入れる。空なら該当なし）
        no_, area_, name_, vsec_, avail_, avail_up_ = data[0], data[1], data[2], data[5], data[12], data[13]
        pub = judge_row(vsec_, avail_, avail_up_)["judge"]
        a = adj.adjust(area_, name_, no_, pub)
        ws.cell(r, 31, a["resultNg"] or None)
        ws.cell(r, 32, a["upperWarn"] or None)
        f = {
            # R ゲート: 空容量(上位系考慮)＜最低必要容量(2MW) かつ N-1電制≠可 → 警告（ランクには影響させない）
            18: f'=IF(COUNT($M{r}:$N{r})=0,"",IF(AND({eff(r)}<{p["s1_t3"]["thr"]},$O{r}<>"可"),"警告",""))',
            # S 判定(高圧2MW)＝最終判定: 実績NG(AE)があればNG、上位系の実績警告(AF)があり公表データの判定がOKなら要検討、それ以外は公表データの判定(AD)
            19: f'=IF($AE{r}<>"","{JUDGE_NG}",IF(AND($AF{r}<>"",$AD{r}="{JUDGE_OK}"),"{JUDGE_CHECK}",$AD{r}))',
            # AD 判定(公表データのみ): データ無 → 単体＜2MW=NG → 二次電圧7kV超=要検討 → 上位系公表かつ＜2MW=要検討 → OK
            30: f'=IF(COUNT($M{r}:$N{r})=0,"{JUDGE_NODATA}",IF({own(r)}<{p["s1_t3"]["thr"]},"{JUDGE_NG}",'
                f'IF(AND(ISNUMBER($F{r}),$F{r}>{p["s5_1"]["thr"]}),"{JUDGE_CHECK}",'
                f'IF(AND($N{r}<>"",$N{r}<{p["s1b_ok"]["thr"]}),"{JUDGE_CHECK}","{JUDGE_OK}"))))',
            # T S1a 変電所単体（当該設備）の空容量
            20: f'=IF(COUNT($M{r}:$N{r})=0,0,IF({own(r)}>={p["s1_t1"]["thr"]},{p["s1_t1"]["pts"]},'
                f'IF({own(r)}>={p["s1_t2"]["thr"]},{p["s1_t2"]["pts"]},'
                f'IF({own(r)}>={p["s1_t3"]["thr"]},{p["s1_t3"]["pts"]},0))))',
            # U S1b 上位系の空き（上位系考慮が未公表なら不明の点。変電所単体＜2MWなら0）
            21: f'=IF(OR(COUNT($M{r}:$N{r})=0,{own(r)}<{p["s1_t3"]["thr"]}),0,'
                f'IF($N{r}="",{p["s1b_unk"]["pts"]},IF($N{r}>={p["s1b_ok"]["thr"]},{p["s1b_ok"]["pts"]},0)))',
            # V S2 N-1電制
            22: f'=IF($O{r}="可",{p["s2"]["pts"]},0)',
            # W S3 潮流余裕 |予想潮流|/運用容量
            23: f'=IF(OR(COUNT($K{r}:$L{r})<2,$K{r}=0),0,'
                f'IF(ABS($L{r})/$K{r}<={p["s3_t1"]["thr"]},{p["s3_t1"]["pts"]},'
                f'IF(ABS($L{r})/$K{r}<={p["s3_t2"]["thr"]},{p["s3_t2"]["pts"]},'
                f'IF(ABS($L{r})/$K{r}<={p["s3_t3"]["thr"]},{p["s3_t3"]["pts"]},0))))',
            # X S4 出力制御
            24: f'=IF($P{r}="なし",{p["s4"]["pts"]},0)',
            # Y S5 配電用変電所か(二次電圧)
            25: f'=IF($F{r}="",0,IF(ISNUMBER($F{r}),'
                f'IF($F{r}<={p["s5_1"]["thr"]},{p["s5_1"]["pts"]},{p["s5_2"]["pts"]}),0))',
            # Z S6 エリア点
            26: f'=IFERROR(VLOOKUP($B{r},{p["area_range"]},2,FALSE),0)',
            # AA 系統スコア(80)
            27: f'=SUM($T{r}:$Z{r})',
            # AB ランク
            28: f'=IF(COUNT($M{r}:$N{r})=0,"データ無",'
                f'IF($AA{r}>={p["rank_s"]["thr"]},"S",IF($AA{r}>={p["rank_a"]["thr"]},"A",'
                f'IF($AA{r}>={p["rank_b"]["thr"]},"B","C"))))',
            # AC 実績メモ（実績DBのエリア＋変電所名キーと一致すれば要約を表示）
            29: f'=IFERROR(INDEX({memo_range},MATCH($B{r}&$C{r},{key_range},0)),"")',
        }
        for c, formula in f.items():
            ws.cell(r, c, formula)
        for c in list(range(18, 29)) + [30]:
            ws.cell(r, c).alignment = CENTER

    last = 2 + len(rows)
    rng = f"A3:AF{last}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=['$AE3<>""'], fill=RESULT_NG_FILL, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['$R3="警告"'], fill=GATE_FILL, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['OR($AB3="S",$AB3="A")'], fill=RANK_FILL))
    ws.auto_filter.ref = f"A2:AF{last}"
    ws.freeze_panes = "D3"

    widths = [6, 8, 26, 11, 10, 10, 11, 11, 14, 10, 10, 10, 12, 13, 16, 9, 22,
              7, 12, 12, 10, 10, 10, 9, 10, 10, 12, 9, 40, 14, 48, 48]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    for col in ("G", "H"):
        for r in range(3, last + 1):
            ws.cell(r, {"G": 7, "H": 8}[col]).number_format = "0.000000"
    return (f"スコアリング!$AB$3:$AB${last}", f"スコアリング!$R$3:$R${last}", f"スコアリング!$S$3:$S${last}",
            f"スコアリング!$AE$3:$AE${last}", f"スコアリング!$AF$3:$AF${last}")


def main():
    ap = argparse.ArgumentParser(description="全国変電所リストxlsxから変電所スコアリングxlsxを生成")
    ap.add_argument("--in", dest="src", required=True, help="全国変電所リスト（66kV以上）のxlsxパス")
    ap.add_argument("--out", dest="out", required=True, help="出力xlsxパス")
    ap.add_argument("--list-sheet", default="全国変電所リスト", help="入力の変電所リストのシート名")
    args = ap.parse_args()

    rows = read_substations(args.src, args.list_sheet)
    print(f"読込: {len(rows)}行")

    wb = Workbook()
    ws_score = wb.active
    ws_score.title = "スコアリング"
    ws_base = wb.create_sheet(BASE)
    ws_db = wb.create_sheet("実績DB")

    params = build_base_sheet(ws_base)
    memo_range, key_range = build_results_sheet(ws_db)
    rank_range, gate_range, judge_range, resng_range, upwarn_range = build_scoring_sheet(ws_score, rows, params, memo_range, key_range)
    add_rank_counts(ws_base, params["rank_count_row"], rank_range, gate_range, judge_range, resng_range, upwarn_range)

    r = params["rank_count_row"] + RANK_COUNT_ROWS + 2
    ws_base.cell(r, 1, "前提・注意").font = HDR_FONT
    for i, note in enumerate(NOTES):
        ws_base.cell(r + 1 + i, 1, note).font = SMALL

    wb.save(args.out)
    print(f"出力: {args.out}")


if __name__ == "__main__":
    main()
