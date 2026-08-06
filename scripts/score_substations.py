#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
変電所スコアリング xlsx 生成スクリプト（系統用蓄電池・高圧2MW前提 v1）

全国変電所リスト（66kV以上・系統接続検討用）xlsx を読み込み、
接続検討の実績12件から校正した配点（変電所スコアリング v1 2026-07-28）で
全変電所を「接続検討を申し込む価値」の順に点数化した3タブのxlsxを出力する。

  スコアリング     : 全行のゲート判定・S1〜S6・系統スコア(80点)・ランク・実績メモ。
                     スコア列は全て数式で、入力値や基準パラメータを変えると再計算される。
                     ランクS/Aの行は条件付き書式で黄色ハイライト。
  スコアリング基準 : 配点定義・エリア点・ランク閾値のパラメータ表（数式の参照元）と前提。
  実績DB           : 接続検討回答の実績12件。エリア＋変電所名の一致で実績メモ列に自動表示。

スコア構成（系統スコア80点。土地確定後に近接性S7=20点を加えて100点満点）:
  S1 空容量(上位系考慮) 35 / S2 N-1電制 10 / S3 潮流余裕 15 / S4 出力制御 5 / S5 配変 10 / S6 エリア 5
ゲート: 空容量(上位系考慮)≦0 かつ N-1電制≠可 → 除外（上位系増強リスク。実績: 喜多方=工期10年2ヶ月）
ランク: S:60点以上 / A:48-59 / B:36-47 / C:35以下 / データ無:空容量未公表(OSMのみ等)

使い方:
  python scripts/score_substations.py --in 全国変電所リスト.xlsx --out 変電所スコアリング.xlsx
"""
import argparse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="FCE4D6")
PARAM_FILL = PatternFill("solid", fgColor="FFFF00")   # 変更可能なパラメータセル
RANK_FILL = PatternFill("solid", fgColor="FFF2A8")    # ランクS/Aハイライト（黄色）
GATE_FILL = PatternFill("solid", fgColor="E7E6E6")    # ゲート除外行（グレー）
HDR_FONT = Font(bold=True, size=11)
TITLE_FONT = Font(bold=True, size=14)
SMALL = Font(size=10)
THIN = Side(style="thin", color="D0D0D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP = Alignment(wrap_text=True, vertical="center")

BASE = "スコアリング基準"

# S6 エリア点（実績工期・出力制御環境からの仮置き。実績追加で更新）
AREA_POINTS = [
    ("北海道", 5, "由仁町: 計量のみ410万・6ヶ月"),
    ("東北", 4, "喜多方: 配電のみなら7ヶ月"),
    ("東京", 4, "需要地・配変多い(実績は特高のみ)"),
    ("中部", 2, "関市18ヶ月・野尻24ヶ月と増強系が遅い"),
    ("北陸", 3, "実績なし・中位仮置き"),
    ("関西", 3, "白浜: 12ヶ月"),
    ("中国", 3, "実績未読(PW付)・中位仮置き"),
    ("四国", 3, "実績なし・中位仮置き"),
    ("九州", 5, "神崎3ヶ月・出力制御は多め"),
    ("沖縄", 2, "市場環境考慮・仮置き"),
    ("J-POWER", 0, "対象外"),
]

# 接続検討回答 実績12件（詳細: 接続検討回答_分析メモ_20260724.md）
RESULTS_DB = [
    ("九州", "笠之原変電所", "鹿屋市上小原", "1,999kW", "約6,140万円(付箋値・要確認)", "入金後1年4ヶ月", "張替1,405m+SVC", "細線・SVC発動で高額長期", "負担金約6,140万・1年4ヶ月(張替1.4km+SVC)"),
    ("北海道", "由仁変電所", "由仁町川端(1135)", "1,999kW", "410万円", "入金後6ヶ月", "計量設備のみ", "空容量潤沢なら最安最速", "負担金410万・6ヶ月(計量のみ)"),
    ("九州", "王子変電所", "大分市神崎(東盛)", "高圧", "920万円", "入金後3ヶ月", "張替1,006m+PGB", "バンク逆潮流対策済みは別格", "負担金920万・3ヶ月(バンク対策済み)"),
    ("九州", "久山変電所", "福岡県久山町(KMP)", "高圧", "1,630万円(税抜)", "入金後1年5ヶ月", "SVC300kvar+張替16m", "細線(177A)でSVC発動", "負担金1,630万・1年5ヶ月(SVC)"),
    ("東北", "－(須賀川上位系)", "喜多方市松山(PX No071)", "高圧", "特定負担は僅少", "配電7ヶ月/上位系10年2ヶ月", "SVR5000kVA+須賀川MT取替", "上位系増強=時間で死ぬ", "上位系増強で10年2ヶ月(ゲート根拠)"),
    ("中国", "－", "東広島654(1129)", "高圧", "未読(PW付PDF)", "未読", "−", "−", "PW付未読"),
    ("中部", "小屋名変電所", "関市小屋名(B646G)", "高圧", "620万円+保証金31万", "入金後18ヶ月", "配電線増強", "中部の増強系は遅い", "負担金620万・18ヶ月(増強)"),
    ("中部", "柏原変電所", "信濃町柏原2390-1", "高圧", "850万円+保証金42.5万", "入金後6ヶ月", "新設48m+引込105m", "軽微工事なら中部でも速い", "850万・6ヶ月(新設48m)/同F21野尻は1,782万・24ヶ月"),
    ("中部", "柏原変電所", "信濃町野尻(契約申込=確定)", "1,998kW", "1,782万円(税抜)", "入金後24ヶ月", "電柱5本+新設1,842m+SVR取替", "新設km級は高額長期", "(同上メモ参照)"),
    ("東京", "－(西毛線)", "高崎市吉井町(特高参考)", "特高", "4.64億円+保証金2,551万", "入金後26ヶ月(受付後73ヶ月)", "送電線新設+2回線張替", "特高は桁違い", "特高参考: 4.64億・6年"),
    ("関西", "－(田辺管内)", "白浜町才野1339-1", "1,999kW", "約3,700〜3,900万円", "入金後12ヶ月", "高圧線新設4,701m+SVR2台", "遠距離新設の典型NG", "3,700万・12ヶ月(新設4.7km)"),
    ("九州", "－(野原宇佐線)", "宇佐市蜷木(1486)", "高圧", "2.465億円", "入金後2年3ヶ月", "66kV鉄塔・アクセス送電線新設", "66kV側対策で億超え", "2.5億・2年3ヶ月(66kV側対策)"),
]

# 変数・配点定義（実績12件で校正）— 基準シートの表示用
SCORE_DEFS = [
    ("G1", "ゲート: 上位系増強リスク", "除外", "空容量(上位系考慮)≦0 かつ N-1電制≠可 → 除外", "喜多方: 上位系増強で工期10年2ヶ月"),
    ("G2", "ゲート: 募集プロセス/増強エリア", "除外(手動)", "各社公表資料で該当時に手動除外", "高額・長納期化リスク"),
    ("S1", "空容量(上位系考慮)", "35", "≧20MW(2MWの10倍):35 / ≧6MW:28 / >0:15 / ≦0:0", "最重要。上位系増強の回避確率"),
    ("S2", "N-1電制 可否", "10", "可:10 / それ以外:0", "空容量ゼロでも接続余地"),
    ("S3", "予想潮流の余裕(|潮流|/運用容量)", "15", "≦30%:15 / ≦60%:9 / ≦90%:4 / >90%:0", "混雑度=出力制御・増強リスク"),
    ("S4", "出力制御可能性", "5", "なし:5 / 有り:0", "収益影響(九州で頻発)"),
    ("S5", "配電用変電所か(二次電圧)", "10", "二次≦7kV(6.6/6.0):10 / その他数値:5 / 不明:0", "高圧2MW連系の受け皿"),
    ("S6", "エリア", "5", "エリア点表参照", "実績工期のエリア差(九州3ヶ月〜中部24ヶ月)"),
    ("合計", "系統スコア", "80", "S1+S2+S3+S4+S5+S6", ""),
    ("S7", "(土地確定後)変電所・既設配電線への近接", "20", "≦500m:20 / ≦1km:15 / ≦2km:10 / ≦5km:4 / 超:0", "負担金≒400万+0.8〜1.0万円/m×新設延長"),
    ("", "実績補正", "±", "実績DBに同一変電所があれば優先参照(実績メモ列)", "王子SS(バンク対策済み)=神崎920万/3ヶ月 等"),
]

NOTES = [
    "スコアリング v1 2026-07-28（接続検討の実回答12件の分析から逆算した配点）に準拠。",
    "使い方: ランク S/A の変電所を優先し、その周辺（目安2km以内・既設高圧配電線の近く）で土地を探す→接続検討を申込む。",
    "系統スコアは80点満点。土地確定後に近接性S7(20点)を加えて100点満点で最終評価する。",
    "負担金の事前推定式(実績回帰): 推定負担金(万円) ≒ 固定費400 + 高圧線新設・張替延長(m)×0.8〜1.0 + SVC/SVR台数×500〜。",
    "工期目安: 計量のみ3-6ヶ月 / 軽微(〜100m)3-6ヶ月 / 増強・km級12-24ヶ月 / 66kV側対策・上位系増強は2〜10年。",
    "公表空容量は特高側評価。高圧連系の最終確定は接続検討のみ。バンク・配電線の空きは実績DBタブで補正すること。",
    "【前提】S1・ゲートの空容量は「上位系考慮」を優先し、単一値のみ公表のエリア(北陸・中部・沖縄等)は「当該設備」で代用する。",
    "【前提】ランク「データ無」= 空容量が両列とも未公表(OSMのみ等)。ゲート該当行はランク「除外」と表示する。",
    "【前提】S4は出力制御が「なし」(完全一致)の場合のみ加点。「有り」「あり」表記ゆれ・不明は0点。",
]


def build_base_sheet(ws):
    """基準シートを作成し、数式が参照するパラメータセルの絶対参照を dict で返す。"""
    ws["A1"] = "変電所スコアリング基準（系統用蓄電池・高圧2MW前提）v1"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "作成: 2026-07-28 ｜ 接続検討の実回答12件（実績DBタブ）から逆算した配点で「接続検討を申し込む価値」を点数化する"
    ws["A2"].font = SMALL

    r = 4
    ws.cell(r, 1, "変数・配点定義（実績12件で校正）").font = HDR_FONT
    r += 1
    for c, h in enumerate(["記号", "変数", "配点", "閾値・ルール", "根拠(実績)"], 1):
        cell = ws.cell(r, c, h)
        cell.fill = HEADER_FILL; cell.font = HDR_FONT; cell.alignment = CENTER; cell.border = BORDER
    for row in SCORE_DEFS:
        r += 1
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v); cell.border = BORDER; cell.alignment = WRAP

    # 配点パラメータ（数式の参照元。黄色セルを変えるとスコアリングタブが再計算される）
    r += 2
    ws.cell(r, 1, "配点パラメータ（黄色セルを変更するとスコアが自動再計算されます）").font = HDR_FONT
    r += 1
    for c, h in enumerate(["項目", "閾値", "点数"], 1):
        cell = ws.cell(r, c, h)
        cell.fill = HEADER_FILL; cell.font = HDR_FONT; cell.alignment = CENTER; cell.border = BORDER
    params = {}
    param_rows = [
        ("s1_t1", "S1: 空容量 ≧(MW) で35点", 20, 35),
        ("s1_t2", "S1: 空容量 ≧(MW) で28点", 6, 28),
        ("s1_t3", "S1: 空容量 >0 で15点", None, 15),
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
    for key, label, thr, pts in param_rows:
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
    ws.cell(r, 1, "ランク集計（スコアリングタブに連動）").font = HDR_FONT
    r += 1
    for c, h in enumerate(["ランク", "件数", "定義"], 1):
        cell = ws.cell(r, c, h)
        cell.fill = HEADER_FILL; cell.font = HDR_FONT; cell.alignment = CENTER; cell.border = BORDER
    params["rank_count_row"] = r + 1  # 呼び出し側で件数のCOUNTIF数式を入れる

    return params


def add_rank_counts(ws, start_row, rank_col_range):
    defs = [
        ("S", "60点以上（最優先で周辺の土地を探す）"),
        ("A", "48〜59点（優先候補）"),
        ("B", "36〜47点"),
        ("C", "35点以下"),
        ("除外", "ゲート該当（空容量≦0 かつ N-1電制≠可 → 上位系増強リスク）"),
        ("データ無", "空容量未公表(OSMのみ等)。スコア比較不能"),
    ]
    for i, (rank, note) in enumerate(defs):
        r = start_row + i
        ws.cell(r, 1, rank).border = BORDER
        fc = ws.cell(r, 2, f'=COUNTIF({rank_col_range},"{rank}")')
        fc.border = BORDER; fc.alignment = CENTER
        ws.cell(r, 3, note).border = BORDER
    for col, w in [("A", 44), ("B", 14), ("C", 52), ("D", 52), ("E", 44)]:
        ws.column_dimensions[col].width = w


def build_results_sheet(ws):
    """実績DBシート。戻り値: (要約列レンジ, キー列レンジ)"""
    ws["A1"] = "接続検討回答 実績12件（詳細: 接続検討回答_分析メモ_20260724.md）"
    ws["A1"].font = TITLE_FONT
    headers = ["エリア", "変電所名(リスト表記)", "案件", "出力", "負担金", "連系工期",
               "主な対策工事", "教訓", "要約(スコア表参照用)", "キー(自動)"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(2, c, h)
        cell.fill = HEADER_FILL; cell.font = HDR_FONT; cell.alignment = CENTER; cell.border = BORDER
    for i, row in enumerate(RESULTS_DB):
        r = 3 + i
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v); cell.border = BORDER; cell.alignment = WRAP
        kc = ws.cell(r, 10, f"=A{r}&TRIM(B{r})")
        kc.border = BORDER; kc.font = SMALL
    last = 2 + len(RESULTS_DB)
    widths = [8, 20, 24, 10, 26, 22, 28, 26, 40, 18]
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
    ws["A1"] = "全国変電所スコアリング（高圧2MW BESS）　黄色=ランクS/A ／ グレー=ゲート除外"
    ws["A1"].font = TITLE_FONT
    headers = ["No.", "エリア", "変電所名", "都道府県", "最大電圧kV", "二次電圧kV", "緯度", "経度",
               "座標精度", "設備容量MW", "運用容量MW", "予想潮流MW", "空容量_当該MW", "空容量_上位系MW",
               "N-1電制", "出力制御", "所在地(参考)", "ゲート", "S1空容量(35)", "S2 N-1(10)",
               "S3潮流(15)", "S4出制(5)", "S5配変(10)", "S6エリア(5)", "系統スコア(80)", "ランク", "実績メモ"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(2, c, h)
        cell.fill = HEADER_FILL; cell.font = HDR_FONT; cell.alignment = CENTER; cell.border = BORDER

    # 空容量: 上位系考慮(N列)を優先、無ければ当該設備(M列)で代用（基準シートの前提参照）
    eff = lambda r: f'IF($N{r}<>"",$N{r},$M{r})'

    for i, data in enumerate(rows):
        r = 3 + i
        for c, v in enumerate(data, 1):
            ws.cell(r, c, v)
        f = {
            # R ゲート: 空容量≦0 かつ N-1電制≠可 → 除外
            18: f'=IF(COUNT($M{r}:$N{r})=0,"",IF(AND({eff(r)}<=0,$O{r}<>"可"),"除外",""))',
            # S1 空容量(上位系考慮)
            19: f'=IF(COUNT($M{r}:$N{r})=0,0,IF({eff(r)}>={p["s1_t1"]["thr"]},{p["s1_t1"]["pts"]},'
                f'IF({eff(r)}>={p["s1_t2"]["thr"]},{p["s1_t2"]["pts"]},'
                f'IF({eff(r)}>0,{p["s1_t3"]["pts"]},0))))',
            # S2 N-1電制
            20: f'=IF($O{r}="可",{p["s2"]["pts"]},0)',
            # S3 潮流余裕 |予想潮流|/運用容量
            21: f'=IF(OR(COUNT($K{r}:$L{r})<2,$K{r}=0),0,'
                f'IF(ABS($L{r})/$K{r}<={p["s3_t1"]["thr"]},{p["s3_t1"]["pts"]},'
                f'IF(ABS($L{r})/$K{r}<={p["s3_t2"]["thr"]},{p["s3_t2"]["pts"]},'
                f'IF(ABS($L{r})/$K{r}<={p["s3_t3"]["thr"]},{p["s3_t3"]["pts"]},0))))',
            # S4 出力制御
            22: f'=IF($P{r}="なし",{p["s4"]["pts"]},0)',
            # S5 配電用変電所か(二次電圧)
            23: f'=IF($F{r}="",0,IF(ISNUMBER($F{r}),'
                f'IF($F{r}<={p["s5_1"]["thr"]},{p["s5_1"]["pts"]},{p["s5_2"]["pts"]}),0))',
            # S6 エリア点
            24: f'=IFERROR(VLOOKUP($B{r},{p["area_range"]},2,FALSE),0)',
            # 系統スコア(80)
            25: f'=SUM($S{r}:$X{r})',
            # ランク
            26: f'=IF($R{r}="除外","除外",IF(COUNT($M{r}:$N{r})=0,"データ無",'
                f'IF($Y{r}>={p["rank_s"]["thr"]},"S",IF($Y{r}>={p["rank_a"]["thr"]},"A",'
                f'IF($Y{r}>={p["rank_b"]["thr"]},"B","C")))))',
            # 実績メモ（実績DBのエリア＋変電所名キーと一致すれば要約を表示）
            27: f'=IFERROR(INDEX({memo_range},MATCH($B{r}&$C{r},{key_range},0)),"")',
        }
        for c, formula in f.items():
            ws.cell(r, c, formula)
        for c in range(18, 27):
            ws.cell(r, c).alignment = CENTER

    last = 2 + len(rows)
    rng = f"A3:AA{last}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=['$Z3="除外"'], fill=GATE_FILL, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['OR($Z3="S",$Z3="A")'], fill=RANK_FILL))
    ws.auto_filter.ref = f"A2:AA{last}"
    ws.freeze_panes = "D3"

    widths = [6, 8, 26, 11, 10, 10, 11, 11, 14, 10, 10, 10, 12, 13, 16, 9, 22,
              7, 11, 10, 10, 9, 10, 10, 12, 9, 40]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    for col in ("G", "H"):
        for r in range(3, last + 1):
            ws.cell(r, {"G": 7, "H": 8}[col]).number_format = "0.000000"
    return f"スコアリング!$Z$3:$Z${last}"


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
    rank_range = build_scoring_sheet(ws_score, rows, params, memo_range, key_range)
    add_rank_counts(ws_base, params["rank_count_row"], rank_range)

    r = params["rank_count_row"] + 8
    ws_base.cell(r, 1, "前提・注意").font = HDR_FONT
    for i, note in enumerate(NOTES):
        ws_base.cell(r + 1 + i, 1, note).font = SMALL

    wb.save(args.out)
    print(f"出力: {args.out}")


if __name__ == "__main__":
    main()
