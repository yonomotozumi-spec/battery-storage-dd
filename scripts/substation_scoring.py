#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
変電所スコアリングの配点定義（系統用蓄電池・高圧2MW前提 v2 / 2026-11-06）。

配点・閾値・エリア点・実績DB・接続検討の判定（OK／要検討／NG）の単一ソース。
xlsx生成（score_substations.py）とWeb用データ生成（build_substation_web_data.py）が
ともにこのモジュールを参照するため、定義はここだけを直せば両方に反映される。

スコア構成:
  系統スコア80点 = S1a 変電所単体の空容量25 + S1b 上位系の空き10 + S2 N-1電制10 + S3潮流15 + S4出力制御5 + S5配変10 + S6エリア5
  総合スコア100点 = 系統スコア80 + S7近接20（土地の座標が決まってから加算）
最低必要容量: 空容量が MIN_MW（2MW＝高圧の案件規模）未満は「空き無し」と同じ扱い（S1a=0点・判定NG・地図の色）

v2（2026-11-06）での変更 — 「高圧2MWは上位系が逼迫していても変電所（配電用バンク）に空きがあれば連系できることが多い」を反映:
  v1までは S1(35点) を「空容量(上位系考慮)」で採点していたため、上位系が0の変電所は当該設備（配電用バンク）に
  20MW空いていても S1=0点・ランクC（⚠警告）になり、全6,071件のうち2,347件がCに落ちていた。
  実績DBでは上位系0の変電所への高圧の回答32件のうち31件が連系可（王子3ヶ月・由仁6ヶ月・曽根12〜30ヶ月）で、
  上位系0は「繋がらない」ではなく「工期が長くなりやすい（中央値35ヶ月 vs 14ヶ月）・出力制御あり」の目安。
  そこで S1 を2つに分けた:
    S1a 変電所単体（当該設備＝配電用バンク）の空容量 25点  ≧20MW:25 / ≧6MW:18 / ≧2MW:5 / ＜2MW:0
    S1b 上位系の空き                             10点  空きあり(≧2MW):10 / 未公表(不明):10 / 空き無し(＜2MW):0（単体が＜2MWなら0）
  未公表（東北・北陸・中部・四国・沖縄）は v1 の「当該設備で代用」と同じく確認済みと同点にしてある。S1aの段階は
  S1a＋S1b(10) が v1 の S1（35/28/15）と同じ点になるよう 25/18/5 にした。つまり上位系0の変電所だけが変わり（S1=0 → S1a分を加点）、
  他の変電所の点数・ランクは v1 から1件も動かない。未公表を5点に下げると中部などで S→A が435件動くため見送り。
  変えるなら PARAM_ROWS の s1b_unk の1か所。判定・画面では「上位系未確認」と表示して注意を促す。
ゲート（⚠上位系0）: 空容量(上位系考慮)＜2MW かつ N-1電制≠可 → 警告（ランクは通常どおり算出）。v2では判定「要検討」に対応する。

接続検討の判定（高圧2MW・1変電所ごと。judge_row）:
  接続検討NG  変電所単体（当該設備）の空容量が2MW未満（空き無し）
  要検討      変電所単体は空きあり だが 上位系の空き無し（工期・出力制御のリスクを人が判断）、
              または 二次電圧が7kV超（配電用変電所でなく高圧2MWの受け皿にならない可能性）
  接続検討OK  変電所単体に空きあり かつ 上位系が空きあり または 未公表（未公表は「上位系未確認」を付ける）
  データ無    空容量が未公表で判定できない
候補地（緯度経度）の判定（judge_site）: 2km以内に「接続検討OK」の変電所があればOK、10km以内にOK／要検討の変電所しか
  無ければ要検討（距離・上位系の理由を出す）、10km以内の空容量データありの変電所がすべてNGならNG、データが無ければ要検討。
実績による格下げ（2026-11-06、余野さん判断。ResultAdjuster）: 直近1年以内の接続検討の回答が 連系不可／負担金4,000万円以上／
  工期24ヶ月以上 のいずれか（＝実績NG）なら、その変電所（名前が確定している実績のみ）は公表の空容量が空きありでも「接続検討NG」。
  実績NGの回答が名指しした上位系の設備（上位変電所の変圧器など）を上位系の制約設備として持つ変電所には「⚠上位系に実績NG」の
  警告を付け、接続検討OKなら「要検討」に落とす。公表データだけの判定は judgePub として残す。
"""

import json
import os

MIN_MW = 2       # 最低必要容量（高圧2MWの案件規模）。これ未満の空容量は「空き無し」と同じ扱い（2026-10-20 余野さん判断）

# ── 実績による格下げの閾値（2026-11-06 余野さん判断。詳細はファイル末尾「実績による格下げ」）──
RESULT_NG_YEN_MAN = 4000      # 負担金（万円）がこれ以上なら実績NG（4,000万円）
RESULT_NG_MONTHS = 24         # 連系工期（ヶ月）がこれ以上なら実績NG
RESULT_NG_WINDOW_DAYS = 365   # 回答日がこの日数以内の実績だけを見る（直近1年）
RESULT_NG_TAG = "実績NG"
UPPER_WARN_TAG = "⚠上位系に実績NG"


def fmt_yen_man(v):
    """万円の数値を「4,000万円」「2.07億円」のように表示する。"""
    return f"{v / 10000:g}億円" if v >= 10000 else f"{v:,.0f}万円"


RESULT_NG_YEN_TXT = fmt_yen_man(RESULT_NG_YEN_MAN)   # "4,000万円"
GRID_MAX = 80    # 系統スコア満点
S7_MAX = 20      # 近接スコア満点
TOTAL_MAX = 100  # 総合スコア満点

# ── 配点パラメータ（key, ラベル, 閾値, 点数）──
PARAM_ROWS = [
    ("s1_t1", "S1a: 変電所単体（当該設備）の空容量 ≧(MW) で25点", 20, 25),
    ("s1_t2", "S1a: 変電所単体の空容量 ≧(MW) で18点", 6, 18),
    ("s1_t3", "S1a: 変電所単体の空容量 ≧(MW) で5点（最低必要容量）", MIN_MW, 5),
    ("s1b_ok", "S1b: 上位系の空き（空容量(上位系考慮)）≧(MW) で10点", MIN_MW, 10),
    ("s1b_unk", "S1b: 上位系が未公表（不明）のとき（確認済みと同点。5点に下げると中部などS→Aが435件動く）", None, 10),
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

# ── 接続検討の判定（高圧2MW）のラベルと、候補地の判定に使う距離 ──
JUDGE_OK, JUDGE_CHECK, JUDGE_NG, JUDGE_NODATA = "接続検討OK", "要検討", "接続検討NG", "データ無"
SITE_OK_M = 2000      # 候補地の判定: この距離以内に「接続検討OK」の変電所があれば候補地もOK（周辺2km以内で土地を探す目安と同じ）
SITE_RANGE_M = 10000  # 候補地の判定: この距離以内の変電所だけを材料にする（それより遠い変電所は候補地の判定に使わない）
JUDGE_MEANING = {
    JUDGE_OK: "接続検討OK：変電所単体（配電用バンク）に空きがあり、上位系も空きあり（または未公表）。接続検討を申し込む",
    JUDGE_CHECK: "要検討：変電所単体は空きありだが上位系の空き無し（高圧2MWは実績でほぼ連系可だが工期が長く出力制御あり）、"
                 "または配電用変電所でない。人が工期・出力制御・近くの別候補と比べて判断する",
    JUDGE_NG: "接続検討NG：変電所単体（配電用バンク）の空容量が2MW未満で空き無し。別の変電所を探す",
    JUDGE_NODATA: "データ無：空容量が未公表で判定できない。送配電事業者の公表資料か接続検討で確認",
}

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
AREA_POINT_MAP = {area: pts for area, pts, _ in AREA_POINTS}

# 接続検討回答 実績（手入力分。2026-08-06更新。詳細: 接続検討回答_分析メモ_20260724.md＋追加実績メモ20260806）
RESULTS_DB = [
    ("九州", "笠之原変電所", "鹿屋市上小原", "1,999kW", "約6,140万円(付箋値・要確認)", "入金後1年4ヶ月", "張替1,405m+SVC", "細線・SVC発動で高額長期", "負担金約6,140万・1年4ヶ月(張替1.4km+SVC)"),
    ("北海道", "由仁変電所", "由仁町川端(1135)", "1,999kW", "410万円", "入金後6ヶ月", "計量設備のみ", "空容量潤沢なら最安最速", "負担金410万・6ヶ月(計量のみ)"),
    ("九州", "王子変電所", "大分市神崎(東盛)", "高圧", "920万円", "入金後3ヶ月", "張替1,006m+PGB", "バンク逆潮流対策済みは別格", "負担金920万・3ヶ月(バンク対策済み)"),
    ("九州", "久山変電所", "福岡県久山町(KMP)", "高圧", "1,630万円(税抜)", "入金後1年5ヶ月", "SVC300kvar+張替16m", "細線(177A)でSVC発動", "負担金1,630万・1年5ヶ月(SVC)"),
    ("東北", "－(須賀川上位系)", "喜多方市松山(PX No071)", "高圧", "特定負担は僅少", "配電7ヶ月/上位系10年2ヶ月", "SVR5000kVA+須賀川MT取替", "上位系増強=時間で死ぬ", "上位系増強で10年2ヶ月(ゲート根拠)"),
    ("中部", "小屋名変電所", "関市小屋名(B646G)", "高圧", "620万円+保証金31万", "入金後18ヶ月", "配電線増強", "中部の増強系は遅い", "負担金620万・18ヶ月(増強)"),
    ("中部", "柏原変電所", "信濃町柏原2390-1", "高圧", "850万円+保証金42.5万", "入金後6ヶ月", "新設48m+引込105m", "軽微工事なら中部でも速い", "850万・6ヶ月(新設48m)/同F21野尻は1,782万・24ヶ月"),
    ("中部", "柏原変電所", "信濃町野尻(契約申込=確定)", "1,998kW", "1,782万円(税抜)", "入金後24ヶ月", "電柱5本+新設1,842m+SVR取替", "新設km級は高額長期", "(同上メモ参照)"),
    ("東京", "－(西毛線)", "高崎市吉井町(特高参考)", "特高", "4.64億円+保証金2,551万", "入金後26ヶ月(受付後73ヶ月)", "送電線新設+2回線張替", "特高は桁違い", "特高参考: 4.64億・6年"),
    ("関西", "－(田辺管内)", "白浜町才野1339-1", "1,999kW", "約3,700〜3,900万円", "入金後12ヶ月", "高圧線新設4,701m+SVR2台", "遠距離新設の典型NG", "3,700万・12ヶ月(新設4.7km)"),
    ("九州", "－(野原宇佐線)", "宇佐市蜷木(1486)", "高圧", "2.465億円→契約時点2.704億円(税抜)", "入金後2年3ヶ月→2年10ヶ月(契約時No24595)", "66kV鉄塔・アクセス送電線新設", "66kV側対策で億超え", "2.5億・2年3ヶ月(66kV側対策)"),
    ("九州", "－(西佐世保相浦線)", "伊万里市東山代町東大久保(特高参考)", "特高40,232kW", "2.68億円(税込・税抜2.43億)", "入金後3年2ヶ月", "アクセス線新設0.1km+鉄塔建替2基+張替0.9km×2回線+転送遮断装置3面", "特高40MW級は短距離でも2.5億・3年級", "特高参考: 2.68億・3年2ヶ月(九電回答2026/7/14)"),
    ("中部", "東青島変電所", "藤枝市(連系承諾=確定)", "1,998kW", "2,536万円(税込)", "入金後26ヶ月(受付後73ヶ月)", "電柱5本+新設373m+張替735m+SVR1台+バンク逆潮流対策(PT+OVGR)", "バンク逆潮流対策込でも2,500万台・ただし26ヶ月", "確定値: 2,536万・26ヶ月(承諾2026/5/18)"),
    ("中部", "片田変電所", "志摩市大王町波切(B647G)", "1,999kW", "3,246万円(税抜)", "入金後18ヶ月", "電柱9本新設+8本建替+新設617m+張替1,064m+SVR2台", "新設0.6km+張替1km+SVR2台で3,200万・18ヶ月(工事量比例)", "負担金3,246万・18ヶ月(回答2026/5/13)"),
    ("九州", "－(八代配電管内)", "八代市鏡町上鏡(東盛 No30979)", "1,987kW", "1,170万円(税込)", "入金後6ヶ月", "張替1,235m(200SQ→400SQ)+電柱建替6本・上位系統対策なし", "張替のみは1,000万・6ヶ月(九州の勝ち筋)", "負担金1,170万・6ヶ月(張替1.2km)"),
    ("中部", "初倉変電所", "島田市中河(GEネックス)", "1,997kW", "198万円(税込)+保証金9.9万", "入金後18ヶ月", "電柱新設2本+高圧線新設150m+引込15m(OCW-60)", "実績中最安198万・ただし軽微工事でも中部は18ヶ月", "負担金198万・18ヶ月(回答2026/3/17)"),
    ("九州", "－(西山鹿変電所系)", "熊本県南関町B(西日本PE・特高参考・A/B/C同規模3件)", "特高67,668kW", "1.69億円(税抜・3件按分後)", "入金後2年8ヶ月", "連系線0.1km+光ケーブル13km+転送遮断(220kV三池)", "特高でも共同按分が効くと1.7億に低減・光ケーブル13kmの通信費が重い", "特高参考: 1.69億・2年8ヶ月"),
    ("中部", "－(度会橋連絡線)", "伊勢市矢持町(PX自社・特高参考)", "特高49,880kW", "9.23億円(税込・概算工事費23.6億)", "入金後57カ月", "鉄塔建替1基+引出設備 12回線建替+調相開閉器9台+N-1充電停止装置(77kV側広範)", "中部特高は9億・57カ月と別次元(希望日に間に合わず)", "特高参考: 9.23億・57カ月(回答2026/6/16)"),
]

# 分析メモ（2026-08-06追記・19件ベース）
ANALYSIS_NOTES = [
    "①工期は「エリア×工事種別」で決まる: 九州は張替のみなら3〜6ヶ月(神崎3ヶ月・上鏡6ヶ月)。中部は最軽微(新設150mだけ)でも18ヶ月と標準工程が長い。S6エリア点(九州5・中部2)は実績と整合しており変更不要",
    "②負担金は工事量に比例: 198万(新設150m)→1,170万(張替1.2km)→2,536万(新設373m+張替735m+SVR+バンク対策)→3,246万(新設617m+張替1km+SVR2台)。推定式は「固定費200〜400万+延長(m)×0.7〜1.0万+SVR/SVC×500万+バンク対策500〜1,000万」に更新余地",
    "③契約段階の増額リスク: 宇佐蜷木は接続検討→契約申込で負担金+2,400万(+10%)・工期+7ヶ月。接続検討回答値には+10〜15%・+半年のバッファを見る",
    "④特高は4件(吉井4.64億/伊万里2.68億/南関1.69億/伊勢9.23億)で負担金・工期(26〜73ヶ月)とも桁違い。高圧2MW戦略の正しさを再確認。ただし南関のように共同按分が効くと1.7億まで下がる例あり",
    "⑤上位系統対策の有無が時間の分水嶺: 上鏡は「上位系統対策なし」で6ヶ月。喜多方(上位系10年)との対比でゲート足切りの妥当性を再確認",
    "⑥バンク逆潮流対策は2例目(藤枝2,536万)。対策済み変電所(神崎920万)との差≒約1,000万強=対策費。対策済みバンクの再利用価値大。残課題: 東電JK2件・PW付PDF2件(東広島・山口)・鹿屋原本確認",
]

# 変数・配点定義（実績12件で校正）— 基準シートの表示用
SCORE_DEFS = [
    ("判定", "接続検討の判定（高圧2MW）", "OK/要検討/NG", "NG: 変電所単体(当該設備)＜2MW ／ 要検討: 単体は空きありだが上位系空き無し、または二次電圧7kV超 ／ OK: 単体空きあり かつ 上位系空きあり・未公表 ／ データ無: 空容量未公表", "上位系0の変電所への高圧の回答32件中31件が連系可（王子3ヶ月・由仁6ヶ月・曽根12〜30ヶ月）。ただし工期中央値35ヶ月（非該当14ヶ月）"),
    ("G1", "ゲート: 上位系0（警告）", "警告", "空容量(上位系考慮)＜2MW かつ N-1電制≠可 → ⚠警告（ランクは通常算出。判定は要検討）", "喜多方: 上位系増強で工期10年2ヶ月／由仁: 上位系0でも410万・6ヶ月で連系（2026-10-20に除外→警告）"),
    ("G2", "ゲート: 募集プロセス/増強エリア", "除外(手動)", "各社公表資料で該当時に手動除外", "高額・長納期化リスク"),
    ("S1a", "変電所単体（当該設備＝配電用バンク）の空容量", "25", "≧20MW(2MWの10倍):25 / ≧6MW:18 / ≧2MW:5 / ＜2MW:0（最低必要容量2MW未満は空き無し扱い。S1a+S1bがv1のS1 35/28/15と同じ点になる段階）", "最重要。高圧2MWの受け皿はこのバンクの空き（v2で上位系考慮から当該設備に変更）"),
    ("S1b", "上位系の空き（空容量(上位系考慮)）", "10", "≧2MW:10 / 未公表(不明):10（上位系未確認と表示） / ＜2MW:0（変電所単体が＜2MWのときも0）", "上位系0でも高圧はほぼ繋がる（31/32）が工期中央値35ヶ月・出力制御あり"),
    ("S2", "N-1電制 可否", "10", "可:10 / それ以外:0", "空容量ゼロでも接続余地"),
    ("S3", "予想潮流の余裕(|潮流|/運用容量)", "15", "≦30%:15 / ≦60%:9 / ≦90%:4 / >90%:0", "混雑度=出力制御・増強リスク"),
    ("S4", "出力制御可能性", "5", "なし:5 / 有り:0", "収益影響(九州で頻発)"),
    ("S5", "配電用変電所か(二次電圧)", "10", "二次≦7kV(6.6/6.0):10 / その他数値:5 / 不明:0", "高圧2MW連系の受け皿"),
    ("S6", "エリア", "5", "エリア点表参照", "実績工期のエリア差(九州3ヶ月〜中部24ヶ月)"),
    ("合計", "系統スコア", "80", "S1a+S1b+S2+S3+S4+S5+S6", ""),
    ("実績", "実績による判定の格下げ（直近1年の接続検討の回答）", "NG／要検討",
     f"実績NG＝回答日が直近{RESULT_NG_WINDOW_DAYS // 365}年以内で 連系不可／負担金{RESULT_NG_YEN_TXT}以上／工期{RESULT_NG_MONTHS}ヶ月以上 のいずれか。"
     "変電所名が確定している実績NGがあれば判定を接続検討NGに。名指しされた上位系の設備を制約設備として持つ変電所は⚠上位系に実績NG（OK→要検討）",
     "牟礼（公表8MW・OKだが連系不可 2.07億円）、遠野（岩手(変)MT取替 80.8億円）など。推定の実績は変電所のNGには使わない"),
    ("S7", "(土地確定後)変電所・既設配電線への近接", "20", "≦500m:20 / ≦1km:15 / ≦2km:10 / ≦5km:4 / 超:0", "負担金≒400万+0.8〜1.0万円/m×新設延長"),
    ("", "実績補正", "±", "実績DBに同一変電所があれば優先参照(実績メモ列)", "王子SS(バンク対策済み)=神崎920万/3ヶ月 等"),
]

NOTES = [
    "スコアリング v2 2026-11-06。v1（2026-07-28、接続検討の実回答12件から逆算）の S1空容量35点を、S1a 変電所単体（当該設備）25点＋S1b 上位系の空き10点に分けた。高圧2MWは上位系が逼迫していても配電用バンクに空きがあれば連系できることが多い（実績32件中31件）ため、上位系0を一律0点にしない。",
    "接続検討の判定（高圧2MW）: 変電所単体＜2MW＝接続検討NG ／ 単体は空きありだが上位系空き無し・配電用変電所でない＝要検討（人が判断） ／ 単体空きあり かつ 上位系空きあり・未公表＝接続検討OK。候補地は2km以内にOKの変電所があればOK、10km以内にOK・要検討しか無ければ要検討、すべてNGならNG。",
    "使い方: ランク S/A の変電所を優先し、その周辺（目安2km以内・既設高圧配電線の近く）で土地を探す→接続検討を申込む。",
    "系統スコアは80点満点。土地確定後に近接性S7(20点)を加えて100点満点で最終評価する。",
    "負担金の事前推定式(実績回帰・2026-08-06更新): 推定負担金(万円) ≒ 固定費200〜400 + 高圧線新設・張替延長(m)×0.7〜1.0 + SVC/SVR台数×500〜 + バンク逆潮流対策500〜1,000。",
    "工期目安: 九州は計量・張替のみ3-6ヶ月 / 中部は最軽微工事でも18ヶ月・増強系18-26ヶ月 / SVC込16-17ヶ月 / 特高・上位系増強は2〜10年。",
    "接続検討→契約申込で +10〜15%・+半年 の増額リスクあり(宇佐蜷木実績・2026-08-06更新)。",
    "公表空容量は特高側評価。高圧連系の最終確定は接続検討のみ。バンク・配電線の空きは実績DBタブで補正すること。",
    "【前提】S1aは「当該設備」の空容量（未公表なら「上位系考慮」で代用）、S1bとゲートは「上位系考慮」。上位系を公表していないエリア(東北・北陸・中部・四国・沖縄)はS1bを確認済みと同点にし、判定は「接続検討OK（上位系未確認）」と表示する。",
    "【前提】ランク「データ無」= 空容量が両列とも未公表(OSMのみ等)。ゲート該当行（上位系0かつN-1不可）は「⚠上位系0」の警告を付けて通常どおりランク付けし、判定は「要検討」（2026-10-20までは「除外」として非表示、2026-11-06までは S1=0でランクC）。",
    "【前提】最低必要容量は2MW（高圧の案件規模）。空容量が2MW未満の変電所は「空き無し」と同じ扱い（S1a=0点・判定NG・ゲート判定・地図の色）。2026-10-20に「0以下」から変更。",
    "【前提】S4は出力制御が「なし」(完全一致)の場合のみ加点。「有り」「あり」表記ゆれ・不明は0点。",
    f"【実績による格下げ】直近{RESULT_NG_WINDOW_DAYS // 365}年以内の接続検討の回答が 連系不可／負担金{RESULT_NG_YEN_TXT}以上／工期{RESULT_NG_MONTHS}ヶ月以上 のいずれか（実績NG）なら、"
    "その変電所は公表の空容量が空きありでも判定「接続検討NG」（2026-11-06 余野さん判断。変電所名が確定している実績のみ。推定は使わない）。"
    "実績NGが名指しした上位系の設備（上位変電所の変圧器など）を上位系の制約設備として持つ変電所は「⚠上位系に実績NG」の警告を付け、接続検討OKは要検討に落とす。"
    "公表データだけの判定は「判定(公表データのみ)」に残す。xlsxの判定(高圧2MW)列は 判定(公表データのみ)・実績NG・上位系の実績警告 の3列から数式で決まる（実績の2列はPython側で算出）。"
    "閾値は substation_scoring.py の RESULT_NG_YEN_MAN・RESULT_NG_MONTHS・RESULT_NG_WINDOW_DAYS。",
    "【参考】外部DB: 蓄電所ネット「系統連系診断」(https://bess-net.jp/tools/grid-connection-check) は10社8,346変電所の公表データで連系候補Top5を出す無料ツール(2026-10-12更新)。"
    "空容量は当該設備ベースで上位系のゲートが無いため本ツールより甘く出る(例: 旭川中央は上位系考慮0MW・N-1不可でも55点)。本ツールでは上位系0は S1b=0点・判定「要検討」になる。クロスチェック用に使い、採否は本ツールの判定と接続検討で判断する。",
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


def own_capacity(avail_cur, avail_up):
    """S1a・判定で使う「変電所単体（当該設備＝配電用バンク）」の空容量。

    当該設備を優先し、無ければ上位系考慮で代用（上位系考慮≦当該設備なので控えめな値になる）。
    どちらも未公表なら None（= 空容量データ無し）。
    """
    cur, up = _num(avail_cur), _num(avail_up)
    return cur if cur is not None else up


def judge_row(vsec, avail_cur, avail_up):
    """1変電所の接続検討の判定（高圧2MW）。戻り値: {"judge", "reason", "up_known"}

    判定の順序（score_substations.py の数式・substation.html の judgeOf と同じにすること）:
      1. 空容量が両列とも未公表 → データ無
      2. 変電所単体（当該設備）の空容量 ＜ MIN_MW → 接続検討NG
      3. 二次電圧が数値で 7kV 超（配電用変電所でない） → 要検討
      4. 上位系考慮が公表されていて ＜ MIN_MW → 要検討
      5. それ以外 → 接続検討OK（上位系考慮が未公表なら up_known=False＝「上位系未確認」）
    """
    cur, up = _num(avail_cur), _num(avail_up)
    own = own_capacity(avail_cur, avail_up)
    if own is None:
        return {"judge": JUDGE_NODATA, "up_known": False,
                "reason": "空容量が未公表（OSM由来など）で判定できない。送配電事業者の公表資料か接続検討で確認"}
    own_txt = f"{own:g}MW" + ("" if cur is not None else "（当該設備は未公表のため上位系考慮の値）")
    if own < MIN_MW:
        return {"judge": JUDGE_NG, "up_known": up is not None,
                "reason": f"変電所単体（当該設備＝配電用バンク）の空容量 {own_txt} が最低必要容量{MIN_MW}MW未満で空き無し"}
    v = _num(vsec)
    if v is not None and v > P["s5_1"]["thr"]:
        how = (f"高圧2MWの連系には{v:g}kV/6kV柱上変圧器の新設などが要る（中国エリアの22kV配電で実績あり）"
               if v <= 33 else "高圧2MWの連系先（6.6kV配電線）の受け皿にならない可能性がある")
        return {"judge": JUDGE_CHECK, "up_known": up is not None,
                "reason": f"二次電圧 {v:g}kV で6.6kVの配電用変電所ではない。{how}ため、近くの配電用変電所と比べて判断する"
                          f"（変電所単体の空容量 {own_txt}）"}
    if up is not None and up < MIN_MW:
        extra = (f"。当該設備≧{P['s1_t2']['thr']}MWは実績で上位系0でも通った例（王子・由仁・曽根）と同じ条件"
                 if own >= P["s1_t2"]["thr"] else "")
        return {"judge": JUDGE_CHECK, "up_known": True,
                "reason": f"変電所単体は空きあり {own_txt} だが上位系の空き無し（空容量(上位系考慮) {up:g}MW）。"
                          f"高圧2MWは実績32件中31件が連系可だが工期中央値35ヶ月（非該当14ヶ月）・出力制御の見込み。工期を許容できるか人が判断{extra}"}
    if up is None:
        return {"judge": JUDGE_OK, "up_known": False,
                "reason": f"変電所単体に空きあり {own_txt}。上位系は未公表（上位系未確認）のため、接続検討で上位系の扱いを確認する"}
    return {"judge": JUDGE_OK, "up_known": True,
            "reason": f"変電所単体に空きあり {own_txt}、上位系も空きあり（空容量(上位系考慮) {up:g}MW）"}


def judge_site(cands):
    """候補地（緯度経度）の判定。cands は [(judge, distance_m, name), ...]（判定済みの変電所と直線距離）。

    戻り値: {"judge", "reason", "key"（判定の決め手になった変電所の名前）}
      OK   : SITE_OK_M 以内に「接続検討OK」の変電所がある
      要検討: SITE_RANGE_M 以内にOKの変電所があるが SITE_OK_M より遠い／OKは無いが要検討の変電所がある
      NG   : SITE_RANGE_M 以内の空容量データありの変電所がすべてNG
      要検討: SITE_RANGE_M 以内に空容量データありの変電所が無い（判定材料なし）
    """
    inrange = sorted([c for c in cands if c[1] is not None and c[1] <= SITE_RANGE_M and c[0] != JUDGE_NODATA], key=lambda c: c[1])
    km = lambda m: f"{round(m / 1000, 1):g}km" if m >= 1000 else f"{m:.0f}m"
    oks = [c for c in inrange if c[0] == JUDGE_OK]
    checks = [c for c in inrange if c[0] == JUDGE_CHECK]
    if oks and oks[0][1] <= SITE_OK_M:
        j, d, n = oks[0]
        return {"judge": JUDGE_OK, "key": n, "cause": f"{km(SITE_OK_M)}以内にOKの変電所",
                "reason": f"{km(SITE_OK_M)}以内に接続検討OKの変電所あり（{n}・{km(d)}）。周辺の既設高圧配電線を確認して接続検討を申し込む"}
    if oks:
        j, d, n = oks[0]
        return {"judge": JUDGE_CHECK, "key": n, "cause": f"OKの変電所が{km(SITE_OK_M)}より遠い（{km(d)}）",
                "reason": f"最も近い接続検討OKの変電所（{n}）まで {km(d)} と{km(SITE_OK_M)}を超える。負担金は高圧線の亘長にほぼ比例するため、"
                          f"既設高圧配電線までの距離を確認するか、より近い候補地を探す"
                          + (f"。{km(SITE_OK_M)}以内には要検討の変電所あり（{checks[0][2]}・{km(checks[0][1])}）" if checks and checks[0][1] <= SITE_OK_M else "")}
    if checks:
        j, d, n = checks[0]
        return {"judge": JUDGE_CHECK, "key": n, "cause": "OKの変電所が無く要検討のみ",
                "reason": f"{km(SITE_RANGE_M)}以内に接続検討OKの変電所が無く、最も近いのは要検討の {n}（{km(d)}）。その変電所の理由（上位系の空き無し等）を見て人が判断"}
    if inrange:
        j, d, n = inrange[0]
        return {"judge": JUDGE_NG, "key": n, "cause": "周辺の変電所がすべて空き無し",
                "reason": f"{km(SITE_RANGE_M)}以内で空容量が公表されている変電所 {len(inrange)}件はすべて接続検討NG（変電所単体の空き無し。最寄りは {n}・{km(d)}）"}
    return {"judge": JUDGE_CHECK, "key": None, "cause": "判定材料なし",
            "reason": f"{km(SITE_RANGE_M)}以内に空容量が公表されている変電所が無く判定材料がない。送配電事業者の公表資料か接続検討で確認"}


def score_row(area, vsec, opcap, flow, avail_cur, avail_up, n1, curtail):
    """1変電所の系統スコアを算出する。xlsxの数式と同じ判定。

    戻り値: {"gate", "judge", "s1a", "s1b", "s1"(=s1a+s1b), "s2".."s6", "grid_score", "rank"}
    """
    eff = effective_capacity(avail_cur, avail_up)   # 上位系考慮を優先した値（ゲート・データ有無の判定用）
    own = own_capacity(avail_cur, avail_up)         # 変電所単体（当該設備）の値（S1a用）
    up = _num(avail_up)
    gate = "警告" if (eff is not None and eff < MIN_MW and n1 != "可") else ""

    if own is None:
        s1a = 0
    elif own >= P["s1_t1"]["thr"]:
        s1a = P["s1_t1"]["pts"]
    elif own >= P["s1_t2"]["thr"]:
        s1a = P["s1_t2"]["pts"]
    elif own >= P["s1_t3"]["thr"]:
        s1a = P["s1_t3"]["pts"]
    else:
        s1a = 0

    if eff is None or own < MIN_MW:   # 変電所単体に空きが無ければ上位系の点は付けない
        s1b = 0
    elif up is None:
        s1b = P["s1b_unk"]["pts"]
    elif up >= P["s1b_ok"]["thr"]:
        s1b = P["s1b_ok"]["pts"]
    else:
        s1b = 0

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

    grid = s1a + s1b + s2 + s3 + s4 + s5 + s6
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

    return {"gate": gate, "judge": judge_row(vsec, avail_cur, avail_up)["judge"],
            "s1a": s1a, "s1b": s1b, "s1": s1a + s1b, "s2": s2, "s3": s3, "s4": s4, "s5": s5, "s6": s6,
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


# 接続検討回答 検証一覧からの取込分（scripts/import_connection_results.py で生成）。
# 手入力の RESULTS_DB の後ろに追加する。ファイルが無ければ手入力分のみ。
IMPORTED_RESULTS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                     "..", "data", "connection_results.json")
RESULTS_MANUAL_COUNT = len(RESULTS_DB)
if os.path.exists(IMPORTED_RESULTS_PATH):
    with open(IMPORTED_RESULTS_PATH, encoding="utf-8") as _f:
        RESULTS_DB = RESULTS_DB + [tuple(r) for r in json.load(_f)["records"]]

# 同じ変電所に複数の回答があるときは要約を「／」でつなぐ
RESULTS_MEMO = {}
for _r in RESULTS_DB:
    _k = results_key(_r[0], _r[1])
    RESULTS_MEMO[_k] = f"{RESULTS_MEMO[_k]} ／ {_r[8]}" if _k in RESULTS_MEMO else _r[8]


# ══════════════════════════════════════════════════════════════════════════════
# 実績による格下げ（2026-11-06 余野さん判断。閾値はここだけを直す）
#
#   直近 RESULT_NG_WINDOW_DAYS 日以内の接続検討の回答で、
#     連系不可 ／ 負担金 ≧ RESULT_NG_YEN_MAN 万円 ／ 連系工期 ≧ RESULT_NG_MONTHS ヶ月
#   のいずれかに当たる回答を「実績NG」と呼ぶ。
#
#   (a) 変電所名が確定している実績NGがある変電所 → 判定を「接続検討NG」に落とす（公表の空容量が空きありでも）。
#       回答書に変電所名が無く住所から推定した実績（「変電所は推定」）は、別の変電所の回答かもしれないので (a) には使わない。
#   (b) 実績NGの回答が上位系の設備（例: 岩手(変)275/154kV主要変圧器、金城変電所110/66kV、新岡山(変)）を名指ししているとき、
#       同じ設備を「上位系の制約設備」として持つ変電所（系統トポロジー data/grid_topology.js の upRefs）にも
#       「⚠上位系に実績NG」の警告を付け、判定が接続検討OKなら「要検討」に落とす（NG・データ無はそのまま）。
#       上位設備の名前は回答書に明記されているので、(b) は推定の実績も使う。回答を受けた変電所自身の名前は上位設備に数えない。
#   公表データだけの判定（judge_row）は judgePub として残し、画面・CSV・xlsxには両方と理由を出す。
#   実績DB（connection_results.json）が無い環境（実績を公開リポジトリから外した後の公開版）では何も格下げしない。
# ══════════════════════════════════════════════════════════════════════════════
import datetime as _dt
import re as _re

TOPOLOGY_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "grid_topology.js")


def _parse_yen_man(text):
    """負担金の文字列から万円単位の金額（最大値）を返す。「4.64億円」「3,246万円」「約3,700〜3,900万円」。無ければ None。"""
    vals = []
    for m in _re.finditer(r"([\d,]+(?:\.\d+)?)\s*億円", text or ""):
        vals.append(float(m.group(1).replace(",", "")) * 10000)
    for m in _re.finditer(r"([\d,]+(?:\.\d+)?)\s*万円", text or ""):
        vals.append(float(m.group(1).replace(",", "")))
    return max(vals) if vals else None


def _parse_months(text):
    """連系工期の文字列から月数（複数あれば最大）を返す。「入金後1年4ヶ月」「配電7ヶ月/上位系10年2ヶ月」「記載なし」→ None。"""
    vals = []
    body = _re.sub(r"[（(][^）)]*[）)]", "", text or "")   # 「(受付後73ヶ月)」のような括弧内の補足は見ない
    for seg in _re.split(r"[/／]", body):
        m = _re.search(r"(?:(\d+)\s*年)?\s*(\d+)\s*[ヶかカケ]月", seg)
        if m:
            vals.append(int(m.group(1) or 0) * 12 + int(m.group(2)))
        else:
            m = _re.search(r"(\d+)\s*年(?![\d])", seg)
            if m:
                vals.append(int(m.group(1)) * 12)
    return max(vals) if vals else None


def _parse_date(text):
    """「(2026-04-24)」「(回答2026/5/13)」のような回答日。複数あれば最後のもの。無ければ None。"""
    ds = _re.findall(r"(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})", text or "")
    if not ds:
        return None
    y, mo, d = ds[-1]
    try:
        return _dt.date(int(y), int(mo), int(d))
    except ValueError:
        return None


def norm_sub_name(name):
    """変電所名の表記ゆれを吸収する（「岩手(変)」「岩手変電所」「岩手」→「岩手」）。"""
    s = str(name or "")
    s = _re.sub(r"[\s　]", "", s)
    s = _re.sub(r"[（(]変[)）]$", "", s)
    s = _re.sub(r"変電所$", "", s)
    s = _re.sub(r"[（(]変[)）]", "", s)
    return s


def _upper_names(record):
    """回答の「主な対策工事」「教訓・律速」から上位系の設備として名指しされた変電所名（正規化済み）を返す。"""
    text = " ".join(str(x or "") for x in (record[6], record[7]))
    # 「上位系」以降の記述を優先（無ければ全文）
    m = _re.search(r"上位系[^:：]*[:：](.*)$", text)
    body = m.group(1) if m else text
    names = set()
    for n in _re.findall(r"([一-龥ァ-ヶー]{1,8})[（(]変[)）]", body):
        names.add(norm_sub_name(n))
    for n in _re.findall(r"([一-龥ァ-ヶー]{1,8}変電所)", body):
        names.add(norm_sub_name(n))
    own = norm_sub_name(_re.sub(r"^－\(|\)$", "", str(record[1] or "")))
    names.discard(own)
    # 固有名でない語（「配電用変電所の増強」「上位変電所」など）は設備名に数えない
    names -= {"", "配電用", "上位", "基幹", "当該", "連系", "既設", "新設", "他社", "自社", "最寄りの", "最寄り"}
    return sorted(names)


def parse_result(record):
    """実績DBの1件を解析する。戻り値: dict(area, name, est, date, yen_man, months, no_connect, upper, ng, why)。"""
    area, name = record[0], str(record[1] or "").strip()
    memo = str(record[8] or "")
    text78 = " ".join(str(x or "") for x in (record[7], record[8]))
    est = "変電所は推定" in text78
    date = _parse_date(memo) or _parse_date(str(record[7] or ""))
    yen = _parse_yen_man(str(record[4] or "")) or _parse_yen_man(memo)
    months = _parse_months(str(record[5] or ""))
    if months is None:
        months = _parse_months(memo)
    no_connect = "連系不可" in text78
    why = []
    if no_connect:
        why.append("連系不可")
    if yen is not None and yen >= RESULT_NG_YEN_MAN:
        why.append(f"負担金 {fmt_yen_man(yen)}")
    if months is not None and months >= RESULT_NG_MONTHS:
        why.append(f"工期 {months}ヶ月")
    return {"area": area, "name": name, "est": est, "date": date, "yen_man": yen, "months": months,
            "no_connect": no_connect, "upper": _upper_names(record), "ng": bool(why), "why": why,
            "case": str(record[2] or ""), "kw": str(record[3] or "")}


def result_ng_records(as_of=None):
    """直近 RESULT_NG_WINDOW_DAYS 日以内の実績NG（parse_result の dict のリスト）。回答日が無い実績は除く。"""
    as_of = as_of or _dt.date.today()
    out = []
    for rec in RESULTS_DB:
        p = parse_result(rec)
        if not p["ng"] or p["date"] is None:
            continue
        if (as_of - p["date"]).days > RESULT_NG_WINDOW_DAYS or p["date"] > as_of:
            continue
        out.append(p)
    return out


def _fmt_ng(p):
    # 案件名・住所は出さない（回答書の目的外使用・第三者開示の禁止に配慮）。出力だけ添える
    return f"{p['date'].isoformat()} {'・'.join(p['why'])}" + (f"（{p['kw']}）" if p["kw"] else "")


class ResultAdjuster:
    """実績NGによる判定の格下げを、変電所ごとに引き当てるための索引。"""

    def __init__(self, as_of=None, topology_path=TOPOLOGY_PATH):
        self.as_of = as_of or _dt.date.today()
        self.ngs = result_ng_records(self.as_of)
        # (a) 変電所名が確定している実績NG → キー（エリア＋名前）ごとの理由
        self.direct = {}
        # (b) 名指しされた上位設備 → (エリア, 正規化名) ごとの理由
        self.upper = {}
        for p in self.ngs:
            if not p["est"] and p["name"] and not p["name"].startswith("－"):
                self.direct.setdefault(results_key(p["area"], p["name"]), []).append(p)
            for u in p["upper"]:
                self.upper.setdefault((p["area"], u), []).append(p)
        # 系統トポロジー: 変電所No → 上位系の制約設備として参照するノード名（正規化）の集合
        self.up_names_by_no = {}
        self.topo_ok = False
        if os.path.exists(topology_path):
            try:
                self._load_topology(topology_path)
                self.topo_ok = True
            except Exception:  # noqa: BLE001  トポロジーが読めなくても (a) は動かす
                self.topo_ok = False

    def _load_topology(self, path):
        raw = open(path, encoding="utf-8").read()
        t = json.loads(raw[raw.index("{"): raw.rindex("}") + 1])
        NF = {n: i for i, n in enumerate(t["nodeFields"])}
        BF = {n: i for i, n in enumerate(t["bankFields"])}
        LF = {n: i for i, n in enumerate(t["lineFields"])}
        areas = t["meta"]["areas"]
        nodes, lines = t["nodes"], t["lines"]
        node_lines = {}
        for li, l in enumerate(lines):
            for e in (l[LF["from"]], l[LF["to"]]):
                if isinstance(e, int):
                    node_lines.setdefault(e, []).append(li)

        def node_name(ni):
            return norm_sub_name(nodes[ni][NF["name"]])

        for ni, node in enumerate(nodes):
            ups = set()
            for b in node[NF["banks"]]:
                for ref in b[BF["upRefs"]] or []:
                    if ref[0] == "N":
                        ups.add(ref[1])
                    elif ref[0] == "L":
                        l = lines[ref[1]]
                        for e in (l[LF["from"]], l[LF["to"]]):
                            if isinstance(e, int):
                                ups.add(e)
            for li in node_lines.get(ni, []):
                for ref in lines[li][LF["upRefs"]] or []:
                    if ref[0] == "N":
                        ups.add(ref[1])
            ups.discard(ni)
            names = {node_name(u) for u in ups}
            area = areas[node[NF["area"]]] if isinstance(node[NF["area"]], int) else node[NF["area"]]
            for no in node[NF["nos"]]:
                self.up_names_by_no[str(no)] = (area, names)

    def upper_hits(self, area, no):
        """この変電所の上位系の制約設備のうち、実績NGで名指しされたもの → [(設備名, [実績...]), ...]"""
        ent = self.up_names_by_no.get(str(no)) if no is not None else None
        if not ent:
            return []
        hits = []
        for nm in sorted(ent[1]):
            ps = self.upper.get((area, nm))
            if ps:
                hits.append((nm, ps))
        return hits

    def adjust(self, area, name, no, judge_pub):
        """公表データの判定 judge_pub に実績NGを反映する。

        戻り値: {"judge": 最終判定, "resultNg": 理由（(a)に該当すれば文字列、無ければ ""）,
                 "upperWarn": 理由（(b)に該当すれば文字列、無ければ ""）}
        """
        direct = self.direct.get(results_key(area, name), [])
        hits = self.upper_hits(area, no)
        result_ng = ""
        upper_warn = ""
        judge = judge_pub
        if direct:
            result_ng = (f"{RESULT_NG_TAG}: 直近{RESULT_NG_WINDOW_DAYS // 365}年以内の接続検討で "
                         + "／".join(_fmt_ng(p) for p in sorted(direct, key=lambda p: p["date"], reverse=True))
                         + f"。基準: 連系不可・負担金{RESULT_NG_YEN_TXT}以上・工期{RESULT_NG_MONTHS}ヶ月以上のいずれか")
            judge = JUDGE_NG
        if hits:
            parts = []
            for nm, ps in hits:
                p0 = sorted(ps, key=lambda p: p["date"], reverse=True)[0]
                parts.append(f"{nm}（{p0['date'].isoformat()} {p0['name'] or '変電所名なし'}: {'・'.join(p0['why'])}"
                             + (f" ほか{len(ps) - 1}件" if len(ps) > 1 else "") + "）")
            upper_warn = (f"{UPPER_WARN_TAG}: この変電所の上位系の制約設備 " + "、".join(parts)
                          + " が直近の接続検討で増強の対象になった。同じ上位問題で高額・長工期の回答になる可能性があるため、"
                          + ("接続検討OKから要検討に格下げ" if judge_pub == JUDGE_OK and not direct else "注意して判断"))
            if judge == JUDGE_OK:
                judge = JUDGE_CHECK
        return {"judge": judge, "resultNg": result_ng, "upperWarn": upper_warn}


_ADJUSTER = None


def result_adjuster(as_of=None):
    """ResultAdjuster の共有インスタンス（生成が重いので1回だけ作る）。"""
    global _ADJUSTER
    if _ADJUSTER is None or (as_of is not None and _ADJUSTER.as_of != as_of):
        _ADJUSTER = ResultAdjuster(as_of)
    return _ADJUSTER


RESULT_NG_MEANING = {
    RESULT_NG_TAG: (f"実績NG：直近{RESULT_NG_WINDOW_DAYS // 365}年以内の接続検討の回答（変電所名が確定しているもの）が "
                    f"連系不可・負担金{RESULT_NG_YEN_TXT}以上・工期{RESULT_NG_MONTHS}ヶ月以上のいずれか。"
                    "公表の空容量が空きありでも判定を接続検討NGに落とす（住所から推定した実績は使わない）"),
    UPPER_WARN_TAG: ("上位系に実績NG：実績NGの回答が名指しした上位系の設備（上位変電所の変圧器など）を、この変電所も上位系の制約設備として持つ。"
                     "同じ上位問題で高額・長工期になる可能性があるため、接続検討OKは要検討に格下げ（要検討・NG・データ無はそのまま）"),
}
