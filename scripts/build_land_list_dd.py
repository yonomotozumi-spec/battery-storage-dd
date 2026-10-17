#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
土地一覧（案件候補地リスト）に、変電所スコア・条例リスク・DD許認可の確認結果を
突き合わせて「確認用の一覧xlsx」を作る。

ネットワークはしない。緯度経度・自治体名は呼び出し側が解決して geo.json で渡す
（GoogleマップのピンURL解決／国土地理院の逆ジオコーダ）。reinfolib の自動判定は
scripts/reinfolib_judge.py の出力 values.json を --values-dir <No>.json で受け取る。

  python scripts/build_land_list_dd.py \
    --in "土地一覧.xlsx" --sheet Sheet1 --geo geo.json \
    [--values-dir reinfolib_values] --out "土地一覧_DD確認.xlsx"
"""
import argparse, json, math, os, re, urllib.parse, datetime
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 見た目（既存の案件チェックシートに合わせる）──
HDR_FILL = PatternFill("solid", fgColor="FCE4D6")
SUB_FILL = PatternFill("solid", fgColor="FFF2CC")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
RISK_FILL = PatternFill("solid", fgColor="FCE4EC")
GREY_FILL = PatternFill("solid", fgColor="F2F2F2")
HDR_FONT = Font(bold=True, size=11)
SMALL = Font(size=10)
LINK_FONT = Font(color="1155CC", underline="single", size=10)
THIN = Side(style="thin", color="D0D0D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

# ── 確認リンク（すべて緯度経度で開く）──
def gmaps(lat, lon): return f"https://www.google.com/maps?q={lat},{lon}"
def gsi_air(lat, lon): return f"https://maps.gsi.go.jp/#18/{lat}/{lon}/&base=ort&ls=ort&disp=1&vs=c1g1j0h0k0l0u0t0z0r0s0m0f1"
FLOOD_LS = "flood_l2_kaokutoukai_kagan,0.8%7Cflood_l2_kaokutoukai_hanran,0.8%7Cflood_l2_keizoku,0.8%7Cflood_list,0.8%7Cflood_l1,0.8%7Cflood_list_l2,0.75"
DOSHA_LS = "dosha_kiken_nadare,0.8%7Cdosha_keikai_jisuberi,0.8%7Cdosha_keikai_dosekiryu,0.8%7Cdosha_keikai_kyukeisha,0.8"
EKI_LS = "ekijouka_zenkoku,0.8"
def haz(lat, lon, ls): return f"https://disaportal.gsi.go.jp/hazardmap/maps/index.html?ll={lat},{lon}&z=16&ls={ls}"
def hazall(lat, lon): return f"https://disaportal.gsi.go.jp/hazardmap/maps/index.html?ll={lat},{lon}&z=16&base=pale"
def search(q): return "https://www.google.com/search?q=" + urllib.parse.quote(q)

def dms(dec, is_lat):
    hemi = ("N" if dec >= 0 else "S") if is_lat else ("E" if dec >= 0 else "W")
    dec = abs(dec); d = int(dec); mf = (dec - d) * 60; m = int(mf); s = round((mf - m) * 60, 1)
    return f"{d}°{m}'{s}\"{hemi}"

def distance_m(lat1, lon1, lat2, lon2):
    """Haversine（substation.html と同じ地球半径）"""
    R = 6371008.8; rad = math.pi / 180
    dlat = (lat2 - lat1) * rad; dlon = (lon2 - lon1) * rad
    a = math.sin(dlat / 2) ** 2 + math.cos(lat1 * rad) * math.cos(lat2 * rad) * math.sin(dlon / 2) ** 2
    return 2 * R * math.asin(min(1, math.sqrt(a)))

def fmt_dist(m):
    return f"{round(m)} m" if m < 1000 else f"{m/1000:.2f} km" if m < 10000 else f"{m/1000:.1f} km"

# ── 変電所データ（data/substations.js）──
def load_substations(path):
    s = open(path, encoding="utf-8").read()
    i = s.index("window.SUBSTATION_DATA=")
    obj = json.loads(s[i + len("window.SUBSTATION_DATA="):].rstrip().rstrip(";"))
    fields = obj["fields"]; idx = {f: n for n, f in enumerate(fields)}
    return obj, idx

def s7_points(m, steps):
    for limit, pts in steps:
        if m <= limit:
            return pts
    return 0

def total_rank(total, steps):
    for thr, rk in steps:
        if total >= thr:
            return rk
    return "C"

# ── DD 18項目（build_px_checksheet.py と同じ並び・同じ確認先）──
def dd_items(lat, lon, address, muni, pref):
    m = muni or ""; p = pref or ""
    coord = f"{dms(lat, True)} {dms(lon, False)}（{lat}, {lon}）"
    return [
        ("1", "住所", address, gmaps(lat, lon)),
        ("2", "座標", coord, gmaps(lat, lon)),
        ("3", "架線の敷地内縦横断", "", gsi_air(lat, lon)),
        ("4", "ハザード・浸水", "", haz(lat, lon, FLOOD_LS)),
        ("5", "ハザード・土砂災害", "", haz(lat, lon, DOSHA_LS)),
        ("6", "液状化危険度", "", haz(lat, lon, EKI_LS)),
        ("7", "垂直積雪量", "", search(f"{p} {m} 垂直積雪量 建築基準法施行細則")),
        ("8", "市街化区域", "", search(f"{m} 都市計画情報 区域区分 市街化区域 マップ")),
        ("9", "市街化調整区域", "", search(f"{m} 都市計画情報 市街化調整区域 マップ")),
        ("10", "用途地域", "", search(f"{m} 用途地域 都市計画情報 マップ")),
        ("11", "農地該当", "", gmaps(lat, lon)),
        ("12", "森林", "", search(f"{p} {m} 森林情報 GIS 地域森林計画対象民有林")),
        ("13", "公園", "", search(f"{p} {m} 自然公園 区域 都市公園 区域図")),
        ("14", "鳥獣保護区", "", search(f"{p} 鳥獣保護区 位置図")),
        ("15", "埋蔵文化財包蔵地", "", search(f"{p} {m} 遺跡地図 埋蔵文化財 包蔵地 GIS")),
        ("16", "景観区域", "", search(f"{m} 景観計画区域 図")),
        ("17", "騒音規制", "", search(f"{m} 騒音規制法 地域指定 規制基準")),
        ("18", "振動規制", "", search(f"{m} 振動規制法 地域指定 規制基準")),
    ]

AUTO_ITEMS = {"4", "5", "6", "8", "9", "10", "13"}   # reinfolib で自動判定できる項目

# reinfolib の判定から「要」が立ちうる許認可（build_px_checksheet.py の PERMITS と同番号）
PERMIT_NAMES = {
    "7": "宅地造成及び特定盛土等規制法（宅地造成等工事許可）",
    "13": "地すべり等防止法（防止区域内の許認可）",
    "14": "急傾斜地崩壊防止法（危険区域内の許認可）",
    "15": "土砂災害防止法（特別警戒区域の特定開発許可）",
    "16": "土砂災害危険箇所（区域確認）",
}

def judge_dd(no, value):
    """値から 良/注意/リスク を決める。値が無ければ要確認。"""
    if not value:
        return "要確認"
    v = str(value)
    if no == "4":
        return "良" if "該当なし" in v else "リスク"
    if no == "5":
        if "該当なし" in v: return "良"
        return "リスク" if "特別警戒" in v else "注意"
    if no == "6":
        return "良" if "該当なし" in v else "注意"
    if no == "8":
        return "良" if "指定あり" in v else "注意"
    if no == "9":
        return "リスク" if "指定あり" in v else "良"
    if no == "10":
        if "指定なし" in v: return "良"
        return "リスク" if ("低層住居専用" in v or "中高層住居専用" in v or "住居" in v) else "注意"
    if no == "13":
        return "良" if "該当なし" in v else "リスク"
    return "記入済"

KANJI_NUM = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5", "六": "6", "七": "7",
             "八": "8", "九": "9", "十": "10", "十一": "11", "十二": "12"}

def norm_addr(t):
    """丁目の漢数字・全角数字・ケ/ヶ の表記ゆれを吸収して照合できる形にする。"""
    t = str(t or "")
    t = t.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    t = t.replace("ケ", "ヶ").replace("ガ", "ヶ")
    for k in sorted(KANJI_NUM, key=len, reverse=True):
        t = t.replace(k + "丁目", KANJI_NUM[k] + "丁目")
    return t


YOUTO_ABBR = [
    ("1低専", "第１種低層住居専用"), ("2低専", "第２種低層住居専用"),
    ("1中専", "第１種中高層住居専用"), ("2中専", "第２種中高層住居専用"),
    ("1住居", "第１種住居"), ("2住居", "第２種住居"), ("準住居", "準住居"),
    ("田園住居", "田園住居"), ("近商", "近隣商業"), ("商業", "商業"),
    ("準工", "準工業"), ("工専", "工業専用"), ("工業", "工業"),
]

def youto_key(t):
    """『1低専（50/80）』と『第１種低層住居専用地域（…）』を同じキーに寄せる。"""
    t = str(t or "").replace("第1種", "第１種").replace("第2種", "第２種")
    for abbr, full in YOUTO_ABBR:
        if abbr in t or full in t:
            return full
    return ""


def compare_youto(sheet_val, api_val):
    """元表の用途地域と reinfolib 判定の一致を返す。"""
    if not api_val:
        return ""
    if "指定なし" in str(api_val):
        return "APIでは用途地域の指定なし（元表も記載なし）" if not str(sheet_val).strip() else \
               f"要確認（元表:{sheet_val} / API:指定なし）"
    if not str(sheet_val).strip():
        return f"元表に記載なし（API:{youto_key(api_val) or api_val}）"
    a, b = youto_key(sheet_val), youto_key(api_val)
    if a and b and a == b:
        return "○ 一致"
    return f"要確認（元表:{sheet_val} / API:{b or api_val}）"


def norm_muni(muni):
    """仙台市青葉区 → 仙台市（開発許可権者・条例台帳の突き合わせ用）"""
    m = re.match(r"^(.+?市)(.+区)$", muni or "")
    return m.group(1) if m else (muni or "")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="src", required=True)
    ap.add_argument("--sheet", default="Sheet1")
    ap.add_argument("--geo", required=True)
    ap.add_argument("--subs", default="data/substations.js")
    ap.add_argument("--ordinance", default="data/ordinance_registry.json")
    ap.add_argument("--authorities", default="data/kaihatsu_kyokasha.json")
    ap.add_argument("--priority", default="data/priority_munis.json")
    ap.add_argument("--values-dir", default=None)
    ap.add_argument("--title", default="笹川総研様土地一覧")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()

    geo = {str(k): v for k, v in json.load(open(a.geo, encoding="utf-8")).items()}
    sub, IX = load_substations(a.subs)
    ordn = json.load(open(a.ordinance, encoding="utf-8"))
    auth = json.load(open(a.authorities, encoding="utf-8"))["authorities"]
    prio = json.load(open(a.priority, encoding="utf-8"))["munis"]
    prio_key = {(m["pref"], m["muni"]): m for m in prio}

    dict_area = sub["dict"]["area"]; dict_acc = sub["dict"]["acc"]
    dict_n1 = sub["dict"]["n1"]; dict_curtail = sub["dict"]["curtail"]; dict_rank = sub["dict"]["rank"]
    s7steps = sub["params"]["s7Steps"]; trsteps = sub["params"]["totalRankSteps"]

    # ── 元シートを読む（数式セルは計算済み値で読み直す）──
    wb_in = openpyxl.load_workbook(a.src, data_only=False)
    ws_in = wb_in[a.sheet]
    rows = []
    for r in ws_in.iter_rows(min_row=4):
        v = {c.column_letter: c.value for c in r if c.value is not None}
        if not v.get("B"):
            continue
        link = None
        for c in r:
            if c.hyperlink:
                link = c.hyperlink.target
        rows.append((v, link))

    def area_m2(v):
        """地積セル。=1066.4+463.9 のような数式は足し合わせて数値にする。"""
        x = v.get("J")
        if isinstance(x, (int, float)):
            return float(x)
        if isinstance(x, str) and x.startswith("="):
            nums = re.findall(r"\d+(?:\.\d+)?", x)
            return sum(float(n) for n in nums) if nums else None
        return None

    sites = []
    for n, (v, link) in enumerate(rows, start=1):
        g = geo.get(str(n), {})
        lat, lon = g.get("lat"), g.get("lon")
        muni_full = g.get("muni", ""); pref = g.get("pref", "")
        sites.append({
            "no": n, "label": str(v.get("B")).replace(".0", ""),
            "area_power": v.get("C", ""), "scale": v.get("D", ""), "impression": v.get("E", ""),
            "address": v.get("F", ""), "maplink": link or v.get("G", ""),
            "chimoku": v.get("H", ""), "genkyo": v.get("I", ""), "m2": area_m2(v),
            "memo1": v.get("P", "") or "", "memo2": v.get("Q", "") or "",
            "toshi": v.get("R", "") or "", "youto": v.get("S", "") or "",
            "kodo": v.get("T", "") or "", "kotai": v.get("U", "") or "",
            "takasa": v.get("V", "") or "", "chiku": v.get("W", "") or "",
            "road_w": v.get("X", ""), "road_c": v.get("Y", ""),
            "lat": lat, "lon": lon, "pref": pref, "muni_full": muni_full,
            "muni": norm_muni(muni_full), "lv01": g.get("lv01", ""),
        })

    # ── 変電所（半径10km以内・総合スコア順に上位5件）──
    for s in sites:
        s["subs"] = []
        if s["lat"] is None:
            continue
        hits = []
        for row in sub["rows"]:
            la, lo = row[IX["lat"]], row[IX["lon"]]
            if la is None or lo is None:
                continue
            d = distance_m(s["lat"], s["lon"], la, lo)
            if d > 10000:
                continue
            s7 = s7_points(d, s7steps); grid = row[IX["grid"]] or 0
            hits.append({"row": row, "dist": d, "s7": s7, "grid": grid, "total": grid + s7})
        hits.sort(key=lambda h: (-h["total"], h["dist"]))
        s["subs"] = hits[:5]
        s["nearest"] = min(hits, key=lambda h: h["dist"]) if hits else None
        s["n_in_10km"] = len(hits)
        s["n_sa"] = sum(1 for h in hits if dict_rank[h["row"][IX["rank"]]] in ("S", "A"))

    # ── reinfolib 自動判定（あれば取り込む）──
    for s in sites:
        s["values"] = {}
        s["permits_auto"] = {}
        if a.values_dir:
            p = os.path.join(a.values_dir, f"{s['no']}.json")
            if os.path.exists(p):
                d = json.load(open(p, encoding="utf-8"))
                s["values"] = d.get("values", {})
                s["permits_auto"] = d.get("permits", {})

    # ── 条例・開発許可 ──
    entries = ordn["entries"]
    national = [e for e in entries if e["level"] == "国"]
    for s in sites:
        matched = []
        for e in entries:
            if e["level"] == "都道府県" and e["pref"] == s["pref"]:
                matched.append(e)
            elif e["level"] == "市区町村" and e["pref"] == s["pref"] and e["muni"] == s["muni"]:
                matched.append(e)
        s["ord_matched"] = matched
        # 開発許可権者：市が権者ならその市、そうでなければ都道府県
        au_muni = next((x for x in auth if x["pref"] == s["pref"] and x["name"] == s["muni"]), None)
        au_pref = next((x for x in auth if x["type"] == "都道府県" and x["name"] == s["pref"]), None)
        s["authority"] = au_muni or au_pref
        s["authority_kind"] = ("市区町村（" + au_muni["type"] + "）") if au_muni else "都道府県"
        s["priority"] = prio_key.get((s["pref"], s["muni"]))
        # リスク判定
        if any(e["level"] == "市区町村" for e in matched):
            lv, why = "高", "所在市区町村が系統用蓄電池の取扱いを公表済み"
        elif any(e["level"] == "都道府県" for e in matched):
            lv, why = "中", "都道府県が取扱いを公表済み（市区町村は未公表）"
        elif s["authority"] and s["authority"].get("status") == "公表あり":
            lv, why = "中", "開発許可権者が公表あり（台帳未収録・内容要確認）"
        else:
            lv, why = "低（要照会）", "公表確認できず。国都計第7号の運用は窓口照会が必要"
        chosei = "調整区域" in (s["toshi"] or "") or \
                 "指定あり" in ((s["values"].get("9") or {}).get("value") or "")
        if chosei:
            lv = "高"; why += " ／ 市街化調整区域のため開発許可の要否が論点（国都計第7号）"
        s["ord_level"] = lv; s["ord_why"] = why

    for s in sites:
        s["dd"] = []
        if s["lat"] is None:
            continue
        for no, name, val, url in dd_items(s["lat"], s["lon"], s["address"], s["muni_full"], s["pref"]):
            got = s["values"].get(no)
            value = (got or {}).get("value") or val
            comment = (got or {}).get("comment") or ""
            s["dd"].append({"no": no, "name": name, "value": value, "url": url,
                            "comment": comment, "judge": judge_dd(no, value),
                            "auto": no in AUTO_ITEMS})

    # ══════════ 出力 ══════════
    wb = Workbook()
    today = datetime.date.today().isoformat()

    def head(ws, headers, widths, title, note):
        ws["A1"] = title; ws["A1"].font = Font(bold=True, size=13)
        ws["A2"] = note; ws["A2"].font = Font(size=9, color="666666")
        for i, h in enumerate(headers, start=1):
            c = ws.cell(row=3, column=i, value=h)
            c.fill = HDR_FILL; c.font = HDR_FONT; c.alignment = CENTER; c.border = BORDER
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[3].height = 34
        ws.freeze_panes = "C4"

    def put(ws, r, c, v, *, link=None, fill=None, font=None, align=None):
        cell = ws.cell(row=r, column=c, value=v)
        cell.border = BORDER; cell.font = font or SMALL
        cell.alignment = align or WRAP
        if link:
            cell.hyperlink = link; cell.font = LINK_FONT
        if fill:
            cell.fill = fill
        return cell

    # ── ① 一覧（DD総括）──
    ws = wb.active; ws.title = "①一覧（DD総括）"
    H = ["No.", "電力\nエリア", "想定\n規模", "所感\n(元表)", "所在（元表）", "都道府県", "市区町村",
         "緯度", "経度", "座標の\n所在確認", "地目", "現況", "地積\n(㎡)", "地積\n(坪)",
         "都市計画\n(元表)", "用途地域\n(元表)", "用途地域\n(API判定)", "元表との照合", "最有力変電所\n(総合スコア最大)", "その距離", "系統\nランク", "系統\n(80)",
         "近接\n(20)", "総合\n(100)", "総合\nランク", "最寄変電所", "最寄\n距離", "10km内\n変電所", "うち\nS/A", "条例\nリスク",
         "条例リスクの理由", "該当条例\n(件)", "開発許可権者", "権者の\n公表状況", "DD\n要確認", "DD\nリスク", "自動判定で「要」になった許認可",
         "Googleマップ", "ハザードマップ", "航空写真(GSI)", "元表の地図リンク"]
    W = [7, 7, 7, 9, 42, 10, 13, 10, 10, 13, 12, 14, 9, 8, 12, 16, 26, 26, 22, 9, 7, 7, 7, 8, 7, 22, 9, 8, 7, 9,
         40, 8, 16, 12, 8, 8, 42, 12, 14, 13, 14]
    head(ws, H, W, f"{a.title}｜DD・変電所・条例 確認一覧",
         f"作成 {today}／座標は元表のGoogleマップ・ピンURLを解決した実測値。変電所は data/substations.js（66kV以上6,071件・系統スコア80点＋近接S7 20点）、"
         f"条例は data/ordinance_registry.json（{ordn['updated']}時点30件）と開発許可権者151者のスイープ結果に基づく。")
    r = 4
    for s in sites:
        best = s["subs"][0] if s["subs"] else None
        row = best["row"] if best else None
        nrow = s["nearest"]["row"] if s.get("nearest") else None
        rk = dict_rank[row[IX["rank"]]] if row else ""
        trk = total_rank(best["total"], trsteps) if best else ""
        dd_need = sum(1 for d in s["dd"] if d["judge"] == "要確認")
        dd_risk = sum(1 for d in s["dd"] if d["judge"] == "リスク")
        match_ok = "○" if (s["lv01"] and norm_addr(s["lv01"]) in norm_addr(s["address"])) else "要確認"
        vals = [s["label"], s["area_power"], s["scale"], s["impression"], s["address"], s["pref"], s["muni_full"],
                s["lat"], s["lon"], f"{match_ok}\n{s['lv01']}", s["chimoku"], s["genkyo"],
                s["m2"], round(s["m2"] / 3.30578, 2) if s["m2"] else None,
                s["toshi"], s["youto"],
                (s["values"].get("10") or {}).get("value") or "（APIキー未設定）",
                compare_youto(s["youto"], (s["values"].get("10") or {}).get("value")),
                (str(row[IX["name"]]) + f"\n({dict_area[row[IX['area']]]})") if row else "半径10km内になし",
                fmt_dist(best["dist"]) if best else "", rk, best["grid"] if best else "",
                best["s7"] if best else "", best["total"] if best else "", trk,
                (str(nrow[IX["name"]]) + f"\n({dict_rank[nrow[IX['rank']]]})") if nrow is not None else "",
                fmt_dist(s["nearest"]["dist"]) if s.get("nearest") else "",
                s.get("n_in_10km", 0), s.get("n_sa", 0), s["ord_level"], s["ord_why"], len(s["ord_matched"]),
                (s["authority"]["name"] + f"\n({s['authority_kind']})") if s["authority"] else "要確認",
                s["authority"]["status"] if s["authority"] else "", dd_need, dd_risk,
                "\n".join(f"No.{k} {PERMIT_NAMES.get(k, '')}" for k in sorted(s["permits_auto"], key=int)) or "（自動判定では該当なし）"]
        for i, v in enumerate(vals, start=1):
            f = None
            if i == 30:  # 条例リスク
                f = RISK_FILL if s["ord_level"] == "高" else (WARN_FILL if s["ord_level"].startswith("中") else OK_FILL)
            if i == 25 and trk:  # 総合ランク
                f = OK_FILL if trk in ("S", "A") else (WARN_FILL if trk == "B" else RISK_FILL)
            if i == 10 and match_ok != "○":
                f = WARN_FILL
            if i == 18 and str(v).startswith("要確認"):
                f = WARN_FILL
            if i == 37 and s["permits_auto"]:
                f = RISK_FILL
            put(ws, r, i, v, fill=f, align=CENTER if i in (1, 2, 3, 4, 10, 21, 22, 23, 24, 25, 27, 28, 29, 30, 32, 34, 35, 36) else None)
        if s["lat"] is not None:
            put(ws, r, 38, "Googleマップ", link=gmaps(s["lat"], s["lon"]))
            put(ws, r, 39, "重ねるハザード", link=hazall(s["lat"], s["lon"]))
            put(ws, r, 40, "GSI航空写真", link=gsi_air(s["lat"], s["lon"]))
        if s["maplink"]:
            put(ws, r, 41, "元表のピン", link=str(s["maplink"]))
        ws.row_dimensions[r].height = 46
        r += 1
    ws.auto_filter.ref = f"A3:{get_column_letter(len(H))}{r-1}"

    # ── ② 変電所（最寄5件）──
    ws2 = wb.create_sheet("②変電所_最寄5件")
    H2 = ["No.", "候補地", "順位", "変電所名", "エリア", "所在(県)", "距離", "座標精度",
          "最高電圧\n(kV)", "二次電圧\n(kV)", "運用容量\n(MW)", "予想潮流\n(MW)", "空容量\n(現状MW)",
          "空容量\n(上位系考慮MW)", "N-1電制", "出力制御", "系統\n(80)", "近接S7\n(20)", "総合\n(100)",
          "系統\nランク", "総合\nランク", "変電所所在地", "地図"]
    W2 = [6, 30, 6, 26, 8, 10, 10, 12, 9, 9, 10, 10, 10, 12, 14, 10, 7, 8, 8, 7, 7, 34, 10]
    head(ws2, H2, W2, "候補地ごとの最寄変電所（半径10km・総合スコア順 上位5件）",
         "総合＝系統スコア80点（空容量・N-1電制・潮流余裕・出力制御・配電用変電所・エリア）＋近接S7 20点（≦500m:20 ／ ≦1km:15 ／ ≦2km:10 ／ ≦5km:4 ／ 超:0）。"
         "近接は直線距離。実際の負担金は既設高圧配電線までの亘長で決まるため目安。")
    r = 4
    for s in sites:
        if not s["subs"]:
            put(ws2, r, 1, s["label"], align=CENTER)
            put(ws2, r, 2, s["address"])
            put(ws2, r, 4, "半径10km内に66kV以上の変電所なし（データ上）", fill=WARN_FILL)
            r += 1
            continue
        for i, h in enumerate(s["subs"], start=1):
            row = h["row"]; rk = dict_rank[row[IX["rank"]]]
            trk = total_rank(h["total"], trsteps)
            vals = [s["label"] if i == 1 else "", s["address"] if i == 1 else "", i,
                    row[IX["name"]], dict_area[row[IX["area"]]], row[IX["pref"]], fmt_dist(h["dist"]),
                    dict_acc[row[IX["acc"]]], row[IX["vmax"]], row[IX["vsec"]], row[IX["opcap"]],
                    row[IX["flow"]], row[IX["availCur"]], row[IX["availUp"]],
                    dict_n1[row[IX["n1"]]], dict_curtail[row[IX["curtail"]]],
                    h["grid"], h["s7"], h["total"], rk, trk, row[IX["addr"]]]
            for j, v in enumerate(vals, start=1):
                f = None
                if j == 20:
                    f = OK_FILL if rk in ("S", "A") else (WARN_FILL if rk == "B" else (RISK_FILL if rk == "除外" else GREY_FILL))
                if j == 21:
                    f = OK_FILL if trk in ("S", "A") else (WARN_FILL if trk == "B" else RISK_FILL)
                put(ws2, r, j, v, fill=f, align=CENTER if j in (3, 5, 7, 9, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 21) else None)
            put(ws2, r, 23, "地図", link=gmaps(row[IX["lat"]], row[IX["lon"]]))
            r += 1
    ws2.auto_filter.ref = f"A3:{get_column_letter(len(H2))}{r-1}"

    # ── ③ 条例・開発許可 ──
    ws3 = wb.create_sheet("③条例・開発許可")
    H3 = ["No.", "候補地", "都道府県", "市区町村", "都市計画\n(元表)", "条例\nリスク", "判定の理由",
          "開発許可権者", "権者区分", "権者の\n公表状況", "権者調査メモ（151者スイープ）", "権者の公表URL",
          "該当する台帳エントリ（都道府県・市区町村）", "S/Aランク変電所の\n優先監視自治体", "次アクション"]
    W3 = [6, 30, 10, 14, 12, 8, 34, 16, 12, 12, 54, 22, 48, 14, 40]
    head(ws3, H3, W3, "条例・開発許可の照合結果",
         f"台帳 data/ordinance_registry.json（{ordn['updated']}時点・30件）と開発許可権者151者のスイープ（2026-08-21）に照合。"
         "国レベル4件（国都計第7号ほか）は全案件共通のため⑤タブに別掲。")
    r = 4
    for s in sites:
        au = s["authority"]
        lst = "\n".join(f"・[{e['level']}] {e['title']}（{e['reg_type']}／{e.get('effective_date') or '施行日不明'}）" for e in s["ord_matched"]) or "（台帳に該当エントリなし）"
        if s["ord_level"] == "高":
            act = "所在市区町村の運用基準を一次情報で確認し、開発許可の要否・事前協議の要否を窓口照会。着工スケジュールに手続き期間を織り込む。"
        elif s["ord_level"].startswith("中"):
            act = "都道府県／権者の公表内容を確認のうえ、所在市区町村の運用（未公表分）を窓口照会。"
        else:
            act = "国都計第7号の運用（危険物該当・第一種特定工作物の扱い）を開発許可窓口に照会。公表が無くても内規運用の可能性あり。"
        vals = [s["label"], s["address"], s["pref"], s["muni_full"], s["toshi"], s["ord_level"], s["ord_why"],
                au["name"] if au else "要確認", s["authority_kind"], au["status"] if au else "",
                (au.get("note") or "") if au else "", "", lst,
                (f"S{s['priority']['S']}／A{s['priority']['A']}件" if s["priority"] else "対象外"), act]
        for i, v in enumerate(vals, start=1):
            f = None
            if i == 6:
                f = RISK_FILL if s["ord_level"] == "高" else (WARN_FILL if s["ord_level"].startswith("中") else OK_FILL)
            if i == 10 and au and au["status"] == "公表あり":
                f = WARN_FILL
            put(ws3, r, i, v, fill=f, align=CENTER if i in (1, 3, 6, 9, 10, 14) else None)
        if au and au.get("url"):
            put(ws3, r, 12, "権者ページ", link=au["url"])
        ws3.row_dimensions[r].height = 60
        r += 1
    ws3.auto_filter.ref = f"A3:{get_column_letter(len(H3))}{r-1}"

    # ── ④ DD許認可18項目 ──
    ws4 = wb.create_sheet("④DD許認可18項目")
    H4 = ["No.", "候補地", "所在", "項目\nNo.", "確認項目", "値・判定", "判定", "自動/手動",
          "コメント・根拠", "確認先URL", "担当者", "確認日"]
    W4 = [6, 22, 34, 6, 20, 40, 8, 10, 44, 14, 10, 10]
    auto_note = "reinfolib（不動産情報ライブラリAPI）の自動判定を取り込み済み" if a.values_dir else \
                "reinfolib（不動産情報ライブラリAPI）のキー未設定のため、7項目（浸水・土砂災害・液状化・市街化区域・調整区域・用途地域・自然公園）は「要確認」。確認先URLは全て緯度経度で開く。"
    head(ws4, H4, W4, "許認可事前確認（DD）18項目 × 全候補地",
         f"PX案件チェックシートのDD項目と同じ並び。{auto_note}")
    r = 4
    for s in sites:
        for d in s["dd"]:
            vals = [s["label"] if d["no"] == "1" else "", s["muni_full"] if d["no"] == "1" else "",
                    s["address"] if d["no"] == "1" else "", d["no"], d["name"], d["value"], d["judge"],
                    ("自動(reinfolib)" if (d["auto"] and d["comment"]) else ("自動対象" if d["auto"] else "手動")),
                    d["comment"]]
            for i, v in enumerate(vals, start=1):
                f = None
                if i == 7:
                    f = {"良": OK_FILL, "注意": WARN_FILL, "リスク": RISK_FILL, "要確認": GREY_FILL}.get(d["judge"])
                put(ws4, r, i, v, fill=f, align=CENTER if i in (4, 7, 8) else None)
            put(ws4, r, 10, "確認先", link=d["url"])
            put(ws4, r, 11, "")
            put(ws4, r, 12, "")
            r += 1
    ws4.auto_filter.ref = f"A3:{get_column_letter(len(H4))}{r-1}"

    # ── ⑤ 条例台帳（全国共通＋該当分）──
    ws5 = wb.create_sheet("⑤条例台帳（根拠）")
    H5 = ["区分", "レベル", "都道府県", "市区町村", "種別", "名称", "対象", "手続き", "要約",
          "DD・スケジュールへの影響", "公表日", "施行日", "最終確認", "URL", "該当候補地"]
    W5 = [12, 10, 10, 12, 12, 40, 36, 20, 60, 50, 11, 11, 11, 14, 16]
    head(ws5, H5, W5, "根拠となる条例・通知の台帳（該当分＋全国共通）",
         f"出典 data/ordinance_registry.json（{ordn['updated']}時点・全30件）。週次監視 scripts/ordinance_watch.py。")
    r = 4
    used = []
    for e in national:
        used.append(("全国共通", e, "全19地点"))
    for e in entries:
        if e["level"] == "国":
            continue
        hit = [s["label"] for s in sites if e in s["ord_matched"]]
        if hit:
            used.append(("該当あり", e, "、".join(hit)))
    for kind, e, who in used:
        vals = [kind, e["level"], e.get("pref") or "", e.get("muni") or "", e["reg_type"], e["title"],
                e.get("scope") or "", e.get("requirement") or "", e.get("summary") or "",
                e.get("dd_impact") or "", e.get("announced_date") or "", e.get("effective_date") or "",
                e.get("last_verified") or "", "", who]
        for i, v in enumerate(vals, start=1):
            f = SUB_FILL if (i == 1 and kind == "該当あり") else None
            put(ws5, r, i, v, fill=f, align=CENTER if i in (1, 2, 11, 12, 13) else None)
        if e.get("url"):
            put(ws5, r, 14, "一次情報", link=e["url"])
        ws5.row_dimensions[r].height = 70
        r += 1

    # ── ⑥ 凡例・出典 ──
    ws6 = wb.create_sheet("⑥凡例・出典")
    ws6.column_dimensions["A"].width = 26; ws6.column_dimensions["B"].width = 110
    ws6["A1"] = "凡例・出典・注意点"; ws6["A1"].font = Font(bold=True, size=13)
    notes = [
        ("作成日", today),
        ("元データ", f"{os.path.basename(a.src)}（{a.sheet}）の19筆。No.・所在・地目・地積・都市計画は元表のまま転記。"),
        ("緯度経度", "元表のGoogleマップ短縮URLを解決して得たピンの実座標。住所からの代表点ではないため、筆単位の判定に使える精度。"
                     "「座標の所在確認」列は国土地理院の逆ジオコーダで得た大字・丁目と元表の所在が一致するかの照合結果。"),
        ("変電所", "data/substations.js（全国66kV以上6,071件／座標あり5,329件）。半径10km以内を総合スコア順に上位5件。"
                   "系統スコア80点＝S1空容量35＋S2 N-1電制10＋S3潮流余裕15＋S4出力制御5＋S5配電用変電所10＋S6エリア5。"
                   "近接S7 20点＝≦500m:20／≦1km:15／≦2km:10／≦5km:4／超:0。総合ランク S:75+／A:60+／B:45+／C:44以下。"),
        ("座標精度", "変電所側の座標は OSM実測／公表資料／Google Places／推定 が混在。「座標精度」列で確認のこと。推定値は距離が数百m〜数km ずれる。"),
        ("条例", f"data/ordinance_registry.json（{ordn['updated']}時点・30件）。国4件は全国共通、都道府県・市区町村エントリは所在で照合。"
                 "リスク=高: 所在市区町村が公表済 ／ 中: 都道府県のみ公表 or 権者が公表あり ／ 低（要照会）: 公表確認できず。"
                 "「低」は規制が無いという意味ではなく、内規運用の可能性があるため窓口照会が必要。"),
        ("開発許可権者", "data/kaihatsu_kyokasha.json（151者・2026-08-21スイープ／公表あり17者・公表確認できず134者）。"
                        "政令市・中核市・施行時特例市は市が権者、それ以外は都道府県。"),
        ("DD18項目", "PX案件チェックシートと同じ項目・同じ確認先。確認先URLはすべて緯度経度で開くため、住所検索によるズレが起きない。"),
        ("判定の意味", "良=支障になる指定なし ／ 注意=指定ありだが対応可能な範囲 ／ リスク=許認可・設置可否に直結 ／ 要確認=未取得。"),
        ("免責", "本一覧はスクリーニング補助です。用途地域・農地・森林・埋蔵文化財等の最終確認は各行政窓口で行ってください。"),
    ]
    r = 3
    for k, v in notes:
        c = ws6.cell(row=r, column=1, value=k); c.font = HDR_FONT; c.fill = SUB_FILL; c.alignment = WRAP; c.border = BORDER
        c2 = ws6.cell(row=r, column=2, value=v); c2.font = SMALL; c2.alignment = WRAP; c2.border = BORDER
        ws6.row_dimensions[r].height = 46
        r += 1

    wb.save(a.out)
    print("wrote", a.out)
    for s in sites:
        b = s["subs"][0] if s["subs"] else None
        print(f"{s['label']:>7} {s['muni_full']:<12} 最寄{fmt_dist(b['dist']) if b else '—':>9} "
              f"総合{b['total'] if b else '—':>4} 条例{s['ord_level']}")

if __name__ == "__main__":
    main()
