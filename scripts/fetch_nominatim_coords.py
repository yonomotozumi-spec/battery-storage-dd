#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OpenStreetMap の検索サービス Nominatim から変電所座標を取得する。

fetch_substation_coords.py（Overpass）は power=substation タグが付いた地物しか拾わない。
一方 OSM には「○○変電所」という名前だけ付いた建物（building=yes）や敷地（landuse=industrial）
が相当数あり、Overpass の照合では取り逃していた。Nominatim は名前で地物を探すので、
タグに関係なく「○○変電所」を引ける。応答には operator（運営者）や substation=traction
（鉄道用）といった付随タグも入るため、鉄道会社や他電力の同名変電所を弾ける。

【Nominatim は「該当なし」を返さない】Google と同じく、無ければ似た名前の別の地物
（例: 「北島変電所」→「変電所沢川」「琴似変電所」）を返すので、名称の完全一致を必須にする。

【取り違えを防ぐ判定】採用は次をすべて満たす候補だけ。
  1. 名称が正規化後に完全一致（fetch_substation_coords.match_key。施設種別は区別、かな表記ゆれは吸収）
  2. 運営者が鉄道（鉄道・電鉄・旅客・メトロ・交通・JR）や substation=traction でない
  3. 運営者・名称に他エリアの電力会社名が入っていない
  4. 所在都道府県（応答の ISO3166-2 コード）が行の都道府県と一致。隣県なら同エリアかつ
     リストの県の既存座標から10km以内のときだけ採用（元リストの都道府県は管轄区分で
     所在地とずれる行があるため。fetch_google_coords.py と同じ規則）。
     都道府県が空の行（関西・九州）はエリアの一致と実測分布内で判定
  5. 候補が複数残ったら、互いに500m以内（同一敷地の建物など）のときだけ代表点を採用

利用規約: 1秒1リクエスト・連絡先入りの User-Agent が必須。結果は osm_cache/_nominatim.json
に貯めるので、同じ名前を再照会しない（--max-age-days で古いものだけ取り直す）。

使い方:
  # ① 精度を測る（座標が既知の行で誤差を出す。新しい情報源は必ずこれを先に行う）
  python scripts/fetch_nominatim_coords.py validate \
    --in "data/全国変電所リスト_66kV以上_座標追加.xlsx" --acc "公表資料(手動)" --sample 120

  # ② 未取得リストの空欄を埋める
  python scripts/list_missing_coords.py --in "data/全国変電所リスト_66kV以上_座標追加.xlsx" \
    --out missing.xlsx --ranks "S,A,B,C,データ無" --include-gated
  python scripts/fetch_nominatim_coords.py fill \
    --in "data/全国変電所リスト_66kV以上_座標追加.xlsx" --list missing.xlsx \
    --sheets 優先S_A --report nominatim_fill_report.csv

  # ③ 本体へ反映（出所は OSM実測(名称一致)）
  python scripts/merge_coords.py --in "data/全国変電所リスト_66kV以上_座標追加.xlsx" \
    --filled missing.xlsx --filled-sheets 優先S_A --acc "OSM実測(名称一致)" \
    --out "data/全国変電所リスト_66kV以上_座標追加.xlsx" --report nominatim_merge_report.csv
"""
import argparse
import collections
import csv
import json
import os
import random
import re
import statistics
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from fetch_substation_coords import (ACC_OSM_NAME, NEIGHBOR_MAX_KM, NEIGHBOR_PREFS, area_ranges,
                                     foreign_operator, haversine_km, in_area, match_key,
                                     pref_points, read_rows)

URL = "https://nominatim.openstreetmap.org/search"
USER_AGENT = ("battery-storage-dd substation coord fetcher "
              "(+https://github.com/yonomotozumi-spec/battery-storage-dd)")
SLEEP = 1.1          # 利用規約: 1リクエスト/秒
LIMIT = 40           # 同名が多い名前でも目的の県の候補が入るように多めに取る
CLUSTER_KM = 0.5     # 複数候補を同一施設とみなす距離
FILLED_FILL = PatternFill("solid", fgColor="E2EFDA")   # 緑=OSM
RAIL_RE = re.compile(r"鉄道|電鉄|旅客|メトロ|交通|JR|Railway|Rail|地下鉄|モノレール|新交通", re.I)
# 名前が一致しても施設ではない地物（バス停・駅・川・地名など）
NOT_FACILITY = {"highway", "railway", "waterway", "natural", "place", "public_transport",
                "boundary", "route", "tourism", "shop", "historic"}

PREF_ORDER = ["北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
              "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
              "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
              "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
              "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
              "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
              "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"]
CODE2PREF = {f"JP-{i:02d}": p for i, p in enumerate(PREF_ORDER, start=1)}


# ---------------------------------------------------------------- 取得とキャッシュ
def load_cache(path):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(path, cache):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)
    os.replace(tmp, path)


class Searcher:
    """名前で Nominatim を引く。結果はキャッシュに貯め、1秒1件のペースを守る。"""

    def __init__(self, cache_path, max_age_days=None, retries=3):
        self.path = cache_path
        self.cache = load_cache(cache_path)
        self.max_age = max_age_days
        self.retries = retries
        self.last = 0.0
        self.requests = 0
        self.dirty = 0

    def _fresh(self, ent):
        if self.max_age is None:
            return True
        return (time.time() - ent.get("ts", 0)) / 86400 <= self.max_age

    def search(self, name):
        key = unicodedata.normalize("NFKC", str(name)).strip()
        ent = self.cache.get(key)
        if ent and self._fresh(ent) and "error" not in ent:
            return ent["results"], None
        q = urllib.parse.urlencode({
            "q": key, "format": "jsonv2", "countrycodes": "jp", "limit": LIMIT,
            "addressdetails": 1, "extratags": 1, "namedetails": 1, "accept-language": "ja"})
        req = urllib.request.Request(f"{URL}?{q}", headers={"User-Agent": USER_AGENT})
        err = None
        for attempt in range(self.retries + 1):
            wait = SLEEP - (time.time() - self.last)
            if wait > 0:
                time.sleep(wait)
            self.last = time.time()
            self.requests += 1
            try:
                with urllib.request.urlopen(req, timeout=30) as res:
                    results = json.load(res)
                err = None
                break
            except urllib.error.HTTPError as e:
                err = f"HTTP {e.code}"
                if e.code in (429, 503, 502) and attempt < self.retries:
                    time.sleep(5 * (attempt + 1))
                    continue
                break
            except Exception as e:
                err = str(e)[:80]
                if attempt < self.retries:
                    time.sleep(3)
                    continue
        if err:
            self.cache[key] = {"ts": time.time(), "results": [], "error": err}
            self._flush()
            return [], err
        # 判定に使う項目だけ残す（キャッシュを小さく保つ）
        slim = []
        for r in results:
            slim.append({
                "name": r.get("name") or "",
                "name_ja": (r.get("namedetails") or {}).get("name:ja") or "",
                "lat": float(r["lat"]), "lon": float(r["lon"]),
                "category": r.get("category") or r.get("class") or "",
                "type": r.get("type") or "",
                "osm": f"{r.get('osm_type', '')}/{r.get('osm_id', '')}",
                "pref_code": (r.get("address") or {}).get("ISO3166-2-lvl4", ""),
                "pref_text": (r.get("address") or {}).get("province")
                             or (r.get("address") or {}).get("state") or "",
                "display": r.get("display_name", ""),
                "tags": {k: v for k, v in (r.get("extratags") or {}).items()
                         if k in ("power", "substation", "operator", "operator:ja", "voltage",
                                  "building", "landuse", "railway")},
            })
        self.cache[key] = {"ts": time.time(), "results": slim}
        self._flush()
        return slim, None

    def _flush(self, force=False):
        self.dirty += 1
        if force or self.dirty >= 10:
            save_cache(self.path, self.cache)
            self.dirty = 0

    def close(self):
        self._flush(force=True)


# ---------------------------------------------------------------- 判定
def pref_to_area(rows):
    tally = {}
    for r in rows:
        if r["pref"] and r["area"]:
            tally.setdefault(r["pref"], collections.Counter())[r["area"]] += 1
    return {p: c.most_common(1)[0][0] for p, c in tally.items()}


def result_pref(c):
    p = CODE2PREF.get(c.get("pref_code") or "")
    if p:
        return p
    t = c.get("pref_text") or ""
    return t if t in PREF_ORDER else ""


def pick(cands, row, p2a, pts, box):
    """候補から採用する1点を決める。(採用候補 or None, 理由) を返す。"""
    want = match_key(row["name"])
    if not want:
        return None, "正規化して固有名が残らない名前"
    if not cands:
        return None, "該当なし"
    same, why = [], collections.Counter()
    for c in cands:
        nm = c.get("name_ja") or c.get("name") or ""
        if match_key(nm) != want:
            why["名称不一致"] += 1
            continue
        tags = c.get("tags") or {}
        op = " ".join(v for k, v in tags.items() if k.startswith("operator"))
        if tags.get("substation") == "traction" or RAIL_RE.search(op):
            why[f"鉄道用（{op or 'traction'}）"] += 1
            continue
        if foreign_operator(op, row["area"]) or foreign_operator(nm, row["area"]):
            why[f"他社（{op or nm}）"] += 1
            continue
        if c.get("category") in NOT_FACILITY:
            why[f"施設でない（{c.get('category')}={c.get('type')}）"] += 1
            continue
        got = result_pref(c)
        if row["pref"]:
            if not got:
                why["所在地を確認できない"] += 1
                continue
            if got != row["pref"]:
                near = min((haversine_km(c["lat"], c["lon"], a, b) for a, b in pts.get(row["pref"], [])),
                           default=float("inf"))
                if not (got in NEIGHBOR_PREFS.get(row["pref"], []) and p2a.get(got) == row["area"]
                        and near <= NEIGHBOR_MAX_KM):
                    why[f"所在地が{got}"] += 1
                    continue
                c = dict(c, note=f"所在地は隣県の{got}・{row['pref']}の既存座標から{near:.1f}km")
        else:
            if got and p2a.get(got) != row["area"]:
                why[f"所在地が{got}（{p2a.get(got)}エリア）"] += 1
                continue
            if not in_area(box, row["area"], c["lat"], c["lon"]):
                why[f"{row['area']}エリアの実測分布外"] += 1
                continue
        same.append(c)
    if not same:
        top = why.most_common(1)[0][0] if why else "該当なし"
        return None, top
    # 複数残ったら同一敷地（互いに500m以内）のときだけ採用。power=substation を優先
    if len(same) > 1:
        spread = max(haversine_km(a["lat"], a["lon"], b["lat"], b["lon"]) for a in same for b in same)
        if spread > CLUSTER_KM:
            return None, f"候補{len(same)}件が{spread:.1f}km離れていて特定不可"
    same.sort(key=lambda c: (c.get("tags", {}).get("power") != "substation",
                             c.get("category") not in ("power", "building")))
    best = same[0]
    return best, "採用" + (f"（{best['note']}）" if best.get("note") else "")


def describe(c):
    tags = c.get("tags") or {}
    bits = [f"{c.get('category')}={c.get('type')}"]
    for k in ("power", "substation", "operator", "voltage"):
        if tags.get(k):
            bits.append(f"{k}={tags[k]}")
    return f"Nominatim {c.get('name_ja') or c.get('name')} ({c.get('osm')}; " + ", ".join(bits) + ")"


# ---------------------------------------------------------------- コマンド
def cmd_validate(args):
    rows, _, _ = read_rows(args.src, args.list_sheet)
    truth = [r for r in rows if str(r["acc"] or "") == args.acc and isinstance(r["lat"], (int, float))]
    if args.area:
        truth = [r for r in truth if r["area"] == args.area]
    if not truth:
        raise SystemExit(f"座標精度が '{args.acc}' の行がありません")
    if args.sample and len(truth) > args.sample:
        random.seed(args.seed)
        truth = random.sample(truth, args.sample)
    p2a, pts, box = pref_to_area(rows), pref_points(rows), area_ranges(rows)
    s = Searcher(args.cache, args.max_age_days)
    kept, dropped, errors = [], collections.Counter(), 0
    for r in truth:
        cands, err = s.search(r["name"])
        if err:
            errors += 1
        best, why = pick(cands, r, p2a, pts, box)
        if best:
            kept.append((r, haversine_km(r["lat"], r["lon"], best["lat"], best["lon"]) * 1000, best))
        else:
            dropped[why.split("（")[0]] += 1
    s.close()
    ds = sorted(d for _, d, _ in kept)
    print(f"照会 {len(truth)}件（通信 {s.requests}回・失敗 {errors}）→ 採用 {len(kept)} / 見送り {len(truth) - len(kept)}")
    if ds:
        print(f"採用分の誤差: 中央値 {statistics.median(ds):.0f}m / 最大 {max(ds):.0f}m")
        for th in (100, 500, 1000, 5000):
            print(f"  <{th}m: {sum(1 for x in ds if x < th)}/{len(ds)}")
    for k, v in dropped.most_common():
        print(f"  見送り: {k} {v}件")
    bad = [(r, d, b) for r, d, b in kept if d >= 1000]
    if bad:
        print("1km以上ずれた採用（要確認）:")
        for r, d, b in bad:
            print(f"  {r['area']} {r['name']} {r['pref']} {d/1000:.1f}km  {describe(b)}")


def cmd_fill(args):
    rows, _, _ = read_rows(args.src, args.list_sheet)
    p2a, pts, box = pref_to_area(rows), pref_points(rows), area_ranges(rows)
    by_no = {r["no"]: r for r in rows}
    s = Searcher(args.cache, args.max_age_days)
    wb = load_workbook(args.list)
    report, filled, skipped = [], 0, 0
    for sheet in [x.strip() for x in args.sheets.split(",")]:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        hdr = {}
        for c in range(1, ws.max_column + 1):
            h = str(ws.cell(2, c).value or "").replace("\n", "")
            if h:
                hdr[h] = c
        for token in ("No.", "変電所名", "エリア", "都道府県", "緯度", "経度"):
            if token not in hdr:
                raise SystemExit(f"シート'{sheet}'に列'{token}'がありません")
        for r in range(3, ws.max_row + 1):
            name = ws.cell(r, hdr["変電所名"]).value
            if not name:
                continue
            if args.area and ws.cell(r, hdr["エリア"]).value != args.area:
                continue
            if ws.cell(r, hdr["緯度"]).value is not None and not args.overwrite:
                continue
            no = ws.cell(r, hdr["No."]).value
            row = by_no.get(no) or {"name": str(name), "pref": ws.cell(r, hdr["都道府県"]).value or "",
                                    "area": ws.cell(r, hdr["エリア"]).value or "", "vmax": None}
            if not match_key(row["name"]):
                skipped += 1
                report.append([sheet, r, no, name, row["pref"], "", "", "不採用",
                               "正規化して固有名が残らない名前（数字だけの名前）", ""])
                continue
            cands, err = s.search(row["name"])
            if err:
                skipped += 1
                report.append([sheet, r, no, name, row["pref"], "", "", "不採用", f"照会に失敗: {err}", ""])
                continue
            best, why = pick(cands, row, p2a, pts, box)
            if not best:
                skipped += 1
                report.append([sheet, r, no, name, row["pref"], "", "", "不採用", why, ""])
                continue
            ws.cell(r, hdr["緯度"]).value = round(best["lat"], 6)
            ws.cell(r, hdr["経度"]).value = round(best["lon"], 6)
            if "出典メモ" in hdr:
                ws.cell(r, hdr["出典メモ"]).value = describe(best)
            for k in ("緯度", "経度", "出典メモ"):
                if k in hdr:
                    ws.cell(r, hdr[k]).fill = FILLED_FILL
            filled += 1
            report.append([sheet, r, no, name, row["pref"], round(best["lat"], 6), round(best["lon"], 6),
                           why, describe(best), best.get("display", "")])
    s.close()
    wb.save(args.out or args.list)
    print(f"記入 {filled}件 / 見送り {skipped}件（通信 {s.requests}回）")
    print(f"出力: {args.out or args.list}")
    if args.report:
        with open(args.report, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["シート", "行", "No.", "変電所名", "都道府県", "緯度", "経度",
                        "判定", "OSM側の名称/理由", "所在地"])
            w.writerows(report)
        print(f"レポート: {args.report}")
    if filled:
        print(f"※ 次に merge_coords.py --acc \"{ACC_OSM_NAME}\" で本体xlsxへ反映してください。")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name, help_ in (("validate", "座標が既知の行で誤差を測る"), ("fill", "未取得リストの緯度・経度欄を埋める")):
        p = sub.add_parser(name, help=help_)
        p.add_argument("--in", dest="src", required=True, help="本体xlsx")
        p.add_argument("--list-sheet", default="全国変電所リスト")
        p.add_argument("--cache", default="osm_cache/_nominatim.json")
        p.add_argument("--max-age-days", type=float, help="この日数より古いキャッシュは取り直す（既定: 取り直さない）")
        p.add_argument("--area", help="このエリアだけに限定する")
        if name == "validate":
            p.add_argument("--acc", default="公表資料(手動)", help="正解として使う座標精度")
            p.add_argument("--sample", type=int, help="無作為にこの件数だけ照会する")
            p.add_argument("--seed", type=int, default=1)
        else:
            p.add_argument("--list", required=True, help="list_missing_coords.py が出した未取得リスト")
            p.add_argument("--out")
            p.add_argument("--sheets", default="優先S_A,劣後S_A")
            p.add_argument("--report")
            p.add_argument("--overwrite", action="store_true")
    args = ap.parse_args()
    cmd_validate(args) if args.cmd == "validate" else cmd_fill(args)


if __name__ == "__main__":
    main()
