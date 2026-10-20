#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
系統トポロジー（変電所ごとの「つながっている送電線」「相手側の変電所」「上位系の制約設備」）を、
各送配電事業者の公表CSV（変圧器＋送電線）から作り、Web画面（substation.html）用の
data/grid_topology.js を出力する。

対象は、全国共通形式（OCCTOの系統空容量公表フォーマット）のCSVを公開している6社。
  東北・北陸・関西・中国・四国・九州
東京（403）・中部（403）・北海道（PDFのみ）・沖縄（403）は自動取得できないため対象外。
中部は公表マッピングのGeoJSON（chuden_cache/）から送電線の実形状を別途表示する（build_chuden_grid_web.py）。

  # ① 取得：変圧器CSVと送電線CSVを grid_cache/ に保存（要ネットワーク）
  python scripts/build_grid_topology.py fetch --cache-dir grid_cache

  # ② 生成：本体xlsxの変電所に紐づけて data/grid_topology.js と照合レポートを出す（ネットワーク不要）
  python scripts/build_grid_topology.py build \
    --in "data/全国変電所リスト_66kV以上_座標追加.xlsx" --cache-dir grid_cache \
    --out data/grid_topology.js --report "docs/系統トポロジー_照合レポート.md"

【つながりの出どころ】送電線CSVの「潮流方向」欄（例: 青森変電所→瀬辺地変電所、変1→変6、A→B、浦桑→小値賀）。
  両端を変電所CSVの行（変電所No または 変電所名）に引き当てる。引き当てられない端点
  （発電所・需要家・○○線分岐 など）は文字のまま残す。推測で埋めない。
【上位系の制約設備】変圧器CSV・送電線CSVの「平常時出力制御が必要となりうる設備(上位系設備)」欄。
  各社の設備番号（送0104・変002・H012・基L22・基変31・19地区送電線№6 など）を同じ公表物の
  送電線名・変電所名に引き当てる。引き当てられないもの（フェンス名など）は文字のまま残す。
【本体xlsxとの紐づけ】エリア内で変電所名（種別を外した名前）が一致する行に紐づける（update_capacity.py と同じ）。
"""
import argparse
import csv
import datetime
import io
import json
import os
import re
import sys
import unicodedata
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fetch_substation_coords import loose_name  # noqa: E402
from update_capacity import (SOURCES, DASHES, http_get, decode, cache_name, text, num,  # noqa: E402
                             n1_value, curtail_value, find_col)
from score_substations import read_substations  # noqa: E402

# 送電線CSVも拾うように、update_capacity.py の変圧器CSVの正規表現を広げる
LINE_PATTERNS = {
    "東北": r"sys_capa_(?:kikan|local)\d+_(?:tr|line)_\d{6}_\d+\.csv",
    "北陸": r"sys_capa_(?:kikan|local)\d+_(?:tr|line)_\d{6}_\d+\.csv",
    "関西": r"154kv_(?:more|less)_(?:trans|line)\.csv",
    "中国": r"zip/csv_[a-z0-9]+\.zip",
    "四国": r"sys_capa_(?:kikan|local)\d+_(?:tr|line)_\d{6}_\d+\.csv",
    "九州": r"/var/rev0/[0-9/]+/td_[A-Za-z0-9]+\.zip",
}

# 端点・設備名として出てくるが変電所ではない語（文字のまま残す）
GENERIC = {"発電所", "需要家", "開閉所", "配電塔", "電源", "需要家等", "他社", "需要", "負荷"}


def kind_of(fname, member=None):
    s = (member or fname)
    if "送電線" in s or "_line_" in s or "line" in s.lower():
        return "line"
    if "変圧器" in s or "_tr_" in s or "trans" in s:
        return "tr"
    return None


def cmd_fetch(args):
    os.makedirs(args.cache_dir, exist_ok=True)
    manifest = {"fetched": datetime.date.today().isoformat(), "areas": {}}
    for area, (page, _) in SOURCES.items():
        pat = LINE_PATTERNS[area]
        try:
            html = decode(http_get(page))
        except Exception as e:  # noqa: BLE001
            print(f"{area}: ページ取得失敗 {e}", file=sys.stderr)
            manifest["areas"][area] = {"page": page, "error": f"ページ取得失敗: {e}", "files": []}
            continue
        import urllib.parse
        links = sorted({urllib.parse.urljoin(page, h) for h in re.findall(r'href="([^"]+)"', html)
                        if re.search(pat, h)})
        files = []
        for url in links:
            try:
                b = http_get(url)
            except Exception as e:  # noqa: BLE001
                print(f"{area}: {url} 取得失敗 {e}", file=sys.stderr)
                continue
            if url.endswith(".zip"):
                try:
                    zf = zipfile.ZipFile(io.BytesIO(b))
                except zipfile.BadZipFile:
                    continue
                for n in zf.namelist():
                    if not n.lower().endswith(".csv") or kind_of(n) is None:
                        continue
                    fn = cache_name(area, url) + "__" + re.sub(r"[^\w.\-]", "_", os.path.basename(n))
                    with open(os.path.join(args.cache_dir, fn), "wb") as f:
                        f.write(zf.read(n))
                    files.append({"file": fn, "url": url, "member": n, "kind": kind_of(n)})
            else:
                fn = cache_name(area, url)
                with open(os.path.join(args.cache_dir, fn), "wb") as f:
                    f.write(b)
                files.append({"file": fn, "url": url, "kind": kind_of(fn)})
        manifest["areas"][area] = {"page": page, "files": files}
        print(f"{area}: {len(files)}ファイル")
    with open(os.path.join(args.cache_dir, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)


# ---------- CSVの読み取り ----------

def clean(s):
    """値の前後の空白・CSV由来の先頭アポストロフィ（'0104）を落とす。"""
    return text(s).strip("'").strip()


def parse_file(path):
    t = decode(open(path, "rb").read())
    rows = list(csv.reader(io.StringIO(t)))
    date = clean(rows[0][0]) if rows and rows[0] else ""
    hi = next((i for i, r in enumerate(rows[:12])
               if any(("変電所名" in (c or "")) or ("送電線名" in (c or "")) for c in r)), None)
    if hi is None:
        return date, None, []
    h = [c or "" for c in rows[hi]]
    is_line = any("送電線名" in c for c in h)
    col = {
        "no": find_col(h, "送電線No") if is_line else find_col(h, "変電所No"),
        "name": find_col(h, "送電線名") if is_line else find_col(h, "変電所名"),
        "kv": find_col(h, "電圧", avoid=("一次", "二次")) if is_line else None,
        "v1": None if is_line else find_col(h, "一次"),
        "v2": None if is_line else find_col(h, "二次"),
        "circ": find_col(h, "回線数") if is_line else find_col(h, "台数"),
        "cap": find_col(h, "設備容量"),
        "opcap": find_col(h, "運用容量値"),
        "factor": find_col(h, "制約要因"),
        "dir": find_col(h, "潮流方向") if find_col(h, "潮流方向") is not None else find_col(h, "潮流正方向"),
        "dir_to": find_col(h, "潮流方向(至)"),
        "flow": find_col(h, "予想潮流"),
        "cur": find_col(h, "当該設備", "空") if find_col(h, "当該設備", "空") is not None
        else (find_col(h, "空き容量") if find_col(h, "空き容量") is not None else find_col(h, "空容量", avoid=("上位",))),
        "up": find_col(h, "上位系", "空"),
        "n1": find_col(h, "N-1電制適用可否"),
        "n1amt": find_col(h, "N-1電制適用可能量"),
        "curtail": find_col(h, "平常時出力制御", avoid=("設備", "系統")),
        "cur_eq": find_col(h, "当該設備", avoid=("空",)),
        "up_eq": find_col(h, "上位系", avoid=("空",)),
        "remark": find_col(h, "備考") if find_col(h, "備考") is not None else find_col(h, "備 考"),
    }
    out = []
    for r in rows[hi + 1:]:
        g = lambda k: (r[col[k]] if col.get(k) is not None and col[k] < len(r) else "")  # noqa: E731
        name = clean(g("name"))
        no = clean(g("no"))
        if not name and not no:
            continue
        rec = {
            "no": no, "name": name,
            "circ": num(g("circ")), "cap": num(g("cap")), "opcap": num(g("opcap")),
            "factor": clean(g("factor")), "flow": num(g("flow")),
            "cur": num(g("cur")), "up": num(g("up")) if col["up"] is not None else None,
            "n1": n1_value(g("n1")), "n1amt": num(g("n1amt")),
            "curtail": curtail_value(g("curtail")),
            "cur_eq": clean(g("cur_eq")), "up_eq": clean(g("up_eq")),
            "remark": clean(g("remark")),
        }
        if is_line:
            rec["kv"] = num(g("kv"))
            d = clean(g("dir"))
            if col["dir_to"] is not None:
                d = f"{d}→{clean(g('dir_to'))}" if d or clean(g("dir_to")) else ""
            rec["dir"] = d
        else:
            rec["v1"], rec["v2"] = num(g("v1")), num(g("v2"))
        out.append(rec)
    return date, ("line" if is_line else "tr"), out


# ---------- 名前・番号の引き当て ----------

def node_key(name):
    """変電所名の照合キー。先頭のアポストロフィ、「（変）」の略記、末尾の「開」（北陸の開閉所略記）を吸収する。"""
    s = unicodedata.normalize("NFKC", str(name or "")).strip("'").strip()
    s = s.replace("ヶ", "ケ").replace("ｹ", "ケ").replace("ケ", "ケ")
    s = re.sub(r"\(変\)$", "変電所", s)
    s = re.sub(r"\(開\)$", "開閉所", s)
    s = re.sub(r"\(発\)$", "発電所", s)
    s = re.sub(r"\(.*?\)", "", s)  # 「新改開閉所(変14)」の括弧書き
    if re.search(r"[^\s]開$", s) and not s.endswith("開閉所"):
        s = s[:-1] + "開閉所"
    return loose_name(s)


def is_generic(tok):
    t = re.sub(r"[<＜《].*?[>＞》]", "", tok).strip()
    return t in GENERIC or t.endswith("需要家") or (t.endswith("発電所") and len(t) <= 5)


class Area:
    """1エリア分の公表データ。変電所ノード・送電線・引き当て用の索引を持つ。"""

    def __init__(self, name):
        self.name = name
        self.dates = set()
        self.files = []          # (tag, kind, path)
        self.nodes = {}          # key -> {"name","banks":[...],"ids":set(),"files":set()}
        self.node_order = []
        self.lines = []          # dict
        self.tr_by_no = {}       # (tag, no) -> key
        self.tr_by_no_area = {}  # no -> key（エリア内で一意なときだけ）
        self.line_by_no = {}     # (tag, no) -> line idx
        self.line_by_no_area = {}
        self.file_pref = {}      # tag -> 都道府県（四国の「徳島変N」などの解決用）

    def add_tr(self, tag, rec):
        key = node_key(rec["name"])
        if not key:
            return
        n = self.nodes.get(key)
        if n is None:
            n = {"name": rec["name"].strip("'"), "banks": [], "ids": set(), "files": set(), "key": key}
            self.nodes[key] = n
            self.node_order.append(key)
        n["banks"].append({"tag": tag, **rec})
        n["files"].add(tag)
        if rec["no"]:
            n["ids"].add(rec["no"])
            self.tr_by_no[(tag, rec["no"])] = key
            prev = self.tr_by_no_area.get(rec["no"])
            self.tr_by_no_area[rec["no"]] = key if prev in (None, key) else False

    def add_line(self, tag, rec):
        idx = len(self.lines)
        self.lines.append({"tag": tag, **rec})
        if rec["no"]:
            self.line_by_no[(tag, rec["no"])] = idx
            prev = self.line_by_no_area.get(rec["no"])
            self.line_by_no_area[rec["no"]] = idx if prev in (None, idx) else False

    # 端点（潮流方向の片側）を変電所ノードに引き当てる。戻り値は ("node", key) か ("text", 文字列)
    def resolve_endpoint(self, tag, tok):
        t = unicodedata.normalize("NFKC", tok).strip().strip("'")
        if not t or t in DASHES:
            return None
        if is_generic(t):
            return ("text", t)
        # 四国: 「変6」「徳島変3」「新改開閉所(変14)」
        m = re.fullmatch(r"(\D*?)変(\d+)", t) or re.search(r"\((?:\D*?)変(\d+)\)", t)
        if m and self.name == "四国":
            if re.fullmatch(r"(\D*?)変(\d+)", t):
                pref, no = m.group(1), m.group(2)
            else:
                pref, no = "", m.group(1)
            if not pref:
                k = self.tr_by_no.get((tag, no))
                if k:
                    return ("node", k)
            else:
                for ftag, fp in self.file_pref.items():
                    if fp.startswith(pref):
                        k = self.tr_by_no.get((ftag, no))
                        if k:
                            return ("node", k)
        # 関西・四国など: 変電所No そのもの（同一ファイル優先、次にエリア内で一意なもの）
        k = self.tr_by_no.get((tag, t))
        if k:
            return ("node", k)
        k = self.tr_by_no_area.get(t)
        if k:
            return ("node", k)
        # 名前で引く（東北・北陸・中国・九州）
        key = node_key(t)
        if key and key in self.nodes:
            return ("node", key)
        return ("text", t)

    # 上位系設備の欄（設備番号の列挙）を送電線・変電所に引き当てる
    def resolve_equipment(self, tag, s):
        refs, prefix = [], ""
        s = unicodedata.normalize("NFKC", s or "")
        if not s or s in DASHES or s in ("対象", "-"):
            return refs
        # 九州: 「19地区送電線№6」「5地区 送電線No.2」「15地区 変電所No.5,送電線No.9」「変電所No2(1)」。
        # 地区の指定が無ければその行のファイル（地区）の設備。NFKC で「№」は「No」になる
        if self.name == "九州":
            def district_tag(dist):
                for ftag in self.file_pref:
                    nt = unicodedata.normalize("NFKC", ftag)
                    if (dist == "離島" and "離島" in nt) or re.match(rf"★?\s*{re.escape(dist)}[\s_]", nt):
                        return ftag
                return None
            cur_tag = tag
            pat = (r"(?P<dist>\d+|離島)地区|"
                   r"(?P<kind>送電線|変電所|変圧器)No\.?(?P<no>\d+?(?:-\d+)?)(?:\((?P<sub>\d+)\))?"
                   r"(?=(?:\d+|離島)地区|[^\d\-]|$)")
            for m in re.finditer(pat, s):
                if m.group("dist"):
                    cur_tag = district_tag(m.group("dist")) or cur_tag
                    continue
                kind, no = m.group("kind"), m.group("no")
                if kind == "送電線":
                    i = self.line_by_no.get((cur_tag, no))
                    refs.append(("L", i) if i is not None else ("T", m.group(0)))
                else:
                    sub = m.group("sub")
                    # 変電所No は「1(1)」「1-2」「2」の形。括弧・枝番付きで無ければ本体番号で引く
                    k = (self.tr_by_no.get((cur_tag, f"{no}({sub})")) if sub else None) \
                        or self.tr_by_no.get((cur_tag, no)) \
                        or self.tr_by_no.get((cur_tag, no.split("-")[0])) \
                        or self.tr_by_no.get((cur_tag, no.split("-")[0] + "(1)"))
                    refs.append(("N", k) if k else ("T", m.group(0)))
            rest = re.sub(pat, " ", s)
            s = rest
        def find_line(no, kikan_only=False):
            for (t2, n2), i in self.line_by_no.items():
                if n2 == no and (not kikan_only or "kikan" in t2):
                    return ("L", i)
            return None

        def find_tr(no, kikan_only=False):
            for (t2, n2), k in self.tr_by_no.items():
                if n2 == no and (not kikan_only or "kikan" in t2):
                    return ("N", k)
            return None

        for tok in re.split(r"[,，、;；/／\s]+", s):
            tok = tok.strip()
            if not tok or tok in DASHES:
                continue
            m = re.fullmatch(r"(基送|基変|基L|基S|送|変)?(.+)", tok)
            pre, body = (m.group(1) or ""), m.group(2)
            if pre:
                prefix = pre
            elif re.fullmatch(r"\d+[A-Z]?", body) and prefix:
                pre = prefix  # 「基送27,28,29」の 28・29 は直前の種別を引き継ぐ
            ref = None
            if pre in ("送", "基送"):
                ref = find_line(body, kikan_only=(pre == "基送"))
            elif pre == "基L":
                ref = find_line("基L" + body, kikan_only=True)
            elif pre in ("変", "基変"):
                ref = find_tr(body, kikan_only=(pre == "基変"))
            elif pre == "基S":
                ref = find_tr("基S" + body)
            if ref is None:
                i = self.line_by_no.get((tag, tok))
                if i is None:
                    i2 = self.line_by_no_area.get(tok)
                    i = i2 if i2 not in (None, False) else None
                if i is not None:
                    ref = ("L", i)
                else:
                    k = self.tr_by_no.get((tag, tok)) or (self.tr_by_no_area.get(tok) or None)
                    if k:
                        ref = ("N", k)
            refs.append(ref or ("T", tok))
        # 重複を除く
        seen, out = set(), []
        for r in refs:
            if r not in seen:
                seen.add(r); out.append(r)
        return out


def load_areas(cache_dir):
    with open(os.path.join(cache_dir, "manifest.json"), encoding="utf-8") as f:
        manifest = json.load(f)
    areas = {}
    for area, info in manifest["areas"].items():
        A = Area(area)
        for fi in info.get("files", []):
            path = os.path.join(cache_dir, fi["file"])
            if not os.path.exists(path):
                continue
            date, kind, recs = parse_file(path)
            if kind is None:
                continue
            tag = fi.get("member") or fi["file"]
            tag = re.sub(r"\.csv$", "", os.path.basename(tag))
            tag = re.sub(r"_(line|tr)_\d{6}_\d+$", "", tag).replace("_line", "").replace("_trans", "")
            tag = tag.replace("_送電線", "").replace("_変圧器", "")
            if date:
                A.dates.add(date)
            A.files.append((tag, kind, fi["file"], len(recs)))
            if kind == "tr":
                for r in recs:
                    A.add_tr(tag, r)
            else:
                for r in recs:
                    A.add_line(tag, r)
        areas[area] = A
    return manifest, areas


def cmd_build(args):
    manifest, areas = load_areas(args.cache_dir)
    master = read_substations(args.src, args.list_sheet)
    # (エリア, 名前キー) -> 本体No.のリスト／都道府県
    m_index, m_pref = {}, {}
    for (no, area, name, pref, *_rest) in master:
        k = node_key(name)
        if area and k:
            m_index.setdefault((area, k), []).append(no)
            if pref:
                m_pref.setdefault((area, k), pref)

    strings, s_idx = [], {}

    def sid(s):
        s = s or ""
        if s not in s_idx:
            s_idx[s] = len(strings); strings.append(s)
        return s_idx[s]

    area_list = list(areas.keys())
    nodes_out, lines_out, by_no = [], [], {}
    report = {"generated": datetime.date.today().isoformat(), "areas": {}}
    node_idx_global = {}   # (area, key) -> idx

    for ai, area in enumerate(area_list):
        A = areas[area]
        # ファイルごとの都道府県（四国の「徳島変N」の解決用）：そのファイルの変電所が本体で最も多く属する県
        for tag in {t for t, k, _, _ in A.files if k == "tr"}:
            votes = {}
            for key, n in A.nodes.items():
                if tag in n["files"]:
                    p = m_pref.get((area, key))
                    if p:
                        votes[p] = votes.get(p, 0) + 1
            A.file_pref[tag] = max(votes, key=votes.get) if votes else ""
        for tag in {t for t, k, _, _ in A.files if k == "line"}:
            A.file_pref.setdefault(tag, "")

        # 変圧器CSVには無いが本体リストにはある設備（開閉所など）を、送電線の端点から拾ってノードにする。
        # 座標があれば地図に線を引けるため。バンク情報は無い（名前・本体No.だけ）
        for L in A.lines:
            d = unicodedata.normalize("NFKC", L["dir"] or "")
            for p in [x.strip() for x in re.split(r"→|->|⇒", d) if x.strip()]:
                r = A.resolve_endpoint(L["tag"], p)
                if r is None or r[0] == "node" or is_generic(p):
                    continue
                k = node_key(p)
                if k and (area, k) in m_index and k not in A.nodes:
                    A.nodes[k] = {"name": p, "banks": [], "ids": set(), "files": set(), "key": k}
                    A.node_order.append(k)
        # ノード
        for key in A.node_order:
            n = A.nodes[key]
            node_idx_global[(area, key)] = len(nodes_out)
            nos = m_index.get((area, key), [])
            for no in nos:
                by_no[str(no)] = len(nodes_out)
            nodes_out.append([ai, n["name"], nos, [], [], sorted(n["ids"])])
        # 送電線（端点を引き当てる）
        stat = {"tr_rows": sum(len(n["banks"]) for n in A.nodes.values()), "nodes": len(A.nodes),
                "nodes_matched": sum(1 for k in A.nodes if m_index.get((area, k))),
                "lines": len(A.lines), "ends": 0, "ends_node": 0, "ends_text": 0, "unresolved": {}}
        line_base = len(lines_out)
        for li, L in enumerate(A.lines):
            ends = []
            d = L["dir"]
            parts = [p for p in re.split(r"→|->|⇒", unicodedata.normalize("NFKC", d))] if d else []
            parts = [p.strip() for p in parts]
            if len(parts) > 2:
                parts = [parts[0], parts[-1]]
            for p in parts:
                r = A.resolve_endpoint(L["tag"], p)
                if r is None:
                    continue
                stat["ends"] += 1
                if r[0] == "node":
                    stat["ends_node"] += 1
                    ends.append(node_idx_global[(area, r[1])])
                else:
                    stat["ends_text"] += 1
                    ends.append(r[1])
                    if not is_generic(r[1]):
                        stat["unresolved"][r[1]] = stat["unresolved"].get(r[1], 0) + 1
            while len(ends) < 2:
                ends.append("")
            lines_out.append([ai, L["no"], L["name"], L["kv"], L["circ"], ends[0], ends[1],
                              L["cap"], L["opcap"], L["factor"], L["flow"], L["cur"], L["up"],
                              L["n1"] or "", L["n1amt"], L["curtail"] or "", sid(L["up_eq"]), [], L["remark"]])
        # 上位系設備の引き当て（送電線・変電所の各バンク）
        def refs_out(refs):
            out = []
            for kind, v in refs:
                if kind == "L":
                    out.append(["L", line_base + v])
                elif kind == "N":
                    out.append(["N", node_idx_global[(area, v)]])
                else:
                    out.append(["T", v])
            return out
        for li, L in enumerate(A.lines):
            lines_out[line_base + li][17] = refs_out(A.resolve_equipment(L["tag"], L["up_eq"]))
        for key in A.node_order:
            n = A.nodes[key]
            ni = node_idx_global[(area, key)]
            for b in n["banks"]:
                nodes_out[ni][3].append([b["v1"], b["v2"], b["circ"], b["cap"], b["opcap"], b["flow"],
                                         b["cur"], b["up"], b["n1"] or "", b["curtail"] or "", b["factor"],
                                         sid(b["up_eq"]), refs_out(A.resolve_equipment(b["tag"], b["up_eq"])),
                                         b["remark"], b["no"]])
        stat["unresolved"] = sorted(stat["unresolved"].items(), key=lambda x: -x[1])[:15]
        stat["dates"] = sorted(A.dates)
        stat["files"] = [[t, k, c] for t, k, _, c in A.files]
        report["areas"][area] = stat

    # 本体側のカバー率：エリアごとに、容量データのある本体行のうちノードに紐づいた行
    cover = {}
    for (no, area, name, pref, vmax, vsec, lat, lon, acc, cap, opcap, flow, avail, avail_up, n1, curtail, addr) in master:
        if area not in areas:
            continue
        c = cover.setdefault(area, {"rows": 0, "with_data": 0, "linked": 0, "linked_sa": 0})
        c["rows"] += 1
        has = any(v is not None for v in (opcap, avail, avail_up))
        if has:
            c["with_data"] += 1
        if str(no) in by_no:
            c["linked"] += 1
    for area, c in cover.items():
        report["areas"][area]["master"] = c

    payload = {
        "meta": {"generated": report["generated"], "fetched": manifest.get("fetched", ""),
                 "areas": area_list, "nodes": len(nodes_out), "lines": len(lines_out),
                 "linkedRows": len(by_no),
                 "sources": {a: {"page": manifest["areas"][a].get("page", ""), "dates": report["areas"][a]["dates"]}
                             for a in area_list}},
        "nodeFields": ["area", "name", "nos", "banks", "_", "ids"],
        "bankFields": ["v1", "v2", "units", "cap", "opcap", "flow", "availCur", "availUp", "n1", "curtail",
                       "factor", "upEq", "upRefs", "remark", "no"],
        "lineFields": ["area", "no", "name", "kv", "circuits", "from", "to", "cap", "opcap", "factor", "flow",
                       "availCur", "availUp", "n1", "n1amt", "curtail", "upEq", "upRefs", "remark"],
        "strings": strings,
        "nodes": nodes_out,
        "lines": lines_out,
        "byNo": by_no,
    }
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    header = ("// 自動生成ファイル - 直接編集しないこと\n"
              "// 生成: scripts/build_grid_topology.py（各社の公表CSV grid_cache/ から）\n")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(header + "window.GRID_TOPOLOGY=" + body + ";\n")
    print(f"出力: {args.out}  変電所ノード{len(nodes_out)}件・送電線{len(lines_out)}件・本体紐づけ{len(by_no)}行 "
          f"{os.path.getsize(args.out) / 1024:.0f}KB")

    if args.report:
        write_report(args.report, report, area_list)
        print(f"レポート: {args.report}")


def write_report(path, report, area_list):
    L = [f"# 系統トポロジー 照合レポート（{report['generated']}）", "",
         "各送配電事業者の公表CSV（変圧器・送電線）から、変電所ごとの「つながっている送電線」「相手側の変電所」"
         "「上位系の制約設備」を引き当てた結果です。`scripts/build_grid_topology.py build` が生成します。", "",
         "| エリア | 公表日 | 変電所ノード | うち本体リストに紐づいた | 本体の行（紐づいた／容量データあり／全行） | 送電線 | 端点の引き当て（変電所／文字のまま） |",
         "|---|---|---|---|---|---|---|"]
    for a in area_list:
        s = report["areas"][a]
        m = s.get("master", {})
        L.append(f"| {a} | {'、'.join(s['dates'])} | {s['nodes']} | {s['nodes_matched']} | "
                 f"{m.get('linked', 0)}／{m.get('with_data', 0)}／{m.get('rows', 0)} | {s['lines']} | "
                 f"{s['ends_node']}／{s['ends_text']}（全{s['ends']}） |")
    L += ["", "「文字のまま」の端点は、発電所・需要家・○○線分岐（送電線の途中の分岐点）など変電所CSVに無いものです。推測で埋めていません。", ""]
    for a in area_list:
        s = report["areas"][a]
        if s["unresolved"]:
            L.append(f"## {a}：変電所に引き当てられなかった端点（発電所・需要家を除く上位15件）")
            L.append("")
            L += [f"- {n}（{c}本）" for n, c in s["unresolved"]]
            L.append("")
    L.append("## 取り込んだファイル")
    L.append("")
    for a in area_list:
        s = report["areas"][a]
        L.append(f"- {a}: " + "、".join(f"{t}（{k}・{c}行）" for t, k, c in s["files"]))
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")


def main():
    ap = argparse.ArgumentParser(description="系統トポロジー（接続送電線・上位系設備）の生成")
    sub = ap.add_subparsers(dest="cmd", required=True)
    f = sub.add_parser("fetch", help="各社の変圧器・送電線CSVを取得")
    f.add_argument("--cache-dir", default="grid_cache")
    f.set_defaults(func=cmd_fetch)
    b = sub.add_parser("build", help="data/grid_topology.js を生成")
    b.add_argument("--in", dest="src", required=True)
    b.add_argument("--cache-dir", default="grid_cache")
    b.add_argument("--out", default="data/grid_topology.js")
    b.add_argument("--report", default="")
    b.add_argument("--list-sheet", default="全国変電所リスト")
    b.set_defaults(func=cmd_build)
    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
