#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
変電所ごとの「上位系統の状態」「要接続検討」「接続検討の実績の分類」「上位側の送電線」「上位系の制約設備」を、
地図モード（substation.html）と同じ判定ルールで計算する（CSV払い出し用）。

data/substations.js（本体行・実績DB）と data/grid_topology.js（上位系統のつながり）を読む。
判定ルールは substation.html の upperStatus / needCheck / memoUpper / lineUpStatus / ehvOf と同じ内容を
Python に写したもの。**画面側のルールを変えたらここも同じに直す**（両方の件数が一致することを確認する）。

特高パターン（2026-11-20〜）: ehv() が「想定電圧（66kV以上で最も低い母線）の上位系設備の受入余力」を
公表の空容量 → 公表のN-1電制適用可能量 → 余裕(運用容量−|予想潮流|・推定) の優先順で求める。
閾値・案件規模の既定値は substation_scoring.py の EHV_* が単一ソース。

  python scripts/grid_status.py            # 判定の件数を表示（画面の表示件数との突き合わせ用）
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import substation_scoring as S  # noqa: E402  配点・閾値・特高の定義の単一ソース

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOPO_JS = os.path.join(ROOT, "data", "grid_topology.js")
SUBS_JS = os.path.join(ROOT, "data", "substations.js")

STATUS_LABEL = {
    "ok": "上位系に空きあり",
    "part": "上位側に空きあり（この変電所の上位系値は未公表）",
    "bad": "上位系がダメ（空き無し）",
    "check": "要接続検討（上位系は0だが当該設備に空き／実績で上位系対策なし）",
    "unk": "上位系統が不明",
    "none": "空容量データ無し",
}
MEMO_LABEL = {"none": "対策なし", "bank": "バンク等の対策待ち", "yes": "上位系対策あり", "unk": "記載なし"}
GS_NAME = {"both": "緑", "upunk": "黄", "upbad": "紫", "ownbad": "赤", "none": "灰"}
GS_LABEL = {
    "both": "変電所も上位系も空きあり",
    "upunk": "変電所は空きあり・上位系は不明",
    "upbad": "変電所は空きあり・上位系は空き無し（要接続検討）",
    "ownbad": "変電所単体で空き無し",
    "none": "空容量データ無し",
}


def _load_js(path):
    raw = open(path, encoding="utf-8").read()
    return json.loads(raw[raw.index("{"): raw.rindex("}") + 1])


def fmt(v):
    return "—" if v is None else str(v)


class GridStatus:
    def __init__(self, subs_js=SUBS_JS, topo_js=TOPO_JS):
        d = _load_js(subs_js)
        self.F = {n: i for i, n in enumerate(d["fields"])}
        self.dict = d["dict"]
        self.rows = d["rows"]
        self.results = d["results"]
        s1 = d.get("params", {}).get("s1Steps") or []
        self.need_check_mw = s1[1][0] if len(s1) > 1 else 6
        self.grid_rank = d.get("params", {}).get("gridRank", {"S": 60, "A": 48, "B": 36})
        # 最低必要容量（画面の MIN_MW と同じ）。これ未満の空容量は「空き無し」
        self.min_mw = d.get("params", {}).get("minMw", 2)
        self.memo = {}
        self.result_index = {id(r): i for i, r in enumerate(self.results)}
        for r in self.results:
            self.memo.setdefault(r[0] + str(r[1]).strip(), []).append(r)
        self.row_by_no = {str(r[self.F["no"]]): r for r in self.rows}
        self.topo = _load_js(topo_js) if os.path.exists(topo_js) else None
        self.node_lines = {}
        if self.topo:
            self.LF = {n: i for i, n in enumerate(self.topo["lineFields"])}
            self.BF = {n: i for i, n in enumerate(self.topo["bankFields"])}
            for i, l in enumerate(self.topo["lines"]):
                for e in (l[self.LF["from"]], l[self.LF["to"]]):
                    if isinstance(e, int):
                        self.node_lines.setdefault(e, []).append(i)
        self._cache = {}

    # ── 行アクセス ──
    def v(self, row, k):
        return row[self.F[k]]

    def area(self, row):
        return self.dict["area"][row[self.F["area"]]]

    def rank(self, row):
        return self.dict["rank"][row[self.F["rank"]]]

    def has_data(self, row):
        return self.v(row, "availUp") is not None or self.v(row, "availCur") is not None

    def memo_of(self, row):
        return self.memo.get(self.area(row) + str(self.v(row, "name")).strip(), [])

    def node_of(self, row):
        if not self.topo:
            return None
        return self.topo["byNo"].get(str(self.v(row, "no")))

    def node_rows(self, ni):
        return [self.row_by_no[str(no)] for no in self.topo["nodes"][ni][2] if str(no) in self.row_by_no]

    # ── 送電線・バンク ──
    def line_cap(self, l):
        return l[self.LF["availUp"]] if l[self.LF["availUp"]] is not None else l[self.LF["availCur"]]

    def line_kv(self, l):
        return l[self.LF["kv"]]

    def bank_cap(self, b):
        return b[self.BF["availUp"]] if b[self.BF["availUp"]] is not None else b[self.BF["availCur"]]

    def no_room(self, v):
        return v < self.min_mw

    def node_bad(self, ni):
        return any(self.bank_cap(b) is not None and self.no_room(self.bank_cap(b)) for b in self.topo["nodes"][ni][3])

    def node_links(self, ni):
        out, seen = [], {}
        for li in self.node_lines.get(ni, []):
            l = self.topo["lines"][li]
            other = l[self.LF["to"]] if l[self.LF["from"]] == ni else l[self.LF["from"]]
            key = f"{l[self.LF['name']]}|{'n' + str(other) if isinstance(other, int) else 't' + str(other)}"
            if key in seen:
                seen[key]["n"] += 1
                continue
            seen[key] = {"line": l, "li": li, "other": other, "n": 1}
            out.append(seen[key])
        out.sort(key=lambda x: -(self.line_kv(x["line"]) or 0))
        return out

    def vmax_of_node(self, ni):
        v1 = [b[self.BF["v1"]] for b in self.topo["nodes"][ni][3] if b[self.BF["v1"]] is not None]
        if v1:
            return max(v1)
        rows = self.node_rows(ni)
        return (self.v(rows[0], "vmax") or 0) if rows else 0

    def node_status_for_kv(self, ni, kv):
        banks = self.topo["nodes"][ni][3]
        if not banks:
            return {"st": "unk", "cap": None}
        use = banks
        if kv is not None:
            use = [b for b in banks if b[self.BF["v2"]] is not None and abs(b[self.BF["v2"]] - kv) < 3]
            if not use:
                return {"st": "peer", "cap": None}
        caps = [self.bank_cap(b) for b in use if self.bank_cap(b) is not None]
        if not caps:
            return {"st": "unk", "cap": None}
        m = max(caps)
        return {"st": "bad" if self.no_room(m) else "ok", "cap": m}

    def line_up_status(self, l, from_ni):
        c = self.line_cap(l)
        if c is not None:
            return {"st": "bad" if self.no_room(c) else "ok", "cap": c, "via": "送電線"}
        kv = self.line_kv(l)
        ends = [e for e in (l[self.LF["from"]], l[self.LF["to"]]) if isinstance(e, int) and e != from_ni]
        if not ends:
            return {"st": "unk", "cap": None, "via": ""}
        up = ends[0] if len(ends) == 1 else (ends[0] if self.vmax_of_node(ends[0]) >= self.vmax_of_node(ends[1]) else ends[1])
        r = self.node_status_for_kv(up, kv)
        return {"st": r["st"], "cap": r["cap"], "via": f"{self.topo['nodes'][up][1]}の該当バンク", "up": up}

    # ── 特高パターン（2026-11-20）。想定電圧の上位系設備（送電線・上位変圧器）の受入余力を出す ──
    #   画面（substation.html の ehvOf）と同じルール。**どちらかを変えたら必ず両方直す**。
    #   優先順: 公表の空容量 → 公表のN-1電制適用可能量 → 余裕（運用容量−|予想潮流|。推定）
    def _ehv_room(self, pub_cap, n1amt, opcap, flow):
        """1設備の受入余力を (値, 出所) で返す。公表値が無ければ余裕（推定）。"""
        if pub_cap is not None:
            return pub_cap, S.EHV_SRC_CAP
        if isinstance(n1amt, (int, float)) and not isinstance(n1amt, bool):
            return float(n1amt), S.EHV_SRC_N1
        if (isinstance(opcap, (int, float)) and not isinstance(opcap, bool)
                and isinstance(flow, (int, float)) and not isinstance(flow, bool)):
            return float(opcap) - abs(float(flow)), S.EHV_SRC_EST
        return None, None

    def ehv_kv(self, row, ni):
        """特高連系の想定電圧(kV)。その変電所にある EHV_KV_MIN 以上の母線のうち最も低い電圧。"""
        vs = set()
        for b in self.topo["nodes"][ni][3]:
            for k in ("v1", "v2"):
                v = b[self.BF[k]]
                if isinstance(v, (int, float)) and v >= S.EHV_KV_MIN:
                    vs.add(float(v))
        v = self.v(row, "vmax")
        if isinstance(v, (int, float)) and v >= S.EHV_KV_MIN:
            vs.add(float(v))
        return min(vs) if vs else None

    def ehv(self, row):
        """特高の受入余力。戻り値は
        {"kv", "cap", "src", "what", "n", "over", "min_est", "items"} か {"why": 計算できない理由}。
        """
        key = ("ehv", id(row))
        if key in self._cache:
            return self._cache[key]
        r = self._ehv(row)
        self._cache[key] = r
        return r

    def _ehv(self, row):
        if not self.topo:
            return {"why": "系統トポロジーのデータが無い"}
        if self.area(row) not in self.topo["meta"]["areas"]:
            return {"why": f"{self.area(row)}エリアは上位系の公表データを取得できず対象外（東京・中部・沖縄）"}
        ni = self.node_of(row)
        if ni is None:
            return {"why": "公表物にこの変電所と同名のノードが無く、つながりをたどれない"}
        kv = self.ehv_kv(row, ni)
        if kv is None:
            return {"why": f"{S.EHV_KV_MIN}kV以上の母線が見つからない（22/33kVのみなど）"}
        items = []
        for x in self.node_links(ni):
            l = x["line"]
            lkv = self.line_kv(l)
            if not isinstance(lkv, (int, float)) or abs(lkv - kv) >= 3:
                continue
            room, src = self._ehv_room(self.line_cap(l), l[self.LF["n1amt"]], l[self.LF["opcap"]], l[self.LF["flow"]])
            if room is None and isinstance(x["other"], int):
                # 送電線自体に公表値が無ければ、相手側（上位側）変電所の同電圧バンクで代用する
                for b in self.topo["nodes"][x["other"]][3]:
                    v2 = b[self.BF["v2"]]
                    if isinstance(v2, (int, float)) and abs(v2 - kv) < 3:
                        room, src = self._ehv_room(self.bank_cap(b), None, b[self.BF["opcap"]], b[self.BF["flow"]])
                        if room is not None:
                            items.append({"what": f"{self.topo['nodes'][x['other']][1]}の{kv:g}kVバンク",
                                          "room": room, "src": src})
                            break
                continue
            if room is not None:
                items.append({"what": f"{l[self.LF['name']]}（{kv:g}kV）", "room": room, "src": src})
        for b in self.topo["nodes"][ni][3]:
            v2 = b[self.BF["v2"]]
            if isinstance(v2, (int, float)) and abs(v2 - kv) < 3:
                room, src = self._ehv_room(self.bank_cap(b), None, b[self.BF["opcap"]], b[self.BF["flow"]])
                if room is not None:
                    items.append({"what": f"この変電所の{fmt(b[self.BF['v1']])}/{kv:g}kVバンク", "room": room, "src": src})
        if not items:
            return {"why": f"{kv:g}kVの設備に公表データ（空容量・N-1電制適用可能量・運用容量と予想潮流）が無い", "kv": kv}
        # 公表値（空容量→N-1電制適用可能量）がある設備を優先し、同じ段の中では最も余力の大きい設備を根拠にする
        pick = None
        for src in (S.EHV_SRC_CAP, S.EHV_SRC_N1, S.EHV_SRC_EST):
            cands = [i for i in items if i["src"] == src]
            if cands:
                pick = max(cands, key=lambda i: i["room"])
                break
        ests = [i["room"] for i in items if i["src"] == S.EHV_SRC_EST]
        return {"kv": kv, "cap": pick["room"], "src": pick["src"], "what": pick["what"], "n": len(items),
                "over": sum(1 for e in ests if e < 0), "min_est": min(ests) if ests else None, "items": items}

    def ehv_judge(self, row, need_mw=None):
        e = self.ehv(row)
        return S.ehv_judge(e.get("cap"), need_mw)

    def ehv_reason(self, row, need_mw=None):
        """特高の判定の理由。画面のツールチップ・CSVの理由列と同じ文言。"""
        need = S.EHV_MW_DEFAULT if need_mw is None else need_mw
        e = self.ehv(row)
        if "cap" not in e:
            return e.get("why", "")
        head = (f"想定電圧 {e['kv']:g}kV の上位系設備 {e['n']}件のうち受入余力が最大のもの＝{e['what']} で "
                f"{e['cap']:g}MW（{S.EHV_SRC_LABEL[e['src']]}）")
        if e["cap"] >= need:
            tail = f"。案件規模 {need:g}MW 以上の余力あり"
        elif e["cap"] > 0:
            tail = f"。案件規模 {need:g}MW に届かない（規模を落とすか増強前提か人が判断）"
        else:
            tail = "。余力0以下（予想潮流が運用容量を超えている等）"
        if e["over"]:
            tail += f"。{S.EHV_OVER_TAG}：同じ電圧の設備 {e['over']}件で予想潮流が運用容量を超えている（最小 {e['min_est']:g}MW）"
        if e["src"] == S.EHV_SRC_EST:
            tail += "。根拠が推定値なので接続検討で必ず確認する"
        return head + tail

    def ehv_items_text(self, row):
        """想定電圧の設備ごとの受入余力（CSV用）。"""
        e = self.ehv(row)
        if "items" not in e:
            return ""
        return " / ".join(f"{i['what']}[{i['room']:g}MW・{S.EHV_SRC_LABEL[i['src']]}]"
                          for i in sorted(e["items"], key=lambda i: -i["room"]))

    # ── 実績 ──
    @staticmethod
    def memo_upper(m):
        t = " ".join(str(x or "") for x in (m[6], m[7], m[8]))
        mm = re.search(r"上位系[^:：]*[:：]\s*(.*)$", t)
        after = mm.group(1) if mm else ""
        head = after or t
        est = "変電所は推定" in t
        cls = "unk"
        if re.match(r"^(なし|無し)", after) or re.search(r"上位系(統)?(増強|対策)(工事)?(なし|無し|0円)", t):
            cls = "bank" if re.search(r"別途|バンク逆潮流対策工事あり|バンク逆潮流対策（|変圧器の更新|変電設備工事|配電設備工事|工程表に上位系統|上位系統の現場調査|対策待ち", head) else "none"
        elif (re.match(r"^あり", after) or re.search(r"上位系(統)?(増強|対策)(工事)?\s*[:：]", t)
              or re.search(r"上位系(統)?増強(で|が必要)", t) or re.search(r"(変圧器|変電所)(の)?増強(が必要|（)", head)):
            cls = "yes"
        return cls, est

    def memo_summary(self, row):
        ms = self.memo_of(row)
        c = {"n": len(ms), "none": 0, "bank": 0, "yes": 0, "unk": 0, "noneSure": 0, "est": 0}
        for m in ms:
            cls, est = self.memo_upper(m)
            c[cls] += 1
            if est:
                c["est"] += 1
            elif cls == "none":
                c["noneSure"] += 1
        return c

    def need_check(self, row):
        why = []
        cur = self.v(row, "availCur")
        if cur is not None and cur >= self.need_check_mw:
            why.append(f"当該設備の空容量 {cur}MW（≧{self.need_check_mw}MW）")
        c = self.memo_summary(row)
        if c["noneSure"]:
            why.append(f"接続検討の実績で上位系・バンク等の対策なしの回答 {c['noneSure']}件")
        return why or None

    # ── 上位側の送電線（判定に使うもの。同階級の変電所どうしを結ぶ線は除く）──
    def up_lines(self, row):
        ni = self.node_of(row)
        if ni is None:
            return []
        banks = self.topo["nodes"][ni][3]
        v1s = [b[self.BF["v1"]] for b in banks if b[self.BF["v1"]] is not None]
        v1max = max(v1s) if v1s else None
        out = []
        for x in self.node_links(ni):
            kv = self.line_kv(x["line"])
            if not isinstance(x["other"], int):
                continue
            if v1max is not None and not (kv is not None and kv >= v1max):
                continue
            u = self.line_up_status(x["line"], ni)
            if u["st"] == "peer":
                continue
            out.append({**x, "u": u})
        return out

    def upper_status(self, row):
        key = id(row)
        if key in self._cache:
            return self._cache[key]
        ups = self.up_lines(row)
        bad = [x for x in ups if x["u"]["st"] == "bad"]
        ok = [x for x in ups if x["u"]["st"] == "ok"]

        def desc(x):
            u = x["u"]
            return f"{x['line'][self.LF['name']]}（{fmt(self.line_kv(x['line']))}kV）→ {u['via']}" + (f" {u['cap']}MW" if u["cap"] is not None else "")

        up = self.v(row, "availUp")
        ni = self.node_of(row)
        if up is not None and self.no_room(up):
            r = {"cls": "bad", "reason": f"この変電所の空容量(上位系考慮)が {up}MW" + (f"。上位側も空き無し: {'、'.join(map(desc, bad))}" if bad else "")}
        elif up is not None:
            r = {"cls": "ok", "reason": f"この変電所の空容量(上位系考慮)が {up}MW" + (
                f"（ただし上位側に空き{fmt(self.min_mw)}MW未満の設備あり: {'、'.join(map(desc, bad))}）" if bad else (f"。上位側も空きあり: {'、'.join(map(desc, ok))}" if ok else ""))}
        elif bad and not ok:
            r = {"cls": "bad", "reason": f"上位側の変電所・送電線に空き無し: {'、'.join(map(desc, bad))}。この変電所の「空容量(上位系考慮)」は未公表"}
        elif bad and ok:
            r = {"cls": "part", "reason": f"上位側に空きのある系統とない系統が混在。空きあり: {'、'.join(map(desc, ok))}／空き無し: {'、'.join(map(desc, bad))}"}
        elif ok:
            r = {"cls": "part", "reason": f"上位側に空きあり: {'、'.join(map(desc, ok))}。この変電所の「空容量(上位系考慮)」は未公表"}
        elif self.has_data(row):
            r = {"cls": "unk", "reason": "「空容量(上位系考慮)」が未公表" + ("。上位側の変電所・送電線の空容量も非公表" if ups else ("。上位側の送電線がつながりデータに無い" if ni is not None else "。つながりデータなし"))}
        else:
            r = {"cls": "none", "reason": "送配電事業者の空容量が未取得。上位側の空容量も非公表" if ups else "送配電事業者の空容量が未取得"}
        r["label"] = STATUS_LABEL[r["cls"]]
        self._cache[key] = r
        return r

    # ── 整理した判定：変電所単体（当該設備）の空き × 上位系の空き → 5区分（画面の gridStatus と同じ）──
    def own_status(self, row):
        c, u = self.v(row, "availCur"), self.v(row, "availUp")
        if c is not None:
            return {"st": "bad" if self.no_room(c) else "ok", "cap": c}
        if u is not None and not self.no_room(u):
            return {"st": "ok", "cap": u}
        return {"st": "unk", "cap": None}

    def grid_status(self, row):
        key = ("gs", id(row))
        if key in self._cache:
            return self._cache[key]
        o, u = self.own_status(row), self.upper_status(row)
        up = "ok" if u["cls"] in ("ok", "part") else "bad" if u["cls"] == "bad" else "unk"
        if o["st"] == "bad":
            cls = "ownbad"
        elif o["st"] == "ok":
            cls = "both" if up == "ok" else "upbad" if up == "bad" else "upunk"
        else:
            cls = "ownbad" if up == "bad" else "none"
        own_text = f"空きあり {o['cap']}MW" if o["st"] == "ok" else f"空き無し {o['cap']}MW" if o["st"] == "bad" else "非公表"
        up_text = {"ok": "空きあり", "bad": "空き無し"}.get(up, "不明")
        nc = self.need_check(row) if cls == "upbad" else None
        reason = f"変電所単体（当該設備）: {own_text}／上位系: {up_text}（{u['reason']}）"
        if cls == "upbad":
            reason += ("。" + "、".join(nc) + " → 実績で上位系0でも通った例（王子・由仁・曽根・引野）と同じ条件") if nc else "。接続検討で上位系の扱いを確認する価値あり"
        r = {"cls": cls, "own": o, "up": up, "label": GS_LABEL[cls], "reason": reason}
        self._cache[key] = r
        return r

    # ── CSV向けの文字列 ──
    def up_lines_text(self, row):
        out = []
        for x in self.up_lines(row):
            u, l = x["u"], x["line"]
            other = self.topo["nodes"][x["other"]][1]
            st = {"ok": f"空きあり{fmt(u['cap'])}MW", "bad": f"空き無し{fmt(u['cap'])}MW"}.get(u["st"], "非公表")
            out.append(f"{l[self.LF['name']]}({fmt(self.line_kv(l))}kV)→{other}[{st}]" + (f"×{x['n']}" if x["n"] > 1 else ""))
        return " / ".join(out)

    def links_count(self, row):
        ni = self.node_of(row)
        return len(self.node_links(ni)) if ni is not None else None

    def upper_equipment_text(self, row):
        ni = self.node_of(row)
        if ni is None:
            return ""
        labels, seen = [], set()
        for b in self.topo["nodes"][ni][3]:
            for k, v in b[self.BF["upRefs"]]:
                if k == "L":
                    l = self.topo["lines"][v]
                    c = self.line_cap(l)
                    t = f"{l[self.LF['name']]}({fmt(self.line_kv(l))}kV)" + ("▼" if c is not None and self.no_room(c) else "")
                elif k == "N":
                    t = self.topo["nodes"][v][1] + ("▼" if self.node_bad(v) else "")
                else:
                    t = str(v)
                if t not in seen:
                    seen.add(t)
                    labels.append(t)
        return "、".join(labels)

    def memo_text(self, row):
        c = self.memo_summary(row)
        if not c["n"]:
            return ""
        parts = [f"{MEMO_LABEL[k]}{c[k]}" for k in ("none", "bank", "yes", "unk") if c[k]]
        return f"{c['n']}件（" + "・".join(parts) + (f"・うち変電所は推定{c['est']}" if c["est"] else "") + "）"

    def rank_meaning(self, rk):
        """系統ランク（80点満点）の意味。画面の rankMeaning と同じ文言"""
        g = self.grid_rank
        return {
            "S": f"S：系統スコア{g['S']}点以上（最優先で接続検討を申し込む）",
            "A": f"A：系統スコア{g['A']}〜{g['S'] - 1}点（接続検討の候補）",
            "B": f"B：系統スコア{g['B']}〜{g['A'] - 1}点（近くにS・Aが無いときの次点）",
            "C": f"C：系統スコア{g['B'] - 1}点以下（見送り推奨）",
            "データ無": "データ無：送配電事業者の空容量が未公表・未取得で採点できない",
        }.get(rk, "")

    def category_text(self, row):
        """地図と同じカテゴリ（5色）。例「紫：変電所は空きあり・上位系は空き無し（要接続検討）」"""
        g = self.grid_status(row)
        return f"{GS_NAME[g['cls']]}：{g['label']}"

    def memo_detail_text(self, row):
        """接続検討の実績の全項目（回答ごとに1行。画面の memoDetailText と同じ）"""
        out = []
        for m in self.memo_of(row):
            cls, est = self.memo_upper(m)
            i = self.result_index.get(id(m), -1)
            out.append(f"#{i + 1}【{MEMO_LABEL[cls]}{'・変電所は推定' if est else ''}】{m[8]}／案件:{m[2]}／出力:{m[3]}／"
                       f"負担金:{m[4]}／工期:{m[5]}／対策:{m[6]}／教訓・上位系:{m[7]}")
        return "\n".join(out)

    def topo_available(self, row):
        if not self.topo:
            return ""
        if self.area(row) not in self.topo["meta"]["areas"]:
            return "対象外エリア"
        return "あり" if self.node_of(row) is not None else "なし（公表物に同名なし）"


def main():
    gs = GridStatus()
    from collections import Counter
    c = Counter(gs.grid_status(r)["cls"] for r in gs.rows if gs.v(r, "lat") is not None)
    print("座標あり行の判定（変電所単体×上位系）:", dict(c))
    c2 = Counter(gs.upper_status(r)["cls"] for r in gs.rows if gs.v(r, "lat") is not None)
    print("座標あり行の上位系統の状態:", dict(c2))
    m = Counter()
    for r in gs.results:
        cls, est = gs.memo_upper(r)
        m[cls + ("(推定)" if est else "")] += 1
    print("実績の分類:", dict(m))
    print("要接続検討（紫）:", sum(1 for r in gs.rows if gs.grid_status(r)["cls"] == "upbad"),
          "実績あり:", sum(1 for r in gs.rows if gs.memo_of(r)))
    # 特高パターン（画面の件数と一致することを確認する）
    coord = [r for r in gs.rows if gs.v(r, "lat") is not None]
    src = Counter(S.EHV_SRC_LABEL[gs.ehv(r)["src"]] for r in gs.rows if "cap" in gs.ehv(r))
    print(f"特高の受入余力が出せた行: {sum(1 for r in gs.rows if 'cap' in gs.ehv(r))}件（根拠 {dict(src)}）")
    for mw in S.EHV_MW_OPTIONS:
        c = Counter(S.EHV_JUDGE_LABEL[S.ehv_judge(gs.ehv(r).get("cap"), mw)] for r in coord)
        print(f"  座標あり行の特高判定（案件規模 {mw}MW）:", dict(c))


if __name__ == "__main__":
    main()
