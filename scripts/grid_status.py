#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
変電所ごとの「上位系統の状態」「要接続検討」「接続検討の実績の分類」「上位側の送電線」「上位系の制約設備」を、
地図モード（substation.html）と同じ判定ルールで計算する（CSV払い出し用）。

data/substations.js（本体行・実績DB）と data/grid_topology.js（上位系統のつながり）を読む。
判定ルールは substation.html の upperStatus / needCheck / memoUpper / lineUpStatus と同じ内容を Python に写したもの。
**画面側のルールを変えたらここも同じに直す**（両方の件数が一致することを確認する）。

  python scripts/grid_status.py            # 判定の件数を表示（画面の表示件数との突き合わせ用）
"""
import json
import os
import re
import sys

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
        self.memo = {}
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

    def node_bad(self, ni):
        return any(self.bank_cap(b) is not None and self.bank_cap(b) <= 0 for b in self.topo["nodes"][ni][3])

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
        return {"st": "bad" if m <= 0 else "ok", "cap": m}

    def line_up_status(self, l, from_ni):
        c = self.line_cap(l)
        if c is not None:
            return {"st": "bad" if c <= 0 else "ok", "cap": c, "via": "送電線"}
        kv = self.line_kv(l)
        ends = [e for e in (l[self.LF["from"]], l[self.LF["to"]]) if isinstance(e, int) and e != from_ni]
        if not ends:
            return {"st": "unk", "cap": None, "via": ""}
        up = ends[0] if len(ends) == 1 else (ends[0] if self.vmax_of_node(ends[0]) >= self.vmax_of_node(ends[1]) else ends[1])
        r = self.node_status_for_kv(up, kv)
        return {"st": r["st"], "cap": r["cap"], "via": f"{self.topo['nodes'][up][1]}の該当バンク", "up": up}

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
        if self.rank(row) != "除外":
            return None
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
        if up is not None and up <= 0:
            r = {"cls": "bad", "reason": f"この変電所の空容量(上位系考慮)が {up}MW" + (f"。上位側も空き無し: {'、'.join(map(desc, bad))}" if bad else "")}
        elif up is not None:
            r = {"cls": "ok", "reason": f"この変電所の空容量(上位系考慮)が {up}MW" + (
                f"（ただし上位側に空き0の設備あり: {'、'.join(map(desc, bad))}）" if bad else (f"。上位側も空きあり: {'、'.join(map(desc, ok))}" if ok else ""))}
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
        if r["cls"] == "bad":
            nc = self.need_check(row)
            if nc:
                r = {"cls": "check", "reason": r["reason"] + "。ただし " + "、".join(nc) + " → 接続検討を申し込む価値あり（実績: 王子・由仁・曽根・引野は上位系0でも上位系対策なし）"}
        r["label"] = STATUS_LABEL[r["cls"]]
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
                    t = f"{l[self.LF['name']]}({fmt(self.line_kv(l))}kV)" + ("▼" if c is not None and c <= 0 else "")
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

    def topo_available(self, row):
        if not self.topo:
            return ""
        if self.area(row) not in self.topo["meta"]["areas"]:
            return "対象外エリア"
        return "あり" if self.node_of(row) is not None else "なし（公表物に同名なし）"


def main():
    gs = GridStatus()
    from collections import Counter
    c = Counter(gs.upper_status(r)["cls"] for r in gs.rows if gs.v(r, "lat") is not None)
    print("座標あり行の上位系統の状態:", dict(c))
    m = Counter()
    for r in gs.results:
        cls, est = gs.memo_upper(r)
        m[cls + ("(推定)" if est else "")] += 1
    print("実績の分類:", dict(m))
    print("要接続検討:", sum(1 for r in gs.rows if gs.upper_status(r)["cls"] == "check"),
          "実績あり:", sum(1 for r in gs.rows if gs.memo_of(r)))


if __name__ == "__main__":
    main()
