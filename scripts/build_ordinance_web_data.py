#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""条例リスクDBのWeb用データ（data/ordinances.js）を生成する。

ordinance_registry.json（台帳・定点）、kaihatsu_kyokasha.json（開発許可権者151者）、
priority_munis.json（S/Aランク変電所の所在自治体）を1本のJSにまとめ、
ordinance.html が file:// でも読めるよう window.ORDINANCE_DATA に代入する。

台帳・権者リスト・優先自治体を更新したら再生成すること:
  python scripts/build_ordinance_web_data.py
"""

import json
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "data" / "ordinances.js"


def main():
    reg = json.loads((ROOT / "data" / "ordinance_registry.json").read_text(encoding="utf-8"))
    kk = json.loads((ROOT / "data" / "kaihatsu_kyokasha.json").read_text(encoding="utf-8"))
    pm_path = ROOT / "data" / "priority_munis.json"
    pm = json.loads(pm_path.read_text(encoding="utf-8")) if pm_path.exists() else {"munis": []}

    # 市街化調整区域が実質NG（基準未整備・不許可・設置不可）のエントリを判定
    def chosei_ng(e):
        text = " ".join(filter(None, [e.get("requirement"), e.get("summary"), e.get("dd_impact")]))
        return any(k in text for k in ("設置不可", "原則不許可", "許可基準未整備", "基準は未整備", "許可基準が存在しない", "基準未策定", "許可対象外"))

    data = {
        "meta": {
            "generated": date.today().isoformat(),
            "registry_updated": reg["updated"],
            "authorities_updated": kk["updated"],
            "initialImport": reg.get("initial_import"),
        },
        "entries": [
            {
                "id": e["id"], "level": e["level"], "pref": e["pref"], "muni": e["muni"],
                "regType": e["reg_type"], "title": e["title"], "url": e["url"],
                "announced": e["announced_date"], "effective": e["effective_date"],
                "scope": e["scope"], "requirement": e["requirement"],
                "summary": e["summary"], "ddImpact": e["dd_impact"],
                "verified": e["last_verified"], "notes": e.get("notes"),
                "added": e.get("added", e["last_verified"]),
                "choseiNg": chosei_ng(e),
            }
            for e in reg["entries"]
        ],
        "watchpoints": reg.get("watchpoints", []),
        "authorities": [
            {"pref": a["pref"], "name": a["name"], "type": a["type"],
             "status": a["status"], "url": a["url"], "lastSwept": a["last_swept"]}
            for a in kk["authorities"]
        ],
        "priority": [
            {"pref": m["pref"], "muni": m["muni"], "S": m["S"], "A": m["A"]}
            for m in pm["munis"]
        ],
    }

    js = (
        "// 自動生成ファイル - 直接編集しないこと\n"
        "// 生成: scripts/build_ordinance_web_data.py\n"
        "// 元データ: data/ordinance_registry.json, data/kaihatsu_kyokasha.json, data/priority_munis.json\n"
        "window.ORDINANCE_DATA=" + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n"
    )
    OUT.write_text(js, encoding="utf-8")
    print(f"生成: {OUT.relative_to(ROOT)}  台帳{len(data['entries'])}件 / 権者{len(data['authorities'])}者 / 優先{len(data['priority'])}自治体")


if __name__ == "__main__":
    main()
