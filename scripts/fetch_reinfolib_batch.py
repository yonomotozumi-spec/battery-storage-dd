#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
複数地点ぶんの reinfolib（不動産情報ライブラリ）GeoJSON をまとめて取得し、
scripts/reinfolib_judge.py の judge を各地点に回して values.json を書き出す。

  export REINFOLIB_KEY=<APIキー>
  python scripts/fetch_reinfolib_batch.py --geo geo.json --outdir reinfolib_values

geo.json は {"1": {"lat":..., "lon":...}, ...}。取得したタイルは
<outdir>/<No>/<CODE>.geojson に、判定結果は <outdir>/<No>.json に保存する。
APIキーは環境変数からのみ読み、ファイルにも出力にも残さない。
"""
import argparse, json, os, subprocess, sys, time, urllib.request, urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
JUDGE = os.path.join(HERE, "reinfolib_judge.py")

def fetch(url, key, retries=3):
    req = urllib.request.Request(url, headers={"Ocp-Apim-Subscription-Key": key})
    for n in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=40) as r:
                return r.read().decode("utf-8")
        except urllib.error.HTTPError as e:
            if e.code in (401, 403):
                raise SystemExit(f"認証エラー({e.code})。APIキーを確認してください。")
            if e.code == 429:
                time.sleep(5 * (n + 1)); continue
            if n == retries - 1:
                return ""
            time.sleep(2 * (n + 1))
        except Exception:
            if n == retries - 1:
                return ""
            time.sleep(2 * (n + 1))
    return ""

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--geo", required=True)
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--sleep", type=float, default=0.7, help="リクエスト間隔(秒)")
    a = ap.parse_args()

    key = os.environ.get("REINFOLIB_KEY", "").strip()
    if not key:
        raise SystemExit("環境変数 REINFOLIB_KEY にAPIキーを設定してください。")

    geo = json.load(open(a.geo, encoding="utf-8"))
    os.makedirs(a.outdir, exist_ok=True)

    for no, g in geo.items():
        lat, lon = g["lat"], g["lon"]
        d = os.path.join(a.outdir, str(no))
        os.makedirs(d, exist_ok=True)
        urls = subprocess.run([sys.executable, JUDGE, "urls", "--lat", str(lat), "--lon", str(lon)],
                              capture_output=True, text=True, check=True).stdout.strip().splitlines()
        got = 0
        for line in urls:
            code, url = line.split("\t", 1)
            body = fetch(url, key)
            open(os.path.join(d, code + ".geojson"), "w", encoding="utf-8").write(body)
            if body.strip():
                got += 1
            time.sleep(a.sleep)
        subprocess.run([sys.executable, JUDGE, "judge", "--lat", str(lat), "--lon", str(lon),
                        "--dir", d, "--out", os.path.join(a.outdir, f"{no}.json")], check=True)
        vals = json.load(open(os.path.join(a.outdir, f"{no}.json"), encoding="utf-8")).get("values", {})
        print(f"No.{no} {lat},{lon}  取得{got}/{len(urls)}レイヤ  判定{len(vals)}項目", flush=True)

if __name__ == "__main__":
    main()
