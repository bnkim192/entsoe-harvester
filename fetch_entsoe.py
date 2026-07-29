# -*- coding: utf-8 -*-
"""ENTSO-E 폴란드 day-ahead → 월평균 EUR/MWh. Secret: ENTSOE_TOKEN"""
import os, sys, json, csv
import datetime as dt
import xml.etree.ElementTree as ET
import requests

TOKEN = os.environ.get("ENTSOE_TOKEN", "").strip()
ZONE = "10YPL-AREA-----S"
BASE = "https://web-api.tp.entsoe.eu/api"
START_YEAR = 2020

def fetch_year(y):
    params = {"securityToken": TOKEN, "documentType": "A44",
              "in_Domain": ZONE, "out_Domain": ZONE,
              "periodStart": f"{y}01010000", "periodEnd": f"{y+1}01010000"}
    r = requests.get(BASE, params=params, timeout=120)
    print(f"[{y}] HTTP {r.status_code} len={len(r.text)}", flush=True)
    if r.status_code != 200:
        print("[BODY]", r.text[:600], flush=True); return []
    return parse(r.text)

def parse(xml):
    root = ET.fromstring(xml)
    ns = {"ns": root.tag.split("}")[0].strip("{")} if "}" in root.tag else {}
    fa = lambda el, tag: el.findall("ns:"+tag, ns) if ns else el.findall(tag)
    def tx(el, tag):
        f = el.find("ns:"+tag, ns) if ns else el.find(tag)
        return f.text if f is not None else None
    pts = []
    for ts in fa(root, "TimeSeries"):
        for per in fa(ts, "Period"):
            ti = per.find("ns:timeInterval", ns) if ns else per.find("timeInterval")
            start = tx(ti, "start"); res = tx(per, "resolution") or "PT60M"
            step = 15 if "PT15M" in res else 60
            t0 = dt.datetime.strptime(start[:16], "%Y-%m-%dT%H:%M")
            for pt in fa(per, "Point"):
                pos = int(tx(pt, "position")); price = float(tx(pt, "price.amount"))
                pts.append((t0 + dt.timedelta(minutes=step*(pos-1)), price))
    return pts

def main():
    if not TOKEN:
        print("!! ENTSOE_TOKEN 없음", flush=True); sys.exit(1)
    now = dt.datetime.utcnow(); allpts = []
    for y in range(START_YEAR, now.year+1):
        allpts += fetch_year(y)
    if not allpts:
        print("!! 0 수집", flush=True); sys.exit(1)
    agg = {}
    for t, p in allpts:
        agg.setdefault(f"{t.year}-{t.month:02d}", []).append(p)
    series = [{"ym": k, "eur_mwh": round(sum(v)/len(v),2), "hours": len(v)} for k, v in sorted(agg.items())]
    print("[월 수]", len(series), " 최신", series[-1], flush=True)
    out = {"source":"ENTSO-E A44 PL day-ahead","zone":"PL","unit":"EUR/MWh monthly avg","months":len(series),"series":series}
    with open("pl_wholesale_monthly.json","w",encoding="utf-8") as f: json.dump(out,f,ensure_ascii=False,indent=2)
    with open("pl_wholesale_monthly.csv","w",newline="",encoding="utf-8-sig") as f:
        w=csv.writer(f); w.writerow(["ym","eur_mwh","hours"])
        for s in series: w.writerow([s["ym"],s["eur_mwh"],s["hours"]])
    print("[OK] pl_wholesale_monthly.json/.csv", flush=True)

if __name__ == "__main__": main()
