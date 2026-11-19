"""
폴란드 과거 겨울(10~1월) 바람·수요·가격 한 번 받기 (2026-11-19 신규, 12월 판단용 과거 비교)

기간  2021~2024년 겨울(10~1월) 4번 + 2025-10-01 ~ 2025-12-31
      (2026-01-01 이후는 pl_drivers_daily.json 에 이미 있음)
자료  ENTSO-E  A44 PL 하루전 가격(EUR) · A65 수요 · A75 전원별 발전(풍력 육상·해상, 태양광 등)
      Open-Meteo Archive(ERA5)  폴란드 3개 지점 100m 풍속·기온 (pl_drivers_daily 와 같은 지점)
      ECB  EUR→PLN 기준환율
산출  pl_winter_hist.json  {daily:[...], monthly:[...]}   ※ 공개 자료만, 회사 정보 없음
      fetch_pl_drivers.py 의 수집 함수를 그대로 쓰고, 기간만 바꾼다.
환경변수  ENTSOE_TOKEN (GitHub Secrets)
"""
import json, time
import datetime as dt
import requests
import fetch_pl_drivers as fd

PERIODS = [("2021-10-01", "2022-02-01"), ("2022-10-01", "2023-02-01"), ("2023-10-01", "2024-02-01"),
           ("2024-10-01", "2025-02-01"), ("2025-10-01", "2026-01-01")]   # 2026-11-19 2021·2022 겨울 추가(에너지 위기 시기 포함)
OUT = "pl_winter_hist.json"


def wins():
    out = []
    for a, b in PERIODS:
        s, e = dt.datetime.strptime(a, "%Y-%m-%d"), dt.datetime.strptime(b, "%Y-%m-%d")
        while s < e:
            t = min(s + dt.timedelta(days=31), e)
            out.append((s, t))
            s = t
    return out


fd.windows = wins          # fetch_points 가 이 기간만 받도록


def in_periods(d):
    return any(a <= d < b for a, b in PERIODS)


def weather():
    keys = (("wind_speed_100m", "wind100"), ("temperature_2m", "temp"))
    acc = {k: {} for _, k in keys}
    for nm, lat, lon in fd.WX_POINTS:
        for k in range(3):
            try:
                r = requests.get("https://archive-api.open-meteo.com/v1/archive",
                                 params={"latitude": lat, "longitude": lon, "timezone": "Europe/Warsaw",
                                         "start_date": PERIODS[0][0], "end_date": "2025-12-31",
                                         "hourly": ",".join(s for s, _ in keys)}, timeout=120)
                if r.status_code == 200:
                    break
                fd.log("   [기상 %s] HTTP %s (시도 %d)" % (nm, r.status_code, k + 1))
            except Exception as e:
                fd.log("   [기상 %s] 실패 %s (시도 %d)" % (nm, type(e).__name__, k + 1))
            time.sleep(10)
        else:
            continue
        h = r.json().get("hourly", {})
        tl = h.get("time", [])
        for src, key in keys:
            vals = h.get(src) or []
            byd = {}
            for i, ts in enumerate(tl):
                v = vals[i] if i < len(vals) else None
                if v is not None:
                    byd.setdefault(ts[:10], []).append(v)
            for d, vs in byd.items():
                if len(vs) >= fd.MIN_HOURS and in_periods(d):
                    acc[key].setdefault(d, []).append(sum(vs) / len(vs))
        fd.log("[기상 %-8s] 시각 %d" % (nm, len(tl)))
    return {k: {d: sum(v) / len(v) for d, v in dd_.items() if len(v) == len(fd.WX_POINTS)} for k, dd_ in acc.items()}


def ecb_pln():
    try:
        r = requests.get("https://data-api.ecb.europa.eu/service/data/EXR/D.PLN.EUR.SP00.A",
                         params={"startPeriod": "2021-09-20", "endPeriod": "2026-01-05", "format": "csvdata"}, timeout=60)
        out = {}
        lines = r.text.splitlines()
        hdr = lines[0].split(",")
        it, iv = hdr.index("TIME_PERIOD"), hdr.index("OBS_VALUE")
        for ln in lines[1:]:
            c = ln.split(",")
            if len(c) > max(it, iv) and c[iv]:
                out[c[it]] = float(c[iv])
        return out
    except Exception as e:
        fd.log("   [ECB] 실패 %r" % e)
        return {}


def main():
    pts = fd.merged(fd.fetch_points({"documentType": "A44", "in_Domain": fd.PL, "out_Domain": fd.PL}, "A44 PL", "price.amount"))
    px = fd.dmean(fd.by_day(pts))
    fd.log("[가격 PL ] 일수 %d" % len(px))
    ld = fd.get_load()
    gn = fd.get_gen()
    wx = weather()
    fx = ecb_pln()
    fk = sorted(fx)

    def fx_on(d):
        prior = [k for k in fk if k <= d]
        return fx[prior[-1]] if prior else None

    days = sorted(d for d in set(px) | set(ld["load"]) if in_periods(d))
    rows = []
    for d in days:
        w_on, w_off = gn.get("wind_on", {}).get(d), gn.get("wind_off", {}).get(d)
        wind = (w_on or 0) + (w_off or 0) if (w_on is not None or w_off is not None) else None
        p, f = px.get(d), fx_on(d)
        rows.append({"date": d, "px_PL": p, "px_PL_pln": p * f if (p is not None and f) else None, "eurpln": f,
                     "load": ld["load"].get(d), "wind": wind, "wind_on": w_on, "wind_off": w_off,
                     "solar": gn.get("solar", {}).get(d), "hardcoal": gn.get("hardcoal", {}).get(d),
                     "lignite": gn.get("lignite", {}).get(d), "gas": gn.get("gas", {}).get(d),
                     "wind100": wx.get("wind100", {}).get(d), "temp": wx.get("temp", {}).get(d)})
    mon = {}
    for r in rows:
        mon.setdefault(r["date"][:7], []).append(r)

    def m(vs, k, div=1.0):
        xs = [v[k] for v in vs if v.get(k) is not None]
        return round(sum(xs) / len(xs) / div, 2) if xs else None
    monthly = [{"month": k, "days": len(v), "px_PL": m(v, "px_PL"), "px_PL_pln": m(v, "px_PL_pln"),
                "load_gw": m(v, "load", 1000), "wind_gw": m(v, "wind", 1000), "solar_gw": m(v, "solar", 1000),
                "wind100": m(v, "wind100"), "temp": m(v, "temp")} for k, v in sorted(mon.items())]
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"generated": dt.datetime.now(fd.TZ).strftime("%Y-%m-%d %H:%M"), "periods": PERIODS,
                   "units": {"px_PL": "EUR/MWh", "px_PL_pln": "PLN/MWh(ECB)", "load·wind·solar": "MW 일평균",
                             "wind100": "km/h (3지점 평균)", "temp": "°C"},
                   "collect": fd.LOG, "monthly": monthly, "daily": rows}, f, ensure_ascii=False, indent=1)
    for x in monthly:
        fd.log("%s 가격 %s PLN · 풍력 %s GW · 수요 %s GW · 풍속 %s" % (x["month"], x["px_PL_pln"], x["wind_gw"], x["load_gw"], x["wind100"]))


if __name__ == "__main__":
    main()
