"""
폴란드 5개 도시 100m 풍속 (2026-11-26, WA법인 요청 2번 풍속 예보 시험 수집)

도시  바르샤바(동) · 슈체친(서) · 브로츠와프(남) · 그단스크(북) · 포즈난(중앙)  (WA법인이 준 좌표)
1) 과거  Open-Meteo Archive(ERA5) 시간별 wind_speed_100m → 하루 평균. 최근 며칠은 Forecast API past_days 로 메움
2) 전망  Open-Meteo Seasonal API(ECMWF 계절예보) daily wind_speed_100m_mean, 앙상블 멤버 평균 = 그날 풍속
         (웹 화면에서 XLSX 를 받아 한 행의 평균을 쓰던 방법과 같은 값)
단위 m/s. 5개 도시 단순 평균(avg)도 함께 저장.
주의 Open-Meteo 무료 API 는 비상업용 조건. 상시 업무용이면 유료 키 검토 필요.

사용  python fetch_pl_wind5.py              → 전망 + 과거 최근 30일 (매일용)
      python fetch_pl_wind5.py 2023-01-01   → 과거를 그날부터 다시 받음 (백필)
산출  pl_wind5.json   (공개 자료만)
"""
import json, os, sys, time
import datetime as dt
import requests

CITIES = [("warszawa", "바르샤바", 52.2297, 21.0122), ("szczecin", "슈체친", 53.4285, 14.5528),
          ("wroclaw", "브로츠와프", 51.1079, 17.0385), ("gdansk", "그단스크", 54.3520, 18.6466),
          ("poznan", "포즈난", 52.4064, 16.9252)]
ARCH = "https://archive-api.open-meteo.com/v1/archive"
FCST = "https://api.open-meteo.com/v1/forecast"
SEAS = "https://seasonal-api.open-meteo.com/v1/seasonal"
OUT = "pl_wind5.json"
TZ = "Europe/Warsaw"
LOG = []
TODAY = dt.date.today()


def log(msg):
    print(msg, flush=True)


def get(url, params):
    for k in range(3):
        try:
            r = requests.get(url, params=params, timeout=180)
            if r.status_code == 200:
                return r.json()
            LOG.append({"url": url, "params": {a: b for a, b in params.items() if a not in ("latitude", "longitude")},
                        "status": r.status_code, "head": r.text[:300]})
            if r.status_code == 400:
                return None
        except (requests.RequestException, ValueError) as e:
            LOG.append({"url": url, "error": repr(e)[:200]})
        time.sleep(5 * (k + 1))
    return None


def hourly_to_daily(js):
    h = (js or {}).get("hourly", {})
    acc = {}
    for t, v in zip(h.get("time", []), h.get("wind_speed_100m", [])):
        if v is not None:
            acc.setdefault(t[:10], []).append(v)
    return {d: round(sum(v) / len(v), 2) for d, v in acc.items() if len(v) >= 20}


def run_hist(start):
    old = {}
    if os.path.exists(OUT):
        try:
            with open(OUT, encoding="utf-8") as f:
                old = json.load(f)
        except ValueError:
            old = {}
    hist = old.get("hist", {})
    end_arch = TODAY - dt.timedelta(days=6)
    for key, ko, lat, lon in CITIES:
        got = {}
        y = start
        while y <= end_arch:                                    # 1년씩
            y2 = min(dt.date(y.year, 12, 31), end_arch)
            js = get(ARCH, {"latitude": lat, "longitude": lon, "start_date": y.isoformat(), "end_date": y2.isoformat(),
                            "hourly": "wind_speed_100m", "wind_speed_unit": "ms", "timezone": TZ})
            got.update(hourly_to_daily(js))
            y = y2 + dt.timedelta(days=1)
            time.sleep(1)
        js = get(FCST, {"latitude": lat, "longitude": lon, "past_days": 14, "forecast_days": 1,
                        "hourly": "wind_speed_100m", "wind_speed_unit": "ms", "timezone": TZ})
        for d, v in hourly_to_daily(js).items():
            if d < TODAY.isoformat() and d not in got:           # archive 가 없는 최근 며칠만
                got[d] = v
        for d, v in got.items():
            hist.setdefault(d, {})[key] = v
        log("[과거] %s %d일" % (ko, len(got)))
    for d, row in hist.items():
        vs = [row[k] for k, *_ in CITIES if row.get(k) is not None]
        row["avg"] = round(sum(vs) / len(vs), 2) if len(vs) == len(CITIES) else None
    return {d: hist[d] for d in sorted(hist)}


def run_seasonal():
    days = {}
    meta = {}
    for key, ko, lat, lon in CITIES:
        base = {"latitude": lat, "longitude": lon, "daily": "wind_speed_100m_mean", "wind_speed_unit": "ms", "timezone": TZ}
        js = get(SEAS, dict(base, forecast_days=183))
        if js is None:
            js = get(SEAS, base)
        if not js:
            log("[전망] %s 받지 못함" % ko)
            continue
        dd = js.get("daily", {})
        mem = [k for k in dd if k.startswith("wind_speed_100m_mean") and "member" in k]
        if not mem:                                         # 멤버가 없으면 평균 열 하나만
            mem = [k for k in dd if k.startswith("wind_speed_100m_mean")]
        meta[key] = {"members": len(mem), "first": (dd.get("time") or [None])[0], "last": (dd.get("time") or [None])[-1]}
        for i, t in enumerate(dd.get("time", [])):
            vs = [dd[k][i] for k in mem if i < len(dd[k]) and dd[k][i] is not None]
            if vs:
                days.setdefault(t, {})[key] = round(sum(vs) / len(vs), 2)
        log("[전망] %s %s ~ %s · 멤버 %d" % (ko, meta[key]["first"], meta[key]["last"], len(mem)))
        time.sleep(1)
    for d, row in days.items():
        vs = [row[k] for k, *_ in CITIES if row.get(k) is not None]
        row["avg"] = round(sum(vs) / len(vs), 2) if len(vs) == len(CITIES) else None
    return {"fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"), "cities": meta,
            "days": {d: days[d] for d in sorted(days)}}


def main():
    start = dt.date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 else TODAY - dt.timedelta(days=30)
    hist, seas = {}, {}
    try:
        hist = run_hist(start)
    except Exception as e:
        log("[과거] 오류 %r" % e)
    try:
        seas = run_seasonal()
    except Exception as e:
        log("[전망] 오류 %r" % e)
    out = {"fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
           "source": "Open-Meteo Archive(ERA5) · Forecast · Seasonal(ECMWF) API, wind_speed_100m",
           "unit": "m/s (하루 평균)", "cities": [{"key": k, "name": n, "lat": a, "lon": b} for k, n, a, b in CITIES],
           "note": "avg = 5개 도시 단순 평균. 전망은 앙상블 멤버 평균. 무료 API 는 비상업용 조건",
           "hist": hist, "seasonal": seas, "log": LOG[-100:]}
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    hd = sorted(hist)
    log("[pl_wind5] 과거 %s ~ %s %d일 · 전망 %d일" % (hd[0] if hd else None, hd[-1] if hd else None, len(hd),
                                                  len((seas or {}).get("days", {}))))


if __name__ == "__main__":
    main()
