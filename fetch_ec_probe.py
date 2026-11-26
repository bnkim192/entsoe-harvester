"""
독일 전력 선물(energy-charts price_futures_bars) 시험 수집 (2026-11-26 v3)

futures.js 에서 찾은 주소 규칙:
  /charts/price_futures_bars/data/{국가}/power_{기간}_{YYYY}_{MM}_{DD}.json
  기간 = year · quarter · month · week, 날짜 = 거래일. 화면도 없으면 하루씩 최대 30일 전까지 거슬러 찾음.
기간마다 가장 최근 거래일 파일을 찾아 원자료(시리즈 이름·x축·값)를 ec_probe.json 에 저장한다 (공개 자료만).
원자료는 EEX 시세 기반이라 상시 업무용이면 이용조건 확인이 필요하다.
"""
import json, time
import datetime as dt
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "application/json, text/javascript, */*; q=0.01", "X-Requested-With": "XMLHttpRequest",
      "Referer": "https://www.energy-charts.info/charts/price_futures_bars/chart.htm?l=en&c=DE"}
CH = "https://www.energy-charts.info/charts/price_futures_bars/data"
OUT = "ec_probe.json"
TODAY = dt.date.today()
ses = requests.Session()


def find(country, interval):
    tried = []
    for back in range(0, 15):
        d = TODAY - dt.timedelta(days=back)
        url = "%s/%s/power_%s_%04d_%02d_%02d.json" % (CH, country, interval, d.year, d.month, d.day)
        try:
            r = ses.get(url, headers=UA, timeout=60)
        except requests.RequestException as e:
            tried.append({"url": url, "error": repr(e)[:150]})
            continue
        tried.append({"url": url, "status": r.status_code})
        if r.status_code == 200:
            try:
                return d.isoformat(), url, r.json(), tried
            except ValueError:
                return d.isoformat(), url, {"raw": r.text[:3000]}, tried
        time.sleep(0.3)
    return None, None, None, tried


def summarize(js):
    """시리즈마다 이름·x축·값을 짧게 정리 (전체 원자료도 따로 저장)."""
    out = []
    for s in (js if isinstance(js, list) else []):
        name = s.get("name")
        if isinstance(name, list):
            name = name[0]
        if isinstance(name, dict):
            name = name.get("en") or name.get("de") or str(name)
        out.append({"name": name, "x": s.get("xAxisValues"), "data": s.get("data"), "unit": s.get("unit") or s.get("yAxisTitle")})
    return out


def main():
    res = {"fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
           "rule": "/charts/price_futures_bars/data/{country}/power_{interval}_{YYYY}_{MM}_{DD}.json", "results": {}}
    for country in ("de", "pl"):
        for interval in ("month", "quarter", "year", "week"):
            day, url, js, tried = find(country, interval)
            key = "%s_%s" % (country, interval)
            res["results"][key] = {"trade_day": day, "url": url, "tried": tried[-3:],
                                   "series": summarize(js) if js is not None else None,
                                   "raw": js if (js is not None and len(json.dumps(js)) < 60000) else None}
            print(key, day, url, "시리즈 %d" % len(res["results"][key]["series"] or []), flush=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
