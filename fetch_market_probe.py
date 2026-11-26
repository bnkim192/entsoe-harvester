"""
선물 시세 재확인 (2026-11-26): 독일 전력 선물 · TTF 가스 선물 · EUA 선물을 ICE 말고 다른 경로로 받을 수 있는지 시험
그리고 이미 되는 PSE 전망 · Open-Meteo 풍속도 한 번 더 확인.

경로마다 상태 코드, 응답 길이, 앞부분, 숫자가 들어 있는지(시세로 보이는 값)를 market_probe.json 에 남긴다.
여기서 "됨"으로 나온 경로만 다음에 정식 수집 코드로 만든다.

사용  python fetch_market_probe.py
산출  market_probe.json   (공개 자료만)
"""
import json, re, time
import datetime as dt
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "application/json, text/plain, */*", "Accept-Language": "en-US,en;q=0.9"}
TODAY = dt.date.today()
D0 = (TODAY - dt.timedelta(days=10)).isoformat()
D1 = TODAY.isoformat()
OUT = "market_probe.json"
RES = []


def yahoo(sym):
    return "https://query1.finance.yahoo.com/v8/finance/chart/%s?range=1mo&interval=1d" % sym


TESTS = [
    # 1) 독일 전력 선물
    ("de_futures", "energy-charts API 목록(openapi)", "https://api.energy-charts.info/openapi.json", {}),
    ("de_futures", "energy-charts 차트 화면", "https://www.energy-charts.info/charts/price_futures_bars/chart.htm?l=en&interval_select=month", {}),
    ("de_futures", "energy-charts 차트 데이터(월물)", "https://www.energy-charts.info/charts/price_futures_bars/data/de/month.json", {}),
    ("de_futures", "energy-charts 차트 데이터(월물, 올해)", "https://www.energy-charts.info/charts/price_futures_bars/data/de/month_%d.json" % TODAY.year, {}),
    ("de_futures", "energy-charts 현물 API(비교용)", "https://api.energy-charts.info/price?bzn=DE-LU&start=%s&end=%s" % (D0, D1), {}),
    # 2) TTF 가스 선물
    ("ttf", "Yahoo TTF 근월물(TTF=F)", yahoo("TTF=F"), {}),
    ("ttf", "ICE 리포트 282 화면", "https://www.ice.com/report/282", {}),
    ("ttf", "ICE 지연시세(이전 403 재확인)",
     "https://www.ice.com/marketdata/DelayedMarkets.shtml?getContractsAsJson=&productId=4331&hubId=7979",
     {"X-Requested-With": "XMLHttpRequest", "Referer": "https://www.ice.com/products/27996665/Dutch-TTF-Natural-Gas-Futures/data"}),
    # 3) EUA 선물
    ("eua", "ICE 지연시세 EUA(이전 403 재확인)",
     "https://www.ice.com/marketdata/DelayedMarkets.shtml?getContractsAsJson=&productId=390&hubId=564",
     {"X-Requested-With": "XMLHttpRequest", "Referer": "https://www.ice.com/products/197/EUA-Futures/data"}),
    ("eua", "EEX EUA 선물 화면", "https://www.eex.com/en/market-data/market-data-hub/environmentals/futures", {}),
    ("eua", "Yahoo 탄소 ETF KRBN(대용 지표)", yahoo("KRBN"), {}),
    ("eua", "Yahoo EUA 12월물 후보(CKZ26.NYM)", yahoo("CKZ26.NYM"), {}),
    ("eua", "Yahoo SparkChange 탄소 ETC(CO2.L)", yahoo("CO2.L"), {}),
    # 4) 폴란드 수요·풍력·태양광 전망 (이미 됨, 재확인)
    ("pse_fc", "PSE pk5l-wp", "https://api.raporty.pse.pl/api/pk5l-wp?$filter=business_date%%20eq%%20'%s'" % D1, {}),
    # 5) 과거 풍속 (이미 됨, 재확인)
    ("wind_hist", "Open-Meteo 과거 바르샤바",
     "https://archive-api.open-meteo.com/v1/archive?latitude=52.2297&longitude=21.0122&start_date=%s&end_date=%s"
     "&daily=wind_speed_100m_mean&wind_speed_unit=ms" % ((TODAY - dt.timedelta(days=15)).isoformat(), (TODAY - dt.timedelta(days=7)).isoformat()), {}),
]


def numbers_seen(text):
    """응답 안에 시세처럼 보이는 숫자가 몇 개인지 (판단 참고용)."""
    return len(re.findall(r"(?<![\d.])\d{1,4}\.\d{1,3}(?![\d.])", text[:200000]))


def main():
    ses = requests.Session()
    for item, name, url, extra in TESTS:
        rec = {"item": item, "name": name, "url": url}
        try:
            r = ses.get(url, headers=dict(UA, **extra), timeout=60)
            txt = r.text or ""
            rec.update({"status": r.status_code, "len": len(txt), "ctype": r.headers.get("content-type", ""),
                        "numbers": numbers_seen(txt), "head": txt[:400]})
            if "openapi" in url and r.status_code == 200:
                try:
                    rec["paths"] = sorted(r.json().get("paths", {}).keys())
                except ValueError:
                    pass
            if "yahoo" in url and r.status_code == 200:
                try:
                    res = r.json()["chart"]["result"][0]
                    q = res["indicators"]["quote"][0]["close"]
                    rec["last_close"] = [x for x in q if x is not None][-3:]
                    rec["meta"] = {k: res["meta"].get(k) for k in ("symbol", "currency", "exchangeName", "longName", "shortName")}
                except (KeyError, IndexError, TypeError, ValueError):
                    pass
            if "chart.htm" in url and r.status_code == 200:
                rec["json_refs"] = sorted(set(re.findall(r"[\w/\-.]+\.json", txt)))[:40]
        except requests.RequestException as e:
            rec["error"] = repr(e)[:300]
        RES.append(rec)
        print("%-10s %-36s %s %s" % (item, name, rec.get("status", "ERR"), rec.get("last_close", rec.get("len", ""))), flush=True)
        time.sleep(1)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"), "results": RES},
                  f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
