"""
폴란드 지하 가스 저장소 저장률 수집기 (2026-11-13)

1순위  GIE AGSI+ (유럽 가스 저장소 공식 공시 플랫폼, 폴란드 운영사 GSP 자료 포함)
       https://agsi.gie.eu/api?country=pl  — 무료 API 키 필요(헤더 x-key). GitHub Secrets 의 AGSI_KEY.
2순위  키가 없거나 실패하면 폴란드 저장소 운영사 사이트를 구조 확인용으로만 조사(값은 쓰지 않음)
       https://ipi.gasstoragepoland.pl/.../dane-operacyjne/

산출물
  pl_gas_storage.json        {daily: [{date, full(%), gas_twh, inj_gwh, wd_gwh, wgv_twh, trend(%p)}]}  — 기존 파일과 합쳐 쌓음
  gas_storage_probe.json     요청 결과 (상태코드·건수·오류, 2순위 조사 결과)

환경변수  AGSI_KEY (선택, 없으면 2순위 조사만), PLG_START (기본 2025-01-01)
"""
import os, json, re, time
import datetime as dt
import requests

KEY = os.environ.get("AGSI_KEY", "").strip()
START = os.environ.get("PLG_START", "").strip() or "2025-01-01"
AGSI = "https://agsi.gie.eu/api"
GSP = ("https://ipi.gasstoragepoland.pl/uslugi-magazynowania/dane-operacyjne-i-eksploatacyjne/dane-operacyjne/")
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
OUT = "pl_gas_storage.json"
PROBE = "gas_storage_probe.json"


def log(*a):
    print("[gas]", *a, flush=True)


def fnum(v):
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return None


def get_agsi(probe):
    today = dt.date.today().isoformat()
    rows, page, last = [], 1, 1
    hdr = dict(UA)
    if KEY:
        hdr["x-key"] = KEY
    while page <= last and page <= 20:
        try:
            r = requests.get(AGSI, params={"country": "pl", "from": START, "to": today, "size": 300, "page": page},
                             headers=hdr, timeout=(10, 60))
        except requests.RequestException as e:
            probe["agsi"] = {"error": type(e).__name__}
            return rows
        probe.setdefault("agsi_pages", []).append({"page": page, "status": r.status_code, "len": len(r.text)})
        if r.status_code != 200:
            probe["agsi"] = {"status": r.status_code, "head": r.text[:300], "key_set": bool(KEY)}
            return rows
        try:
            js = r.json()
        except ValueError:
            probe["agsi"] = {"status": r.status_code, "head": r.text[:300], "note": "JSON 아님"}
            return rows
        last = int(js.get("last_page") or 1)
        for d in js.get("data", []):
            day = d.get("gasDayStart")
            if not day:
                continue
            rows.append({"date": day, "full": fnum(d.get("full")), "gas_twh": fnum(d.get("gasInStorage")),
                         "inj_gwh": fnum(d.get("injection")), "wd_gwh": fnum(d.get("withdrawal")),
                         "wgv_twh": fnum(d.get("workingGasVolume")), "trend": fnum(d.get("trend"))})
        page += 1
        time.sleep(1)
    probe["agsi"] = {"status": 200, "rows": len(rows), "key_set": bool(KEY), "last_page": last}
    return rows


def probe_gsp(probe):
    try:
        r = requests.get(GSP, headers=UA, timeout=(10, 40))
        txt = r.text
        heads = re.findall(r"<th[^>]*>(.*?)</th>", txt, flags=re.S | re.I)[:30]
        probe["gsp"] = {"status": r.status_code, "len": len(txt), "tables": txt.lower().count("<table"),
                        "th": [re.sub(r"<[^>]+>|\s+", " ", h).strip() for h in heads],
                        "links": [u for u in re.findall(r'href="([^"]+)"', txt) if re.search(r"(?i)xls|csv|json|api|export", u)][:20]}
    except requests.RequestException as e:
        probe["gsp"] = {"error": type(e).__name__}


def main():
    probe = {"fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"), "start": START}
    rows = get_agsi(probe)
    if not rows:
        log("AGSI 자료 없음 (키 %s) → GSP 사이트 구조 조사" % ("있음" if KEY else "없음"))
        probe_gsp(probe)
    old = {}
    if os.path.exists(OUT):
        try:
            with open(OUT, encoding="utf-8") as f:
                old = {r["date"]: r for r in json.load(f).get("daily", [])}
        except (ValueError, KeyError):
            old = {}
    for r in rows:
        old[r["date"]] = r
    daily = [old[d] for d in sorted(old)]
    if daily:
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump({"source": "GIE AGSI+ (country=PL)", "units": {"full": "%", "gas_twh": "TWh", "inj_gwh": "GWh/일",
                       "wd_gwh": "GWh/일", "wgv_twh": "TWh", "trend": "%p(전일 대비)"}, "daily": daily}, f, ensure_ascii=False, indent=1)
        log("저장 %d일 · 마지막 %s 저장률 %s%%" % (len(daily), daily[-1]["date"], daily[-1].get("full")))
    with open(PROBE, "w", encoding="utf-8") as f:
        json.dump(probe, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
