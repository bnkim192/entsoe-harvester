"""
PSE 5개년 조정계획(Plan Koordynacyjny 5-letni) — 중앙급전 발전기(JWCD) 계획 정비 (2026-11-23 신규)

엔드포인트(2026-11-23 프로브로 확인, PSE EndpointsMap.pdf):
  unav-pk5l   Remonty — 발전소·발전기·정비 종류·시작·종료 등 (v_pk5l_*)
  pk5l-wp     기본 수치 — 계획 기간 수요·가용 용량·정지 합계 등
PSE API 는 $top 을 받지 않아 조건 없이 받고 @odata.nextLink 로 다음 쪽을 따라간다(최대 MAX_PAGES 쪽).
산출  pse_pk5l.json  {unav:{columns, rows}, wp:{columns, rows}, log}   (공개 자료만)
"""
import json, time
import datetime as dt
import requests

API = "https://api.raporty.pse.pl/api"
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "application/json"}
MAX_PAGES = 40
OUT = "pse_pk5l.json"


def pull(name, log):
    url, rows, page = "%s/%s" % (API, name), [], 0
    while url and page < MAX_PAGES:
        r = None
        for k in range(3):
            try:
                r = requests.get(url, headers=UA, timeout=120)
                if r.status_code == 200:
                    break
            except requests.RequestException as e:
                log.append({"name": name, "page": page, "error": type(e).__name__})
            time.sleep(5)
        if r is None or r.status_code != 200:
            log.append({"name": name, "page": page, "status": r.status_code if r is not None else None,
                        "head": r.text[:300] if r is not None else ""})
            break
        js = r.json()
        vals = js.get("value", [])
        rows += vals
        log.append({"name": name, "page": page, "rows": len(vals)})
        url = js.get("@odata.nextLink")
        page += 1
    cols = sorted({k for v in rows[:200] for k in v.keys()}) if rows else []
    return {"columns": cols, "rows": rows, "pages": page, "more": bool(url)}


def main():
    log = []
    out = {"fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
           "source": "PSE raporty.pse.pl API — Plan Koordynacyjny 5-letni"}
    out["unav"] = pull("unav-pk5l", log)
    out["wp"] = pull("pk5l-wp", log)
    out["log"] = log
    print("[pse] unav-pk5l %d줄(%d쪽, 더 있음 %s) · pk5l-wp %d줄" % (len(out["unav"]["rows"]), out["unav"]["pages"],
          out["unav"]["more"], len(out["wp"]["rows"])), flush=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1, default=str)


if __name__ == "__main__":
    main()
