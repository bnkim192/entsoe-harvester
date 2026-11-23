"""
PSE 5개년 조정계획(Plan Koordynacyjny 5-letni) — 중앙급전 발전기(JWCD) 계획 정비 (2026-11-23 v3)

엔드포인트: unav-pk5l (정비 일정). 열: business_date(= 정비 시작일), power_plant, unit_code, type(POS=정지),
            reason(RK 대정비·RS 중정비·RB 일상·RA 긴급), state, start_dtime, end_dtime, publication_ts, mrid_zas
확인된 제약(2026-11-23): 한 번에 최대 100줄, @odata.nextLink 없음, $skip·$top 불가, 문자열 날짜 비교만 됨('YYYY-MM-DD').
→ business_date 구간으로 나눠 받고, 100줄이 차면 구간을 반으로 쪼개 다시 받는다.
기간: 2026-12-01 ~ 2028-01-01 시작분 + 2026년 하반기에 시작해 2027년까지 이어지는 정비
산출  pse_pk5l.json  {y2027:[...], dec2026:[...], rows, log}   (공개 자료만)
"""
import json, time
import datetime as dt
import requests

API = "https://api.raporty.pse.pl/api/unav-pk5l"
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "application/json"}
OUT = "pse_pk5l.json"
REASON = {"RK": "대정비", "RS": "중정비", "RB": "일상 정비", "RA": "긴급 정비"}
LOG = []


def query(flt):
    for k in range(3):
        try:
            r = requests.get(API, params={"$filter": flt}, headers=UA, timeout=120)
            if r.status_code == 200:
                return r.json().get("value", [])
            LOG.append({"filter": flt, "status": r.status_code, "head": r.text[:200]})
            return None
        except (requests.RequestException, ValueError) as e:
            LOG.append({"filter": flt, "error": type(e).__name__})
            time.sleep(5)
    return None


def pull_range(d0, d1, extra=""):
    """business_date 가 [d0, d1) 인 줄. 100줄이면 반으로 쪼갠다."""
    flt = "business_date ge '%s' and business_date lt '%s'%s" % (d0.isoformat(), d1.isoformat(), extra)
    rows = query(flt)
    if rows is None:
        return []
    LOG.append({"from": d0.isoformat(), "to": d1.isoformat(), "rows": len(rows)})
    if len(rows) >= 100 and (d1 - d0).days > 1:
        mid = d0 + dt.timedelta(days=(d1 - d0).days // 2)
        time.sleep(0.5)
        return pull_range(d0, mid, extra) + pull_range(mid, d1, extra)
    if len(rows) >= 100:
        LOG.append({"from": d0.isoformat(), "note": "하루에 100줄 이상 — 일부 빠졌을 수 있음"})
    return rows


def main():
    rows = []
    m = dt.date(2026, 12, 1)
    while m < dt.date(2028, 1, 1):                       # 2026-12 ~ 2027-12 시작분, 달 단위
        n = dt.date(m.year + (m.month == 12), (m.month % 12) + 1, 1)
        rows += pull_range(m, n)
        m = n
        time.sleep(0.5)
    # 2026-07 ~ 2026-11 에 시작해 2027년까지 이어지는 것
    rows += pull_range(dt.date(2026, 7, 1), dt.date(2026, 12, 1), " and end_dtime ge '2027-01-01 00:00:00'")
    # 중복 제거: 같은 정비(mrid·발전기·시작)는 가장 최근 공시만
    best = {}
    for r in rows:
        k = (r.get("mrid_zas"), r.get("unit_code"), r.get("start_dtime"))
        if k not in best or (r.get("publication_ts") or "") > (best[k].get("publication_ts") or ""):
            best[k] = r
    cur = sorted(best.values(), key=lambda r: (r.get("start_dtime") or "", r.get("power_plant") or ""))

    def slim(r):
        return {"plant": r.get("power_plant"), "unit": r.get("unit_code"), "type": r.get("type"), "reason": r.get("reason"),
                "reason_ko": REASON.get(r.get("reason"), r.get("reason")), "state": r.get("state"),
                "start": r.get("start_dtime"), "end": r.get("end_dtime"), "published": r.get("publication_ts")}
    y27 = [slim(r) for r in cur if (r.get("end_dtime") or "") >= "2027-01-01" and (r.get("start_dtime") or "") < "2028-01-01"]
    dec = [slim(r) for r in cur if (r.get("end_dtime") or "") >= "2026-12-01" and (r.get("start_dtime") or "") < "2027-01-01"]
    out = {"fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
           "source": "PSE raporty.pse.pl API unav-pk5l (Plan Koordynacyjny 5-letni, Remonty)",
           "reason_codes": REASON, "rows_total": len(cur), "y2027": y27, "dec2026": dec,
           "states": sorted({r.get("state") for r in cur if r.get("state")}),
           "types": sorted({r.get("type") for r in cur if r.get("type")}), "log": LOG}
    print("[pse] 정비 %d건 · 2027년에 걸친 정비 %d건 · 2026년 12월 %d건" % (len(cur), len(y27), len(dec)), flush=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1, default=str)


if __name__ == "__main__":
    main()
