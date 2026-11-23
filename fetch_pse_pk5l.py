"""
PSE 5개년 조정계획(Plan Koordynacyjny 5-letni) — 중앙급전 발전기(JWCD) 계획 정비 (2026-11-23, v2: 날짜 조건·쪽 넘김)

엔드포인트(2026-11-23 확인): unav-pk5l (정비 일정), pk5l-wp (기본 수치)
  unav-pk5l 열: business_date, power_plant, unit_code, type(POS=정지 등), reason(RK 대정비·RS 중정비·RB 일상·RA 긴급 등),
               state, start_dtime, end_dtime, publication_ts, mrid_zas, direction
v1 은 조건 없이 받아 2024-06 자료 100줄만 왔다 → v2 는
  1) 최근 계획일(business_date) 조건, 2) 시작·종료일 조건 등 여러 $filter 형태를 시험하고,
  2) 100줄이 꽉 차면 @odata.nextLink 또는 $skip 으로 다음 쪽을 받는다.
  받은 줄을 합쳐 중복을 빼고, 2027년에 걸친 정비만 따로 정리한다.
산출  pse_pk5l.json  {tries, unav_rows(최신 계획 기준), y2027, wp_sample}   (공개 자료만)
"""
import json, time
import datetime as dt
import requests

API = "https://api.raporty.pse.pl/api"
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "application/json"}
MAX_PAGES = 60
OUT = "pse_pk5l.json"
REASON = {"RK": "대정비", "RS": "중정비", "RB": "일상 정비", "RA": "긴급 정비"}


def get(url, params=None):
    for k in range(3):
        try:
            r = requests.get(url, params=params, headers=UA, timeout=120)
            return r
        except requests.RequestException:
            time.sleep(5)
    return None


def pull(name, flt, log, tag):
    """조건 하나로 끝까지(쪽 넘김) 받는다."""
    rows, page, skip, prev_first = [], 0, 0, None
    params = {"$filter": flt} if flt else {}
    url = "%s/%s" % (API, name)
    while page < MAX_PAGES:
        r = get(url, params)
        if r is None or r.status_code != 200:
            log.append({"tag": tag, "filter": flt, "page": page, "status": r.status_code if r is not None else None,
                        "head": (r.text[:200] if r is not None else "")})
            break
        try:
            js = r.json()
        except ValueError:
            log.append({"tag": tag, "filter": flt, "page": page, "head": r.text[:200]})
            break
        vals = js.get("value", [])
        first = json.dumps(vals[0], sort_keys=True) if vals else None
        if first is not None and first == prev_first:     # $skip 이 무시되어 같은 쪽이 다시 오면 멈춤
            log.append({"tag": tag, "filter": flt, "page": page, "note": "같은 쪽 반복 → 중단"})
            break
        prev_first = first
        rows += vals
        nxt = js.get("@odata.nextLink")
        log.append({"tag": tag, "filter": flt, "page": page, "rows": len(vals), "next": bool(nxt)})
        page += 1
        if nxt:
            url, params = nxt, None
        elif len(vals) >= 100:                            # 다음 쪽 링크가 없으면 $skip 시도
            skip += len(vals)
            params = dict({"$filter": flt} if flt else {}, **{"$skip": str(skip)})
        else:
            break
    return rows


def key(r):
    return (r.get("mrid_zas"), r.get("unit_code"), r.get("start_dtime"), r.get("end_dtime"), r.get("publication_ts"))


def main():
    log, tries = [], []
    today = dt.datetime.now(dt.timezone.utc).date()
    cands = []
    for k in range(0, 8):
        d = (today - dt.timedelta(days=k)).isoformat()
        cands += [("bd_q_%s" % d, "business_date eq '%s'" % d), ("bd_%s" % d, "business_date eq %s" % d)]
    cands += [("start_2027_q", "start_dtime ge '2027-01-01 00:00:00'"), ("end_2027_q", "end_dtime ge '2027-01-01 00:00:00'"),
              ("start_2027", "start_dtime ge 2027-01-01T00:00:00"), ("bd_ge_q", "business_date ge '%s'" % (today - dt.timedelta(days=7)).isoformat())]
    allrows = {}
    for tag, flt in cands:
        rs = pull("unav-pk5l", flt, log, tag)
        tries.append({"tag": tag, "filter": flt, "rows": len(rs),
                      "max_end": max((r.get("end_dtime") or "" for r in rs), default=None),
                      "bd": sorted({r.get("business_date") for r in rs})[-3:] if rs else []})
        for r in rs:
            allrows[key(r)] = r
        time.sleep(1)
    rows = list(allrows.values())
    # 최신 계획: business_date 가 가장 최근인 줄들
    bds = sorted({r.get("business_date") for r in rows if r.get("business_date")})
    latest = bds[-1] if bds else None
    cur = [r for r in rows if r.get("business_date") == latest] if latest else rows
    y27 = [r for r in cur if (r.get("end_dtime") or "") >= "2027-01-01" and (r.get("start_dtime") or "") < "2028-01-01"]
    y27.sort(key=lambda r: (r.get("start_dtime") or "", r.get("power_plant") or ""))
    out = {"fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
           "source": "PSE raporty.pse.pl API unav-pk5l (Plan Koordynacyjny 5-letni, Remonty)",
           "reason_codes": REASON, "tries": tries, "latest_business_date": latest, "rows_total": len(rows),
           "unav_rows": cur, "y2027": [{"plant": r.get("power_plant"), "unit": r.get("unit_code"), "type": r.get("type"),
                                       "reason": r.get("reason"), "reason_ko": REASON.get(r.get("reason"), r.get("reason")),
                                       "start": r.get("start_dtime"), "end": r.get("end_dtime"), "state": r.get("state"),
                                       "published": r.get("publication_ts")} for r in y27],
           "log": log[-120:]}
    wp = pull("pk5l-wp", "business_date eq '%s'" % latest if latest else None, log, "wp") if latest else []
    out["wp_rows"] = wp[:2000]
    print("[pse] 최신 계획일 %s · 줄 %d · 2027년 정비 %d건 · 기본수치 %d줄" % (latest, len(cur), len(y27), len(wp)), flush=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1, default=str)


if __name__ == "__main__":
    main()
