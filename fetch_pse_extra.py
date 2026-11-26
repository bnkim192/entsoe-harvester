"""
PSE 공개 API 추가 수집 (2026-11-26, WA법인 요청 2·3번 시험 수집)

1) pk5l-wp  : 5개년 조정계획 기본값 = 전력망 수요 전망 · 풍력 전망 · 태양광 전망 (오늘부터 15일)
2) pdwkseub : 15분 단위 수요 (요청 3번 PEAK 영향도 분석용)
              하루씩 받아 '하루 요약'(평균·최대·최대 시각·저녁 16~20시 평균)을 쌓고, 최근 14일은 15분 원자료도 남김

PSE 제약(fetch_pse_pk5l.py 에서 확인): 한 번에 최대 100줄, 다음 쪽 없음, 날짜는 문자열 비교('YYYY-MM-DD')
→ business_date 구간으로 나눠 받고 100줄이 차면 반으로 쪼갬. 하루가 100줄을 넘으면 시각 열로 다시 쪼갬.
열 이름은 아직 모르므로 받은 열 목록(columns)과 원자료를 그대로 남긴다.

사용  python fetch_pse_extra.py              → 전망 + 15분 수요 최근 14일 (매일용)
      python fetch_pse_extra.py 2023-01-01   → 15분 수요를 그날부터 다시 받음 (백필)
산출  pse_pk5l_wp.json, pse_demand15.json   (공개 자료만)
"""
import json, os, sys, time
import datetime as dt
import requests

API = "https://api.raporty.pse.pl/api"
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "application/json"}
EP_FC = "pk5l-wp"
EP_15 = "pdwkseub"
OUT_FC = "pse_pk5l_wp.json"
OUT_15 = "pse_demand15.json"
LOG = []
TODAY = dt.date.today()


def log(msg):
    print(msg, flush=True)


def query(ep, flt):
    """200 이면 줄 목록, 400·404 면 None(재시도 안 함), 그 밖의 오류는 3번까지 다시."""
    for k in range(3):
        try:
            r = requests.get("%s/%s" % (API, ep), params={"$filter": flt}, headers=UA, timeout=120)
            if r.status_code == 200:
                return r.json().get("value", [])
            LOG.append({"ep": ep, "filter": flt, "status": r.status_code, "head": r.text[:300]})
            if r.status_code in (400, 404):
                return None
        except (requests.RequestException, ValueError) as e:
            LOG.append({"ep": ep, "filter": flt, "error": repr(e)[:200]})
        time.sleep(3 * (k + 1))
    return None


def time_col(rows):
    for r in rows[:1]:
        for k in r:
            if "dtime" in k.lower() and "utc" not in k.lower() and "publication" not in k.lower():
                return k
        for k in r:
            if "dtime" in k.lower() or "period" in k.lower():
                return k
    return None


def pull(ep, d0, d1):
    """business_date 가 [d0, d1) 인 줄."""
    flt = "business_date ge '%s' and business_date lt '%s'" % (d0.isoformat(), d1.isoformat())
    rows = query(ep, flt)
    if rows is None:
        return None
    if len(rows) >= 100 and (d1 - d0).days > 1:
        mid = d0 + dt.timedelta(days=(d1 - d0).days // 2)
        time.sleep(0.3)
        a = pull(ep, d0, mid)
        b = pull(ep, mid, d1)
        return (a or []) + (b or [])
    if len(rows) >= 100:
        tc = time_col(rows)
        if tc:                                             # 하루가 100줄 이상 → 시각으로 6시간씩
            out = []
            for h0 in range(0, 24, 6):
                f2 = flt + " and %s ge '%s %02d:00:00' and %s lt '%s %02d:00:00'" % (
                    tc, d0.isoformat(), h0, tc, (d0 if h0 < 18 else d1).isoformat(), (h0 + 6) % 24)
                part = query(ep, f2)
                if part is None:
                    LOG.append({"ep": ep, "day": d0.isoformat(), "note": "시각 쪼개기 실패, 100줄만 남김"})
                    return rows
                if len(part) >= 100:
                    LOG.append({"ep": ep, "day": d0.isoformat(), "hour": h0, "note": "6시간에 100줄 이상, 일부 빠졌을 수 있음"})
                out += part
                time.sleep(0.3)
            return out
        LOG.append({"ep": ep, "day": d0.isoformat(), "note": "하루 100줄 이상, 시각 열 없음, 일부 빠졌을 수 있음"})
    return rows


def numeric_cols(rows):
    skip = ("id", "mrid", "business_date", "publication", "source")
    cols = []
    for k in (rows[0] if rows else {}):
        if any(s in k.lower() for s in skip) or "dtime" in k.lower() or "period" in k.lower():
            continue
        vals = [r.get(k) for r in rows[:50]]
        try:
            if any(v is not None for v in vals) and all(v is None or isinstance(float(v), float) for v in vals):
                cols.append(k)
        except (TypeError, ValueError):
            pass
    return cols


def day_summary(rows, tc, cols):
    """하루 요약: 열마다 평균 · 최대 · 최대 시각 · 저녁(16~20시) 평균."""
    s = {"n": len(rows)}
    for c in cols:
        pts = []
        for r in rows:
            try:
                pts.append(((r.get(tc) or "") if tc else "", float(r[c])))
            except (TypeError, ValueError, KeyError):
                pass
        if not pts:
            continue
        mx = max(pts, key=lambda p: p[1])
        eve = [v for t, v in pts if len(t) >= 13 and "16" <= t[11:13] <= "19"]
        s[c] = {"mean": round(sum(v for _, v in pts) / len(pts), 2), "max": round(mx[1], 2), "max_at": mx[0][11:16],
                "eve_mean": round(sum(eve) / len(eve), 2) if eve else None}
    return s


def run_forecast():
    rows = pull(EP_FC, TODAY - dt.timedelta(days=1), TODAY + dt.timedelta(days=16))
    if rows is None:
        log("[pk5l-wp] 받지 못함 (LOG 확인)")
        rows = []
    tc = time_col(rows)
    cols = numeric_cols(rows)
    by_day = {}
    for r in rows:
        d = ((r.get(tc) or "")[:10] if tc else None) or r.get("business_date")
        by_day.setdefault(d, []).append(r)
    daily = {d: day_summary(v, tc, cols) for d, v in sorted(by_day.items()) if d}
    out = {"fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
           "source": "PSE raporty.pse.pl API %s (Plan Koordynacyjny 5-letni, wielkosci podstawowe)" % EP_FC,
           "columns": list(rows[0].keys()) if rows else [], "time_col": tc, "numeric_cols": cols,
           "rows_total": len(rows), "daily": daily, "rows": rows, "log": [x for x in LOG if x.get("ep") == EP_FC]}
    with open(OUT_FC, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1, default=str)
    log("[pk5l-wp] %d줄 · %d일 · 시각 열 %s · 숫자 열 %s" % (len(rows), len(daily), tc, cols))


def run_demand15(start):
    old = {}
    if os.path.exists(OUT_15):
        try:
            with open(OUT_15, encoding="utf-8") as f:
                old = json.load(f)
        except ValueError:
            old = {}
    daily = old.get("daily", {})
    recent = {}
    tc_all, cols_all, cols_first = None, None, None
    d = start
    n_ok = n_empty = n_fail = 0
    t0 = time.time()
    while d <= TODAY:
        rows = pull(EP_15, d, d + dt.timedelta(days=1))
        if rows is None:
            n_fail += 1
            if n_fail >= 5 and n_ok == 0:
                log("[pdwkseub] 처음부터 5번 연속 실패, 중단 (LOG 확인)")
                break
        elif not rows:
            n_empty += 1
        else:
            n_ok += 1
            tc = time_col(rows)
            cols = numeric_cols(rows)
            if cols_first is None:
                cols_first = list(rows[0].keys())
            tc_all, cols_all = tc, cols
            daily[d.isoformat()] = day_summary(rows, tc, cols)
            if (TODAY - d).days <= 14:
                recent[d.isoformat()] = rows
        if (n_ok + n_empty + n_fail) % 100 == 0:
            log("   %s 까지 · 받음 %d · 빈 날 %d · 실패 %d · %.0f초" % (d, n_ok, n_empty, n_fail, time.time() - t0))
        d += dt.timedelta(days=1)
        time.sleep(0.2)
    days = sorted(daily)
    out = {"fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
           "source": "PSE raporty.pse.pl API %s (15분 수요)" % EP_15,
           "columns": cols_first or old.get("columns", []), "time_col": tc_all or old.get("time_col"),
           "numeric_cols": cols_all or old.get("numeric_cols", []),
           "first_day": days[0] if days else None, "last_day": days[-1] if days else None, "n_days": len(days),
           "summary_note": "열마다 mean=하루 평균, max=하루 최대 15분 값, max_at=최대 시각, eve_mean=16~20시 평균",
           "daily": {k: daily[k] for k in days}, "recent_15min": recent,
           "log": [x for x in LOG if x.get("ep") == EP_15][-200:]}
    with open(OUT_15, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1, default=str)
    log("[pdwkseub] %s ~ %s · %d일 (이번에 받음 %d · 빈 날 %d · 실패 %d) · 숫자 열 %s" %
        (out["first_day"], out["last_day"], len(days), n_ok, n_empty, n_fail, out["numeric_cols"]))


def main():
    start = dt.date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 else TODAY - dt.timedelta(days=14)
    try:
        run_forecast()
    except Exception as e:
        log("[pk5l-wp] 오류 %r" % e)
    try:
        run_demand15(start)
    except Exception as e:
        log("[pdwkseub] 오류 %r" % e)


if __name__ == "__main__":
    main()
