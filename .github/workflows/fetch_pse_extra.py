"""
PSE 공개 API 추가 수집 (2026-11-26 v4, WA법인 요청 2번)

pk5l-wp : 5개년 조정계획 기본값 = 전력망 수요 전망(grid_demand_fcst) · 풍력 전망(fcst_wi_tot_gen)
          · 태양광 전망(fcst_pv_tot_gen) 등, 시간별, 오늘부터 약 15~18일

v4 변경
 - pdwkseub(15분 자료) 수집 삭제: 받아 보니 수요가 아니라 발전기별 15분 가용용량이었고,
   발전기 수만큼 줄이 생겨 하루 수천 줄 → 80분 넘게 걸림. 15분 수요 PEAK 는 ENTSO-E 실수요로 따로 계산
 - 하루 요약에서 grid_demand_fcst 가 빠지던 오류 수정 ('grid' 안의 'id' 글자를 id 열로 잘못 봄)

사용  python fetch_pse_extra.py   (몇 초면 끝남, 인자는 받아도 무시)
산출  pse_pk5l_wp.json   (공개 자료만)
"""
import json, time
import datetime as dt
import requests

API = "https://api.raporty.pse.pl/api"
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "application/json"}
EP_FC = "pk5l-wp"
OUT_FC = "pse_pk5l_wp.json"
LOG = []
TODAY = dt.date.today()
SKIP_WORDS = {"id", "mrid", "source"}                  # 열 이름을 '_' 로 나눈 단어 기준
SKIP_PREFIX = ("business_date", "publication", "period", "plan_dtime", "dtime")


def log(msg):
    print(msg, flush=True)


def query(ep, flt):
    """200 이면 줄 목록, 400·404 면 None(재시도 안 함), 그 밖의 오류는 3번까지 다시."""
    for k in range(3):
        try:
            r = requests.get("%s/%s" % (API, ep), params={"$filter": flt}, headers=UA, timeout=60)
            if r.status_code == 200:
                return r.json().get("value", [])
            LOG.append({"ep": ep, "filter": flt, "status": r.status_code, "head": r.text[:300]})
            if r.status_code in (400, 404):
                return None
        except (requests.RequestException, ValueError) as e:
            LOG.append({"ep": ep, "filter": flt, "error": repr(e)[:200]})
        time.sleep(3 * (k + 1))
    return None


def pull(ep, d0, d1):
    """business_date 가 [d0, d1) 인 줄. 한 번에 100줄 제한이라 100줄이 차면 기간을 반으로 나눔."""
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
        LOG.append({"ep": ep, "day": d0.isoformat(), "note": "하루 100줄 이상, 일부 빠졌을 수 있음"})
    return rows


def time_col(rows):
    for r in rows[:1]:
        for k in r:
            if "dtime" in k.lower() and "utc" not in k.lower() and "publication" not in k.lower():
                return k
    return None


def numeric_cols(rows):
    cols = []
    for k in (rows[0] if rows else {}):
        kl = k.lower()
        if kl.startswith(SKIP_PREFIX) or SKIP_WORDS & set(kl.split("_")):
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


def main():
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
           "key_cols": {"grid_demand_fcst": "수요 전망 MW", "fcst_wi_tot_gen": "풍력 전망 MW",
                        "fcst_pv_tot_gen": "태양광 전망 MW"},
           "columns": list(rows[0].keys()) if rows else [], "time_col": tc, "numeric_cols": cols,
           "rows_total": len(rows), "daily": daily, "rows": rows, "log": LOG}
    with open(OUT_FC, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1, default=str)
    log("[pk5l-wp] %d줄 · %d일 · 시각 열 %s · 수요 전망 포함 %s" % (len(rows), len(daily), tc, "grid_demand_fcst" in cols))


if __name__ == "__main__":
    main()
