"""
폴란드 발전설비 정비 '1월 계획 vs 연말 실적' 비교용 자료 한 번 받기 (2026-11-23 신규)

1) PSE 5개년 조정계획 정비(unav-pk5l) — 정비 시작일(business_date) 2021-10 ~ 2026-12 의 모든 줄(판·공시 시각 포함)
   → 'Y년 1월 31일까지 공시된 판' = 1월 계획, '마지막 판' = 실적에 가까운 최종 일정 으로 PC 에서 비교
   PSE 제약: 한 번에 100줄, 쪽 넘김 없음 → business_date 구간을 반씩 쪼개 받음 (fetch_pse_pk5l.pull_range 재사용)
2) ENTSO-E A80 발전기 정지(최종판) 월별 하루 평균 정지 MW (계획·고장, 전원별) 2022-01 ~ 이번 달 (2026-11-23: 2026년 추가)
   (fetch_pl_outages_hist.one_month 재사용)
산출  pse_pk5l_hist.json, pl_outages_hist_2022_2026.json   (공개 자료만)
"""
import json, time
import datetime as dt
import fetch_pse_pk5l as pk
import fetch_pl_outages_hist as oh


def main():
    # ── PSE
    rows, m = [], dt.date(2021, 10, 1)
    while m < dt.date(2027, 1, 1):
        n = dt.date(m.year + (m.month == 12), (m.month % 12) + 1, 1)
        rows += pk.pull_range(m, n)
        m = n
        time.sleep(0.3)
    seen, uniq = set(), []
    for r in rows:
        k = (r.get("mrid_zas"), r.get("unit_code"), r.get("start_dtime"), r.get("end_dtime"), r.get("publication_ts"), r.get("state"))
        if k not in seen:
            seen.add(k)
            uniq.append(r)
    per_mrid = {}
    for r in uniq:
        per_mrid.setdefault(r.get("mrid_zas"), set()).add(r.get("publication_ts"))
    multi = sum(1 for v in per_mrid.values() if len(v) > 1)
    with open("pse_pk5l_hist.json", "w", encoding="utf-8") as f:
        json.dump({"fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
                   "note": "PSE unav-pk5l, business_date 2021-10~2026-12. 같은 mrid_zas 의 여러 publication_ts = 판 이력",
                   "rows": uniq, "n_mrid": len(per_mrid), "n_mrid_multi_versions": multi, "log": pk.LOG[-200:]},
                  f, ensure_ascii=False, indent=0, default=str)
    print("[pse] 줄 %d · 정비(mrid) %d · 판이 여러 개인 정비 %d" % (len(uniq), len(per_mrid), multi), flush=True)
    # ── ENTSO-E
    res = {"fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
           "note": "ENTSO-E A80 최신판(최종) 기준 하루 정지 MW 월평균. 계획=A53, 고장=A54.", "months": []}
    today = dt.date.today()
    for y in range(2022, 2027):
        for mo in range(1, 13):
            if (y, mo) > (today.year, today.month):
                break
            ym = "%04d-%02d" % (y, mo)
            try:
                x = oh.one_month(ym)
                x.pop("revisions", None)
                res["months"].append(x)
            except Exception as e:
                res["months"].append({"month": ym, "error": repr(e)})
            time.sleep(0.5)
    with open("pl_outages_hist_2022_2026.json", "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=0)
    print("[entsoe] 월 %d개" % len(res["months"]), flush=True)


if __name__ == "__main__":
    main()
