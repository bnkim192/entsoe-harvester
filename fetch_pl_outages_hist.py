"""
폴란드 발전설비 정지 과거 월별 비교 (2026-11-23 신규, 법인 '11월 계획정지가 작년보다 많았다' 정량 확인용)

자료  ENTSO-E A80 발전기 정지 공시(REMIT) — fetch_pl_outages.py 의 수집·해석 함수를 그대로 쓴다.
기간  2024-10 ~ 2024-12, 2025-10 ~ 2025-12, 2026-10 ~ 이번 달
계산  달마다: 그 달과 겹치는 공시를 받아 같은 공시는 최신판만, 취소(A09) 제외,
      날짜별로 같은 발전기는 그날 가장 큰 정지 1건만 더해 '하루 정지 MW' → 월평균 (계획 A53 · 고장 A54 · 전원별)
      공시 판(revision)이 여러 개 받아지면 첫 판과 최신판의 정지 기간·MW 차이도 기록(계획 대비 실제 변경 정도)
산출  pl_outages_hist.json   (공개 자료만)
환경변수  ENTSOE_TOKEN
"""
import json, time
import datetime as dt
import xml.etree.ElementTree as ET
import fetch_pl_outages as fo

MONTHS = ["2024-10", "2024-11", "2024-12", "2025-10", "2025-11", "2025-12"]
OUT = "pl_outages_hist.json"


def months_now():
    t = dt.datetime.now(fo.TZ).date()
    ms = list(MONTHS)
    y, m = 2026, 10
    while (y, m) <= (t.year, t.month):
        ms.append("%04d-%02d" % (y, m))
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return ms


def month_bounds(ym):
    y, m = int(ym[:4]), int(ym[5:])
    d0 = dt.date(y, m, 1)
    d1 = dt.date(y + (m == 12), (m % 12) + 1, 1)
    return d0, d1


def one_month(ym):
    d0, d1 = month_bounds(ym)
    s = dt.datetime.combine(d0, dt.time(0), fo.TZ).astimezone(fo.UTC)
    e = dt.datetime.combine(d1, dt.time(0), fo.TZ).astimezone(fo.UTC)
    recs = []
    for doc_type in ("A80",):
        for x in fo.get_docs(doc_type, s, e):
            try:
                recs += fo.parse(x, doc_type)
            except ET.ParseError:
                pass
    # 판별 묶음: 같은 공시(mRID)·발전기·시작 → 첫 판/최신판
    groups = {}
    for r in recs:
        k = (r["mrid"], r["unit"] or r["plant"])
        groups.setdefault(k, []).append(r)
    revs = []
    best = {}
    for k, rs in groups.items():
        rs.sort(key=lambda r: r["rev"])
        last = rs[-1]
        if len(set(r["rev"] for r in rs)) > 1:
            first = rs[0]
            revs.append({"unit": k[1], "type": last["type"], "rev_first": first["rev"], "rev_last": last["rev"],
                         "first_start": first["start_utc"], "first_end": first["end_utc"], "first_mw": first["unavail_mw"],
                         "last_start": last["start_utc"], "last_end": last["end_utc"], "last_mw": last["unavail_mw"],
                         "last_status": last["status"]})
        for r in rs:
            kk = (r["mrid"], r["unit"] or r["plant"], r["start_utc"])
            if kk not in best or r["rev"] > best[kk]["rev"]:
                best[kk] = r
    cur = [r for r in best.values() if r["status"] != "A09" and r["unavail_mw"]]
    daily = {}
    d = d0
    while d < d1:
        a = dt.datetime.combine(d, dt.time(0), fo.TZ).astimezone(fo.UTC)
        b = a + dt.timedelta(days=1)
        per = {}
        for r in cur:
            rs_ = dt.datetime.fromisoformat(r["start_utc"]) if r["start_utc"] else None
            re_ = dt.datetime.fromisoformat(r["end_utc"]) if r["end_utc"] else None
            if (rs_ and rs_ >= b) or (re_ and re_ <= a):
                continue
            ku = r["unit"] or r["plant"] or r["mrid"]
            if ku not in per or r["unavail_mw"] > per[ku]["unavail_mw"]:
                per[ku] = r
        tot = {"all": 0.0, "계획": 0.0, "고장": 0.0}
        fuel = {}
        for r in per.values():
            v = r["unavail_mw"]
            tot["all"] += v
            tot[r["type"] if r["type"] in ("계획", "고장") else "all"] += 0 if r["type"] not in ("계획", "고장") else v
            if r["type"] == "계획":
                fuel[r["fuel"] or "미상"] = fuel.get(r["fuel"] or "미상", 0) + v
        daily[d.isoformat()] = {"n": len(per), "all": round(tot["all"], 1), "planned": round(tot["계획"], 1),
                                "forced": round(tot["고장"], 1), "planned_by_fuel": {k: round(v, 1) for k, v in fuel.items()}}
        d += dt.timedelta(days=1)
    t_today = dt.datetime.now(fo.TZ).date().isoformat()
    past = [v for k, v in daily.items() if k < t_today]
    def av(key):
        return round(sum(v[key] for v in past) / len(past), 1) if past else None
    fuels = {}
    for v in past:
        for k, x in v["planned_by_fuel"].items():
            fuels[k] = fuels.get(k, 0) + x
    fuels = {k: round(x / len(past), 1) for k, x in sorted(fuels.items(), key=lambda kv: -kv[1])} if past else {}
    fo.log("%s 문서 %d · 정지 %d건 · 지난 날 %d일 · 평균 정지 %s MW (계획 %s · 고장 %s) · 판 변경 %d건"
           % (ym, len(groups), len(cur), len(past), av("all"), av("planned"), av("forced"), len(revs)))
    return {"month": ym, "days": len(past), "avg_all_mw": av("all"), "avg_planned_mw": av("planned"),
            "avg_forced_mw": av("forced"), "planned_by_fuel_mw": fuels, "outages": len(cur), "revisions": revs[:300],
            "n_revised": len(revs), "daily": daily}


def main():
    res = {"fetched_utc": dt.datetime.now(fo.UTC).isoformat(timespec="seconds"),
           "note": "ENTSO-E A80 발전기 정지 공시. 하루 정지 MW = 그날 겹치는 정지(발전기당 최대 1건)의 합 → 월평균. 계획=A53, 고장=A54.",
           "months": []}
    for ym in months_now():
        try:
            res["months"].append(one_month(ym))
        except Exception as e:
            fo.log("%s 실패 %r" % (ym, e))
            res["months"].append({"month": ym, "error": repr(e)})
        time.sleep(1)
    res["diag"] = fo.DIAG[-80:]
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
