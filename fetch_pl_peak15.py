"""
폴란드 15분 수요 PEAK (2026-11-26, WA법인 요청 3번)

ENTSO-E A65(실제 수요, processType A16, 폴란드)를 15분 단위로 받아 '하루 요약'만 저장한다.
기존 fetch_pl_drivers.py 는 하루 평균만 남기므로, PEAK 분석용으로 따로 받는다(기존 코드 변경 없음).

하루 요약 (현지 시각 Europe/Warsaw 기준)
  mean      하루 평균 MW
  max       하루 최대 15분 MW,  max_at  그 시각(HH:MM)
  min       하루 최소 MW
  eve_mean  저녁 16:00~19:45 평균 MW
  peak_gap  max − mean (MW)
  n, res    점 개수, 분해능(분)

사용  python fetch_pl_peak15.py              → 최근 40일만 다시 받아 기존 파일에 덮어 붙임(매일용)
      python fetch_pl_peak15.py 2023-01-01   → 그날부터 전부 (백필)
산출  pl_peak15.json   (공개 자료만, 토큰은 GitHub Secrets ENTSOE_TOKEN)
"""
import json, os, sys, time
import datetime as dt
import xml.etree.ElementTree as ET
from zoneinfo import ZoneInfo
import requests

TOKEN = os.environ.get("ENTSOE_TOKEN", "").strip()
BASE_URL = "https://web-api.tp.entsoe.eu/api"
PL = "10YPL-AREA-----S"
TZ = ZoneInfo("Europe/Warsaw")
UTC = dt.timezone.utc
OUT = "pl_peak15.json"
MIN_HOURS = 20
LOG = []


def log(msg):
    print(msg, flush=True)


def entsoe(params):
    p = dict(params, securityToken=TOKEN)
    for k in range(3):
        try:
            r = requests.get(BASE_URL, params=p, timeout=180)
        except requests.RequestException as e:
            LOG.append({"start": params.get("periodStart"), "error": type(e).__name__})
            time.sleep(5 * (k + 1))
            continue
        txt = r.text or ""
        if r.status_code == 200 and txt.lstrip().startswith("<"):
            if "Acknowledgement_MarketDocument" in txt[:400]:
                LOG.append({"start": params.get("periodStart"), "note": "자료 없음"})
                return None
            return txt
        LOG.append({"start": params.get("periodStart"), "status": r.status_code, "head": txt[:160]})
        if r.status_code in (429, 503):
            time.sleep(20)
            continue
        return None
    return None


def parse(xml):
    """[(utc 시각, MW, 분해능 분)] · curveType A03(같은 값 생략)은 앞 값으로 채움."""
    root = ET.fromstring(xml)
    ns = root.tag.split("}")[0].strip("{") if "}" in root.tag else ""
    q = (lambda t: "{%s}%s" % (ns, t)) if ns else (lambda t: t)
    out = []
    for ts in root.iter(q("TimeSeries")):
        ce = ts.find(q("curveType"))
        curve = (ce.text or "").strip() if ce is not None else "A01"
        for per in ts.iter(q("Period")):
            ti = per.find(q("timeInterval"))
            if ti is None:
                continue
            st, en = ti.find(q("start")), ti.find(q("end"))
            rs = per.find(q("resolution"))
            res = rs.text if rs is not None and rs.text else "PT60M"
            step = 15 if "PT15M" in res else (30 if "PT30M" in res else 60)
            t0 = dt.datetime.strptime(st.text.strip()[:16], "%Y-%m-%dT%H:%M").replace(tzinfo=UTC)
            t1 = dt.datetime.strptime(en.text.strip()[:16], "%Y-%m-%dT%H:%M").replace(tzinfo=UTC)
            nslot = int((t1 - t0).total_seconds() // (step * 60))
            vals = {}
            for pt in per.iter(q("Point")):
                pe, qe = pt.find(q("position")), pt.find(q("quantity"))
                if pe is not None and qe is not None:
                    vals[int(pe.text)] = float(qe.text)
            last = None
            for pos in range(1, nslot + 1):
                if pos in vals:
                    last = vals[pos]
                elif curve != "A03":
                    continue
                if last is not None:
                    out.append((t0 + dt.timedelta(minutes=step * (pos - 1)), last, step))
    return out


def windows(start):
    s = dt.datetime.combine(start, dt.time()) - dt.timedelta(days=1)
    end = dt.datetime.now(UTC).replace(tzinfo=None, minute=0, second=0, microsecond=0)
    out = []
    while s < end:
        e = min(s + dt.timedelta(days=31), end)
        out.append((s, e))
        s = e
    return out


def summarize(points):
    """{현지 날짜: 요약}. 가장 촘촘한 분해능만 쓰고, 20시간 미만인 날은 버림."""
    g = {}
    for t, v, step in points:
        lt = t.astimezone(TZ)
        g.setdefault(lt.strftime("%Y-%m-%d"), {}).setdefault(step, {})[lt.strftime("%H:%M")] = v
    out = {}
    for d, bys in g.items():
        ok = [s for s, vs in bys.items() if len(vs) * s / 60 >= MIN_HOURS]
        if not ok:
            continue
        s = min(ok)
        vs = bys[s]
        items = sorted(vs.items())
        vals = [v for _, v in items]
        mx = max(items, key=lambda x: x[1])
        eve = [v for h, v in items if "16:00" <= h <= "19:59"]
        mean = sum(vals) / len(vals)
        out[d] = {"mean": round(mean, 1), "max": round(mx[1], 1), "max_at": mx[0], "min": round(min(vals), 1),
                  "eve_mean": round(sum(eve) / len(eve), 1) if eve else None, "peak_gap": round(mx[1] - mean, 1),
                  "n": len(vals), "res": s}
    return out


def main():
    if not TOKEN:
        log("ENTSOE_TOKEN 없음 (GitHub Secrets 확인)")
        sys.exit(1)
    today = dt.date.today()
    start = dt.date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].strip() else today - dt.timedelta(days=40)
    old = {}
    if os.path.exists(OUT):
        try:
            with open(OUT, encoding="utf-8") as f:
                old = json.load(f).get("daily", {})
        except ValueError:
            old = {}
    pts = []
    ws = windows(start)
    for i, (s, e) in enumerate(ws):
        x = entsoe({"documentType": "A65", "processType": "A16", "outBiddingZone_Domain": PL,
                    "periodStart": s.strftime("%Y%m%d%H00"), "periodEnd": e.strftime("%Y%m%d%H00")})
        if x:
            try:
                got = parse(x)
                pts += got
            except ET.ParseError as ex:
                LOG.append({"start": s.isoformat(), "error": "XML %s" % ex})
                got = []
            log("  %s ~ %s  %d점" % (s.date(), e.date(), len(got)))
        time.sleep(1)
    new = summarize(pts)
    today_s = today.isoformat()
    daily = dict(old)
    for d, v in new.items():
        if d >= start.isoformat() and d < today_s:
            daily[d] = v
    days = sorted(daily)
    res = sorted({v["res"] for v in daily.values()})
    out = {"fetched_utc": dt.datetime.now(UTC).isoformat(timespec="seconds"),
           "source": "ENTSO-E Transparency A65 Actual Total Load (A16), Poland",
           "unit": "MW", "tz": "Europe/Warsaw",
           "note": "mean·max·min·eve_mean(16~20시)·peak_gap(max−mean) · 분해능 res 분 · 20시간 미만인 날 제외",
           "first_day": days[0] if days else None, "last_day": days[-1] if days else None, "n_days": len(days),
           "resolutions": res, "daily": {d: daily[d] for d in days}, "log": LOG[-100:]}
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    log("[pl_peak15] %s ~ %s · %d일 · 분해능 %s분 · 이번에 받은 점 %d" % (out["first_day"], out["last_day"], len(days), res, len(pts)))


if __name__ == "__main__":
    main()
