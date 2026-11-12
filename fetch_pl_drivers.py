# -*- coding: utf-8 -*-
"""
폴란드 전력가격 원인 판정용 일별 데이터 수집기 (2026-11-09)
— 주간 보고(pl_weekly_report.py)와 10/27 급등 원인 판정의 원자료.

■ 왜 만드나
  TGE 파일 분석 결과 상승이 두 겹이었다.
    · 6~10월 +252 PLN/MWh — 가스 83% · 탄소배출권 8% · 석탄 3% · 기타 6% (원가식 분해)
    · 10/27 급등 (10/26 759 → 10/27 1,384 PLN/MWh) — 연료로는 15% 만 설명, 85% 미설명
  미설명 몫이 날씨(바람·햇빛·기온)인지 공급 사건(정비·고장·연계선)인지 판정할 자료를 받는다.

■ 수집 항목 (전부 무료)
  ENTSO-E (기존 ENTSOE_TOKEN 그대로)
    A44  하루전 가격 — PL + 주변국 DE·CZ·SK·LT·SE4
    A65  폴란드 실제 수요
    A75  폴란드 전원별 실제 발전 — 풍력(육상·해상)·태양광·가스·유연탄·갈탄
  Open-Meteo (키 없음)
    기온·100m 풍속·일사 — 바르샤바·포모제(북부 풍력)·실롱스크(남부) 3지점 평균 + 향후 7일 예보

■ 판정 (로그와 pl_drivers_diag.json)
  ① 급등 전(9/1~급등 전날) 일별로 '가격 ~ 잔여부하 + 시간추세' 직선을 잡는다
     잔여부하 = 수요 − 풍력 − 태양광 (화력이 채워야 하는 몫, 가격을 가장 잘 설명)
     시간추세 = 연료가격처럼 서서히 움직이는 요인을 따로 떼어 내기 위한 항
  ② 급등기에 대입 → 잔여부하 몫 = 날씨·수요 / 추세 몫 = 연료 등 / 남는 몫 = 공급 사건·기타
  ③ 독일도 비슷하게 뛰었으면 지역 공통(날씨·연료), 폴란드만 뛰었으면 폴란드 고유 요인
  ※ 발전소 정지 정보(A80)는 zip 형식이라 이번 판에 넣지 않았다 → '남는 몫'에 포함된다.

■ 원칙
  1. 값을 만들어내지 않는다 — 못 받은 항목, 20시간 미만인 날은 빈칸.
  2. 날짜는 폴란드 현지 시각(Europe/Warsaw)으로 자른다.
  3. 같은 날 15분·60분 자료가 함께 오면 촘촘한 쪽 하나만 쓴다(중복 계산 방지).
  4. 값이 바뀔 때만 점을 찍는 문서(curveType A03)는 다음 점까지 같은 값으로 채운다.
  5. 수요·발전·기상 실측은 어제까지만 쓴다(오늘은 하루가 덜 찼다). 가격은 확정된 내일분까지 받는다.
  6. 토큰은 로그에 찍지 않는다(예외 메시지도 종류만 찍는다 — URL 에 토큰이 들어 있다).

■ 산출
  pl_drivers_daily.csv / .json   일별 표 (+ 향후 7일 기상 예보)
  pl_drivers_diag.json           급등 판정표

■ 환경변수 (선택)
  PLD_START  수집 시작일 (기본 2026-01-01)
  PLD_SPIKE  급등 시작일 (기본 2026-10-27)
"""
import os, json, csv, time
import datetime as dt
import xml.etree.ElementTree as ET
from zoneinfo import ZoneInfo
import requests

TOKEN = os.environ.get("ENTSOE_TOKEN", "").strip()
BASE_URL = "https://web-api.tp.entsoe.eu/api"
PL = "10YPL-AREA-----S"
NEIGH = {"DE": "10Y1001A1001A82H", "CZ": "10YCZ-CEPS-----N", "SK": "10YSK-SEPS-----K",
         "LT": "10YLT-1001A0008Q", "SE4": "10Y1001A1001A47J", "UA": "10Y1001C--000182"}   # UA = 우크라이나 IPS(2026-11-13 추가)
PSR = {"B19": "wind_on", "B18": "wind_off", "B16": "solar",
       "B04": "gas", "B05": "hardcoal", "B02": "lignite",
       "B01": "biomass", "B10": "hydro_ps", "B11": "hydro_ror", "B12": "hydro_res"}   # 바이오매스·수력(양수·유입식·저수지) 2026-11-13 추가
UA_ALT = {"UA": ["10YUA-WEPS-----0", "10Y1001C--00003F"]}   # UA IPS 다음 후보: Burshtyn 섬, UA 전체
START = os.environ.get("PLD_START", "").strip() or "2026-01-01"
SPIKE_FROM = os.environ.get("PLD_SPIKE", "").strip() or "2026-10-27"
BASE_FROM = "2026-09-01"
BASE_TO = (dt.date.fromisoformat(SPIKE_FROM) - dt.timedelta(days=1)).isoformat()
TZ = ZoneInfo("Europe/Warsaw")
UTC = dt.timezone.utc
MIN_HOURS = 20
WX_POINTS = [("Warszawa", 52.23, 21.01), ("Pomorze", 54.35, 17.00), ("Slask", 50.26, 19.02)]

OUT_CSV = "pl_drivers_daily.csv"
OUT_JSON = "pl_drivers_daily.json"
OUT_DIAG = "pl_drivers_diag.json"
LOG = []
SKIPPED_CUR = set()
CUR_WANT = ["EUR"]          # 가격 통화. 우크라이나만 잠깐 UAH 로 바꿔 받은 뒤 NBU 환율로 유로 환산
FORECAST = {}


def log(*a):
    print(*a, flush=True)


def today_local():
    return dt.datetime.now(TZ).date().isoformat()


# ══════════════════════════════════════════════════════════════════
#  공통 계산
# ══════════════════════════════════════════════════════════════════
def mean(xs):
    xs = [x for x in xs if x is not None]
    return sum(xs) / len(xs) if xs else None


def _solve(A, b):
    n = len(b)
    M = [list(A[i]) + [b[i]] for i in range(n)]
    for c in range(n):
        p = max(range(c, n), key=lambda r: abs(M[r][c]))
        if abs(M[p][c]) < 1e-12:
            return None
        M[c], M[p] = M[p], M[c]
        for r in range(n):
            if r != c:
                f = M[r][c] / M[c][c]
                for k in range(c, n + 1):
                    M[r][k] -= f * M[c][k]
    return [M[i][n] / M[i][i] for i in range(n)]


def ols(cols, y):
    """최소제곱 직선. cols = 설명변수 목록(각각 y 와 같은 길이), 절편 자동.
    반환 {"coef": [절편, b1, b2, ...], "r2": 설명력(0~1), "n": 표본수} — 표본이 모자라면 None."""
    n, k = len(y), len(cols) + 1
    if n < k + 8:
        return None
    X = [[1.0] + [c[i] for c in cols] for i in range(n)]
    A = [[sum(X[r][i] * X[r][j] for r in range(n)) for j in range(k)] for i in range(k)]
    bb = [sum(X[r][i] * y[r] for r in range(n)) for i in range(k)]
    coef = _solve(A, bb)
    if coef is None:
        return None
    my = sum(y) / n
    syy = sum((v - my) ** 2 for v in y)
    sse = sum((y[r] - sum(coef[i] * X[r][i] for i in range(k))) ** 2 for r in range(n))
    return {"coef": coef, "r2": (1 - sse / syy) if syy else None, "n": n}


# ══════════════════════════════════════════════════════════════════
#  ENTSO-E
# ══════════════════════════════════════════════════════════════════
def windows():
    """START ~ 모레(UTC)를 31일 단위로 자른다(문서 크기·타임아웃 방지)."""
    s = dt.datetime.strptime(START, "%Y-%m-%d")
    end = dt.datetime.now(UTC).replace(tzinfo=None, minute=0, second=0, microsecond=0) + dt.timedelta(days=2)
    out = []
    while s < end:
        e = min(s + dt.timedelta(days=31), end)
        out.append((s, e))
        s = e
    return out


def entsoe(params, tag):
    """XML 텍스트 반환, 실패·자료없음이면 None."""
    p = dict(params)
    p["securityToken"] = TOKEN
    for _ in range(3):
        try:
            r = requests.get(BASE_URL, params=p, timeout=180)
        except Exception as e:
            log("   [%s] 접속 실패 (%s)" % (tag, type(e).__name__))
            time.sleep(5)
            continue
        txt = r.text
        if r.status_code == 200 and txt.lstrip().startswith("<"):
            if "Acknowledgement_MarketDocument" in txt[:400]:
                i, j = txt.find("<text>"), txt.find("</text>")
                log("   [%s] 자료 없음 %s" % (tag, txt[i + 6:j][:120] if 0 <= i < j else ""))
                return None
            return txt
        log("   [%s] HTTP %s %s" % (tag, r.status_code, txt[:160].replace("\n", " ")))
        if r.status_code in (429, 503):
            time.sleep(15)
            continue
        return None
    return None


def parse_t(s):
    return dt.datetime.strptime(s.strip()[:16], "%Y-%m-%dT%H:%M").replace(tzinfo=UTC)


def parse_series(xml, value_tag, need_in_domain=False):
    """[(psrType 또는 None, utc 시각, 값, 분해능(분))] — 태그 이름으로만 읽는다."""
    root = ET.fromstring(xml)
    ns = root.tag.split("}")[0].strip("{") if "}" in root.tag else ""
    q = (lambda t: "{%s}%s" % (ns, t)) if ns else (lambda t: t)
    out = []
    for ts in root.iter(q("TimeSeries")):
        if need_in_domain and ts.find(q("inBiddingZone_Domain.mRID")) is None:
            continue                                   # 발전 방향만 (양수 소비 방향 제외)
        cur = ts.find(q("currency_Unit.name"))
        if value_tag == "price.amount" and cur is not None and (cur.text or "").strip() != CUR_WANT[0]:
            SKIPPED_CUR.add((cur.text or "").strip())     # 유로가 아닌 가격은 버리고 통화만 기록
            continue                                   # 유로 가격만
        psr = None
        mk = ts.find(q("MktPSRType"))
        if mk is not None:
            pe = mk.find(q("psrType"))
            psr = pe.text.strip() if pe is not None and pe.text else None
        ce = ts.find(q("curveType"))
        curve = (ce.text or "").strip() if ce is not None else "A01"
        for per in ts.iter(q("Period")):
            ti = per.find(q("timeInterval"))
            if ti is None:
                continue
            st, en = ti.find(q("start")), ti.find(q("end"))
            if st is None or not st.text:
                continue
            rs = per.find(q("resolution"))
            res = rs.text if rs is not None and rs.text else "PT60M"
            step = 15 if "PT15M" in res else (30 if "PT30M" in res else 60)
            t0 = parse_t(st.text)
            nslot = None
            if en is not None and en.text:
                nslot = int((parse_t(en.text) - t0).total_seconds() // (step * 60))
            pts = []
            for pt in per.iter(q("Point")):
                pos, val = pt.find(q("position")), pt.find(q(value_tag))
                if pos is None or val is None:
                    continue
                try:
                    pts.append((int(pos.text), float(val.text)))
                except (TypeError, ValueError):
                    continue
            pts.sort()
            if curve == "A03" and nslot:
                for i, (p, v) in enumerate(pts):
                    nxt = pts[i + 1][0] if i + 1 < len(pts) else nslot + 1
                    for k in range(p, nxt):
                        out.append((psr, t0 + dt.timedelta(minutes=step * (k - 1)), v, step))
            else:
                for p, v in pts:
                    out.append((psr, t0 + dt.timedelta(minutes=step * (p - 1)), v, step))
    return out


def fetch_points(base, tag, value_tag, need_in=False):
    """{psr: {(분해능, utc시각): 값}}"""
    pts = {}
    for s, e in windows():
        p = dict(base)
        p["periodStart"] = s.strftime("%Y%m%d%H%M")
        p["periodEnd"] = e.strftime("%Y%m%d%H00")
        x = entsoe(p, tag)
        if not x:
            continue
        try:
            for psr, t, v, step in parse_series(x, value_tag, need_in):
                pts.setdefault(psr, {})[(step, t)] = v
        except ET.ParseError as ex:
            log("   [%s] XML 해석 실패 %s" % (tag, ex))
    return pts


def by_day(points, until=None):
    """{(분해능, utc시각): 값} → {현지날짜: (분해능, [값...])}. 20시간 미만인 날은 버린다."""
    g = {}
    for (step, t), v in points.items():
        d = t.astimezone(TZ).strftime("%Y-%m-%d")
        if until and d >= until:
            continue
        g.setdefault(d, {}).setdefault(step, []).append(v)
    out = {}
    for d, bys in g.items():
        ok = [s for s, vs in bys.items() if len(vs) * s / 60 >= MIN_HOURS]
        if ok:
            s = min(ok)
            out[d] = (s, bys[s])
    return out


def dmean(bd):
    return {d: sum(vs) / len(vs) for d, (s, vs) in bd.items()}


def dsum_mwh(bd):
    """{날짜: (분해능 분, [MW...])} → {날짜: MWh}. 20시간 미만인 날은 by_day 에서 이미 빠짐."""
    return {d: sum(vs) * s / 60 for d, (s, vs) in bd.items()}


def nbu_eur_rates(d0, d1):
    """우크라이나 중앙은행 공식 환율 {날짜: UAH per EUR}. 키 불필요."""
    out = {}
    s, e = dt.date.fromisoformat(d0) - dt.timedelta(days=7), dt.date.fromisoformat(d1)
    while s <= e:                                   # 한 번에 최대 약 1년이라 180일씩 나눠 받음
        t = min(s + dt.timedelta(days=180), e)
        try:
            r = requests.get("https://bank.gov.ua/NBU_Exchange/exchange_site",
                             params={"start": s.strftime("%Y%m%d"), "end": t.strftime("%Y%m%d"), "valcode": "eur",
                                     "sort": "exchangedate", "order": "asc", "json": ""}, timeout=60)
            for x in r.json():
                ds = x.get("exchangedate") or ""            # dd.mm.yyyy
                rate = x.get("rate_per_unit") or (x.get("rate") / (x.get("units") or 1) if x.get("rate") else None)
                if ds and rate:
                    out[dt.datetime.strptime(ds, "%d.%m.%Y").date().isoformat()] = float(rate)
        except Exception as ex:
            log("   [NBU 환율] 실패 %s" % type(ex).__name__)
        s = t + dt.timedelta(days=1)
    return out


def merged(pts):
    allp = {}
    for d in pts.values():
        allp.update(d)
    return allp


# ══════════════════════════════════════════════════════════════════
#  수집
# ══════════════════════════════════════════════════════════════════
def get_prices():
    out = {}
    for name, eic in [("PL", PL)] + list(NEIGH.items()):
        pts = merged(fetch_points({"documentType": "A44", "in_Domain": eic, "out_Domain": eic},
                                  "A44 " + name, "price.amount"))
        if not pts and name in UA_ALT:                 # 우크라이나: ENTSO-E 는 UAH 로만 줌 → UAH 로 받아 NBU 환율로 유로 환산
            CUR_WANT[0] = "UAH"
            try:
                for code in [eic] + UA_ALT[name]:
                    pts = merged(fetch_points({"documentType": "A44", "in_Domain": code, "out_Domain": code},
                                              "A44 %s(%s, UAH)" % (name, code), "price.amount"))
                    if pts:
                        LOG.append({"item": "가격 %s 코드·통화" % name, "eic": code, "currency": "UAH"})
                        break
            finally:
                CUR_WANT[0] = "EUR"
            if pts:
                bd = by_day(pts)
                uah = dmean(bd)
                fx = nbu_eur_rates(min(uah), max(uah))
                ks = sorted(fx)
                conv = {}
                for d, v in uah.items():
                    prior = [k for k in ks if k <= d]
                    if prior:
                        conv[d] = v / fx[prior[-1]]       # 주말·휴일은 직전 공시 환율
                out["px_UA"] = conv
                out["px_UA_uah"] = uah
                LOG.append({"item": "가격 UA", "points": len(pts), "days": len(conv), "fx_days": len(fx)})
                log("[가격 UA ] UAH 포인트 %6d · 유로 환산 일수 %3d (NBU 환율 %d일)" % (len(pts), len(conv), len(fx)))
                continue
        bd = by_day(pts)
        out["px_" + name] = dmean(bd)
        if name == "PL":
            out["px_PL_max"] = {d: max(vs) for d, (s, vs) in bd.items()}
            out["px_PL_min"] = {d: min(vs) for d, (s, vs) in bd.items()}
            out["px_PL_neg_h"] = {d: sum(1 for v in vs if v < 0) * s / 60 for d, (s, vs) in bd.items()}
        LOG.append({"item": "가격 " + name, "points": len(pts), "days": len(bd)})
        log("[가격 %-3s] 포인트 %6d · 일수 %3d" % (name, len(pts), len(bd)))
    return out


def get_load():
    pts = merged(fetch_points({"documentType": "A65", "processType": "A16", "outBiddingZone_Domain": PL},
                              "A65 수요", "quantity"))
    bd = by_day(pts, until=today_local())
    d = dmean(bd)
    LOG.append({"item": "수요", "points": len(pts), "days": len(d)})
    log("[수요    ] 포인트 %6d · 일수 %3d" % (len(pts), len(d)))
    return {"load": d, "load_mwh": dsum_mwh(bd)}


def get_gen():
    pts = fetch_points({"documentType": "A75", "processType": "A16", "in_Domain": PL},
                       "A75 발전", "quantity", need_in=True)
    out, t = {}, today_local()
    for code, key in PSR.items():
        bd = by_day(pts.get(code, {}), until=t)
        out[key] = dmean(bd)
        out[key + "_mwh"] = dsum_mwh(bd)                       # 하루 발전량 MWh (MW × 분해능 합)
        if key == "solar":
            out["solar_peak"] = {d: max(vs) for d, (s, vs) in bd.items()}   # 태양광 하루 최고 출력 MW
        LOG.append({"item": "발전 " + key, "points": len(pts.get(code, {})), "days": len(out[key])})
        log("[발전 %-9s] 포인트 %6d · 일수 %3d" % (key, len(pts.get(code, {})), len(out[key])))
    others = sorted(str(k) for k in pts if k not in PSR)
    log("   (받았지만 쓰지 않은 전원코드: %s)" % (", ".join(others) or "없음"))
    return out


def get_weather():
    """Open-Meteo — 과거(archive) 우선, 최근 며칠은 forecast 의 past_days 로 메운다. 오늘부터는 예보."""
    today = dt.datetime.now(TZ).date()
    tstr = today.isoformat()
    keys = (("temperature_2m", "temp"), ("wind_speed_100m", "wind100"), ("shortwave_radiation", "rad"))
    hist = {k: {} for _, k in keys}
    fc = {k: {} for _, k in keys}
    for nm, lat, lon in WX_POINTS:
        ser = {}
        reqs = [("https://archive-api.open-meteo.com/v1/archive",
                 {"start_date": START, "end_date": (today - dt.timedelta(days=1)).isoformat()}),
                ("https://api.open-meteo.com/v1/forecast", {"past_days": 14, "forecast_days": 8})]
        for url, extra in reqs:
            p = {"latitude": lat, "longitude": lon, "timezone": "Europe/Warsaw",
                 "hourly": ",".join(s for s, _ in keys)}
            p.update(extra)
            try:
                r = requests.get(url, params=p, timeout=90)
            except Exception as e:
                log("   [기상 %s] 접속 실패 (%s)" % (nm, type(e).__name__))
                continue
            if r.status_code != 200:
                log("   [기상 %s] HTTP %s %s" % (nm, r.status_code, r.text[:120]))
                continue
            h = r.json().get("hourly", {})
            tl = h.get("time", [])
            for src, key in keys:
                vals = h.get(src) or []
                for i, ts in enumerate(tl):
                    v = vals[i] if i < len(vals) else None
                    if v is not None:
                        ser.setdefault(key, {}).setdefault(ts, v)      # 먼저 받은 아카이브 우선
        for _, key in keys:
            byd = {}
            for ts, v in ser.get(key, {}).items():
                byd.setdefault(ts[:10], []).append(v)
            for d, vs in byd.items():
                if len(vs) < MIN_HOURS:
                    continue
                val = sum(vs) / 1000 if key == "rad" else sum(vs) / len(vs)   # 일사: Wh/m2 합 → kWh/m2
                (fc if d >= tstr else hist)[key].setdefault(d, []).append(val)
        log("[기상 %-8s] 시각 %5d" % (nm, len(ser.get("temp", {}))))

    def avg3(dd):          # 3지점이 다 있는 날만
        return {d: sum(v) / len(v) for d, v in dd.items() if len(v) == len(WX_POINTS)}
    FORECAST["wx"] = {k: avg3(fc[k]) for k in fc}
    out = {k: avg3(hist[k]) for k in hist}
    LOG.append({"item": "기상", "days": len(out["temp"]), "forecast_days": len(FORECAST["wx"]["temp"])})
    log("[기상 평균 ] 실측 일수 %3d · 예보 일수 %d" % (len(out["temp"]), len(FORECAST["wx"]["temp"])))
    return out


# ══════════════════════════════════════════════════════════════════
#  판정
# ══════════════════════════════════════════════════════════════════
KEYS_CMP = ["px_PL", "px_DE", "px_CZ", "px_SK", "px_LT", "px_SE4", "load", "wind", "solar",
            "resload", "gas", "coal", "temp", "wind100", "rad"]


def diagnose(rows):
    act = [r for r in rows if r.get("resload") is not None and r.get("px_PL") is not None]
    last = act[-1]["date"] if act else SPIKE_FROM
    inb = lambda r: BASE_FROM <= r["date"] <= BASE_TO
    ins = lambda r: SPIKE_FROM <= r["date"] <= last
    comp = {}
    for k in KEYS_CMP:
        b = mean([r.get(k) for r in rows if inb(r)])
        s = mean([r.get(k) for r in rows if ins(r)])
        comp[k] = {"base": b, "spike": s,
                   "chg_pct": ((s / b - 1) * 100) if (b not in (None, 0) and s is not None and b > 0) else None}
    base, spike = [r for r in act if inb(r)], [r for r in act if ins(r)]
    v = {"base_days": len(base), "spike_days": len(spike)}
    if len(base) >= 15 and spike:
        d0 = dt.date.fromisoformat(BASE_FROM)
        tt = lambda r: (dt.date.fromisoformat(r["date"]) - d0).days
        y = [r["px_PL"] for r in base]
        x = [r["resload"] / 1000 for r in base]
        f1, f2 = ols([x], y), ols([x, [tt(r) for r in base]], y)
        pb, ps = mean(y), mean([r["px_PL"] for r in spike])
        xb, xs = mean(x), mean([r["resload"] / 1000 for r in spike])
        tb, tsp = mean([tt(r) for r in base]), mean([tt(r) for r in spike])
        rise = ps - pb
        v.update({"px_base": pb, "px_spike": ps, "rise": rise, "res_base_gw": xb, "res_spike_gw": xs})
        if f1:
            v["fit_simple"] = {"per_gw": f1["coef"][1], "r2": f1["r2"], "explained": f1["coef"][1] * (xs - xb)}
        if f2:
            e_res, e_tr = f2["coef"][1] * (xs - xb), f2["coef"][2] * (tsp - tb)
            v["fit_trend"] = {"per_gw": f2["coef"][1], "per_day": f2["coef"][2], "r2": f2["r2"],
                              "explained_res": e_res, "explained_trend": e_tr,
                              "unexplained": rise - e_res - e_tr}
    pl, de = comp["px_PL"]["chg_pct"], comp["px_DE"]["chg_pct"]
    if pl is not None and de is not None:
        v["pl_vs_de"] = ("지역 공통(독일도 비슷하게 상승)" if (pl * de > 0 and abs(de) >= abs(pl) * 0.75)
                         else "폴란드 고유 요인 가능성(독일 변동폭이 작거나 반대)")
    return {"base": [BASE_FROM, BASE_TO], "spike": [SPIKE_FROM, last], "compare": comp, "verdict": v}


def fmt(v, n=0):
    return "-" if v is None else ("{:,.%df}" % n).format(v)


def pct_of(a, b):
    return fmt(a / b * 100, 0) if (a is not None and b) else "-"


# ══════════════════════════════════════════════════════════════════
def main():
    log("=== fetch_pl_drivers.py  start=%s  급등=%s  기준=%s~%s ===" % (START, SPIKE_FROM, BASE_FROM, BASE_TO))
    data = {}
    if TOKEN:
        for fn in (get_prices, get_load, get_gen):
            try:
                data.update(fn())
            except Exception as e:
                log("!! %s 실패: %s" % (fn.__name__, type(e).__name__))
    else:
        log("!! ENTSOE_TOKEN 비어있음 — ENTSO-E 항목 건너뜀")
    try:
        data.update(get_weather())
    except Exception as e:
        log("!! 기상 실패: %s" % type(e).__name__)

    days = sorted(set().union(*[set(v.keys()) for v in data.values()])) if data else []
    rows = []
    for d in [x for x in days if x >= START]:
        r = {"date": d}
        for k, v in data.items():
            r[k] = v.get(d)
        w = [r.get("wind_on"), r.get("wind_off")]
        r["wind"] = sum(x for x in w if x is not None) if any(x is not None for x in w) else None
        c = [r.get("hardcoal"), r.get("lignite")]
        r["coal"] = sum(x for x in c if x is not None) if any(x is not None for x in c) else None
        for sfx in ("", "_mwh"):                                  # 합계 항목 (MW 평균과 MWh 둘 다)
            for tot, parts in (("wind", ("wind_on", "wind_off")), ("coal", ("hardcoal", "lignite")),
                               ("hydro", ("hydro_ps", "hydro_ror", "hydro_res"))):
                if sfx == "" and tot in ("wind", "coal"):
                    continue
                v = [r.get(p + sfx) for p in parts]
                r[tot + sfx] = sum(x for x in v if x is not None) if any(x is not None for x in v) else None
        okm = all(r.get(k) is not None for k in ("load_mwh", "wind_mwh", "solar_mwh"))
        r["resload_mwh"] = (r["load_mwh"] - r["wind_mwh"] - r["solar_mwh"]) if okm else None
        ok = all(r.get(k) is not None for k in ("load", "wind", "solar"))
        r["resload"] = (r["load"] - r["wind"] - r["solar"]) if ok else None
        rows.append(r)
    log("\n[표] 일수 %d  (%s ~ %s)" % (len(rows), rows[0]["date"] if rows else "-", rows[-1]["date"] if rows else "-"))

    cols = ["date", "px_PL", "px_PL_max", "px_PL_min", "px_PL_neg_h", "px_DE", "px_CZ", "px_SK", "px_LT",
            "px_SE4", "px_UA", "px_UA_uah", "load", "wind", "wind_on", "wind_off", "solar", "resload", "gas", "coal", "hardcoal",
            "lignite", "biomass", "hydro", "temp", "wind100", "rad",
            "load_mwh", "wind_mwh", "wind_on_mwh", "wind_off_mwh", "solar_mwh", "solar_peak", "resload_mwh",
            "gas_mwh", "coal_mwh", "hardcoal_mwh", "lignite_mwh", "biomass_mwh", "hydro_mwh",
            "hydro_ps_mwh", "hydro_ror_mwh", "hydro_res_mwh"]
    if rows:
        with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(cols)
            for r in rows:
                w.writerow([r.get(c) if (c == "date" or r.get(c) is None) else round(r[c], 2) for c in cols])
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump({"source": "ENTSO-E A44·A65·A75 + Open-Meteo", "tz": "Europe/Warsaw",
                       "generated": dt.datetime.now(TZ).strftime("%Y-%m-%d %H:%M"),
                       "units": {"px": "EUR/MWh", "load·gen": "MW(일평균)", "*_mwh": "MWh(하루 합)", "solar_peak": "MW(하루 최고)", "temp": "°C",
                                 "wind100": "km/h", "rad": "kWh/m2(일합계)", "px_PL_neg_h": "시간"},
                       "collect": LOG, "forecast": FORECAST, "daily": rows}, f, ensure_ascii=False, indent=1)
        log("[OK] %s / %s" % (OUT_CSV, OUT_JSON))

    dg = diagnose(rows)
    with open(OUT_DIAG, "w", encoding="utf-8") as f:
        json.dump({"collect": LOG, "skipped_currency": sorted(SKIPPED_CUR), **dg}, f, ensure_ascii=False, indent=2)

    log("\n================ 급등 원인 판정 (EUR/MWh) ================")
    log("기준 %s~%s  vs  급등기 %s~%s" % (dg["base"][0], dg["base"][1], dg["spike"][0], dg["spike"][1]))
    names = {"px_PL": "폴란드 가격", "px_DE": "독일 가격", "px_CZ": "체코 가격", "px_SK": "슬로바키아",
             "px_LT": "리투아니아", "px_SE4": "스웨덴 남부", "load": "수요 MW", "wind": "풍력 MW",
             "solar": "태양광 MW", "resload": "잔여부하 MW", "gas": "가스발전 MW", "coal": "석탄발전 MW",
             "temp": "기온 °C", "wind100": "풍속 km/h", "rad": "일사 kWh/m2"}
    for k, nm in names.items():
        c = dg["compare"].get(k, {})
        log("  %-12s 기준 %10s → 급등기 %10s  (%s%%)" % (nm, fmt(c.get("base"), 1), fmt(c.get("spike"), 1),
                                                      fmt(c.get("chg_pct"), 0)))
    v = dg["verdict"]
    if v.get("rise") is not None:
        log("  상승폭 %s (기준 %s → 급등기 %s) · 잔여부하 %s → %s GW" % (
            fmt(v["rise"], 1), fmt(v["px_base"], 1), fmt(v["px_spike"], 1),
            fmt(v["res_base_gw"], 2), fmt(v["res_spike_gw"], 2)))
    if v.get("fit_simple"):
        fs = v["fit_simple"]
        log("  [단순]     1GW당 %s · 설명력 %s%% → 잔여부하 몫 %s (%s%%)" % (
            fmt(fs["per_gw"], 1), fmt((fs["r2"] or 0) * 100, 0), fmt(fs["explained"], 1),
            pct_of(fs["explained"], v["rise"])))
    if v.get("fit_trend"):
        ft = v["fit_trend"]
        log("  [추세분리] 1GW당 %s · 하루당 %s · 설명력 %s%%" % (
            fmt(ft["per_gw"], 1), fmt(ft["per_day"], 2), fmt((ft["r2"] or 0) * 100, 0)))
        log("             잔여부하(날씨·수요) %s (%s%%) | 추세(연료 등) %s (%s%%) | 남는 몫(공급사건·기타) %s (%s%%)" % (
            fmt(ft["explained_res"], 1), pct_of(ft["explained_res"], v["rise"]),
            fmt(ft["explained_trend"], 1), pct_of(ft["explained_trend"], v["rise"]),
            fmt(ft["unexplained"], 1), pct_of(ft["unexplained"], v["rise"])))
    if v.get("pl_vs_de"):
        log("  [주변국] %s" % v["pl_vs_de"])
    log("[OK] %s" % OUT_DIAG)


if __name__ == "__main__":
    main()
