"""
폴란드 국경 간 전력 흐름(수입·수출) 수집기

1) ENTSO-E A11 실제 물리 흐름 (Cross-Border Physical Flows) — 국경별, 방향별
   폴란드 ↔ 독일·체코·슬로바키아·리투아니아·스웨덴 남부(SE4)·우크라이나
   in_Domain 쪽으로 들어오는 흐름. 폴란드가 in 이면 수입, out 이면 수출.
2) PSE(폴란드 송전망 운영사) 공개 API — 첫 판은 구조 확인용 프로브만 (값은 아직 쓰지 않음)

산출물
  pl_flows_daily.csv / .json   현지 날짜별 국경별 수입·수출·순수입 (GWh), 전체 합계
  pl_flows_monthly.json        월별 합계 (GWh)
  pl_flows_diag.json           요청별 결과 (자료 없음·오류 확인용)
  pse_probe.json               PSE API 엔드포인트 목록과 국경 흐름 후보 표본

환경변수
  ENTSOE_TOKEN  (GitHub Secrets)
  PLF_START     수집 시작일 (기본 2026-01-01)
"""
import os, json, csv, time
import datetime as dt
import xml.etree.ElementTree as ET
from zoneinfo import ZoneInfo
import requests

TOKEN = os.environ.get("ENTSOE_TOKEN", "").strip()
BASE_URL = "https://web-api.tp.entsoe.eu/api"
PL = "10YPL-AREA-----S"
# 국경 상대편. 첫 코드로 자료가 없으면 다음 코드로 다시 시도한다.
BORDERS = {
    "DE": ["10Y1001A1001A82H", "10YDE-VE-------2"],      # DE-LU 입찰구역, 50Hertz 제어구역
    "CZ": ["10YCZ-CEPS-----N"],
    "SK": ["10YSK-SEPS-----K"],
    "LT": ["10YLT-1001A0008Q"],
    "SE4": ["10Y1001A1001A47J"],
    "UA": ["10Y1001C--000182", "10Y1001C--00003F"],     # UA IPS, UA
}
START = os.environ.get("PLF_START", "").strip() or "2026-01-01"
TZ = ZoneInfo("Europe/Warsaw")
UTC = dt.timezone.utc
MIN_HOURS = 20
PSE_API = "https://api.raporty.pse.pl/api"
PSE_KEYS = ("przep", "transgr", "wymian", "sald", "flow", "exch", "crb", "import", "eksport", "export", "granic")
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}

OUT_CSV = "pl_flows_daily.csv"
OUT_JSON = "pl_flows_daily.json"
OUT_MON = "pl_flows_monthly.json"
OUT_DIAG = "pl_flows_diag.json"
OUT_PSE = "pse_probe.json"
DIAG = []


def log(*a):
    print("[flows]", *a, flush=True)


# ══════════════════════════════════════════════════════════════════
#  ENTSO-E
# ══════════════════════════════════════════════════════════════════
def windows():
    """START ~ 내일(UTC)을 31일 단위로 자른다."""
    s = dt.datetime.strptime(START, "%Y-%m-%d")
    end = dt.datetime.now(UTC).replace(tzinfo=None, minute=0, second=0, microsecond=0) + dt.timedelta(days=1)
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
            log("  [%s] 접속 실패 (%s)" % (tag, type(e).__name__))
            time.sleep(5)
            continue
        txt = r.text
        if r.status_code == 200 and txt.lstrip().startswith("<"):
            if "Acknowledgement_MarketDocument" in txt[:400]:
                i, j = txt.find("<text>"), txt.find("</text>")
                DIAG.append({"tag": tag, "ok": False, "note": "자료 없음 " + (txt[i + 6:j][:100] if 0 <= i < j else "")})
                return None
            DIAG.append({"tag": tag, "ok": True, "note": "len=%d" % len(txt)})
            return txt
        DIAG.append({"tag": tag, "ok": False, "note": "HTTP %s %s" % (r.status_code, txt[:120].replace("\n", " "))})
        if r.status_code in (429, 503):
            time.sleep(15)
            continue
        return None
    return None


def parse_t(s):
    return dt.datetime.strptime(s.strip()[:16], "%Y-%m-%dT%H:%M").replace(tzinfo=UTC)


def parse_flow(xml):
    """[(utc 시각, MW, 분해능(분))]"""
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
            if st is None or not st.text:
                continue
            rs = per.find(q("resolution"))
            res = rs.text if rs is not None and rs.text else "PT60M"
            step = 15 if "PT15M" in res else (30 if "PT30M" in res else 60)
            t0 = parse_t(st.text)
            nslot = int((parse_t(en.text) - t0).total_seconds() // (step * 60)) if en is not None and en.text else None
            pts = []
            for pt in per.iter(q("Point")):
                pos, val = pt.find(q("position")), pt.find(q("quantity"))
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
                        out.append((t0 + dt.timedelta(minutes=step * (k - 1)), v, step))
            else:
                for p, v in pts:
                    out.append((t0 + dt.timedelta(minutes=step * (p - 1)), v, step))
    return out


def fetch_dir(in_eic, out_eic, tag):
    """{(분해능, utc시각): MW}"""
    pts = {}
    for s, e in windows():
        x = entsoe({"documentType": "A11", "in_Domain": in_eic, "out_Domain": out_eic,
                    "periodStart": s.strftime("%Y%m%d%H%M"), "periodEnd": e.strftime("%Y%m%d%H00")},
                   "%s %s" % (tag, s.strftime("%Y-%m")))
        time.sleep(0.3)
        if not x:
            continue
        try:
            for t, v, step in parse_flow(x):
                pts[(step, t)] = v
        except ET.ParseError as ex:
            DIAG.append({"tag": tag, "ok": False, "note": "XML 해석 실패 %s" % ex})
    return pts


def daily_gwh(pts):
    """{(분해능, utc시각): MW} → {현지날짜: GWh}. 한 날짜에 분해능이 섞이면 가장 촘촘한 것, 20시간 미만인 날은 버림."""
    g = {}
    for (step, t), v in pts.items():
        d = t.astimezone(TZ).strftime("%Y-%m-%d")
        g.setdefault(d, {}).setdefault(step, []).append(v)
    out = {}
    for d, bys in g.items():
        ok = [s for s, vs in bys.items() if len(vs) * s / 60 >= MIN_HOURS]
        if ok:
            s = min(ok)
            out[d] = sum(bys[s]) * s / 60 / 1000.0
    return out


def get_flows():
    res, used = {}, {}
    for name, codes in BORDERS.items():
        for code in codes:
            imp = daily_gwh(fetch_dir(PL, code, "A11 %s→PL" % name))
            exp = daily_gwh(fetch_dir(code, PL, "A11 PL→%s" % name))
            if imp or exp:
                res[name] = {"imp": imp, "exp": exp}
                used[name] = code
                log(name, code, "수입 %d일 수출 %d일" % (len(imp), len(exp)))
                break
            log(name, code, "자료 없음")
    return res, used


# ══════════════════════════════════════════════════════════════════
#  PSE 프로브
# ══════════════════════════════════════════════════════════════════
def probe_pse():
    out = {"base": PSE_API, "fetched_utc": dt.datetime.now(UTC).isoformat(timespec="seconds")}
    try:
        r = requests.get(PSE_API + "/", headers=UA, timeout=60)
        out["status"] = r.status_code
        out["content_type"] = r.headers.get("Content-Type", "")
        try:
            js = r.json()
        except ValueError:
            out["head"] = r.text[:1500]
            return out
        names = [v.get("name") or v.get("url") for v in js.get("value", []) if isinstance(v, dict)]
        out["entities"] = names
        cand = [n for n in names if n and any(k in n.lower() for k in PSE_KEYS)]
        out["candidates"] = {}
        for n in cand[:12]:
            try:
                r2 = requests.get("%s/%s" % (PSE_API, n), params={"$top": "5"}, headers=UA, timeout=60)
                item = {"status": r2.status_code}
                try:
                    j2 = r2.json()
                    vals = j2.get("value", []) if isinstance(j2, dict) else []
                    item["keys"] = sorted(vals[0].keys()) if vals and isinstance(vals[0], dict) else []
                    item["sample"] = vals[:3]
                    if not vals:
                        item["head"] = json.dumps(j2, ensure_ascii=False)[:600]
                except ValueError:
                    item["head"] = r2.text[:600]
                out["candidates"][n] = item
            except Exception as e:
                out["candidates"][n] = {"error": type(e).__name__}
            time.sleep(0.5)
    except Exception as e:
        out["error"] = type(e).__name__ + " " + str(e)[:200]
    return out


# ══════════════════════════════════════════════════════════════════
def main():
    if TOKEN:
        flows, used = get_flows()
        days = sorted({d for b in flows.values() for k in ("imp", "exp") for d in b[k]})
        names = list(flows.keys())
        rows = []
        for d in days:
            row = {"date": d}
            ti = te = 0.0
            complete = True
            for n in names:
                i, e = flows[n]["imp"].get(d), flows[n]["exp"].get(d)
                row[n + "_imp"] = round(i, 3) if i is not None else None
                row[n + "_exp"] = round(e, 3) if e is not None else None
                row[n + "_net"] = round(i - e, 3) if i is not None and e is not None else None
                if i is None or e is None:
                    complete = False
                else:
                    ti += i
                    te += e
            row["total_imp"] = round(ti, 3) if complete else None
            row["total_exp"] = round(te, 3) if complete else None
            row["total_net"] = round(ti - te, 3) if complete else None
            rows.append(row)
        if rows:
            cols = list(rows[0].keys())
            with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=cols)
                w.writeheader()
                w.writerows(rows)
            with open(OUT_JSON, "w", encoding="utf-8") as f:
                json.dump({"source": "ENTSO-E A11 Cross-Border Physical Flows", "unit": "GWh/일",
                           "tz": "Europe/Warsaw", "sign": "net = 수입 − 수출 (양수 = 순수입)",
                           "borders": used, "rows": rows}, f, ensure_ascii=False, indent=1)
            mon = {}
            for r in rows:
                if r["total_net"] is None:
                    continue
                m = mon.setdefault(r["date"][:7], {"days": 0, "imp": 0.0, "exp": 0.0})
                m["days"] += 1
                m["imp"] += r["total_imp"]
                m["exp"] += r["total_exp"]
            for m in mon.values():
                m["imp"], m["exp"] = round(m["imp"], 1), round(m["exp"], 1)
                m["net"] = round(m["imp"] - m["exp"], 1)
            with open(OUT_MON, "w", encoding="utf-8") as f:
                json.dump({"unit": "GWh/월", "note": "모든 국경 자료가 있는 날만 합산", "months": mon},
                          f, ensure_ascii=False, indent=1)
            log("rows", len(rows), "last", rows[-1]["date"], "net", rows[-1]["total_net"])
        else:
            log("ENTSO-E 흐름 자료 없음 (파일 유지)")
    else:
        log("ENTSOE_TOKEN 없음, ENTSO-E 건너뜀")
        DIAG.append({"tag": "token", "ok": False, "note": "ENTSOE_TOKEN 없음"})

    with open(OUT_DIAG, "w", encoding="utf-8") as f:
        json.dump({"fetched_utc": dt.datetime.now(UTC).isoformat(timespec="seconds"), "log": DIAG},
                  f, ensure_ascii=False, indent=1)

    pse = probe_pse()
    with open(OUT_PSE, "w", encoding="utf-8") as f:
        json.dump(pse, f, ensure_ascii=False, indent=1)
    log("PSE", pse.get("status"), "엔티티", len(pse.get("entities", [])), "후보", list(pse.get("candidates", {}).keys()))


if __name__ == "__main__":
    main()
