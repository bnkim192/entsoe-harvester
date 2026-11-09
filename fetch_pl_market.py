# -*- coding: utf-8 -*-
"""
폴란드 주간 보고용 원자재·환율 수집기 (2026-11-09 v2 — 실패 시 재시도·직전 자료 유지)

■ 수집 항목 (전부 무료, 키 없음)
  탄소배출권 EUA   EEX 1차시장 경매 결과 xlsx (경매일마다 1건 이상 → 일별 평균)
                   주소는 eia-harvester 프로브가 EEX 페이지에서 찾은 것:
                   public.eex-group.com/eex/eua-auction-report/emission-spot-primary-market-auction-report-YYYY-data.xlsx
                   + 페이지를 다시 읽어 주소가 바뀌었으면 새 주소도 시도한다.
  브렌트유         미국 EIA 일별 현물 xls (RBRTEd.xls — 키 불필요, EIA 는 주 1회 갱신)
  가스 TTF·석탄    World Bank 월평균 — eia-harvester 가 이미 받아 둔 fuel_monthly.json 을 그대로 읽는다
                   (가스·석탄의 무료 일별 자료는 아직 못 찾았다 — Yahoo 429, Stooq 차단, EEX 데이터 서버 접속 불가)
  환율             ECB 기준환율 EUR→PLN·KRW·USD (일별)
  TGE (폴란드 거래소) 공개 페이지 — 구조 확인용 프로브만. 값은 아직 쓰지 않는다(본문 보고 파서를 짠다).

■ 원칙
  1. 값을 만들어내지 않는다. 못 받은 항목은 로그와 collect 에 이유를 남긴다.
  2. 종료코드는 항상 0 — 다른 수집을 막지 않는다.
  3. 서버 오류(예: ECB 504)로 한 항목을 통째로 못 받으면, 저장소에 있던 직전 수집분을 그대로 둔다
     (이번 실행에서 비면 보고서의 PLN 환산·가격 구성이 전부 빠지므로). collect 에 '직전 자료 유지'로 남긴다.

■ 산출
  pl_market_daily.json   {eua, brent, fx, monthly, collect}
  tge_probe.json         TGE 페이지 구조 (다음 판 파서용)

■ 환경변수 (선택)
  PLD_START  수집 시작일 (기본 2026-01-01)
"""
import os, io, re, csv, json, time
import datetime as dt
import requests

START = os.environ.get("PLD_START", "").strip() or "2026-01-01"
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0 Safari/537.36"}
EEX_PAGE = "https://www.eex.com/en/market-data/environmental-markets/eua-primary-auction-spot-download"
EEX_FILE = ("https://public.eex-group.com/eex/eua-auction-report/"
            "emission-spot-primary-market-auction-report-%d-data.xlsx")
EIA_BRENT = "https://www.eia.gov/dnav/pet/hist_xls/RBRTEd.xls"
FUEL_RAW = "https://raw.githubusercontent.com/bnkim192/eia-harvester/main/fuel_monthly.json"
ECB = "https://data-api.ecb.europa.eu/service/data/EXR/D.PLN+KRW+USD.EUR.SP00.A"
MON_KEYS = ["WB_TTF_EU", "WB_COAL_AU", "WB_COAL_ZA", "WB_HH_US", "WB_LNG_JP", "EUA_EUA_EEXauc"]
TGE_PAGES = [("gas_rdn", "https://tge.pl/gaz-rdn"),
             ("power_rdn", "https://tge.pl/energia-elektryczna-rdn"),
             ("power_otf", "https://tge.pl/energia-elektryczna-otf"),
             ("gas_otf", "https://tge.pl/gaz-otf")]
OUT = "pl_market_daily.json"
OUT_PROBE = "tge_probe.json"
LOG = []


def log(*a):
    print(*a, flush=True)


def rec(item, ok, note):
    LOG.append({"item": item, "ok": bool(ok), "note": note})
    log("[%s] %s %s" % ("OK" if ok else "--", item, note))


def get(url, **kw):
    kw.setdefault("timeout", 90)
    kw.setdefault("headers", UA)
    return requests.get(url, **kw)


# ══════════════════════════════════════════════════════════════════
#  탄소배출권 — EEX 경매
# ══════════════════════════════════════════════════════════════════
def parse_eex(content):
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    acc, info = {}, []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        if not rows:
            continue
        ncol = max(len(r) for r in rows[:60])
        top = rows[:15]
        hdr = [" ".join(str(r[c]) for r in top if c < len(r) and isinstance(r[c], str)).lower()
               for c in range(ncol)]
        pcol = next((c for c, h in enumerate(hdr) if "auction price" in h), None)
        if pcol is None:
            info.append("%s: 'auction price' 열 없음" % ws.title)
            continue
        cnt = [sum(1 for r in rows if c < len(r) and isinstance(r[c], dt.date)) for c in range(ncol)]
        dcol = max(range(ncol), key=lambda c: cnt[c])
        if cnt[dcol] == 0:
            info.append("%s: 날짜열 없음" % ws.title)
            continue
        used = skip = 0
        for r in rows:
            if dcol >= len(r) or pcol >= len(r):
                continue
            d, p = r[dcol], r[pcol]
            if not isinstance(d, dt.date) or not isinstance(p, (int, float)) or not (5 < p < 400):
                continue
            txt = " ".join(str(x) for x in r if isinstance(x, str)).upper()
            if "EUAA" in txt or "ETS2" in txt or "ETS 2" in txt:
                skip += 1                      # 항공용·ETS2 경매 제외
                continue
            acc.setdefault(d.strftime("%Y-%m-%d"), []).append(float(p))
            used += 1
        info.append("%s: 날짜열%d 가격열%d(%s) 사용%d 제외%d" % (ws.title, dcol, pcol, hdr[pcol][:40], used, skip))
    return {d: sum(v) / len(v) for d, v in acc.items()}, info


def get_eua():
    urls = [EEX_FILE % y for y in range(int(START[:4]), dt.date.today().year + 1)]
    try:
        r = get(EEX_PAGE)
        found = re.findall(r'(https?://[^"\'\s<>]+auction-report[^"\'\s<>]*\.xlsx)', r.text, re.I)
        for u in found:
            if u not in urls:
                urls.append(u)
        rec("EUA 페이지", r.status_code == 200, "HTTP %s · 새로 찾은 주소 %d건" % (r.status_code, len(found)))
    except Exception as e:
        rec("EUA 페이지", False, type(e).__name__)
    out = {}
    for u in urls:
        try:
            r = get(u, timeout=120)
            if r.status_code != 200 or not r.content[:2] == b"PK":
                rec("EUA " + u.rsplit("/", 1)[-1], False, "HTTP %s len=%d" % (r.status_code, len(r.content)))
                continue
            d, info = parse_eex(r.content)
            out.update({k: v for k, v in d.items() if k >= START})
            rec("EUA " + u.rsplit("/", 1)[-1], bool(d), "%d일 | %s" % (len(d), " / ".join(info)))
        except Exception as e:
            rec("EUA " + u.rsplit("/", 1)[-1], False, type(e).__name__)
    if out:
        k = sorted(out)
        log("   EUA %d일 %s~%s 최근 %.2f €/t" % (len(k), k[0], k[-1], out[k[-1]]))
    return out


# ══════════════════════════════════════════════════════════════════
#  브렌트유 — EIA xls
# ══════════════════════════════════════════════════════════════════
def get_brent():
    out = {}
    try:
        r = get(EIA_BRENT, timeout=120)
        if r.status_code != 200:
            rec("브렌트 EIA", False, "HTTP %s" % r.status_code)
            return out
        import xlrd
        bk = xlrd.open_workbook(file_contents=r.content)
        names = bk.sheet_names()
        sh = bk.sheet_by_name("Data 1") if "Data 1" in names else bk.sheet_by_index(min(1, bk.nsheets - 1))
        for i in range(sh.nrows):
            if sh.ncols < 2:
                break
            c0, c1 = sh.cell(i, 0), sh.cell(i, 1)
            if c1.ctype != xlrd.XL_CELL_NUMBER:
                continue
            if c0.ctype == xlrd.XL_CELL_DATE or (c0.ctype == xlrd.XL_CELL_NUMBER and 30000 < c0.value < 60000):
                d = xlrd.xldate_as_datetime(c0.value, bk.datemode).date().isoformat()
                if d >= START:
                    out[d] = float(c1.value)
        k = sorted(out)
        rec("브렌트 EIA", bool(out), "시트=%s %d일 %s~%s" % (sh.name, len(k), k[0] if k else "-", k[-1] if k else "-"))
    except Exception as e:
        rec("브렌트 EIA", False, type(e).__name__)
    return out


# ══════════════════════════════════════════════════════════════════
#  월평균 가스·석탄 — eia-harvester 산출물 재사용
# ══════════════════════════════════════════════════════════════════
def get_monthly():
    out = {}
    try:
        r = get(FUEL_RAW, timeout=60)
        if r.status_code != 200:
            rec("월평균 fuel_monthly", False, "HTTP %s" % r.status_code)
            return out
        data = r.json().get("data", {})
        for k in MON_KEYS:
            s = data.get(k)
            if s and s.get("points"):
                out[k] = {"label": s.get("label", k), "unit": s.get("unit", ""), "source": s.get("source", ""),
                          "points": sorted(s["points"], key=lambda p: p["period"])}
        rec("월평균 fuel_monthly", bool(out), "계열 %s · 최신 %s" % (
            list(out), {k: v["points"][-1]["period"] for k, v in out.items()}))
    except Exception as e:
        rec("월평균 fuel_monthly", False, type(e).__name__)
    return out


# ══════════════════════════════════════════════════════════════════
#  환율 — ECB
# ══════════════════════════════════════════════════════════════════
def get_fx():
    out = {"PLN": {}, "KRW": {}, "USD": {}}
    try:
        # 1/1 같은 휴일에도 '직전 영업일' 환율이 있도록 열흘 앞부터 받는다
        st = (dt.date.fromisoformat(START) - dt.timedelta(days=10)).isoformat()
        r, why = None, ""
        for k in range(3):                                  # ECB 는 가끔 504 — 30초 간격 3번
            try:
                r = get(ECB, params={"startPeriod": st, "format": "csvdata"}, timeout=60)
                if r.status_code == 200:
                    break
                why = "HTTP %s" % r.status_code
            except Exception as e:
                why = type(e).__name__
            r = None
            if k < 2:
                time.sleep(30)
        if r is None:
            rec("환율 ECB", False, why + " (3회 시도)")
            return out
        for row in csv.DictReader(io.StringIO(r.text)):
            cur, d, v = row.get("CURRENCY"), row.get("TIME_PERIOD"), row.get("OBS_VALUE")
            if cur in out and d and v:
                try:
                    out[cur][d] = float(v)
                except ValueError:
                    pass
        rec("환율 ECB", all(out.values()), " · ".join("%s %d일" % (c, len(v)) for c, v in out.items()))
    except Exception as e:
        rec("환율 ECB", False, type(e).__name__)
    return out


# ══════════════════════════════════════════════════════════════════
#  TGE 프로브 — 값은 쓰지 않는다
# ══════════════════════════════════════════════════════════════════
def probe_tge():
    res = []
    for pid, u in TGE_PAGES:
        try:
            r = get(u, timeout=60)
            t = r.text
            title = re.search(r"<title>(.*?)</title>", t, re.I | re.S)
            plain = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", re.sub(r"(?is)<(script|style).*?</\1>", " ", t)))
            hits = []
            for kw in ("TGEgasDA", "TGEgas", "TGeBase", "BASE_M", "BASE_Q", "BASE_Y", "GAS_BASE", "Kurs", "PLN/MWh"):
                for m in list(re.finditer(re.escape(kw), plain))[:2]:
                    hits.append(plain[max(0, m.start() - 60): m.end() + 140])
            item = {"id": pid, "url": u, "status": r.status_code, "len": len(t),
                    "tables": len(re.findall(r"<table", t, re.I)),
                    "title": (title.group(1).strip()[:120] if title else ""),
                    "csv_xlsx_links": re.findall(r'href="([^"]+\.(?:csv|xlsx?)[^"]*)"', t, re.I)[:10],
                    "hits": hits[:12]}
            res.append(item)
            rec("TGE " + pid, r.status_code == 200, "len=%d 표=%d 키워드=%d" % (len(t), item["tables"], len(hits)))
        except Exception as e:
            res.append({"id": pid, "url": u, "error": type(e).__name__})
            rec("TGE " + pid, False, type(e).__name__)
    with open(OUT_PROBE, "w", encoding="utf-8") as f:
        json.dump({"generated": dt.datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"), "pages": res},
                  f, ensure_ascii=False, indent=2)


# ══════════════════════════════════════════════════════════════════
def main():
    log("=== fetch_pl_market.py  start=%s ===" % START)
    out = {"generated": dt.datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")}
    try:
        with open(OUT, encoding="utf-8") as f:
            prev = json.load(f)                             # 저장소에 있던 직전 수집분
    except Exception:
        prev = {}
    for key, fn in (("eua", get_eua), ("brent", get_brent), ("monthly", get_monthly), ("fx", get_fx)):
        try:
            out[key] = fn()
        except Exception as e:
            out[key] = {}
            rec(key, False, "전체 실패 %s" % type(e).__name__)
        old = prev.get(key) or {}
        if key == "fx":                                     # 통화별로: 빈 통화만 직전 값으로
            for c in ("PLN", "KRW", "USD"):
                out[key].setdefault(c, {})
            kept = [c for c in out[key] if not out[key][c] and old.get(c)]
            for c in kept:
                out[key][c] = old[c]
            if kept:
                rec("환율 ECB", False, "이번 수집 실패 → 직전 자료 유지(%s 수집분): %s" % (prev.get("generated", "?"), ", ".join(kept)))
        elif not out[key] and old:
            out[key] = old
            rec(key, False, "이번 수집 실패 → 직전 자료 유지(%s 수집분)" % prev.get("generated", "?"))
    try:
        probe_tge()
    except Exception as e:
        rec("TGE", False, type(e).__name__)
    out["units"] = {"eua": "EUR/t (경매가, 같은 날 여러 경매는 평균)", "brent": "USD/bbl",
                    "fx": "통화 per 1 EUR (ECB 기준환율)", "monthly": "계열별 unit 참조"}
    out["collect"] = LOG
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    log("[OK] %s" % OUT)


if __name__ == "__main__":
    main()
