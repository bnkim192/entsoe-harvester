"""
TTF 가스 · EU 탄소배출권(EUA) 선물 가격(월물·연물 곡선) 수집기  (2026-11-18 신규, 법인 요청)

자료  ICE 공개 지연시세(무료, 로그인 없음). 공식 API 문서는 없어 결과를 fwd_curves_probe.json 에 남긴다.
  TTF  Dutch TTF Natural Gas Futures (ICE Endex)  productId 4331 · hubId 7979 (공개 사례로 확인된 번호)
  EUA  EUA Futures (ICE Endex, 상품번호 197)       번호 미확인 → 상품 페이지에서 productId·hubId 를 찾아 시도

산출물
  fwd_curves.json            {fetched_utc, ttf:{unit, contracts:[{strip, month, price, time}]}, eua:{...}}
  fwd_curves_history.jsonl   하루 1줄 누적(같은 날 재실행 시 덮어씀) → 1주 전 대비 변화 계산용
  fwd_curves_probe.json      요청별 상태·건수·앞부분
"""
import os, re, json, time
import datetime as dt
import requests

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "application/json, text/javascript, */*; q=0.01", "X-Requested-With": "XMLHttpRequest"}
DELAYED = "https://www.ice.com/marketdata/DelayedMarkets.shtml"
PAGES = {"ttf": "https://www.ice.com/products/27996665/Dutch-TTF-Natural-Gas-Futures/data",
         "eua": "https://www.ice.com/products/197/EUA-Futures/data"}
KNOWN = {"ttf": [(4331, 7979)], "eua": []}
UNIT = {"ttf": "EUR/MWh", "eua": "EUR/tCO2"}
OUT, HIST, PROBE = "fwd_curves.json", "fwd_curves_history.jsonl", "fwd_curves_probe.json"
MON = {m: i + 1 for i, m in enumerate(["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"])}


def log(*a):
    print("[fwd]", *a, flush=True)


def fnum(v):
    try:
        return float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return None


def strip_month(s):
    """'Dec26' / 'DEC 26' / 'Dec-2026' → '2026-12'. 분기·연물 등은 None."""
    m = re.match(r"^\s*([A-Za-z]{3})[\s\-]?(\d{2}|\d{4})\s*$", str(s or ""))
    if not m or m.group(1).upper() not in MON:
        return None
    y = int(m.group(2))
    y = y + 2000 if y < 100 else y
    return "%04d-%02d" % (y, MON[m.group(1).upper()])


def get_json(pid, hid, probe, tag):
    for k in range(3):
        try:
            r = requests.get(DELAYED, params={"getContractsAsJson": "", "productId": pid, "hubId": hid},
                             headers=UA, timeout=(10, 60))
            ok = r.status_code == 200 and r.text.strip()[:1] in "[{"
            probe.append({"tag": tag, "pid": pid, "hid": hid, "try": k + 1, "status": r.status_code,
                          "len": len(r.text), "head": r.text[:300]})
            if ok:
                js = r.json()
                return js if isinstance(js, list) else js.get("data") or js.get("contracts") or []
        except (requests.RequestException, ValueError) as e:
            probe.append({"tag": tag, "pid": pid, "hid": hid, "try": k + 1, "error": type(e).__name__})
        time.sleep(5)
    return []


def discover(key, probe):
    """상품 페이지 HTML 에서 productId·hubId 숫자 쌍을 찾는다."""
    pairs = []
    try:
        r = requests.get(PAGES[key], headers={"User-Agent": UA["User-Agent"]}, timeout=(10, 60))
        txt = r.text
        pids = re.findall(r"productId[\"'\s:=]+(\d+)", txt)
        hids = re.findall(r"hubId[\"'\s:=]+(\d+)", txt)
        for a, b in re.findall(r"productId=(\d+)&(?:amp;)?hubId=(\d+)", txt):
            pairs.append((int(a), int(b)))
        if pids and hids:
            pairs.append((int(pids[0]), int(hids[0])))
        probe.append({"tag": key + " page", "status": r.status_code, "len": len(txt),
                      "productId": sorted(set(pids))[:10], "hubId": sorted(set(hids))[:10]})
    except requests.RequestException as e:
        probe.append({"tag": key + " page", "error": type(e).__name__})
    seen, out = set(), []
    for p in pairs:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


def parse(rows):
    out = []
    for d in rows:
        if not isinstance(d, dict):
            continue
        strip = d.get("marketStrip") or d.get("strip") or d.get("contract") or d.get("marketName")
        price = next((fnum(d.get(k)) for k in ("lastPrice", "last", "settlementPrice", "settle", "lastSettle")
                      if fnum(d.get(k)) is not None), None)
        if not strip or price is None:
            continue
        out.append({"strip": str(strip).strip(), "month": strip_month(strip), "price": price,
                    "time": d.get("lastTime") or d.get("tradeDate") or d.get("endDate"),
                    "volume": fnum(d.get("volume"))})
    return out


def main():
    now = dt.datetime.now(dt.timezone.utc)
    probe, res = [], {"fetched_utc": now.isoformat(timespec="seconds"),
                      "source": "ICE 공개 지연시세(DelayedMarkets). 최근 체결가 기준, 정산가 아님"}
    for key in ("ttf", "eua"):
        cands = KNOWN[key] + [p for p in discover(key, probe) if p not in KNOWN[key]]
        got, used = [], None
        for pid, hid in cands[:6]:
            got = parse(get_json(pid, hid, probe, key))
            if got:
                used = (pid, hid)
                break
        res[key] = {"unit": UNIT[key], "productId": used[0] if used else None, "hubId": used[1] if used else None,
                    "contracts": got}
        log(key, "계약 %d개" % len(got), "번호 %s" % (used,), "후보 %s" % cands[:6])
        time.sleep(2)

    old = {}
    if os.path.exists(OUT):
        try:
            with open(OUT, encoding="utf-8") as f:
                old = json.load(f)
        except ValueError:
            old = {}
    for key in ("ttf", "eua"):                       # 이번에 못 받으면 예전 값을 그대로 두고 표시
        if not res[key]["contracts"] and (old.get(key) or {}).get("contracts"):
            res[key] = dict(old[key], stale=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)

    day = now.date().isoformat()
    line = {"date": day}
    for key in ("ttf", "eua"):
        if not res[key].get("stale"):
            line[key] = {c["month"] or c["strip"]: c["price"] for c in res[key]["contracts"]}
    lines = []
    if os.path.exists(HIST):
        with open(HIST, encoding="utf-8") as f:
            for ln in f:
                ln = ln.strip()
                try:
                    if ln and json.loads(ln).get("date") != day:
                        lines.append(ln)
                except ValueError:
                    pass
    if len(line) > 1:
        lines.append(json.dumps(line, ensure_ascii=False))
    with open(HIST, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + ("\n" if lines else ""))
    with open(PROBE, "w", encoding="utf-8") as f:
        json.dump({"fetched_utc": res["fetched_utc"], "log": probe}, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
