"""
폴란드 발전설비 정지(출력 감소) 수집기 — ENTSO-E 판

TGE GPI(gpi.tge.pl)는 GitHub 서버에서도 접속이 막혀(ConnectTimeout) ENTSO-E 로 바꿨다.
같은 REMIT 정지 공시를 발전사가 ENTSO-E Transparency 에도 올린다.

자료
  A80  발전기(유닛) 정지   Unavailability of Generation Units
  A77  발전소 정지          Unavailability of Production Units
  businessType A53 = 계획정지, A54 = 고장정지(불시)
  정지 MW = 정격용량(nominalP) − 정지 중 가용용량(가장 낮은 값). 가용용량이 없으면 전량 정지로 본다.
  같은 공시의 여러 판(revision)은 가장 최신판만, 취소된 공시(docStatus A09)는 뺀다.

산출물
  pl_outages_latest.json    오늘 기준 −1일 ~ +7일에 걸친 정지 목록 + 날짜별 정지 MW 합계(계획·고장, 전원별)
  pl_outages_history.jsonl  하루 1줄 누적(오늘 날짜 합계, 같은 날 재실행 시 덮어씀)
  pl_outages_diag.json      요청별 결과

환경변수  ENTSOE_TOKEN (GitHub Secrets)
"""
import os, io, json, time, zipfile
import datetime as dt
import xml.etree.ElementTree as ET
from zoneinfo import ZoneInfo
import requests

TOKEN = os.environ.get("ENTSOE_TOKEN", "").strip()
BASE_URL = "https://web-api.tp.entsoe.eu/api"
PL = "10YPL-AREA-----S"
TZ = ZoneInfo("Europe/Warsaw")
UTC = dt.timezone.utc
BACK_DAYS, AHEAD_DAYS = 1, 7
PAGE = 200                      # ENTSO-E 정지 자료는 한 번에 최대 200건, offset 으로 넘김
PSR = {"B01": "바이오", "B02": "갈탄", "B03": "석탄가스", "B04": "가스", "B05": "유연탄", "B06": "석유",
       "B10": "양수", "B11": "수력(유입식)", "B12": "수력(저수지)", "B14": "원자력", "B15": "기타 재생",
       "B16": "태양광", "B17": "폐기물", "B18": "해상풍력", "B19": "육상풍력", "B20": "기타"}
BT = {"A53": "계획", "A54": "고장"}

OUT_LATEST = "pl_outages_latest.json"
OUT_HIST = "pl_outages_history.jsonl"
OUT_DIAG = "pl_outages_diag.json"
DIAG = []


def log(*a):
    print("[outages]", *a, flush=True)


def local(tag):
    return tag.split("}")[-1]


def first(el, name):
    """하위 요소 중 이름이 name 인 첫 요소의 글자 (이름공간 무시)."""
    for x in el.iter():
        if local(x.tag) == name and x.text and x.text.strip():
            return x.text.strip()
    return None


def children(el, name):
    return [x for x in el.iter() if local(x.tag) == name]


def get_docs(doc_type, s, e):
    """XML 문서 목록. 응답이 zip 이면 안의 xml 을 모두 꺼낸다."""
    docs, offset = [], 0
    while True:
        p = {"securityToken": TOKEN, "documentType": doc_type, "biddingZone_Domain": PL,
             "periodStart": s.strftime("%Y%m%d%H%M"), "periodEnd": e.strftime("%Y%m%d%H%M"), "offset": str(offset)}
        tag = "%s offset %d" % (doc_type, offset)
        r = None
        for _ in range(2):
            try:
                r = requests.get(BASE_URL, params=p, timeout=(10, 120))
                break
            except requests.RequestException as ex:
                DIAG.append({"tag": tag, "ok": False, "note": type(ex).__name__})
                time.sleep(3)
        if r is None:
            return docs
        body = r.content
        if body[:2] == b"PK":
            with zipfile.ZipFile(io.BytesIO(body)) as z:
                got = [z.read(n).decode("utf-8", errors="replace") for n in z.namelist() if n.lower().endswith(".xml")]
        else:
            txt = body.decode("utf-8", errors="replace")
            if "Acknowledgement_MarketDocument" in txt[:400] or r.status_code != 200:
                i, j = txt.find("<text>"), txt.find("</text>")
                DIAG.append({"tag": tag, "ok": False,
                             "note": "HTTP %s %s" % (r.status_code, txt[i + 6:j][:150] if 0 <= i < j else txt[:150])})
                return docs
            got = [txt]
        DIAG.append({"tag": tag, "ok": True, "note": "문서 %d건" % len(got)})
        docs += got
        if len(got) < PAGE:
            return docs
        offset += PAGE
        time.sleep(0.5)


def ptime(d, t):
    if not d:
        return None
    t = (t or "00:00:00Z").replace("Z", "")
    return dt.datetime.strptime(d + "T" + t[:5], "%Y-%m-%dT%H:%M").replace(tzinfo=UTC)


def parse(xml, doc_type):
    root = ET.fromstring(xml)
    st = None
    for x in root.iter():
        if local(x.tag) == "docStatus":
            st = first(x, "value")
            break
    reason = "; ".join(t for t in (first(x, "text") for x in children(root, "Reason")) if t)
    out = []
    for ts in children(root, "TimeSeries"):
        s = ptime(first(ts, "start_DateAndOrTime.date"), first(ts, "start_DateAndOrTime.time"))
        e = ptime(first(ts, "end_DateAndOrTime.date"), first(ts, "end_DateAndOrTime.time"))
        try:
            nom = float(first(ts, "production_RegisteredResource.pSRType.powerSystemResources.nominalP") or "nan")
        except ValueError:
            nom = float("nan")
        avail = []
        for pt in children(ts, "Point"):
            q = first(pt, "quantity")
            try:
                avail.append(float(q))
            except (TypeError, ValueError):
                pass
        mn = min(avail) if avail else 0.0
        unav = (nom - mn) if nom == nom else None
        psr = first(ts, "psrType")
        out.append({
            "mrid": first(root, "mRID"), "rev": int(first(root, "revisionNumber") or 0), "status": st,
            "doc": doc_type, "type": BT.get(first(ts, "businessType") or "", first(ts, "businessType")),
            "plant": first(ts, "production_RegisteredResource.name"),
            "unit": first(ts, "production_RegisteredResource.pSRType.powerSystemResources.name"),
            "fuel": PSR.get(psr or "", psr), "nominal_mw": nom if nom == nom else None,
            "avail_min_mw": mn if avail else None,
            "unavail_mw": round(max(0.0, unav), 1) if unav is not None else None,
            "start_utc": s.isoformat() if s else None, "end_utc": e.isoformat() if e else None,
            "reason": reason[:200] or None,
        })
    return out


def main():
    now = dt.datetime.now(UTC)
    today = now.astimezone(TZ).date()
    if not TOKEN:
        log("ENTSOE_TOKEN 없음 — 건너뜀")
        return
    s = dt.datetime.combine(today - dt.timedelta(days=BACK_DAYS), dt.time(0), TZ).astimezone(UTC)
    e = dt.datetime.combine(today + dt.timedelta(days=AHEAD_DAYS + 1), dt.time(0), TZ).astimezone(UTC)

    recs = []
    for doc_type in ("A80", "A77"):
        docs = get_docs(doc_type, s, e)
        n0 = len(recs)
        for x in docs:
            try:
                recs += parse(x, doc_type)
            except ET.ParseError as ex:
                DIAG.append({"tag": doc_type, "ok": False, "note": "XML 해석 실패 %s" % ex})
        log(doc_type, "문서", len(docs), "정지 건", len(recs) - n0)

    # 같은 공시는 최신판만, 취소 공시 제외
    best = {}
    for r in recs:
        k = (r["mrid"], r["doc"], r["unit"] or r["plant"], r["start_utc"])
        if k not in best or r["rev"] > best[k]["rev"]:
            best[k] = r
    cur = [r for r in best.values() if r["status"] != "A09" and r["unavail_mw"]]
    cur.sort(key=lambda r: (-(r["unavail_mw"] or 0), r["plant"] or ""))

    # 날짜별 합계 (그날 하루 중 일부라도 겹치면 포함)
    daily = {}
    for k in range(-BACK_DAYS, AHEAD_DAYS + 1):
        d = today + dt.timedelta(days=k)
        d0 = dt.datetime.combine(d, dt.time(0), TZ).astimezone(UTC)
        d1 = d0 + dt.timedelta(days=1)
        tot, by_t, by_f, n = 0.0, {}, {}, 0
        for r in cur:
            rs = dt.datetime.fromisoformat(r["start_utc"]) if r["start_utc"] else None
            re_ = dt.datetime.fromisoformat(r["end_utc"]) if r["end_utc"] else None
            if (rs and rs >= d1) or (re_ and re_ <= d0):
                continue
            v = r["unavail_mw"]
            tot += v
            n += 1
            by_t[r["type"] or "미상"] = round(by_t.get(r["type"] or "미상", 0) + v, 1)
            by_f[r["fuel"] or "미상"] = round(by_f.get(r["fuel"] or "미상", 0) + v, 1)
        daily[d.isoformat()] = {"n": n, "unavail_mw": round(tot, 1), "by_type": by_t, "by_fuel": by_f}

    with open(OUT_LATEST, "w", encoding="utf-8") as f:
        json.dump({"source": "ENTSO-E A80(발전기)·A77(발전소) Unavailability", "zone": "PL",
                   "fetched_utc": now.isoformat(timespec="seconds"), "today_pl": today.isoformat(),
                   "note": "정지 MW = 정격 − 정지 중 최저 가용용량. 하루 중 일부라도 겹치면 그날 합계에 포함(최대치에 가까움). "
                           "A80과 A77에 같은 설비가 겹쳐 올라올 수 있음(첫 실행 후 확인).",
                   "daily": daily, "outages": cur}, f, ensure_ascii=False, indent=1)

    t0 = daily.get(today.isoformat())
    if t0 is not None:
        lines = []
        if os.path.exists(OUT_HIST):
            with open(OUT_HIST, encoding="utf-8") as f:
                for ln in f:
                    ln = ln.strip()
                    try:
                        if ln and json.loads(ln).get("date") != today.isoformat():
                            lines.append(ln)
                    except ValueError:
                        pass
        lines.append(json.dumps(dict({"date": today.isoformat(), "fetched_utc": now.isoformat(timespec="seconds")}, **t0),
                                ensure_ascii=False))
        with open(OUT_HIST, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")

    with open(OUT_DIAG, "w", encoding="utf-8") as f:
        json.dump({"fetched_utc": now.isoformat(timespec="seconds"), "log": DIAG}, f, ensure_ascii=False, indent=1)
    log("정지 %d건, 오늘 정지 %s MW" % (len(cur), t0["unavail_mw"] if t0 else "-"))


if __name__ == "__main__":
    main()
