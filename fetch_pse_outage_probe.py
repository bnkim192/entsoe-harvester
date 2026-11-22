"""
PSE(폴란드 송전망 운영사) 계획정지(유지보수) 자료 엔드포인트 찾기 — 1회용 프로브 (2026-11-23)

목적  PSE 공개 API(api.raporty.pse.pl)에서 '5개년 조정계획(Plan Koordynacyjny 5-letni)의 JWCD 계획 정비(Planowane remonty JWCD)'와
      '출력 감소(ubytki mocy)' 자료가 어느 엔드포인트로 나오는지 확인한다. 값 수집은 확인 후 별도 수집기에서.
방법  1) PSE 엔드포인트 지도 PDF(EndpointsMap.pdf)를 받아 글자를 뽑고, pk5l·ubyt·remont 등이 들어간 이름을 모은다.
      2) 후보 이름마다 $top=3 으로 시험 요청해 상태·열 이름·표본을 남긴다.
산출  pse_outage_probe.json   (공개 자료만)
"""
import io, re, json
import datetime as dt
import requests

API = "https://api.raporty.pse.pl/api"
PDF = "https://api.raporty.pse.pl/EndpointsMap.pdf"
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "application/json"}
KEYS = ("pk5l", "pk5", "ubyt", "remont", "jwcd", "outage", "unavail", "plan", "pkd", "bpkd", "pdgo")
GUESS = ["pk5l-wp", "pk5l-remonty", "pk5l-remonty-jwcd", "pk5l-ubytki", "pk5l-jwcd", "pk5l-wp-jwcd", "ubytki", "ubytki-mocy",
         "pkd-ubytki", "pk5l", "plan-remontow", "remonty-jwcd", "pdgobpkd", "pdgopkd"]
OUT = "pse_outage_probe.json"


def main():
    out = {"fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"), "pdf": {}, "tries": []}
    names = []
    try:
        r = requests.get(PDF, headers={"User-Agent": UA["User-Agent"]}, timeout=90)
        out["pdf"] = {"status": r.status_code, "bytes": len(r.content)}
        if r.status_code == 200:
            try:
                from pypdf import PdfReader
                txt = "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(r.content)).pages)
                out["pdf"]["chars"] = len(txt)
                toks = sorted(set(re.findall(r"[a-z0-9]+(?:-[a-z0-9]+){1,5}", txt.lower())))
                out["pdf"]["all_names"] = toks[:400]
                names = [t for t in toks if any(k in t for k in KEYS)]
                lines = [ln.strip() for ln in txt.splitlines() if any(k in ln.lower() for k in ("ubyt", "remont", "pk5", "jwcd", "plan koord"))]
                out["pdf"]["lines"] = lines[:120]
            except Exception as e:
                out["pdf"]["parse_error"] = repr(e)
    except requests.RequestException as e:
        out["pdf"] = {"error": type(e).__name__}
    for n in list(dict.fromkeys(names + GUESS))[:60]:
        try:
            r = requests.get("%s/%s" % (API, n), params={"$top": "3"}, headers=UA, timeout=60)
            rec = {"name": n, "status": r.status_code}
            if r.status_code == 200:
                try:
                    js = r.json()
                    vals = js.get("value", js if isinstance(js, list) else [])
                    rec["rows"] = len(vals)
                    rec["columns"] = sorted(vals[0].keys()) if vals and isinstance(vals[0], dict) else []
                    rec["sample"] = vals[:3]
                    rec["next"] = js.get("@odata.nextLink")
                except ValueError:
                    rec["head"] = r.text[:300]
            else:
                rec["head"] = r.text[:200]
            out["tries"].append(rec)
        except requests.RequestException as e:
            out["tries"].append({"name": n, "error": type(e).__name__})
    ok = [t["name"] for t in out["tries"] if t.get("status") == 200]
    out["ok"] = ok
    print("[pse] PDF %s · 후보 %d · 성공 %s" % (out["pdf"].get("status"), len(out["tries"]), ok), flush=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1, default=str)


if __name__ == "__main__":
    main()
