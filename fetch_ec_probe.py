"""
독일 전력 선물(energy-charts price_futures_bars) 숫자 파일 위치 찾기 (2026-11-26)

1) 차트 화면(chart.htm)이 불러오는 스크립트 파일 목록을 받고
2) 스크립트 안에서 'data/' · '.json' · 'futures' 가 들어간 주소 조각을 모은 뒤
3) 그 조각과 흔한 이름 후보로 실제 숫자 파일이 있는지 시험한다.
결과는 ec_probe.json (공개 자료만). 여기서 찾은 주소로 다음에 정식 수집 코드를 만든다.
"""
import json, re, time
import datetime as dt
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "*/*", "Referer": "https://www.energy-charts.info/charts/price_futures_bars/chart.htm?l=en&c=DE"}
BASE = "https://www.energy-charts.info"
CH = BASE + "/charts/price_futures_bars"
OUT = "ec_probe.json"
RES = {"scripts": [], "refs": [], "tests": []}
Y = dt.date.today().year
W = dt.date.today().isocalendar()[1]
ses = requests.Session()


def get(url):
    try:
        r = ses.get(url, headers=UA, timeout=60)
        return r.status_code, r.text or "", r.headers.get("content-type", "")
    except requests.RequestException as e:
        return None, repr(e)[:200], ""


def main():
    st, html, _ = get(CH + "/chart.htm?l=en&c=DE&interval_select=month")
    RES["page"] = {"status": st, "len": len(html), "head": html[:1500]}
    srcs = sorted(set(re.findall(r"""src=["']?([^"'\s>]+\.js[^"'\s>]*)""", html)))
    refs = set()
    for s in srcs[:25]:
        url = s if s.startswith("http") else (BASE + s if s.startswith("/") else CH + "/" + s)
        st2, js, _ = get(url)
        RES["scripts"].append({"url": url, "status": st2, "len": len(js)})
        for m in re.findall(r"""["'`]([^"'`\s]{0,120}(?:data/|\.json|futures)[^"'`\s]{0,120})["'`]""", js):
            refs.add(m)
        time.sleep(0.5)
    RES["refs"] = sorted(refs)[:200]
    cands = set()
    for r in refs:
        if r.endswith(".json") and "{" not in r and "$" not in r:
            cands.add(r if r.startswith("http") else (BASE + r if r.startswith("/") else CH + "/" + r))
    for c in ["de", "DE", "de-lu"]:
        for nm in ["month", "quarter", "year", "month_%d" % Y, "week_%d_%02d" % (Y, W), "all", "index", "latest"]:
            cands.add("%s/data/%s/%s.json" % (CH, c, nm))
    cands.add(CH + "/data/")
    cands.add(CH + "/data/de/")
    for url in sorted(cands)[:80]:
        st3, txt, ct = get(url)
        RES["tests"].append({"url": url, "status": st3, "len": len(txt), "ctype": ct, "head": txt[:300] if st3 == 200 else ""})
        print(st3, len(txt), url, flush=True)
        time.sleep(0.4)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(dict(RES, fetched_utc=dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")), f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
