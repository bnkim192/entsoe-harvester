"""
TGE GPI 'Zestawienie ubytkow' (폴란드 발전설비 출력 감소 공시) 수집기

출처  https://gpi.tge.pl/zestawienie-ubytkow  (공개 페이지, 키 불필요)

첫 판은 페이지 구조를 모르는 채로 짠다 (회사 PC에서는 사이트가 막혀 있어 미리 볼 수 없음).
그래서 표를 일반적인 방법으로 읽고, 원본 HTML과 진단 파일을 같이 남긴다.
첫 실행 결과(gpi_probe.json, gpi_raw/*.html)를 보고 다음 판에서 열 이름을 확정한다.

산출물
  pl_outages_latest.json    오늘부터 +7일까지 날짜별 표 (머리글, 행, 자동 합계)
  pl_outages_history.jsonl  하루 1줄 누적 (조회일 = 오늘인 표만, 같은 날 재실행 시 덮어씀)
  gpi_probe.json            진단 (상태코드, 표 개수, 머리글, 페이지에 보이는 날짜, 내보내기·AJAX 링크 후보)
  gpi_raw/YYYY-MM-DD.html   오늘 날짜 원본 HTML (최근 7개만 유지)
"""
import datetime as dt
import json
import os
import re
import time
from html.parser import HTMLParser

import requests

BASE = "https://gpi.tge.pl/zestawienie-ubytkow"
PID = "gpicalendar_WAR_gpicalendarportlet"
UA = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "pl,en;q=0.8",
}
DAYS_AHEAD = 7
RAW_DIR = "gpi_raw"
RAW_KEEP = 7
OUT_LATEST = "pl_outages_latest.json"
OUT_HIST = "pl_outages_history.jsonl"
OUT_PROBE = "gpi_probe.json"

# 열 자동 인식용 키워드 (폴란드어 소문자, 앞부분만). 첫 실행 뒤 실제 머리글로 확정한다.
COLS = {
    "unit": ["jednost", "nazwa", "blok", "obiekt", "unit"],
    "mw": ["ubyt", "niedost", "wielko", "unavail", "mw"],
    "type": ["rodzaj", "typ", "charakter", "planow"],
    "start": ["począt", "poczat", "od", "start", "rozpocz"],
    "end": ["koniec", "zakończ", "zakoncz", "do", "end"],
    "reason": ["przyczyn", "powód", "powod", "reason", "uwag"],
    "company": ["uczestnik", "podmiot", "spółk", "spolk", "firma", "właści", "wlasci"],
}


def log(*a):
    print("[gpi]", *a, flush=True)


def norm(s):
    return re.sub(r"\s+", " ", (s or "").replace("\xa0", " ")).strip()


class Tables(HTMLParser):
    """모든 <table>을 행·칸 목록으로 읽는다. 중첩 표는 표마다 상태를 따로 둔다."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self.stack = []
        self.title = []
        self._in_title = False
        self._skip = 0  # script/style 안

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in ("script", "style"):
            self._skip += 1
        elif tag == "title":
            self._in_title = True
        elif tag == "table":
            t = {"id": a.get("id") or "", "cls": a.get("class") or "", "rows": [], "row": None, "cell": None}
            self.tables.append(t)
            self.stack.append(t)
        elif not self.stack:
            return
        elif tag == "tr":
            self.stack[-1]["row"] = []
        elif tag in ("td", "th"):
            t = self.stack[-1]
            if t["row"] is None:
                t["row"] = []
            t["cell"] = {"th": tag == "th", "txt": [], "span": int(a.get("colspan") or 1) if str(a.get("colspan") or "1").isdigit() else 1}
        elif tag == "br" and self.stack[-1]["cell"] is not None:
            self.stack[-1]["cell"]["txt"].append(" ")

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self._skip = max(0, self._skip - 1)
        elif tag == "title":
            self._in_title = False
        elif not self.stack:
            return
        elif tag in ("td", "th"):
            self._close_cell(self.stack[-1])
        elif tag == "tr":
            t = self.stack[-1]
            self._close_cell(t)
            if t["row"]:
                t["rows"].append(t["row"])
            t["row"] = None
        elif tag == "table":
            t = self.stack.pop()
            self._close_cell(t)
            if t["row"]:
                t["rows"].append(t["row"])
            t["row"] = None

    @staticmethod
    def _close_cell(t):
        c = t["cell"]
        if c is not None and t["row"] is not None:
            txt = norm("".join(c["txt"]))
            t["row"].append((c["th"], txt))
            for _ in range(max(0, c["span"] - 1)):
                t["row"].append((c["th"], txt))
        t["cell"] = None

    def handle_data(self, d):
        if self._skip:
            return
        if self._in_title:
            self.title.append(d)
        if self.stack and self.stack[-1]["cell"] is not None:
            self.stack[-1]["cell"]["txt"].append(d)


def to_table(t):
    """머리글 행(전부 th, 없으면 첫 행)과 자료 행으로 나눈다."""
    rows = t["rows"]
    if not rows:
        return None
    hi = next((i for i, r in enumerate(rows[:3]) if r and all(c[0] for c in r)), 0)
    header = [c[1] or ("col%d" % (j + 1)) for j, c in enumerate(rows[hi])]
    data = []
    for r in rows[hi + 1:]:
        vals = [c[1] for c in r]
        if not any(vals):
            continue
        if len(vals) < len(header):
            vals += [""] * (len(header) - len(vals))
        data.append(vals[:len(header)] if len(vals) > len(header) else vals)
    return {"id": t["id"], "cls": t["cls"], "header": header, "rows": data,
            "width_mismatch": sum(1 for r in rows[hi + 1:] if len(r) != len(header))}


def find_col(header, keys):
    h = [x.lower() for x in header]
    for k in keys:
        for i, x in enumerate(h):
            if k in x:
                return i
    return None


def num(s):
    s = norm(s).replace(" ", "")
    m = re.search(r"-?\d+(?:[.,]\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0).replace(",", "."))
    except ValueError:
        return None


def summarize(tab):
    """열을 자동으로 찾아 합계를 낸다. 열을 못 찾으면 비워 둔다 (추정값을 만들지 않음)."""
    cols = {k: find_col(tab["header"], v) for k, v in COLS.items()}
    out = {"cols_guess": {k: (tab["header"][i] if i is not None else None) for k, i in cols.items()},
           "n_rows": len(tab["rows"])}
    mi = cols["mw"]
    if mi is None:
        return out
    vals = [num(r[mi]) for r in tab["rows"] if mi < len(r)]
    vals = [v for v in vals if v is not None]
    out["mw_sum"] = round(sum(vals), 1) if vals else None
    out["mw_n"] = len(vals)
    ti = cols["type"]
    if ti is not None:
        by = {}
        for r in tab["rows"]:
            if mi < len(r) and ti < len(r):
                v = num(r[mi])
                if v is not None:
                    by[r[ti] or "(빈칸)"] = round(by.get(r[ti] or "(빈칸)", 0) + v, 1)
        out["mw_by_type"] = by
    return out


DATE_RX = re.compile(r"\b(20\d\d-\d\d-\d\d|\d\d\.\d\d\.20\d\d)\b")
LINK_RX = re.compile(r"""["'](https?://[^"'\s<>]+|/[^"'\s<>]+)["']""")
LINK_KEYS = ("p_p_lifecycle=2", "p_p_resource", "resource", "ajax", "export", "eksport", "pobierz",
             ".csv", ".xls", ".json", "/api/", "download")


def link_candidates(html):
    seen = []
    for m in LINK_RX.finditer(html):
        u = m.group(1)
        lu = u.lower()
        if any(k in lu for k in LINK_KEYS) and not lu.endswith((".css", ".png", ".svg", ".ico", ".woff", ".woff2")):
            if u not in seen:
                seen.append(u)
        if len(seen) >= 40:
            break
    return seen


def fetch(sess, params=None):
    for k in range(3):
        try:
            r = sess.get(BASE, params=params, headers=UA, timeout=60)
            # charset 표기가 없으면 requests가 ISO-8859-1로 잡아 폴란드어가 깨지므로 그때는 UTF-8
            enc = r.encoding if "charset" in r.headers.get("Content-Type", "").lower() else "utf-8"
            html = r.content.decode(enc or "utf-8", errors="replace")
            return r.status_code, r.url, html
        except requests.RequestException as e:
            log("retry", k + 1, type(e).__name__)
            time.sleep(5 * (k + 1))
    return None, BASE, ""


def day_params(d, direction=None):
    p = {"p_p_id": PID, "p_p_lifecycle": "0", "p_p_state": "normal", "p_p_mode": "view",
         "p_p_col_id": "column-1", "p_p_col_count": "1",
         "_%s_current" % PID: d.isoformat()}
    if direction is not None:
        p["_%s_direction" % PID] = str(direction)
    return p


def analyze(status, url, html):
    p = Tables()
    try:
        p.feed(html)
    except Exception as e:  # 깨진 HTML도 진단은 남긴다
        log("parse error", type(e).__name__)
    tabs = [x for x in (to_table(t) for t in p.tables) if x]
    main = max(tabs, key=lambda x: len(x["rows"]), default=None)
    text = re.sub(r"<[^>]+>", " ", re.sub(r"(?is)<(script|style).*?</\1>", " ", html))
    return {
        "status": status, "url": url, "len": len(html), "title": norm("".join(p.title)),
        "n_tables": len(tabs),
        "tables": [{"id": t["id"], "cls": t["cls"], "header": t["header"], "n_rows": len(t["rows"]),
                    "width_mismatch": t["width_mismatch"], "sample": t["rows"][:3]} for t in tabs[:10]],
        "dates_seen": sorted(set(DATE_RX.findall(text)))[:20],
        "cookie_or_block": [k for k in ("cookie", "captcha", "access denied", "forbidden", "cloudflare")
                            if k in html.lower()][:5],
    }, main


def write_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


def main():
    now = dt.datetime.now(dt.timezone.utc)
    try:
        from zoneinfo import ZoneInfo
        today = now.astimezone(ZoneInfo("Europe/Warsaw")).date()
    except Exception:
        today = now.date()
    sess = requests.Session()
    probe = {"fetched_utc": now.isoformat(timespec="seconds"), "today_pl": today.isoformat(), "variants": {}}
    latest = {"source": BASE, "fetched_utc": probe["fetched_utc"], "note": "열 이름은 자동 추정. gpi_probe.json으로 확인 후 확정",
              "days": {}}

    # 1) 주소 형식 비교 (어떤 형식이 그 날짜를 보여 주는지 첫 실행에서 확인)
    variants = [("plain", None),
                ("current_today", day_params(today)),
                ("user_form_dir-1_next", day_params(today + dt.timedelta(days=1), -1))]
    today_html = ""
    for name, params in variants:
        st, url, html = fetch(sess, params)
        info, main_tab = analyze(st, url, html)
        info["main_rows"] = len(main_tab["rows"]) if main_tab else 0
        probe["variants"][name] = info
        if name == "plain":
            probe["link_candidates"] = link_candidates(html)
        if name == "current_today" and html:
            today_html = html
        log(name, st, "len", len(html), "tables", info["n_tables"], "rows", info["main_rows"], info["dates_seen"][:3])
        time.sleep(1.5)

    # 2) 오늘~+7일 날짜별 표
    for k in range(DAYS_AHEAD + 1):
        d = today + dt.timedelta(days=k)
        if k == 0 and today_html:
            st, url, html = probe["variants"]["current_today"]["status"], probe["variants"]["current_today"]["url"], today_html
        else:
            st, url, html = fetch(sess, day_params(d))
            time.sleep(1.5)
        info, tab = analyze(st, url, html)
        day = {"status": st, "dates_seen": info["dates_seen"][:5]}
        if tab:
            day.update({"header": tab["header"], "rows": tab["rows"], "summary": summarize(tab)})
        latest["days"][d.isoformat()] = day
        log(d.isoformat(), st, "rows", len(tab["rows"]) if tab else 0)

    write_json(OUT_LATEST, latest)

    # 3) 누적 (조회일 = 오늘 표만, 같은 날짜는 덮어씀)
    t0 = latest["days"].get(today.isoformat(), {})
    if t0.get("rows"):
        lines = []
        if os.path.exists(OUT_HIST):
            with open(OUT_HIST, encoding="utf-8") as f:
                for ln in f:
                    ln = ln.strip()
                    if not ln:
                        continue
                    try:
                        if json.loads(ln).get("date") != today.isoformat():
                            lines.append(ln)
                    except ValueError:
                        pass
        lines.append(json.dumps({"date": today.isoformat(), "fetched_utc": probe["fetched_utc"],
                                 "header": t0["header"], "rows": t0["rows"], "summary": t0.get("summary")},
                                ensure_ascii=False))
        with open(OUT_HIST, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        log("history lines", len(lines))
    else:
        log("history skip: 오늘 표 없음")

    # 4) 원본 HTML (구조 확인용, 최근 7개)
    if today_html:
        os.makedirs(RAW_DIR, exist_ok=True)
        with open(os.path.join(RAW_DIR, today.isoformat() + ".html"), "w", encoding="utf-8") as f:
            f.write(today_html)
        old = sorted(x for x in os.listdir(RAW_DIR) if x.endswith(".html"))[:-RAW_KEEP]
        for x in old:
            os.remove(os.path.join(RAW_DIR, x))

    write_json(OUT_PROBE, probe)
    log("done")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log("FATAL", type(e).__name__, str(e)[:200])
        write_json(OUT_PROBE, {"fatal": type(e).__name__, "msg": str(e)[:500],
                               "fetched_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")})
