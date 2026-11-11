# -*- coding: utf-8 -*-
"""
폴란드 전력단가 및 원자재 가격동향 — 주간 PPT 4장 + Excel raw (2026-11-10 v3 — 영문·숫자 Arial Narrow)
  1장 주간 요약 · 2장 가격 구성 · 3장 상승 원인 · 4장 다음 주 전망
  (2~4장은 장마다 따로 만들어, 한 장이 실패해도 나머지는 그대로 나온다)

■ 입력 (같은 실행에서 먼저 만든다)
  pl_drivers_daily.json  fetch_pl_drivers.py — 전력가격·주변국·수요·발전·기상
  pl_market_daily.json   fetch_pl_market.py  — 탄소배출권·브렌트·월평균 가스/석탄·환율
■ 산출
  weekly/<YYYY-Www>/PL_power_commodity_<YYYY-Www>.pptx · .xlsx   주차별 보관
  weekly/latest/PL_power_commodity_latest.pptx · .xlsx            주소 고정 (늘 최신호)
  weekly/latest/summary.json                                      숫자·문장 원본
■ 원칙
  1. 값은 공개자료 원값만 쓴다. 문장은 숫자에서 규칙으로 만든다 — 규칙은 s_* 함수에 그대로 있다.
  2. 공개 저장소이므로 회사·법인·계약 정보는 넣지 않는다.
  3. 주 = 폴란드 현지 월~일. 자료가 5일 미만인 주간 값은 빈칸(—).
  4. 변화가 ±5% 안이면 '비슷'으로 쓴다.
  5. PLN·원 환산은 ECB 기준환율(주말·공휴일은 직전 영업일 값). 원본 EUR 값은 Excel 에 그대로 둔다.
■ 환경변수 (선택)
  PLW_WEEK_END  보고 주의 마지막 일요일 YYYY-MM-DD (기본: 실행일 직전 일요일)
"""
import os, re, json, math, shutil, bisect, zipfile, traceback
import datetime as dt
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_MARK, XL_MARKER_STYLE
from pptx.enum.chart import XL_TICK_LABEL_POSITION
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from fetch_pl_drivers import ols, TZ

FLAT, NEED, FIT_DAYS, CHART_DAYS, CHART_MONTHS = 5.0, 5, 56, 84, 18
WD = "월화수목금토일"
KO, LAT = "LG스마트체 Regular", "Arial Narrow"   # 한글 LG스마트체 Regular · 영문/숫자 Arial Narrow (사용자 지정, 2026-11-10)
INK, INK2, MUTED, GRID, AXIS, CARD = "0B0B0B", "52514E", "898781", "E1E0D9", "C3C2B7", "F4F3EF"
COL = {"PL": "2A78D6", "DE": "EB6834", "EUA": "1BAF7A", "TTF": "6250D6", "COAL": "008300", "BRENT": "E34948"}
SRC = {"px": "ENTSO-E", "fx": "ECB", "eua": "EEX 경매", "brent": "EIA", "wb": "World Bank", "wx": "Open-Meteo"}
BULLET_SRC = {"가격": "ENTSO-E · ECB", "수급": "ENTSO-E · 자체 계산", "주변국": "ENTSO-E",
              "원자재": "EEX · EIA · World Bank", "다음 주": "Open-Meteo 예보 · ENTSO-E"}
SRC_LONG = [
    ["전력 하루전 가격", "ENTSO-E Transparency Platform, A44 Day-ahead Prices (PL·DE-LU·CZ·SK·LT·SE4)",
     "https://transparency.entsoe.eu"],
    ["수요", "ENTSO-E Transparency Platform, A65 Actual Total Load (PL)", "https://transparency.entsoe.eu"],
    ["국경 간 흐름", "ENTSO-E Transparency Platform, A11 Cross-Border Physical Flows (PL↔DE·CZ·SK·LT·SE4·UA)",
     "https://transparency.entsoe.eu"],
    ["전원별 발전", "ENTSO-E Transparency Platform, A75 Actual Generation per Production Type (PL)",
     "https://transparency.entsoe.eu"],
    ["탄소배출권 EUA", "EEX, EU ETS Primary Market Auction Report (경매가)",
     "https://www.eex.com/en/market-data/environmental-markets/eua-primary-auction-spot-download"],
    ["브렌트유", "U.S. EIA, Europe Brent Spot Price FOB (RBRTE, 일별)",
     "https://www.eia.gov/dnav/pet/hist/RBRTED.htm"],
    ["가스 TTF·석탄", "World Bank, Commodity Price Data (Pink Sheet, 월평균)",
     "https://www.worldbank.org/en/research/commodity-markets"],
    ["환율", "ECB, Euro foreign exchange reference rates (EUR→PLN·KRW·USD)",
     "https://data.ecb.europa.eu"],
    ["기상·예보", "Open-Meteo, Historical Weather API(ERA5) · Forecast API", "https://open-meteo.com"],
]
OUT_DIR = "weekly"
BASENAME = "PL_power_commodity_%s"


# ══════════════════════════════════════════════════════════════════
#  공통
# ══════════════════════════════════════════════════════════════════
def f0(v, nd=0):
    return "없음" if v is None else ("{:,.%df}" % nd).format(v)


def sg(v, nd=0, unit=""):
    if v is None:
        return "없음"
    return ("{:+,.%df}" % nd).format(v).replace("-", "−") + unit


def pct(a, b):
    return (a / b - 1) * 100 if (a is not None and b is not None and b > 0) else None


def dd(s):
    return dt.date.fromisoformat(s)


def md(s):
    x = dd(s)
    return "%d/%d" % (x.month, x.day)


def mdw(s):
    x = dd(s)
    return "%d/%d(%s)" % (x.month, x.day, WD[x.weekday()])


def drange(a, b):
    out, d = [], a
    while d <= b:
        out.append(d.isoformat())
        d += dt.timedelta(days=1)
    return out


def load(path):
    if not os.path.exists(path):
        print("!! %s 없음" % path, flush=True)
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)


class FxLine:
    """날짜별 환율 — 그날 값이 없으면 직전 영업일 값."""
    def __init__(self, src):
        self.k = sorted(src)
        self.v = [src[x] for x in self.k]

    def on(self, d):
        i = bisect.bisect_right(self.k, d) - 1
        return self.v[i] if i >= 0 else None

    def avg(self, d0, d1):
        xs = [v for k, v in zip(self.k, self.v) if d0 <= k <= d1]
        return sum(xs) / len(xs) if xs else self.on(d1)


def week_bounds():
    env = os.environ.get("PLW_WEEK_END", "").strip()
    if env:
        end = dd(env)
    else:
        today = dt.datetime.now(TZ).date()
        end = today - dt.timedelta(days=today.isoweekday())
    return end - dt.timedelta(days=6), end


# ══════════════════════════════════════════════════════════════════
#  주간 통계
# ══════════════════════════════════════════════════════════════════
def avg(R, ds, k, need=NEED):
    xs = [R[d][k] for d in ds if d in R and R[d].get(k) is not None]
    return sum(xs) / len(xs) if len(xs) >= need else None


def wstats(R, ds):
    o = {k: avg(R, ds, k) for k in ("px_PL", "px_PL_pln", "px_PL_krw", "px_DE", "px_DE_pln", "px_CZ",
                                    "px_SK", "px_LT", "px_SE4", "temp", "wind100", "rad")}
    for k in ("gas", "coal"):
        v = avg(R, ds, k)
        o[k] = v / 1000 if v is not None else None
    cd = [d for d in ds if d in R and R[d].get("resload") is not None]      # 분해가 정확히 맞도록 같은 날만
    for k, src in (("load", "load"), ("wind", "wind"), ("solar", "solar"), ("res", "resload")):
        o[k] = (sum(R[d][src] for d in cd) / len(cd) / 1000) if len(cd) >= NEED else None
    wk = [d for d in ds if dd(d).weekday() < 5]
    we = [d for d in ds if dd(d).weekday() >= 5]
    o["wkday_pln"] = avg(R, wk, "px_PL_pln", 3)
    o["wkend_pln"] = avg(R, we, "px_PL_pln", 1)
    day = [(R[d]["px_PL_pln"], d) for d in ds if d in R and R[d].get("px_PL_pln") is not None]
    o["dmax"] = max(day) if day else None
    o["dmin"] = min(day) if day else None
    ng = [R[d]["px_PL_neg_h"] for d in ds if d in R and R[d].get("px_PL_neg_h") is not None]
    o["neg_h"] = sum(ng) if len(ng) >= NEED else None
    return o


def fit_resload(R, week_start):
    """보고 주 직전 8주: 가격(PLN) ~ 잔여부하(GW) + 시간추세. 추세는 연료처럼 서서히 움직이는 몫을 떼어 낸다."""
    d1 = dd(week_start) - dt.timedelta(days=1)
    ds = drange(d1 - dt.timedelta(days=FIT_DAYS - 1), d1)
    pts = [(i, R[d]["resload"] / 1000, R[d]["px_PL_pln"]) for i, d in enumerate(ds)
           if d in R and R[d].get("resload") is not None and R[d].get("px_PL_pln") is not None]
    if len(pts) < 28:
        return {"ok": False, "why": "자료 %d일(28일 미만)" % len(pts)}
    f = ols([[p[1] for p in pts], [p[0] for p in pts]], [p[2] for p in pts])
    if not f:
        return {"ok": False, "why": "계산 불가"}
    b, r2 = f["coef"][1], f["r2"] or 0
    ok = b > 0 and r2 >= 0.2
    return {"ok": ok, "per_gw": b, "r2": r2, "n": f["n"], "from": ds[0], "to": ds[-1],
            "why": "" if ok else "관계 약함(1GW당 %s, 설명력 %.0f%%)" % (f0(b), r2 * 100)}


def decompose(W, P, fit):
    if not fit.get("ok") or None in (W["res"], P["res"], W["px_PL_pln"], P["px_PL_pln"]):
        return None
    b = fit["per_gw"]
    d = {k: W[k] - P[k] for k in ("load", "wind", "solar", "res")}
    eff = {"load": b * d["load"], "wind": -b * d["wind"], "solar": -b * d["solar"]}
    dpx = W["px_PL_pln"] - P["px_PL_pln"]
    return {"d": d, "eff": eff, "expl": b * d["res"], "dpx": dpx, "other": dpx - b * d["res"]}


def last_vs_week(src, asof):
    ks = sorted(k for k in src if k <= asof)
    if not ks:
        return None
    k = ks[-1]
    ref = (dd(k) - dt.timedelta(days=7)).isoformat()
    kr = [x for x in ks if x <= ref]
    pv = src[kr[-1]] if kr else None
    return {"date": k, "value": src[k], "prev_date": kr[-1] if kr else None, "prev": pv,
            "chg_pct": pct(src[k], pv)}


def last_month(mon, key):
    s = mon.get(key)
    if not s or len(s.get("points", [])) < 2:
        return None
    a, b = s["points"][-1], s["points"][-2]
    return {"period": a["period"], "value": a["value"], "prev_period": b["period"], "prev": b["value"],
            "chg_pct": pct(a["value"], b["value"]), "unit": UNIT_FIX.get(key, s.get("unit", "")),
            "label": MON_NAME.get(key, s.get("label", key))}


# World Bank 원문 단위 (fuel_monthly.json 에는 "USD(원문단위)" 로만 적혀 있다)
UNIT_FIX = {"WB_TTF_EU": "$/MMBtu", "WB_HH_US": "$/MMBtu", "WB_LNG_JP": "$/MMBtu",
            "WB_COAL_AU": "$/t", "WB_COAL_ZA": "$/t"}
MON_HDR = {"WB_TTF_EU": "가스 TTF 유럽 ($/MMBtu)", "WB_COAL_AU": "석탄 뉴캐슬 호주 ($/t)",
           "WB_COAL_ZA": "석탄 남아공 ($/t)", "WB_HH_US": "가스 헨리허브 미국 ($/MMBtu)",
           "WB_LNG_JP": "LNG 일본 ($/MMBtu)", "EUA_EUA_EEXauc": "탄소배출권 EUA 월평균 (EUR/t)"}
MON_NAME = {k: v.rsplit(" (", 1)[0] for k, v in MON_HDR.items()}


# ══════════════════════════════════════════════════════════════════
#  문장 — 숫자에서 규칙으로 (개조식)
# ══════════════════════════════════════════════════════════════════
NAMES = {"load": "수요", "wind": "풍력 발전", "solar": "태양광 발전"}


def s_price(W, P):
    v = W["px_PL_pln"]
    if v is None:
        return "폴란드 하루전 가격 자료 부족(5일 미만)"
    p = pct(v, P["px_PL_pln"])
    s = "주평균 %s PLN/MWh (%s EUR/MWh), " % (f0(v), f0(W["px_PL"]))
    s += "전주와 비슷(%s)" % sg(p, 0, "%") if (p is not None and abs(p) < FLAT) else "전주 대비 %s" % sg(p, 0, "%")
    ext = []
    if W["wkday_pln"] is not None and W["wkend_pln"] is not None:
        ext.append("평일 %s · 주말 %s" % (f0(W["wkday_pln"]), f0(W["wkend_pln"])))
    if W["dmax"]:
        ext.append("최고 %s %s" % (mdw(W["dmax"][1]), f0(W["dmax"][0])))
    if W["neg_h"]:
        ext.append("마이너스 가격 %s시간" % f0(W["neg_h"], 1))
    return s + (" · " + " · ".join(ext) if ext else "")


def s_cause(dec, fit, W, P):
    if dec is None:
        if not fit.get("ok"):
            return "직전 8주 가격-잔여부하 %s → 수급 분해 생략" % (fit.get("why") or "자료 부족")
        return "수요·풍력·태양광 자료 부족 → 수급 분해 생략"
    res = dec["d"]["res"]
    if abs(res) < 0.1:
        head = "화력 몫(잔여부하 = 수요 − 풍력 − 태양광) 전주와 비슷"
    else:
        cand = [k for k in dec["eff"] if dec["eff"][k] * res > 0] or list(dec["eff"])
        k = max(cand, key=lambda z: abs(dec["eff"][z]))
        dv = dec["d"][k]
        head = "%s %.1fGW %s → 화력 몫(잔여부하) %.1fGW %s" % (
            NAMES[k], abs(dv), "증가" if dv > 0 else "감소", abs(res), "증가" if res > 0 else "감소")
    dpx, ex = dec["dpx"], dec["expl"]
    if abs(dpx) < 10:
        return head
    if ex * dpx > 0:
        sh = ex / dpx * 100
        if sh >= 100:
            return head + ", 가격 변화(%s PLN/MWh)를 모두 설명" % sg(dpx)
        return head + ", 가격 변화 %s 중 약 %s PLN/MWh(%.0f%%) 설명" % (sg(dpx), sg(ex), sh)
    return head + ", 가격은 반대로 움직임(%s PLN/MWh) → 연료·발전소·주변국 등 다른 요인 우세" % sg(dpx)


def s_region(W, P):
    pl, de = pct(W["px_PL"], P["px_PL"]), pct(W["px_DE"], P["px_DE"])
    if pl is None or de is None:
        return None
    if abs(pl) < FLAT and abs(de) < FLAT:
        mv = "폴란드·독일 모두 전주와 비슷"
    elif abs(pl) < FLAT:
        mv = "독일 %s 변동, 폴란드는 비슷" % sg(de, 0, "%")
    elif pl * de > 0 and abs(de) >= 0.75 * abs(pl):
        mv = "독일 %s로 동반 %s → 지역 공통 요인(날씨·연료)" % (sg(de, 0, "%"), "상승" if de > 0 else "하락")
    else:
        mv = "독일 %s 대비 폴란드 %s → 폴란드 내부 요인 우세" % (sg(de, 0, "%"), sg(pl, 0, "%"))
    if W["px_PL_pln"] is not None and W["px_DE_pln"] is not None:
        mv += " · 폴란드−독일 가격차 %s PLN/MWh" % sg(W["px_PL_pln"] - W["px_DE_pln"])
    return mv


def s_commod(K):
    parts = []
    for key, nm, unit, nd in (("EUA", "탄소배출권", "€/t", 2), ("BRENT", "브렌트유", "$/bbl", 1)):
        k = K.get(key)
        if k:
            parts.append("%s %s%s(%s, 1주 전 대비 %s)" % (nm, f0(k["value"], nd), unit, md(k["date"]),
                                                       sg(k["chg_pct"], 1, "%")))
    for key, nm in (("TTF", "가스 TTF"), ("COAL", "석탄 뉴캐슬")):
        k = K.get(key)
        if k:
            parts.append("%s %d월 평균 %s %s(전월 대비 %s)" % (nm, int(k["period"][5:7]), f0(k["value"], 1),
                                                         k["unit"], sg(k["chg_pct"], 0, "%")))
    return " · ".join(parts) if parts else "원자재 자료 수집 실패"


def s_outlook(R, FC, W, week_end, fxp):
    wx = FC.get("wx", {})
    t = dt.datetime.now(TZ).date().isoformat()
    nd = sorted(d for d in wx.get("wind100", {}) if d >= t)[:7]
    parts, info = [], {"days": nd}
    if len(nd) >= 3 and W["wind100"] is not None:
        wf = sum(wx["wind100"][d] for d in nd) / len(nd)
        p = pct(wf, W["wind100"])
        info["wind_fc"], info["wind_pct"] = wf, p
        if p is not None and p >= FLAT:
            parts.append("풍속 예보(%s~%s) 이번 주 대비 %s → 가격 하방 요인" % (md(nd[0]), md(nd[-1]), sg(p, 0, "%")))
        elif p is not None and p <= -FLAT:
            parts.append("풍속 예보(%s~%s) 이번 주 대비 %s → 가격 상방 요인" % (md(nd[0]), md(nd[-1]), sg(p, 0, "%")))
        else:
            parts.append("풍속 예보(%s~%s) 이번 주와 비슷" % (md(nd[0]), md(nd[-1])))
    tl = [wx.get("temp", {}).get(d) for d in nd if wx.get("temp", {}).get(d) is not None]
    if len(tl) >= 3 and W["temp"] is not None:
        tf = sum(tl) / len(tl)
        dtp = tf - W["temp"]
        info["temp_fc"], info["temp_diff"] = tf, dtp
        if dtp <= -2:
            parts.append("기온 예보 평균 %.1f°C(%s°C) → 수요 증가, 상방 요인" % (tf, sg(dtp, 1)))
        elif dtp >= 2:
            parts.append("기온 예보 평균 %.1f°C(%s°C) → 수요 감소, 하방 요인" % (tf, sg(dtp, 1)))
        else:
            parts.append("기온 예보 평균 %.1f°C, 이번 주와 비슷" % tf)
    known = [d for d in sorted(R) if d > week_end.isoformat() and R[d].get("px_PL") is not None][:3]
    if known:
        parts.append("확정 하루전 가격 " + ", ".join(
            "%s %s" % (mdw(d), f0(R[d]["px_PL"] * fxp.on(d) if fxp.on(d) else None)) for d in known) + " PLN/MWh")
    return (" · ".join(parts) if parts else "예보 자료 없음"), info


# ══════════════════════════════════════════════════════════════════
#  PPT
# ══════════════════════════════════════════════════════════════════
def rgb(h):
    return RGBColor.from_string(h)


def style(font, size, bold=False, color=INK):
    font.size = Pt(size)
    font.bold = bold
    font.color.rgb = rgb(color)
    font.name = LAT
    rpr = font._element
    ea = rpr.find(qn("a:ea"))
    if ea is None:
        ea = rpr.makeelement(qn("a:ea"), {})
        lat = rpr.find(qn("a:latin"))
        if lat is not None:
            lat.addnext(ea)
        else:
            rpr.append(ea)
    ea.set("typeface", KO)


def fill_tf(tf, lines, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, after=0):
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.08)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    tf.vertical_anchor = anchor
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if after:
            p.space_after = Pt(after)
        for text, size, bold, color in line:
            r = p.add_run()
            r.text = text
            style(r.font, size, bold, color)


def textbox(sl, x, y, w, h, lines, **kw):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    fill_tf(tb.text_frame, lines, **kw)
    return tb


def rect(sl, x, y, w, h, color, rounded=False):
    s = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
                            Inches(x), Inches(y), Inches(w), Inches(h))
    if rounded:
        s.adjustments[0] = 0.06
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(color)
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def kpi(sl, x, y, w, h, color, label, value, unit, sub1, sub2, src):
    c = rect(sl, x, y, w, h, CARD, rounded=True)
    rect(sl, x, y + 0.12, 0.05, h - 0.24, color)
    fill_tf(c.text_frame, [[(label, 9, True, INK2)],
                           [(value, 18, True, INK), ("  " + unit, 8.5, False, INK2)],
                           [(sub1, 8, False, INK2)], [(sub2, 8, False, INK2)],
                           [("자료: " + src, 7, False, MUTED)]])
    c.text_frame.margin_left = Inches(0.16)


def nice_range(vals, n=4):
    lo, hi = min(vals), max(vals)
    if hi == lo:
        hi = lo + 1
    pad = (hi - lo) * 0.08
    floor0 = lo >= 0                      # 값이 모두 양수면 축을 0 아래로 내리지 않는다
    lo, hi = lo - pad, hi + pad
    if floor0:
        lo = max(lo, 0)
    raw = (hi - lo) / n
    mag = 10 ** math.floor(math.log10(raw))
    step = next(m * mag for m in (1, 2, 2.5, 5, 10) if m * mag >= raw)
    return math.floor(lo / step) * step, math.ceil(hi / step) * step, step


def step_fmt(step):
    """눈금 간격이 정수가 아니면(0.5·2.5 등) 소수 1자리 — 같은 숫자가 두 번 찍히는 것 방지."""
    return "#,##0" if abs(step - round(step)) < 1e-9 else "#,##0.0"


def set_skip(ax, n):
    el = ax._element
    for tag in ("c:tickLblSkip", "c:tickMarkSkip"):
        for e in el.findall(qn(tag)):
            el.remove(e)
    a = el.makeelement(qn("c:tickLblSkip"), {"val": str(n)})
    b = el.makeelement(qn("c:tickMarkSkip"), {"val": str(n)})
    nm = el.find(qn("c:noMultiLvlLbl"))
    if nm is not None:
        nm.addprevious(a)
        nm.addprevious(b)
    else:
        el.append(a)
        el.append(b)


def line_chart(sl, x, y, w, h, cats, series, title, nd=0, legend=False, skip=1, tsize=9.5):
    """series = [(이름, 값목록, 색)] — 한 축만 쓴다(단위가 다른 값은 별도 차트)."""
    vals = [v for _, vs, _ in series for v in vs if v is not None]
    if not cats or not vals:
        c = rect(sl, x, y, w, h, CARD, rounded=True)
        fill_tf(c.text_frame, [[(title, 9, True, INK2)], [("자료 없음 (수집 실패, Actions 로그 확인)", 8.5, False, MUTED)]])
        return
    cd = CategoryChartData()
    cd.categories = cats
    for name, vs, _ in series:
        cd.add_series(name, [None if v is None else round(v, nd) for v in vs])
    ch = sl.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    style(ch.font, 8, False, MUTED)                # 영문·숫자 Arial + 한글 LG스마트체
    ch.has_title = True
    ch.chart_title.text_frame.text = title
    for p in ch.chart_title.text_frame.paragraphs:
        for r in p.runs:
            style(r.font, tsize, True, INK2)
    ch.has_legend = legend
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.TOP
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(8.5)
    lo, hi, step = nice_range(vals)
    va = ch.value_axis
    va.minimum_scale, va.maximum_scale, va.major_unit = lo, hi, step
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = rgb(GRID)
    va.major_gridlines.format.line.width = Pt(0.5)
    va.format.line.fill.background()
    va.major_tick_mark = XL_TICK_MARK.NONE
    va.tick_labels.number_format = step_fmt(step)
    va.tick_labels.number_format_is_linked = False
    va.tick_labels.font.size = Pt(8)
    va.tick_labels.font.color.rgb = rgb(MUTED)
    ca = ch.category_axis
    ca.format.line.color.rgb = rgb(AXIS)
    ca.major_tick_mark = XL_TICK_MARK.NONE
    ca.tick_labels.font.size = Pt(8)
    ca.tick_labels.font.color.rgb = rgb(MUTED)
    if skip > 1:
        set_skip(ca, skip)
    plot = ch.plots[0]
    for i, (name, vs, col) in enumerate(series):
        s = plot.series[i]
        s.smooth = False
        s.marker.style = XL_MARKER_STYLE.NONE
        s.format.line.color.rgb = rgb(col)
        s.format.line.width = Pt(1.75)
        idx = [j for j, v in enumerate(vs) if v is not None]
        if idx:                                    # 마지막 값에만 숫자 (모든 점에 붙이지 않는다)
            dl = s.points[idx[-1]].data_label
            dl.has_text_frame = True
            dl.text_frame.text = f0(vs[idx[-1]], nd)
            for p in dl.text_frame.paragraphs:
                for r in p.runs:
                    style(r.font, 8.5, True, INK2)
            dl.position = XL_LABEL_POSITION.RIGHT


def table(sl, x, y, w, h, rows, widths):
    t = sl.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w), Inches(h)).table
    for j, cw in enumerate(widths):
        t.columns[j].width = Inches(cw)
    for i, row in enumerate(rows):
        t.rows[i].height = Inches(h / len(rows))
        for j, val in enumerate(row):
            c = t.cell(i, j)
            c.fill.solid()
            c.fill.fore_color.rgb = rgb("E8E7E1" if i == 0 else ("FCFCFB" if i % 2 else CARD))
            c.margin_left = c.margin_right = Inches(0.05)
            c.margin_top = c.margin_bottom = Inches(0.01)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = c.text_frame
            tf.text = str(val)
            for p in tf.paragraphs:
                p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT
                for r in p.runs:
                    style(r.font, 8, i == 0, INK if i else INK2)


def build_ppt(path, ctx):
    W, P, K, ws, we, tag = ctx["W"], ctx["P"], ctx["K"], ctx["ws"], ctx["we"], ctx["tag"]
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    # ── 제목
    textbox(sl, 0.35, 0.22, 8.6, 0.55, [[("폴란드 전력단가 및 원자재 가격동향", 22, True, INK)]],
            anchor=MSO_ANCHOR.MIDDLE)
    textbox(sl, 8.9, 0.22, 4.1, 0.55,
            [[("%d년 %d주차 (%s~%s)" % (ctx["iso"][0], ctx["iso"][1], md(ws.isoformat()), md(we.isoformat())), 12, True, INK2)],
             [("본사 에너지효율화팀 · 발행 %s · 폴란드 현지 월~일 기준" % ctx["issued"], 8.5, False, MUTED)]],
            anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
    rect(sl, 0.4, 0.86, 12.53, 0.02, AXIS)

    # ── KPI 6장
    cw, gap, y0, h0 = 1.99, 0.118, 0.98, 1.14

    def kx(i):
        return 0.4 + i * (cw + gap)
    p_pl = pct(W["px_PL_pln"], P["px_PL_pln"])
    kpi(sl, kx(0), y0, cw, h0, COL["PL"], "폴란드 전력 (주평균)", f0(W["px_PL_pln"]), "PLN/MWh",
        "%s EUR/MWh · %s 원/kWh" % (f0(W["px_PL"]), f0(W["px_PL_krw"], 1)), "전주 대비 %s" % sg(p_pl, 1, "%"),
        SRC["px"] + " · 환율 " + SRC["fx"])
    kpi(sl, kx(1), y0, cw, h0, COL["DE"], "독일 전력 (주평균)", f0(W["px_DE"]), "EUR/MWh",
        "%s PLN/MWh" % f0(W["px_DE_pln"]), "전주 대비 %s" % sg(pct(W["px_DE"], P["px_DE"]), 1, "%"), SRC["px"])
    k = K.get("EUA")
    kpi(sl, kx(2), y0, cw, h0, COL["EUA"], "탄소배출권 EUA", f0(k["value"], 2) if k else "없음", "€/t",
        ("%s 경매" % mdw(k["date"])) if k else "자료 없음",
        ("1주 전 대비 %s" % sg(k["chg_pct"], 1, "%")) if k else "", SRC["eua"])
    for i, (key, label) in enumerate((("TTF", "가스 TTF (월평균)"), ("COAL", "석탄 뉴캐슬 (월평균)"))):
        k = K.get(key)
        kpi(sl, kx(3 + i), y0, cw, h0, COL[key], label, f0(k["value"], 1) if k else "없음",
            k["unit"] if k else "", ("%s년 %d월" % (k["period"][:4], int(k["period"][5:7]))) if k else "자료 없음",
            ("전월 대비 %s" % sg(k["chg_pct"], 1, "%")) if k else "", SRC["wb"])
    k = K.get("BRENT")
    kpi(sl, kx(5), y0, cw, h0, COL["BRENT"], "브렌트유", f0(k["value"], 1) if k else "없음", "$/bbl",
        ("%s 기준" % mdw(k["date"])) if k else "자료 없음",
        ("1주 전 대비 %s" % sg(k["chg_pct"], 1, "%")) if k else "", SRC["brent"])

    # ── 차트: 전력 (왼쪽 큰 것)
    cs = ctx["chart_days"]
    line_chart(sl, 0.4, 2.15, 7.3, 2.85, [md(d) for d in cs],
               [("폴란드", [ctx["R"].get(d, {}).get("px_PL_pln") for d in cs], COL["PL"]),
                ("독일", [ctx["R"].get(d, {}).get("px_DE_pln") for d in cs], COL["DE"])],
               "하루전 전력가격 일평균 (PLN/MWh, 최근 12주) · 자료: %s, 환율 %s" % (SRC["px"], SRC["fx"]),
               legend=True, skip=7)

    # ── 차트: 원자재 2×2 (각자 단위 → 각자 축)
    mx, my, mw, mh = 7.85, 2.15, 2.5, 1.4
    for i, (key, title, nd) in enumerate((("EUA", "탄소배출권 EUA (€/t) · %s" % SRC["eua"], 2),
                                          ("BRENT", "브렌트유 ($/bbl) · %s" % SRC["brent"], 1))):
        src = ctx["daily_src"][key]
        ks = [d for d in sorted(src) if cs[0] <= d <= we.isoformat()]
        line_chart(sl, mx + i * (mw + 0.08), my, mw, mh, [md(d) for d in ks],
                   [(title, [src[d] for d in ks], COL[key])], title, nd=nd, skip=max(1, len(ks) // 4), tsize=8.5)
    for i, (key, wbk, title) in enumerate((("TTF", "WB_TTF_EU", "가스 TTF (%s, 월평균) · " + SRC["wb"]),
                                           ("COAL", "WB_COAL_AU", "석탄 뉴캐슬 (%s, 월평균) · " + SRC["wb"]))):
        s = ctx["MON"].get(wbk, {})
        pts = s.get("points", [])[-CHART_MONTHS:]
        t = title % UNIT_FIX.get(wbk, s.get("unit", ""))
        line_chart(sl, mx + i * (mw + 0.08), my + mh + 0.05, mw, mh, [p["period"][2:].replace("-", ".") for p in pts],
                   [(t, [p["value"] for p in pts], COL[key])], t, nd=1, skip=3, tsize=8.5)

    # ── 이번 주 포인트
    rect(sl, 0.4, 5.12, 8.15, 1.86, CARD, rounded=True)
    lines = [[("이번 주 포인트", 10.5, True, INK)]]
    for lab, txt in ctx["bullets"]:
        if txt:
            lines.append([("▪ %s  " % lab, 8.5, True, INK), (txt, 8.5, False, INK2),
                          ("  [" + BULLET_SRC.get(lab, "") + "]", 7.5, False, MUTED)])
    textbox(sl, 0.48, 5.16, 8.0, 1.8, lines, after=2)

    # ── 수급 표
    rows = [["수급·기상 (일평균)", "이번 주", "전주", "차이", "자료"]]
    for key, nm, nd, unit, src in (("load", "수요", 1, "GW", "ENTSO-E"), ("wind", "풍력 발전", 1, "GW", "ENTSO-E"),
                                   ("solar", "태양광 발전", 1, "GW", "ENTSO-E"),
                                   ("res", "잔여부하(화력 몫)", 1, "GW", "ENTSO-E 계산"),
                                   ("coal", "석탄·갈탄 발전", 1, "GW", "ENTSO-E"), ("gas", "가스 발전", 1, "GW", "ENTSO-E"),
                                   ("temp", "기온", 1, "°C", "Open-Meteo"), ("wind100", "풍속(100m)", 1, "km/h", "Open-Meteo")):
        a, b = W.get(key), P.get(key)
        rows.append(["%s (%s)" % (nm, unit), f0(a, nd), f0(b, nd),
                     sg(a - b, nd) if (a is not None and b is not None) else "없음", src])
    table(sl, 8.7, 5.12, 4.23, 1.86, rows, [1.45, 0.62, 0.62, 0.6, 0.94])

    # ── 출처 (원 자료명까지)
    textbox(sl, 0.35, 7.0, 12.6, 0.42,
            [[("출처  ", 7, True, INK2),
              ("전력가격·수요·발전: ENTSO-E Transparency Platform (A44 Day-ahead Prices, A65 Actual Total Load, "
               "A75 Actual Generation per Type) · 탄소배출권: EEX EU ETS Primary Auction Report · 브렌트유: U.S. EIA "
               "Europe Brent Spot Price FOB (RBRTE) · 가스·석탄: World Bank Commodity Price Data (Pink Sheet) · "
               "환율: ECB Euro foreign exchange reference rates · 기상: Open-Meteo (ERA5·예보 모델)", 7, False, MUTED)],
             [("자료 기준 %s 수집 · 잔여부하 = 수요 − 풍력 − 태양광 · 문장은 숫자에서 규칙으로 자동 작성 · "
               "가스·석탄은 무료 일별 자료가 없어 월평균" % ctx["collected"], 7, False, MUTED)]])

    # ── 2~4장: 장마다 따로 — 실패한 장은 지우고 로그만 남긴다 (1장은 항상 나온다)
    for fn in (slide_composition, slide_cause, slide_outlook):
        n0 = len(prs.slides)
        try:
            fn(prs, ctx)
        except Exception as e:
            print("!! %s 실패 → 이 장은 빼고 진행: %r" % (fn.__name__, e), flush=True)
            traceback.print_exc()
            lst = prs.slides._sldIdLst
            while len(prs.slides) > n0:
                sid = lst[-1]
                prs.part.drop_rel(sid.rId)
                lst.remove(sid)
    prs.save(path)
    try:
        patch_theme(path)
    except Exception as e:
        print("!! 테마 글꼴 정리 실패(PPT 는 정상): %r" % e, flush=True)
    print("PPT %d장" % len(prs.slides), flush=True)


# ══════════════════════════════════════════════════════════════════
#  2~4장 — 가격 구성 · 상승 원인 · 다음 주 전망 (HTML ②③④ 탭과 같은 계산)
# ══════════════════════════════════════════════════════════════════
COMP_MONTHS = 12
GAS_EFF, COAL_EFF, GAS_EF, COAL_EF, VOM_EUR, COAL_MWH_T, MMBTU_MWH = 0.55, 0.41, 0.37, 0.83, 2.0, 6.98, 0.293071
STACKS = [("가스 연료비", "gF", "2A78D6"), ("가스 탄소비", "gC", "86B6EF"), ("석탄 연료비", "cF", "008300"),
          ("석탄 탄소비", "cC", "1BAF7A"), ("운전비", "V", "C3C2B7"), ("희소성 프리미엄", "prem", "EB6834"),
          ("재생에너지 과잉 할인", "disc", "6250D6")]
DARK_FILL = {"2A78D6", "008300", "6250D6"}          # 이 색 위의 숫자는 흰색


def month_avg(src, m):
    xs = [float(v) for k, v in (src or {}).items() if k.startswith(m) and v is not None]
    return sum(xs) / len(xs) if xs else None


def composition(R, M, upto):
    """월별 전력가격 = 가스·석탄 발전원가 가중합 + 희소성 프리미엄(또는 재생에너지 과잉 할인).
    전력 = ENTSO-E 하루전 월평균 × ECB, 가스 = World Bank TTF 월평균(열량 환산), 석탄 = World Bank 뉴캐슬.
    그 달 가스·석탄 월평균이 아직 없으면 직전 달 값(note 에 표시)."""
    by = {}
    for d, r in R.items():
        if d <= upto and r.get("px_PL") is not None:
            by.setdefault(d[:7], []).append(r["px_PL"])
    mon, fx, eua = M.get("monthly", {}), M.get("fx", {}), M.get("eua", {})

    def ser(k):
        return {p["period"]: float(p["value"]) for p in mon.get(k, {}).get("points", []) if p.get("value") is not None}

    def last_le(o, m):
        ks = [k for k in o if k <= m]
        return (o[max(ks)], max(ks)) if ks else (None, None)
    ttf, coal = ser("WB_TTF_EU"), ser("WB_COAL_AU")
    out = []
    for m in sorted(by)[-COMP_MONTHS:]:
        if len(by[m]) < 5:
            continue
        pln, usd, e = month_avg(fx.get("PLN"), m), month_avg(fx.get("USD"), m), month_avg(eua, m)
        t, tk = last_le(ttf, m)
        c, ck = last_le(coal, m)
        if not pln or not usd or e is None or t is None or c is None:
            continue
        P = sum(by[m]) / len(by[m]) * pln
        gas_mwh = t / usd / MMBTU_MWH * pln                     # $/MMBtu → PLN/MWh(열량)
        gF, gC, V = gas_mwh / GAS_EFF, GAS_EF * e * pln, VOM_EUR * pln
        cF, cC = c * (pln / usd) / COAL_MWH_T / COAL_EFF, COAL_EF * e * pln
        G, Cc = gF + gC + V, cF + cC + V
        w = (P - Cc) / (G - Cc) if G != Cc else 0.0
        w = max(0.0, min(1.0, w))
        base = w * G + (1 - w) * Cc
        out.append({"m": m, "days": len(by[m]), "P": P, "G": G, "C": Cc, "w": w,
                    "gF": w * gF, "gC": w * gC, "cF": (1 - w) * cF, "cC": (1 - w) * cC, "V": V,
                    "prem": P - base if P > base else 0.0, "disc": P - base if P < base else 0.0,
                    "coal_carb_share": cC / Cc if Cc else None, "gas_mwh": gas_mwh, "eua": e, "coal_usd": c,
                    "ttf_usd": t, "pln": pln, "note": "" if (tk == m and ck == m) else "가스 %s·석탄 %s 월평균 사용" % (tk, ck)})
    return out


def gen_share(R, ref):
    od = [R[d] for d in sorted(R) if d.startswith(ref["m"])
          and all(R[d].get(k) is not None for k in ("load", "coal", "gas", "wind", "solar"))]
    if len(od) < 10:
        return None
    g = {k: sum(r[k] for r in od) / len(od) for k in ("load", "coal", "gas", "wind", "solar")}
    other = max(0.0, g["load"] - g["coal"] - g["gas"] - g["wind"] - g["solar"])
    parts = [("석탄·갈탄", g["coal"], 100 - ref["w"] * 100), ("가스", g["gas"], ref["w"] * 100),
             ("풍력", g["wind"], 0.0), ("태양광", g["solar"], 0.0), ("기타 (수입·수력 등)", other, 0.0)]
    return {"days": len(od), "load": g["load"], "parts": parts}


def mon_label(c):
    s = "%d월" % int(c["m"][5:7])
    if c["days"] < 25:
        s += "(%d일)" % c["days"]
    return s + ("*" if c["note"] else "")


def wk_label(d0, d1):
    iso = d0.isocalendar()
    return "%d년 %d주차 (%s~%s)" % (iso[0], iso[1], md(d0.isoformat()), md(d1.isoformat()))


def header(prs, ctx, title, sub=None):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    textbox(sl, 0.35, 0.22, 8.6, 0.55, [[(title, 22, True, INK)]], anchor=MSO_ANCHOR.MIDDLE)
    textbox(sl, 8.9, 0.22, 4.1, 0.55,
            [[(sub or wk_label(ctx["ws"], ctx["we"]), 12, True, INK2)],
             [("본사 에너지효율화팀 · 발행 %s" % ctx["issued"], 8.5, False, MUTED)]],
            anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
    rect(sl, 0.4, 0.86, 12.53, 0.02, AXIS)
    return sl


def footer(sl, lines):
    textbox(sl, 0.35, 6.98, 12.6, 0.45, [[(t, 7, i == 0, INK2 if i == 0 else MUTED)] for i, t in enumerate(lines)])


def note_box(sl, x, y, w, h, title, items, size=9):
    rect(sl, x, y, w, h, CARD, rounded=True)
    lines = [[(title, 10.5, True, INK)]]
    for it in items:
        txt, bold = (it, False) if isinstance(it, str) else it
        lines.append([("▪ " + txt, size, bold, INK if bold else INK2)])
    textbox(sl, x + 0.08, y + 0.06, w - 0.16, h - 0.12, lines, after=3)


def placeholder(sl, x, y, w, h, title, why):
    c = rect(sl, x, y, w, h, CARD, rounded=True)
    fill_tf(c.text_frame, [[(title, 9.5, True, INK2)], [(why, 8.5, False, MUTED)]])


def chart_frame(ch, title, legend=None, tsize=10.5):
    style(ch.font, 8, False, MUTED)
    ch.has_title = True
    ch.chart_title.text_frame.text = title
    for p in ch.chart_title.text_frame.paragraphs:
        for r in p.runs:
            style(r.font, tsize, True, INK2)
    ch.has_legend = legend is not None
    if legend is not None:
        ch.legend.position = legend
        ch.legend.include_in_layout = False
        style(ch.legend.font, 8, False, INK2)


def axes_style(ch, numfmt="#,##0"):
    va = ch.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = rgb(GRID)
    va.major_gridlines.format.line.width = Pt(0.5)
    va.format.line.fill.background()
    va.major_tick_mark = XL_TICK_MARK.NONE
    va.tick_labels.number_format = numfmt
    va.tick_labels.number_format_is_linked = False
    va.tick_labels.font.size = Pt(8)
    va.tick_labels.font.color.rgb = rgb(MUTED)
    ca = ch.category_axis
    ca.format.line.color.rgb = rgb(AXIS)
    ca.major_tick_mark = XL_TICK_MARK.NONE
    ca.tick_labels.font.size = Pt(8)
    ca.tick_labels.font.color.rgb = rgb(INK2)
    return va, ca


def reverse_cat(ax):
    """가로 막대에서 첫 항목이 맨 위에 오도록."""
    sc = ax._element.find(qn("c:scaling"))
    o = sc.find(qn("c:orientation"))
    if o is None:
        o = sc.makeelement(qn("c:orientation"), {})
        lb = sc.find(qn("c:logBase"))
        if lb is not None:
            lb.addnext(o)
        else:
            sc.insert(0, o)
    o.set("val", "maxMin")


def point_label(series, j, text, pos, color=INK2, size=8.5, bold=True):
    dl = series.points[j].data_label
    dl.has_text_frame = True
    dl.text_frame.text = text
    for p in dl.text_frame.paragraphs:
        for r in p.runs:
            style(r.font, size, bold, color)
    dl.position = pos


def val_el(tag, val):
    """<c:xxx val="..."/> — OxmlElement 의 두 번째 인자는 속성이 아니라 이름공간이라 따로 set 한다."""
    e = OxmlElement(tag)
    e.set("val", val)
    return e


def last_bar_to_line(ch):
    """누적 막대 차트의 마지막 계열을 꺾은선으로 옮긴다(막대 + 선 혼합 차트). 같은 축을 쓴다."""
    pa = ch._chartSpace.chart.plotArea
    bar = pa.find(qn("c:barChart"))
    s = bar.findall(qn("c:ser"))[-1]
    bar.remove(s)
    for t in ("c:invertIfNegative", "c:shape", "c:pictureOptions"):
        for e in s.findall(qn(t)):
            s.remove(e)
    lc = OxmlElement("c:lineChart")
    for tag, val in (("c:grouping", "standard"), ("c:varyColors", "0")):
        lc.append(val_el(tag, val))
    lc.append(s)
    lc.append(val_el("c:marker", "1"))
    for a in bar.findall(qn("c:axId")):
        lc.append(deepcopy(a))
    bar.addnext(lc)
    return ch.plots[1].series[0]


def signed_bars(sl, x, y, w, h, cats, vals, colors, title):
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("값", [round(v, 1) for v in vals])
    ch = sl.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    chart_frame(ch, title)
    va, ca = axes_style(ch)
    reverse_cat(ca)
    ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW          # 음수 막대와 항목 이름이 겹치지 않게
    plot = ch.plots[0]
    plot.gap_width = 60
    s = plot.series[0]
    s.invert_if_negative = False
    big = max(abs(v) for v in vals) or 1
    for j, v in enumerate(vals):
        pt = s.points[j]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb(colors[j])
        dpt = s._element.get_or_add_dPt_for_point(j)           # 점마다 '음수 반전 안 함' — 없으면 음수 막대가 흰색으로 나온다
        if dpt.find(qn("c:invertIfNegative")) is None:
            dpt.find(qn("c:idx")).addnext(val_el("c:invertIfNegative", "0"))
        inside = v < 0 and abs(v) >= 0.12 * big
        point_label(s, j, sg(v), XL_LABEL_POSITION.INSIDE_END if inside else XL_LABEL_POSITION.OUTSIDE_END,
                    ("FFFFFF" if colors[j] in DARK_FILL else INK) if inside else INK2)


# ── 2장 가격 구성
def slide_composition(prs, ctx):
    C = ctx.get("COMP") or []
    sl = header(prs, ctx, "폴란드 전력가격 구성")
    if len(C) < 2:
        placeholder(sl, 0.35, 0.98, 12.6, 5.9, "월별 전력가격 구성", "자료 부족. 월평균 가스·석탄·배출권·환율 중 수집 실패 (Actions 로그 확인)")
        return
    st = [s for s in STACKS if any(abs(c[s[1]]) >= 0.5 for c in C)]

    # A. 월별 구성 (누적 막대) + 실제 가격 (검은 선)
    cd = CategoryChartData()
    cd.categories = [mon_label(c) for c in C]
    for nm, k, _ in st:
        cd.add_series(nm, [round(c[k]) for c in C])
    cd.add_series("실제 가격", [round(c["P"]) for c in C])
    ch = sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(0.35), Inches(0.98), Inches(8.3), Inches(3.75), cd).chart
    chart_frame(ch, "월별 전력가격 구성 (PLN/MWh), 막대는 원가 항목 · 검은 선은 실제 가격", XL_LEGEND_POSITION.BOTTOM)
    _, ca = axes_style(ch)
    ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW          # 월 이름을 음수 막대 아래로
    plot = ch.plots[0]
    plot.gap_width = 60
    for i, (_, _, col) in enumerate(st):
        s = plot.series[i]
        s.invert_if_negative = False
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = rgb(col)
        s.format.line.color.rgb = rgb("FFFFFF")
        s.format.line.width = Pt(0.5)
    ln = last_bar_to_line(ch)
    ln.smooth = False
    ln.format.line.color.rgb = rgb(INK)
    ln.format.line.width = Pt(2)
    ln.marker.style = XL_MARKER_STYLE.CIRCLE
    ln.marker.size = 6
    ln.marker.format.fill.solid()
    ln.marker.format.fill.fore_color.rgb = rgb(INK)
    ln.marker.format.line.color.rgb = rgb(INK)
    point_label(ln, len(C) - 1, f0(C[-1]["P"]), XL_LABEL_POSITION.ABOVE, INK)

    full = [c for c in C if c["days"] >= 25]
    ref, cur = (full[-1] if full else C[-1]), C[-1]

    # E. 발전량 비중 vs 가격을 정한 비중 (기준월)
    gs = gen_share(ctx["R"], ref)
    if gs:
        cd = CategoryChartData()
        cd.categories = [p[0] for p in gs["parts"]]
        cd.add_series("발전량 비중 (%)", [round(p[1] / gs["load"] * 100) for p in gs["parts"]])
        cd.add_series("가격에 실린 무게 (%, 환산)", [round(p[2]) for p in gs["parts"]])
        ch = sl.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(8.78), Inches(0.98), Inches(4.2), Inches(3.75), cd).chart
        chart_frame(ch, "석탄은 많이 돌지만, 가격은 가스가 정한다 (%d월)" % int(ref["m"][5:7]), XL_LEGEND_POSITION.BOTTOM)
        va, ca = axes_style(ch, '0"%"')
        va.minimum_scale, va.maximum_scale, va.major_unit = 0, 100, 25
        reverse_cat(ca)
        plot = ch.plots[0]
        plot.gap_width = 60
        plot.overlap = -10
        for i, col in enumerate((AXIS, "EB6834")):
            plot.series[i].format.fill.solid()
            plot.series[i].format.fill.fore_color.rgb = rgb(col)
        plot.has_data_labels = True
        dls = plot.data_labels
        dls.number_format = '0"%";;;'                     # 0% 는 숨김
        dls.number_format_is_linked = False
        dls.position = XL_LABEL_POSITION.OUTSIDE_END
        style(dls.font, 8, True, INK2)
    else:
        placeholder(sl, 8.78, 0.98, 4.2, 3.75, "발전량 비중과 가격에 실린 무게", "전원별 발전 자료 부족")

    # B. 기준월 vs 이번 달 (막대 길이 = 실제 가격)
    cmp_ = [ref, cur] if cur["m"] != ref["m"] else [ref]
    st2 = [s for s in STACKS if any(abs(c[s[1]]) >= 0.5 for c in cmp_)]
    cd = CategoryChartData()
    cd.categories = ["%s  %s" % (mon_label(c), f0(c["P"])) for c in cmp_]
    for nm, k, _ in st2:
        cd.add_series(nm, [round(c[k]) for c in cmp_])
    ch = sl.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, Inches(0.35), Inches(4.83), Inches(6.3), Inches(2.1), cd).chart
    chart_frame(ch, "전력가격 구성 비교 (PLN/MWh), 막대 길이가 실제 가격", XL_LEGEND_POSITION.BOTTOM, tsize=10)
    va, ca = axes_style(ch)
    reverse_cat(ca)
    plot = ch.plots[0]
    plot.gap_width = 50
    for i, (_, k, col) in enumerate(st2):
        s = plot.series[i]
        s.invert_if_negative = False
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = rgb(col)
        s.format.line.color.rgb = rgb("FFFFFF")
        s.format.line.width = Pt(0.5)
        for j, c in enumerate(cmp_):
            if abs(c[k]) >= 0.07 * c["P"]:
                point_label(s, j, f0(c[k]), XL_LABEL_POSITION.CENTER, "FFFFFF" if col in DARK_FILL else INK, 8)

    # 구성 해석
    gas, coal = ref["gF"] + ref["gC"], ref["cF"] + ref["cC"]
    li = []
    if gs:
        pc = {p[0]: p[1] / gs["load"] * 100 for p in gs["parts"]}
        li.append("발전량은 석탄·갈탄 %s%%, 가스 %s%%지만, 평균 가격으로 환산하면 가스 쪽 무게가 %s%%입니다 (%d월)"
                  % (f0(pc["석탄·갈탄"]), f0(pc["가스"]), f0(ref["w"] * 100), int(ref["m"][5:7])))
    li.append("%d월 %s PLN/MWh = 가스 관련 %s%% (연료 %s%% + 탄소 %s%%) · 석탄 관련 %s%% (연료 %s%% + 탄소 %s%%)"
              % (int(ref["m"][5:7]), f0(ref["P"]), f0(gas / ref["P"] * 100), f0(ref["gF"] / ref["P"] * 100),
                 f0(ref["gC"] / ref["P"] * 100), f0(coal / ref["P"] * 100), f0(ref["cF"] / ref["P"] * 100),
                 f0(ref["cC"] / ref["P"] * 100)))
    if ref["coal_carb_share"]:
        li.append("석탄 발전원가의 약 %s%%는 탄소배출권 비용 → 배출권 가격이 오르면 석탄 원가가 가장 크게 오름"
                  % f0(ref["coal_carb_share"] * 100))
    if cur["m"] != ref["m"] and cur["prem"] > 0:
        li.append(("%s %s PLN/MWh는 가스 발전원가(%s)보다 %s 높음 → 발전기 부족에 따른 희소성 프리미엄 %s%%"
                   % (mon_label(cur), f0(cur["P"]), f0(cur["G"]), f0(cur["prem"]), f0(cur["prem"] / cur["P"] * 100)), True))
    elif cur["m"] != ref["m"] and cur["disc"] < 0:
        li.append(("%s %s PLN/MWh는 석탄 발전원가(%s)보다 %s 낮음 → 재생에너지 발전이 많아 화력 원가 아래로 거래"
                   % (mon_label(cur), f0(cur["P"]), f0(cur["C"]), f0(-cur["disc"])), True))
    note_box(sl, 6.78, 4.83, 6.2, 2.1, "구성 해석", li, size=8.5)
    notes = sorted({c["note"] for c in C if c["note"]})
    footer(sl, ["출처  전력: ENTSO-E 하루전 월평균 × ECB 환율 · 가스: World Bank TTF 월평균(열량 환산) · 석탄: World Bank 뉴캐슬 월평균 · "
                "탄소배출권: EEX 경매 월평균 · 발전량: ENTSO-E A75·A65",
                "발전원가 = 연료비 ÷ 효율(가스 %.2f·석탄 %.2f) + 배출계수(가스 %.2f·석탄 %.2f t/MWh) × 배출권가 + 운전비 %.0f€/MWh · "
                "가스 쪽 무게 = (실제 가격 − 석탄 원가) ÷ (가스 원가 − 석탄 원가), 실제 시간 비율은 아닌 환산값%s"
                % (GAS_EFF, COAL_EFF, GAS_EF, COAL_EF, VOM_EUR, (" · * " + ", ".join(notes)) if notes else "")])


# ── 3장 상승 원인
def slide_cause(prs, ctx):
    DG, dec, W, P, K = ctx.get("DG") or {}, ctx.get("dec"), ctx["W"], ctx["P"], ctx["K"]
    sl = header(prs, ctx, "폴란드 전력가격 상승 원인" if (dec is None or dec["dpx"] >= 0) else "폴란드 전력가격 변동 원인")
    v, cp = DG.get("verdict") or {}, DG.get("compare") or {}
    ft = v.get("fit_trend")
    if ft and v.get("rise") is not None:
        sp, bs = DG.get("spike") or ["", ""], DG.get("base") or ["", ""]
        signed_bars(sl, 0.35, 0.98, 6.2, 3.95, ["날씨·수요 (바람·햇빛·기온)", "완만한 추세 (기준 기간 흐름)", "그 밖 (가스값·주변국·발전소 정지 등)"],
                    [ft["explained_res"], ft["explained_trend"], ft["unexplained"]], [COL["PL"], "C3C2B7", "EB6834"],
                    "%s 이후 급등 %s EUR/MWh의 원인 (%s~%s 평균 대비)" % (md(sp[0]), sg(v["rise"]), md(bs[0]), md(bs[1])))
    else:
        placeholder(sl, 0.35, 0.98, 6.2, 3.95, "급등 원인", "진단 자료(pl_drivers_diag.json) 없음")
    if dec:
        e = dec["eff"]
        vals = [e["wind"], e["load"], e["solar"], dec["other"]]
        signed_bars(sl, 6.78, 0.98, 6.2, 3.95,
                    ["풍력 " + ("감소" if dec["d"]["wind"] < 0 else "증가"), "수요 " + ("증가" if dec["d"]["load"] > 0 else "감소"),
                     "태양광 " + ("증가" if dec["d"]["solar"] > 0 else "감소"), "그 밖 (연료·주변국·발전소)"],
                    vals, [("6250D6" if x < 0 else COL["PL"]) for x in vals[:3]] + ["C3C2B7"],
                    "이번 주 전주 대비 %s PLN/MWh의 원인" % sg(dec["dpx"]))
    else:
        placeholder(sl, 6.78, 0.98, 6.2, 3.95, "이번 주 원인", s_cause(dec, ctx.get("fit") or {}, W, P))

    ls = []
    if ft and v.get("rise"):
        w_, s_, t_, r_ = cp.get("wind", {}), cp.get("solar", {}), cp.get("temp", {}), cp.get("resload", {})
        ls.append("%s 이후 급등(%s EUR/MWh, %s)의 약 %s%%는 날씨·수요: 풍력 %s→%sGW(%s), 태양광 %s, 기온 %s→%s°C → 화력 몫 %s→%sGW"
                  % (md((DG.get("spike") or [""])[0]), sg(v["rise"]), sg(cp.get("px_PL", {}).get("chg_pct"), 0, "%"),
                     f0(ft["explained_res"] / v["rise"] * 100), f0((w_.get("base") or 0) / 1000, 1),
                     f0((w_.get("spike") or 0) / 1000, 1), sg(w_.get("chg_pct"), 0, "%"), sg(s_.get("chg_pct"), 0, "%"),
                     f0(t_.get("base"), 1), f0(t_.get("spike"), 1), f0((r_.get("base") or 0) / 1000, 1),
                     f0((r_.get("spike") or 0) / 1000, 1)))
    if cp.get("px_DE"):
        ls.append("주변국: 독일 %s, 체코 %s, 리투아니아 %s, 스웨덴 남부 %s → %s"
                  % (sg(cp["px_DE"].get("chg_pct"), 0, "%"), sg(cp.get("px_CZ", {}).get("chg_pct"), 0, "%"),
                     sg(cp.get("px_LT", {}).get("chg_pct"), 0, "%"), sg(cp.get("px_SE4", {}).get("chg_pct"), 0, "%"),
                     v.get("pl_vs_de") or ""))
    if dec:
        ls.append("이번 주 전주 대비 %s PLN/MWh: 풍력 %s, 수요 %s, 태양광 %s, 그 밖(연료·주변국·발전소) %s"
                  % (sg(dec["dpx"]), sg(dec["eff"]["wind"]), sg(dec["eff"]["load"]), sg(dec["eff"]["solar"]), sg(dec["other"])))
    if K.get("EUA") and K.get("BRENT"):
        a, b = K["EUA"]["chg_pct"], K["BRENT"]["chg_pct"]
        tail = " → 배출권·유가는 이번 주 상승의 원인이 아님" if (dec and dec["dpx"] > 0 and a is not None and b is not None
                                                  and a < FLAT and b < FLAT) else ""
        ls.append("연료비: 탄소배출권 %s, 브렌트유 %s (1주 전 대비)%s. 가스는 월평균 자료뿐이라 주간 비교에서 빠짐" % (sg(a, 1, "%"), sg(b, 1, "%"), tail))
    oc = outlook_line(ctx)
    if oc:
        ls.append((oc, True))
    note_box(sl, 0.35, 5.08, 12.6, 1.82, "핵심 정리", ls)
    footer(sl, ["출처  하루전 가격·수요·발전: ENTSO-E · 기상·예보: Open-Meteo · 탄소배출권: EEX · 브렌트유: EIA · 환율: ECB",
                "원인 분해 = 직전 일별 자료에서 '가격 ~ 화력 몫(수요 − 풍력 − 태양광) + 시간추세' 직선을 잡아 화력 몫 변화 × 1GW당 가격으로 계산"
                " · 발전소 정지 정보는 미반영('그 밖'에 포함) · 문장은 숫자에서 규칙으로 자동 작성"])


def outlook_line(ctx):
    o, W = ctx.get("out_info") or {}, ctx["W"]
    p = o.get("wind_pct")
    if p is None:
        return None
    wx = (ctx.get("FC") or {}).get("wx", {}).get("wind100", {})
    days = [d for d in o.get("days", []) if wx.get(d) is not None]
    top = max(days, key=lambda d: wx[d]) if days else None
    if p >= FLAT:
        s = "다음 주 풍속 예보 %s → 재생에너지 발전이 늘어 가격 안정 가능성 높음" % sg(p, 0, "%")
        if top:
            s += " (바람이 가장 강한 날 %s %s km/h)" % (mdw(top), f0(wx[top], 1))
    elif p <= -FLAT:
        s = "다음 주 풍속 예보 %s → 화력 몫이 늘어 가격 상승 압력" % sg(p, 0, "%")
    else:
        s = "다음 주 풍속은 이번 주와 비슷(%s) → 가격은 기온(수요)·연료비 영향이 클 것" % sg(p, 0, "%")
    return s


# ── 4장 다음 주 전망
def fc_chart(sl, x, y, w, h, days, vals, ref, color, title):
    cd = CategoryChartData()
    cd.categories = [mdw(d) for d in days]
    cd.add_series("예보", [None if v is None else round(v, 1) for v in vals])
    if ref is not None:
        cd.add_series("이번 주 평균", [round(ref, 1)] * len(days))
    ch = sl.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    chart_frame(ch, title, XL_LEGEND_POSITION.BOTTOM)
    lo, hi, step = nice_range([v for v in vals if v is not None] + ([ref] if ref is not None else []))
    va, ca = axes_style(ch, step_fmt(step))
    va.minimum_scale, va.maximum_scale, va.major_unit = lo, hi, step
    plot = ch.plots[0]
    s = plot.series[0]
    s.smooth = False
    s.format.line.color.rgb = rgb(color)
    s.format.line.width = Pt(2)
    s.marker.style = XL_MARKER_STYLE.CIRCLE
    s.marker.size = 7
    s.marker.format.fill.solid()
    s.marker.format.fill.fore_color.rgb = rgb(color)
    s.marker.format.line.color.rgb = rgb("FFFFFF")
    for j, v in enumerate(vals):
        if v is not None:
            point_label(s, j, f0(v, 1), XL_LABEL_POSITION.ABOVE, INK2, 8)
    if ref is not None:
        r = plot.series[1]
        r.smooth = False
        r.marker.style = XL_MARKER_STYLE.NONE
        r.format.line.color.rgb = rgb(MUTED)
        r.format.line.width = Pt(1.25)
        r.format.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        point_label(r, len(days) - 1, "이번 주 %s" % f0(ref, 1), XL_LABEL_POSITION.BELOW, MUTED, 8, False)


def slide_outlook(prs, ctx):
    W, we, o = ctx["W"], ctx["we"], ctx.get("out_info") or {}
    nws, nwe = we + dt.timedelta(days=1), we + dt.timedelta(days=7)
    sl = header(prs, ctx, "폴란드 전력가격 다음 주 전망", wk_label(nws, nwe))
    wx = (ctx.get("FC") or {}).get("wx", {})
    t0 = dt.datetime.now(TZ).date().isoformat()
    days = sorted(d for d in wx.get("wind100", {}) if d >= t0)[:8]
    if len(days) < 3:
        placeholder(sl, 0.35, 0.98, 12.6, 5.9, "다음 주 전망", "Open-Meteo 예보 자료 없음 (Actions 로그 확인)")
        return
    fc_chart(sl, 0.35, 0.98, 6.2, 3.85, days, [wx["wind100"].get(d) for d in days], W.get("wind100"), "1BAF7A",
             "풍속 예보 (km/h, 100m · 3지점 평균), 점선은 이번 주 평균")
    fc_chart(sl, 6.78, 0.98, 6.2, 3.85, days, [wx.get("temp", {}).get(d) for d in days], W.get("temp"), "EB6834",
             "기온 예보 (°C, 3지점 평균), 점선은 이번 주 평균")

    li = []
    p = o.get("wind_pct")
    if p is not None:
        s = "풍속 예보 평균 %s km/h로 이번 주(%s) 대비 %s, %s" % (
            f0(o["wind_fc"], 1), f0(W.get("wind100"), 1), sg(p, 0, "%"),
            "풍력 발전 증가, 가격 하방 요인" if p >= FLAT else ("화력 몫 증가, 가격 상방 요인" if p <= -FLAT else "영향 작음"))
        bw = ((ctx.get("DG") or {}).get("compare") or {}).get("wind100", {}).get("base")
        bs = (ctx.get("DG") or {}).get("base") or []
        if bw and len(bs) == 2:
            s += " (급등 전 %s~%s 평균 %s km/h)" % (md(bs[0]), md(bs[1]), f0(bw, 1))
        li.append(s)
        dd_ = [d for d in o.get("days", []) if wx.get("wind100", {}).get(d) is not None]
        if dd_:
            top = max(dd_, key=lambda d: wx["wind100"][d])
            li.append("바람이 가장 강한 날: %s %s km/h" % (mdw(top), f0(wx["wind100"][top], 1)))
    if o.get("temp_fc") is not None:
        dtp = o["temp_diff"]
        eff = ("수요 증가, 상방 요인" if dtp <= -2 else "수요 다소 증가" if dtp <= -1 else
               "수요 감소, 하방 요인" if dtp >= 2 else "수요 다소 감소" if dtp >= 1 else "수요 영향 작음")
        li.append("기온 예보 평균 %s°C (이번 주 대비 %s°C) → %s" % (f0(o["temp_fc"], 1), sg(dtp, 1), eff))
    R, fxp = ctx["R"], ctx["fxp"]
    known = [d for d in sorted(R) if d > we.isoformat() and R[d].get("px_PL") is not None and fxp.on(d)][:3]
    if known:
        kv = [(d, R[d]["px_PL"] * fxp.on(d)) for d in known]
        s = "확정된 하루전 가격: " + ", ".join("%s %s PLN/MWh" % (mdw(d), f0(x)) for d, x in kv)
        if W.get("px_PL_pln") and kv[0][1] > W["px_PL_pln"] * 1.05:
            s += " → 주초는 아직 이번 주 평균(%s)보다 높음" % f0(W["px_PL_pln"])
        li.append(s)
    oc = outlook_line(ctx)
    if oc:
        li.append(("종합: " + oc, True))
    note_box(sl, 0.35, 4.98, 12.6, 1.92, "전망", li)
    footer(sl, ["출처  기상 예보: Open-Meteo Forecast API (바르샤바·포모제·실롱스크 3지점 평균) · 하루전 가격: ENTSO-E · 환율: ECB",
                "예보는 매일 바뀜 · 판단 기준: 풍속·기온 예보를 이번 주 평균과 비교(풍속 ±5%·기온 ±1°C 이내는 영향 작음) · "
                "문장은 숫자에서 규칙으로 자동 작성"])


def patch_theme(path):
    """기본 테마의 Calibri·빈 한글 글꼴을 Arial·LG스마트체로 — 표·차트의 기본 글꼴까지 통일."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            b = zi.read(it.filename)
            if re.match(r"ppt/theme/theme\d+\.xml$", it.filename):
                x = b.decode("utf-8")
                x = x.replace('typeface="Calibri Light"', 'typeface="%s"' % LAT).replace('typeface="Calibri"', 'typeface="%s"' % LAT)
                x = x.replace('<a:ea typeface=""/>', '<a:ea typeface="%s"/>' % KO)
                x = re.sub(r'<a:font script="Hang" typeface="[^"]*"/>', '<a:font script="Hang" typeface="%s"/>' % KO, x)
                b = x.encode("utf-8")
            zo.writestr(it, b)
    os.replace(tmp, path)


# ══════════════════════════════════════════════════════════════════
#  Excel raw
# ══════════════════════════════════════════════════════════════════
HDR = PatternFill("solid", fgColor="E8E7E1")


def sheet(wb, title, headers, rows, fmts=None, widths=None, first=False):
    sh = wb.active if first else wb.create_sheet()
    sh.title = title
    sh.append(headers)
    for c in range(1, len(headers) + 1):
        cell = sh.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.fill = HDR
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in rows:
        sh.append(r)
    for ci, f in (fmts or {}).items():  # noqa — 열 서식
        for row in sh.iter_rows(min_row=2, min_col=ci, max_col=ci):
            for cell in row:
                cell.number_format = f
    for i, w in enumerate(widths or [], start=1):
        sh.column_dimensions[get_column_letter(i)].width = w
    sh.row_dimensions[1].height = 32
    sh.freeze_panes = "B2"
    return sh


def r2(v, n=2):
    return None if v is None else round(v, n)


def build_xlsx(path, ctx):
    R, W, P, F4, K, MON, FXD = ctx["R"], ctx["W"], ctx["P"], ctx["F4"], ctx["K"], ctx["MON"], ctx["FXD"]
    wb = Workbook()
    rows = []

    def add(nm, unit, a, b, basis, src, f4=None, nd=2):
        rows.append([nm, unit, r2(a, nd), r2(b, nd), r2(pct(a, b), 1), r2(f4, nd), basis, src])
    wk = "%s~%s 주평균 vs 전주" % (ctx["ws"].isoformat(), ctx["we"].isoformat())
    add("폴란드 전력 하루전", "PLN/MWh", W["px_PL_pln"], P["px_PL_pln"], wk, "ENTSO-E × ECB", F4["px_PL_pln"])
    add("폴란드 전력 하루전", "EUR/MWh", W["px_PL"], P["px_PL"], wk, "ENTSO-E", F4["px_PL"])
    add("폴란드 전력 하루전", "원/kWh", W["px_PL_krw"], P["px_PL_krw"], wk, "ENTSO-E × ECB", F4["px_PL_krw"])
    for c, nm in (("DE", "독일"), ("CZ", "체코"), ("SK", "슬로바키아"), ("LT", "리투아니아"), ("SE4", "스웨덴 남부")):
        add(nm + " 전력 하루전", "EUR/MWh", W["px_" + c], P["px_" + c], wk, "ENTSO-E", F4["px_" + c])
    for key, nm, unit in (("EUA", "탄소배출권 EUA 경매", "EUR/t"), ("BRENT", "브렌트유 현물", "USD/bbl")):
        k = K.get(key)
        if k:
            rows.append([nm, unit, r2(k["value"]), r2(k["prev"]), r2(k["chg_pct"], 1), None,
                         "%s vs %s" % (k["date"], k["prev_date"]), "EEX" if key == "EUA" else "EIA"])
    for key in ("WB_TTF_EU", "WB_COAL_AU", "WB_COAL_ZA", "WB_HH_US", "WB_LNG_JP"):
        k = last_month(MON, key)
        if k:
            rows.append([k["label"], k["unit"], r2(k["value"]), r2(k["prev"]), r2(k["chg_pct"], 1), None,
                         "%s vs %s 월평균" % (k["period"], k["prev_period"]), "World Bank"])
    for key, nm, unit in (("load", "수요", "GW"), ("wind", "풍력 발전", "GW"), ("solar", "태양광 발전", "GW"),
                          ("res", "잔여부하(수요−풍력−태양광)", "GW"), ("coal", "석탄·갈탄 발전", "GW"),
                          ("gas", "가스 발전", "GW"), ("temp", "기온(3지점)", "°C"), ("wind100", "풍속 100m(3지점)", "km/h"),
                          ("rad", "일사(3지점)", "kWh/m2·일")):
        add(nm, unit, W[key], P[key], wk, "ENTSO-E" if key not in ("temp", "wind100", "rad") else "Open-Meteo", F4[key])
    FL = ctx.get("FLOW") or {}
    FLW = {r["date"]: r for r in FL.get("rows", [])}

    def fsum(key, d0, d1):                     # 주 합계 (5일 미만이면 빈칸)
        xs = [FLW[d][key] for d in drange(d0, d1) if d in FLW and FLW[d].get(key) is not None]
        return sum(xs) if len(xs) >= 5 else None
    if FLW:
        pws, pwe = ctx["ws"] - dt.timedelta(days=7), ctx["we"] - dt.timedelta(days=7)
        for key, nm in (("total_net", "국경 순수입(수입−수출) 주 합계"), ("total_imp", "국경 수입 주 합계"),
                        ("total_exp", "국경 수출 주 합계")):
            add(nm, "GWh", fsum(key, ctx["ws"], ctx["we"]), fsum(key, pws, pwe),
                "%s~%s 합계 vs 전주" % (ctx["ws"].isoformat(), ctx["we"].isoformat()), "ENTSO-E A11")
    add("환율 EUR→PLN", "PLN", ctx["fx_w"]["PLN"], ctx["fx_p"]["PLN"], wk, "ECB")
    add("환율 EUR→KRW", "KRW", ctx["fx_w"]["KRW"], ctx["fx_p"]["KRW"], wk, "ECB")
    rows.append([])
    rows.append(["이번 주 포인트 (PPT 문장)"])
    for lab, txt in ctx["bullets"]:
        rows.append([lab, txt, None, None, None, None, None, BULLET_SRC.get(lab, "")])
    sheet(wb, "요약", ["항목", "단위", "이번 주/최근", "비교값(전주·직전)", "변화율(%)", "직전 4주 평균", "기준", "출처"],
          rows, {3: "#,##0.00", 4: "#,##0.00", 5: "0.0", 6: "#,##0.00"}, [30, 12, 14, 16, 10, 14, 34, 16], first=True)

    days = sorted(R)
    fxp, fxk = ctx["fxp"], ctx["fxk"]
    sheet(wb, "전력_일별",
          ["날짜", "요일", "폴란드 EUR/MWh", "폴란드 PLN/MWh", "폴란드 원/kWh", "독일 EUR/MWh", "체코", "슬로바키아",
           "리투아니아", "스웨덴 남부", "폴란드 15분 최고 EUR", "폴란드 15분 최저 EUR", "마이너스 가격 시간", "EUR→PLN 적용"],
          [[dd(d), WD[dd(d).weekday()], r2(R[d].get("px_PL")), r2(R[d].get("px_PL_pln")), r2(R[d].get("px_PL_krw")),
            r2(R[d].get("px_DE")), r2(R[d].get("px_CZ")), r2(R[d].get("px_SK")), r2(R[d].get("px_LT")),
            r2(R[d].get("px_SE4")), r2(R[d].get("px_PL_max")), r2(R[d].get("px_PL_min")), r2(R[d].get("px_PL_neg_h")),
            r2(fxp.on(d), 4)] for d in days if R[d].get("px_PL") is not None],
          {1: "yyyy-mm-dd", 3: "#,##0.00", 4: "#,##0.00", 5: "#,##0.0", 6: "#,##0.00"},
          [12, 6, 12, 12, 11, 12, 10, 10, 10, 10, 13, 13, 11, 11])
    sheet(wb, "수급기상_일별",
          ["날짜", "수요 MW", "풍력 MW", "  육상 MW", "  해상 MW", "태양광 MW", "잔여부하 MW", "석탄·갈탄 MW",
           "  유연탄 MW", "  갈탄 MW", "가스 MW", "기온 °C", "풍속 100m km/h", "일사 kWh/m2"],
          [[dd(d), r2(R[d].get("load"), 0), r2(R[d].get("wind"), 0), r2(R[d].get("wind_on"), 0),
            r2(R[d].get("wind_off"), 0), r2(R[d].get("solar"), 0), r2(R[d].get("resload"), 0), r2(R[d].get("coal"), 0),
            r2(R[d].get("hardcoal"), 0), r2(R[d].get("lignite"), 0), r2(R[d].get("gas"), 0), r2(R[d].get("temp"), 1),
            r2(R[d].get("wind100"), 1), r2(R[d].get("rad"), 2)] for d in days
           if any(R[d].get(k) is not None for k in ("load", "wind", "temp"))],
          {1: "yyyy-mm-dd"}, [12] + [11] * 13)
    if FLW:
        bn = [b for b in ("DE", "CZ", "SK", "LT", "SE4", "UA") if b in FL.get("borders", {})]
        nmk = {"DE": "독일", "CZ": "체코", "SK": "슬로바키아", "LT": "리투아니아", "SE4": "스웨덴 남부", "UA": "우크라이나"}
        hdr = (["날짜", "요일", "순수입 합계 GWh", "수입 합계 GWh", "수출 합계 GWh"] +
               ["%s 순수입" % nmk[b] for b in bn] + [x for b in bn for x in ("%s 수입" % nmk[b], "%s 수출" % nmk[b])])
        fr = []
        for d in sorted(FLW):
            r = FLW[d]
            fr.append([dd(d), WD[dd(d).weekday()], r2(r.get("total_net"), 1), r2(r.get("total_imp"), 1),
                       r2(r.get("total_exp"), 1)] + [r2(r.get(b + "_net"), 1) for b in bn] +
                      [r2(r.get(b + k), 1) for b in bn for k in ("_imp", "_exp")])
        sheet(wb, "국경흐름_일별", hdr, fr, dict([(1, "yyyy-mm-dd")] + [(c, "#,##0.0") for c in range(3, len(hdr) + 1)]),
              [12, 6, 13, 12, 12] + [11] * (len(hdr) - 5))
    eua, brent = ctx["daily_src"]["EUA"], ctx["daily_src"]["BRENT"]
    cd = sorted(set(eua) | set(brent))
    sheet(wb, "원자재_일별", ["날짜", "탄소배출권 EUA 경매 EUR/t", "브렌트유 USD/bbl"],
          [[dd(d), r2(eua.get(d)), r2(brent.get(d))] for d in cd], {1: "yyyy-mm-dd"}, [12, 18, 16])
    mk = [k for k in ("WB_TTF_EU", "WB_COAL_AU", "WB_COAL_ZA", "WB_HH_US", "WB_LNG_JP", "EUA_EUA_EEXauc") if k in MON]
    per = sorted({p["period"] for k in mk for p in MON[k]["points"]})
    per = [p for p in per if p >= "2020-01"]
    look = {k: {p["period"]: p["value"] for p in MON[k]["points"]} for k in mk}
    sheet(wb, "원자재_월별", ["월"] + [MON_HDR.get(k, MON[k]["label"]) for k in mk],
          [[p] + [r2(look[k].get(p)) for k in mk] for p in per], None, [10] + [20] * len(mk))
    fd = sorted(set(FXD.get("PLN", {})) | set(FXD.get("KRW", {})) | set(FXD.get("USD", {})))
    sheet(wb, "환율_ECB", ["날짜", "EUR→PLN", "EUR→KRW", "EUR→USD"],
          [[dd(d), FXD.get("PLN", {}).get(d), FXD.get("KRW", {}).get(d), FXD.get("USD", {}).get(d)] for d in fd],
          {1: "yyyy-mm-dd"}, [12, 11, 11, 11])
    C = ctx.get("COMP") or []
    if C:
        sheet(wb, "가격구성_월별",
              ["월", "자료 일수", "실제 가격 PLN/MWh", "가스 연료비", "가스 탄소비", "석탄 연료비", "석탄 탄소비", "운전비",
               "희소성 프리미엄", "재생에너지 과잉 할인", "가스 쪽 무게(%, 환산)", "가스 발전원가", "석탄 발전원가",
               "가스 TTF 환산 PLN/MWh(열량)", "가스 TTF $/MMBtu", "석탄 뉴캐슬 $/t", "EUA EUR/t", "EUR→PLN", "비고"],
              [[c["m"], c["days"], r2(c["P"], 0), r2(c["gF"], 0), r2(c["gC"], 0), r2(c["cF"], 0), r2(c["cC"], 0),
                r2(c["V"], 0), r2(c["prem"], 0), r2(c["disc"], 0), r2(c["w"] * 100, 0), r2(c["G"], 0), r2(c["C"], 0),
                r2(c["gas_mwh"], 0), r2(c["ttf_usd"]), r2(c["coal_usd"], 1), r2(c["eua"]), r2(c["pln"], 4), c["note"]]
               for c in C],
              {k: "#,##0" for k in range(3, 15)}, [10, 8, 12, 11, 11, 11, 11, 9, 12, 12, 13, 12, 12, 14, 12, 12, 10, 10, 28])
    notes = [["전력 하루전 가격", "ENTSO-E Transparency A44 (EUR/MWh). 15분·60분 자료는 현지 날짜별 평균. PLN·원 환산은 ECB 기준환율(주말은 직전 영업일)."],
             ["수요·발전", "ENTSO-E A65(실제 수요), A75(전원별 실제 발전). 20시간 미만인 날은 빈칸."],
             ["잔여부하", "수요 − 풍력 − 태양광. 화력(석탄·가스)이 채워야 하는 몫이라 가격을 가장 잘 설명."],
             ["수급 분해", "보고 주 직전 8주 일별 자료로 '가격 ~ 잔여부하 + 시간추세' 직선을 잡아, 잔여부하 변화 × 1GW당 가격으로 계산. 시간추세는 연료가격처럼 서서히 움직이는 몫을 떼어 내는 항."],
             ["탄소배출권", "EEX 1차시장 경매가(EUR/t). 같은 날 여러 경매는 평균. 항공용(EUAA)·ETS2 제외."],
             ["브렌트유", "미국 EIA 일별 현물(USD/bbl). EIA 는 주 1회 갱신이라 최근 며칠이 비어 있을 수 있음."],
             ["가스·석탄", "World Bank 월평균(원 단위 그대로). 무료 일별 자료가 없어 월평균만 제공."],
             ["기상", "Open-Meteo. 바르샤바·포모제·실롱스크 3지점 평균. 오늘 이후는 예보."],
             ["문장", "모든 문장은 숫자에서 규칙으로 자동 작성(변화 ±5% 안은 '비슷'). 사람이 쓴 해석이 아님."],
             ["주", "폴란드 현지 월~일. 자료 5일 미만 주간 값은 빈칸."],
             ["국경 흐름", "ENTSO-E A11 실제 물리 흐름(MW)을 현지 날짜별로 더해 GWh로 환산(20시간 미만인 날 제외). 순수입 = 수입 − 수출, 양수면 폴란드가 들여온 양. "
                         "국경 하나라도 빈 날은 합계 빈칸. 물리 흐름이라 거래소 계약량(상업 거래)과 다를 수 있음. 주변국 상황은 가격을 움직이는 5가지 요인 중 하나."],
             ["가격 구성", "발전원가 = 연료비 ÷ 효율(가스 0.55·석탄 0.41) + 배출계수(가스 0.37·석탄 0.83 t/MWh) × 배출권가 + 운전비 2€/MWh. "
                         "가스 쪽 무게 = (실제 가격 − 석탄 원가) ÷ (가스 원가 − 석탄 원가), 0~100%로 제한한 환산값(실제 시간 비율 아님). 가스 원가보다 비싸면 초과분 = 희소성 프리미엄, "
                         "석탄 원가보다 싸면 = 재생에너지 과잉 할인. 전력 = ENTSO-E × ECB, 가스 = World Bank TTF 월평균(공개자료 기준, TGE 아님)."]]
    sheet(wb, "설명", ["항목", "내용"], notes, None, [16, 120])
    sheet(wb, "출처", ["항목", "자료명", "주소"],
          SRC_LONG + [[], ["수집 시각", ctx["collected"], ""]], None, [16, 80, 70])
    hangul = re.compile("[ㄱ-ㆎ가-힣]")          # 한글 셀 = LG스마트체 · 나머지 = Arial Narrow
    for sh in wb.worksheets:
        for row in sh.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                f = c.font
                c.font = Font(name=KO if (isinstance(c.value, str) and hangul.search(c.value)) else "Arial Narrow",
                              bold=f.bold, italic=f.italic, size=f.size, color=f.color)
    wb.save(path)


# ══════════════════════════════════════════════════════════════════
def main():
    D, M = load("pl_drivers_daily.json"), load("pl_market_daily.json")
    R = {r["date"]: r for r in D.get("daily", [])}
    if not R:
        print("!! 일별 자료 없음 — 보고서 생성 중단", flush=True)
        return
    FXD = M.get("fx", {})
    fxp, fxk = FxLine(FXD.get("PLN", {})), FxLine(FXD.get("KRW", {}))
    for d, r in R.items():
        p, k = fxp.on(d), fxk.on(d)
        r["px_PL_pln"] = r["px_PL"] * p if (r.get("px_PL") is not None and p) else None
        r["px_DE_pln"] = r["px_DE"] * p if (r.get("px_DE") is not None and p) else None
        r["px_PL_krw"] = r["px_PL"] * k / 1000 if (r.get("px_PL") is not None and k) else None

    ws, we = week_bounds()
    iso = ws.isocalendar()
    tag = "%d-W%02d" % (iso[0], iso[1])
    ds = drange(ws, we)
    dp = drange(ws - dt.timedelta(days=7), we - dt.timedelta(days=7))
    d4 = drange(ws - dt.timedelta(days=28), ws - dt.timedelta(days=1))
    W, P, F4 = wstats(R, ds), wstats(R, dp), wstats(R, d4)
    fit = fit_resload(R, ws.isoformat())
    dec = decompose(W, P, fit)
    EUA, BRENT, MON = M.get("eua", {}), M.get("brent", {}), M.get("monthly", {})
    K = {"EUA": last_vs_week(EUA, we.isoformat()), "BRENT": last_vs_week(BRENT, we.isoformat()),
         "TTF": last_month(MON, "WB_TTF_EU"), "COAL": last_month(MON, "WB_COAL_AU")}
    K = {k: v for k, v in K.items() if v}
    out_s, out_info = s_outlook(R, D.get("forecast", {}), W, we, fxp)
    bullets = [("가격", s_price(W, P)), ("수급", s_cause(dec, fit, W, P)), ("주변국", s_region(W, P)),
               ("원자재", s_commod(K)), ("다음 주", out_s)]
    ctx = {"R": R, "W": W, "P": P, "F4": F4, "K": K, "MON": MON, "FXD": FXD, "fxp": fxp, "fxk": fxk,
           "ws": ws, "we": we, "tag": tag, "iso": (iso[0], iso[1]),
           "issued": dt.datetime.now(TZ).strftime("%Y-%m-%d"),
           "collected": "%s (폴란드 시각)" % D.get("generated", "미상"),
           "chart_days": drange(ws - dt.timedelta(days=CHART_DAYS - 7), we),
           "daily_src": {"EUA": EUA, "BRENT": BRENT}, "bullets": bullets,
           "fx_w": {"PLN": fxp.avg(ds[0], ds[-1]), "KRW": fxk.avg(ds[0], ds[-1])},
           "fx_p": {"PLN": fxp.avg(dp[0], dp[-1]), "KRW": fxk.avg(dp[0], dp[-1])},
           "M": M, "DG": load("pl_drivers_diag.json"), "FC": D.get("forecast", {}),
           "dec": dec, "fit": fit, "out_info": out_info, "FLOW": load("pl_flows_daily.json")}
    try:
        ctx["COMP"] = composition(R, M, we.isoformat())
    except Exception as e:
        print("!! 가격 구성 계산 실패: %r" % e, flush=True)
        ctx["COMP"] = []

    wdir, ldir = os.path.join(OUT_DIR, tag), os.path.join(OUT_DIR, "latest")
    os.makedirs(wdir, exist_ok=True)
    os.makedirs(ldir, exist_ok=True)
    base = os.path.join(wdir, BASENAME % tag)
    build_ppt(base + ".pptx", ctx)
    build_xlsx(base + ".xlsx", ctx)
    for ext in (".pptx", ".xlsx"):
        shutil.copyfile(base + ext, os.path.join(ldir, (BASENAME % "latest") + ext))
    with open(os.path.join(ldir, "summary.json"), "w", encoding="utf-8") as f:
        json.dump({"week": tag, "from": ws.isoformat(), "to": we.isoformat(), "issued": ctx["issued"],
                   "this_week": W, "prev_week": P, "commodities": K, "fit": fit, "decompose": dec,
                   "outlook": out_info, "bullets": bullets}, f, ensure_ascii=False, indent=2, default=str)

    print("\n=== %s (%s~%s) ===" % (tag, ws, we), flush=True)
    for lab, txt in bullets:
        print("  [%s] %s" % (lab, txt), flush=True)
    print("[OK] %s.pptx / .xlsx  → %s 에도 복사" % (base, ldir), flush=True)


if __name__ == "__main__":
    main()
