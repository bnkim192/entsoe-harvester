# -*- coding: utf-8 -*-
"""
폴란드 전력단가 및 원자재 가격동향 — 주간 1장 PPT + Excel raw (2026-11-09)

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
import os, json, math, shutil, bisect
import datetime as dt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_MARK, XL_MARKER_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from fetch_pl_drivers import ols, TZ

FLAT, NEED, FIT_DAYS, CHART_DAYS, CHART_MONTHS = 5.0, 5, 56, 84, 18
WD = "월화수목금토일"
KO, LAT = "LG스마트체 Regular", "Arial Narrow"
INK, INK2, MUTED, GRID, AXIS, CARD = "0B0B0B", "52514E", "898781", "E1E0D9", "C3C2B7", "F4F3EF"
COL = {"PL": "2A78D6", "DE": "EB6834", "EUA": "1BAF7A", "TTF": "6250D6", "COAL": "008300", "BRENT": "E34948"}
OUT_DIR = "weekly"
BASENAME = "PL_power_commodity_%s"


# ══════════════════════════════════════════════════════════════════
#  공통
# ══════════════════════════════════════════════════════════════════
def f0(v, nd=0):
    return "—" if v is None else ("{:,.%df}" % nd).format(v)


def sg(v, nd=0, unit=""):
    if v is None:
        return "—"
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
            "chg_pct": pct(a["value"], b["value"]), "unit": s.get("unit", ""), "label": s.get("label", key)}


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


def kpi(sl, x, y, w, h, color, label, value, unit, sub1, sub2):
    c = rect(sl, x, y, w, h, CARD, rounded=True)
    rect(sl, x, y + 0.12, 0.05, h - 0.24, color)
    fill_tf(c.text_frame, [[(label, 9, True, INK2)],
                           [(value, 18, True, INK), ("  " + unit, 8.5, False, INK2)],
                           [(sub1, 8, False, INK2)], [(sub2, 8, False, MUTED)]])
    c.text_frame.margin_left = Inches(0.16)


def nice_range(vals, n=4):
    lo, hi = min(vals), max(vals)
    if hi == lo:
        hi = lo + 1
    pad = (hi - lo) * 0.08
    lo, hi = lo - pad, hi + pad
    raw = (hi - lo) / n
    mag = 10 ** math.floor(math.log10(raw))
    step = next(m * mag for m in (1, 2, 2.5, 5, 10) if m * mag >= raw)
    return math.floor(lo / step) * step, math.ceil(hi / step) * step, step


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
        fill_tf(c.text_frame, [[(title, 9, True, INK2)], [("자료 없음 (수집 실패 — Actions 로그 확인)", 8.5, False, MUTED)]])
        return
    cd = CategoryChartData()
    cd.categories = cats
    for name, vs, _ in series:
        cd.add_series(name, [None if v is None else round(v, nd) for v in vs])
    ch = sl.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    ch.font.size = Pt(8)
    ch.font.name = LAT
    ch.font.color.rgb = rgb(MUTED)
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
    va.tick_labels.number_format = "#,##0" if step >= 1 else "#,##0.0"
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
             [("발행 %s · 폴란드 현지 월~일 기준" % ctx["issued"], 8.5, False, MUTED)]],
            anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
    rect(sl, 0.4, 0.86, 12.53, 0.02, AXIS)

    # ── KPI 6장
    cw, gap, y0, h0 = 1.99, 0.118, 1.0, 1.02

    def kx(i):
        return 0.4 + i * (cw + gap)
    p_pl = pct(W["px_PL_pln"], P["px_PL_pln"])
    kpi(sl, kx(0), y0, cw, h0, COL["PL"], "폴란드 전력 (주평균)", f0(W["px_PL_pln"]), "PLN/MWh",
        "%s EUR/MWh · %s 원/kWh" % (f0(W["px_PL"]), f0(W["px_PL_krw"], 1)), "전주 대비 %s" % sg(p_pl, 1, "%"))
    kpi(sl, kx(1), y0, cw, h0, COL["DE"], "독일 전력 (주평균)", f0(W["px_DE"]), "EUR/MWh",
        "%s PLN/MWh" % f0(W["px_DE_pln"]), "전주 대비 %s" % sg(pct(W["px_DE"], P["px_DE"]), 1, "%"))
    k = K.get("EUA")
    kpi(sl, kx(2), y0, cw, h0, COL["EUA"], "탄소배출권 EUA", f0(k["value"], 2) if k else "—", "€/t",
        ("%s 경매" % mdw(k["date"])) if k else "자료 없음",
        ("1주 전 대비 %s" % sg(k["chg_pct"], 1, "%")) if k else "")
    for i, (key, label) in enumerate((("TTF", "가스 TTF (월평균)"), ("COAL", "석탄 뉴캐슬 (월평균)"))):
        k = K.get(key)
        kpi(sl, kx(3 + i), y0, cw, h0, COL[key], label, f0(k["value"], 1) if k else "—",
            k["unit"] if k else "", ("%s년 %d월" % (k["period"][:4], int(k["period"][5:7]))) if k else "자료 없음",
            ("전월 대비 %s" % sg(k["chg_pct"], 1, "%")) if k else "")
    k = K.get("BRENT")
    kpi(sl, kx(5), y0, cw, h0, COL["BRENT"], "브렌트유", f0(k["value"], 1) if k else "—", "$/bbl",
        ("%s 기준" % mdw(k["date"])) if k else "자료 없음",
        ("1주 전 대비 %s" % sg(k["chg_pct"], 1, "%")) if k else "")

    # ── 차트: 전력 (왼쪽 큰 것)
    cs = ctx["chart_days"]
    line_chart(sl, 0.4, 2.15, 7.3, 2.85, [md(d) for d in cs],
               [("폴란드", [ctx["R"].get(d, {}).get("px_PL_pln") for d in cs], COL["PL"]),
                ("독일", [ctx["R"].get(d, {}).get("px_DE_pln") for d in cs], COL["DE"])],
               "하루전 전력가격 일평균 (PLN/MWh, 최근 12주 · ECB 환율 환산)", legend=True, skip=7)

    # ── 차트: 원자재 2×2 (각자 단위 → 각자 축)
    mx, my, mw, mh = 7.85, 2.15, 2.5, 1.4
    for i, (key, title, nd) in enumerate((("EUA", "탄소배출권 EUA (€/t, 경매일)", 2),
                                          ("BRENT", "브렌트유 ($/bbl, 거래일)", 1))):
        src = ctx["daily_src"][key]
        ks = [d for d in sorted(src) if cs[0] <= d <= we.isoformat()]
        line_chart(sl, mx + i * (mw + 0.08), my, mw, mh, [md(d) for d in ks],
                   [(title, [src[d] for d in ks], COL[key])], title, nd=nd, skip=max(1, len(ks) // 4), tsize=8.5)
    for i, (key, wbk, title) in enumerate((("TTF", "WB_TTF_EU", "가스 TTF (%s, 월평균)"),
                                           ("COAL", "WB_COAL_AU", "석탄 뉴캐슬 (%s, 월평균)"))):
        s = ctx["MON"].get(wbk, {})
        pts = s.get("points", [])[-CHART_MONTHS:]
        t = title % s.get("unit", "")
        line_chart(sl, mx + i * (mw + 0.08), my + mh + 0.05, mw, mh, [p["period"][2:].replace("-", ".") for p in pts],
                   [(t, [p["value"] for p in pts], COL[key])], t, nd=1, skip=3, tsize=8.5)

    # ── 이번 주 포인트
    rect(sl, 0.4, 5.12, 8.15, 1.86, CARD, rounded=True)
    lines = [[("이번 주 포인트", 10.5, True, INK)]]
    for lab, txt in ctx["bullets"]:
        if txt:
            lines.append([("▪ %s  " % lab, 8.5, True, INK), (txt, 8.5, False, INK2)])
    textbox(sl, 0.48, 5.16, 8.0, 1.8, lines, after=2)

    # ── 수급 표
    rows = [["수급·기상 (일평균)", "이번 주", "전주", "차이"]]
    for key, nm, nd, unit in (("load", "수요", 1, "GW"), ("wind", "풍력 발전", 1, "GW"),
                              ("solar", "태양광 발전", 1, "GW"), ("res", "잔여부하(화력 몫)", 1, "GW"),
                              ("coal", "석탄·갈탄 발전", 1, "GW"), ("gas", "가스 발전", 1, "GW"),
                              ("temp", "기온", 1, "°C"), ("wind100", "풍속(100m)", 1, "km/h")):
        a, b = W.get(key), P.get(key)
        rows.append(["%s (%s)" % (nm, unit), f0(a, nd), f0(b, nd),
                     sg(a - b, nd) if (a is not None and b is not None) else "—"])
    table(sl, 8.7, 5.12, 4.23, 1.86, rows, [1.75, 0.82, 0.82, 0.84])

    # ── 출처
    textbox(sl, 0.35, 7.05, 12.6, 0.32,
            [[("자료: ENTSO-E(하루전 가격·수요·발전), EEX(탄소배출권 경매), 미국 EIA(브렌트), World Bank(가스·석탄 월평균), "
               "ECB(환율), Open-Meteo(기상·예보). 잔여부하 = 수요 − 풍력 − 태양광. 문장은 숫자에서 규칙으로 자동 작성. "
               "가스·석탄은 무료 일별 자료가 없어 월평균.", 7, False, MUTED)]])
    prs.save(path)


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
    for ci, f in (fmts or {}).items():
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
    add("환율 EUR→PLN", "PLN", ctx["fx_w"]["PLN"], ctx["fx_p"]["PLN"], wk, "ECB")
    add("환율 EUR→KRW", "KRW", ctx["fx_w"]["KRW"], ctx["fx_p"]["KRW"], wk, "ECB")
    rows.append([])
    rows.append(["이번 주 포인트 (PPT 문장)"])
    for lab, txt in ctx["bullets"]:
        rows.append([lab, txt])
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
    eua, brent = ctx["daily_src"]["EUA"], ctx["daily_src"]["BRENT"]
    cd = sorted(set(eua) | set(brent))
    sheet(wb, "원자재_일별", ["날짜", "탄소배출권 EUA 경매 EUR/t", "브렌트유 USD/bbl"],
          [[dd(d), r2(eua.get(d)), r2(brent.get(d))] for d in cd], {1: "yyyy-mm-dd"}, [12, 18, 16])
    mk = [k for k in ("WB_TTF_EU", "WB_COAL_AU", "WB_COAL_ZA", "WB_HH_US", "WB_LNG_JP", "EUA_EUA_EEXauc") if k in MON]
    per = sorted({p["period"] for k in mk for p in MON[k]["points"]})
    per = [p for p in per if p >= "2020-01"]
    look = {k: {p["period"]: p["value"] for p in MON[k]["points"]} for k in mk}
    sheet(wb, "원자재_월별", ["월"] + ["%s (%s)" % (MON[k]["label"], MON[k]["unit"]) for k in mk],
          [[p] + [r2(look[k].get(p)) for k in mk] for p in per], None, [10] + [20] * len(mk))
    fd = sorted(set(FXD.get("PLN", {})) | set(FXD.get("KRW", {})) | set(FXD.get("USD", {})))
    sheet(wb, "환율_ECB", ["날짜", "EUR→PLN", "EUR→KRW", "EUR→USD"],
          [[dd(d), FXD.get("PLN", {}).get(d), FXD.get("KRW", {}).get(d), FXD.get("USD", {}).get(d)] for d in fd],
          {1: "yyyy-mm-dd"}, [12, 11, 11, 11])
    notes = [["전력 하루전 가격", "ENTSO-E Transparency A44 (EUR/MWh). 15분·60분 자료는 현지 날짜별 평균. PLN·원 환산은 ECB 기준환율(주말은 직전 영업일)."],
             ["수요·발전", "ENTSO-E A65(실제 수요), A75(전원별 실제 발전). 20시간 미만인 날은 빈칸."],
             ["잔여부하", "수요 − 풍력 − 태양광. 화력(석탄·가스)이 채워야 하는 몫이라 가격을 가장 잘 설명."],
             ["수급 분해", "보고 주 직전 8주 일별 자료로 '가격 ~ 잔여부하 + 시간추세' 직선을 잡아, 잔여부하 변화 × 1GW당 가격으로 계산. 시간추세는 연료가격처럼 서서히 움직이는 몫을 떼어 내는 항."],
             ["탄소배출권", "EEX 1차시장 경매가(EUR/t). 같은 날 여러 경매는 평균. 항공용(EUAA)·ETS2 제외."],
             ["브렌트유", "미국 EIA 일별 현물(USD/bbl). EIA 는 주 1회 갱신이라 최근 며칠이 비어 있을 수 있음."],
             ["가스·석탄", "World Bank 월평균(원 단위 그대로). 무료 일별 자료가 없어 월평균만 제공."],
             ["기상", "Open-Meteo. 바르샤바·포모제·실롱스크 3지점 평균. 오늘 이후는 예보."],
             ["문장", "모든 문장은 숫자에서 규칙으로 자동 작성(변화 ±5% 안은 '비슷'). 사람이 쓴 해석이 아님."],
             ["주", "폴란드 현지 월~일. 자료 5일 미만 주간 값은 빈칸."]]
    sheet(wb, "설명", ["항목", "내용"], notes, None, [16, 120])
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
           "chart_days": drange(ws - dt.timedelta(days=CHART_DAYS - 7), we),
           "daily_src": {"EUA": EUA, "BRENT": BRENT}, "bullets": bullets,
           "fx_w": {"PLN": fxp.avg(ds[0], ds[-1]), "KRW": fxk.avg(ds[0], ds[-1])},
           "fx_p": {"PLN": fxp.avg(dp[0], dp[-1]), "KRW": fxk.avg(dp[0], dp[-1])}}

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
