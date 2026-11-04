# -*- coding: utf-8 -*-
"""계산 결과를 회사 양식(경영기획팀 '사업비 예산편성 양식')에 그대로 채워 보고서 파일을 만든다.

회사 양식 파일의 서식을 그대로 쓰기 위해 Excel(COM)로 연다. 행이 더 필요하면 양식의 원래 행
(첫 줄·중간 줄·소계·합계)을 복사해 끼워 넣고, 원래 견본 행은 지운다. PM팀 2026 제출본과 같은 방식.

  build_report(xl, 양식경로, 출력경로, res, data, preview_dir=None) → dict(previews=[…], sheets=…)

res  : site_io.excel_calc 결과(groups·rates·calc·price·interest·stage·year)
data : 화면 입력값(fields·types·monthly) — 보고서 안 '_자동작성입력'(아주 숨김) 시트에 넣어 다시 불러올 수 있게 한다
"""
import datetime
import json
import math
import os
import re
import shutil

XL_DOWN = -4121
XL_HAIR, XL_THIN = 1, 2
EDGE_TOP, EDGE_BOTTOM = 8, 9
XL_INSIDE_H, XL_CONTINUOUS, XL_NONE = 12, 1, -4142
HIDDEN, VERY_HIDDEN, VISIBLE = 0, 2, -1
DATA_SHEET = "_자동작성입력"
ERRORS = []          # 미리보기 등 치명적이지 않은 오류(서버 기록용)


def xround(x):
    """Excel ROUND(x,0) 과 같게(.5 는 올림)"""
    return int(math.floor(x + 0.5)) if x >= 0 else -int(math.floor(-x + 0.5))


def L(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def latest_form(folder):
    """회사양식 폴더에서 연도가 가장 큰(같으면 최근 수정) 양식"""
    best = None
    for fn in os.listdir(folder):
        if not fn.lower().endswith(".xlsx") or fn.startswith("~$"):
            continue
        p = os.path.join(folder, fn)
        m = re.search(r"(20\d{2})", fn)
        key = (int(m.group(1)) if m else 0, os.path.getmtime(p))
        if best is None or key > best[0]:
            best = (key, p)
    return best[1] if best else None


# ── 공통 ──────────────────────────────────────────────────────────
def _vals(ws, last_col="V", last_row=None):
    last_row = last_row or (ws.UsedRange.Row + ws.UsedRange.Rows.Count + 2)
    v = ws.Range(f"A1:{last_col}{last_row}").Value
    return [list(r) for r in v]


def _txt(v):
    return "" if v is None else re.sub(r"\s+", "", str(v))


def _find(vals, col, pred, start=0):
    for i in range(start, len(vals)):
        if pred(_txt(vals[i][col])):
            return i + 1
    return None


def _copy_rows(ws, src, dst_first, dst_last):
    ws.Rows(src).Copy(ws.Range(f"A{dst_first}:A{dst_last}").EntireRow)


def rebuild(ws, start, end, layout, last_col):
    """start..end 행(견본)을 layout 대로 새로 짠다. layout = [견본행번호, …] (원래 번호).
    end 바로 아래에 새 행을 끼우고 견본 서식을 복사한 뒤 견본 행을 지운다. 새 행은 start 부터."""
    n = len(layout)
    ws.Rows(f"{end + 1}:{end + n}").Insert(XL_DOWN)
    i = 0
    while i < n:                     # 같은 견본이 이어지는 구간은 한 번에 복사
        j = i
        while j + 1 < n and layout[j + 1] == layout[i]:
            j += 1
        _copy_rows(ws, layout[i], end + 1 + i, end + 1 + j)
        i = j + 1
    ws.Rows(f"{start}:{end}").Delete()
    return start, start + n - 1


def write_block(ws, r0, rows, first_col, last_col):
    """rows: [[값…], …] 를 r0 행 first_col 열부터 한 번에(.Formula)"""
    if not rows:
        return
    ncol = ord(last_col) - ord(first_col) + 1 if len(last_col) == 1 else None
    data = []
    for row in rows:
        row = list(row) + [""] * (ncol - len(row))
        data.append(tuple("" if v is None else v for v in row[:ncol]))
    ws.Range(f"{first_col}{r0}:{last_col}{r0 + len(rows) - 1}").Formula = tuple(data)


def shift_years(wb, fy, ty):
    """양식 연도(fy) 글자를 예산연도(ty)에 맞게: '25년 실적'→'26년 실적', '2026년'→'2027년' (상수 글만)"""
    d = ty - fy
    if not d:
        return 0
    cnt = 0

    def repl(m):
        y = int(m.group(1))
        full = y if y >= 2000 else 2000 + y
        if fy - 3 <= full <= fy + 1:
            full += d
            return str(full) if y >= 2000 else f"{full % 100:02d}"
        return m.group(1)
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        f = ur.Formula
        if not isinstance(f, tuple):
            f = ((f,),)
        for i, row in enumerate(f):
            for j, s in enumerate(row):
                if isinstance(s, str) and s and not s.startswith("=") and "년" in s:
                    new = re.sub(r"(?<!\d)(20\d{2}|\d{2})(?=\s*년)", repl, s)
                    if new != s:
                        ws.Cells(ur.Row + i, ur.Column + j).Value = new
                        cnt += 1
    return cnt


def widen(ws, cols):
    """값이 ##### 로 보이지 않게 필요한 열만 넓힘(양식보다 좁히지는 않음)"""
    for col in cols:
        c = ws.Columns(col)
        w0 = c.ColumnWidth
        c.AutoFit()
        if c.ColumnWidth < w0:
            c.ColumnWidth = w0


def _lines(text, width_units=72):
    """산출근거 글이 차지하는 줄 수(줄바꿈 + 칸 너비를 넘치는 만큼)"""
    n = 0
    for line in (text or "").split("\n"):
        w = sum(1.9 if ord(ch) > 0x2E80 else 1.0 for ch in line)
        n += max(1, math.ceil(w / width_units))
    return n


# ── 27년 예산 ─────────────────────────────────────────────────────
def fill_budget(ws, res, title, date_txt, site, keep_interest=False):
    v = _vals(ws, "U")
    hr = _find(v, 3, lambda t: t == "세목")
    tot = _find(v, 0, lambda t: t.startswith("합"), hr)
    sub = _find(v, 3, lambda t: t == "소계", hr)
    T = dict(first=hr + 1, mid=hr + 2, sub=sub, first2=sub + 1, tot=tot)
    for g, gf in zip(res["groups"], res.get("formulas") or []):      # 월별 산식(계산 과정) — 줄마다 12개월, 없으면 None
        for r, rf in zip(g["rows"], gf or []):
            r["fx"] = rf
    groups = []
    for g in res["groups"]:
        rows = [r for r in g["rows"] if r["total"] or (r["manual"] and r["g"])]
        if not rows and keep_interest and g["sub"] == "대납이자":       # 무이자 현장: 대출잔액은 엑셀에서 채움
            rows = [r for r in g["rows"] if not r["manual"]][:1]
        if rows:
            groups.append((g, rows))
    layout, kinds = [], []
    for gi, (g, rows) in enumerate(groups):
        for k, r in enumerate(rows):
            layout.append(T["first"] if (gi == 0 and k == 0) else (T["first2"] if k == 0 else T["mid"]))
            kinds.append(("row", gi, k, r))
        layout.append(T["sub"])
        kinds.append(("sub", gi, None, None))
    a, b = rebuild(ws, T["first"], T["tot"] - 1, layout, "U")
    tot_row = b + 1
    out, subs, sub_of = [], [], {}
    start_of = {}
    for i, (kind, gi, k, r) in enumerate(kinds):
        rr = a + i
        g = groups[gi][0]
        if kind == "row":
            start_of.setdefault(gi, rr)
            fx = r.get("fx") or [None] * 12
            months = [("=" + f) if f else ("" if not m else m) for m, f in zip(r["months"], fx)]   # 산식이 있으면 수식으로
            out.append([g["item"] if k == 0 else "", "", g["sub"] if k == 0 else "", "", f"=I{rr}", " " + r["g"].strip(),
                        "", f"=SUM(J{rr}:U{rr})"] + months)
            ws.Rows(rr).RowHeight = 19.5 if _lines(r["g"]) == 1 else 15 * _lines(r["g"]) + 6
        else:
            s0 = start_of[gi]
            sums = [f"=SUM({L(10 + j)}{s0}:{L(10 + j)}{rr - 1})" for j in range(12)]
            out.append(["", "", "소  계", f"=SUM(E{s0}:E{rr - 1})", f"=SUM(F{s0}:F{rr - 1})", "", "", f"=SUM(J{rr}:U{rr})"] + sums)
            subs.append(rr)
            sub_of[gi] = rr
    write_block(ws, a, out, "B", "U")
    ws.Range(f"G{a}:G{b}").WrapText = True        # 산출근거가 길거나 여러 줄(합친 세목)이어도 칸 안에 다 보이게
    # 회사 양식은 '실적' 열(E) 가운데 줄의 윗선(점선)이 빠져 있어 행을 복사하면 실적 칸 사이 점선이 사라짐
    # → 항목 묶음마다 실적 칸 사이에 세목 열(D)과 같은 점선을 다시 그린다
    for gi in start_of:
        s0, s1 = start_of[gi], sub_of[gi] - 1
        if s1 > s0:
            src = ws.Range(f"D{s0 + 1}").Borders(EDGE_TOP)
            b = ws.Range(f"E{s0}:E{s1}").Borders(XL_INSIDE_H)
            b.LineStyle = src.LineStyle if src.LineStyle != XL_NONE else XL_CONTINUOUS
            b.Weight = src.Weight if src.LineStyle != XL_NONE else XL_HAIR
            if src.LineStyle != XL_NONE:
                b.Color = src.Color
    plus = lambda c: "=" + "+".join(f"{c}{s}" for s in subs) if subs else ""
    write_block(ws, tot_row, [[plus("E"), plus("F"), "", "", f"=SUM(J{tot_row}:U{tot_row})"]
                              + [plus(L(10 + j)) for j in range(12)]], "E", "U")
    ws.Range("A2").Value = title
    ws.Range("A3").Value = date_txt
    ws.Range(f"A{hr - 1}").Value = f"[ {site} ]"
    labels = res.get("stage") or []
    write_block(ws, hr - 1, [[x or "" for x in labels]], "J", "U")
    return dict(groups=[(groups[gi][0], start_of[gi], sub_of[gi]) for gi in range(len(groups))], total=tot_row, header=hr)


# ── 예산안 보고서 ─────────────────────────────────────────────────
def fill_report(ws, budget_name, binfo, title, date_txt, site, note_txt):
    v = _vals(ws, "P")
    hr = _find(v, 1, lambda t: t == "항목")
    first = hr + 2
    sub = _find(v, 3, lambda t: t == "소계", hr)
    tot = _find(v, 0, lambda t: t.startswith("합"), hr)
    T = dict(first=first, mid=first + 1, sub=sub, first2=sub + 1, tot=tot)
    items = []
    for g, s0, s1 in binfo["groups"]:
        if items and items[-1][0] == g["item"]:
            items[-1][1].append((g, s0, s1))
        else:
            items.append((g["item"], [(g, s0, s1)]))
    layout, kinds = [], []
    for ii, (item, gs) in enumerate(items):
        for k, x in enumerate(gs):
            layout.append(T["first"] if (ii == 0 and k == 0) else (T["first2"] if k == 0 else T["mid"]))
            kinds.append(("row", ii, k, x))
        layout.append(T["sub"])
        kinds.append(("sub", ii, None, None))
    a, b = rebuild(ws, T["first"], T["tot"] - 1, layout, "P")
    tot_row = b + 1
    out, subs, start = [], [], None
    for i, (kind, ii, k, x) in enumerate(kinds):
        rr = a + i
        if kind == "row":
            if k == 0:
                start = rr
            g, s0, s1 = x
            out.append([items[ii][0] if k == 0 else "", "", g["sub"], "", "", "", "", "", "",
                        f'=IF(H{rr}-E{rr}=0,"",IF(E{rr}=0,100%,(H{rr}-E{rr})/E{rr}))', "",
                        f"='{budget_name}'!F{s1}",
                        f'=IF(M{rr}-H{rr}=0,"",IF(H{rr}=0,100%,(M{rr}-H{rr})/H{rr}))', note_txt, ""])
        else:
            sm = lambda c: f"=SUM({c}{start}:{c}{rr - 1})"
            out.append(["", "", "소  계", sm("E"), sm("F"), sm("G"), sm("H"), sm("I"), sm("J"),
                        f'=IF(H{rr}-E{rr}=0,"",IF(E{rr}=0,100%,(H{rr}-E{rr})/E{rr}))', "", sm("M"),
                        f'=IF(M{rr}-H{rr}=0,"",IF(H{rr}=0,100%,(M{rr}-H{rr})/H{rr}))', "", ""])
            subs.append(rr)
    write_block(ws, a, out, "B", "P")
    plus = lambda c: "=" + "+".join(f"{c}{s}" for s in subs) if subs else ""
    write_block(ws, tot_row, [[plus(c) for c in "EFGHIJ"] +
                              [f'=IF(H{tot_row}-E{tot_row}=0,"",IF(E{tot_row}=0,100%,(H{tot_row}-E{tot_row})/E{tot_row}))', "",
                               plus("M"), f'=IF(M{tot_row}-H{tot_row}=0,"",IF(H{tot_row}=0,100%,(M{tot_row}-H{tot_row})/H{tot_row}))']],
                "E", "N")
    ws.Range("A2").Value = title
    ws.Range("A3").Value = date_txt
    ws.Range(f"A{hr - 1}").Value = f"[ {site} ]"
    return dict(header=hr, total=tot_row)


# ── 기본정보 ─────────────────────────────────────────────────────
def fill_basic(ws, fields, types):
    v = _vals(ws, "G", 40)
    rcnt = _find(v, 3, lambda t: t == "계")
    t_rows = [i + 1 for i in range(rcnt, len(v)) if _txt(v[i][3]) == "평형입력"]
    if not t_rows:
        t_rows = [rcnt + 1, rcnt + 2]
    t1, tn = t_rows[0], t_rows[-1]
    m = len(types)
    if m > len(t_rows):
        k = m - len(t_rows)
        ws.Rows(f"{tn}:{tn + k - 1}").Insert(XL_DOWN)
        ws.Range(f"D{t1}:G{t1}").Copy(ws.Range(f"D{tn}:G{tn + k - 1}"))
        for c in ("D", "E", "F", "G"):
            rg = ws.Range(f"{c}{tn}:{c}{tn + k - 1}")
            rg.Borders(EDGE_TOP).Weight = XL_HAIR
            rg.Borders(EDGE_BOTTOM).Weight = XL_HAIR
        tn += k
    rows = []
    for i in range(tn - t1 + 1):
        t = types[i] if i < len(types) else None
        rows.append([str(t[0]) if t else "", "", (t[1] if t and t[1] not in (None, "") else "")])
    write_block(ws, t1, rows, "D", "F")
    ws.Range(f"F{rcnt}").Formula = f"=SUM(F{t1}:F{tn})"
    v = _vals(ws, "G", 40)
    put = {"현장명": fields.get("현장명"), "담당소장": fields.get("담당소장"), "착공년월": fields.get("착공"),
           "준공년월": fields.get("준공"), "시행법인": fields.get("시행"),
           "시공법인": fields.get("시공") or fields.get("시행")}
    for i, row in enumerate(v):
        lab = _txt(row[2]) or _txt(row[3])
        if lab in put and put[lab] not in (None, ""):
            ws.Range(f"F{i + 1}").Value = str(put[lab])
    return dict(total=f"F{rcnt}")


# ── 분양율,입주율 ─────────────────────────────────────────────────
def fill_rates(ws, res, basic_total):
    v = _vals(ws, "U", 30)
    r_s = _find(v, 2, lambda t: t == "분양세대")
    r_i = _find(v, 2, lambda t: t == "입주세대")
    ws.Range("C4").Formula = f"=기본정보!{basic_total}"
    R = res["rates"]

    def nums(xs):
        return ["" if not x else x for x in xs]
    if r_s:
        write_block(ws, r_s, [nums(R["전년분양"]) + nums(R["분양"])], "E", "T")
    if r_i:
        write_block(ws, r_i, [nums(R["전년입주"]) + nums(R["입주"])], "E", "T")


# ── 분양대행 및 촉진수수료 ─────────────────────────────────────────
def _groups_by_size(types):
    order, units = [], {}
    for t in types:
        if not t or t[1] in (None, ""):
            continue
        m = re.match(r"\s*(\d+)", str(t[0]))
        key = m.group(1) if m else str(t[0])
        if key not in units:
            order.append(key)
            units[key] = 0
        units[key] += float(t[1])
    return order, units


def _distribute(total_by_month, keys, units):
    tot = sum(units[k] for k in keys)
    if not tot:
        return {k: [0] * 12 for k in keys}
    out = {k: [] for k in keys}
    for n in total_by_month:
        cum, prev = 0.0, 0
        for k in keys:
            cum += units[k] / tot
            cur = xround(n * cum)
            out[k].append(cur - prev)
            prev = cur
    return out


def fill_agency(ws, res, types):
    """분양대행수수료 / 분양대행수수료(조직) / 총계 / 분양촉진수수료(MGM) — 평형별 줄(84·101…), 2026 PM팀 제출본 구성"""
    c = res["calc"]
    keys, units = _groups_by_size(types)
    if not keys:
        keys, units = ["전체"], {"전체": 1.0}
    series = {"대행": c.get("대행세대") or [0] * 12, "조직": c.get("조직세대") or [0] * 12, "MGM": c.get("MGM") or [0] * 12}
    if not any(sum(x) for x in series.values()):
        ws.Visible = HIDDEN
        return False
    v = _vals(ws, "V", 40)
    r0 = _find(v, 0, lambda t: t == "분양대행수수료")                 # 표 안 구분 칸(제목 줄 제외)
    rm = _find(v, 0, lambda t: t.startswith("분양촉진수수료("))
    if not (r0 and rm and _txt(v[r0 - 1][1]) == "세대수" and _txt(v[r0 + 3][1]) == "금액" and rm - r0 == 10):
        return False                   # 양식이 예상과 다르면 손대지 않음
    S = dict(u1=r0, um=r0 + 1, ul=r0 + 2, us=r0 + 3, a1=r0 + 4, am=r0 + 5, al=r0 + 6, as_=r0 + 7, cs=r0 + 8, tot=r0 + 9,
             m1=rm, mm=rm + 1, ml=rm + 2, ms=rm + 3, ma1=rm + 4, mam=rm + 5, mal=rm + 6, mas=rm + 7)
    n = len(keys)

    def pick(f, m, l):
        return [f] if n == 1 else [f] + [m] * (n - 2) + [l]
    lay, kinds = [], []

    def add(name, mgm):
        p = (S["m1"], S["mm"], S["ml"], S["ms"], S["ma1"], S["mam"], S["mal"], S["mas"]) if mgm else \
            (S["u1"], S["um"], S["ul"], S["us"], S["a1"], S["am"], S["al"], S["as_"])
        lay.extend(pick(p[0], p[1], p[2]))
        kinds.extend((name, "u", i) for i in range(n))
        lay.append(p[3])
        kinds.append((name, "usum", None))
        lay.extend(pick(p[4], p[5], p[6]))
        kinds.extend((name, "a", i) for i in range(n))
        lay.append(p[7])
        kinds.append((name, "asum", None))
        if not mgm:
            lay.extend([S["cs"], S["tot"]])
            kinds.extend([(name, "consult", None), (name, "total", None)])
    add("대행", False)
    if sum(series["조직"]) > 0:
        add("조직", False)
        lay.append(S["tot"])
        kinds.append(("총계", "grand", None))
    add("MGM", True)
    ws.Range(f"A{r0}:C{S['mas']}").UnMerge()
    a, _ = rebuild(ws, r0, S["mas"], lay, "V")
    pos = {}
    for i, (name, kind, gi) in enumerate(kinds):
        pos.setdefault((name, kind), []).append(a + i)
    price = {"대행": res["price"].get("분양대행") or 0, "조직": res["price"].get("조직") or 0, "MGM": res["price"].get("MGM") or 0}
    dist = {k: _distribute(series[k], keys, units) for k in series}
    cols = ["F", "G", "H"] + [L(10 + j) for j in range(12)]          # 금액이 들어가는 열(전년 3개월 + 1~12월)
    out = []
    for i, (name, kind, gi) in enumerate(kinds):
        rr = a + i
        row = {c_: "" for c_ in ["C", "D", "E", "I"] + cols}
        if kind == "u":
            row["C"] = int(keys[gi]) if keys[gi].isdigit() else keys[gi]
            row["D"] = price[name]
            for j in range(12):
                row[L(10 + j)] = dist[name][keys[gi]][j] or ""
        elif kind == "a":
            u = pos[(name, "u")][gi]
            row["C"], row["D"] = f"=+C{u}", f"=+D{u}"
            for col in cols:
                row[col] = f"=+{col}{u}*$D${rr}"
        elif kind in ("usum", "asum"):
            src = pos[(name, "u" if kind == "usum" else "a")]
            for col in cols:
                row[col] = f"=+SUM({col}{src[0]}:{col}{src[-1]})"
        elif kind == "total":
            s_, c_r = pos[(name, "asum")][0], pos[(name, "consult")][0]
            for col in cols:
                row[col] = f"=+{col}{s_}+{col}{c_r}"
        elif kind == "grand":
            t1, t2 = pos[("대행", "total")][0], pos[("조직", "total")][0]
            for col in cols:
                row[col] = f"=+{col}{t1}+{col}{t2}"
        row["E"] = f"=+SUM(F{rr}:H{rr})"
        row["I"] = f"=+SUM(J{rr}:U{rr})"
        out.append([row[c_] for c_ in ["C", "D", "E", "F", "G", "H", "I"] + [L(10 + j) for j in range(12)]])
    write_block(ws, a, out, "C", "U")
    labels = {"대행": "분양대행\n수수료", "조직": "분양대행\n수수료\n(조직)", "MGM": "분양촉진\n수수료\n(MGM 등)"}
    for name in ("대행", "조직", "MGM"):
        if (name, "u") not in pos:
            continue
        u, am = pos[(name, "u")], pos[(name, "a")]
        last = pos[(name, "total")][0] if (name, "total") in pos else pos[(name, "asum")][0]
        ws.Range(f"A{u[0]}:A{last}").Merge()
        ws.Range(f"A{u[0]}").Value = labels[name]
        ws.Range(f"B{u[0]}:B{u[-1]}").Merge()
        ws.Range(f"B{u[0]}").Value = "세대수"
        ws.Range(f"B{am[0]}:B{am[-1]}").Merge()
        ws.Range(f"B{am[0]}").Value = "금 액"
        for k in ("usum", "asum"):
            ws.Range(f"B{pos[(name, k)][0]}").Value = ""
        if (name, "consult") in pos:
            cr, tr = pos[(name, "consult")][0], pos[(name, "total")][0]
            ws.Range(f"B{cr}:C{cr}").Merge()
            ws.Range(f"B{cr}").Value = "상담사 인건비"
            ws.Range(f"B{tr}:C{tr}").Merge()
            ws.Range(f"B{tr}").Value = "계"
    if ("총계", "grand") in pos:
        gr = pos[("총계", "grand")][0]
        ws.Range(f"A{gr}:C{gr}").Merge()
        ws.Range(f"A{gr}").Value = "총 계"
    widen(ws, [L(i) for i in range(5, 22)])
    return True


# ── 공가세대관리비 ───────────────────────────────────────────────
def fill_vacancy(ws, res, types):
    c = res["calc"]
    vac = c.get("미입주", [0] * 12)
    if not any(vac):
        ws.Visible = HIDDEN
        return False
    tl = [t for t in types if t and t[1] not in (None, "")]
    keys = [str(t[0]) for t in tl]
    units = {str(t[0]): float(t[1]) for t in tl}
    py = {str(t[0]): (math.ceil(float(t[2]) * 0.3025) if t[2] not in (None, "") else 0) for t in tl}
    v = _vals(ws, "V", 40)
    r7 = _find(v, 0, lambda t: t.startswith("평당"))
    r0 = r7 + 1 if r7 else None
    if not r0 or _txt(v[r0 - 1][1]) != "세대수" or _txt(v[r0 + 3][1]) != "금액":
        return False
    S = dict(u1=r0, um=r0 + 1, ul=r0 + 2, us=r0 + 3, a1=r0 + 4, am=r0 + 5, al=r0 + 6, as_=r0 + 7)
    n = max(1, len(keys))
    pick = lambda f, m_, l_: [f] if n == 1 else [f] + [m_] * (n - 2) + [l_]
    lay = pick(S["u1"], S["um"], S["ul"]) + [S["us"]] + pick(S["a1"], S["am"], S["al"]) + [S["as_"]]
    ws.Range(f"A{r0}:C{S['as_']}").UnMerge()
    a, b = rebuild(ws, r0, S["as_"], lay, "V")
    dist = _distribute(vac, keys, units) if keys else {}
    per = res["price"].get("공가") or 0
    months = [L(9 + j) for j in range(12)]      # I..T
    rows = []
    us_row, as_row = a + n, a + 2 * n + 1
    for i in range(n):                           # 세대수
        rr = a + i
        k = keys[i] if i < len(keys) else ""
        rows.append([k, f"=+SUM(E{rr}:G{rr})", "", "", "", f"=+SUM(I{rr}:T{rr})"] +
                    [(dist[k][j] or "") if k else "" for j in range(12)] + [f"{py.get(k, 0)}평" if k else ""])
    rows.append(["", f"=+SUM(E{us_row}:G{us_row})"] + [f"=+SUM({col}{a}:{col}{us_row - 1})" for col in "EFG"] +
                [f"=+SUM(I{us_row}:T{us_row})"] + [f"=+SUM({m}{a}:{m}{us_row - 1})" for m in months] + [""])
    for i in range(n):                           # 금액
        rr = as_row - n + i
        u = a + i
        k = keys[i] if i < len(keys) else ""
        rows.append([f"=+C{u}", f"=+SUM(E{rr}:G{rr})", "", "", "", f"=+SUM(I{rr}:T{rr})"] +
                    [f"=+{m}$7*{m}{u}*{py.get(k, 0)}" for m in months] + [""])
    rows.append(["", f"=+SUM(E{as_row}:G{as_row})"] + [f"=+SUM({col}{as_row - n}:{col}{as_row - 1})" for col in "EFG"] +
                [f"=+SUM(I{as_row}:T{as_row})"] + [f"=+SUM({m}{as_row - n}:{m}{as_row - 1})" for m in months] + [""])
    write_block(ws, a, rows, "C", "U")
    write_block(ws, r7, [[per if vac[j] else "" for j in range(12)]], "I", "T")
    ws.Range(f"A{a}:A{as_row}").Merge()
    ws.Range(f"A{a}").Value = "공가세대\n관리비"
    ws.Range(f"B{a}:B{us_row - 1}").Merge()
    ws.Range(f"B{a}").Value = "세대수"
    ws.Range(f"B{us_row + 1}:B{as_row - 1}").Merge()
    ws.Range(f"B{us_row + 1}").Value = "금 액"
    widen(ws, [L(i) for i in range(4, 21)])
    return dict(total=as_row)


# ── 대납이자 ─────────────────────────────────────────────────────
def fill_interest(ws, res, fields, ty):
    v = _vals(ws, "W", 45)
    r1 = _find(v, 0, lambda t: t == "중도금대출")
    r2 = _find(v, 0, lambda t: t == "대출이자")
    if r1 and r2:                                # 견본 숫자(10000, 5000) 지우기
        ws.Range(f"E{r1}:I{r2 - 2}").ClearContents()
        ws.Range(f"K{r1}:V{r2 - 2}").ClearContents()
    rr = _find(v, 0, lambda t: t == "구분", (r2 or 30))
    if rr:
        rates = res.get("interest") or [None] * 4
        ws.Range(f"A{rr + 1}").Value = ty - 1
        ws.Range(f"A{rr + 2}").Value = ty
        for addr, val in zip((f"B{rr + 1}", f"C{rr + 1}", f"B{rr + 2}", f"C{rr + 2}"), rates):
            if val is not None:
                ws.Range(addr).Value = val
    if fields.get("무이자") != "Y":
        ws.Visible = HIDDEN


def link_interest(bs, di, binfo):
    """27년 예산 대납이자 줄의 월 금액을 대납이자 시트 '계'(백만원)에 연결 — 대출잔액을 채우면 바로 반영"""
    v = _vals(di, "W", 45)
    r2 = _find(v, 0, lambda t: t == "대출이자")
    rs = _find(v, 1, lambda t: t == "계", r2 or 0) if r2 else None
    hdr = _find(v, 10, lambda t: t == "1월")
    if not (rs and hdr):
        return
    c0 = next(j for j, x in enumerate(v[hdr - 1]) if _txt(x) == "1월") + 1
    for g, s0, s1 in binfo["groups"]:
        if g["sub"] == "대납이자":
            write_block(bs, s0, [[f"='{di.Name}'!{L(c0 + j)}{rs}*1000000" for j in range(12)]], "J", "U")


# ── 입력값 보관(아주 숨김 시트) ─────────────────────────────────────
def store_inputs(wb, data, year):
    try:
        old = wb.Worksheets(DATA_SHEET)
        old.Visible = VISIBLE
        old.Delete()
    except Exception:
        pass
    last = wb.Worksheets(wb.Worksheets.Count)
    ws = wb.Worksheets.Add(None, last)
    ws.Name = DATA_SHEET
    ws.Range("A1").Value = "판관비 예산안 자동작성 입력값 — 화면에서 다시 불러올 때 씁니다(지우지 마세요)."
    s = json.dumps(dict(year=year, data=data), ensure_ascii=False)
    chunks = [s[i:i + 30000] for i in range(0, len(s), 30000)] or [""]
    ws.Range("A2").Value = len(chunks)
    for i, ch in enumerate(chunks):
        ws.Range(f"A{3 + i}").NumberFormat = "@"
        ws.Range(f"A{3 + i}").Value = ch
    ws.Visible = VERY_HIDDEN


# ── 전체 ─────────────────────────────────────────────────────────
def build_report(xl, form_path, out_path, res, data, preview_dir=None):
    shutil.copyfile(form_path, out_path)
    wb = xl.Workbooks.Open(os.path.abspath(out_path), 0, False)
    try:
        fields = data.get("fields") or {}
        types = [t for t in (data.get("types") or []) if t and (t[0] not in (None, "") or t[1] not in (None, ""))]
        ty = int(res["year"])
        bname = next(ws.Name for ws in wb.Worksheets if re.fullmatch(r"\d{2}년 예산", ws.Name))
        fy = 2000 + int(bname[:2])
        shift_years(wb, fy, ty)
        bs = wb.Worksheets(bname)
        new_name = f"{ty % 100:02d}년 예산"
        if bname != new_name:
            bs.Name = new_name
        site = fields.get("현장명") or "현장명"
        d = fields.get("작성일")
        try:
            dt = datetime.datetime.strptime(str(d)[:10], "%Y-%m-%d") if d else datetime.datetime.now()
        except ValueError:
            dt = datetime.datetime.now()
        date_txt = f"( {dt:%Y. %m. %d} )"
        title = f"[ {site} ] {ty}년 판관비 예산안"
        open_idx = None
        o = fields.get("오픈") or [None, None]
        if o[0] and o[1]:
            open_idx = int(o[0]) * 12 + int(o[1])
        note = " 신규현장" if open_idx and ty * 12 + 1 <= open_idx <= ty * 12 + 12 else ""
        binfo = fill_budget(bs, res, title, date_txt, site, keep_interest=fields.get("무이자") == "Y")
        fill_report(wb.Worksheets("예산안 보고서"), new_name, binfo, title, date_txt, site, note)
        basic = fill_basic(wb.Worksheets("기본정보"), fields, types)
        fill_rates(wb.Worksheets("분양율,입주율"), res, basic["total"])
        fill_agency(wb.Worksheets("분양대행 및 촉진수수료"), res, types)
        fill_vacancy(wb.Worksheets("공가세대관리비"), res, types)
        fill_interest(wb.Worksheets("대납이자"), res, fields, ty)
        if fields.get("무이자") == "Y":
            link_interest(bs, wb.Worksheets("대납이자"), binfo)
        for nm in ("프로젝트(수주)", "수주실패비"):
            try:
                wb.Worksheets(nm).Visible = HIDDEN
            except Exception:
                pass
        store_inputs(wb, data, ty)
        rp = wb.Worksheets("예산안 보고서")
        rp.Activate()
        try:
            rp.Range("A1").Select()
        except Exception:
            pass
        xl.CalculateFull()
        wb.Save()
        previews = []
        if preview_dir:
            previews = _previews(wb, rp, bs, binfo, preview_dir)
        return dict(previews=previews, budget_sheet=new_name)
    finally:
        wb.Close(False)


def _previews(wb, rp, bs, binfo, outdir):
    """보고서·예산 시트 첫 블록만 그림으로(미리보기용 인쇄 설정은 저장하지 않음)"""
    import pymupdf
    os.makedirs(outdir, exist_ok=True)
    out = []
    v = _vals(rp, "P")
    hr = _find(v, 1, lambda t: t == "항목")
    tot = _find(v, 0, lambda t: t.startswith("합"), hr)
    jobs = [(rp, f"$A$1:$P${tot}", "보고서"), (bs, f"$A$1:$U${binfo['total']}", "예산")]
    for attempt in range(2):                   # 인쇄 설정이 가끔 일시적으로 실패 → 한 번 더
        jobs = [j for j in jobs if not any(o.startswith(j[2] + "_") for o in out)]
        if not jobs:
            break
        if attempt:
            import time
            time.sleep(1.5)
        out += _export(jobs, outdir, pymupdf)
    return out


def _export(jobs, outdir, pymupdf):
    out = []
    for ws, area, name in jobs:
        try:
            ws.Activate()
            ps = ws.PageSetup
            ps.PrintArea = area
            ps.Orientation = 2
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = False
            pdf = os.path.join(outdir, f"{name}.pdf")
            ws.ExportAsFixedFormat(0, pdf)
            doc = pymupdf.open(pdf)
            for i, pg in enumerate(doc):
                if i >= 4:
                    break
                fn = os.path.join(outdir, f"{name}_{i + 1}.png")
                pg.get_pixmap(matrix=pymupdf.Matrix(1.6, 1.6)).save(fn)
                out.append(os.path.basename(fn))
            doc.close()
        except Exception as e:
            import traceback
            ERRORS.append(f'{name}: {e!r}\n{traceback.format_exc()}')
    return out
