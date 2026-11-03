# -*- coding: utf-8 -*-
"""화면(app\\server.py) ↔ 자동작성 양식 사이의 읽기·쓰기.

  schema(양식)            현장입력 칸 목록(구역·칸·기본값) — build_template.INPUTS 와 양식의 이름(입력_…)으로 만든다
  template_info(양식)     기준안 파일·작성일·예산연도·기준단가 표
  read_inputs(xlsx)       작성한 현장 파일에서 입력값 읽기(불러오기)
  write_inputs(양식, 출력, data)   입력값을 양식 복사본에 쓰기(openpyxl)
  excel_calc(xl, 경로)    Excel 로 다시 계산하고 결과(월별 금액·단계·점검) 읽기, 계산값을 담아 저장
"""
import datetime
import os
import re
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_template as BT     # noqa: E402
import kijun_parser as KP       # noqa: E402

XL_ERR = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!"}


def _ref(wb, name):
    """정의된 이름 → (시트, 셀주소)"""
    dn = wb.defined_names.get(name)
    if dn is None:
        return None
    m = re.match(r"'?(.+?)'?!\$?([A-Z]+)\$?(\d+)$", dn.attr_text)
    if not m:
        return None
    return m.group(1), f"{m.group(2)}{m.group(3)}"


def _cell(wb, name):
    r = _ref(wb, name)
    return wb[r[0]][r[1]] if r else None


def _jsonable(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return v


def _resolve(wb, v):
    """'=기준_…' 같은 기본 수식을 값으로"""
    if isinstance(v, str) and v.startswith("=기준_"):
        c = _cell(wb, v[1:])
        return c.value if c is not None else None
    return v


# 현장 구분별로 보여 줄 칸: new=신규분양, remain=잔여분양(기공급), movein=입주, office=분양사무실 (없으면 모두)
KIND_LABELS = [("new", "신규분양", "올해(또는 내년 1월) 견본주택 오픈"),
               ("remain", "잔여분양", "이미 오픈한 현장 — M/H·조직분양 계속"),
               ("movein", "입주", "사전점검·입주"),
               ("office", "분양사무실", "M/H 철거 뒤 분양사무실 운영")]
_NR, _N, _R, _M, _O = "new remain", "new", "remain", "movein", "office"
FIELD_TAGS = {
    "오픈": _NR, "사전영업": _N, "MH종료": _NR, "목표분양률": _NR, "전년분양": "remain movein", "분양대행": _NR,
    "조직": _NR, "조직시작": _NR, "조직종료": _NR, "조직비율": _NR, "조직단가": _NR,
    "전기미인입": _N, "주차장": _N, "MH중개": _N, "분양가용역": _N, "유니트": _NR, "숙소분양": _N, "PC": _NR,
    "복합기": "new remain movein", "보험월": _NR, "예비비율": _N, "등기월1": _NR, "등기월2": _NR,
    "아르바이트": _N, "카페테리아": _N, "무순위": _N, "계약취소": _N,
    "갑티슈사전": _N, "갑티슈오픈": _N, "물티슈": _N, "종이컵": _N, "크린홈": _N, "마스크": _N, "대형마트": _N,
    "현수막주수": _NR, "현수막감축": _NR, "현수막감축주수": _NR, "오픈선물인원": _N, "청약선물인원": _N,
    "잔여갑티슈": _R, "잔여판촉": _R, "해지세대": _NR, "해지월": _NR, "상담사": _R, "상담사시작": _R,
    "무이자": _NR, "최고이율": _NR,
    "사전점검": _M, "입주": _M, "지정기간": _M, "전년입주": _M, "입주율": _M, "시행사2": _M, "입주촉진": _M,
    "숙소입주": _M, "입주PC": _M, "어린이집": _M, "입주안내": _M,
    "사무실시작": _O, "사무실종료": _O, "통합센터": _O, "사무실복합기": _O, "사무실보안": _O,
}


# 화면에 바로 보이는 칸. 나머지는 기본값(기준안·과거 사례)이 있어 '세부 조정'에 접어 둔다
MAIN_FIELDS = {"현장명", "지역", "급지", "사업구분", "상품", "상가",
               "오픈", "MH종료", "전년분양", "조직시작", "조직종료", "조직비율",
               "사전점검", "입주", "전년입주", "사무실시작", "사무실종료"}


def guess_kinds(fields, year):
    """입력값(날짜)으로 현장 구분 추정 — 예전에 만든 파일을 불러올 때"""
    def idx(v):
        try:
            return int(v[0]) * 12 + int(v[1])
        except Exception:
            return None
    out = []
    o = idx(fields.get("오픈") or [None, None])
    if o is not None and year:
        out.append("new" if o >= year * 12 else "remain")
    if idx(fields.get("입주") or [None, None]) or idx(fields.get("사전점검") or [None, None]):
        out.append("movein")
    if idx(fields.get("사무실시작") or [None, None]):
        out.append("office")
    return out


class _ModelView:
    """xlcalc.Model 을 openpyxl 통합문서처럼 읽기(이름→값, 시트 셀 값). 웹판에서 양식을 다시 읽지 않으려고"""

    def __init__(self, model):
        self.m = model
        self.sheetnames = model.sheetnames

    def raw(self, sheet, r, c):
        v = self.m.cells.get(sheet, {}).get((r, c))
        if isinstance(v, tuple) and v and v[0] == "F":
            return "=" + v[1]
        return v

    def by_name(self, name):
        nr = self.m.name_ref(name)
        if not nr:
            return None
        return self.raw(nr[0], nr[1], nr[2])

    def max_row(self, sheet):
        return self.m.maxrow.get(sheet, 0)


def _reader(src):
    """경로(openpyxl 로 읽음) 또는 xlcalc.Model → (이름값 함수, 셀값 함수, 시트이름, max_row 함수, 표시 이름)"""
    if hasattr(src, "name_ref"):
        mv = _ModelView(src)
        return mv.by_name, mv.raw, mv.sheetnames, mv.max_row, getattr(src, "name", "template")
    wb = openpyxl.load_workbook(src, data_only=False)

    def by_name(name):
        c = _cell(wb, name)
        return c.value if c is not None else None

    def raw(sheet, r, c):
        return wb[sheet].cell(r, c).value
    return by_name, raw, wb.sheetnames, (lambda s: wb[s].max_row), os.path.basename(src)


def schema(template):
    """현장입력 칸 목록(구역·칸·기본값). template = 양식 경로 또는 xlcalc.Model"""
    by_name, raw, sheetnames, max_row, _ = _reader(template)

    def resolve(v):
        if isinstance(v, str) and v.startswith("=기준_"):
            return by_name(v[1:])
        return v
    sections = []
    cur = None
    for it in BT.INPUTS:
        if it[1] is None or it[1] == "TYPES":
            cur = dict(title=it[0], types=it[1] == "TYPES", fields=[])
            sections.append(cur)
            continue
        key, label, kind, default, unit, help_, opts = it[:7]
        is_def = len(it) > 7 and it[7] == "def"
        f = dict(key=key, label=label, kind=kind, unit=unit, help=help_,
                 options=opts.split(",") if opts else None, auto=is_def,
                 tags=FIELD_TAGS.get(key, "").split() or None, main=key in MAIN_FIELDS)
        if kind == "ym":
            f["default"] = [by_name(f"입력_{key}연"), by_name(f"입력_{key}월")]
        else:
            v = by_name(f"입력_{key}")
            f["default"] = _jsonable(resolve(v if v is not None else default))
        cur["fields"].append(f)
    return dict(sections=sections, ntypes=BT.NTYPES, kinds=[dict(key=k, label=l, help=h) for k, l, h in KIND_LABELS])


def template_info(template):
    """기준안 파일·작성일·예산연도·기준단가 표. template = 양식 경로 또는 xlcalc.Model"""
    by_name, raw, sheetnames, max_row, name = _reader(template)
    ws = "기준단가"
    C = KP.COL
    params = []
    check = 0
    for r in range(KP.FIRST_ROW, max_row(ws) + 1):
        key = raw(ws, r, C["name"])
        if not key:
            sec = raw(ws, r, 1)
            if isinstance(sec, str) and sec.startswith("■"):
                params.append(dict(section=sec[1:].strip()))
            continue
        res = raw(ws, r, C["result"]) or ""
        if str(res).startswith("확인"):
            check += 1
        params.append(dict(key=key, item=raw(ws, r, C["item"]), sub=raw(ws, r, C["sub"]),
                           desc=raw(ws, r, C["desc"]), value=raw(ws, r, C["val"]),
                           unit=raw(ws, r, C["unit"]), rule=raw(ws, r, C["rule"]),
                           src=raw(ws, r, C["src_text"]), result=res))
    budget = next((n for n in sheetnames if re.fullmatch(r"\d{2}년 예산", n)), None)
    serial = lambda v: _jsonable(datetime.datetime(1899, 12, 30) + datetime.timedelta(days=v)) if isinstance(v, (int, float)) else _jsonable(v)
    return dict(name=name, path=template if isinstance(template, str) else None, kijun=raw(ws, 2, 3),
                written=serial(raw(ws, 2, 7)), year=raw(ws, 2, 10), applied=serial(raw(ws, 3, 10)),
                count=sum(1 for p in params if "key" in p), check=check, params=params, budget=budget)


def _type_rows(wb):
    tot = _ref(wb, "입력_총세대")
    row = int(re.sub(r"\D", "", tot[1]))
    return row - BT.NTYPES, row - 1


def read_inputs(path):
    """작성한 현장 파일 → 입력값. 수식(자동 기본값) 칸은 None.
    회사 양식 보고서(화면에서 저장한 파일)는 아주 숨긴 '_자동작성입력' 시트의 값을 읽는다."""
    import json
    wb = openpyxl.load_workbook(path)
    if "_자동작성입력" in wb.sheetnames:
        ws = wb["_자동작성입력"]
        n = int(ws["A2"].value or 0)
        js = "".join(str(ws[f"A{3 + i}"].value or "") for i in range(n))
        obj = json.loads(js)
        data = obj.get("data") or {}
        data["year"] = obj.get("year")
        return data
    if "현장입력" not in wb.sheetnames or wb.defined_names.get("입력_현장명") is None:
        raise ValueError("자동작성으로 만든 파일이 아닙니다(입력값이 들어 있지 않음).")
    fields = {}
    for it in BT.INPUTS:
        if it[1] is None or it[1] == "TYPES":
            continue
        key, kind = it[0], it[2]

        def val(name):
            c = _cell(wb, name)
            if c is None:
                return None
            v = c.value
            if isinstance(v, str) and v.startswith("="):
                return None
            return _jsonable(v)
        if kind == "ym":
            fields[key] = [val(f"입력_{key}연"), val(f"입력_{key}월")]
        else:
            fields[key] = val(f"입력_{key}")
    ws = wb["현장입력"]
    t1, t2 = _type_rows(wb)
    types = []
    for r in range(t1, t2 + 1):
        t, n, a = ws[f"C{r}"].value, ws[f"D{r}"].value, ws[f"E{r}"].value
        if t not in (None, "") or n not in (None, ""):
            types.append([t, n, a])
    rs = wb["분양율,입주율"]

    def months(row):
        out = []
        for i in range(12):
            v = rs.cell(row, 9 + i).value
            out.append(None if v is None or (isinstance(v, str) and v.startswith("=")) else v)
        return out
    monthly = dict(분양=months(8), 입주=months(13),
                   전년분양=[rs.cell(8, 6 + i).value for i in range(3)],
                   전년입주=[rs.cell(13, 6 + i).value for i in range(3)])
    return dict(fields=fields, types=types, monthly=monthly)


def _num(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).replace(",", "").strip()
    if s.endswith("%"):
        return float(s[:-1]) / 100
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except ValueError:
        return v


def _col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def input_items(names_ref, data):
    """화면 입력값 → [(시트, 셀, 값)]. 양식 파일(openpyxl)과 수식 계산기(xlcalc)가 같이 쓴다.
    names_ref(이름) → (시트, 셀) 또는 None. 비운 자동 기본값 칸·안 보낸 칸·비운 월별 칸은 목록에 없다(양식 수식 그대로)."""
    items = []
    fields = data.get("fields") or {}
    for it in BT.INPUTS:
        if it[1] is None or it[1] == "TYPES":
            continue
        key, kind = it[0], it[2]
        is_def = len(it) > 7 and it[7] == "def"
        if key not in fields:          # 화면에서 보내지 않은 칸은 양식 기본값 그대로
            continue
        v = fields.get(key)
        if kind == "ym":
            v = v or [None, None]
            for suffix, x in zip(("연", "월"), v):
                ref = names_ref(f"입력_{key}{suffix}")
                if ref:
                    items.append((ref[0], ref[1], _num(x)))
            continue
        ref = names_ref(f"입력_{key}")
        if not ref:
            continue
        if v in (None, ""):
            if not is_def:
                items.append((ref[0], ref[1], None))
            continue
        if kind in ("int", "num", "pct", "month"):
            val = _num(v)
        elif kind == "date":
            try:
                val = datetime.datetime.strptime(str(v)[:10], "%Y-%m-%d")
            except ValueError:
                val = None
        else:
            val = str(v)
        items.append((ref[0], ref[1], val))
    tot = names_ref("입력_총세대")
    row = int(re.sub(r"\D", "", tot[1]))
    t1, t2 = row - BT.NTYPES, row - 1
    types = data.get("types") or []
    for k, r in enumerate(range(t1, t2 + 1)):
        t = types[k] if k < len(types) else [None, None, None]
        t = (list(t) + [None, None, None])[:3]
        items.append(("현장입력", f"C{r}", None if t[0] in (None, "") else str(t[0])))
        items.append(("현장입력", f"D{r}", _num(t[1])))
        items.append(("현장입력", f"E{r}", _num(t[2])))
    m = data.get("monthly") or {}
    for key, row in (("분양", 8), ("입주", 13)):
        vals = m.get(key) or []
        for i in range(12):
            v = _num(vals[i]) if i < len(vals) else None
            if v is not None:
                items.append(("분양율,입주율", f"{_col(9 + i)}{row}", v))      # 비우면 자동 곡선 수식 그대로
    for key, row in (("전년분양", 8), ("전년입주", 13)):
        vals = m.get(key) or []
        for i in range(3):
            items.append(("분양율,입주율", f"{_col(6 + i)}{row}", _num(vals[i]) if i < len(vals) else None))
    return items


def write_inputs(template, out, data):
    """입력값을 양식 복사본에 써서 out 으로 저장. 비운 자동 기본값 칸은 수식을 그대로 둔다."""
    wb = openpyxl.load_workbook(template)
    for sheet, cell, v in input_items(lambda n: _ref(wb, n), data):
        wb[sheet][cell].value = v
    wb.active = wb.sheetnames.index("현장입력")
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)


def calc_model(model, data):
    """xlcalc.Model(계산용 양식)에 입력값을 넣고 계산 → excel_calc 와 같은 결과 구조 (Excel 없이)"""
    import xlcalc
    model.reset()

    def names_ref(name):
        nr = model.name_ref(name)
        return (nr[0], f"{_col(nr[2])}{nr[1]}") if nr else None
    model.set_many(input_items(names_ref, data))
    errors = []
    for sheet, cells in model.cells.items():
        for (r, c), raw in cells.items():
            if isinstance(raw, tuple) and raw and raw[0] == "F":
                v = model.cell(sheet, r, c)
                if isinstance(v, xlcalc.XLError):
                    errors.append(f"{sheet}!R{r}C{c} {v.code}")

    def conv(v):
        return v.code if isinstance(v, xlcalc.XLError) else v

    def get2d(sheet, a1):
        return [[conv(x) for x in row] for row in model.values(sheet, a1)]

    def name_value(name):
        return conv(model.name_value(name))
    return read_results_from(model.sheetnames, get2d, name_value, lambda s: model.maxrow.get(s, 0), errors,
                             explain=model.explain)


def attach_formulas(model, data, res):
    """Excel 로 계산한 결과(res)에 수식 계산기(xlcalc)로 푼 월별 산식(res['formulas'])을 붙인다 — 보고서 월별 칸용.
    두 결과의 구조(묶음·줄 수·줄 합계)가 다르면 붙이지 않는다(값만 쓴다). 붙였으면 True."""
    m = calc_model(model, data)
    fx = m.get("formulas")
    if not fx or len(m["groups"]) != len(res.get("groups") or []):
        return False
    for g1, g2, f in zip(m["groups"], res["groups"], fx):
        if len(g1["rows"]) != len(g2["rows"]) or len(f) != len(g2["rows"]):
            return False
        for r1, r2 in zip(g1["rows"], g2["rows"]):
            if abs(float(r1["total"] or 0) - float(r2["total"] or 0)) > 1:
                return False
    res["formulas"] = fx
    return True


def _v(x):
    if isinstance(x, int) and x in XL_ERR:
        return XL_ERR[x]
    return x


def excel_calc(xl, path):
    """Excel 로 열어 다시 계산 → 결과 읽기 → 계산값을 담아 저장"""
    wb = xl.Workbooks.Open(os.path.abspath(path), 0, False)
    try:
        xl.CalculateFull()
        res = read_results(wb)
        wb.Save()
    finally:
        wb.Close(False)
    return res


def read_results(wb):
    """Excel(COM)로 연 계산 파일 → 결과. 값 읽기는 read_results_from 과 같은 규칙"""
    errors = []
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        vals = ur.Value
        if vals is None:
            continue
        if not isinstance(vals, tuple):
            vals = ((vals,),)
        for i, row in enumerate(vals):
            for j, v in enumerate(row):
                if isinstance(v, int) and v in XL_ERR:
                    errors.append(f"{ws.Name}!R{ur.Row + i}C{ur.Column + j} {XL_ERR[v]}")

    def get2d(sheet, a1):
        v = wb.Worksheets(sheet).Range(a1).Value
        if not isinstance(v, tuple):
            v = ((v,),)
        return [[_v(x) for x in row] for row in v]

    def name_value(name):
        try:
            return _v(wb.Names(name).RefersToRange.Value)
        except Exception:
            return None

    def last_row(sheet):
        ur = wb.Worksheets(sheet).UsedRange
        return ur.Row + ur.Rows.Count
    return read_results_from([ws.Name for ws in wb.Worksheets], get2d, name_value, last_row, errors)


def read_results_from(sheetnames, get2d, name_value, last_row, errors, explain=None):
    """get2d(시트, 'A1:U300') → 2차원 값(오류는 '#…' 글자), name_value(이름) → 값, last_row(시트) → 마지막 행.
    explain(시트, 행, 열) → 그 칸의 산식 글(없으면 None)을 주면 결과에 formulas(묶음별·줄별 12개월 산식)를 함께 담는다."""
    bname = next(n for n in sheetnames if re.fullmatch(r"\d{2}년 예산", n))
    last = last_row(bname)
    data = get2d(bname, f"A1:U{last}")
    stage = [_v(x) or "" for x in data[4][9:21]]
    groups, cur = [], None
    fxs, cur_fx = [], None
    total = dict(total=0, months=[0] * 12)
    for i, row in enumerate(data, start=1):
        if i < 7:
            continue
        a, item, sub, g, tot, note = row[0], row[1], row[3], row[6], row[5], row[7]
        months = [_v(x) if x is not None else 0 for x in row[9:21]]
        if isinstance(a, str) and a.startswith("합"):
            total = dict(total=_v(tot) or 0, months=months)
            break
        if item:
            cur = dict(item=item, sub=sub or "", rows=[], total=0, months=[0] * 12)
            groups.append(cur)
            cur_fx = []
            fxs.append(cur_fx)
        if cur is None:
            continue
        if isinstance(sub, str) and sub.replace(" ", "") == "소계":
            cur["total"] = _v(tot) or 0
            cur["months"] = months
            cur = None
            cur_fx = None
            continue
        manual = note == "직접입력"
        if manual and not any(isinstance(m, (int, float)) and m for m in months) and not g:
            continue
        cur["rows"].append(dict(g=(g or "").strip(), total=_v(tot) or 0, months=months, manual=manual))
        if explain is not None:
            fx = None if manual else [explain(bname, i, 10 + j) for j in range(12)]
            cur_fx.append(fx if fx and any(fx) else None)
    rs = get2d("분양율,입주율", "E8:T16")          # 행 8~16, 열 E(5)~T(20)

    def rr(row, c0=9, n=12):
        return [rs[row - 8][c0 - 5 + i] for i in range(n)]
    rates = dict(분양=rr(8), 분양누계=rr(9), 입주=rr(13), 미입주=rr(15), 분양율=rr(11), 입주율=rr(16),
                 전년분양=rr(8, 5, 4), 전년입주=rr(13, 5, 4))
    checks = [row[0] for row in get2d("현장입력", "I14:I22") if row[0]]
    # 보고서(회사 양식) 작성용: 계산 시트의 월별 행과 단가
    calc = {}
    for row in get2d("계산", "B1:V220"):
        key = row[1]
        if key in ("대행세대", "조직세대", "MGM", "미입주", "분양", "입주", "k", "mk"):
            calc[key] = [_v(x) or 0 for x in row[8:20]]
    price = dict(분양대행=name_value("기준_분양대행"), 조직=name_value("계산_조직단가"), MGM=name_value("기준_MGM"),
                 공가=name_value("기준_공가관리비"))
    di = get2d("대납이자", "B36:C37")
    rates_int = [di[0][0], di[0][1], di[1][0], di[1][1]]
    out = dict(sheet=bname, year=int("20" + bname[:2]), stage=stage, groups=groups, total=total,
               rates=rates, checks=checks, errors=errors[:30], nerrors=len(errors),
               calc=calc, price=price, interest=rates_int)
    if explain is not None:
        out["formulas"] = fxs
    return out
