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
    "복합기": "new remain movein", "보험월": _NR, "특공세대": _N, "예비비율": _N, "등기월1": _NR, "등기월2": _NR,
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


def schema(template):
    wb = openpyxl.load_workbook(template)
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
            a, b = _cell(wb, f"입력_{key}연"), _cell(wb, f"입력_{key}월")
            f["default"] = [a.value if a is not None else None, b.value if b is not None else None]
        else:
            c = _cell(wb, f"입력_{key}")
            v = c.value if c is not None else default
            f["default"] = _jsonable(_resolve(wb, v))
        cur["fields"].append(f)
    return dict(sections=sections, ntypes=BT.NTYPES, kinds=[dict(key=k, label=l, help=h) for k, l, h in KIND_LABELS])


def template_info(template):
    wb = openpyxl.load_workbook(template, data_only=False)
    ws = wb["기준단가"]
    C = KP.COL
    params = []
    check = 0
    for r in range(KP.FIRST_ROW, ws.max_row + 1):
        key = ws.cell(r, C["name"]).value
        if not key:
            sec = ws.cell(r, 1).value
            if isinstance(sec, str) and sec.startswith("■"):
                params.append(dict(section=sec[1:].strip()))
            continue
        res = ws.cell(r, C["result"]).value or ""
        if str(res).startswith("확인"):
            check += 1
        params.append(dict(key=key, item=ws.cell(r, C["item"]).value, sub=ws.cell(r, C["sub"]).value,
                           desc=ws.cell(r, C["desc"]).value, value=ws.cell(r, C["val"]).value,
                           unit=ws.cell(r, C["unit"]).value, rule=ws.cell(r, C["rule"]).value,
                           src=ws.cell(r, C["src_text"]).value, result=res))
    written = ws["G2"].value
    applied = ws["J3"].value
    budget = next((n for n in wb.sheetnames if re.fullmatch(r"\d{2}년 예산", n)), None)
    return dict(name=os.path.basename(template), path=template, kijun=ws["C2"].value,
                written=_jsonable(written), year=ws["J2"].value, applied=_jsonable(applied),
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


def write_inputs(template, out, data):
    """입력값을 양식 복사본에 써서 out 으로 저장. 비운 자동 기본값 칸은 수식을 그대로 둔다."""
    wb = openpyxl.load_workbook(template)
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
                c = _cell(wb, f"입력_{key}{suffix}")
                if c is not None:
                    c.value = _num(x)
            continue
        c = _cell(wb, f"입력_{key}")
        if c is None:
            continue
        if v in (None, ""):
            if not is_def:
                c.value = None
            continue
        if kind in ("int", "num", "pct", "month"):
            c.value = _num(v)
        elif kind == "date":
            try:
                c.value = datetime.datetime.strptime(str(v)[:10], "%Y-%m-%d")
            except ValueError:
                c.value = None
        else:
            c.value = str(v)
    ws = wb["현장입력"]
    t1, t2 = _type_rows(wb)
    types = data.get("types") or []
    for k, r in enumerate(range(t1, t2 + 1)):
        t = types[k] if k < len(types) else [None, None, None]
        t = (list(t) + [None, None, None])[:3]
        ws[f"C{r}"].value = None if t[0] in (None, "") else str(t[0])
        ws[f"D{r}"].value = _num(t[1])
        ws[f"E{r}"].value = _num(t[2])
    rs = wb["분양율,입주율"]
    m = data.get("monthly") or {}
    for key, row in (("분양", 8), ("입주", 13)):
        vals = m.get(key) or []
        for i in range(12):
            v = _num(vals[i]) if i < len(vals) else None
            if v is not None:
                rs.cell(row, 9 + i).value = v          # 비우면 자동 곡선 수식 그대로
    for key, row in (("전년분양", 8), ("전년입주", 13)):
        vals = m.get(key) or []
        for i in range(3):
            rs.cell(row, 6 + i).value = _num(vals[i]) if i < len(vals) else None
    wb.active = wb.sheetnames.index("현장입력")
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)


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
    bname = next(ws.Name for ws in wb.Worksheets if re.fullmatch(r"\d{2}년 예산", ws.Name))
    b = wb.Worksheets(bname)
    last = b.UsedRange.Row + b.UsedRange.Rows.Count
    data = b.Range(f"A1:U{last}").Value
    stage = [_v(x) or "" for x in data[4][9:21]]
    groups, cur = [], None
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
        if cur is None:
            continue
        if isinstance(sub, str) and sub.replace(" ", "") == "소계":
            cur["total"] = _v(tot) or 0
            cur["months"] = months
            cur = None
            continue
        manual = note == "직접입력"
        if manual and not any(isinstance(m, (int, float)) and m for m in months) and not g:
            continue
        cur["rows"].append(dict(g=(g or "").strip(), total=_v(tot) or 0, months=months, manual=manual))
    rs = wb.Worksheets("분양율,입주율")

    def rr(row, c0=9, n=12):
        return [_v(x) for x in rs.Range(rs.Cells(row, c0), rs.Cells(row, c0 + n - 1)).Value[0]]
    rates = dict(분양=rr(8), 분양누계=rr(9), 입주=rr(13), 미입주=rr(15), 분양율=rr(11), 입주율=rr(16),
                 전년분양=rr(8, 5, 4), 전년입주=rr(13, 5, 4))
    ci = wb.Worksheets("현장입력")
    checks = [c for c in (ci.Range(f"I{r}").Value for r in range(14, 23)) if c]
    # 보고서(회사 양식) 작성용: 계산 시트의 월별 행과 단가
    calc = {}
    for row in wb.Worksheets("계산").Range("B1:V220").Value:
        key = row[1]
        if key in ("대행세대", "조직세대", "MGM", "미입주", "분양", "입주", "k", "mk"):
            calc[key] = [_v(x) or 0 for x in row[8:20]]

    def nm(name):
        try:
            return _v(wb.Names(name).RefersToRange.Value)
        except Exception:
            return None
    price = dict(분양대행=nm("기준_분양대행"), 조직=nm("계산_조직단가"), MGM=nm("기준_MGM"), 공가=nm("기준_공가관리비"))
    di = wb.Worksheets("대납이자")
    rates_int = [_v(di.Range(a).Value) for a in ("B36", "C36", "B37", "C37")]
    return dict(sheet=bname, year=int("20" + bname[:2]), stage=stage, groups=groups, total=total,
                rates=rates, checks=checks, errors=errors[:30], nerrors=len(errors),
                calc=calc, price=price, interest=rates_int)
