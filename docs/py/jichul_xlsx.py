# -*- coding: utf-8 -*-
"""지출계획서 엑셀 작성(openpyxl) — 과거 PM팀 지출계획서 양식과 같은 모양(굴림, 회색 머리글, 결재란, 월별 칸은 산식).

  build_xlsx(res, data, year=None) → bytes     jichul.calc() 결과를 양식에 채운 xlsx (다시 불러올 입력값은 숨긴 시트에)
  read_inputs(path 또는 bytes) → dict|None     저장한 파일 → {"year":…, "data":…}

양식 구조(열): A 항목 · B 세목 · C 합계 · D~ 월별 · 그 다음 '비고' — 비고는 결재자 수+1 개의 작은 열을 합쳐 쓴다.
결재란(제목 아래 오른쪽)은 그 작은 열들로 그린다(과거 파일은 엑셀 '카메라' 그림이라 자동 생성이 안 되므로 실제 칸으로).
"""
import io
import json

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

DATA_SHEET = "_자동작성입력"
SHEET = "지출계획서"
NF = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-'
FONT = "굴림"
GRAY = "D9D9D9"
LIGHT = "F2F2F2"
NOTE_WIDTH = 65.75          # 과거 양식의 비고 열 너비(글자 수)
MARK_WIDTH = 4.25           # 결재란 '결재' 세로 글 칸


def _side(style):
    return Side(style=style) if style else Side()


def _border(left=None, right=None, top=None, bottom=None):
    return Border(left=_side(left), right=_side(right), top=_side(top), bottom=_side(bottom))


def _font(size=11, bold=False):
    return Font(name=FONT, size=size, bold=bold)


def _fill(color):
    return PatternFill("solid", fgColor=color) if color else PatternFill()


def _al(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


class _Form:
    def __init__(self, ws, n_months, n_appr, appr_width):
        self.ws = ws
        self.n = n_months
        self.cA, self.cB, self.cC = 1, 2, 3
        self.cM = [4 + i for i in range(n_months)]            # 월별 열
        self.cMark = 4 + n_months                               # 결재 '결재' 칸
        self.cAp = [self.cMark + 1 + i for i in range(n_appr)]  # 결재자 칸들
        self.cN1, self.cN2 = self.cMark, self.cAp[-1]           # 비고 = cMark ~ 마지막 결재자 칸
        self.last = self.cN2
        ws.column_dimensions["A"].width = 11.25
        ws.column_dimensions["B"].width = 14.25
        ws.column_dimensions["C"].width = 14.25
        for c in self.cM:
            ws.column_dimensions[L(c)].width = 14.25
        ws.column_dimensions[L(self.cMark)].width = MARK_WIDTH
        for c in self.cAp:
            ws.column_dimensions[L(c)].width = appr_width
        for c in range(self.last + 1, self.last + 4):
            ws.column_dimensions[L(c)].width = 3

    def put(self, r, c, v=None, font=None, fill=None, border=None, al=None, nf=None):
        cell = self.ws.cell(r, c)
        if v is not None:
            cell.value = v
        cell.font = font or _font()
        if fill:
            cell.fill = fill
        if border:
            cell.border = border
        cell.alignment = al or _al()
        if nf:
            cell.number_format = nf
        return cell

    def merge(self, r1, c1, r2, c2):
        if (r1, c1) != (r2, c2):
            self.ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    def note_merge(self, r):
        self.merge(r, self.cN1, r, self.cN2)


def _approval_box(fm, ap, r0):
    """결재란: r0 머리글(담당/검토/승인) · r0+1 서명 · r0+2 날짜(/) · r0+3 직위(이름). 같은 역할이 이어지면 머리글 합침"""
    ws = fm.ws
    lst = ap["list"]
    bd = _border("thin", "thin", "thin", "thin")
    fb = _font(10, True)
    gray = _fill(GRAY)
    for r in range(r0, r0 + 4):
        fm.put(r, fm.cMark, None, fb, gray, bd, _al(wrap=True))
    fm.put(r0, fm.cMark, "결\n\n재", fb, gray, bd, _al(wrap=True))
    fm.merge(r0, fm.cMark, r0 + 3, fm.cMark)
    i = 0
    while i < len(lst):
        j = i
        while j + 1 < len(lst) and lst[j + 1]["role"] == lst[i]["role"]:
            j += 1
        role = lst[i]["role"]
        label = {"담당": "담  당", "검토": "검   토", "승인": "승  인", "협조": "협  조"}.get(role, role)
        for k in range(i, j + 1):
            fm.put(r0, fm.cAp[k], None, fb, gray, bd)
        fm.put(r0, fm.cAp[i], label, fb, gray, bd)
        fm.merge(r0, fm.cAp[i], r0, fm.cAp[j])
        i = j + 1
    for k, a in enumerate(lst):
        c = fm.cAp[k]
        fm.put(r0 + 1, c, None, fb, None, bd)
        fm.put(r0 + 2, c, "/", fb, None, bd)
        txt = (a.get("pos") or "") + ((" " + a["name"]) if a.get("name") else "")
        fm.put(r0 + 3, c, txt.strip() or None, fb, gray, bd, _al(wrap=True))


def build_xlsx(res, data, year=None):
    """계산 결과(res)와 입력값(data) → 지출계획서 xlsx(bytes)"""
    n = res["n"]
    ap = res.get("approvers") or dict(list=[], width=None, height=None)
    lst = ap.get("list") or [dict(role="담당", pos="과장"), dict(role="검토", pos="부장"), dict(role="검토", pos="임원"),
                             dict(role="승인", pos="실장")]
    ap = dict(ap, list=lst)
    nA = len(lst)
    width = ap.get("width") or round((NOTE_WIDTH - MARK_WIDTH) / nA, 2)
    sign_h = ap.get("height") or 48
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    fm = _Form(ws, n, nA, width)
    last = fm.last
    gray = _fill(GRAY)

    # 제목·작성일
    rh = ws.row_dimensions
    rh[1].height = 9
    fm.put(2, 1, res["title"], _font(28, True), None, None, _al("center", "center"))
    fm.merge(2, 1, 2, last)
    rh[2].height = 31.5
    fm.put(3, 1, res["date_txt"], _font(14), None, None, _al("center", "center"))
    fm.merge(3, 1, 3, last)
    rh[3].height = 17.25
    rh[4].height = 6
    # 결재란(5~8행)
    _approval_box(fm, ap, 5)
    rh[5].height = 24
    rh[6].height = sign_h
    rh[7].height = 19.5
    rh[8].height = 30 if any(a.get("name") for a in lst) else 18
    # ( 금액 : 원 )
    fm.put(9, fm.cN1, "( 금액 : 원 )", _font(10), None, None, _al("right", "bottom"))
    fm.note_merge(9)
    rh[9].height = 20

    # 머리글(10~11행)
    H = 10
    rh[H].height = 45
    rh[H + 1].height = 45
    hal = _al("center", "center", True)
    fm.put(H, fm.cA, "구      분", _font(), gray, _border("medium", None, "medium", None), hal)
    fm.put(H, fm.cB, None, _font(), gray, _border(None, "thin", "medium", None), hal)
    fm.put(H + 1, fm.cA, None, _font(), gray, _border("medium", None, None, "double"), hal)
    fm.put(H + 1, fm.cB, None, _font(), gray, _border(None, "thin", None, "double"), hal)
    fm.merge(H, fm.cA, H + 1, fm.cB)
    fm.put(H, fm.cC, "합 계", _font(), gray, _border("medium", "medium", "medium", None), hal)
    fm.put(H + 1, fm.cC, None, _font(), gray, _border("medium", "medium", None, "double"), hal)
    fm.merge(H, fm.cC, H + 1, fm.cC)
    for i, c in enumerate(fm.cM):
        left = "medium" if i == 0 else "hair"
        right = "thin" if i == n - 1 else "hair"
        fm.put(H, c, "예 상 비 용   지 출 계 획" if i == 0 else None, _font(), gray, _border(left, right, "medium", "hair"), hal)
        m = res["months"][i]
        txt = f"{m['m']}월" + (("\n" + m["label"]) if m.get("label") else "")
        fm.put(H + 1, c, txt, _font(), gray, _border(left, right, "hair", "double"), hal)
    fm.merge(H, fm.cM[0], H, fm.cM[-1])
    for c in range(fm.cN1, fm.cN2 + 1):
        fm.put(H, c, "비 고" if c == fm.cN1 else None, _font(), gray,
               _border("thin" if c == fm.cN1 else None, "medium" if c == fm.cN2 else None, "medium", None), hal)
        fm.put(H + 1, c, None, _font(), gray,
               _border("thin" if c == fm.cN1 else None, "medium" if c == fm.cN2 else None, None, "double"), hal)
    fm.merge(H, fm.cN1, H + 1, fm.cN2)

    # 본문
    r = H + 2
    sum_refs = []           # 합계 줄에서 더할 행(소계 또는 단일 줄)
    first_body = r
    for gi, g in enumerate(res["groups"]):
        rows = g["rows"]
        g_first = r
        g_rows = len(rows) + (1 if g["subtotal"] else 0)
        for ri, row in enumerate(rows):
            top = "thin" if ri == 0 else "hair"
            is_last = (ri == len(rows) - 1) and not g["subtotal"]
            bottom = "thin" if is_last else "hair"
            rh[r].height = 53.25
            if g["merge"]:
                fm.put(r, fm.cA, g["item"], _font(), gray, _border("medium", None, top, bottom), _al(wrap=True))
                fm.put(r, fm.cB, None, _font(), gray, _border(None, None, top, bottom), _al(wrap=True))
                fm.merge(r, fm.cA, r, fm.cB)
            else:
                fm.put(r, fm.cA, g["item"] if ri == 0 else None, _font(), gray,
                       _border("medium", None, "thin" if ri == 0 else None, None), _al(wrap=True))
                fm.put(r, fm.cB, row["name"], _font(), None, _border("thin", None, top, bottom), _al(wrap=True))
            fm.put(r, fm.cC, f"=SUM({L(fm.cM[0])}{r}:{L(fm.cM[-1])}{r})", _font(), None,
                   _border("medium", "medium", top, bottom), _al("right"), NF)
            for i, c in enumerate(fm.cM):
                cell = row["cells"][i]
                v = f"={cell['f']}" if cell.get("f") else (cell["v"] if cell["v"] else 0)
                fm.put(r, c, v, _font(), None, _border("medium" if i == 0 else "hair", "hair", top, bottom), _al("right"), NF)
            for c in range(fm.cN1, fm.cN2 + 1):
                fm.put(r, c, row["note"] if c == fm.cN1 else None, _font(), None,
                       _border("thin" if c == fm.cN1 else None, "medium" if c == fm.cN2 else None, top, bottom),
                       _al("left", "center", True), "@")
            fm.note_merge(r)
            row["_row"] = r
            r += 1
        if g["subtotal"]:
            rh[r].height = 53.25
            fm.put(r, fm.cA, None, _font(), gray, _border("medium", None, None, "thin"))
            fm.put(r, fm.cB, "소   계", _font(), _fill(LIGHT), _border(None, None, "hair", "thin"))
            fm.put(r, fm.cC, f"=SUM({L(fm.cM[0])}{r}:{L(fm.cM[-1])}{r})", _font(), _fill(LIGHT),
                   _border("medium", "medium", "hair", "thin"), _al("right"), NF)
            r1, r2 = rows[0]["_row"], rows[-1]["_row"]
            for i, c in enumerate(fm.cM):
                fm.put(r, c, f"=SUM({L(c)}{r1}:{L(c)}{r2})", _font(), _fill(LIGHT),
                       _border("medium" if i == 0 else "hair", "hair", "hair", "thin"), _al("right"), NF)
            for c in range(fm.cN1, fm.cN2 + 1):
                fm.put(r, c, None, _font(), _fill(LIGHT),
                       _border("thin" if c == fm.cN1 else None, "medium" if c == fm.cN2 else None, "hair", "thin"))
            fm.note_merge(r)
            sum_refs.append(r)
            r += 1
        else:
            sum_refs.append(rows[0]["_row"])
        if not g["merge"]:
            fm.merge(g_first, fm.cA, g_first + g_rows - 1, fm.cA)
            # 묶음 왼쪽 칸 아래선
            ws.cell(g_first + g_rows - 1, fm.cA).border = _border("medium", None, None, "thin")

    # 합계
    rh[r].height = 60
    fb = _font(11, True)
    fm.put(r, fm.cA, "합      계", fb, gray, _border("medium", None, "double", "medium"))
    fm.put(r, fm.cB, None, fb, gray, _border(None, None, "double", "medium"))
    fm.merge(r, fm.cA, r, fm.cB)
    fm.put(r, fm.cC, "=" + "+".join(f"C{x}" for x in sum_refs) if sum_refs else 0, fb, gray,
           _border("medium", "medium", "double", "medium"), _al("right"), NF)
    for i, c in enumerate(fm.cM):
        fm.put(r, c, "=" + "+".join(f"{L(c)}{x}" for x in sum_refs) if sum_refs else 0, fb, gray,
               _border("medium" if i == 0 else "hair", "hair", "double", "medium"), _al("right"), NF)
    for c in range(fm.cN1, fm.cN2 + 1):
        fm.put(r, c, None, _font(10), gray, _border("thin" if c == fm.cN1 else None, "medium" if c == fm.cN2 else None, "double", "medium"))
    fm.note_merge(r)
    last_row = r

    # 인쇄 설정: A4 세로, 한 쪽에 맞춤(가로·세로), 가운데
    ws.print_area = f"A1:{L(last)}{last_row}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.2
    ws.page_margins.top, ws.page_margins.bottom = 0.55, 0.31
    ws.page_margins.header, ws.page_margins.footer = 0.39, 0.28

    # 다시 불러오기용 입력값
    s = json.dumps(dict(year=year, data=data), ensure_ascii=False)
    chunks = [s[i:i + 30000] for i in range(0, len(s), 30000)] or [""]
    ds = wb.create_sheet(DATA_SHEET)
    ds["A1"] = "지출계획서 자동작성 입력값 — 화면에서 다시 불러올 때 씁니다(지우지 마세요)."
    ds["A2"] = len(chunks)
    for i, ch in enumerate(chunks):
        ds.cell(3 + i, 1).value = ch
    ds.sheet_state = "veryHidden"
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def read_inputs(src):
    """저장한 지출계획서(숨긴 '_자동작성입력' 시트) → dict(year, data). 자동작성 파일이 아니면 None"""
    if isinstance(src, (bytes, bytearray)):
        src = io.BytesIO(src)
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    try:
        if DATA_SHEET not in wb.sheetnames:
            return None
        ws = wb[DATA_SHEET]
        rows = [[c for c in row] for row in ws.iter_rows(min_row=1, max_row=200, max_col=1, values_only=True)]
        n = int((rows[1][0] if len(rows) > 1 else 0) or 0)
        js = "".join(str(rows[2 + i][0] or "") for i in range(n) if 2 + i < len(rows))
        return json.loads(js) if js else None
    finally:
        wb.close()
