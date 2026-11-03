# -*- coding: utf-8 -*-
"""PM팀 판관비 예산기준(안) 엑셀을 읽어 '기준단가' 시트 값을 갱신한다.

기준안은 해마다 " · 항목 : 금액" 형태의 글로 되어 있고, 바뀐 값은 "이전값 → 새값" 으로 적힌다.
'기준단가' 시트의 숨긴 열(L~R)에 단가마다 '어느 줄의 어느 숫자'인지 찾는 규칙이 적혀 있어,
기준안 문구가 바뀌면 그 칸만 고치면 된다.

찾는 순서
  1) 구분(분양/입주 · 분양사무실) → 항목(B열) → 줄찾기 정규식(세목|기준안|비고 를 '|'로 이은 글에 적용)
  2) 같은 줄에서 찾힌 단가들의 '위치' 글자로 줄을 조각낸다(예: 수도권 / 비수도권)
  3) 조각 안에 '→' 가 있으면 마지막 '→' 뒤만 본다(바뀐 값)
  4) 단위가 붙은 숫자(예: 3,000천원, 350원, 30%, 2명) 중 순번째 값을 쓴다(-1 = 마지막)
"""
import re
import datetime

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

# 기준단가 시트 열 위치 (1부터)
COL = dict(no=1, sec=2, item=3, sub=4, desc=5, val=6, unit=7, rule=8, src_text=9, result=10,
           name=11, f_sec=12, f_item=13, f_line=14, f_clause=15, f_unit=16, f_nth=17, f_src=18)
FIRST_ROW = 6          # 기준단가 표 첫 행
NUM = r"(\d[\d,]*(?:\.\d+)?)"

FILL_VALUE = PatternFill("solid", fgColor="FFF2CC")     # 평소 값 칸
FILL_CHANGED = PatternFill("solid", fgColor="FFC000")   # 이번에 바뀐 값
FILL_CHECK = PatternFill("solid", fgColor="FF9999")     # 확인 필요


def _num(s):
    v = float(s.replace(",", ""))
    return int(v) if v == int(v) else v


def _clean(v):
    if v is None:
        return ""
    return str(v).replace("\r", "").strip("\n")


def read_kijun(path):
    """기준안 엑셀 → (줄 목록, 작성일, 시트이름)"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = None
    for name in wb.sheetnames:
        if "예산기준" in name.replace(" ", ""):
            ws = wb[name]
            break
    if ws is None:
        ws = wb.worksheets[0]

    # 병합 칸 값을 채움 (A 구분, B 항목, C 세목, E 비고)
    fill = {}
    for rng in ws.merged_cells.ranges:
        v = ws.cell(rng.min_row, rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                fill[(r, c)] = v

    def get(r, c):
        v = ws.cell(r, c).value
        if v is None:
            v = fill.get((r, c))
        return _clean(v)

    written = None
    for r in range(1, 6):
        for c in range(1, 6):
            m = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일", get(r, c))
            if m:
                written = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                break
        if written:
            break

    head_row = None
    for r in range(1, min(ws.max_row, 15) + 1):
        if get(r, 4).replace(" ", "") == "기준안":
            head_row = r
            break
    lines = []
    for r in range((head_row or 4) + 1, ws.max_row + 1):
        a, b, c, d, e = (get(r, k) for k in range(1, 6))
        if not d:
            continue
        lines.append(dict(row=r, sec=a, item=b, sub=c, text=d, note=e))
    return lines, written, ws.title


def _sec_ok(spec_sec, line_sec):
    s = line_sec.replace("\n", "").replace(" ", "")
    if spec_sec == "사무실":
        return "사무실" in s
    return "사무실" not in s


def _tokens(text, unit):
    """단위가 붙은 숫자 목록. unit 은 정규식(천원, %, 명 …), '원' 은 '천원' 과 구별한다."""
    if unit == "원":
        pat = NUM + r"\s*원"
    else:
        pat = NUM + r"\s*(?:" + unit + ")"
    return [_num(m.group(1)) for m in re.finditer(pat, text)]


def _hay(ln):
    return "|".join(x.replace("\n", " ") for x in (ln["sub"], ln["text"], ln["note"]))


def find_values(lines, specs):
    """specs: [{key, sec, item, line, clause, unit, nth, src}] → ({key: (값|None, 줄, 사유)}, 쓴 줄 번호들, {key: 줄})"""
    picked = {}
    for sp in specs:
        if not sp.get("line"):
            continue
        item_re = re.compile(sp["item"]) if sp.get("item") else None
        line_re = re.compile(sp["line"])
        for ln in lines:
            if not _sec_ok(sp["sec"], ln["sec"]):
                continue
            if item_re and not item_re.search(ln["item"].replace("\n", "")):
                continue
            if line_re.search(_hay(ln)):
                picked[sp["key"]] = ln
                break

    by_line = {}
    for sp in specs:
        ln = picked.get(sp["key"])
        if ln is not None:
            by_line.setdefault(ln["row"], []).append(sp)

    result = {}
    for sp in specs:
        key = sp["key"]
        if not sp.get("line"):
            continue
        ln = picked.get(key)
        if ln is None:
            result[key] = (None, "", "기준안에서 해당 줄을 못 찾음")
            continue
        text = ln["note"] if sp.get("src") == "비고" else ln["text"]
        start = 0
        if sp.get("clause"):
            m = re.search(sp["clause"], text)
            if not m and sp.get("src") != "비고":
                alt = ln["note"]
                m2 = re.search(sp["clause"], alt)
                if m2:
                    text, m = alt, m2
            if not m:
                result[key] = (None, ln["text"], "줄은 찾았으나 위치 글자가 없음")
                continue
            start = m.start()
        ends = []
        for other in by_line.get(ln["row"], []):
            if other["key"] == key or not other.get("clause") or other.get("src") == "비고":
                continue
            m = re.search(other["clause"], text)
            if m and m.start() > start:
                ends.append(m.start())
        piece = text[start:min(ends) if ends else len(text)]
        # 칸 안 줄바꿈 뒤 '→ 입주 첫달은…' 같은 설명 줄은 값이 아님 → 첫 줄만 본다
        piece = piece.split("\n")[0]
        if "→" in piece:
            piece = piece.split("→")[-1]
        toks = _tokens(piece, sp["unit"])
        nth = int(sp.get("nth") or 1)
        if not toks or (nth > 0 and len(toks) < nth):
            result[key] = (None, ln["text"], "숫자를 못 찾음")
            continue
        result[key] = (toks[nth - 1] if nth > 0 else toks[nth], ln["text"], "")
    return result, {ln["row"] for ln in picked.values()}, picked


def specs_from_sheet(ws):
    specs = []
    for r in range(FIRST_ROW, ws.max_row + 1):
        key = ws.cell(r, COL["name"]).value
        if not key:
            continue
        specs.append(dict(
            row=r, key=key,
            sec=ws.cell(r, COL["f_sec"]).value or "분양",
            item=ws.cell(r, COL["f_item"]).value or "",
            line=ws.cell(r, COL["f_line"]).value or "",
            clause=ws.cell(r, COL["f_clause"]).value or "",
            unit=ws.cell(r, COL["f_unit"]).value or "천원",
            nth=ws.cell(r, COL["f_nth"]).value or 1,
            src=ws.cell(r, COL["f_src"]).value or "",
        ))
    return specs


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return f"{v:,.2f}".rstrip("0").rstrip(".")
    return f"{v:,}"


def apply_to_workbook(wb, kijun_path, force=False):
    """openpyxl 통합문서의 '기준단가' 시트에 기준안 값을 넣고 '기준안 반영결과' 시트를 새로 쓴다."""
    ws = wb["기준단가"]
    lines, written, sheet_title = read_kijun(kijun_path)
    specs = specs_from_sheet(ws)
    found, used_rows, picked = find_values(lines, specs)

    f_res = lambda color, bold=False: Font(name="맑은 고딕", size=9, color=color, bold=bold)
    report = []
    for sp in specs:
        if sp["key"] not in found:          # 기준안에서 읽지 않는 고정값
            continue
        r = sp["row"]
        cell = ws.cell(r, COL["val"])
        old = cell.value
        val, text, why = found[sp["key"]]
        desc = f'{ws.cell(r, COL["item"]).value} / {ws.cell(r, COL["desc"]).value}'
        ws.cell(r, COL["src_text"]).value = " ".join(text.split()) if text else None
        res = ws.cell(r, COL["result"])
        ln = picked.get(sp["key"])
        if ln is not None and ln["sub"]:
            new_sub = " ".join(ln["sub"].split())
            if ws.cell(r, COL["sub"]).value != new_sub:
                report.append(("세목", desc, ws.cell(r, COL["sub"]).value, new_sub, "세목 이름 변경"))
                ws.cell(r, COL["sub"]).value = new_sub
        if val is None:
            res.value = f"확인 필요: {why} (기존값 유지)"
            res.font = f_res("C00000", True)
            cell.fill = FILL_CHECK
            report.append(("확인", desc, old, None, why))
            continue
        if isinstance(old, (int, float)) and abs(float(old) - float(val)) < 1e-9:
            res.value = "동일"
            res.font = f_res("808080")
            cell.fill = FILL_VALUE
            report.append(("동일", desc, old, val, ""))
            continue
        if isinstance(old, (int, float)) and old and not force:
            ratio = float(val) / float(old)
            if ratio > 3 or ratio < 1 / 3:
                res.value = f"확인 필요: 읽은 값 {fmt(val)} 이 기존 {fmt(old)} 과 차이가 커서 반영 안 함"
                res.font = f_res("C00000", True)
                cell.fill = FILL_CHECK
                report.append(("확인", desc, old, val, "읽은 값 차이가 큼(자동 반영 안 함)"))
                continue
        cell.value = val
        cell.fill = FILL_CHANGED
        res.value = f"변경 {fmt(old)} → {fmt(val)}"
        res.font = f_res("0000C0", True)
        report.append(("변경", desc, old, val, ""))

    unused = [ln for ln in lines
              if ln["row"] not in used_rows and re.search(r"\d\s*(천원|원|%)", ln["text"])]

    fname = kijun_path.replace("\\", "/").split("/")[-1]
    ws.cell(2, 3).value = fname
    if written:
        ws.cell(2, 7).value = written
        ws.cell(2, 7).number_format = "yyyy.mm.dd"
        ws.cell(2, 10).value = written.year + 1
    ws.cell(3, 10).value = datetime.datetime.now().replace(microsecond=0)
    ws.cell(3, 10).number_format = "yyyy.mm.dd hh:mm"

    _write_result_sheet(wb, fname, written, report, unused)
    return dict(report=report, unused=unused, written=written, sheet=sheet_title, lines=len(lines))


def _write_result_sheet(wb, fname, written, report, unused):
    name = "기준안 반영결과"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = "FFC000"
    f = lambda **k: Font(name="맑은 고딕", size=k.pop("size", 10), **k)
    ws["A1"] = "기준안 반영 결과"
    ws["A1"].font = f(size=14, bold=True)
    ws["A2"] = f"기준안 파일 : {fname}" + (f"   (작성일 {written:%Y.%m.%d} → {written.year + 1}년 예산)" if written else "")
    ws["A3"] = f"반영 일시 : {datetime.datetime.now():%Y.%m.%d %H:%M}"
    cnt = {k: sum(1 for x in report if x[0] == k) for k in ("변경", "동일", "확인", "세목")}
    ws["A4"] = (f"바뀐 값 {cnt['변경']}개 · 같은 값 {cnt['동일']}개 · 확인 필요 {cnt['확인']}개"
                + (f" · 세목 이름 변경 {cnt['세목']}개" if cnt["세목"] else "")
                + f" · 단가로 읽지 않은 기준안 줄 {len(unused)}개")
    ws["A4"].font = f(bold=True, color="C00000" if cnt["확인"] or unused else "000000")
    r = 6
    for title, kinds in (("1. 확인 필요 (기존 값 유지 — 기준단가 시트에서 직접 확인)", ("확인",)),
                         ("2. 바뀐 값", ("변경", "세목")),):
        ws.cell(r, 1, title).font = f(bold=True)
        r += 1
        for c, h in enumerate(("구분", "항목 / 내용", "기존", "기준안 값", "비고"), start=1):
            ws.cell(r, c, h).font = f(bold=True)
            ws.cell(r, c).fill = PatternFill("solid", fgColor="D9E1F2")
        r += 1
        rows = [x for x in report if x[0] in kinds]
        if not rows:
            ws.cell(r, 2, "없음").font = f(color="808080")
            r += 1
        for kind, desc, old, new, why in rows:
            ws.cell(r, 1, kind)
            ws.cell(r, 2, desc)
            ws.cell(r, 3, fmt(old) if not isinstance(old, str) else old)
            ws.cell(r, 4, fmt(new) if not isinstance(new, str) else new)
            ws.cell(r, 5, why)
            r += 1
        r += 1
    ws.cell(r, 1, "3. 기준안에는 있는데 어느 단가로도 읽지 않은 줄 (새 항목이면 양식에 행 추가가 필요할 수 있음)").font = f(bold=True)
    r += 1
    for c, h in enumerate(("기준안 행", "구분 / 항목", "기준안 글"), start=1):
        ws.cell(r, c, h).font = f(bold=True)
        ws.cell(r, c).fill = PatternFill("solid", fgColor="D9E1F2")
    r += 1
    if not unused:
        ws.cell(r, 2, "없음").font = f(color="808080")
    for ln in unused:
        ws.cell(r, 1, ln["row"])
        ws.cell(r, 2, f'{" ".join(ln["sec"].split())} / {" ".join(ln["item"].split())}')
        ws.cell(r, 3, " ".join(ln["text"].split()))
        r += 1
    for col, w in zip("ABCDE", (10, 46, 16, 16, 60)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=6, max_row=r):
        for c in row:
            if c.font is None or c.font.name != "맑은 고딕":
                c.font = f()
            c.alignment = Alignment(vertical="center", wrap_text=c.column in (2, 3, 5))
