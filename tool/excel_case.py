# -*- coding: utf-8 -*-
"""양식에 현장 값을 넣고 Excel 로 계산해 결과를 읽는 시험 도구.

    from excel_case import run_case
    res = run_case(양식경로, inputs={...}, types=[("84A", 343, 112.5), ...], save_as=None)

inputs 의 키는 현장입력 이름(입력_ 뒤 부분). 연·월 입력은 ("오픈", 2027, 3) 처럼 키 '오픈' 에 (연, 월) 튜플.
분양율 시트 월별 값을 직접 줄 때는 inputs["@분양"] = [12개], inputs["@입주"] = [12개].
"""
import os
import sys
import time

import pythoncom
import win32com.client

ERRORS = ("#NAME?", "#VALUE!", "#REF!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!", "#SPILL!", "#CALC!")
XL_ERR = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!"}


def _val(v):
    if isinstance(v, int) and v in XL_ERR:
        return XL_ERR[v]
    return v


def run_case(template, inputs, types, save_as=None, budget_sheet=None, keep_open=False):
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(template), 0, False)
        wb.Application.Calculation = -4135   # 수동
        ws_in = wb.Worksheets("현장입력")
        # 타입표
        t_first = None
        for r in range(1, 80):
            if ws_in.Cells(r, 3).Value == "타입" and ws_in.Cells(r, 4).Value == "세대수":
                t_first = r + 1
                break
        for k in range(10):
            row = t_first + k
            if k < len(types):
                tp = types[k]
                ws_in.Cells(row, 3).Value = str(tp[0])
                ws_in.Cells(row, 4).Value = tp[1]
                ws_in.Cells(row, 5).Value = tp[2] if len(tp) > 2 else None
            else:
                for c in (3, 4, 5):
                    ws_in.Cells(row, c).Value = None
        for key, v in inputs.items():
            if key.startswith("@"):
                continue
            if isinstance(v, tuple):
                wb.Names(f"입력_{key}연").RefersToRange.Value = v[0]
                wb.Names(f"입력_{key}월").RefersToRange.Value = v[1]
            else:
                wb.Names(f"입력_{key}").RefersToRange.Value = v
        rates = wb.Worksheets("분양율,입주율")
        for key, row in (("@분양", 8), ("@입주", 13)):
            if key in inputs:
                for i, v in enumerate(inputs[key]):
                    if v is not None:
                        rates.Cells(row, 9 + i).Value = v
        for key, row in (("@전년분양", 8), ("@전년입주", 13)):
            if key in inputs:       # F~H (10~12월)
                for i, v in enumerate(inputs[key]):
                    rates.Cells(row, 6 + i).Value = v
        xl.CalculateFull()

        # 오류 검사
        errors = []
        for ws in wb.Worksheets:
            ur = ws.UsedRange
            vals = ur.Value
            if vals is None:
                continue
            if not isinstance(vals, tuple):
                vals = ((vals,),)
            r0, c0 = ur.Row, ur.Column
            for i, row in enumerate(vals):
                for j, v in enumerate(row):
                    v = _val(v)
                    if isinstance(v, str) and v in ERRORS:
                        errors.append(f"{ws.Name}!{_addr(r0 + i, c0 + j)} {v} :: {ws.Cells(r0 + i, c0 + j).Formula[:120]}")

        bname = budget_sheet
        if bname is None:
            for ws in wb.Worksheets:
                if ws.Name.endswith("년 예산"):
                    bname = ws.Name
        b = wb.Worksheets(bname)
        rows = []
        last = b.UsedRange.Row + b.UsedRange.Rows.Count
        data = b.Range(f"A1:U{last}").Value
        cur_item = None
        for i, row in enumerate(data, start=1):
            if i < 7:
                continue
            item, sub, g, f_tot = row[1], row[3], row[6], row[5]
            if item:
                cur_item = item
            if isinstance(row[0], str) and "합" in row[0]:
                rows.append(dict(r=i, item="합계", sub="", g="", total=f_tot, months=[_val(x) for x in row[9:21]]))
                break
            rows.append(dict(r=i, item=cur_item, sub=sub, g=g, total=f_tot, months=[_val(x) for x in row[9:21]]))
        stage = [b.Cells(5, 10 + i).Value for i in range(12)]
        rates_v = {nm: [rates.Cells(rr, 9 + i).Value for i in range(12)] for nm, rr in (("분양", 8), ("분양누계", 9), ("입주", 13), ("미입주", 15))}
        checks = [wb.Worksheets("현장입력").Range(f"I{r}").Value for r in range(14, 23)]
        result = dict(errors=errors, rows=rows, stage=stage, rates=rates_v, checks=[c for c in checks if c], sheet=bname)
        if save_as:
            wb.SaveAs(os.path.abspath(save_as))
        if keep_open:
            return result, xl, wb
        wb.Close(False)
        return result
    finally:
        if not keep_open:
            xl.Quit()


def _addr(r, c):
    s = ""
    while c:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return f"{s}{r}"


def show(res, only_nonzero=True, width=100):
    print("단계:", " | ".join(f"{i + 1}:{s or '-'}" for i, s in enumerate(res["stage"])))
    for k, v in res["rates"].items():
        print(f"  {k:5}", v)
    if res["checks"]:
        print("점검:", res["checks"])
    print("오류 셀:", len(res["errors"]))
    for e in res["errors"][:30]:
        print("   ", e)
    for row in res["rows"]:
        tot = row["total"] or 0
        if only_nonzero and not tot and row["item"] != "합계":
            continue
        ms = " ".join(f"{(m or 0) / 1000:>8,.0f}" if isinstance(m, (int, float)) else f"{str(m):>8}" for m in row["months"])
        g = (row["g"] or "")[:width]
        print(f"r{row['r']:<3} {str(row['item'])[:6]:6} {(tot or 0) / 1000:>12,.1f} | {g}")
        print(f"      {ms}")
