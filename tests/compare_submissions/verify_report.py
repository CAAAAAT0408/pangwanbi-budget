# -*- coding: utf-8 -*-
"""엑셀 보고서를 Excel COM 으로 열어 재계산 → 오류 셀 검사 → 요약 시트 PDF → PNG 렌더링(육안 확인용)"""
import sys, os, time
sys.stdout.reconfigure(encoding="utf-8")
import site
site.addsitedir(site.getusersitepackages())
import pythoncom, win32com.client
import fitz

src = os.path.abspath(sys.argv[1])
outdir = os.path.dirname(src)
pythoncom.CoInitialize()
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(src)
    xl.CalculateFull()
    errs = []
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        vals = ur.Value
        if vals is None:
            continue
        if not isinstance(vals, tuple):
            vals = ((vals,),)
        for i, row in enumerate(vals):
            for j, v in enumerate(row):
                if isinstance(v, str) and v.startswith("#") and v.endswith(("!", "?", "A")):
                    errs.append(f"{ws.Name}!R{ur.Row+i}C{ur.Column+j} {v}")
                if isinstance(v, int) and v in (-2146826281, -2146826246, -2146826259, -2146826288, -2146826252, -2146826265, -2146826273):
                    errs.append(f"{ws.Name}!R{ur.Row+i}C{ur.Column+j} err{v}")
        print(f"[{ws.Name}] 사용범위 {ur.Address}")
    print("오류 셀:", len(errs), errs[:10])
    s = wb.Worksheets("요약")
    for a in ("D21", "E21", "F21", "G21", "H21"):
        pass
    # 요약 합계 행 값 출력
    for r in range(1, s.UsedRange.Rows.Count + 1):
        b = s.Cells(r, 2).Value
        if isinstance(b, str) and (b.startswith("합계") or b.startswith("일치·근접")):
            print("  ", b, "|", s.Cells(r, 4).Value, s.Cells(r, 5).Value, s.Cells(r, 6).Value, s.Cells(r, 7).Value, s.Cells(r, 8).Value)
    pdf = os.path.join(outdir, "요약.pdf")
    s.ExportAsFixedFormat(0, pdf)
    wb.Close(False)
finally:
    xl.Quit()
doc = fitz.open(pdf)
print("요약 PDF 쪽수:", len(doc))
for i, page in enumerate(doc):
    png = os.path.join(outdir, f"요약_{i+1}.png")
    page.get_pixmap(dpi=110).save(png)
    print("렌더:", png)
