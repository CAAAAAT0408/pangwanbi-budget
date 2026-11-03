"""두 xlsx(회사 양식 보고서)의 값·수식·병합·숨김 비교"""
import sys, re
sys.stdout.reconfigure(encoding="utf-8")
import warnings
warnings.filterwarnings("ignore")
import openpyxl
A = openpyxl.load_workbook(sys.argv[1])
B = openpyxl.load_workbook(sys.argv[2])
print("시트 같음:", A.sheetnames == B.sheetnames, "| 숨김:", [(w.title, w.sheet_state) for w in B.worksheets if w.sheet_state != "visible"])
tot = 0
for name in A.sheetnames:
    if name.startswith("_"):
        continue
    wa, wb = A[name], B[name]
    diffs = []
    for r in range(1, max(wa.max_row, wb.max_row) + 1):
        for c in range(1, max(wa.max_column, wb.max_column) + 1):
            va, vb = wa.cell(r, c).value, wb.cell(r, c).value
            if isinstance(va, str) and va.startswith("="):
                va = va.replace("=+", "=")
            if isinstance(vb, str) and vb.startswith("="):
                vb = vb.replace("=+", "=")
            if va in (None, "") and vb in (None, ""):
                continue
            if isinstance(va, (int, float)) and isinstance(vb, (int, float)) and abs(va - vb) < 1e-6:
                continue
            if va != vb:
                diffs.append(f"{wa.cell(r, c).coordinate}: {va!r} vs {vb!r}")
    ma = sorted(str(m) for m in wa.merged_cells.ranges)
    mb = sorted(str(m) for m in wb.merged_cells.ranges)
    tot += len(diffs)
    print(f"  [{name}] 값·수식 차이 {len(diffs)} 병합 {'같음' if ma == mb else '다름'}", diffs[:4])
print("총 차이", tot)
