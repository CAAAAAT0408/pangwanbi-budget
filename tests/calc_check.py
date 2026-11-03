"""xlcalc 검증: Excel 이 계산해 둔 작업 파일(work.xlsx)의 수식을 다시 계산해 캐시값과 전부 비교"""
import sys, os, time, datetime, glob
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # 이 저장소 폴더(외장하드 E: 또는 F:)
sys.path.insert(0, os.path.join(ROOT, "tool"))
sys.stdout.reconfigure(encoding="utf-8")
import warnings
warnings.filterwarnings("ignore")
import openpyxl
import xlcalc

paths = sys.argv[1:] or sorted(glob.glob(os.path.join(ROOT, "app", "_jobs", "*", "work.xlsx")), key=os.path.getmtime)[-3:]
for path in paths:
    t = time.time()
    m = xlcalc.Model(path)
    t_load = time.time() - t
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    t = time.time()
    n = bad = err = 0
    shown = 0
    by_sheet = {}
    for ws in wb.worksheets:
        for ri, row in enumerate(ws.iter_rows(), 1):
            for ci, cell in enumerate(row, 1):
                raw = m.cells.get(ws.title, {}).get((ri, ci))
                if not (isinstance(raw, tuple) and raw and raw[0] == "F"):
                    continue
                n += 1
                coord = f"{xlcalc.num2col(ci)}{ri}"
                cell = type("C", (), {"coordinate": coord, "row": ri, "column": ci, "value": cell.value})()
                exp = cell.value
                if isinstance(exp, (datetime.datetime, datetime.date)):
                    exp = xlcalc.to_serial(exp)
                got = m.cell(ws.title, cell.row, cell.column)
                ok = False
                if isinstance(got, xlcalc.XLError):
                    err += 1
                    ok = isinstance(exp, str) and exp.startswith("#") and exp == got.code
                elif isinstance(got, bool) or isinstance(exp, bool):
                    ok = (got == exp) or (exp is None and got is False)
                elif isinstance(got, (int, float)) and isinstance(exp, (int, float)):
                    ok = abs(float(got) - float(exp)) <= 1e-6 * max(1.0, abs(float(exp)))
                elif isinstance(got, str) or isinstance(exp, str):
                    ok = (got or "") == (exp or "") or (exp is None and got == "")
                else:
                    ok = got == exp or (got in (None, 0, 0.0) and exp in (None, 0))
                if not ok:
                    bad += 1
                    by_sheet[ws.title] = by_sheet.get(ws.title, 0) + 1
                    if shown < 25:
                        shown += 1
                        print(f"   ✗ {ws.title}!{cell.coordinate}: 기대 {exp!r} / 계산 {got!r}   ← {raw[1][:110]}")
    t_calc = time.time() - t
    print(f"■ {os.path.relpath(path, ROOT)}: 수식 {n}개, 불일치 {bad}개 {by_sheet}, 오류값 {err}개, 읽기 {t_load:.1f}초 계산 {t_calc:.1f}초")
    wb.close()
