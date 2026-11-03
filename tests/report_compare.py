"""report_xl(xlsx XML 편집) 결과 ↔ report(Excel COM) 결과 비교: 같은 res/data 로 만든 두 파일의 값·수식·병합·행높이·숨김·테두리"""
import sys, os, json, glob, time, re
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # 이 저장소 폴더(외장하드 E: 또는 F:)
sys.path.insert(0, os.path.join(ROOT, "tool"))
sys.stdout.reconfigure(encoding="utf-8")
import warnings
warnings.filterwarnings("ignore")
import openpyxl
import report_xl

jobs = sys.argv[1:] or ["bfa16d5fa55f"]
form = open(glob.glob(os.path.join(ROOT, "회사양식", "*.xlsx"))[0], "rb").read()
for jid in jobs:
    d = os.path.join(ROOT, "app", "_jobs", jid)
    res = json.load(open(os.path.join(d, "result.json"), encoding="utf-8"))
    data = json.load(open(os.path.join(d, "input.json"), encoding="utf-8"))["data"]
    t = time.time()
    out, info = report_xl.build_report(form, res, data)
    dt = time.time() - t
    outp = os.path.join(d, "report_xl.xlsx")
    open(outp, "wb").write(out)
    A = openpyxl.load_workbook(os.path.join(d, "report.xlsx"))
    B = openpyxl.load_workbook(outp)
    print(f"■ {jid}: report_xl {dt:.1f}초, {len(out):,} bytes | 시트 {A.sheetnames == B.sheetnames} 상태 "
          f"{[(ws.title, ws.sheet_state) for ws in A.worksheets if ws.sheet_state != 'visible']} vs "
          f"{[(ws.title, ws.sheet_state) for ws in B.worksheets if ws.sheet_state != 'visible']}")
    for name in A.sheetnames:
        if name == report_xl.DATA_SHEET:
            continue
        wa, wb = A[name], B[name]
        diffs = []
        maxr = max(wa.max_row, wb.max_row)
        maxc = max(wa.max_column, wb.max_column)
        for r in range(1, maxr + 1):
            for c in range(1, maxc + 1):
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
        hd = []
        for r in range(1, maxr + 1):
            ha, hb = wa.row_dimensions[r].height, wb.row_dimensions[r].height
            if (ha or 0) != (hb or 0) and not (ha is None and hb is None):
                hd.append(f"{r}:{ha}/{hb}")
        bd = []
        for r in range(1, min(maxr, 120) + 1):
            for c in range(1, min(maxc, 22) + 1):
                ba, bb = wa.cell(r, c).border, wb.cell(r, c).border
                st = lambda s: getattr(s, "style", None)
                sa = (st(ba.top), st(ba.bottom), st(ba.left), st(ba.right))
                sb = (st(bb.top), st(bb.bottom), st(bb.left), st(bb.right))
                if sa != sb:
                    bd.append(f"{wa.cell(r, c).coordinate}:{sa}/{sb}")
        print(f"   [{name}] 값·수식 차이 {len(diffs)}, 병합 {'같음' if ma == mb else '다름 ' + str(set(ma) ^ set(mb))[:120]}, "
              f"행높이 차이 {len(hd)}, 테두리 차이 {len(bd)}")
        for x in diffs[:6]:
            print("      ", x)
        for x in hd[:4]:
            print("       높이", x)
        for x in bd[:6]:
            print("       테두리", x)
    print("   저장:", outp)
