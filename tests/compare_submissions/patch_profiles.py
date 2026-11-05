# -*- coding: utf-8 -*-
"""profiles/*.json 보강: 분양대행 시트 '(조직)' 블록의 타입별 세대수 합(월별) + 공가세대관리비 시트의 타입별 공급평(㎡ 환산)"""
import sys, os, re, json, glob, warnings
warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl
SP = os.path.dirname(os.path.abspath(__file__))


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


for p in sorted(glob.glob(os.path.join(SP, "profiles", "*.json"))):
    pr = json.load(open(p, encoding="utf-8"))
    wb = openpyxl.load_workbook(pr["path"], data_only=True)
    # 조직 블록 세대수 합
    aname = next((n for n in wb.sheetnames if n.startswith("분양대행")), None)
    org = [0] * 12
    init = [0] * 12
    if aname:
        ws = wb[aname]
        r = 1
        while r <= ws.max_row:
            a = ws.cell(r, 1).value
            if isinstance(a, str) and "분양대행" in a and ws.cell(r, 2).value == "세대수":
                is_org = "조직" in a
                rr = r
                while rr <= ws.max_row and ws.cell(rr, 2).value != "금 액":
                    c = ws.cell(rr, 3).value
                    if c not in (None, "", 0) and ws.cell(rr, 4).value not in (None, ""):
                        for i in range(12):
                            v = num(ws.cell(rr, 10 + i).value)
                            if is_org:
                                org[i] += v
                            else:
                                init[i] += v
                    rr += 1
                r = rr
            r += 1
    pr["agency_org_units"] = [round(x, 1) for x in org]
    pr["agency_init_units"] = [round(x, 1) for x in init]
    # 공급평: 공가 시트 '금 액' 블록 부근 V~X 열의 숫자(㎡ 60~250 또는 평 18~75)를 타입 순서대로
    supply = {}
    if "공가세대관리비" in wb.sheetnames:
        ws = wb["공가세대관리비"]
        types_in_sheet = []
        amt_row = None
        for r in range(1, ws.max_row + 1):
            if ws.cell(r, 2).value == "금 액":
                amt_row = r
                for k in range(0, 25):
                    c = ws.cell(r + k, 3).value
                    if c in (None, "", 0):
                        break
                    types_in_sheet.append(str(c))
                break
        if amt_row and types_in_sheet:
            m2, py = [], []
            for r in range(amt_row, amt_row + len(types_in_sheet) + 8):
                for col in (22, 23, 24):
                    v = ws.cell(r, col).value
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        if 60 <= v <= 250:
                            m2.append((r, col, float(v)))
                        elif 18 <= v < 60:
                            py.append((r, col, float(v)))
            vals = None
            if len(m2) >= len(types_in_sheet):
                vals = [v for _, _, v in m2[:len(types_in_sheet)]]
            elif len(py) >= len(types_in_sheet):
                vals = [round(v / 0.3025, 2) for _, _, v in py[:len(types_in_sheet)]]
            if vals:
                supply = dict(zip(types_in_sheet, vals))
    pr["supply"] = supply
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(pr, fh, ensure_ascii=False, indent=1)
    print(pr["key"][:45], "조직세대", sum(org), "초기", sum(init), "공급㎡", supply)
