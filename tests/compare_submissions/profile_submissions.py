# -*- coding: utf-8 -*-
"""과거 제출본(회사 양식) 엑셀에서 현장 조건·실제 예산(항목·세목·월별)을 뽑아 JSON 으로 저장 + 요약 출력"""
import sys, os, re, json, glob, warnings
warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl

BUD = r"E:\1. 우미\1. 본사업무\★ 판관비 예산안"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "profiles")
os.makedirs(OUT, exist_ok=True)

FILES = {
    2027: [r"★ 2027 예산안\기공급\평택고덕(A66BL)\(평택고덕(A66BL)) 2027년 사업비 예산편성 26.11.02.xlsx"],
    2026: [r"★ 2026 예산안\신규분양\평택고덕 A36BL\(평택고덕 A36BL) 2026년 사업비 예산편성 25.11.08.xlsx",
           r"★ 2026 예산안\신규분양\평택고덕 A33BL\(평택고덕 A33BL) 2026년 사업비 예산편성 25.11.08.xlsx",
           r"★ 2026 예산안\기공급\김포북변\(김포북변) 2026년 사업비 예산편성 25.11.08.xlsx",
           r"★ 2026 예산안\기공급\오산 서2\(오산서2구역) 2026년 사업비 예산편성_25.11.21.xlsx",
           r"★ 2026 예산안\기공급\오산 서2\(오산서2구역) 2026년 사업비 예산편성_26.02.10_재작성.xlsx",
           r"★ 2026 예산안\미결사항\화성태안\(화성태안2차) 2026년 사업비 예산편성 25.11.08.xlsx",
           r"★ 2026 예산안\미결사항\화성태안\(화성태안2차) 2026년 사업비 예산편성 260210.xlsx",
           r"★ 2026 예산안\미결사항\과천지식정보\(과천지식정보) 2026년 사업비 예산편성 25.11.08.xlsx",
           r"★ 2026 예산안\미결사항\과천지식정보\(과천지식정보) 2026년 사업비 예산편성 260210.xlsx",
           r"★ 2026 예산안\미결사항\화성동탄2차\(화성동탄2차) 2026년 사업비 예산편성 25.11.08.xlsx",
           r"★ 2026 예산안\미결사항\화성비봉\(화성비봉) 2026년 사업비 예산편성 25.11.08_수정.xlsx",
           r"★ 2026 예산안\미결사항\화성비봉\(화성비봉) 2026년 사업비 예산편성 260210.xlsx"],
    2025: [r"★ 2025 예산안\신규분양\(오산서2구역) 2025년 사업비 예산편성_24.11.25.xlsx",
           r"★ 2025 예산안\기분양\과천지식정보\(과천지식정보) 2025년 사업비 예산편성 24.12.03.xlsx",
           r"★ 2025 예산안\기분양\화성동탄2차\(화성동탄2차) 2025년 사업비 추가 예산편성_24.11.26.xlsx",
           r"★ 2025 예산안\기분양\화성비봉\(화성비봉) 2025년 사업비 예산편성 24.11.26.xlsx",
           r"★ 2025 예산안\기분양\화성태안\(화성태안2차) 2025년 사업비 예산편성_24.11.26.xlsx",
           r"★ 2025 예산안\기분양\오산세교1차\(오산세교1차) 2025년 사업비 예산편성_24.11.26.xlsx",
           r"★ 2025 예산안\기분양\김포북변\(김포북변) 2025년 사업비 예산편성 24.11.26.xlsx"],
    2024: [r"★ 2024 예산안\기분양\과천지식정보\(과천지식정보) 2024년 사업비 예산편성 23.12.13.xlsx",
           r"★ 2024 예산안\기분양\이천중리\(이천중리B1BL) 2024년 사업비 예산편성_23.12.13.xlsx",
           r"★ 2024 예산안\기분양\이천중리\(이천중리B2BL) 2024년 사업비 예산편성_23.12.13.xlsx",
           r"★ 2024 예산안\기분양\화성동탄2차\(화성동탄2차) 2024년 사업비 추가 예산편성_23.11.17.xlsx",
           r"★ 2024 예산안\기분양\화성비봉\(화성비봉) 2024년 사업비 예산편성 23.12.13.xlsx",
           r"★ 2024 예산안\기분양\화성태안\(화성태안2차) 2024년 사업비 예산편성_23.12.14.xlsx",
           r"★ 2024 예산안\신규분양\김포북변\(김포북변) 2024년 사업비 예산편성 23.12.13.xlsx",
           r"★ 2024 예산안\신규분양\냠양뉴타운\(남양뉴타운2차) 2024년 사업비 예산편성_23.12.14.xlsx",
           r"★ 2024 예산안\신규분양\냠양뉴타운\보류\(남양뉴타운1차) 2024년 사업비 예산편성_23.11.07.xlsx",
           r"★ 2024 예산안\신규분양\오산세교2차\(오산세교2차) 2024년 사업비 예산편성_23.12.14.xlsx"],
    2023: [r"★ 2023 예산안\기분양\과천지식정보\(과천지식정보 MH) 2023년 사업비 예산편성_23.01.03.xlsx",
           r"★ 2023 예산안\기분양\과천지식정보\(과천지식정보) 2023년 사업비 예산편성_23.01.03.xlsx",
           r"★ 2023 예산안\기분양\김포한강\(김포 한강 AC-02BL) 2023년 사업비 추가 예산편성_23.01.03.xlsx",
           r"★ 2023 예산안\기분양\오산세교2차\(오산세교2차 MH) 2023년 사업비 예산편성_23.01.03.xlsx",
           r"★ 2023 예산안\기분양\오산세교2차\(오산세교2차) 2023년 사업비 예산편성_23.01.03.xlsx",
           r"★ 2023 예산안\기분양\위례1차\(위례신도시1차) 2023년 사업비 추가 예산편성_23.01.03.xlsx",
           r"★ 2023 예산안\기분양\위례2차\(위례신도시2차 MH) 2023년 사업비 예산편성_23.01.03.xlsx",
           r"★ 2023 예산안\기분양\위례2차\(위례신도시2차) 2023년 사업비 예산편성_23.01.03.xlsx",
           r"★ 2023 예산안\기분양\평택소사벌2차\(평택소사벌2차) 2023년 사업비 추가 예산편성_23.04.26.xlsx",
           r"★ 2023 예산안\기분양\화성동탄1차\(화성동탄1차) 2023년 사업비 추가 예산편성_23.01.19.xlsx",
           r"★ 2023 예산안\기분양\화성동탄2차(신규작성)\(화성동탄2차) 2023년 사업비 추가 예산편성_23.01.19_수정완료.xlsx",
           r"★ 2023 예산안\기분양\화성비봉\(화성비봉 MH) 2023년 사업비 예산편성_23.01.04.xlsx",
           r"★ 2023 예산안\기분양\화성비봉\(화성비봉) 2023년 사업비 예산편성_23.01.04.xlsx",
           r"★ 2023 예산안\기분양\화성태안\(화성태안1차) 2023년 사업비 예산편성_23.02.13(최종).xlsx",
           r"★ 2023 예산안\기분양\화성태안\(화성태안2차 MH) 2023년 사업비 예산편성_23.01.04.xlsx",
           r"★ 2023 예산안\기분양\화성태안\(화성태안2차) 2023년 사업비 예산편성_23.02.13(최종) (1).xlsx",
           r"★ 2023 예산안\수정안\이천중리(완)\(이천중리) 2023년 사업비 예산편성_23.03.29_수정안.xlsx",
           r"★ 2023 예산안\신규분양\이천중리1차\(이천중리B2BL) 2023년 사업비 예산편성_23.04.20.xlsx",
           r"★ 2023 예산안\신규분양\이천중리2차(완)\(이천중리B1BL) 2023년 사업비 예산편성_23.01.05.xlsx"],
}


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


def profile(path, year):
    wbv = openpyxl.load_workbook(path, data_only=True)
    wbf = openpyxl.load_workbook(path, data_only=False)
    out = dict(path=path, year=year, file=os.path.basename(path), sheets=wbv.sheetnames)
    # 기본정보
    if "기본정보" in wbv.sheetnames:
        ws = wbv["기본정보"]
        types, name = [], None
        for r in range(1, ws.max_row + 1):
            c, d, f = ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 6).value
            if c == "현장명":
                name = f
            if d not in (None, "", "계") and c in (None, "") and ws.cell(r, 2).value in (None, "세대수") and isinstance(f, (int, float)):
                types.append([str(d), f])
            if c == "담당소장":
                break
        out["name"] = name
        out["types"] = types
        out["total_units"] = sum(t[1] for t in types)
    # 예산 시트
    bname = next((n for n in wbv.sheetnames if re.fullmatch(r"\d{2}년 예산", n)), None)
    out["budget_sheet"] = bname
    if bname:
        ws, wf = wbv[bname], wbf[bname]
        # 머리글 행 찾기
        hdr = None
        for r in range(1, 12):
            if str(ws.cell(r, 2).value or "").replace(" ", "") == "항목":
                hdr = r
                break
        hdr = hdr or 6
        out["stage"] = [ws.cell(hdr - 1, c).value for c in range(10, 22)]
        rows, cur = [], None
        for r in range(hdr + 1, ws.max_row + 1):
            a, b, d, f, g, h = (ws.cell(r, k).value for k in (1, 2, 4, 6, 7, 8))
            months = [num(ws.cell(r, c).value) for c in range(10, 22)]
            if isinstance(a, str) and a.replace(" ", "").startswith("합계"):
                out["total"] = num(f)
                out["total_row"] = r
                break
            if b not in (None, ""):
                cur = str(b).strip()
            sub = str(d).strip() if d not in (None, "") else ""
            if sub.replace(" ", "") == "소계":
                rows.append(dict(r=r, item=cur, sub="소계", total=num(f), months=months))
                continue
            if f in (None, "") and not any(months) and not g:
                continue
            fx = [wf.cell(r, c).value for c in range(10, 22)]
            rows.append(dict(r=r, item=cur, sub=sub, total=num(f), g=(g or "").strip(), note=h,
                             months=months, formulas=[x if isinstance(x, str) and x.startswith("=") else None for x in fx]))
        out["rows"] = rows
    # 분양율,입주율
    if "분양율,입주율" in wbv.sheetnames:
        ws = wbv["분양율,입주율"]
        rs = {}
        for r in range(1, ws.max_row + 1):
            c = str(ws.cell(r, 3).value or "").replace(" ", "")
            if c in ("분양세대", "입주세대", "분양누계", "입주누계"):
                rs[c] = dict(prev=[num(ws.cell(r, k).value) for k in range(5, 9)],
                             months=[num(ws.cell(r, k).value) for k in range(9, 21)])
        out["rates"] = rs
    # 분양대행 시트 (조직 블록)
    aname = next((n for n in wbv.sheetnames if n.startswith("분양대행")), None)
    if aname:
        ws = wbv[aname]
        blocks = []
        for r in range(1, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if isinstance(a, str) and ("분양대행" in a or "촉진" in a):
                # 단가 D열, 세대수 월별 J..U
                blocks.append(dict(r=r, label=a.replace("\n", " "), price=ws.cell(r, 4).value,
                                   months_units=[num(ws.cell(r, k).value) for k in range(10, 22)]))
        out["agency_blocks"] = blocks
        # 총계 행
        for r in range(1, ws.max_row + 1):
            if ws.cell(r, 1).value == "총계":
                out["agency_total_months"] = [num(ws.cell(r, k).value) for k in range(10, 22)]
    return out


def compact(months):
    return " ".join(f"{int(round(m / 1000)):,}" if m else "·" for m in months)


years = [int(a) for a in sys.argv[1:]] or sorted(FILES)
for y in years:
    for rel in FILES[y]:
        p = os.path.join(BUD, rel)
        if not os.path.exists(p):
            print("!! 없음:", rel)
            continue
        try:
            pr = profile(p, y)
        except Exception as e:
            print("!! 오류:", rel, repr(e))
            continue
        key = f"{y}_{re.sub(r'[^0-9A-Za-z가-힣.]+', '_', pr['file'][:-5])}"
        pr["key"] = key
        with open(os.path.join(OUT, key + ".json"), "w", encoding="utf-8") as f:
            json.dump(pr, f, ensure_ascii=False, indent=1)
        print(f"\n===== [{y}] {pr['file']}")
        print(f"  현장명 {pr.get('name')} | 세대 {pr.get('total_units')} {pr.get('types')}")
        print(f"  단계: {pr.get('stage')}")
        print(f"  합계 {pr.get('total', 0) / 1000:,.0f}천원  (시트 {pr.get('budget_sheet')})")
        rs = pr.get("rates") or {}
        for k in ("분양세대", "입주세대"):
            if k in rs:
                print(f"  {k}: 전년 {rs[k]['prev']} | {rs[k]['months']}")
        for b in pr.get("agency_blocks") or []:
            if any(b["months_units"]):
                print(f"  대행블록 r{b['r']} {b['label'][:14]} 단가 {b['price']} 세대 {b['months_units']}")
        for row in pr.get("rows") or []:
            if row["sub"] == "소계":
                print(f"   ── {row['item']} 소계 {row['total'] / 1000:>14,.0f}")
                continue
            if not row["total"] and not any(row["months"]):
                continue
            print(f"   {str(row['item'])[:6]:<6} {row['sub'][:7]:<7} {row['total'] / 1000:>12,.0f} | {row['g'][:48]:<48} | {compact(row['months'])}")
