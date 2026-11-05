# -*- coding: utf-8 -*-
"""연도별 기준안으로 자동작성 양식을 만들어 스크래치패드에 저장하고 반영 결과(확인 필요 항목)를 출력"""
import sys, os, warnings
warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
ROOT = r"E:\1. 우미\1. 본사업무\★ 판관비 예산안\☆ 예산안 자동작성"
BUD = r"E:\1. 우미\1. 본사업무\★ 판관비 예산안"
OUT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "tool"))
import openpyxl
import build_template, kijun_parser

KIJUN = {
    2027: os.path.join(BUD, r"★ 2027 예산안\PM팀 판관비 예산기준(안)_2026.10.15.xlsx"),
    2026: os.path.join(BUD, r"★ 2026 예산안\PM팀 판관비 예산기준(안)_2025.10.27.xlsx"),
    2025: os.path.join(BUD, r"★ 2025 예산안\PM팀 판관비 예산기준(안)_2024.11.18.xlsx"),
    2024: os.path.join(BUD, r"★ 2024 예산안\PM팀 판관비 예산기준(안)_2023.11.07.xlsx"),
    2023: os.path.join(BUD, r"마케팅팀 판관비 예산기준(안)_2022.10.31.xlsx"),
}
years = [int(a) for a in sys.argv[1:]] or sorted(KIJUN)
for y in years:
    out = os.path.join(OUT, f"양식_{y}.xlsx")
    print(f"\n######## {y} ← {os.path.basename(KIJUN[y])}")
    build_template.build(out, KIJUN[y])
    wb = openpyxl.load_workbook(out, data_only=False)
    ws = wb["기준단가"]
    C = kijun_parser.COL
    print("기준단가 시트 머리:", ws.cell(2, 3).value, ws.cell(2, 7).value, "예산연도", ws.cell(2, 10).value)
    rows = []
    for r in range(kijun_parser.FIRST_ROW, ws.max_row + 1):
        key = ws.cell(r, C["name"]).value
        if not key:
            continue
        res = ws.cell(r, C["result"]).value or ""
        rows.append((key, ws.cell(r, C["val"]).value, ws.cell(r, C["unit"]).value, res, (ws.cell(r, C["src_text"]).value or "")[:70]))
    print("단가 수", len(rows))
    for key, val, unit, res, src in rows:
        flag = "!!" if str(res).startswith("확인") else ("**" if str(res).startswith("변경") else "  ")
        print(f" {flag} {key:<18} {str(val):>8} {str(unit or ''):<10} {res[:60]:<62} | {src}")
    ws2 = wb["기준안 반영결과"]
    print("--- 단가로 읽지 않은 기준안 줄 ---")
    started = False
    for r in range(1, ws2.max_row + 1):
        a = ws2.cell(r, 1).value
        if isinstance(a, str) and a.startswith("3."):
            started = True
            continue
        if started and ws2.cell(r, 3).value:
            print("   ", ws2.cell(r, 1).value, "|", ws2.cell(r, 2).value, "|", str(ws2.cell(r, 3).value)[:110])
