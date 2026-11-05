# -*- coding: utf-8 -*-
"""비교 결과(compare_results_basic/adjusted.json) + 사이트 동작 테스트 결과 → 엑셀 보고서(요약 1장 + 상세 시트)"""
import sys, os, json, re, glob, datetime, warnings
warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SP = os.path.dirname(os.path.abspath(__file__))
ROOT = r"E:\1. 우미\1. 본사업무\★ 판관비 예산안\☆ 예산안 자동작성"
sys.path.insert(0, os.path.join(ROOT, "tool"))
import xlcalc, site_io

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(SP, "현장별 비교·동작 테스트 결과.xlsx")
basic = {r["key"]: r for r in json.load(open(os.path.join(SP, "compare_results_basic.json"), encoding="utf-8"))}
adj = json.load(open(os.path.join(SP, "compare_results_adjusted.json"), encoding="utf-8"))

F = lambda **k: Font(name="맑은 고딕", size=k.pop("size", 9), **k)
FILL_H = PatternFill("solid", fgColor="D9E1F2")
FILL_T = PatternFill("solid", fgColor="1F4E78")
FILL_OK = PatternFill("solid", fgColor="E2EFDA")
FILL_NEAR = PatternFill("solid", fgColor="FFF2CC")
FILL_BAD = PatternFill("solid", fgColor="FCE4D6")
FILL_GRAY = PatternFill("solid", fgColor="EDEDED")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
C = Alignment(horizontal="center", vertical="center", wrap_text=True)
L = Alignment(horizontal="left", vertical="center", wrap_text=True)
R = Alignment(horizontal="right", vertical="center")
NF = '#,##0;[Red]-#,##0;"-"'


def hdr(ws, row, labels, widths=None, fill=FILL_H):
    for i, h in enumerate(labels, start=1):
        c = ws.cell(row, i, h)
        c.font = F(bold=True)
        c.fill = fill
        c.alignment = C
        c.border = BOX
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def put(ws, r, c, v, nf=None, al=None, bold=False, fill=None, color=None):
    cell = ws.cell(r, c, v)
    cell.font = F(bold=bold, color=color) if color else F(bold=bold)
    cell.border = BOX
    if nf:
        cell.number_format = nf
    cell.alignment = al or (R if isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("=")) else L)
    if fill:
        cell.fill = fill
    return cell


def k(v):
    return round((v or 0) / 1000)


# ─────────────────────────────── 현장 유형·판정·원인
KIND_KO = {"new": "신규분양", "remain": "잔여분양(기공급)", "movein": "입주", "office": "분양사무실"}

# 줄 키 → 차이 분류
CAT = {
    "수량·시점 선택(세부 조정으로 맞출 수 있음)": ["대형마트", "갑티슈", "물티슈", "종이컵", "마스크", "크린홈", "오픈 선물", "청약 선물", "당첨자 선물", "계약 선물",
                                   "조직분양 현수막", "보험료", "운반비", "PC", "복합기", "숙소비", "숙소비(입주)", "숙소관리비", "청소", "아르바이트", "식대", "여비",
                                   "소모품", "가구", "분양촉진(MGM)", "통신요금", "상하수도", "전기료", "난방비·전기미인입", "피복비", "주차장", "전기·소방·오수 안전관리",
                                   "무인경비·보안", "발급수수료·확정일자", "송금수수료", "프린터·파쇄기", "M/H 중개수수료", "카페테리아", "무순위 전산추첨", "계약취소 전산추첨",
                                   "특별공급·청약대행", "문고", "입주지원센터", "입주촉진책", "입주상담사", "사전점검행사", "건축물관리대장", "이전등기", "어린이집"],
    "규칙 차이(검토 필요)": ["인지세", "분양대행·조직 수수료", "우편·등기·내용증명", "입주매니저", "공가세대관리비", "대납이자"],
    "옛 기준안 항목(2027 기준안에 없음)": ["컨테이너·텐트·난방기·에어렉스", "실내화", "네트워크 공사", "이벤트(네일아트 등)", "방역", "대봉투(서식대)", "보존등기", "임차료(기타)"],
    "현장 특수 항목(도구 범위 밖)": ["법무사", "상담사", "분양사무실 관리비", "(기타)"],
}
CAT_OF = {key: cat for cat, keys in CAT.items() for key in keys}


def category(line):
    key = line["key"]
    if line["class"] == "일치":
        return "일치"
    if key.startswith("(기타"):
        return "현장 특수 항목(도구 범위 밖)"
    if line["tool_g"].startswith("[분양사무실]") and not line["actual"]:
        return "분양사무실 자동 항목(실제 미편성)"
    return CAT_OF.get(key, "기타")


MANUAL = [
    ("A33BL", "제출본 '조직분양 현수막' 산식 오타(=40000*12000*4 → 12,000장/주)로 3,456,000천원 과다. 오타 보정 시 실제 4,725,720 ↔ 도구 4,708,004(99.6%)."),
    ("오산서2구역) 2026", "조직 단가가 1~7월 8,000(B급)·8~12월 10,000(C급)으로 연중 변경됨(도구는 한 단가, -432,000). 아파트 인지세는 도구만 반영(+62,175). 중개수수료 30,000·무인경비 공사비 2,000·운반비를 실제는 재반영."),
    ("오산서2구역) 2025", "실제는 조직분양 세대에 분양대행 2,000 + 조직 10,000 을 중복 적용(도구는 조직만). 현수막 30천원×1,500장×1주/월, MGM 2,000천원/세대, 피복비 30,000(2025 기준) 등 수량·단가 선택 차이. 인지세 실제 7.5천원/세대(산식 7500) ↔ 기준 75천원."),
    ("오산세교1차", "잔여분양 대행 단가 특례(초기 900·조직 8,000 소수 세대)와 1월 누계 장부 입력 — 도구 규칙 밖. 입주 11월 공가관리비는 유사."),
    ("오산세교2차) 2024", "총액 99% 일치. 현수막(45,000/월=1주), MGM 2,000천원/세대(실제) ↔ 1,000(기준), 인지세 미편성(실제), 보험료 사전영업 전월(도구 0)."),
    ("남양뉴타운1차", "보류 폴더의 초안 — F열 합계(3,827,739)와 월별 합(5,896,125)이 다른 줄 6개. 월별 합 기준 비교. 인지세·현수막 미편성(실제), 주차장 2개월 60,000(실제)."),
    ("남양뉴타운2차", "대납이자 500,527(대납이자 시트, 도구는 무이자 Y·대출잔액 입력 필요) 과 현수막 1주/월(실제) ↔ 4주(도구)가 차이의 대부분."),
    ("이천중리B1BL) 2024", "1월 오픈(23.12 사전영업)·3~6월 조직·7월부터 분양사무실. 전기료 4~11월 800천원(실제 축소) ↔ 기준 8,000, 현수막 1주/월, 인지세 미편성이 차이."),
    ("김포북변) 2024", "M/H 1~3월 철거 전: 실제는 전기료·상하수도·무인경비·전기소방만 편성. 도구는 M/H 운영 항목 전체(통신·복합기·PC·프린터·식대·여비·청소·인지세) 자동 반영."),
    ("화성비봉) 2024", "12월 입주개시 → 도구는 입주매니저·건축물대장·이전등기를 다음 달(2025.1)로 넘겨 0, 실제는 12월에 편성(-85,643). 입주상담사 2명×12,000(실제) ↔ 10,000(도구)."),
    ("화성비봉) 2025", "입주 1월(개시 2024.12). 입주상담사 2명×2,000×2개월(실제) ↔ 200×25일×2명×3개월(도구), 숙소비 미편성(실제)."),
    ("과천지식정보) 2024", "총액 95%. 등기 건수(실제 5회분) ↔ 도구 2회, 운반 2회, 입주상담사 월 금액 차이."),
    ("과천지식정보) 2025", "분양대행 30,000천원×3세대(잔여 재분양 특례)·상가 상담사(250×14일×4인)·무순위 8,800 — 도구 범위 밖. 상가 MGM 은 조정 모드에서 반영."),
    ("과천지식정보) 2026", "분양대행 30,000천원×3세대·상가 상담사 14,000·무순위 10,000 — 특수 항목. 상가 MGM 실제 10,000/호실 ↔ 2026 기준 5,000."),
    ("김포북변) 2026", "분양사무실: 실제는 통신 150·우편 50(+등기 2회)·소모품 300 만 편성, 도구는 복합기·관리비·보안·발급까지 자동 반영."),
    ("김포북변) 2025", "분양사무실: 실제는 통신·우편(+등기 1회)만, 도구는 7개 항목 자동 반영."),
    ("이천중리B2BL) 2024", "분양사무실: 실제는 통신·우편·발급·송금만, 도구는 7개 항목 자동 반영."),
    ("화성태안2차) 2026", "기입주 관리현장: 상가 공가관리비 5만원/실(실제) ↔ 아파트 평당 단가(도구), 상가 MGM 10,000(실제) ↔ 5,000(2026 기준). 분양사무실 자동 항목."),
    ("화성태안2차) 2025", "상가 MGM 4건·상가 상담사 14,000 — 조정 모드에서 MGM 반영 시 104%."),
    ("화성태안2차) 2024", "상가 MGM·청약사이트·전산추첨 등 소액 특수 항목."),
    ("화성동탄2차) 2026", "보유 오피스텔 관리비 300×2실 — 도구 범위 밖."),
    ("추가 예산편성", "예산 0 (대납이자 추가 편성용 파일) — 비교 제외."),
    ("화성비봉) 2026", "공가 1세대만 편성. 도구는 미입주 1세대 기준 연간(2,520) ↔ 실제 1월만(203)."),
    (" MH) 2023", "분양사무실(M/H) 운영비 파일: 실제는 통신·우편(·복합기·관리비·소모품·발급·보안 일부)만 편성 — 도구는 7개 항목 자동 반영(과천 MH 는 18,360 으로 일치)."),
    ("이천중리", "2023 기준안은 분양대행이 급지별 정액(500세대 미만 8억 등)+세대당 1,500~1,800 구조라 도구(2,000/세대)와 다름 — 2023 비교는 참고용."),
    ("위례신도시2차) 2023", "실제 공가관리비 214,542 에 2022.12 분(103,596) 포함. 2023 분만 보면 110,946 ↔ 도구 108,738. 잔여세대 전산추첨 20,000·보존등기 9,764(옛 항목)."),
    ("화성태안2차) 2023년 사업비 예산편성_23.02.13", "사전점검 2월·입주 4월. 공가관리비는 유사(313,404 ↔ 317,184). 통신 1,000/월·복합기+PC 등 실제 수량과 보존등기·방역(옛 항목) 차이. M/H 운영은 없는데 통신 1,000 으로 M/H 로 추정된 영향 없음."),
    ("화성태안1차) 2023", "실제는 사전점검행사 58,500·우편만 편성(입주 항목 대부분 미편성) — 도구는 입주 패키지 전체 반영."),
    ("평택소사벌2차", "상가 MGM·공가관리비·청약사이트 등 2023 추가 편성 — 도구 범위 밖."),
    ("2023년 사업비", "기입주·관리현장 소액(우편 4,000/건·법무사 등) — 도구 범위 밖."),
]


def manual_note(file):
    for key, note in MANUAL:
        if key in file:
            return note
    return ""


def judge(r):
    a, t = r["actual_total"], r["tool_total"]
    if a == 0 and t == 0:
        return "비교 제외", FILL_GRAY
    if a == 0:
        return "비교 제외", FILL_GRAY
    ratio = t / a
    special = not any(x in r["kinds"] for x in ("new", "remain")) and a < 50_000_000
    if 0.95 <= ratio <= 1.05:
        return "일치(±5%)", FILL_OK
    if 0.85 <= ratio <= 1.15:
        return "근접(±15%)", FILL_NEAR
    if special:
        return "소액·특수(참고)", FILL_GRAY
    return "차이 큼", FILL_BAD


def top_lines(r, n=3):
    ls = [x for x in r["lines"] if x["class"] != "일치"]
    ls.sort(key=lambda x: -abs(x["diff"]))
    out = []
    for x in ls[:n]:
        out.append(f"{x['key']}: 실제 {k(x['actual']):,} ↔ 도구 {k(x['tool']):,}")
    return " / ".join(out)


# ─────────────────────────────── 요약 시트
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "요약"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 6
ws.column_dimensions["B"].width = 44
ws.column_dimensions["C"].width = 15
ws.column_dimensions["D"].width = 13
ws.column_dimensions["E"].width = 13
ws.column_dimensions["F"].width = 13
ws.column_dimensions["G"].width = 12
ws.column_dimensions["H"].width = 8
ws.column_dimensions["I"].width = 12
ws.column_dimensions["J"].width = 95

ws["A1"] = "판관비 예산안 자동작성 — 사이트 동작 테스트 · 과거 제출본(2023~2027) 비교 결과"
ws["A1"].font = Font(name="맑은 고딕", size=14, bold=True)
ws["A2"] = (f"작성 {datetime.date.today():%Y.%m.%d} · 비교 방법: 제출본의 현장 조건(세대·오픈/입주/사무실 일정·조직분양·월별 계약세대)을 읽어 그 연도 기준안으로 만든 자동작성 양식에 넣고 수식 계산기(xlcalc, 웹판과 동일 엔진)로 계산 → "
            "제출본 'NN년 예산' 시트와 항목·줄 단위 대조. 금액 단위 천원. '도구 기본값' = 세부 조정은 기준안 기본값, '도구 조정' = 제출본에 적힌 수량·시점(대형마트 주수·선물 인원·보험월·등기월·상가 MGM 등)도 입력.")
ws["A2"].font = F()
ws["A2"].alignment = L
ws.merge_cells("A2:J2")
ws.row_dimensions[2].height = 42

row = 4
ws.cell(row, 1, "1. 사이트 동작 테스트").font = F(size=11, bold=True)
row += 1
hdr(ws, row, ["No", "확인 항목", "결과", "", "", "", "", "", "", "비고"])
ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=9)
tests = [
    ("PC 판 서버(페이지 열기.bat 서버) 켜기 → /api/ping · /api/info", "정상", "2027 양식, 기준안 2026.10.15, 단가 127개, 회사양식 2026 — 포트 8792"),
    ("설계개요 읽기 — 엑셀(A-66BL)", "정상", "타입 8개·1,118세대, 세대수·공급면적 합계 일치(0.7초)"),
    ("설계개요 읽기 — 그림 PDF(A33BL 변경후, Windows OCR)", "정상", "타입 10개·623세대 일치(4.7초)"),
    ("예시 3개 계산(PC 판 API) = 기대값", "정상", "6,106,656 / 1,569,221 / 3,381,545천원, 오류 셀 0"),
    ("타입 11개 현장 계산 + 회사 양식 보고서(Excel)", "정상", "4,754,294천원, 보고서 PNG 3장 45초"),
    ("수식 계산기(xlcalc) ↔ Excel 캐시값 (tests/calc_check)", "정상", "작업 파일 3개 × 수식 5,669개, 불일치 0"),
    ("xlcalc ↔ Excel COM 계산 결과 (tests/model_vs_com)", "정상", "저장된 작업 15개, 차이 0"),
    ("월별 산식(계산 수식) 검사 (tests/formula_check)", "정상", "1,299칸 100% 산식, 재계산 불일치 0"),
    ("웹판 공개 주소 https://caaaaat0408.github.io/pangwanbi-budget/ (헤드리스 Edge)", "정상", "준비 8초(캐시), 예시1 6,106,656 동일, 보고서 xlsx 123,567바이트, 콘솔 오류 0"),
    ("웹판에서 실제 현장 조건(A36BL 조건을 2027 로) 계산 ↔ PC 엔진", "동일", "두 쪽 모두 7,549,085천원"),
    ("웹판(저장소 docs 로컬 서빙) E2E", "정상", "예시1 산식 213칸, 설계개요 엑셀 8타입 일치, 보고서 xlsx 생성. 그림 PDF 브라우저 OCR 은 세대수 2/10만 맞음(알려진 한계, 70초)"),
    ("GitHub Pages 배포본 = 저장소 docs/index.html", "동일", "md5 일치(줄바꿈 제외), manifest synced 2026-11-04 14:22"),
    ("연도별 기준안으로 양식 생성(build_template) 2023~2026", "생성됨", "변경/확인필요: 2026 6/7 · 2025 13/14 · 2024 19/21 · 2023 33/34 — '확인 필요' 단가는 2027 값 유지(아래 '연도별 기준단가' 시트)"),
    ("도구로 만든 2027 파일(평택고덕 A66BL) 입력값 재계산", "동일", "2,611,020천원 — 양식 변경(세목 합치기) 뒤에도 같은 값"),
]
for i, (name, res, note) in enumerate(tests, start=1):
    row += 1
    put(ws, row, 1, i, al=C)
    put(ws, row, 2, name)
    c = put(ws, row, 3, res, al=C, bold=True, fill=FILL_OK)
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=9)
    for cc in range(4, 10):
        ws.cell(row, cc).border = BOX
    put(ws, row, 10, note)

row += 2
ws.cell(row, 1, "2. 과거 제출본 ↔ 도구 계산 비교 (현장별)").font = F(size=11, bold=True)
row += 1
ws.cell(row, 1, "판정: 일치(±5%) / 근접(±15%) / 차이 큼 / 소액·특수(분양·M/H 활동 없는 소액 파일) / 비교 제외(예산 0). 실제 예산은 제출본 월별 금액 합(F열 합계와 다른 줄은 월별 합 사용). 2023 제출본은 기준안 구조가 달라 참고용.").font = F(color="595959")
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
row += 1
hdr(ws, row, ["연도", "제출본(파일)", "유형", "실제 예산", "도구 기본값", "도구 조정", "차이(조정-실제)", "비율", "판정", "주요 차이 원인"])
first_data = row + 1
for r in sorted(adj, key=lambda x: (-x["year"], x["file"])):
    row += 1
    b = basic.get(r["key"])
    kinds = "·".join(KIND_KO.get(x, x) for x in r["kinds"]) or "(활동 없음)"
    verdict, fill = judge(r)
    put(ws, row, 1, r["year"], al=C)
    put(ws, row, 2, r["file"].replace(".xlsx", ""))
    put(ws, row, 3, kinds, al=C)
    put(ws, row, 4, k(r["actual_total"]), nf=NF)
    put(ws, row, 5, k(b["tool_total"]) if b else None, nf=NF)
    put(ws, row, 6, k(r["tool_total"]), nf=NF)
    put(ws, row, 7, f"=F{row}-D{row}", nf=NF)
    put(ws, row, 8, f'=IF(D{row}=0,"-",F{row}/D{row})', nf="0.0%")
    put(ws, row, 9, verdict, al=C, fill=fill, bold=True)
    reason = manual_note(r["file"])
    auto = top_lines(r)
    note = (reason + (" ▶ " if reason and auto else "") + auto) if (reason or auto) else ""
    if r.get("inconsistent"):
        note = f"[제출본 F열 합계 {k(r['actual_total_file']):,} ≠ 월별 합 {k(r['actual_total']):,}] " + note
    put(ws, row, 10, note)
    ws.row_dimensions[row].height = max(15, min(60, 15 * (1 + len(note) // 95)))
last_data = row
row += 1
put(ws, row, 2, "합계(2024~2027, 비교 제외 빼고)", bold=True, fill=FILL_H)
put(ws, row, 4, f'=SUMIFS(D{first_data}:D{last_data},A{first_data}:A{last_data},">=2024",I{first_data}:I{last_data},"<>비교 제외")', nf=NF, bold=True, fill=FILL_H)
put(ws, row, 5, f'=SUMIFS(E{first_data}:E{last_data},A{first_data}:A{last_data},">=2024",I{first_data}:I{last_data},"<>비교 제외")', nf=NF, bold=True, fill=FILL_H)
put(ws, row, 6, f'=SUMIFS(F{first_data}:F{last_data},A{first_data}:A{last_data},">=2024",I{first_data}:I{last_data},"<>비교 제외")', nf=NF, bold=True, fill=FILL_H)
put(ws, row, 7, f"=F{row}-D{row}", nf=NF, bold=True, fill=FILL_H)
put(ws, row, 8, f'=IF(D{row}=0,"-",F{row}/D{row})', nf="0.0%", bold=True, fill=FILL_H)
for cc in (1, 3, 9, 10):
    ws.cell(row, cc).border = BOX
    ws.cell(row, cc).fill = FILL_H
row += 1
put(ws, row, 2, "일치·근접 현장 수 / 비교 대상(2024~2027 활동 현장)", bold=True, fill=FILL_H)
put(ws, row, 4, f'=COUNTIFS(A{first_data}:A{last_data},">=2024",I{first_data}:I{last_data},"일치(±5%)")+COUNTIFS(A{first_data}:A{last_data},">=2024",I{first_data}:I{last_data},"근접(±15%)")', nf="0", bold=True, fill=FILL_H)
put(ws, row, 5, f'=COUNTIFS(A{first_data}:A{last_data},">=2024",I{first_data}:I{last_data},"<>비교 제외",I{first_data}:I{last_data},"<>소액·특수(참고)")', nf="0", bold=True, fill=FILL_H)
for cc in (1, 3, 6, 7, 8, 9, 10):
    ws.cell(row, cc).border = BOX
    ws.cell(row, cc).fill = FILL_H

row += 2
ws.cell(row, 1, "3. 발견 사항 · 개선 제안").font = F(size=11, bold=True)
findings = [
    ("동작", "PC 판·웹판 모두 정상 동작. 같은 입력에 대해 웹판(브라우저 Python)과 PC 엔진 결과가 1원 단위까지 같다. 웹판 그림 PDF 설계개요 읽기만 알려진 한계(엑셀 설계개요 권장)."),
    ("정확도", "2026 평택고덕 A36BL(신규분양) 99.1%(기본값)·99.8%(조정), 2026 오산서2(잔여분양) 95.3%, 2024 오산세교 99.5%, 2024 과천(입주) 95.4%, 2025 오산서2 103.6%, 2025 화성비봉(입주) 105%. "
            "신규분양·잔여분양·입주 현장은 현장 조건만 넣어도 제출본과 ±5~15% 안에 들어오고, 남는 차이는 대부분 수량 선택(대형마트 주수·현수막 주수·선물 인원)과 아래 규칙 차이."),
    ("규칙 1", "보험료 반영월이 M/H 운영기간 밖이면 0 이 된다. 제출본은 사전영업 전월(견본주택 준공 시점)에 보험료를 넣는 경우가 많음(2024 남양·오산세교) → '보험월' 입력이 M/H 기간 밖이어도 반영되게 하거나 '사전영업 전월' 선택지 추가 제안."),
    ("규칙 2", "분양사무실 운영 기간을 넣으면 통신·우편·복합기·관리비·소모품·발급·보안 7개가 모두 들어간다. 조합 사무실 안에서 운영하는 현장(김포북변·태안·이천중리B2BL)은 통신·우편(·소모품)만 편성 → 항목별 Y/N(반영 항목) 제안. 복합기·보안은 금액 0 으로 끌 수 있지만 관리비·소모품·발급·우편은 불가."),
    ("규칙 3", "12월 입주개시(2024 화성비봉)면 입주매니저·건축물관리대장·이전등기가 다음 해 1월로 넘어가 그 해 예산에서 빠진다(-85,643천원). 제출본은 입주개시월에 편성 → 연말 입주 현장용 '입주개시월 반영' 선택지 제안."),
    ("규칙 4", "조직분양 단가가 연중에 바뀌는 현장(2026 오산서2: 1~7월 8,000 → 8~12월 10,000)은 한 단가만 적용 가능(-432,000천원) → 단가 변경월 입력 제안."),
    ("규칙 5", "등기 우편은 계약누계 × 4.5천원 × 2회(6·12월)인데 제출본은 총세대 × 3~4건을 쓴다(2026 A36BL 10,030 ↔ 도구 3,343) → 발송 건수 입력 제안. 인지세는 도구가 기준안대로 계약세대 × 75천원을 넣지만 2023~2025 제출본은 대부분 미편성(2026 부터 편성)."),
    ("기준안 반영", "옛 기준안(2024·2023)으로 양식을 만들면 중개수수료(15,000 단일값)·카페테리아(일당제)·잔여판촉 식대(일당 표기)·상가 MGM 등을 못 읽어 2027 값이 남는다(기준안 반영결과 시트에 '확인 필요'로 표시됨). 2026·2025 기준안은 7개/14개만 확인 필요(대납이자·가구·도급 인지세 등 구조적 항목)."),
    ("옛 항목", "컨테이너·대기용 텐트·외부난방기·에어렉스, 실내화, 네트워크 공사, 네일아트 이벤트, 방역, 대봉투, 보존등기, 청약대행(1,000+3천원/세대)은 2027 기준안에서 빠져 도구에 없다 — 2025 이전 제출본과 비교할 때 차이로 나타남(필요하면 직접입력 줄)."),
    ("제출본 오류", "2026 A33BL 조직분양 현수막 산식 =40000*12000*4 (장수 1,200 → 12,000, 3,456,000천원 과다, 이미 알린 건). 2024 남양뉴타운1차(보류 초안) F열 합계 ≠ 월별 합 6줄. 2025 오산서2 발급수수료 F열 1,918,800 ≠ 월별 합 1,978,800. 2026 화성비봉 분양율 시트에 743세대 기본 곡선 잔존(예산 영향 없음)."),
    ("비교 방법 한계", "현장 조건은 제출본에서 자동으로 읽어 넣었다(오픈월·사전영업·조직 시작/비율·입주월·사무실 기간 등). 급지는 조직 단가로 추정, 공급면적은 공가 시트 값 또는 전용×1.33. 자동 추정이 어려운 특수 현장(과천 재분양 3세대, 오산세교1차 잔여 특례)은 '현장 특수 항목'으로 분류."),
]
for tag, text in findings:
    row += 1
    put(ws, row, 1, tag, al=C, bold=True, fill=FILL_H)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=1)
    c = put(ws, row, 2, text)
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=10)
    ws.row_dimensions[row].height = max(18, 15 * (1 + len(text) // 150))
ws.freeze_panes = "A4"
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A3
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "1:3"

# ─────────────────────────────── 항목별 시트
w2 = wb.create_sheet("항목별 비교")
hdr(w2, 1, ["연도", "제출본", "항목", "실제(천원)", "도구 기본값", "도구 조정", "차이(조정-실제)", "비율"], [6, 46, 12, 13, 13, 13, 14, 8])
r2 = 1
for r in sorted(adj, key=lambda x: (-x["year"], x["file"])):
    b = basic.get(r["key"])
    bmap = {it["item"]: it["tool"] for it in (b["items"] if b else [])}
    for it in r["items"]:
        if not it["actual"] and not it["tool"] and not bmap.get(it["item"]):
            continue
        r2 += 1
        put(w2, r2, 1, r["year"], al=C)
        put(w2, r2, 2, r["file"].replace(".xlsx", ""))
        put(w2, r2, 3, it["item"], al=C)
        put(w2, r2, 4, k(it["actual"]), nf=NF)
        put(w2, r2, 5, k(bmap.get(it["item"], 0)), nf=NF)
        put(w2, r2, 6, k(it["tool"]), nf=NF)
        put(w2, r2, 7, f"=F{r2}-D{r2}", nf=NF)
        put(w2, r2, 8, f'=IF(D{r2}=0,"-",F{r2}/D{r2})', nf="0%")
w2.freeze_panes = "A2"
w2.auto_filter.ref = f"A1:H{r2}"

# ─────────────────────────────── 줄별 차이 시트
w3 = wb.create_sheet("줄별 차이(조정)")
hdr(w3, 1, ["연도", "제출본", "항목", "줄(세목·내용 키)", "실제(천원)", "도구 조정", "차이", "분류", "차이 유형", "제출본 산출근거", "도구 산출근거"], [6, 40, 10, 20, 12, 12, 12, 9, 26, 60, 60])
r3 = 1
for r in sorted(adj, key=lambda x: (-x["year"], x["file"])):
    for ln in sorted(r["lines"], key=lambda x: -abs(x["diff"])):
        r3 += 1
        put(w3, r3, 1, r["year"], al=C)
        put(w3, r3, 2, r["file"].replace(".xlsx", ""))
        put(w3, r3, 3, ln["item"], al=C)
        put(w3, r3, 4, ln["key"])
        put(w3, r3, 5, k(ln["actual"]), nf=NF)
        put(w3, r3, 6, k(ln["tool"]), nf=NF)
        put(w3, r3, 7, f"=F{r3}-E{r3}", nf=NF)
        cat = category(ln)
        put(w3, r3, 8, ln["class"], al=C, fill=FILL_OK if ln["class"] == "일치" else None)
        put(w3, r3, 9, cat)
        put(w3, r3, 10, ln["actual_g"])
        put(w3, r3, 11, ln["tool_g"])
w3.freeze_panes = "A2"
w3.auto_filter.ref = f"A1:K{r3}"

# ─────────────────────────────── 입력 조건 시트
w4 = wb.create_sheet("입력 조건")
hdr(w4, 1, ["연도", "제출본", "유형", "세대", "도구에 넣은 현장 조건(기본값과 다른 칸만, 조정 모드)", "세대 구성(타입·세대·공급㎡)", "월별 분양 / 입주 세대", "추정 메모"], [6, 40, 16, 7, 70, 40, 46, 60])
r4 = 1
SKIP = {"현장명", "지역", "사업구분", "상품", "작성일", "담당소장", "착공", "준공", "시행", "시공"}
for r in sorted(adj, key=lambda x: (-x["year"], x["file"])):
    r4 += 1
    f = r["inputs"]["fields"]
    cond = ", ".join(f"{kk}={vv}" for kk, vv in f.items() if kk not in SKIP and vv not in (None, [None, None], "", 0, "N") and not (isinstance(vv, list) and all(x is None for x in vv)))
    m = r["inputs"]["monthly"]
    ms = "분양 " + str([x or 0 for x in (m.get("분양") or [])]) + (" 전년 " + str(m.get("전년분양")) if any(m.get("전년분양") or []) else "") + \
         (" / 입주 " + str([x or 0 for x in m.get("입주")]) if any(m.get("입주") or []) else "")
    put(w4, r4, 1, r["year"], al=C)
    put(w4, r4, 2, r["file"].replace(".xlsx", ""))
    put(w4, r4, 3, "·".join(KIND_KO.get(x, x) for x in r["kinds"]) or "(활동 없음)", al=C)
    put(w4, r4, 4, r["units"], nf="#,##0")
    put(w4, r4, 5, cond)
    put(w4, r4, 6, ", ".join(f"{t[0]}:{t[1]}({t[2]})" for t in r["inputs"]["types"]))
    put(w4, r4, 7, ms)
    put(w4, r4, 8, " | ".join(r["notes"]))
    w4.row_dimensions[r4].height = max(18, 13 * (1 + max(len(cond) // 70, len(ms) // 46, len(" | ".join(r["notes"])) // 60)))
w4.freeze_panes = "A2"

# ─────────────────────────────── 연도별 기준단가 시트
w5 = wb.create_sheet("연도별 기준단가")
years = [2027, 2026, 2025, 2024, 2023]
infos = {}
for y in years:
    tpl = os.path.join(ROOT, "판관비 예산안 자동작성 양식_2027.xlsx") if y == 2027 else os.path.join(SP, f"양식_{y}.xlsx")
    infos[y] = site_io.template_info(xlcalc.Model(tpl))
hdr(w5, 1, ["단가 이름", "항목", "내용", "단위"] + [f"{y} 양식 값" for y in years] + [f"{y} 반영결과" for y in years], [22, 12, 46, 12] + [11] * 5 + [34] * 5)
w5.cell(2, 1, "기준안 파일").font = F(bold=True)
for j, y in enumerate(years):
    w5.cell(2, 5 + j, infos[y]["kijun"]).font = F(size=8)
    w5.cell(2, 5 + j).alignment = Alignment(wrap_text=True, vertical="top")
w5.row_dimensions[2].height = 40
base = infos[2027]["params"]
r5 = 2
for i, p in enumerate(base):
    if "key" not in p:
        r5 += 1
        c = w5.cell(r5, 1, "■ " + p["section"])
        c.font = F(bold=True)
        c.fill = FILL_H
        for cc in range(2, 15):
            w5.cell(r5, cc).fill = FILL_H
        continue
    r5 += 1
    put(w5, r5, 1, p["key"])
    put(w5, r5, 2, p["item"])
    put(w5, r5, 3, p["desc"])
    put(w5, r5, 4, p["unit"], al=C)
    for j, y in enumerate(years):
        q = next((x for x in infos[y]["params"] if x.get("key") == p["key"]), None)
        v = q["value"] if q else None
        cell = put(w5, r5, 5 + j, v, nf="#,##0.##")
        if y != 2027 and q and base and q["value"] != p["value"]:
            cell.fill = FILL_NEAR
        res = (q or {}).get("result") or ""
        c2 = put(w5, r5, 10 + j, res)
        if str(res).startswith("확인"):
            c2.fill = FILL_BAD
w5.freeze_panes = "E3"

# ─────────────────────────────── 테스트 로그 시트
w6 = wb.create_sheet("테스트 로그")
w6.column_dimensions["A"].width = 160
logs = []
for name in ("webtest/web_public_out.txt", "webtest/web_e2e_out.txt"):
    p = os.path.join(SP, name)
    if os.path.exists(p):
        logs.append(f"■ {name}")
        logs += [ln.rstrip()[:300] for ln in open(p, encoding="utf-8", errors="replace").read().splitlines() if "OCR words" not in ln]
extra = os.path.join(SP, "test_logs.txt")
if os.path.exists(extra):
    logs.append("■ PC 판·엔진 회귀 테스트")
    logs += [ln.rstrip()[:300] for ln in open(extra, encoding="utf-8", errors="replace").read().splitlines()]
for i, ln in enumerate(logs, start=1):
    c = w6.cell(i, 1, ("'" + ln) if ln.startswith("=") else ln)
    c.font = Font(name="Consolas", size=9)

wb.save(OUT)
print("저장:", OUT)
