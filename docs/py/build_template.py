# -*- coding: utf-8 -*-
"""판관비 예산안 자동작성 양식(엑셀) 만들기.

사용: python build_template.py <기준안.xlsx> <출력.xlsx>

양식 구조
  현장입력  → 계산(월별 단계) → 공식 양식 시트(27년 예산·예산안 보고서·분양율,입주율·분양대행·공가·대납이자·기본정보)
  기준단가  : 기준안의 단가. 모든 수식이 이 표의 이름(기준_…)을 참조한다.
금액은 원 단위로 계산하고 '#,###,' 서식으로 천원 표시(기존 PM팀 예산편성 파일과 같은 방식).
"""
import os
import sys

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from params import PARAMS                     # noqa: E402
import kijun_parser                           # noqa: E402

FONT = "맑은 고딕"
MC = [CL(c) for c in range(10, 22)]           # J..U = 1~12월 (계산·27년 예산·분양대행 시트)
BUDGET = "27년 예산"                            # 기준안 반영 때 연도에 맞게 이름을 바꾼다
NONE = 999999                                  # '해당 없음' 월 번호

# ── 서식 ─────────────────────────────────────────────────────────
def font(size=10, bold=False, color=None, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)

F10, F10B, F9, F9G = font(), font(bold=True), font(9), font(9, color="808080")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")       # 입력
FILL_DEF = PatternFill("solid", fgColor="E2EFDA")      # 자동 기본값(덮어쓰기 가능)
FILL_BUD = PatternFill("solid", fgColor="DDEBF7")      # 예산 열
FILL_SUB = PatternFill("solid", fgColor="F2F2F2")      # 소계
FILL_HEAD = PatternFill("solid", fgColor="D9E1F2")     # 도구 시트 머리글
FILL_SEC = PatternFill("solid", fgColor="BDD7EE")      # 구역 제목
FILL_TOT = PatternFill("solid", fgColor="FCE4D6")      # 합계
thin, hair, med, dbl = (Side(style=s) for s in ("thin", "hair", "medium", "double"))
NUMF = "#,###,"                                         # 원 → 천원 표시
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")
VC = Alignment(vertical="center")


def box(ws, rng, side=thin, inner=hair):
    """범위에 테두리: 바깥 side, 안쪽 inner"""
    rows = list(ws[rng])
    for i, row in enumerate(rows):
        for j, c in enumerate(row):
            c.border = Border(
                left=side if j == 0 else inner, right=side if j == len(row) - 1 else inner,
                top=side if i == 0 else inner, bottom=side if i == len(rows) - 1 else inner)


def setw(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def put(ws, addr, value, f=F10, fill=None, al=None, nf=None):
    c = ws[addr]
    c.value = value
    c.font = f
    if fill:
        c.fill = fill
    if al:
        c.alignment = al
    if nf:
        c.number_format = nf
    return c


class Names:
    def __init__(self, wb):
        self.wb = wb
        self.ref = {}

    def add(self, name, sheet, addr):
        ref = f"'{sheet}'!{addr}"
        self.ref[name] = ref
        self.wb.defined_names[name] = DefinedName(name=name, attr_text=ref)


# ═════════════════════════════════════════════════════════════════
def build(out_path, kijun_path=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    names = Names(wb)
    order = ["사용법", "현장입력", "기본정보", "예산안 보고서", BUDGET, "분양율,입주율", "분양대행 및 촉진수수료",
             "공가세대관리비", "대납이자", "기준단가", "계산"]
    sh = {n: wb.create_sheet(n) for n in order}

    P = build_kijun(sh["기준단가"], names)          # 기준_이름 → 셀
    I = build_input(sh["현장입력"], names)          # 입력 정보
    K = build_calc(sh["계산"], names, I)            # 계산 행 번호
    build_rates_sheet(sh["분양율,입주율"], K)
    pay = build_agency(sh["분양대행 및 촉진수수료"], K, I)
    vac = build_vacancy(sh["공가세대관리비"], K, I)
    build_interest(sh["대납이자"])
    groups = build_budget(sh[BUDGET], K, pay, vac)
    build_report(sh["예산안 보고서"], groups)
    build_basic(sh["기본정보"], I)
    build_guide(sh["사용법"])
    finish_input(sh["현장입력"], I, K, groups)

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = ws.title in ("계산",)
    sh["사용법"].sheet_properties.tabColor = "70AD47"
    sh["현장입력"].sheet_properties.tabColor = "FFC000"
    sh["기준단가"].sheet_properties.tabColor = "5B9BD5"
    sh["계산"].sheet_properties.tabColor = "A6A6A6"
    wb.active = 1
    wb.calculation.fullCalcOnLoad = True

    if kijun_path:
        info = kijun_parser.apply_to_workbook(wb, kijun_path)
        # 처음 만들 때는 기본값과 같은 값이 대부분이므로 '변경' 표시를 지움
        ws = sh["기준단가"]
        for r in range(kijun_parser.FIRST_ROW, ws.max_row + 1):
            if ws.cell(r, kijun_parser.COL["name"]).value and ws.cell(r, kijun_parser.COL["result"]).value == "동일":
                ws.cell(r, kijun_parser.COL["result"]).value = "기준안 값"
        wb.move_sheet("기준안 반영결과", offset=wb.sheetnames.index("기준단가") + 1 - wb.sheetnames.index("기준안 반영결과"))
        print("기준안 반영:", {k: sum(1 for x in info["report"] if x[0] == k) for k in ("변경", "동일", "확인")},
              "미사용 줄", len(info["unused"]))
    wb.save(out_path)
    return out_path


# ═════════════════════════════════════════════════════════════════
# 기준단가
def build_kijun(ws, names):
    C = kijun_parser.COL
    put(ws, "A1", "PM팀 판관비 예산 기준단가", font(16, True))
    put(ws, "A2", "적용 기준안", F10B)
    put(ws, "C2", "", F10)
    put(ws, "F2", "작성일", F10B, al=RIGHT)
    put(ws, "G2", None, F10, nf="yyyy.mm.dd")
    put(ws, "I2", "예산연도 →", F10B, al=RIGHT)
    put(ws, "J2", 2027, font(12, True, "C00000"), FILL_IN, CENTER)
    names.add("기준_예산연도", "기준단가", "$J$2")
    put(ws, "A3", "※ F열(값)만 고치면 이 파일의 모든 계산에 바로 반영됩니다.  매년 '기준안 반영.bat' 으로 새 기준안을 "
                  "읽으면 F·I·J열이 자동으로 갱신됩니다(주황=이번에 바뀐 값, 빨강=확인 필요).", F9)
    put(ws, "I3", "반영 일시 →", F9, al=RIGHT)
    heads = ["No", "구분", "항목", "세목", "기준 내용", "값", "단위", "자동작성 적용 방식", "기준안 원문", "반영 결과",
             "이름", "찾기:구분", "찾기:항목", "찾기:줄", "찾기:위치", "찾기:단위", "찾기:순번", "찾기:원천"]
    for i, h in enumerate(heads, start=1):
        c = ws.cell(5, i, h)
        c.font, c.fill, c.alignment = F10B, FILL_HEAD, CENTER
    ref = {}
    r = kijun_parser.FIRST_ROW
    no = 0
    for p in PARAMS:
        if p[0] == "■":
            ws.cell(r, 1, f"■ {p[1]}").font = F10B
            for c in range(1, 11):
                ws.cell(r, c).fill = FILL_SEC
            r += 1
            continue
        key, item, sub, desc, val, unit, rule, spec = p
        no += 1
        sec = "분양사무실" if spec and spec[0] == "사무실" or key.startswith("사무실_") else "분양/입주"
        vals = [no, sec, item, sub, desc, val, unit, rule, None, None, key]
        if spec:
            vals += [spec[0], spec[1], spec[2], spec[3], spec[4], spec[5], spec[6] if len(spec) > 6 else None]
        for i, v in enumerate(vals, start=1):
            c = ws.cell(r, i, v)
            c.font = F9 if i >= 8 else F10
            c.alignment = LEFT if i in (5, 8, 9, 10) else Alignment(vertical="center")
        vc = ws.cell(r, C["val"])
        vc.fill = kijun_parser.FILL_VALUE
        vc.number_format = "#,##0.0##" if isinstance(val, float) else "#,##0"
        vc.font = F10B
        if spec is None:
            ws.cell(r, C["result"], "고정값(기준안에서 읽지 않음)").font = F9G
        name = "기준_" + key
        names.add(name, "기준단가", f"$F${r}")
        ref[key] = name
        r += 1
    box(ws, f"A5:J{r - 1}", thin, hair)
    setw(ws, dict(A=5, B=10, C=11, D=11, E=40, F=10, G=11, H=36, I=62, J=30, K=18, L=7, M=10, N=26, O=16, P=10, Q=6, R=6))
    for col in "KLMNOPQR":
        ws.column_dimensions[col].outlineLevel = 1
        ws.column_dimensions[col].hidden = True
    ws.freeze_panes = "F6"
    return ref


# ═════════════════════════════════════════════════════════════════
# 현장입력
INPUTS = [
    ("1. 기본 사항", None),
    ("예산연도", "예산연도", "int", "=기준_예산연도", "년", "기준안 작성일의 다음 해(기준단가 시트). 다른 해로 쓸 때만 고치세요.", None, "def"),
    ("현장명", "현장명", "text", None, "", "예) 평택고덕(A33BL) — 보고서 제목·시트 머리에 표시", None),
    ("작성일", "작성일", "date", None, "", "보고서 '( 2026. 11. 26 )' 표시용. 비우면 오늘 날짜", None),
    ("지역", "지역", "list", "수도권", "", "여비교통비·M/H 주차장 임차료·M/H 중개수수료 단가 구분", "수도권,비수도권"),
    ("급지", "분양성(급지)", "list", "C급", "",
     "조직분양 수수료(B 8,000·C 10,000·D 15,000천원/세대), 부동산MGM·청약/당첨자/계약 선물(C~D급), 오픈 선물(B~D급)",
     "A급,B급,C급,D급"),
    ("사업구분", "사업구분", "list", "자체사업", "", "인지세: 자체사업 75천원 / 도급현장 20천원 (세대당)", "자체사업,도급사업"),
    ("상품", "상품", "list", "아파트", "", "분양가 산출용역: 아파트 50,000 / 주상복합 70,000천원", "아파트,주상복합"),
    ("담당소장", "담당소장", "text", None, "", "기본정보 시트", None),
    ("착공", "착공년월", "text", None, "", "기본정보 시트 (예: 2026.09)", None),
    ("준공", "준공년월", "text", None, "", "기본정보 시트", None),
    ("시행", "시행법인", "text", None, "", "기본정보 시트", None),
    ("시공", "시공법인", "text", None, "", "비우면 시행법인과 같음", None),
    ("2. 세대 구성", "TYPES"),
    ("상가", "상가 호실 수", "int", 0, "호실", "인지세(상가 포함)·이전등기비용에 더함", None),
    ("3. 분양 일정 (견본주택 오픈월 = D)", None),
    ("오픈", "오픈월(분양월)", "ym", None, "연 / 월",
     "견본주택 오픈 연·월. 비우면 분양 관련 항목은 계산하지 않음(입주·분양사무실만 있는 현장)", None),
    ("사전영업", "사전영업 기간", "int", 1, "개월", "오픈 전 사전영업 개월수(기준: 1개월). M/H 운영·보험료·무인경비 공사·M/H 중개수수료 시작", None),
    ("MH종료", "M/H 운영 종료월", "ym", None, "연 / 월", "비우면 연말까지 운영", None),
    ("목표분양률", "목표 분양률(연말 누계)", "pct", 1, "", "분양율 시트 월별 기본값: 오픈 다음 달 30% → 50·60·70·80·85·90·95·100% (목표에서 멈춤)", None),
    ("전년분양", "전년 9월말 분양누계", "int", 0, "세대", "기분양 현장만. 분양율 시트 '9월누계' 칸", None),
    ("분양대행", "분양대행수수료 적용", "yn", "Y", "", "N 이면 분양대행수수료(2,000천원/세대) 제외", None),
    ("조직", "조직분양 진행", "yn", "N", "", "Y 이면 조직분양 수수료·현수막·잔여판촉 식대·청소 2명 반영", None),
    ("조직시작", "조직분양 시작월", "ym", None, "연 / 월", "비우면 오픈+2개월(기준안 D+2M)", None),
    ("조직종료", "조직분양 종료월", "ym", None, "연 / 월", "비우면 연말까지", None),
    ("조직비율", "조직분양 비율(총세대 대비)", "pct", None, "", "예) 70%. 비우면 시작월 이후 계약분 전부 조직분양", None),
    ("조직단가", "조직분양 수수료 단가", "num", None, "천원/세대", "비우면 급지 기준 단가", None),
    ("4. 견본주택 운영 조건", None),
    ("전기미인입", "전기 미인입 현장", "yn", "N", "", "Y 이면 오픈월 30,000천원(오픈 1개월 限)", None),
    ("주차장", "M/H 주차장 임차(오픈 1개월)", "yn", "Y", "", "", None),
    ("MH중개", "M/H 중개수수료", "yn", "Y", "", "사전영업 시작월", None),
    ("분양가용역", "분양가 산출용역", "yn", "Y", "", "오픈 전월", None),
    ("유니트", "유니트 수", "int", 3, "개", "청소 인원: 오픈 1개월차 3개 이하 3명/4개 이상 4명, 2개월차 2명/3명", None),
    ("여직원", "여직원 수(피복비)", "int", 0, "명", "피복비는 기본 0 — 신규 또는 현장 소속법인 여직원이 있을 때만 인원을 넣으면 1월에 반영", None),
    ("숙소분양", "숙소 호실 수(분양)", "int", 3, "호실", "사전영업 시작월에 3개월분 일시 + 다음 달부터 숙소관리비 3개월", None),
    ("PC", "사무실 PC 대수", "int", 6, "대", "M/H 운영 기간", None),
    ("복합기", "복합기 대수", "int", 1, "대", "M/H 운영·입주지원센터 기간", None),
    ("보험월", "보험료 반영월", "month", None, "월", "비우면 사전영업 시작월(이후 매년 같은 달)", None),
    ("예비비율","예비입주자 내용증명 비율", "pct", 5, "", "내용증명 건수 = 총세대 × 비율 (과거 작성: 500%)", None),
    ("등기월1", "사업승인·제품변경 안내 등기 ①", "month", 6, "월", "그 달 계약누계 × 4.5천원 (입주 전까지)", None),
    ("등기월2", "사업승인·제품변경 안내 등기 ②", "month", 12, "월", "비우면 1회만", None),
    ("아르바이트", "아르바이트", "yn", "Y", "", "오픈월부터 4개월, 3,000천원/월", None),
    ("카페테리아", "카페테리아", "yn", "Y", "", "오픈월 20,000천원", None),
    ("무순위", "무순위청약 전산추첨 대행", "yn", "Y", "", "오픈월", None),
    ("계약취소", "계약취소주택 전산추첨 대행", "yn", "Y", "", "오픈월", None),
    ("5. 판촉물·홍보 수량", None),
    ("갑티슈사전", "갑티슈(사전영업)", "int", 10000, "개", "오픈 전월", None),
    ("갑티슈오픈", "갑티슈(오픈)", "int", 10000, "개", "오픈월", None),
    ("물티슈", "물티슈(사전영업)", "int", 20000, "개", "오픈 전월", None),
    ("종이컵", "종이컵(사전영업)", "int", 20, "박스", "기준 50천원/박스. 과거 파일은 '200박스'로 적고 1,000천원으로 계산", None),
    ("크린홈", "크린홈세트(사전영업)", "int", 10000, "개", "오픈 전월", None),
    ("마스크", "마스크(사전영업)", "int", 10000, "개", "오픈 전월", None),
    ("대형마트", "대형마트 홍보", "int", 2, "주", "기준안 '3개소 × 2주 기준' (과거 작성 5주)", None),
    ("현수막주수", "조직분양 현수막 월 게시 주수", "int", 4, "주/월", "조직분양 기간 × 1,200장/주 × 40천원", None),
    ("현수막감축", "현수막 감축 시작월", "ym", None, "연 / 월", "이 달부터 아래 주수로 줄임(과거 작성: 분양 후반 2주/월). 비우면 감축 없음", None),
    ("현수막감축주수", "감축 후 월 게시 주수", "int", 2, "주/월", "", None),
    ("오픈선물인원", "오픈 선물 인원", "int", None, "명", "비우면 기준안 1,000명", None),
    ("청약선물인원", "청약 선물 인원", "int", None, "명", "비우면 기준안 1,000명", None),
    ("잔여갑티슈", "잔여판촉 갑티슈 수량", "int", 0, "개", "기분양 현장 잔여판촉용", None),
    ("잔여판촉", "잔여판촉 갑티슈 반영월", "ym", None, "연 / 월", "", None),
    ("6. 상가·기타 판촉", None),
    ("상가매출", "상가 분양대행 대상 매출액", "num", 0, "천원", "단지내 10% / 주상복합 20%", None),
    ("상가유형", "상가 유형", "list", "단지내 상가", "", "", "단지내 상가,주상복합 상가"),
    ("상가월", "상가 계약(대행수수료·인지세) 반영월", "ym", None, "연 / 월", "비우면 오픈+1개월", None),
    ("상가MGM", "상가 부동산MGM 호실수", "int", 0, "호실", "10,000천원/호실", None),
    ("상가MGM월", "상가 부동산MGM 반영월", "ym", None, "연 / 월", "비우면 오픈+1개월", None),
    ("해지세대", "계약해지 예정세대(장기대여금 대체)", "int", 0, "세대", "이자후불제 현장, 20,000천원/세대", None),
    ("해지월", "계약해지 반영월", "ym", None, "연 / 월", "", None),
    ("상담사", "재분양 상담사", "yn", "N", "", "6,000천원/월 × 3개월", None),
    ("상담사시작", "재분양 상담사 시작월", "ym", None, "연 / 월", "", None),
    ("무이자", "중도금 무이자(대납이자 반영)", "yn", "N", "", "Y 이면 대납이자 시트 이자를 27년 예산에 반영", None),
    ("최고이율", "기분양 현장 최근 결재 최고이율", "pct", None, "", "입력하면 + 가산금리(0.5%p)를 대납이자 이율로, 비우면 신규 분양현장 이율", None),
    ("7. 입주 일정", None),
    ("사전점검", "사전점검월", "ym", None, "연 / 월", "사전점검행사·소모품·입주지원센터 운영 시작", None),
    ("입주", "입주개시월", "ym", None, "연 / 월", "입주지원센터 구축·문고·입주상담사·공가세대관리비 시작", None),
    ("지정기간", "입주지정기간", "int", 2, "개월", "식대·여비·통신·입주상담사: 입주개시~지정기간+1개월", None),
    ("전년입주", "전년 9월말 입주누계", "int", 0, "세대", "기입주 현장만", None),
    ("입주율", "입주율 전망", "list", "양호", "", "입주매니저: 양호 1명/세대, 열위 1.2명/세대", "양호,열위"),
    ("시행사2", "시행사 2개 이상", "yn", "N", "", "이전등기비용 1.2천원 → 2.4천원", None),
    ("입주촉진", "입주촉진책", "yn", "Y", "", "입주개시월 10,000천원", None),
    ("숙소입주", "입주 숙소 호실 수", "int", 2, "호실", "입주개시월 3개월분 일시 + 숙소관리비", None),
    ("입주PC", "입주지원센터 PC 대수", "int", 4, "대", "", None),
    ("어린이집", "국공립어린이집 동의서 발송", "yn", "Y", "", "입주개시 전월 × 세대수", None),
    ("입주안내", "입주 안내 등기 횟수", "int", 2, "회", "1=사전점검 안내(사전점검 전월), 2=+입주 안내(입주 전월)", None),
    ("8. 분양사무실 운영", None),
    ("사무실연중", "연중 운영(전년부터 계속)", "yn", "N", "",
     "Y 이면 시작월 없이 1월부터 운영으로 계산(사무실 이전 운반비 없음) — 1년 내내 분양사무실만 운영하는 현장. 종료월을 넣으면 그 달까지", None),
    ("사무실시작", "분양사무실 운영 시작월", "ym", None, "연 / 월",
     "올해 새로 시작하면 그 달(M/H 철거 뒤 등 — 이 달에 사무실 이전 운반비). 전년부터 운영 중이면 '연중 운영' 체크 또는 실제 시작월(지난해도 됨)", None),
    ("사무실종료", "분양사무실 운영 종료월", "ym", None, "연 / 월", "비우면 연말까지", None),
    ("통합센터", "통합센터 운영", "yn", "N", "", "관리비 600 → 1,000천원/월", None),
    ("사무실복합기", "복합기 월 임차료", "num", "=기준_사무실_복합기", "천원/월", "300~400 탄력 적용", None, "def"),
    ("사무실보안", "보안경비 월 용역비", "num", "=기준_사무실_보안", "천원/월", "250~350 탄력 적용", None, "def"),
]
NTYPES = 20     # 세대 구성 줄 수(용마터널처럼 타입이 14개인 현장도 있음)


def build_input(ws, names):
    put(ws, "B1", "현장 입력", font(16, True))
    put(ws, "B2", "노란 칸 = 입력,  초록 칸 = 자동 기본값(덮어써도 됨).  입력하면 다른 시트가 모두 자동 계산됩니다.", F9)
    info = dict(rows={}, types=None)
    yn = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    month = DataValidation(type="whole", operator="between", formula1="1", formula2="12", allow_blank=True)
    year = DataValidation(type="whole", operator="between", formula1="2015", formula2="2050", allow_blank=True)
    for dv in (yn, month, year):
        ws.add_data_validation(dv)
    lists = {}
    r = 4
    for it in INPUTS:
        if it[1] is None or it[1] == "TYPES":
            r += 1 if r > 4 else 0
            put(ws, f"B{r}", it[0], font(11, True))
            for col in "BCDEFG":
                ws[f"{col}{r}"].fill = FILL_SEC
            r += 1
            if it[1] == "TYPES":
                for c, h in zip("CDEFG", ("타입", "세대수", "공급면적(㎡)", "공급평(자동)", "비고")):
                    put(ws, f"{c}{r}", h, F10B, FILL_HEAD, CENTER)
                r += 1
                t1 = r
                for k in range(NTYPES):
                    put(ws, f"C{r}", None, F10, FILL_IN, CENTER)
                    put(ws, f"D{r}", None, F10, FILL_IN, nf="#,##0")
                    put(ws, f"E{r}", None, F10, FILL_IN, nf="#,##0.0000")
                    put(ws, f"F{r}", f'=IF(E{r}="","",ROUNDUP(E{r}*0.3025,0))', F10, nf="#,##0")
                    if k == 0:
                        put(ws, f"G{r}", "타입명은 '84A' 처럼. 공급면적은 공가세대관리비(평당 관리비)에만 씀", F9G)
                    r += 1
                t2 = r - 1
                put(ws, f"C{r}", "합 계", F10B, al=CENTER)
                put(ws, f"D{r}", f"=SUM(D{t1}:D{t2})", F10B, nf="#,##0")
                put(ws, f"G{r}", "총세대수", F9G)
                box(ws, f"C{t1 - 1}:G{r}")
                names.add("입력_총세대", "현장입력", f"$D${r}")
                info["types"] = (t1, t2)
                r += 1
            continue
        key, label, kind, default, unit, help_, opts = it[:7]
        is_def = len(it) > 7 and it[7] == "def"
        put(ws, f"C{r}", label, F10, al=LEFT)
        fill = FILL_DEF if is_def else FILL_IN
        if kind == "ym":
            put(ws, f"D{r}", None, F10, fill, CENTER, "0")
            put(ws, f"E{r}", None, F10, fill, CENTER, "0")
            names.add(f"입력_{key}연", "현장입력", f"$D${r}")
            names.add(f"입력_{key}월", "현장입력", f"$E${r}")
            year.add(f"D{r}")
            month.add(f"E{r}")
        else:
            nf = {"int": "#,##0", "num": "#,##0", "pct": "0%", "date": "yyyy.mm.dd", "month": "0\"월\""}.get(kind)
            if key == "예산연도":
                nf = "0"
            put(ws, f"D{r}", default, F10, fill, CENTER, nf)
            names.add(f"입력_{key}", "현장입력", f"$D${r}")
            if kind == "yn":
                yn.add(f"D{r}")
            elif kind == "month":
                month.add(f"D{r}")
            elif kind == "list":
                if opts not in lists:
                    dv = DataValidation(type="list", formula1=f'"{opts}"', allow_blank=True)
                    ws.add_data_validation(dv)
                    lists[opts] = dv
                lists[opts].add(f"D{r}")
            if kind == "text":
                ws.merge_cells(f"D{r}:E{r}")
                ws[f"D{r}"].alignment = LEFT
        put(ws, f"F{r}", unit, F9, al=CENTER)
        put(ws, f"G{r}", help_, F9G, al=LEFT)
        box(ws, f"C{r}:G{r}")
        info["rows"][key] = r
        r += 1
    info["last"] = r
    setw(ws, dict(A=2, B=3, C=34, D=14, E=9, F=10, G=78, H=3))
    ws.freeze_panes = "D4"
    return info


def finish_input(ws, I, K, groups):
    """현장입력 오른쪽: 월별 단계 미리보기 + 점검 메시지 + 합계"""
    put(ws, "I4", "월별 단계 미리보기", font(11, True))
    put(ws, "I5", "구 분", F10B, FILL_HEAD, CENTER)
    for i in range(12):
        col = CL(10 + i)           # J..U
        put(ws, f"{col}5", f"{i + 1}월", F10B, FILL_HEAD, CENTER)
        put(ws, f"{col}6", f"='계산'!{MC[i]}{K['row']['단계']}", F9, al=CENTER)
        put(ws, f"{col}7", f"='분양율,입주율'!{CL(9 + i)}8", F9, al=CENTER, nf="#,##0;-#,##0;\"\"")
        put(ws, f"{col}8", f"='분양율,입주율'!{CL(9 + i)}13", F9, al=CENTER, nf="#,##0;-#,##0;\"\"")
        put(ws, f"{col}9", f"='{BUDGET}'!{MC[i]}{groups['total_row']}", F9, al=CENTER, nf=NUMF)
    for rr, t in ((6, "단계"), (7, "분양세대"), (8, "입주세대"), (9, "예산(천원)")):
        put(ws, f"I{rr}", t, F10, al=CENTER)
    box(ws, "I5:U9")
    put(ws, "I11", "연간 예산 합계(천원)", F10B)
    put(ws, "M11", f"='{BUDGET}'!F{groups['total_row']}", font(12, True, "C00000"), FILL_TOT, RIGHT, NUMF)
    ws.merge_cells("M11:O11")
    put(ws, "I13", "점검", font(11, True))
    checks = [
        '=IF(입력_현장명="","· 현장명을 입력하세요","")',
        '=IF(입력_총세대=0,"· 세대 구성(타입·세대수)을 입력하세요","")',
        '=IF(AND(입력_오픈연="",입력_입주연="",입력_사무실시작연="",입력_사무실연중<>"Y"),"· 오픈월·입주개시월·분양사무실(시작월 또는 연중 운영) 중 하나는 있어야 계산됩니다","")',
        '=IF(AND(입력_조직="Y",LEFT(입력_급지,1)="A"),"· A급은 조직분양 기준 단가가 없어 B급 단가를 씁니다","")',
        f"=IF(SUM('분양율,입주율'!E8:T8)>입력_총세대,\"· 분양세대 합이 총세대수보다 많습니다(분양율 시트)\",\"\")",
        f"=IF(SUM('분양율,입주율'!E13:T13)>입력_총세대,\"· 입주세대 합이 총세대수보다 많습니다(분양율 시트)\",\"\")",
        '=IF(입력_예산연도<>기준_예산연도,"· 예산연도가 기준안 연도("&기준_예산연도&")와 다릅니다","")',
        '=IF(AND(입력_오픈연<>"",입력_오픈월<>"",(입력_오픈연*12+입력_오픈월)<(입력_예산연도*12-24)),"· 오픈월이 2년 넘게 지났습니다 — M/H 운영 종료월·분양사무실을 확인하세요","")',
    ]
    for i, f in enumerate(checks):
        put(ws, f"I{14 + i}", f, font(10, color="C00000"))
    put(ws, f"I{14 + len(checks)}", '=IF(COUNTIF(I14:I21,"·*")=0,"✔ 입력 점검 이상 없음","")', font(10, True, "00A050"))
    setw(ws, {"I": 12, **{CL(10 + i): 8.5 for i in range(12)}})


# ═════════════════════════════════════════════════════════════════
# 계산 (월별 단계)
def ym(key, none="NONE"):
    """입력_xx연/월 → 월번호 식. none='NONE' 이면 없을 때 999999, 'END' 면 연말"""
    blank = NONE if none == "NONE" else "계산_연도*12+12"
    return f'IF(OR(입력_{key}연="",입력_{key}월=""),{blank},입력_{key}연*12+입력_{key}월)'


SCALARS = [
    ("연도", "예산연도", "=입력_예산연도"),
    ("총세대", "총세대수", "=입력_총세대"),
    ("오픈", "오픈월 번호(연×12+월)", "=" + ym("오픈")),
    ("사전개월", "사전영업 개월", '=IF(입력_사전영업="",1,입력_사전영업)'),
    ("사전시작", "사전영업 시작월 번호", "=계산_오픈-계산_사전개월"),
    ("MH종료", "M/H 운영 종료월 번호", "=" + ym("MH종료", "END")),
    ("목표율", "목표 분양률", '=IF(입력_목표분양률="",1,입력_목표분양률)'),
    ("목표세대", "목표 분양세대", "=ROUND(계산_총세대*계산_목표율,0)"),
    ("조직여부", "조직분양 여부", '=IF(입력_조직="Y",1,0)'),
    ("조직시작", "조직분양 시작월 번호", '=IF(OR(입력_조직시작연="",입력_조직시작월=""),계산_오픈+2,입력_조직시작연*12+입력_조직시작월)'),
    ("조직종료", "조직분양 종료월 번호", "=" + ym("조직종료", "END")),
    ("조직상한", "조직분양 세대 상한", '=IF(입력_조직비율="",999999,ROUND(계산_총세대*입력_조직비율,0))'),
    ("급지", "급지", '=LEFT(입력_급지,1)'),
    ("조직단가", "조직분양 수수료 단가(천원)",
     '=IF(입력_조직단가<>"",입력_조직단가,IF(계산_급지="D",기준_조직_D,IF(계산_급지="C",기준_조직_C,기준_조직_B)))'),
    ("MGM적용", "MGM·청약/당첨자/계약 선물 적용(C~D급)", '=IF(OR(계산_급지="C",계산_급지="D"),1,0)'),
    ("오픈선물적용", "오픈 선물 적용(B~D급)", '=IF(OR(계산_급지="B",계산_급지="C",계산_급지="D"),1,0)'),
    ("수도권", "수도권 여부", '=IF(입력_지역="비수도권",0,1)'),
    ("여비오픈", "여비 - 오픈 2개월(천원/월)", "=IF(계산_수도권=1,기준_여비_오픈_수도권,기준_여비_오픈_비수도권)"),
    ("여비사전", "여비 - 사전영업·오픈후 3개월", "=IF(계산_수도권=1,기준_여비_사전_수도권,기준_여비_사전_비수도권)"),
    ("여비기타", "여비 - 기타", "=IF(계산_수도권=1,기준_여비_기타_수도권,기준_여비_기타_비수도권)"),
    ("주차장", "M/H 주차장 임차료", "=IF(계산_수도권=1,기준_주차장_수도권,기준_주차장_비수도권)"),
    ("MH중개", "M/H 중개수수료", "=IF(계산_수도권=1,기준_MH중개_수도권,기준_MH중개_비수도권)"),
    ("분양가용역", "분양가 산출용역", '=IF(입력_상품="주상복합",기준_분양가_주상복합,기준_분양가_아파트)'),
    ("인지세", "인지세 단가(천원/세대)", '=IF(입력_사업구분="도급사업",기준_인지세_도급,기준_인지세_자체)'),
    ("특공세대", "특별공급 세대(갑지 총세대수로 구간 판단: 200↑ / 50~200 / 50↓)", "=계산_총세대"),
    ("특공단가", "특별공급대행 금액", "=IF(계산_특공세대>=200,기준_특공_200,IF(계산_특공세대>=50,기준_특공_50,기준_특공_50미만))"),
    ("청소1", "청소 인원 - 오픈 1개월차", "=IF(입력_유니트>=4,기준_청소_1개월_4이상,기준_청소_1개월_3이하)"),
    ("청소2", "청소 인원 - 오픈 2개월차", "=IF(입력_유니트>=4,기준_청소_2개월_4이상,기준_청소_2개월_3이하)"),
    ("오픈선물인원", "오픈 선물 인원", '=IF(입력_오픈선물인원="",기준_오픈선물_인원,입력_오픈선물인원)'),
    ("청약선물인원", "청약 선물 인원", '=IF(입력_청약선물인원="",기준_청약선물_인원,입력_청약선물인원)'),
    ("당첨자세대", "당첨자 선물 세대", "=ROUND(계산_총세대*기준_당첨자_비율/100,0)"),
    ("MGM세대", "부동산MGM 세대", "=ROUND(계산_총세대*기준_MGM_비율/100,0)"),
    ("내용증명건수", "내용증명 건수", "=ROUND(계산_총세대*입력_예비비율,0)"),
    ("사무실시작", "분양사무실 시작월 번호(연중 운영이면 전년 12월 → 1월부터 운영, 이전 운반비 없음)",
     '=IF(입력_사무실연중="Y",계산_연도*12,' + ym("사무실시작") + ")"),
    ("사무실종료", "분양사무실 종료월 번호", "=" + ym("사무실종료", "END")),
    ("사전점검", "사전점검월 번호", "=" + ym("사전점검")),
    ("입주", "입주개시월 번호", "=" + ym("입주")),
    ("지정기간", "입주지정기간(개월)", '=IF(입력_지정기간="",2,입력_지정기간)'),
    ("사전점검단가", "사전점검행사 단가(천원/세대)",
     "=IF(계산_총세대>=1000,기준_사전점검_1000이상,IF(계산_총세대>=600,기준_사전점검_600_1000,기준_사전점검_600미만))"),
    ("입주상담사인원", "입주상담사 인원", "=IF(계산_총세대>=1000,기준_입주상담사_1000이상,기준_입주상담사_1000미만)"),
    ("입주상담사월", "입주상담사 월 금액(천원)", "=계산_입주상담사인원*기준_입주상담사_일당*기준_입주상담사_근무일"),
    ("매니저배수", "입주매니저 배수", '=IF(입력_입주율="열위",기준_입주매니저_열위,기준_입주매니저_양호)'),
    ("이전등기단가", "이전등기 단가", '=IF(입력_시행사2="Y",기준_이전등기_2개이상,기준_이전등기_단가)'),
    ("상가월", "상가 계약 반영월 번호", '=IF(OR(입력_상가월연="",입력_상가월월=""),계산_오픈+1,입력_상가월연*12+입력_상가월월)'),
    ("상가MGM월", "상가 MGM 반영월 번호", '=IF(OR(입력_상가MGM월연="",입력_상가MGM월월=""),계산_오픈+1,입력_상가MGM월연*12+입력_상가MGM월월)'),
    ("해지월", "계약해지 반영월 번호", "=" + ym("해지월")),
    ("상담사시작", "재분양 상담사 시작월 번호", "=" + ym("상담사시작")),
    ("잔여판촉월", "잔여판촉 갑티슈 반영월 번호", "=" + ym("잔여판촉")),
    ("현수막감축", "현수막 감축 시작월 번호", "=" + ym("현수막감축")),
    ("숙소분양", "숙소비(분양) 지급월 번호", "=계산_사전시작"),
    ("숙소입주", "숙소비(입주) 지급월 번호", "=계산_입주"),
    ("사무실복합기", "분양사무실 복합기(천원/월)", '=IF(입력_사무실복합기="",기준_사무실_복합기,입력_사무실복합기)'),
    ("사무실보안", "분양사무실 보안경비(천원/월)", '=IF(입력_사무실보안="",기준_사무실_보안,입력_사무실보안)'),
    ("시공", "시공법인(표시)", '=IF(입력_시공="",입력_시행,입력_시공)'),
    ("작성일", "작성일(표시)", '=IF(입력_작성일="",TODAY(),입력_작성일)'),
]

# 월별 행: (이름, 설명, J열 식 — {c}=같은 열, {p}=앞 열, {r:이름}=그 행 번호)
MONTHLY = [
    ("월", "월", "=COLUMN()-9"),
    ("번호", "월 번호(연×12+월)", "=계산_연도*12+{c}{r:월}"),
    ("k", "오픈 경과월(D=0)", "={c}{r:번호}-계산_오픈"),
    ("사전", "사전영업 기간", "=IF(AND({c}{r:번호}>=계산_사전시작,{c}{r:번호}<계산_오픈),1,0)"),
    ("오픈1", "오픈월(D)", "=IF({c}{r:k}=0,1,0)"),
    ("오픈2개월", "오픈 2개월(D, D+1)", "=IF(OR({c}{r:k}=0,{c}{r:k}=1),1,0)"),
    ("오픈후3", "오픈후 3개월(D+2~D+4)", "=IF(AND({c}{r:k}>=2,{c}{r:k}<=4),1,0)"),
    ("MH", "M/H 운영", "=IF(AND({c}{r:번호}>=계산_사전시작,{c}{r:번호}<=계산_MH종료),1,0)"),
    ("MH오픈후", "M/H 운영(오픈 이후)", "=IF(AND({c}{r:MH}=1,{c}{r:k}>=0),1,0)"),
    ("조직", "조직분양 기간", "=IF(AND(계산_조직여부=1,{c}{r:번호}>=계산_조직시작,{c}{r:번호}<=계산_조직종료),1,0)"),
    ("사무실", "분양사무실 운영", "=IF(AND({c}{r:번호}>=계산_사무실시작,{c}{r:번호}<=계산_사무실종료),1,0)"),
    ("사전점검", "사전점검월", "=IF({c}{r:번호}=계산_사전점검,1,0)"),
    ("mk", "입주 경과월(입주개시=0)", "={c}{r:번호}-계산_입주"),
    ("입주센터", "입주지원센터 운영(사전점검~지정기간+1개월)",
     "=IF(AND({c}{r:번호}>=MIN(계산_사전점검,계산_입주),{c}{r:mk}<=계산_지정기간),1,0)"),
    ("입주통신", "입주기간 통신(입주개시~종료 익월)", "=IF(AND({c}{r:mk}>=0,{c}{r:mk}<=계산_지정기간),1,0)"),
    ("동하절기", "동·하절기(6~8월, 12~2월)", "=IF(OR({c}{r:월}<=2,AND({c}{r:월}>=6,{c}{r:월}<=8),{c}{r:월}=12),1,0)"),
    ("난방", "난방 기간(11~3월)", "=IF(OR({c}{r:월}<=3,{c}{r:월}>=11),1,0)"),
    ("활동", "현장 활동(M/H·조직분양·입주지원센터)", "=IF(OR({c}{r:MH}=1,{c}{r:조직}=1,{c}{r:입주센터}=1),1,0)"),
    # 같은 세목을 한 줄로 합친 항목(식대·여비·전기료)의 구간 깃발 — 월별 금액은 깃발×단가의 합, 산출근거의 개월 수는 V열(합)
    ("식대잔여", "식대 - 조직분양 등 잔여판촉(오픈 2개월 제외)", "=IF(AND({c}{r:조직}=1,{c}{r:오픈2개월}=0),1,0)"),
    ("식대기타", "식대 - 기타(사전영업·오픈 3개월~·입주지원센터, 조직분양 제외)",
     "=IF(AND({c}{r:오픈2개월}=0,{c}{r:조직}=0,OR({c}{r:사전}=1,{c}{r:MH오픈후}=1,{c}{r:입주센터}=1)),1,0)"),
    ("여비사전", "여비 - 사전영업 1개월·오픈후 3개월(D+2~D+4)", "=IF(OR({c}{r:사전}=1,{c}{r:오픈후3}=1),1,0)"),
    ("여비기타", "여비 - 기타(그 밖의 M/H·조직분양·입주 활동 기간)",
     "=IF(AND({c}{r:활동}=1,{c}{r:사전}=0,{c}{r:오픈2개월}=0,{c}{r:오픈후3}=0),1,0)"),
    ("전기동하", "전기료 - 동·하절기(M/H 운영 중 6~8월, 12~2월)", "=IF(AND({c}{r:MH}=1,{c}{r:동하절기}=1),1,0)"),
    ("전기기타", "전기료 - 기타(M/H 운영 중 3~5월, 9~11월)", "=IF(AND({c}{r:MH}=1,{c}{r:동하절기}=0),1,0)"),
    ("발급", "발급·송금수수료 기간", "=IF(OR({c}{r:MH오픈후}=1,{c}{r:입주센터}=1),1,0)"),
    ("보험", "견본주택 보험료 반영",
     '=IF(AND({c}{r:MH}=1,IF(입력_보험월="",MOD({c}{r:번호}-계산_사전시작,12)=0,{c}{r:월}=입력_보험월)),1,0)'),
    ("분양", "분양세대(분양율 시트)", "='분양율,입주율'!{p}8"),
    ("분양누계", "분양누계", "='분양율,입주율'!{p}9"),
    ("조직세대", "조직분양 세대",
     "=IF({c}{r:조직}=1,MIN({c}{r:분양},MAX(0,계산_조직상한-SUM($I{r:조직세대}:{p}{r:조직세대}))),0)"),
    ("대행세대", "분양대행 세대", '=IF(입력_분양대행="N",0,{c}{r:분양}-{c}{r:조직세대})'),
    ("MGM", "부동산MGM 세대(첫 계약월, C~D급)", "=IF(AND(계산_MGM적용=1,{c}{r:k}=1),계산_MGM세대,0)"),
    ("입주", "입주세대(분양율 시트)", "='분양율,입주율'!{p}13"),
    ("미입주", "공가 미입주 세대", "=IF({c}{r:mk}>=0,'분양율,입주율'!{p}15,0)"),
    ("단계", "단계 표시",
     '=IF({c}{r:mk}=0,"입주",IF({c}{r:사전점검}=1,"사전점검",IF(AND({c}{r:mk}>0,{c}{r:mk}<=계산_지정기간),"입주+"&{c}{r:mk},'
     'IF({c}{r:k}=-1,"사전",IF({c}{r:k}=0,"오픈",IF(AND({c}{r:k}>=1,{c}{r:k}<=3),"D+"&{c}{r:k},'
     'IF({c}{r:조직}=1,"조직분양",IF({c}{r:MH}=1,"M/H",IF({c}{r:사무실}=1,"사무실","")))))))))'),
]
SALES_CURVE = [0, 0.3, 0.5, 0.6, 0.7, 0.8, 0.85, 0.9, 0.95, 1, 1, 1, 1]
MOVEIN_CURVE = [0, 0.3, 0.6, 0.8, 0.9, 0.95, 1, 1, 1, 1, 1, 1, 1]


def build_calc(ws, names, I):
    put(ws, "B1", "계산 — 월별 단계 (자동, 고치지 마세요)", font(14, True))
    put(ws, "B2", "현장입력·기준단가로 각 달이 어떤 단계(사전·오픈·조직분양·입주…)인지 1/0 으로 표시합니다. 27년 예산 시트의 모든 식이 이 표를 씁니다.", F9)
    r = 4
    put(ws, f"B{r}", "항 목", F10B, FILL_HEAD, CENTER)
    put(ws, f"D{r}", "값", F10B, FILL_HEAD, CENTER)
    r += 1
    for key, label, formula in SCALARS:
        put(ws, f"B{r}", label, F10)
        put(ws, f"D{r}", formula, F10, al=RIGHT)
        names.add("계산_" + key, "계산", f"$D${r}")
        r += 1
    box(ws, f"B4:D{r - 1}")
    r += 1
    put(ws, f"B{r}", "기본 곡선(과거 작성사례)", F10B)
    put(ws, f"I{r}", "경과월 →", F9, al=RIGHT)
    for i in range(13):
        put(ws, f"{CL(10 + i)}{r}", i, F10B, FILL_HEAD, CENTER)
    r += 1
    for label, curve, nm in (("분양 누계곡선(오픈 경과월별 누계 분양률)", SALES_CURVE, "계산_분양곡선"),
                             ("입주 누계곡선(입주개시 경과월별 누계 입주율, 입주 첫달 0%)", MOVEIN_CURVE, "계산_입주곡선")):
        put(ws, f"B{r}", label, F10)
        for i, v in enumerate(curve):
            put(ws, f"{CL(10 + i)}{r}", v, F10, FILL_IN, CENTER, "0%")
        names.add(nm, "계산", f"$J${r}:$V${r}")
        r += 1
    r += 1
    head = r
    put(ws, f"B{r}", "월별 단계", F10B, FILL_HEAD, CENTER)
    for i in range(12):
        put(ws, f"{MC[i]}{r}", f"{i + 1}월", F10B, FILL_HEAD, CENTER)
    put(ws, f"V{r}", "개월/합", F10B, FILL_HEAD, CENTER)
    r += 1
    rows = {}
    for key, label, _ in MONTHLY:
        rows[key] = r
        r += 1
    for key, label, formula in MONTHLY:
        rr = rows[key]
        put(ws, f"B{rr}", label, F10)
        put(ws, f"C{rr}", key, F9G)
        for i in range(12):
            c, p = MC[i], CL(9 + i)
            f = formula.replace("{c}", c).replace("{p}", p)
            while "{r:" in f:
                a = f.index("{r:")
                b = f.index("}", a)
                f = f[:a] + str(rows[f[a + 3:b]]) + f[b + 1:]
            put(ws, f"{c}{rr}", f, F10, al=CENTER)
        if key not in ("월", "번호", "k", "mk", "단계", "분양누계"):
            put(ws, f"V{rr}", f"=SUM(J{rr}:U{rr})", F10B, al=CENTER)
    box(ws, f"B{head}:V{r - 1}")
    setw(ws, dict(A=2, B=44, C=10, D=12, **{c: 8 for c in MC}, I=8, V=9))
    ws.freeze_panes = f"J{head + 1}"
    return dict(row=rows)


def T(K, key, c):
    return f"'계산'!{c}${K['row'][key]}"


# ═════════════════════════════════════════════════════════════════
# 분양율,입주율 (공식 양식, 월 = I..T)
def build_rates_sheet(ws, K):
    put(ws, "A1", "분양률 및 입주율 세부내용", font(16, True), al=CENTER)
    ws.merge_cells("A1:U1")
    put(ws, "A4", "세대수", F10B, al=CENTER)
    ws.merge_cells("A4:B4")
    put(ws, "C4", "=계산_총세대", F10B, nf="#,##0")
    put(ws, "E4", "※ 초록 칸 = 자동 기본값(오픈월·입주개시월 기준 과거 작성사례 곡선). 실제 계획으로 덮어쓰세요.  노란 칸 = 입력", F9)
    put(ws, "A6", "구      분", F10B, al=CENTER)
    ws.merge_cells("A6:D7")
    put(ws, "E6", '=(계산_연도-1)&"년"', F10B, al=CENTER)
    ws.merge_cells("E6:H6")
    put(ws, "I6", '=계산_연도&"년"', F10B, al=CENTER)
    ws.merge_cells("I6:T6")
    put(ws, "U6", "비 고", F10B, al=CENTER)
    ws.merge_cells("U6:U7")
    for c, t in zip("EFGH", ("9월누계", "10월", "11월", "12월")):
        put(ws, f"{c}7", t, F10B, al=CENTER)
    for i in range(12):
        col = CL(9 + i)
        lab = T(K, "단계", MC[i])
        put(ws, f"{col}7", f'="{i + 1}월"&IF({lab}="","","("&{lab}&")")', F9, al=CENTER)
    put(ws, "A8", "분 양", F10B, al=CENTER)
    ws.merge_cells("A8:A12")
    put(ws, "A13", "입 주", F10B, al=CENTER)
    ws.merge_cells("A13:A17")
    labels = {8: "분양세대", 9: "분양누계", 10: "미분양누계", 11: "분양율", 12: "미분양율",
              13: "입주세대", 14: "입주누계", 15: "미입주누계", 16: "입주율", 17: "미입주율"}
    for rr, t in labels.items():
        put(ws, f"C{rr}", t, F10, al=CENTER)
        ws.merge_cells(f"C{rr}:D{rr}")
    put(ws, "E8", "=입력_전년분양", F10, FILL_DEF, nf="#,##0")
    put(ws, "E13", "=입력_전년입주", F10, FILL_DEF, nf="#,##0")
    for c in "FGH":
        put(ws, f"{c}8", None, F10, FILL_IN, nf="#,##0")
        put(ws, f"{c}13", None, F10, FILL_IN, nf="#,##0")
    k, mk = K["row"]["k"], K["row"]["mk"]
    for i in range(12):
        col, prev, cc = CL(9 + i), CL(8 + i), MC[i]
        kk, mm = f"'계산'!{cc}${k}", f"'계산'!{cc}${mk}"
        put(ws, f"{col}8",
            f"=IF({kk}<1,0,MAX(0,MIN(ROUND(계산_총세대*MIN(INDEX(계산_분양곡선,MIN({kk},12)+1),계산_목표율),0)"
            f"-ROUND(계산_총세대*MIN(INDEX(계산_분양곡선,MIN({kk}-1,12)+1),계산_목표율),0),계산_목표세대-{prev}9)))",
            F10, FILL_DEF, nf="#,##0")
        put(ws, f"{col}13",
            f"=IF({mm}<1,0,MAX(0,MIN(ROUND(계산_총세대*INDEX(계산_입주곡선,MIN({mm},12)+1),0)"
            f"-ROUND(계산_총세대*INDEX(계산_입주곡선,MIN({mm}-1,12)+1),0),계산_총세대-{prev}14)))",
            F10, FILL_DEF, nf="#,##0")
    for j, col in enumerate("EFGHIJKLMNOPQRST"):
        prev = "EFGHIJKLMNOPQRST"[j - 1] if j else None
        put(ws, f"{col}9", f"={col}8" if j == 0 else f"={prev}9+{col}8", F10, nf="#,##0")
        put(ws, f"{col}10", f'=IF($C$4=0,"",$C$4-{col}9)', F10, nf="#,##0")
        put(ws, f"{col}11", f'=IF($C$4=0,"",{col}9/$C$4)', F10, nf="0.0%")
        put(ws, f"{col}12", f'=IF($C$4=0,"",1-{col}11)', F10, nf="0.0%")
        put(ws, f"{col}14", f"={col}13" if j == 0 else f"={prev}14+{col}13", F10, nf="#,##0")
        put(ws, f"{col}15", f'=IF($C$4=0,"",$C$4-{col}14)', F10, nf="#,##0")
        put(ws, f"{col}16", f'=IF($C$4=0,"",{col}14/$C$4)', F10, nf="0.0%")
        put(ws, f"{col}17", f'=IF($C$4=0,"",1-{col}16)', F10, nf="0.0%")
    put(ws, "A18", "※ 색칠한 부분만 작성", F9)
    box(ws, "A6:U17", med, hair)
    for rr in (8, 13):
        for col in "ABCDEFGHIJKLMNOPQRSTU":
            b = ws[f"{col}{rr}"].border
            ws[f"{col}{rr}"].border = Border(left=b.left, right=b.right, top=thin, bottom=b.bottom)
    setw(ws, dict(A=6, B=2, C=8, D=4, E=9, F=7, G=7, H=7, **{CL(9 + i): 9.5 for i in range(12)}, U=10))
    ws.row_dimensions[7].height = 28


# ═════════════════════════════════════════════════════════════════
# 분양대행 및 촉진수수료 (월 = J..U)
def _types_block(ws, r, I, title_rows, units_formula, price_formula, merge_title):
    """타입별 세대수 10줄 + 소계, 금액 10줄 + 소계. units_formula(c, cum, cum_prev) → 식"""
    t1, t2 = I["types"]
    u0 = r
    put(ws, f"B{r}", "세대수", F10, al=CENTER)
    ws.merge_cells(f"B{r}:B{r + NTYPES - 1}")
    for k in range(NTYPES):
        rr = r + k
        tr = t1 + k
        put(ws, f"C{rr}", f"=IF('현장입력'!C{tr}=\"\",\"\",'현장입력'!C{tr})", F10, al=CENTER)
        put(ws, f"D{rr}", f'=IF($C{rr}="","",{price_formula[1:]})', F10, nf="#,##0")
        put(ws, f"W{rr}", f"=IF(계산_총세대=0,0,SUM('현장입력'!$D${t1}:$D${tr})/계산_총세대)", F9G, nf="0.0000")
        for i in range(12):
            c = MC[i]
            prev = "0" if k == 0 else f"ROUND({{N}}*$W{rr - 1},0)"
            put(ws, f"{c}{rr}", units_formula(c, f"$W{rr}", prev), F10, nf="#,##0;-#,##0;\"\"")
        put(ws, f"E{rr}", f"=SUM(F{rr}:H{rr})", F10, nf="#,##0;-#,##0;\"\"")
        put(ws, f"I{rr}", f"=SUM(J{rr}:U{rr})", F10, nf="#,##0;-#,##0;\"\"")
    rs = r + NTYPES
    put(ws, f"C{rs}", "소계", F10B, FILL_SUB, CENTER)
    for col in ["E", "F", "G", "H", "I"] + MC:
        put(ws, f"{col}{rs}", f"=SUM({col}{r}:{col}{rs - 1})", F10B, FILL_SUB, nf="#,##0;-#,##0;\"\"")
    r = rs + 1
    a0 = r
    put(ws, f"B{r}", "금 액", F10, al=CENTER)
    ws.merge_cells(f"B{r}:B{r + NTYPES - 1}")
    for k in range(NTYPES):
        rr = r + k
        put(ws, f"C{rr}", f"=C{u0 + k}", F10, al=CENTER)
        put(ws, f"D{rr}", f"=D{u0 + k}", F10, nf="#,##0")
        for col in ["F", "G", "H"] + MC:
            put(ws, f"{col}{rr}", f"={col}{u0 + k}*N($D{rr})", F10, nf="#,##0;-#,##0;\"\"")
        put(ws, f"E{rr}", f"=SUM(F{rr}:H{rr})", F10, nf="#,##0;-#,##0;\"\"")
        put(ws, f"I{rr}", f"=SUM(J{rr}:U{rr})", F10, nf="#,##0;-#,##0;\"\"")
    ra = r + NTYPES
    put(ws, f"C{ra}", "소계", F10B, FILL_SUB, CENTER)
    for col in ["E", "F", "G", "H", "I"] + MC:
        put(ws, f"{col}{ra}", f"=SUM({col}{r}:{col}{ra - 1})", F10B, FILL_SUB, nf="#,##0;-#,##0;\"\"")
    put(ws, f"A{u0}", merge_title, F10B, al=CENTER)
    ws.merge_cells(f"A{u0}:A{ra}")
    box(ws, f"A{u0}:V{ra}", med, hair)
    return ra + 1, ra, rs


def build_agency(ws, K, I):
    put(ws, "A1", "분양대행수수료 및 분양촉진수수료 산정근거", font(16, True), al=CENTER)
    ws.merge_cells("A1:V1")
    put(ws, "A4", "[자동 계산 — 세대수는 분양율 시트·현장입력 기준]", F9)
    put(ws, "V4", "(금액단위: 천원)", F9, al=RIGHT)
    put(ws, "A5", "구      분", F10B, al=CENTER)
    ws.merge_cells("A5:C6")
    put(ws, "D5", "단가", F10B, al=CENTER)
    ws.merge_cells("D5:D6")
    put(ws, "E5", '=(계산_연도-1)&"년"', F10B, al=CENTER)
    ws.merge_cells("E5:H5")
    put(ws, "I5", '=계산_연도&"년"', F10B, al=CENTER)
    ws.merge_cells("I5:U5")
    put(ws, "V5", "비 고", F10B, al=CENTER)
    ws.merge_cells("V5:V6")
    put(ws, "W5", "누적비율", F9G, al=CENTER)
    for c, t in zip("EFGHI", ("계", "10월", "11월", "12월", "계")):
        put(ws, f"{c}6", t, F10B, al=CENTER)
    for i in range(12):
        put(ws, f"{MC[i]}6", f"{i + 1}월", F10B, al=CENTER)
    box(ws, "A5:V6", med, thin)

    def units(key):
        def f(c, cum, prev):
            n = T(K, key, c)
            return f"=ROUND({n}*{cum},0)-" + prev.replace("{N}", n)
        return f

    r = 7
    r, agency_amt, _ = _types_block(ws, r, I, None, units("대행세대"), "=기준_분양대행", "분양대행\n수수료")
    r, org_amt, _ = _types_block(ws, r, I, None, units("조직세대"), "=계산_조직단가", "분양대행\n수수료\n(조직)")
    put(ws, f"A{r}", "분양대행수수료 총계", F10B, FILL_TOT, CENTER)
    ws.merge_cells(f"A{r}:D{r}")
    for col in ["E", "F", "G", "H", "I"] + MC:
        put(ws, f"{col}{r}", f"={col}{agency_amt}+{col}{org_amt}", F10B, FILL_TOT, nf="#,##0;-#,##0;\"\"")
    box(ws, f"A{r}:V{r}", med, hair)
    r += 1
    r, mgm_amt, _ = _types_block(ws, r, I, None, units("MGM"), "=기준_MGM", "분양촉진\n수수료\n(MGM)")
    put(ws, f"A{r + 1}", "※ 분양대행: 조직분양 외 계약세대 × 2,000천원/세대 · 조직: 조직분양 시작월(기본 D+2M) 이후 계약세대(비율 상한) × 급지 단가 · "
                         "MGM: 첫 계약월(D+1) × 총세대 30% × 1,000천원(C~D급지)", F9)
    setw(ws, dict(A=9, B=6, C=7, D=8, E=8, F=7, G=7, H=7, I=9, **{c: 9 for c in MC}, V=8, W=8))
    ws.freeze_panes = "E7"
    ws.print_area = f"A1:V{r + 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return dict(agency=agency_amt, org=org_amt, mgm=mgm_amt)


# ═════════════════════════════════════════════════════════════════
# 공가세대관리비 (월 = I..T)
def build_vacancy(ws, K, I):
    put(ws, "A1", "공가세대관리비 산정근거", font(16, True), al=CENTER)
    ws.merge_cells("A1:U1")
    put(ws, "A4", "[자동 계산 — 미입주세대는 분양율,입주율 시트, 공급평은 현장입력]", F9)
    put(ws, "U4", "(금액단위: 천원)", F9, al=RIGHT)
    put(ws, "A5", "구      분", F10B, al=CENTER)
    ws.merge_cells("A5:C6")
    put(ws, "D6", "계", F10B, al=CENTER)
    put(ws, "E5", '=(계산_연도-1)&"년"', F10B, al=CENTER)
    ws.merge_cells("E5:G5")
    put(ws, "H6", "계", F10B, al=CENTER)
    put(ws, "I5", '=계산_연도&"년"', F10B, al=CENTER)
    ws.merge_cells("I5:T5")
    put(ws, "U5", "비 고", F10B, al=CENTER)
    ws.merge_cells("U5:U6")
    put(ws, "V6", "공급평", F9G, al=CENTER)
    put(ws, "W6", "누적비율", F9G, al=CENTER)
    for c, t in zip("EFG", ("10월", "11월", "12월")):
        put(ws, f"{c}6", t, F10B, al=CENTER)
    for i in range(12):
        put(ws, f"{CL(9 + i)}6", f"{i + 1}월", F10B, al=CENTER)
    box(ws, "A5:U6", med, thin)
    put(ws, "A7", "평당 공가관리비", F10, al=CENTER)
    ws.merge_cells("A7:C7")
    for i in range(12):
        put(ws, f"{CL(9 + i)}7", "=기준_공가관리비", F10, nf="#,##0.0")
    put(ws, "H7", "=AVERAGE(I7:T7)", F10, nf="#,##0.0")
    t1, t2 = I["types"]
    r = 8
    u0 = r
    put(ws, f"A{r}", "공가세대\n관리비", F10B, al=CENTER)
    put(ws, f"B{r}", "세대수", F10, al=CENTER)
    ws.merge_cells(f"B{r}:B{r + NTYPES - 1}")
    for k in range(NTYPES):
        rr, tr = r + k, t1 + k
        put(ws, f"C{rr}", f"=IF('현장입력'!C{tr}=\"\",\"\",'현장입력'!C{tr})", F10, al=CENTER)
        put(ws, f"V{rr}", f"=IF('현장입력'!F{tr}=\"\",0,'현장입력'!F{tr})", F9G, nf="#,##0")
        put(ws, f"W{rr}", f"=IF(계산_총세대=0,0,SUM('현장입력'!$D${t1}:$D${tr})/계산_총세대)", F9G, nf="0.0000")
        for i in range(12):
            col, n = CL(9 + i), T(K, "미입주", MC[i])
            prev = "0" if k == 0 else f"ROUND({n}*$W{rr - 1},0)"
            put(ws, f"{col}{rr}", f"=ROUND({n}*$W{rr},0)-{prev}", F10, nf="#,##0;-#,##0;\"\"")
        put(ws, f"D{rr}", f"=SUM(E{rr}:G{rr})", F10, nf="#,##0;-#,##0;\"\"")
        put(ws, f"H{rr}", f"=SUM(I{rr}:T{rr})", F10, nf="#,##0;-#,##0;\"\"")
    rs = r + NTYPES
    put(ws, f"C{rs}", "소계", F10B, FILL_SUB, CENTER)
    for col in "DEFGHIJKLMNOPQRST":
        put(ws, f"{col}{rs}", f"=SUM({col}{r}:{col}{rs - 1})", F10B, FILL_SUB, nf="#,##0;-#,##0;\"\"")
    r = rs + 1
    put(ws, f"B{r}", "금 액", F10, al=CENTER)
    ws.merge_cells(f"B{r}:B{r + NTYPES - 1}")
    for k in range(NTYPES):
        rr = r + k
        put(ws, f"C{rr}", f"=C{u0 + k}", F10, al=CENTER)
        for col in "EFG":
            put(ws, f"{col}{rr}", f"=E$7*{col}{u0 + k}*$V{u0 + k}" if False else None, F10, nf="#,##0;-#,##0;\"\"")
        for i in range(12):
            col = CL(9 + i)
            put(ws, f"{col}{rr}", f"={col}$7*{col}{u0 + k}*$V{u0 + k}", F10, nf="#,##0;-#,##0;\"\"")
        put(ws, f"D{rr}", f"=SUM(E{rr}:G{rr})", F10, nf="#,##0;-#,##0;\"\"")
        put(ws, f"H{rr}", f"=SUM(I{rr}:T{rr})", F10, nf="#,##0;-#,##0;\"\"")
    ra = r + NTYPES
    put(ws, f"C{ra}", "계", F10B, FILL_SUB, CENTER)
    for col in "DEFGHIJKLMNOPQRST":
        put(ws, f"{col}{ra}", f"=SUM({col}{r}:{col}{ra - 1})", F10B, FILL_SUB, nf="#,##0;-#,##0;\"\"")
    ws.merge_cells(f"A{u0}:A{ra}")
    box(ws, f"A7:U{ra}", med, hair)
    put(ws, f"A{ra + 2}", "※ 평당 관리비(기준단가) × 미입주세대(타입별) × 공급평(공급면적×0.3025 올림). 입주개시월은 입주세대 0 → 100% 미입주로 반영",
        F9)
    setw(ws, dict(A=9, B=6, C=8, D=8, E=7, F=7, G=7, H=9, **{CL(9 + i): 9 for i in range(12)}, U=8, V=7, W=8))
    ws.print_area = f"A1:U{ra + 2}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return dict(amount=ra)


# ═════════════════════════════════════════════════════════════════
# 대납이자 (공식 양식, 월 = K..V)
def build_interest(ws):
    put(ws, "A1", "중도금 대납이자 산정근거", font(16, True), al=CENTER)
    ws.merge_cells("A1:W1")
    put(ws, "A4", "[음영 부분만 입력]  ※ 현장입력 '중도금 무이자'=Y 일 때만 27년 예산에 반영", F9)
    put(ws, "W4", "(금액단위: 백만원)", F9, al=RIGHT)
    put(ws, "A5", "구      분", F10B, al=CENTER)
    ws.merge_cells("A5:C6")
    put(ws, "F5", '=(계산_연도-1)&"년"', F10B, al=CENTER)
    ws.merge_cells("F5:I5")
    put(ws, "K5", '=계산_연도&"년"', F10B, al=CENTER)
    ws.merge_cells("K5:V5")
    put(ws, "W5", "비  고", F10B, al=CENTER)
    ws.merge_cells("W5:W6")
    for c, t in zip("DEFGHIJ", ("계", "기존", "9월", "10월", "11월", "12월", "계")):
        put(ws, f"{c}6", t, F10B, al=CENTER)
    for i in range(12):
        put(ws, f"{CL(11 + i)}6", f"{i + 1}월", F10B, al=CENTER)
    put(ws, "A7", "중도금\n대 출", F10B, al=CENTER)
    ws.merge_cells("A7:A18")
    put(ws, "B7", "1금융", F10, al=CENTER)
    ws.merge_cells("B7:B12")
    put(ws, "B13", "2금융", F10, al=CENTER)
    ws.merge_cells("B13:B18")
    for k in range(6):
        for base in (7, 13):
            rr = base + k
            put(ws, f"C{rr}", f"중{k + 1}회", F10, al=CENTER)
            put(ws, f"D{rr}", f"=SUM(E{rr}:I{rr})", F10, nf="#,##0;-#,##0;\"\"")
            for col in ["E", "F", "G", "H", "I"] + [CL(11 + i) for i in range(12)]:
                put(ws, f"{col}{rr}", None, F10, FILL_IN, nf="#,##0")
    put(ws, "B19", "계", F10B, al=CENTER)
    ws.merge_cells("B19:C19")
    put(ws, "D19", "=SUM(E19:I19)", F10B, nf="#,##0;-#,##0;\"\"")
    for col in ["E", "F", "G", "H", "I"] + [CL(11 + i) for i in range(12)]:
        put(ws, f"{col}19", f"=SUM({col}7:{col}18)", F10B, nf="#,##0;-#,##0;\"\"")
    put(ws, "J19", "=SUM(K19:V19)", F10B, nf="#,##0;-#,##0;\"\"")
    put(ws, "A20", "대 출\n이 자", F10B, al=CENTER)
    ws.merge_cells("A20:A31")
    put(ws, "B20", "1금융", F10, al=CENTER)
    ws.merge_cells("B20:B25")
    put(ws, "B26", "2금융", F10, al=CENTER)
    ws.merge_cells("B26:B31")
    for k in range(6):
        for base, src, rate_prev, rate_cur in ((20, 7, "$B$36", "$B$37"), (26, 13, "$C$36", "$C$37")):
            rr, s = base + k, src + k
            put(ws, f"C{rr}", f"중{k + 1}회", F10, al=CENTER)
            put(ws, f"D{rr}", f"=SUM(E{rr}:I{rr})", F10, nf="#,##0.0;-#,##0.0;\"\"")
            for col, pcol in (("G", "F"), ("H", "G"), ("I", "H")):
                put(ws, f"{col}{rr}", f"={pcol}{s}*{rate_prev}/12", F10, nf="#,##0.0;-#,##0.0;\"\"")
            put(ws, f"K{rr}", f"=I{s}*{rate_prev}/12", F10, nf="#,##0.0;-#,##0.0;\"\"")
            for i in range(1, 12):
                col, pcol = CL(11 + i), CL(10 + i)
                put(ws, f"{col}{rr}", f"={pcol}{s}*{rate_cur}/12", F10, nf="#,##0.0;-#,##0.0;\"\"")
            put(ws, f"J{rr}", f"=SUM(K{rr}:V{rr})", F10, nf="#,##0.0;-#,##0.0;\"\"")
    put(ws, "B32", "계", F10B, al=CENTER)
    ws.merge_cells("B32:C32")
    for col in ["D", "E", "F", "G", "H", "I", "J"] + [CL(11 + i) for i in range(12)]:
        put(ws, f"{col}32", f"=SUM({col}20:{col}31)", F10B, FILL_SUB, nf="#,##0.0;-#,##0.0;\"\"")
    box(ws, "A5:W32", med, hair)
    put(ws, "A34", "[중도금 이자율]", F10B)
    put(ws, "G34", "[약정조건]", F10B)
    for c, t in zip("ABCD", ("구 분", "1금융", "2금융", "비   고")):
        put(ws, f"{c}35", t, F10B, FILL_HEAD, CENTER)
    ws.merge_cells("D35:E35")
    put(ws, "A36", "=계산_연도-1", F10, al=CENTER)
    put(ws, "A37", "=계산_연도", F10, al=CENTER)
    put(ws, "B37", '=IF(입력_최고이율<>"",입력_최고이율+기준_대납_기분양_가산/100,기준_대납_신규_1금융/100)', F10, FILL_DEF, CENTER, "0.00%")
    put(ws, "C37", '=IF(입력_최고이율<>"",입력_최고이율+기준_대납_기분양_가산/100,기준_대납_신규_2금융/100)', F10, FILL_DEF, CENTER, "0.00%")
    put(ws, "B36", "=B37", F10, FILL_DEF, CENTER, "0.00%")
    put(ws, "C36", "=C37", F10, FILL_DEF, CENTER, "0.00%")
    put(ws, "D37", '=IF(입력_최고이율<>"","기분양: 최근 결재 최고이율 + 가산","신규 분양현장 기준이율")', F9)
    ws.merge_cells("D36:E36")
    ws.merge_cells("D37:E37")
    for c, t in zip("GH", ("구 분", "세부내용")):
        put(ws, f"{c}35", t, F10B, FILL_HEAD, CENTER)
    put(ws, "G36", "1금융", F10, al=CENTER)
    put(ws, "G37", "2금융", F10, al=CENTER)
    for rr in (36, 37):
        put(ws, f"H{rr}", "CD + 몇%", F10, FILL_IN)
        ws.merge_cells(f"H{rr}:I{rr}")
    ws.merge_cells("H35:I35")
    box(ws, "A35:E37")
    box(ws, "G35:I37")
    put(ws, "A38", "※ 이자율 자금팀 협조 要  (초록 칸 = 기준단가 시트의 기준 이율, 덮어쓰기 가능)", F9)
    setw(ws, dict(A=7, B=6, C=7, D=8, E=7, **{CL(6 + i): 7.5 for i in range(17)}, W=8))


# ═════════════════════════════════════════════════════════════════
# 27년 예산
def build_budget(ws, K, pay, vac):
    N = lambda r: f'IF(COUNTIF(J{r}:U{r},"<>0")>0," × "&COUNTIF(J{r}:U{r},"<>0")&"개월","")'
    M1 = lambda r: f'IFERROR(" ("&MATCH(TRUE,INDEX(J{r}:U{r}<>0,0),0)&"월)","")'
    t = lambda name: f'IF({name}=INT({name}),TEXT({name},"#,##0"),TEXT({name},"#,##0.0##"))'
    cnt = lambda key: f"'계산'!$V${K['row'][key]}"      # 그 단계 깃발이 켜진 달 수('계산' 시트 '개월/합' 열)
    PFX = "     - "                                      # 합친 줄의 둘째 줄부터 들여쓰기(" 식대 - " 너비만큼)

    def merged(head, parts):
        """같은 세목의 구간별 줄을 한 줄로: head 뒤에 '구간 : 단가천원/월 × n개월' 을 줄바꿈으로 이어 붙인다(0개월 구간은 생략).
        parts = [(깃발 키, 구간 이름, 단가 식)]. 모든 구간이 0개월이면 구간별 단가만 적는다(그 줄은 금액 0).
        월별 금액 식은 '(깃발×단가 + …)*1000' 으로 따로 적는다(산식은 그 달 구간의 단가만 남음)."""
        segs = "&".join(f'IF({cnt(k)}>0,CHAR(10)&"{PFX}{label} : "&{t(rate)}&"천원/월 × "&{cnt(k)}&"개월","")'
                        for k, label, rate in parts)
        none = '&" · "&'.join(f'"{label} "&{t(rate)}' for _, label, rate in parts) + '&"천원/월 (해당 월 없음)"'
        total = "+".join(cnt(k) for k, _, _ in parts)
        fn = lambda r: f'={head}&IF({total}=0,{none},MID({segs},{len(PFX) + 2},2000))'
        fn.lines = len(parts)                            # 양식 줄 높이(구간 수만큼)
        return fn
    # 그룹: (항목, 세목 이름(기준단가 셀 참조용 키), [(산출근거 식, 월 식, 비고)])
    # 월 식: {c} 열, {T:키} 계산 행, {B:...} 같은 열 다른 시트
    G = []
    G.append(("복리후생비", "식대_오픈", [
        (merged('" 식대 - "', [("오픈2개월", "오픈2개월", "기준_식대_오픈"),
                              ("식대잔여", "조직분양 등 잔여판촉시", "기준_식대_잔여"),
                              ("식대기타", "기타(사전영업·오픈 3개월~·사전점검~입주지정기간+1개월)", "기준_식대_기타")]),
         "=({T:오픈2개월}*기준_식대_오픈+{T:식대잔여}*기준_식대_잔여+{T:식대기타}*기준_식대_기타)*1000", ""),
        (lambda r: f'=" 피복비 : "&{t("기준_피복비")}&"천원/명 × "&입력_여직원&"명(1월 반영)"',
         "=IF(AND({T:월}=1,SUM('계산'!$J${R:활동}:$U${R:활동})>0),입력_여직원*기준_피복비,0)*1000", ""),
        (lambda r: f'=" 숙소비 : "&{t("기준_숙소비_월")}&"천원/월 × "&기준_숙소_개월&"개월(일시) × "&입력_숙소분양&"호실"&{M1(r)}',
         "=IF({T:번호}=계산_숙소분양,입력_숙소분양*기준_숙소비_월*기준_숙소_개월,0)*1000", ""),
        (lambda r: f'=" 입주 여직원(지원인력) 숙소비 : "&{t("기준_숙소비_입주_월")}&"천원/월 × "&기준_숙소_개월&"개월(일시) × "&입력_숙소입주&"호실"&{M1(r)}',
         "=IF({T:번호}=계산_숙소입주,입력_숙소입주*기준_숙소비_입주_월*기준_숙소_개월,0)*1000", ""),
        (lambda r: f'=" 숙소관리비 : "&{t("기준_숙소관리비")}&"천원/월 × 호실 (숙소비 지급 다음 달부터 "&기준_숙소_개월&"개월)"',
         "=(IF(AND({T:번호}>계산_숙소분양,{T:번호}<=계산_숙소분양+기준_숙소_개월),입력_숙소분양,0)"
         "+IF(AND({T:번호}>계산_숙소입주,{T:번호}<=계산_숙소입주+기준_숙소_개월),입력_숙소입주,0))*기준_숙소관리비*1000", ""),
    ]))
    G.append(("여비교통비", "여비_오픈_수도권", [
        (merged('" 유류비, 숙박비("&입력_지역&") - "', [("오픈2개월", "오픈2개월", "계산_여비오픈"),
                                                  ("여비사전", "사전영업1개월 및 오픈후 3개월", "계산_여비사전"),
                                                  ("여비기타", "기타", "계산_여비기타")]),
         "=({T:오픈2개월}*계산_여비오픈+{T:여비사전}*계산_여비사전+{T:여비기타}*계산_여비기타)*1000", ""),
    ]))
    G.append(("통신비", "통신_MH", [
        (merged('" 통신요금 - "', [("MH", "M/H 운영기간", "기준_통신_MH"),
                                ("입주통신", "입주기간(입주 종료일 익월까지)", "기준_통신_입주")]),
         "=({T:MH}*기준_통신_MH+{T:입주통신}*기준_통신_입주)*1000", ""),
        (lambda r: f'=" 통신요금 및 우편료(우편 기본) : "&{t("기준_우편_기본")}&"천원/월"&{N(r)}',
         "=IF(OR({T:MH}=1,{T:입주센터}=1),기준_우편_기본,0)*1000", ""),
        (lambda r: f'=" 등기(사업승인 변경·제품변경 안내) : "&{t("기준_등기")}&"천원/건 × 계약세대 (발송월 "&입력_등기월1&IF(입력_등기월2="","","·"&입력_등기월2)&"월)"',
         "=IF(AND(OR({T:월}=입력_등기월1,{T:월}=입력_등기월2),{T:번호}<계산_입주),기준_등기*{T:분양누계},0)*1000", ""),
        (lambda r: f'=" 국공립어린이집 동의서 : "&{t("기준_어린이집")}&"천원/건 × "&TEXT(계산_총세대,"#,##0")&"세대(입주 전월)"',
         '=IF(AND(입력_어린이집="Y",{T:번호}=계산_입주-1),기준_어린이집*계산_총세대,0)*1000', ""),
        (lambda r: f'=" 입주 안내 등기(사전점검·입주 안내) : "&{t("기준_등기")}&"천원/건 × "&TEXT(계산_총세대,"#,##0")&"세대 × "&입력_입주안내&"회"',
         "=(IF(AND(입력_입주안내>=1,{T:번호}=계산_사전점검-1),1,0)+IF(AND(입력_입주안내>=2,{T:번호}=계산_입주-1),1,0))"
         "*기준_등기*계산_총세대*1000", ""),
        (lambda r: f'=" 내용증명(예비입주자 포함) : "&{t("기준_내용증명")}&"천원/건 × "&TEXT(계산_내용증명건수,"#,##0")&"건(세대수 × "&TEXT(입력_예비비율,"0%")&")"',
         "=IF({T:오픈1}=1,기준_내용증명*계산_내용증명건수,0)*1000", ""),
        (lambda r: f'=" [분양사무실] 통신요금 및 우편료 : "&{t("기준_사무실_통신")}&"천원/월(인터넷 포함)"&{N(r)}',
         "={T:사무실}*기준_사무실_통신*1000", ""),
        (lambda r: f'=" [분양사무실] 우편료 기본 : "&{t("기준_사무실_우편")}&"천원/월"&{N(r)}',
         "={T:사무실}*기준_사무실_우편*1000", ""),
    ]))
    G.append(("수도광열비", "전기_동하절기", [
        (merged('" 전기료 - "', [("전기동하", "동·하절기(6~8월, 12~2월)", "기준_전기_동하절기"),
                               ("전기기타", "기타(3~5월, 9~11월)", "기준_전기_기타")]),
         "=({T:전기동하}*기준_전기_동하절기+{T:전기기타}*기준_전기_기타)*1000", ""),
        (lambda r: f'=" 상하수도 - M/H 운영기간 : "&{t("기준_상하수도")}&"천원/월"&{N(r)}',
         "={T:MH}*기준_상하수도*1000", ""),
        (lambda r: f'=" 난방비 - 간절기·동절기(11~3월) : "&{t("기준_난방비")}&"천원/월"&{N(r)}',
         "=IF(AND({T:MH}=1,{T:난방}=1),기준_난방비,0)*1000", ""),
        (lambda r: f'=" 전기 미인입 현장(오픈 1개월 限) : "&{t("기준_전기미인입")}&"천원"&IF(입력_전기미인입="Y",""," - 해당 없음")',
         '=IF(AND(입력_전기미인입="Y",{T:오픈1}=1),기준_전기미인입,0)*1000', ""),
    ]))
    G.append(("지급임차료", "복합기", [
        (lambda r: f'=" 복합기 : "&{t("기준_복합기")}&"천원/월 × "&입력_복합기&"대"&{N(r)}',
         "=IF(OR({T:MH}=1,{T:입주센터}=1),기준_복합기*입력_복합기,0)*1000", ""),
        (lambda r: f'=" 계약용 프린터 : "&{t("기준_프린터")}&"천원/대 × "&기준_프린터_대수&"대, 파쇄기 : "&{t("기준_파쇄기")}&"천원"&{N(r)}',
         "=IF({T:MH오픈후}=1,기준_프린터*기준_프린터_대수+기준_파쇄기,0)*1000", ""),
        (lambda r: f'=" 사무실PC : "&{t("기준_PC")}&"천원/월 × "&입력_PC&"대(M/H), "&입력_입주PC&"대(입주지원센터)"',
         "=MAX({T:MH}*입력_PC,{T:입주센터}*입력_입주PC)*기준_PC*1000", ""),
        (lambda r: f'=" M/H 주차장 임차료(오픈 1개월) : "&입력_지역&" "&{t("계산_주차장")}&"천원"&IF(입력_주차장="Y",""," - 해당 없음")',
         '=IF(AND(입력_주차장="Y",{T:오픈1}=1),계산_주차장,0)*1000', ""),
        (lambda r: f'=" [분양사무실] 복합기 : "&{t("계산_사무실복합기")}&"천원/월"&{N(r)}',
         "={T:사무실}*계산_사무실복합기*1000", ""),
    ]))
    G.append(("보험료", "보험료", [
        (lambda r: f'=" 견본주택 보험료(신축 기준) : "&{t("기준_보험료")}&"천원/년"&{M1(r)}',
         "=IF({T:보험}=1,기준_보험료,0)*1000", ""),
    ]))
    G.append(("세금과공과", "인지세_자체", [
        (lambda r: f'=" 인지세("&입력_사업구분&") : "&{t("계산_인지세")}&"천원/세대 × 계약세대"&IF(입력_상가>0,", 상가 "&입력_상가&"호실","")',
         "=({T:분양}+IF({T:번호}=계산_상가월,입력_상가,0))*계산_인지세*1000", ""),
    ]))
    G.append(("광고선전비", "갑티슈", [
        (lambda r: f'=" 갑티슈 : "&{t("기준_갑티슈")}&"천원/개 (사전 "&TEXT(입력_갑티슈사전,"#,##0")&"개, 오픈 "&TEXT(입력_갑티슈오픈,"#,##0")&"개"&IF(입력_잔여갑티슈>0,", 잔여판촉 "&TEXT(입력_잔여갑티슈,"#,##0")&"개","")&")"',
         "=(IF({T:k}=-1,입력_갑티슈사전,0)+IF({T:k}=0,입력_갑티슈오픈,0)+IF({T:번호}=계산_잔여판촉월,입력_잔여갑티슈,0))*기준_갑티슈*1000", ""),
        (lambda r: f'=" 물티슈 : "&{t("기준_물티슈")}&"원/개 (사전 "&TEXT(입력_물티슈,"#,##0")&"개)"',
         "=IF({T:k}=-1,입력_물티슈*기준_물티슈,0)", ""),
        (lambda r: f'=" 종이컵 : "&{t("기준_종이컵")}&"천원/박스 (사전 "&TEXT(입력_종이컵,"#,##0")&"박스)"',
         "=IF({T:k}=-1,입력_종이컵*기준_종이컵,0)*1000", ""),
        (lambda r: f'=" 크린홈세트(3종) : "&{t("기준_크린홈")}&"천원/개 (사전 "&TEXT(입력_크린홈,"#,##0")&"개)"',
         "=IF({T:k}=-1,입력_크린홈*기준_크린홈,0)*1000", ""),
        (lambda r: f'=" 마스크 : "&{t("기준_마스크")}&"천원/개 (사전 "&TEXT(입력_마스크,"#,##0")&"개)"',
         "=IF({T:k}=-1,입력_마스크*기준_마스크,0)*1000", ""),
        (lambda r: f'=" 대형마트 홍보 : "&{t("기준_대형마트")}&"천원/주 × "&입력_대형마트&"주 (3개소)"',
         "=IF({T:k}=-1,입력_대형마트*기준_대형마트,0)*1000", ""),
        (lambda r: f'=" 조직분양 현수막 : "&{t("기준_조직현수막")}&"천원/개 × "&TEXT(기준_조직현수막_장수,"#,##0")&"장/주 × "&입력_현수막주수&"주/월(벌금 포함)"&{N(r)}'
                   f'&IF(입력_현수막감축연="","",", "&입력_현수막감축월&"월부터 "&입력_현수막감축주수&"주/월")',
         "={T:조직}*기준_조직현수막*기준_조직현수막_장수*IF({T:번호}>=계산_현수막감축,입력_현수막감축주수,입력_현수막주수)*1000", ""),
    ]))
    G.append(("판매촉진비", "분양대행", [
        (lambda r: f'=" 분양대행수수료 : "&{t("기준_분양대행")}&"천원/세대 × "&TEXT(SUM(\'계산\'!J{K["row"]["대행세대"]}:U{K["row"]["대행세대"]}),"#,##0")&"세대 (분양대행 시트)"',
         f"='분양대행 및 촉진수수료'!{{c}}{pay['agency']}*1000", ""),
        (lambda r: f'=" 조직분양 수수료("&입력_급지&") : "&{t("계산_조직단가")}&"천원/세대 × "&TEXT(SUM(\'계산\'!J{K["row"]["조직세대"]}:U{K["row"]["조직세대"]}),"#,##0")&"세대 (D+2M부터)"',
         f"='분양대행 및 촉진수수료'!{{c}}{pay['org']}*1000", ""),
        (lambda r: f'=" 상가 분양대행 : 매출액 "&TEXT(입력_상가매출,"#,##0")&"천원 × "&IF(입력_상가유형="주상복합 상가",기준_상가대행_주상복합,기준_상가대행_단지내)&"% ("&입력_상가유형&")"',
         '=IF({T:번호}=계산_상가월,입력_상가매출*IF(입력_상가유형="주상복합 상가",기준_상가대행_주상복합,기준_상가대행_단지내)/100,0)*1000', ""),
        (lambda r: f'=" 부동산MGM : "&{t("기준_MGM")}&"천원/세대 × "&TEXT(계산_MGM세대,"#,##0")&"세대(세대수의 "&기준_MGM_비율&"%, C~D급지)"&IF(계산_MGM적용=1,""," - 해당 없음")',
         f"='분양대행 및 촉진수수료'!{{c}}{pay['mgm']}*1000", ""),
        (lambda r: f'=" 상가 부동산MGM : "&{t("기준_상가MGM")}&"천원/호실 × "&입력_상가MGM&"호실"',
         "=IF({T:번호}=계산_상가MGM월,입력_상가MGM*기준_상가MGM,0)*1000", ""),
        (lambda r: f'=" 오픈 선물 : "&{t("기준_오픈선물")}&"천원/개 × "&TEXT(계산_오픈선물인원,"#,##0")&"명(B~D급지)"&IF(계산_오픈선물적용=1,""," - 해당 없음")',
         "=IF(AND(계산_오픈선물적용=1,{T:k}=-1),계산_오픈선물인원*기준_오픈선물,0)*1000", ""),
        (lambda r: f'=" 청약 선물 : "&{t("기준_청약선물")}&"천원/개 × "&TEXT(계산_청약선물인원,"#,##0")&"명(C~D급지)"&IF(계산_MGM적용=1,""," - 해당 없음")',
         "=IF(AND(계산_MGM적용=1,{T:오픈1}=1),계산_청약선물인원*기준_청약선물,0)*1000", ""),
        (lambda r: f'=" 당첨자 선물 : "&{t("기준_당첨자선물")}&"천원/개 × "&TEXT(계산_당첨자세대,"#,##0")&"세대(세대수의 "&기준_당첨자_비율&"%)"',
         "=IF(AND(계산_MGM적용=1,{T:오픈1}=1),계산_당첨자세대*기준_당첨자선물,0)*1000", ""),
        (lambda r: f'=" 계약 선물 : "&{t("기준_계약선물")}&"천원/개 × "&TEXT(계산_총세대,"#,##0")&"세대(상품권 별도)"',
         "=IF(AND(계산_MGM적용=1,{T:오픈1}=1),계산_총세대*기준_계약선물,0)*1000", ""),
        (lambda r: f'=" 계약해지세대 장기대여금 계정대체 : "&{t("기준_해지대체")}&"천원 × "&입력_해지세대&"세대"',
         "=IF({T:번호}=계산_해지월,입력_해지세대*기준_해지대체,0)*1000", ""),
    ]))
    G.append(("판매촉진비", "대납_신규_1금융", [
        (lambda r: '=" 중도금 대납이자 (대납이자 시트 참조"&IF(입력_무이자="Y",")",", 무이자 현장 아님)")',
         "=IF(입력_무이자=\"Y\",'대납이자'!{D}32*1000000,0)", ""),
    ]))
    G.append(("운반비", "운반비", [
        (lambda r: f'=" 오픈·사무실 이전·지원센터 이전 : "&{t("기준_운반비")}&"천원/회(포장이사)"&{N(r)}',
         "=(IF({T:번호}=계산_사전시작,1,0)+IF({T:번호}=계산_입주,1,0)+IF({T:번호}=계산_사무실시작,1,0))*기준_운반비*1000", ""),
    ]))
    G.append(("소모품비", "소모품_D", [
        (lambda r: f'=" 오픈(D) : "&{t("기준_소모품_D")}&"천원, D+1개월 : "&{t("기준_소모품_D1")}&"천원, D+2개월~·사전영업 : "&{t("기준_소모품_D2")}&"천원/월"',
         "=(IF({T:k}=-1,기준_소모품_D2,0)+IF({T:k}=0,기준_소모품_D,0)+IF({T:k}=1,기준_소모품_D1,0)"
         "+IF(AND({T:k}>=2,{T:MH}=1),기준_소모품_D2,0))*1000", ""),
        (lambda r: f'=" 가구비용(오픈) : "&{t("기준_가구")}&"천원"&{M1(r)}',
         "=IF({T:k}=0,기준_가구,0)*1000", ""),
        (lambda r: f'=" 입주(D) : "&{t("기준_소모품_D")}&"천원, D+1 : "&{t("기준_소모품_D1")}&"천원, 사전점검·D+2~ : "&{t("기준_소모품_입주점검")}&"천원/월"',
         "=(IF({T:사전점검}=1,기준_소모품_입주점검,0)+IF({T:mk}=0,기준_소모품_D,0)+IF({T:mk}=1,기준_소모품_D1,0)"
         "+IF(AND({T:mk}>=2,{T:입주센터}=1),기준_소모품_입주점검,0))*1000", ""),
        (lambda r: f'=" [분양사무실] 소모품 기본 : "&{t("기준_사무실_소모품")}&"천원/월"&{N(r)}',
         "={T:사무실}*기준_사무실_소모품*1000", ""),
    ]))
    G.append(("지급수수료", "발급_기본", [
        (lambda r: f'=" 발급수수료 기본 : "&{t("기준_발급_기본")}&"천원/월"&{N(r)}',
         "={T:발급}*기준_발급_기본*1000", ""),
        (lambda r: f'=" 확정일자 : "&{t("기준_확정일자")}&"원/세대 × "&기준_확정일자_횟수&"회(대출확인서·계약서) × 계약세대"',
         "={T:분양}*기준_확정일자*기준_확정일자_횟수", ""),
        (lambda r: f'=" 송금수수료 기본 : "&{t("기준_송금")}&"천원/월"&{N(r)}',
         "={T:발급}*기준_송금*1000", ""),
        (lambda r: f'=" 무인경비용역 공사비 : "&{t("기준_무인경비_공사")}&"천원"&{M1(r)}',
         "=IF({T:번호}=계산_사전시작,기준_무인경비_공사,0)*1000", ""),
        (lambda r: f'=" 무인경비용역비 : "&{t("기준_무인경비_용역")}&"천원/월"&{N(r)}',
         "={T:MH}*기준_무인경비_용역*1000", ""),
        (lambda r: f'=" 특별공급대행 : "&{t("계산_특공단가")}&"천원(특공 "&TEXT(계산_특공세대,"#,##0")&"세대 기준)"',
         "=IF({T:오픈1}=1,계산_특공단가,0)*1000", ""),
        (lambda r: f'=" 무순위청약 전산추첨 대행 : "&{t("기준_무순위")}&"천원"&IF(입력_무순위="Y",""," - 해당 없음")',
         '=IF(AND(입력_무순위="Y",{T:오픈1}=1),기준_무순위,0)*1000', ""),
        (lambda r: f'=" 계약취소주택 전산추첨 대행 : "&{t("기준_계약취소")}&"천원"&IF(입력_계약취소="Y",""," - 해당 없음")',
         '=IF(AND(입력_계약취소="Y",{T:오픈1}=1),기준_계약취소,0)*1000', ""),
        (lambda r: f'=" 전기안전관리 : "&{t("기준_전기안전")}&"천원/월, 소방안전관리 : "&{t("기준_소방안전")}&"천원/월, 오수관리 : "&{t("기준_오수관리")}&"천원/월"&{N(r)}',
         "={T:MH}*(기준_전기안전+기준_소방안전+기준_오수관리)*1000", ""),
        (lambda r: f'=" 분양가 산출용역 : "&입력_상품&"("&{t("계산_분양가용역")}&"천원)"&IF(입력_분양가용역="Y",""," - 해당 없음")',
         '=IF(AND(입력_분양가용역="Y",{T:k}=-1),계산_분양가용역,0)*1000', ""),
        (lambda r: f'=" 카페테리아(오픈/계약) : "&{t("기준_카페테리아")}&"천원"&IF(입력_카페테리아="Y",""," - 해당 없음")',
         '=IF(AND(입력_카페테리아="Y",{T:오픈1}=1),기준_카페테리아,0)*1000', ""),
        (lambda r: f'=" M/H 중개수수료 : "&입력_지역&" "&{t("계산_MH중개")}&"천원"&IF(입력_MH중개="Y",""," - 해당 없음")',
         '=IF(AND(입력_MH중개="Y",{T:번호}=계산_사전시작),계산_MH중개,0)*1000', ""),
        (lambda r: f'=" 청소 : "&{t("기준_청소단가")}&"천원/인 × (오픈1개월 "&계산_청소1&"명, 2개월 "&계산_청소2&"명, 3개월~ "&기준_청소_이후&"명, 잔여판촉 "&기준_청소_잔여&"명, 유니트 "&입력_유니트&"개)"',
         "=(IF({T:k}=-1,기준_청소_이후,0)+IF({T:k}=0,계산_청소1,0)+IF({T:k}=1,계산_청소2,0)"
         "+IF(AND({T:k}>=2,{T:MH}=1),IF({T:조직}=1,기준_청소_잔여,기준_청소_이후),0))*기준_청소단가*1000", ""),
        (lambda r: f'=" 아르바이트(여직원 1명일 경우) : "&{t("기준_아르바이트")}&"천원/월 × "&기준_아르바이트_개월&"개월"&IF(입력_아르바이트="Y",""," - 해당 없음")',
         '=IF(AND(입력_아르바이트="Y",{T:k}>=0,{T:k}<기준_아르바이트_개월),기준_아르바이트,0)*1000', ""),
        (lambda r: f'=" 상담사(재분양) : "&{t("기준_상담사")}&"천원/월 × "&기준_상담사_개월&"개월"&IF(입력_상담사="Y",""," - 해당 없음")',
         '=IF(AND(입력_상담사="Y",{T:번호}>=계산_상담사시작,{T:번호}<계산_상담사시작+기준_상담사_개월),기준_상담사,0)*1000', ""),
        (lambda r: f'=" 사전점검행사 : "&{t("계산_사전점검단가")}&"천원/세대 × "&TEXT(계산_총세대,"#,##0")&"세대(매니저 포함)"&{M1(r)}',
         "=IF({T:사전점검}=1,계산_사전점검단가*계산_총세대,0)*1000", ""),
        (lambda r: f'=" 입주지원센터 구축 : "&{t("기준_입주지원센터")}&"천원(AS센터 별도)"&{M1(r)}',
         "=IF({T:mk}=0,기준_입주지원센터,0)*1000", ""),
        (lambda r: f'=" 입주촉진책(부동산 초청, 기념품 등) : "&{t("기준_입주촉진책")}&"천원"&IF(입력_입주촉진="Y",""," - 해당 없음")',
         '=IF(AND(입력_입주촉진="Y",{T:mk}=0),기준_입주촉진책,0)*1000', ""),
        (lambda r: f'=" 입주상담사 : "&{t("기준_입주상담사_일당")}&"천원/인·일 × "&기준_입주상담사_근무일&"일 × "&계산_입주상담사인원&"명/월"&{N(r)}',
         "=IF(AND({T:mk}>=0,{T:mk}<=계산_지정기간),계산_입주상담사월,0)*1000", ""),
        (lambda r: f'=" 입주매니저 : "&{t("기준_입주매니저")}&"천원/명 × "&TEXT(계산_총세대,"#,##0")&"세대 × "&계산_매니저배수&"명(입주율 "&입력_입주율&")"&{M1(r)}',
         "=IF({T:mk}=1,기준_입주매니저*계산_총세대*계산_매니저배수,0)*1000", ""),
        (lambda r: f'=" 건축물관리대장 발급 : "&{t("기준_건축물대장")}&"천원/세대 × "&TEXT(계산_총세대,"#,##0")&"세대"&{M1(r)}',
         "=IF({T:mk}=1,기준_건축물대장*계산_총세대,0)*1000", ""),
        (lambda r: f'=" 이전등기비용 : "&{t("계산_이전등기단가")}&"천원 × "&TEXT(계산_총세대+입력_상가,"#,##0")&"(세대·상가) × "&기준_이전등기_비율&"%"&{M1(r)}',
         "=IF({T:mk}=1,계산_이전등기단가*(계산_총세대+입력_상가)*기준_이전등기_비율/100,0)*1000", ""),
        (lambda r: f'=" [분양사무실] 발급수수료 : "&{t("기준_사무실_발급")}&"천원/월"&{N(r)}',
         "={T:사무실}*기준_사무실_발급*1000", ""),
        (lambda r: f'=" [분양사무실] 보안경비용역비 : "&{t("계산_사무실보안")}&"천원/월"&{N(r)}',
         "={T:사무실}*계산_사무실보안*1000", ""),
    ]))
    G.append(("도서인쇄비", "문고", [
        (lambda r: f'=" 문고 설치(입주시) : "&{t("기준_문고")}&"천원"&{M1(r)}',
         "=IF({T:mk}=0,기준_문고,0)*1000", ""),
    ]))
    G.append(("관리비", "공가관리비", [
        (lambda r: f'=" 공가세대 관리비 : "&{t("기준_공가관리비")}&"천원/평 × 미입주세대(입주 첫달 100%) (공가세대관리비 시트)"',
         f"='공가세대관리비'!{{B}}{vac['amount']}*1000", ""),
        (lambda r: f'=" [분양사무실] 일반/공용관리비 : "&TEXT(IF(입력_통합센터="Y",기준_사무실_관리비_통합,기준_사무실_관리비),"#,##0")&"천원/월"&IF(입력_통합센터="Y","(통합센터)","")&{N(r)}',
         '={T:사무실}*IF(입력_통합센터="Y",기준_사무실_관리비_통합,기준_사무실_관리비)*1000', ""),
    ]))

    # ── 시트 쓰기
    put(ws, "A2", "='예산안 보고서'!A2", font(22, True), al=CENTER)
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 36
    ws.row_dimensions[3].height = 20
    put(ws, "A3", "='예산안 보고서'!A3", font(12), al=CENTER)
    ws.merge_cells("A3:H3")
    put(ws, "A5", '="[ "&입력_현장명&" ]"', font(14, True))
    put(ws, "H5", "(금액단위: 천원)", F9, al=RIGHT)
    for i in range(12):
        put(ws, f"{MC[i]}5", T(K, "단계", MC[i]).join(("=", "")), font(9, True, "C00000"), al=CENTER)
    heads = {"B": "항   목", "D": "세   목", "E": '=(계산_연도-2001)&"년 실적"', "F": '=(계산_연도-2000)&"년 예산"',
             "G": "예산산출근거(상세)", "H": "비      고", "I": '=(계산_연도-2000)&"년 합계"'}
    for col, h in heads.items():
        put(ws, f"{col}6", h, F10B, al=CENTER)
    for i in range(12):
        put(ws, f"{MC[i]}6", f"{i + 1}월", F10B, al=CENTER)
    ws.merge_cells("B6:C6")
    box(ws, "A6:U6", med, thin)
    for c in ws[6]:
        if c.column <= 21:
            c.border = Border(left=c.border.left, right=c.border.right, top=med, bottom=dbl)

    r = 7
    out = dict(groups=[])
    kij = lambda key: f"=INDEX('기준단가'!$D:$D,MATCH(\"{key}\",'기준단가'!$K:$K,0))"
    calc_rows = K["row"]
    for item, sub_key, lines in G:
        first = r
        for k, (gf, mf, note) in enumerate(lines + [(None, None, None)]):
            if k == 0:
                put(ws, f"B{r}", item, F10, al=Alignment(horizontal="distributed", vertical="center"))
                put(ws, f"D{r}", kij(sub_key), F10, al=Alignment(vertical="center"))
            if gf is None:          # 직접 입력 줄
                put(ws, f"G{r}", None, F10, FILL_IN, LEFT)
                for i in range(12):
                    put(ws, f"{MC[i]}{r}", None, font(10.5), FILL_IN, VC, NUMF)
                put(ws, f"H{r}", "직접입력", F9G, al=CENTER)
            else:
                put(ws, f"G{r}", gf(r), F10, al=LEFT)
                for i in range(12):
                    c = MC[i]
                    f = mf.replace("{c}", c).replace("{B}", CL(9 + i)).replace("{D}", CL(11 + i))
                    while "{T:" in f:
                        a = f.index("{T:")
                        b = f.index("}", a)
                        f = f[:a] + T(K, f[a + 3:b], c) + f[b + 1:]
                    while "{R:" in f:
                        a = f.index("{R:")
                        b = f.index("}", a)
                        f = f[:a] + str(calc_rows[f[a + 3:b]]) + f[b + 1:]
                    put(ws, f"{c}{r}", f, font(10.5), al=VC, nf=NUMF)
                put(ws, f"H{r}", note or None, F9, al=CENTER)
            put(ws, f"E{r}", None, F10, al=VC, nf=NUMF)
            put(ws, f"F{r}", f"=I{r}", F10, FILL_BUD, VC, NUMF)
            put(ws, f"I{r}", f"=SUM(J{r}:U{r})", font(10.5), al=VC, nf=NUMF)
            lines = getattr(gf, "lines", 1)                  # 합친 세목(식대·여비·통신요금·전기료)은 여러 줄
            ws.row_dimensions[r].height = 21 if lines <= 1 else 15 * lines + 6
            r += 1
        last = r - 1
        put(ws, f"D{r}", "소  계", F10, FILL_SUB, CENTER)
        for col in ["E", "F", "I"] + MC:
            put(ws, f"{col}{r}", f"=SUM({col}{first}:{col}{last})", F10B, FILL_BUD if col == "F" else FILL_SUB, nf=NUMF)
        for col in "BCGH":
            ws[f"{col}{r}"].fill = FILL_SUB
        box(ws, f"A{first}:U{r}", thin, hair)
        for c in ws[r]:
            if c.column <= 21:
                c.border = Border(left=c.border.left, right=c.border.right, top=thin, bottom=thin)
        out["groups"].append(dict(item=item, sub_key=sub_key, first=first, last=last, sub=r))
        r += 1
    put(ws, f"A{r}", "합            계", F10B, FILL_TOT, CENTER)
    ws.merge_cells(f"A{r}:D{r}")
    subs = [g["sub"] for g in out["groups"]]
    for col in ["E", "F", "I"] + MC:
        put(ws, f"{col}{r}", "=" + "+".join(f"{col}{s}" for s in subs), F10B, FILL_TOT, nf=NUMF)
    for col in "GH":
        ws[f"{col}{r}"].fill = FILL_TOT
    box(ws, f"A{r}:U{r}", med, thin)
    out["total_row"] = r
    for rr in range(7, r + 1):
        for col in "F":
            c = ws[f"{col}{rr}"]
            c.border = Border(left=med, right=med, top=c.border.top, bottom=c.border.bottom)
    put(ws, f"G{r + 2}", "※ 노란 줄은 직접 입력. 자동 계산 칸도 필요하면 숫자로 덮어써도 됩니다(덮어쓴 칸은 기준이 바뀌어도 그대로).", F9)
    put(ws, f"G{r + 3}", "※ 금액은 원 단위로 계산되어 천원으로 표시됩니다. 0인 칸은 비어 보입니다.", F9)
    setw(ws, dict(A=0.9, B=9.9, C=0.9, D=12, E=10.3, F=11.5, G=72, H=8, I=10.5, **{c: 9 for c in MC}))
    ws.freeze_panes = "E7"
    ws.print_area = f"A1:U{r}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "5:6"
    return out


# ═════════════════════════════════════════════════════════════════
# 예산안 보고서
def build_report(ws, groups):
    put(ws, "A2", '="[ "&입력_현장명&" ] "&계산_연도&"년 판관비 예산안"', font(22, True), al=CENTER)
    ws.merge_cells("A2:P2")
    ws.row_dimensions[2].height = 36
    ws.row_dimensions[3].height = 20
    put(ws, "A3", '="( "&TEXT(계산_작성일,"yyyy. mm. dd")&" )"', font(12), al=CENTER)
    ws.merge_cells("A3:P3")
    put(ws, "A6", '="[ "&입력_현장명&" ]"', font(14, True))
    put(ws, "P6", "(금액단위: 천원)", F9, al=RIGHT)
    py, cy = '(계산_연도-2001)&"년', '(계산_연도-2000)&"년'
    H = {"B7": "항   목", "D7": "세   목", "E7": f'={py} 예산"', "H7": f'={py} 실적"', "K7": "예산 比",
         "L7": f'={py} 예산대비 실적차이 분석(상세)"', "M7": f'={cy} 예산"', "N7": "실적 比",
         "O7": "전년 실적 대비 증감요인(상세)", "P7": "비   고",
         "E8": "총예산", "F8": "최초편성예산", "G8": "증액금액", "H8": "연간전망", "I8": "1~9월 실적", "J8": "10~12월 실적"}
    for a, v in H.items():
        put(ws, a, v, F10B, al=CENTER)
    for m in ("B7:C8", "D7:D8", "E7:G7", "H7:J7", "K7:K8", "L7:L8", "M7:M8", "N7:N8", "O7:O8", "P7:P8", "A7:A8"):
        ws.merge_cells(m)
    box(ws, "A7:P8", med, thin)
    r = 9
    subs = []
    by_item = []
    for g in groups["groups"]:
        if by_item and by_item[-1][0] == g["item"]:
            by_item[-1][1].append(g)
        else:
            by_item.append((g["item"], [g]))
    new_site = ('IF(AND(계산_오픈>=계산_연도*12+1,계산_오픈<=계산_연도*12+12),'
                '" 신규현장("&(계산_연도-2000)&"."&TEXT(계산_오픈-계산_연도*12,"00")&" 오픈)","")')
    for item, gs in by_item:
        first = r
        for k, g in enumerate(gs):
            if k == 0:
                put(ws, f"B{r}", item, F10, al=Alignment(horizontal="distributed", vertical="center"))
            put(ws, f"D{r}", f"='{BUDGET}'!D{g['first']}", F10, al=LEFT)
            for col in "EFGHIJ":
                put(ws, f"{col}{r}", None, F10, FILL_IN, nf=NUMF)
            put(ws, f"K{r}", f'=IF(H{r}-E{r}=0,"",IF(E{r}=0,100%,(H{r}-E{r})/E{r}))', F10, FILL_BUD, nf="0.0%")
            put(ws, f"L{r}", None, F10, FILL_IN, LEFT)
            put(ws, f"M{r}", f"=SUM('{BUDGET}'!F{g['first']}:F{g['last']})", F10B, nf=NUMF)
            put(ws, f"N{r}", f'=IF(M{r}-H{r}=0,"",IF(H{r}=0,100%,(M{r}-H{r})/H{r}))', F10, FILL_BUD, nf="0.0%")
            put(ws, f"O{r}", f'=IF(M{r}=0,"",{new_site})', F10, FILL_DEF, LEFT)
            put(ws, f"P{r}", None, F10, FILL_IN)
            ws.row_dimensions[r].height = 19.5
            r += 1
        put(ws, f"D{r}", "소  계", F10, FILL_SUB, CENTER)
        for col in "EFGHIJM":
            put(ws, f"{col}{r}", f"=SUM({col}{first}:{col}{r - 1})", F10B, FILL_SUB, nf=NUMF)
        put(ws, f"K{r}", f'=IF(H{r}-E{r}=0,"",IF(E{r}=0,100%,(H{r}-E{r})/E{r}))', F10, FILL_SUB, nf="0.0%")
        put(ws, f"N{r}", f'=IF(M{r}-H{r}=0,"",IF(H{r}=0,100%,(M{r}-H{r})/H{r}))', F10, FILL_SUB, nf="0.0%")
        for col in "BCLOP":
            ws[f"{col}{r}"].fill = FILL_SUB
        box(ws, f"A{first}:P{r}", thin, hair)
        subs.append(r)
        r += 1
    put(ws, f"A{r}", "합            계", F10B, FILL_TOT, CENTER)
    ws.merge_cells(f"A{r}:D{r}")
    for col in "EFGHIJM":
        put(ws, f"{col}{r}", "=" + "+".join(f"{col}{s}" for s in subs), F10B, FILL_TOT, nf=NUMF)
    put(ws, f"K{r}", f'=IF(H{r}-E{r}=0,"",IF(E{r}=0,100%,(H{r}-E{r})/E{r}))', F10B, FILL_TOT, nf="0.0%")
    put(ws, f"N{r}", f'=IF(M{r}-H{r}=0,"",IF(H{r}=0,100%,(M{r}-H{r})/H{r}))', F10B, FILL_TOT, nf="0.0%")
    for col in "LOP":
        ws[f"{col}{r}"].fill = FILL_TOT
    box(ws, f"A{r}:P{r}", med, thin)
    for rr in range(9, r + 1):
        c = ws[f"M{rr}"]
        c.border = Border(left=med, right=med, top=c.border.top, bottom=c.border.bottom)
    put(ws, f"B{r + 2}", "※ 노란 칸: 전년 예산·실적(ERP)과 차이 분석 입력.  M열은 27년 예산 시트 자동 합계.  O열(초록)은 기본 문구이며 덮어써도 됩니다.", F9)
    setw(ws, dict(A=0.9, B=9.9, C=0.9, D=13.8, E=11.7, F=11, G=10, H=11, I=10, J=10, K=8.3, L=30.1, M=11.5, N=8.3, O=40.8, P=8.8))
    ws.print_area = f"A1:P{r}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.freeze_panes = "E9"


# ═════════════════════════════════════════════════════════════════
# 기본정보
def build_basic(ws, I):
    t1, t2 = I["types"]
    put(ws, "B2", "■ 기본정보", font(12, True))
    put(ws, "C3", "구  분", F10B, al=CENTER)
    ws.merge_cells("C3:E3")
    put(ws, "F3", "세 부 내 용", F10B, al=CENTER)
    put(ws, "G3", "비        고", F10B, al=CENTER)
    put(ws, "C4", "현장명", F10, FILL_SUB, CENTER)
    ws.merge_cells("C4:E4")
    put(ws, "F4", "=입력_현장명", F10)
    put(ws, "D5", "계", F10, al=CENTER)
    ws.merge_cells("D5:E5")
    put(ws, "F5", "=입력_총세대", F10B, nf="#,##0")
    put(ws, "B6", "세대수", F10, FILL_SUB, CENTER)
    ws.merge_cells(f"B6:C{5 + NTYPES}")
    for k in range(NTYPES):
        rr = 6 + k
        put(ws, f"D{rr}", f"=IF('현장입력'!C{t1 + k}=\"\",\"\",'현장입력'!C{t1 + k})", F10, al=CENTER)
        ws.merge_cells(f"D{rr}:E{rr}")
        put(ws, f"F{rr}", f"=IF('현장입력'!D{t1 + k}=\"\",\"\",'현장입력'!D{t1 + k})", F10, nf="#,##0")
    rr = 6 + NTYPES
    for label, f in (("담당소장", "=입력_담당소장"), ("착공년월", "=입력_착공"), ("준공년월", "=입력_준공"),
                     ("시행법인", "=입력_시행"), ("시공법인", "=계산_시공"), ("분양성(급지)", "=입력_급지"),
                     ("지역", "=입력_지역")):
        put(ws, f"C{rr}", label, F10, FILL_SUB, CENTER)
        ws.merge_cells(f"C{rr}:E{rr}")
        put(ws, f"F{rr}", f + '&""', F10)
        rr += 1
    put(ws, f"B{rr}", "※ PM실 작성 (현장입력 시트에서 자동)", F9)
    box(ws, f"B3:G{rr - 1}", med, thin)
    setw(ws, dict(A=2, B=7, C=6, D=8, E=6, F=44, G=24))


# ═════════════════════════════════════════════════════════════════
def build_guide(ws):
    lines = [
        ("판관비 예산안 자동작성 양식 — 사용법", font(16, True)),
        ("", None),
        ("■ 새 현장 예산안 만들기", font(12, True)),
        ("1. 이 파일을 복사해 '(현장명) 2027년 사업비 예산편성.xlsx' 처럼 다른 이름으로 저장합니다.", F10),
        ("2. [현장입력] 노란 칸을 채웁니다.  필수: 현장명, 지역, 분양성(급지), 세대 구성(타입·세대수), 오픈월(분양월).", F10),
        ("     입주 현장은 사전점검월·입주개시월, 분양사무실만 운영하는 현장은 분양사무실 운영 시작월(전년부터 1년 내내 운영이면 '연중 운영' Y)을 넣습니다.", F10),
        ("     오른쪽 '월별 단계 미리보기'에서 달마다 사전·오픈·D+1·조직분양·입주 단계가 맞게 잡혔는지 확인하세요.", F10),
        ("3. [분양율,입주율] 초록 칸(월별 분양·입주세대)은 과거 작성사례 곡선으로 자동 채워집니다. 실제 계획으로 덮어쓰세요.", F10),
        ("     (오픈 다음 달 30% → 누계 50·60·70·80·85·90·95·100%, 목표 분양률에서 멈춤 / 입주 첫달 0% → 30·60·80·90·95·100%)", F9),
        ("4. [27년 예산] 기준안에 따라 항목별·월별 금액과 산출근거가 자동으로 나옵니다.", F10),
        ("     노란 줄은 기준안에 없는 현장별 비용을 직접 넣는 줄입니다. 자동 칸도 숫자로 덮어쓸 수 있습니다.", F10),
        ("5. [예산안 보고서] 노란 칸에 전년 예산·실적(ERP)과 차이 분석을 넣으면 보고서가 완성됩니다.", F10),
        ("", None),
        ("■ 해마다 새 기준안 반영", font(12, True)),
        ("1. 새 'PM팀 판관비 예산기준(안)_날짜.xlsx' 를 받으면 이 폴더의 '기준안 반영.bat' 을 실행해 그 파일을 고릅니다.", F10),
        ("2. 기준안의 금액을 읽어 [기준단가] 표를 갱신한 새 양식 파일(… 양식_연도.xlsx)을 만듭니다. 원래 양식은 그대로 둡니다.", F10),
        ("3. [기준안 반영결과] 시트에서 바뀐 값·확인 필요 값·읽지 못한 기준안 줄을 확인합니다.", F10),
        ("     주황 칸 = 이번에 바뀐 값, 빨강 칸 = 확인 필요(기존 값 유지).  [기준단가] F열은 언제든 직접 고쳐도 됩니다.", F10),
        ("4. 기준안에 새 항목이 생기면(반영결과 3번 표) 그 항목은 [27년 예산] 노란 줄에 직접 넣거나, 양식에 행을 추가해야 합니다.", F10),
        ("", None),
        ("■ 자동작성 규칙 (과거 PM팀 예산편성 파일에서 정리)", font(12, True)),
        ("· 단계: 사전영업 = 오픈 전 1개월, 오픈월 = D, D+1, D+2 …  조직분양은 D+2M부터(기준안).  M/H 운영 = 사전영업월 ~ 종료월(기본 연말).", F10),
        ("· 오픈 전월(사전): 판촉물(갑티슈·물티슈·종이컵·크린홈·마스크), 대형마트 홍보, 오픈 선물, 분양가 산출용역, 숙소비 일시분", F10),
        ("· 사전영업 시작월: M/H 중개수수료, 무인경비 공사비, 운반비, 보험료(이후 매년 같은 달)", F10),
        ("· 오픈월(D): 식대·여비 오픈단가, 가구, 소모품 4,000, 청약·당첨자·계약 선물, 내용증명, 특공·무순위·계약취소 대행, 카페테리아, 주차장, 청소 3~4명", F10),
        ("· 첫 계약월(D+1): 분양대행수수료(계약세대), 부동산MGM(세대수 30%).  인지세·확정일자는 월별 계약세대 기준", F10),
        ("· 입주: 사전점검월 행사·소모품, 입주개시월 지원센터·문고·촉진책·숙소, 입주+1월 입주매니저·건축물대장·이전등기, 공가세대관리비는 입주개시월부터", F10),
        ("· 분양사무실 운영기간(시작월~종료월, 연중 운영이면 1~12월): 통신·우편·복합기·관리비·소모품·발급수수료·보안경비 월 단가. 운반비(사무실 이전)는 올해 시작월에만", F10),
        ("", None),
        ("■ 색", font(12, True)),
        ("노란 칸 = 입력 · 초록 칸 = 자동 기본값(덮어쓰기 가능) · 파란 칸 = 예산 열 · 회색 = 소계 · 주황 = 합계", F10),
        ("[계산] 시트는 월별 단계 판단표(수정 불필요), [기준단가] 시트는 기준안 단가표입니다.", F10),
    ]
    for i, (t, f) in enumerate(lines, start=1):
        if f:
            put(ws, f"B{i}", t, f)
    fill = PatternFill("solid", fgColor="FFF2CC")
    put(ws, "B34", "노란 칸 예시", F10, fill)
    put(ws, "C34", "초록 칸 예시", F10, FILL_DEF)
    setw(ws, dict(A=2, B=130))


if __name__ == "__main__":
    kijun = sys.argv[1] if len(sys.argv) > 1 else None
    out = sys.argv[2] if len(sys.argv) > 2 else "out.xlsx"
    build(out, kijun)
    print("만듦:", out)
