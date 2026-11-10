# -*- coding: utf-8 -*-
"""사업관리비 실적계획 자동작성 — 시험·예시용 가짜 자료 만들기(실제 현장·사람 자료는 저장소에 두지 않는다).

  python tests\\siljeok_samples.py            docs\\siljeok\\examples\\ 에 예시 3개를 다시 만든다(화면의 '예시로 해 보기' 가 쓴다)

  prev_budget_xlsx() → bytes     전년도 예산안(회사 「사업비 예산편성」 양식의 「26년 예산」 시트 모양, 값은 수식 — 저장 값 없음)
  ledger_xlsx(rows, carry) → bytes   SAP G/L 내역 내보내기 모양(Data 시트, 머리글 21열)
  ledger_rows() → (1-8월 줄, 9-11월 줄, 9-11월 이월잔액 줄)
금년도 예산안 예시는 예산안 자동작성의 「(예시3) 기공급 잔여분양」 결과 파일(예시\\, docs\\examples\\)을 그대로 쓴다.
"""
import datetime
import io
import os
import sys

import openpyxl
from openpyxl.utils import get_column_letter as L

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SITE = "예시3 기공급 잔여분양"
CODE = "P9999"
PREV_YEAR = 2026
NEXT_EXAMPLE = "(예시3) 기공급 잔여분양 - 26.05 오픈, 27.06 MH종료.xlsx"
STAFF = ["홍길동", "김영희"]                     # 현장 담당자(실적에 넣을 작성자)
OTHERS = ["이재무", "박세무"]                    # 재무·세무 처리 전표(빼는 작성자)

# 전년도 예산: (항목, [(세목, 산출근거, {월: 산식})…])
PREV_BUDGET = [
    ("복리후생비", [("식대", " 식대 - 오픈2개월 : 3,000천원/월 × 2개월\n      - 기타 : 1,800천원/월 × 6개월",
                 {5: "3000*1000", 6: "3000*1000", 7: "1800*1000", 8: "1800*1000", 9: "1800*1000", 10: "1800*1000", 11: "1800*1000", 12: "1800*1000"}),
                ("기타", " 숙소비 : 1,000천원/월 × 3개월(일시) × 2호실", {4: "2*1000*3*1000"})]),
    ("여비교통비", [("국내출장비", " 유류비, 숙박비 : 1,500천원/월 × 8개월", {m: "1500*1000" for m in range(5, 13)})]),
    ("통신비", [("통신요금", " 통신요금 및 우편료 : 500천원/월 × 9개월", {m: "500*1000" for m in range(4, 13)}),
             ("우편료", " 등기 : 4.5천원/건 × 400세대", {6: "4.5*400*1000"})]),
    ("수도광열비", [("전기료", " 전기료 : 6,000천원/월 × 9개월", {m: "6000*1000" for m in range(4, 13)})]),
    ("세금과공과", [("제세금&수수료", " 인지세 : 75천원/세대 × 400세대", {6: "75*400*1000"})]),
    ("광고선전비", [("판촉물", " 갑티슈 : 1.5천원/개 × 10,000개", {4: "10000*1.5*1000"}),
               ("사인물", " 현수막 : 40천원/개 × 300장/주 × 4주 × 3개월", {5: "40*300*4*1000", 6: "40*300*4*1000", 7: "40*300*4*1000"})]),
    ("판매촉진비", [("분양대행수수료", " 분양대행수수료 : 2,000천원/세대 × 120세대", {6: "120*2000*1000"})]),
    ("소모품비", [("일반소모품", " 오픈 : 4,000천원, 기타 : 1,000천원/월", {4: "4000*1000", **{m: "1000*1000" for m in range(5, 13)}})]),
    ("지급수수료", [("용역-기타", " 청소 : 3,500천원/인 × 2명 × 8개월", {m: "3500*2*1000" for m in range(5, 13)}),
               ("용역-보안", " 무인경비 : 400천원/월 × 9개월", {m: "400*1000" for m in range(4, 13)})]),
    ("관리비", [("관리비", " 공가세대 관리비 : 6천원/평 × 25평 × 3세대", {m: "6*25*3*1000" for m in range(1, 13)})]),
]


def _eval(expr):
    return eval(expr, {"__builtins__": {}}, {})          # 숫자·사칙연산만 든 위 표의 산식


def prev_budget_total():
    return sum(_eval(f) for _, lines in PREV_BUDGET for _, _, ms in lines for f in ms.values())


def prev_budget_by_item():
    return {item: sum(_eval(f) for _, _, ms in lines for f in ms.values()) for item, lines in PREV_BUDGET}


def prev_budget_xlsx(site=SITE, year=PREV_YEAR):
    """회사 예산편성 양식의 예산 시트 모양(머리글 6행, B 항목 · D 세목 · F 예산 · G 산출근거 · I 합계 · J~U 월별, 소 계·합 계)"""
    yy = f"{year % 100:02d}"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{yy}년 예산"
    ws["A2"] = f"[ {site} ] {year}년 판관비 예산안"
    ws["A3"] = f"( {year - 1}. 11. 08 )"
    ws["A5"] = f"[ {site} ]"
    ws["H5"] = "(금액단위: 천원)"
    heads = {2: "항   목", 4: "세   목", 5: f"{(year - 1) % 100:02d}년 실적", 6: f"{yy}년 예산", 7: "예산산출근거(상세)", 8: "비      고", 9: f"{yy}년 합계"}
    for c, t in heads.items():
        ws.cell(6, c, t)
    for m in range(12):
        ws.cell(6, 10 + m, f"{m + 1}월")
    r, subs = 7, []
    for item, lines in PREV_BUDGET:
        first = r
        for k, (semok, basis, ms) in enumerate(lines):
            if k == 0:
                ws.cell(r, 2, item)
            ws.cell(r, 4, semok)
            ws.cell(r, 6, f"=I{r}")
            ws.cell(r, 7, basis)
            ws.cell(r, 9, f"=SUM(J{r}:U{r})")
            for m, f in ms.items():
                ws.cell(r, 9 + m, "=" + f)
            r += 1
        ws.cell(r, 4, "소  계")
        for c in (5, 6):
            ws.cell(r, c, f"=SUM({L(c)}{first}:{L(c)}{r - 1})")
        ws.cell(r, 9, f"=SUM(J{r}:U{r})")
        for c in range(10, 22):
            ws.cell(r, c, f"=SUM({L(c)}{first}:{L(c)}{r - 1})")
        subs.append(r)
        r += 1
    ws.cell(r, 1, "합  계")
    for c in [5, 6] + list(range(10, 22)):
        ws.cell(r, c, "=" + "+".join(f"{L(c)}{x}" for x in subs))
    ws.cell(r, 9, f"=SUM(J{r}:U{r})")
    for row in ws.iter_rows(min_row=7, max_row=r, min_col=5, max_col=21):
        for c in row:
            if c.column != 7:
                c.number_format = '#,##0,;[Red]-#,##0,;"-"'
    ws.column_dimensions["G"].width = 60
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


HEAD = ["전기일자", "회사코드", "귀속정보", "귀속정보", "적요", "G/L 계정", "G/L 계정명", "차변금액", "대변금액", "차액", "거래처명",
        "사업자번호(주민등록번호)", "거래처코드", "전표번호", "세금코드", "세금코드명", "전표유형", "전표유형명", "생성자", "생성자명", "역분개전표번호"]
ACCT = {"복리후생비": ("51015603", "판매관리비(복리후생비)_분양경비_D"), "통신비": ("51015607", "판매관리비(통신비)_분양경비_D"),
        "수도광열비": ("51015609", "판매관리비(수도광열비)_분양경비_D"), "세금과공과": ("51015614", "판매관리비(세금과공과)_분양경비_D"),
        "광고선전비": ("51015623", "판매관리비(광고선전비)_분양경비_D"), "소모품비": ("51015639", "판매관리비(소모품비)_분양경비_D"),
        "지급수수료": ("51015641", "판매관리비(지급수수료)_분양경비_D"), "지급수수료기타": ("51015648", "판매관리비(지급수수료)_기타_D"),
        "여비교통비": ("51015669", "판매관리비(여비교통비)_분양경비_D"), "입주관리비": ("51015631", "판매관리비(분양관리비)_입주관리비_D"),
        "급료": ("51015601", "판매관리비(급료및수당)_분양경비_D"), "판매촉진비": ("51015627", "판매관리비(판매촉진비)_분양경비_D")}


def ledger_rows(year=PREV_YEAR):
    """(1-8월 파일 줄, 9-11월 파일 줄, 9-11월 파일 맨 위 이월잔액 줄) — 줄 = (일자, 적요, 계정키, 차변, 대변, 거래처, 전표번호, 전표유형명, 작성자)"""
    a, b = [], []
    doc = [2640060000]

    def add(out, m, d, memo, acct, amt, vendor, who, kind="[경비] 법인카드"):
        doc[0] += 1
        debit, credit = (amt, 0) if amt >= 0 else (0, -amt)
        out.append((datetime.datetime(year, m, d), memo, acct, debit, credit, vendor, str(doc[0]), kind, who))

    for m in range(4, 12):
        out = a if m <= 8 else b
        add(out, m, 5, "식대", "복리후생비", 1200000 + 10000 * m, "(법인카드)농협", STAFF[0])
        add(out, m, 12, f"[예시현장] {m:02d}월 업무활동비 - 복리후생비", "복리후생비", 340000, "(법인카드)농협", STAFF[1])
        add(out, m, 20, f"(예시현장) 통신비(대표전화요금) {m:02d}월", "통신비", 181500 + 300 * m, "케이티", STAFF[1])
        add(out, m, 21, f"M/H 전기요금 {m:02d}월", "수도광열비", 4150000 + 50000 * (m % 3), "한국전력공사", STAFF[0], "[경비] 기타증빙")
        add(out, m, 25, f"청소용역 {m:02d}월", "지급수수료", 7000000, "(주)예시클린", STAFF[0], "[경비] 전자세금계산서")
        add(out, m, 26, "소모품 구입", "소모품비", 420000 + 15000 * m, "예시문구", STAFF[1])
        add(out, m, 27, "유류비", "여비교통비", 250000 + 5000 * m, STAFF[0], STAFF[0], "[경비] 개인카드")
        add(out, m, 28, f"공가세대 관리비 ({m - 1:02d}월분)", "입주관리비", 452000 - 1000 * m, "예시아파트 관리사무소", STAFF[1], "[경비] 기타증빙")
    add(a, 6, 15, "조직분양 현수막 6월", "광고선전비", 48000000, "(주)예시광고", STAFF[0], "[경비] 전자세금계산서")
    add(a, 7, 15, "조직분양 현수막 7월", "광고선전비", 48000000, "(주)예시광고", STAFF[0], "[경비] 전자세금계산서")
    add(a, 5, 30, "송금수수료", "지급수수료기타", 30000, "농협은행", STAFF[1], "[경비] 기타증빙")
    add(b, 10, 30, "송금수수료", "지급수수료기타", 30000, "농협은행", STAFF[1], "[경비] 기타증빙")
    # 재무·세무가 처리한 전표(실적에서 빼는 것)
    add(a, 6, 30, "전자수입인지", "세금과공과", 30000000, "예시현장 컨소시엄", OTHERS[0], "[재무] 구매처 전기")
    add(a, 7, 31, "26.5~7월 M/H 인건비", "급료", -42326572, "예시현장 컨소시엄", OTHERS[0], "[재무] 고객 전기")
    add(a, 8, 26, "26.5~7월 M/H 인건비", "급료", 25395944, "예시현장 컨소시엄", OTHERS[0], "[재무] 구매처 전기")
    add(a, 7, 20, "재산세(토지분)", "세금과공과", 1629460, "", OTHERS[1], "[재무] 일반회계")
    add(b, 9, 20, "재산세(주택분)", "세금과공과", 2351760, "", OTHERS[1], "[재무] 일반회계")
    # 다른 현장 비용이 잘못 들어왔다가 재무에서 되돌린 것(담당자 전표만 넣으면 실적이 부풀려짐 → 화면이 알려 줘야 함)
    add(a, 6, 22, "[다른현장] 분양대행 수수료 지급 건(7차)", "판매촉진비", 41100000, "(주)예시대행", STAFF[0], "[경비] 전자세금계산서")
    add(a, 7, 1, "[귀속부서대체] [다른현장] 분양대행 수수료 지급 건(7차)", "판매촉진비", -41100000, "(주)예시대행", OTHERS[0], "[재무] 일반회계")
    # 전년도가 아닌 전표 한 줄(빼야 함), 두 파일에 겹치는 8월 말 전표 두 줄(한 번만 넣어야 함)
    doc[0] += 1
    a.insert(0, (datetime.datetime(year - 1, 12, 31), "전년 귀속 정산", "복리후생비", 999000, 0, "(법인카드)농협", str(doc[0]), "[경비] 법인카드", STAFF[0]))
    dup = [r for r in a if r[0].month == 8 and r[2] in ("소모품비", "여비교통비")]
    b[0:0] = dup
    carry = [("이월잔액", k, 12345678 + i * 1000, 2345678) for i, k in enumerate(["복리후생비", "통신비", "수도광열비"])]
    return a, b, carry


def ledger_xlsx(rows, carry=(), code=CODE + "D", code_name="예시현장(B1)_사업"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(HEAD)
    for memo, key, debit, credit in carry:                       # 이월잔액: 전기일자·전표번호·작성자 없음
        acct, name = ACCT[key]
        ws.append([None, "AA00", code, code_name, memo, acct, name, debit, credit, debit - credit] + [""] * 11)
    for d, memo, key, debit, credit, vendor, doc, kind, who in rows:
        acct, name = ACCT[key]
        ws.append([d, "AA00", code, code_name, memo, acct, name, debit, credit, debit - credit, vendor, "000-00-00000", "3000000", doc, "", "",
                   "EA", kind, "20990000", who, ""])
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def expected_actual(creators, year=PREV_YEAR):
    """고른 작성자의 실적 기대값(엔진과 따로 센 것): {비목: [1~8월, 9~12월]} — 전년도 아닌 줄 제외, 겹친 줄 한 번만"""
    a, b, _ = ledger_rows(year)
    item = {"지급수수료기타": "지급수수료", "입주관리비": "관리비", "급료": "급료및수당"}
    out, seen = {}, set()
    for rows in (a, b):
        for d, memo, key, debit, credit, vendor, doc, kind, who in rows:
            if d.year != year or who not in creators or doc in seen:
                continue
            seen.add(doc)
            slot = out.setdefault(item.get(key, key), [0.0, 0.0])
            slot[0 if d.month <= 8 else 1] += debit - credit
    return out


def files():
    """예시 파일 3개: {파일 이름: bytes}"""
    a, b, carry = ledger_rows()
    return {"(예시) 전년도 예산 - 26년 예산안.xlsx": prev_budget_xlsx(),
            "(예시) 사용내역 1-8월.xlsx": ledger_xlsx(a),
            "(예시) 사용내역 9-11월.xlsx": ledger_xlsx(b, carry)}


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    out = os.path.join(ROOT, "docs", "siljeok", "examples")
    os.makedirs(out, exist_ok=True)
    for name, data in files().items():
        with open(os.path.join(out, name), "wb") as f:
            f.write(data)
        print(f"  {name}  {len(data):,} bytes")
    print("예시 저장:", out, f"(전년도 예산 합계 {prev_budget_total():,.0f}원)")
