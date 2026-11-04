# -*- coding: utf-8 -*-
"""지출계획서 엔진·엑셀 시험(서버 없이): 평택고덕 2차(오픈 '26.11, 10~12월) · 견본주택 운영(오픈 일정 없음, 중도금대출·선착순) · 입주 · 분양사무실 네 사례
  + 운영 구분 변형(오픈 구분과의 차이, 오픈월 무시, M/H 종료월, 둘 다 고른 경우, 숙소 시작월).
  - 월별 금액이 기준단가와 단계 규칙대로 나오는지(기대값 고정) · 산식을 다시 계산한 값 = 금액 · 비고 글
  - 엑셀: 수식 칸 수, 숨긴 입력 시트 되읽기(roundtrip), 결재란 칸 수, 인쇄 설정
  python tests\\jichul_test.py            (끝에 Excel 이 있으면 PDF 로 뽑아 tests/jichul_*.png 로 저장)
"""
import io
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tool"))
import openpyxl                     # noqa: E402
import apply_kijun                  # noqa: E402
import jichul                       # noqa: E402
import jichul_xlsx                  # noqa: E402
import site_io                      # noqa: E402

y, tpl = apply_kijun.templates()[0]
info = site_io.template_info(tpl)
K = {p["key"]: p["value"] for p in info["params"] if "key" in p}
print(f"기준: {y}년 양식 {os.path.basename(tpl)} (기준안 {info['kijun']}, 단가 {len(K)}개)")
fails = []


def check(name, got, exp):
    ok = abs(float(got) - float(exp)) < 0.5
    if not ok:
        fails.append(f"{name}: {got} != {exp}")
    print(f"  {'OK ' if ok else 'XX '} {name}: {got:,.0f}" + ("" if ok else f" (기대 {exp:,.0f})"))


def row(res, key):
    for g in res["groups"]:
        for r in g["rows"]:
            if r["key"] == key:
                return r
    return None


def verify_formulas(res):
    n = 0
    for g in res["groups"]:
        for r in g["rows"]:
            for c in r["cells"]:
                if c["f"]:
                    n += 1
                    v = jichul.evaluate(c["f"])
                    assert v is not None and abs(v - c["v"]) < 0.5, (r["key"], c)
    return n


# ── 사례 1: 평택고덕 2차 오픈(사전 10월 · 오픈 11월 · 계약 12월), 623세대, 유니트 4, 숙소 4호, PC 7
d1 = dict(kinds=["mh"], fields={"현장명": "평택고덕2차", "작성일": "2026-10-06", "시작": [2026, 10], "개월수": 3, "지역": "수도권",
          "세대수": 623, "오픈": [2026, 11], "사전영업": 1, "유니트": 4, "숙소호실": 4, "PC": 7}, months=[{}, {}, {}],
          approvers={"list": [{"pos": "대리"}, {"pos": "과장"}, {"pos": "부장"}, {"pos": "상무"}, {"pos": "실장"}]})
r1 = jichul.calc(d1, K)
print("\n[사례1]", r1["title"], [m["label"] for m in r1["months"]], "합계", f"{r1['total_all']:,.0f}", r1["warnings"])
assert r1["kind"] == "M/H 오픈비용" and [m["label"] for m in r1["months"]] == ["(사전)", "(오픈)", "(계약)"]
check("식대 10월(기타 10천×4명×30)", row(r1, "식대")["cells"][0]["v"], K["식대_기타"] * 1000)
check("식대 11월(오픈 기준 3,000천원)", row(r1, "식대")["cells"][1]["v"], K["식대_오픈"] * 1000)
check("숙박비 10월(1,000×3개월×4호)", row(r1, "숙박비")["cells"][0]["v"], K["숙소비_월"] * K["숙소_개월"] * 4 * 1000)
check("숙소관리비(200×4호)", row(r1, "숙소관리비")["cells"][2]["v"], K["숙소관리비"] * 4 * 1000)
check("사무용가구 10월", row(r1, "사무용가구")["cells"][0]["v"], K["가구"] * 1000)
check("소모품 11월(D 4,000)", row(r1, "소모품")["cells"][1]["v"], K["소모품_D"] * 1000)
check("소모품 12월(D+1 3,000)", row(r1, "소모품")["cells"][2]["v"], K["소모품_D1"] * 1000)
check("집기 10월(복합기1+PC7)", row(r1, "집기비품")["cells"][0]["v"], (K["복합기"] + K["PC"] * 7) * 1000)
check("집기 11월(+프린터5+파쇄기)", row(r1, "집기비품")["cells"][1]["v"], (K["복합기"] + K["PC"] * 7 + K["프린터"] * K["프린터_대수"] + K["파쇄기"]) * 1000)
check("보안 10월(공사+용역)", row(r1, "보안경비용역")["cells"][0]["v"], (K["무인경비_공사"] + K["무인경비_용역"]) * 1000)
check("용역-기타 11월(전기+소방+청소4명)", row(r1, "용역-기타")["cells"][1]["v"], (K["전기안전"] + K["소방안전"] + K["청소단가"] * K["청소_1개월_4이상"]) * 1000)
check("통신 11월(600+500+내용증명 623×500%)", row(r1, "통신비")["cells"][1]["v"], (K["통신_MH"] + K["우편_기본"] + K["내용증명"] * round(623 * 5)) * 1000)
check("운반비 10월", row(r1, "운반비")["cells"][0]["v"], K["운반비"] * 1000)
check("여비 10월(사전 1,500)", row(r1, "여비교통비")["cells"][0]["v"], K["여비_사전_수도권"] * 1000)
check("여비 11월(오픈 2,000)", row(r1, "여비교통비")["cells"][1]["v"], K["여비_오픈_수도권"] * 1000)
assert row(r1, "발급수수료")["cells"][0]["v"] == 0 and row(r1, "발급수수료")["cells"][1]["v"] == K["발급_기본"] * 1000
print("  산식 칸", verify_formulas(r1), "개 모두 재계산 일치 · 비고(식대):", row(r1, "식대")["note"].replace("\n", " / "))
assert "10월 : 10천원/일×4명×30일" in row(r1, "식대")["note"]
assert len(r1["approvers"]["list"]) == 5 and r1["approvers"]["list"][0]["role"] == "담당" and r1["approvers"]["list"][-1]["role"] == "승인"

# ── 사례 2: 견본주택 운영(오픈 일정 없음, 기간 '26.10~12) + 중도금대출(11월) + 선착순, 직접 고친 칸·비고, 추가 줄, 결재란 3명
d2 = dict(kinds=["run"], fields={"현장명": "운영현장", "시작": [2026, 10], "개월수": 3, "지역": "비수도권", "세대수": 1000,
          "숙소호실": 2, "PC": 5},
          months=[{"events": ["선착순"]}, {"events": ["중도금대출", "선착순"], "청소인원": 1}, {"events": ["선착순"], "label": "(선착순 마감)"}],
          overrides={"소모품": {"1": "=1500000*2"}, "여비교통비": {"0": 777000}}, notes={"운반비": "직접 쓴 비고"},
          extra=[{"item": "지급수수료", "name": "CCTV 설치", "vals": ["=500000*3", 0, ""], "note": "CCTV 3대"}],
          approvers={"list": [{"role": "담당", "pos": "과장", "name": "홍길동"}, {"role": "승인", "pos": "부장"}, {"role": "승인", "pos": "실장"}],
                     "width": 14, "height": 60})
r2 = jichul.calc(d2, K)
print("\n[사례2]", r2["title"], [m["label"] for m in r2["months"]], "합계", f"{r2['total_all']:,.0f}", r2["warnings"])
assert r2["kind"] == "M/H 운영비용" and r2["months"][1]["label"] == "(중도금대출,선착순)" and r2["months"][2]["label"] == "(선착순 마감)"
check("식대 10월(잔여판촉 1,800)", row(r2, "식대")["cells"][0]["v"], K["식대_잔여"] * 1000)
check("식대 11월(잔여 + 은행식대 10×7×20)", row(r2, "식대")["cells"][1]["v"], (K["식대_잔여"] + 10 * 7 * 20) * 1000)
check("발급 11월(30 + 확정일자 600×1000)", row(r2, "발급수수료")["cells"][1]["v"], K["발급_기본"] * 1000 + K["확정일자"] * 1000)
check("통신 11월(600+500+등기 4.5×1000)", row(r2, "통신비")["cells"][1]["v"], (K["통신_MH"] + K["우편_기본"] + K["등기"] * 1000) * 1000)
check("용역-기타 10월(청소 잔여 2명)", row(r2, "용역-기타")["cells"][0]["v"], (K["전기안전"] + K["소방안전"] + K["청소단가"] * K["청소_잔여"]) * 1000)
check("용역-기타 11월(청소 인원 직접 1명)", row(r2, "용역-기타")["cells"][1]["v"], (K["전기안전"] + K["소방안전"] + K["청소단가"] * 1) * 1000)
check("숙박비 10월(운영: 기존 숙소 1,000×2호 월납)", row(r2, "숙박비")["cells"][0]["v"], K["숙소비_월"] * 2 * 1000)
check("숙박비 12월(기존 숙소 계속)", row(r2, "숙박비")["cells"][2]["v"], K["숙소비_월"] * 2 * 1000)
check("소모품 12월(운영 D+2~ 1,000)", row(r2, "소모품")["cells"][2]["v"], K["소모품_D2"] * 1000)
check("집기 10월(복합기1+PC5+프린터5+파쇄기)", row(r2, "집기비품")["cells"][0]["v"], (K["복합기"] + K["PC"] * 5 + K["프린터"] * K["프린터_대수"] + K["파쇄기"]) * 1000)
check("보안 10월(용역비만, 공사비 없음)", row(r2, "보안경비용역")["cells"][0]["v"], K["무인경비_용역"] * 1000)
check("소모품 11월 직접 산식", row(r2, "소모품")["cells"][1]["v"], 3000000)
check("여비 10월 직접 값", row(r2, "여비교통비")["cells"][0]["v"], 777000)
check("여비 11월(운영 기타 비수도권 1,500)", row(r2, "여비교통비")["cells"][1]["v"], K["여비_기타_비수도권"] * 1000)
check("여비 12월(운영 기타 비수도권 1,500)", row(r2, "여비교통비")["cells"][2]["v"], K["여비_기타_비수도권"] * 1000)
assert row(r2, "사무용가구") is None and row(r2, "도서인쇄비") is None and "기존 숙소" in row(r2, "숙박비")["note"]
assert not [w for w in r2["warnings"] if "오픈" in w], r2["warnings"]
assert row(r2, "소모품")["cells"][1]["f"] == "1500000*2" and row(r2, "소모품")["cells"][1]["manual"]
assert row(r2, "운반비") is None or row(r2, "운반비")["note"] == "직접 쓴 비고"
ex = row(r2, "extra:0")
assert ex and ex["item"] == "지급수수료" and ex["total"] == 1500000 and ex["cells"][0]["f"] == "500000*3"
assert r2["months"][1]["식대인원"] == 3 and r2["months"][1]["청소인원"] == 1
assert [a["role"] for a in r2["approvers"]["list"]] == ["담당", "승인", "승인"] and r2["approvers"]["width"] == 14
print("  산식 칸", verify_formulas(r2), "개 일치")

# ── 사례 2 변형: 같은 조건을 「견본주택 오픈」 구분 + 오픈 '26.07 로 — 오픈 일정이 있으면 D+4 는 오픈후 3개월 여비(2,500)
d2b = dict(d2, kinds=["mh"], fields=dict(d2["fields"], **{"오픈": [2026, 7], "유니트": 3}))
r2b = jichul.calc(d2b, K)
check("오픈 구분 여비 11월(D+4 오픈후 3개월 비수도권 2,500)", row(r2b, "여비교통비")["cells"][1]["v"], K["여비_사전_비수도권"] * 1000)
check("오픈 구분 숙박비 10월(일시불 3개월 뒤 기존 숙소 1,000×2호)", row(r2b, "숙박비")["cells"][0]["v"], K["숙소비_월"] * 2 * 1000)
assert r2b["kind"] == "M/H 운영비용"
# 운영 구분은 오픈월 칸에 값이 남아 있어도 무시(오픈 단계 없음·경고 없음), 오픈+운영 둘 다 고르면 오픈 구분으로
r2c = jichul.calc(dict(d2, fields=dict(d2["fields"], **{"오픈": [2026, 11], "사전영업": 1})), K)
check("운영 구분 + 오픈월 남아 있음: 식대 11월은 잔여판촉 1,800 + 은행식대", row(r2c, "식대")["cells"][1]["v"], (K["식대_잔여"] + 10 * 7 * 20) * 1000)
assert not [w for w in r2c["warnings"] if "오픈" in w] and not any("오픈" in m["label"] for m in r2c["months"])
r2d = jichul.calc(dict(d2b, kinds=["mh", "run"]), K)
assert r2d["total_all"] == r2b["total_all"] and r2d["warnings"] == r2b["warnings"]
# 운영 구분에서 M/H 종료월(11월): 12월은 선착순을 골라도 비용 없음, 숙소도 종료월까지
r2e = jichul.calc(dict(d2, fields=dict(d2["fields"], **{"MH종료": [2026, 11]}), overrides={}, extra=[]), K)
print("[사례2e] M/H 종료 11월:", [m["label"] for m in r2e["months"]], "월별", [f"{x:,.0f}" for x in r2e["total"]])
assert r2e["total"][2] == 0 and r2e["months"][2]["stage"] == [] and row(r2e, "숙박비")["cells"][1]["v"] == K["숙소비_월"] * 2 * 1000
# 운영 구분에서 숙소 시작월을 넣으면 그 달에 3개월 일시불(새 숙소)
r2f = jichul.calc(dict(d2, fields=dict(d2["fields"], **{"숙소시작": [2026, 11]}), overrides={}, extra=[]), K)
assert row(r2f, "숙박비")["cells"][0]["v"] == 0 and row(r2f, "숙박비")["cells"][1]["v"] == K["숙소비_월"] * K["숙소_개월"] * 2 * 1000 and row(r2f, "숙박비")["cells"][2]["v"] == 0

# ── 사례 3: 입주(사전점검 '27.02, 입주 '27.03, 기간 '27.01~03), 800세대
d3 = dict(kinds=["movein"], fields={"현장명": "입주현장", "시작": [2027, 1], "개월수": 3, "지역": "수도권", "세대수": 800,
          "사전점검": [2027, 2], "입주": [2027, 3], "지정기간": 2, "입주숙소": 2, "입주PC": 4}, months=[{}, {}, {}])
r3 = jichul.calc(d3, K)
print("\n[사례3]", r3["title"], [m["label"] for m in r3["months"]], "합계", f"{r3['total_all']:,.0f}", r3["warnings"])
assert r3["kind"] == "입주비용" and [m["label"] for m in r3["months"]] == ["", "(사전점검)", "(입주)"]
check("식대 1월(센터 전) 0", row(r3, "식대")["cells"][0]["v"], 0)
check("식대 2월(기타 1,200)", row(r3, "식대")["cells"][1]["v"], K["식대_기타"] * 1000)
check("숙박비 3월(입주 숙소 1,000×3×2호)", row(r3, "숙박비")["cells"][2]["v"], K["숙소비_입주_월"] * K["숙소_개월"] * 2 * 1000)
check("소모품 2월(입주/사전점검 1,000)", row(r3, "소모품")["cells"][1]["v"], K["소모품_입주점검"] * 1000)
check("집기 2월(복합기1 + PC4)", row(r3, "집기비품")["cells"][1]["v"], (K["복합기"] + K["PC"] * 4) * 1000)
check("통신 1월(사전점검 안내 4.5×800)", row(r3, "통신비")["cells"][0]["v"], K["등기"] * 800 * 1000)
check("통신 2월(500+500 + 입주안내 + 어린이집)", row(r3, "통신비")["cells"][1]["v"], (K["통신_입주"] + K["우편_기본"] + K["등기"] * 800 + K["어린이집"] * 800) * 1000)
check("운반비 3월(입주지원센터 이전)", row(r3, "운반비")["cells"][2]["v"], K["운반비"] * 1000)
check("여비 2월(기타 1,000)", row(r3, "여비교통비")["cells"][1]["v"], K["여비_기타_수도권"] * 1000)
print("  산식 칸", verify_formulas(r3), "개 일치")

# ── 사례 4: 분양사무실('27.04~06, 4월 이전), 통합센터
d4 = dict(kinds=["office"], fields={"현장명": "사무실현장", "시작": [2027, 4], "개월수": 3, "사무실시작": [2027, 4], "통합센터": "Y"}, months=[{}, {}, {}])
r4 = jichul.calc(d4, K)
print("\n[사례4]", r4["title"], [m["label"] for m in r4["months"]], "합계", f"{r4['total_all']:,.0f}", r4["warnings"])
assert r4["kind"] == "운영비용" and r4["months"][0]["label"] == "(분양사무실 이전)"
check("소모품(300)", row(r4, "소모품")["cells"][0]["v"], K["사무실_소모품"] * 1000)
check("집기(복합기 300)", row(r4, "집기비품")["cells"][0]["v"], K["사무실_복합기"] * 1000)
check("발급(30)", row(r4, "발급수수료")["cells"][0]["v"], K["사무실_발급"] * 1000)
check("보안(300)", row(r4, "보안경비용역")["cells"][0]["v"], K["사무실_보안"] * 1000)
check("통신(200+50)", row(r4, "통신비")["cells"][0]["v"], (K["사무실_통신"] + K["사무실_우편"]) * 1000)
check("관리비(통합 1,000)", row(r4, "관리비")["cells"][0]["v"], K["사무실_관리비_통합"] * 1000)
check("운반비 4월", row(r4, "운반비")["cells"][0]["v"], K["운반비"] * 1000)
print("  산식 칸", verify_formulas(r4), "개 일치")

# ── 엑셀
outs = {}
for name, d, r in (("open", d1, r1), ("run", d2, r2), ("movein", d3, r3), ("office", d4, r4)):
    xb = jichul_xlsx.build_xlsx(r, d, y)
    outs[name] = xb
    back = jichul_xlsx.read_inputs(xb)
    assert back["year"] == y and back["data"] == d, name
    wb = openpyxl.load_workbook(io.BytesIO(xb))
    ws = wb[jichul_xlsx.SHEET]
    nf = sum(1 for row_ in ws.iter_rows() for c in row_ if isinstance(c.value, str) and c.value.startswith("="))
    nap = len(r["approvers"]["list"])
    assert ws.cell(5, 4 + r["n"]).value == "결\n\n재" and ws.cell(8, 4 + r["n"] + nap).value, name
    assert ws.print_area and ws.page_setup.fitToWidth == 1 and ws.page_setup.fitToHeight == 1 and wb[jichul_xlsx.DATA_SHEET].sheet_state == "veryHidden"
    print(f"엑셀 {name}: {len(xb):,} bytes · 수식 칸 {nf} · 결재 {nap}칸(너비 {ws.column_dimensions[openpyxl.utils.get_column_letter(5 + r['n'])].width}) · "
          f"서명칸 높이 {ws.row_dimensions[6].height} · 인쇄 {ws.print_area} · 되읽기 일치")
    assert ws.row_dimensions[6].height == (r["approvers"]["height"] or 48)

print("\n" + ("실패: " + "; ".join(fails) if fails else "모든 검사 통과"))

# ── Excel 이 있으면 PDF → PNG (눈으로 확인용)
if "--render" in sys.argv:
    import pythoncom, win32com.client, pymupdf      # noqa: E401
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    for name, xb in outs.items():
        p = os.path.join(ROOT, "tests", f"jichul_{name}.xlsx")
        open(p, "wb").write(xb)
        pdf = p[:-5] + ".pdf"
        wb = xl.Workbooks.Open(p, 0, True)
        vals = [wb.Worksheets(1).Range("C12").Value]
        wb.Worksheets(1).ExportAsFixedFormat(0, pdf)
        wb.Close(False)
        doc = pymupdf.open(pdf)
        doc[0].get_pixmap(dpi=100).save(p[:-5] + ".png")
        print("렌더", name, "쪽", len(doc), "C12 =", vals[0], "→", p[:-5] + ".png")
        assert len(doc) == 1, "한 쪽에 들어가야 함"
    xl.Quit()
sys.exit(1 if fails else 0)
