# -*- coding: utf-8 -*-
"""사업관리비 실적계획 엔진 시험(가짜 자료): 읽기 → 계획 → 사내 양식 → 근거자료.

  python tests\\siljeok_test.py                 엔진만(Excel 없이)
  python tests\\siljeok_test.py --excel         만든 두 파일을 이 PC 의 Excel 로 열어 다시 계산한 값·'#'(오류/####)·외부 연결까지 확인
  python tests\\siljeok_test.py --keep <폴더>   만든 파일을 남김

확인하는 것: 전년도 예산(값 없는 수식 파일)·금년도 예산안(예시3) 합계, 값 없는 사본도 같은 결과, 작성자 고르기·이월잔액·전년도 아닌 줄·
겹친 줄 처리, 계정 옮기기/빼기, 양식의 블록 배치·수식·저장 값·서식(틀의 씨앗 행과 같은지), 연도 바꾸기, 여러 현장.
"""
import io
import json
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8")
import warnings
warnings.filterwarnings("ignore")

SP = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(SP)
sys.path.insert(0, os.path.join(ROOT, "tool"))
sys.path.insert(0, SP)
import openpyxl                    # noqa: E402
import siljeok as S                # noqa: E402
import siljeok_samples as G        # noqa: E402

TPL = os.path.join(ROOT, "회사양식", "실적계획", "사업관리비 실적계획 양식.xlsx")
NEXT = os.path.join(ROOT, "예시", G.NEXT_EXAMPLE)
NEXT_TOTAL = 3355445000            # 예시3 합계(tests/api_test.py 기대값과 같음)
fails = []


def check(name, cond, detail=""):
    print(("  ✓ " if cond else "  ✗ ") + name + (f" — {detail}" if detail and not cond else ""))
    if not cond:
        fails.append(name)


def near(a, b):
    return abs(float(a) - float(b)) <= max(0.5, 1e-9 * abs(float(b)))


def strip_cache(data):
    """수식 칸의 저장 값을 지운 사본(웹판에서 내려받아 Excel 로 열어 보지 않은 파일과 같은 상태)"""
    zin = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            b = zin.read(it.filename)
            if it.filename.startswith("xl/worksheets/sheet"):
                s = b.decode("utf-8")
                s = re.sub(r"<c [^>]*>(?:(?!</c>).)*</c>",
                           lambda m: re.sub(r' t="(?:str|e|b)"', "", re.sub(r"<v>[^<]*</v>", "", m.group(0))) if "<f" in m.group(0) else m.group(0),
                           s, flags=re.S)
                b = s.encode("utf-8")
            z.writestr(it.filename, b)
    return out.getvalue()


def excel_check(path, label):
    """Excel(별도 인스턴스)로 열어 다시 계산: 저장해 둔 값과 같은지, '#' 로 보이는 칸, 외부 연결"""
    import time
    import pywintypes
    import win32com.client as win32

    def R(fn, tries=240):
        for _ in range(tries):
            try:
                return fn()
            except pywintypes.com_error as e:
                if e.hresult in (-2147418111, -2147417846):
                    time.sleep(0.5)
                    continue
                raise
        raise RuntimeError("Excel 이 응답하지 않습니다")

    wf, wv = openpyxl.load_workbook(path), openpyxl.load_workbook(path, data_only=True)
    stored = {(ws.title, c.coordinate): wv[ws.title][c.coordinate].value for ws in wf.worksheets for row in ws.iter_rows() for c in row
              if isinstance(c.value, str) and c.value.startswith("=")}
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    diffs, hashes, links, errors = [], [], None, []
    try:
        wb = R(lambda: xl.Workbooks.Open(path, 0, True))
        R(lambda: xl.CalculateFull())
        links = R(lambda: wb.LinkSources(1))
        for i in range(1, R(lambda: wb.Worksheets.Count) + 1):
            ws = R(lambda: wb.Worksheets(i))
            name = R(lambda: ws.Name)
            ur = R(lambda: ws.UsedRange)
            r0, c0, nr, nc = R(lambda: ur.Row), R(lambda: ur.Column), R(lambda: ur.Rows.Count), R(lambda: ur.Columns.Count)
            vals = R(lambda: ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + nr - 1, c0 + nc - 1)).Value2)
            for (sn, addr), sv in stored.items():
                if sn != name:
                    continue
                col, row = openpyxl.utils.cell.coordinate_from_string(addr)
                x = vals[row - r0][openpyxl.utils.column_index_from_string(col) - c0]
                if isinstance(x, int) and x < -2000000000:
                    errors.append((name, addr))
                elif sv is not None and isinstance(x, (int, float)) and not near(x, sv) and not (abs(x) < 5 and abs(x - sv) < 1e-6):
                    diffs.append((name, addr, sv, x))
            if nr * nc <= 40000:
                for rr in range(nr):
                    for cc in range(nc):
                        if isinstance(vals[rr][cc], (int, float)):
                            tx = R(lambda: ws.Cells(r0 + rr, c0 + cc).Text)
                            if isinstance(tx, str) and tx.startswith("#"):
                                hashes.append((name, r0 + rr, c0 + cc, tx))
        R(lambda: wb.Close(False))
    finally:
        try:
            R(lambda: xl.Quit(), tries=40)
        except Exception:      # noqa: BLE001
            pass
    check(f"{label}: Excel 로 열림 · 수식 {len(stored)}칸 다시 계산해도 같은 값", not diffs, str(diffs[:5]))
    check(f"{label}: 수식 오류 없음", not errors, str(errors[:5]))
    check(f"{label}: '#' 로 보이는 칸 없음", not hashes, str(hashes[:5]))
    check(f"{label}: 외부 연결 없음", not links, str(links))
    return stored


def main():
    keep = sys.argv[sys.argv.index("--keep") + 1] if "--keep" in sys.argv else None
    use_excel = "--excel" in sys.argv
    tpl = open(TPL, "rb").read()
    files = G.files()
    names = list(files)
    next_raw = open(NEXT, "rb").read()

    print("■ 읽기")
    prev = S.parse_file(files[names[0]], names[0])
    check("전년도 예산: 예산안으로 알아봄 · 2026년 · 값 없는 수식을 직접 계산", prev["kind"] == "budget" and prev["year"] == 2026 and prev["calc"] == "computed")
    check("전년도 예산 합계", near(prev["total"], G.prev_budget_total()), f"{prev['total']:,.0f} vs {G.prev_budget_total():,.0f}")
    check("전년도 예산 비목별", all(near(sum(ln["amount"] for ln in it["lines"]), G.prev_budget_by_item()[it["item"]]) for it in prev["items"])
          and len(prev["items"]) == len(G.PREV_BUDGET))
    check("전년도 예산 현장명", prev["site"] == G.SITE, prev["site"])
    nxt = S.parse_file(next_raw, G.NEXT_EXAMPLE)
    check("금년도 예산안(예시3): 2027년 · 합계 3,355,445,000 · Excel 저장 값 사용", nxt["year"] == 2027 and near(nxt["total"], NEXT_TOTAL) and nxt["calc"] == "cached",
          f"{nxt['total']:,.0f} {nxt['calc']}")
    check("금년도 예산안: 경고 없음", not nxt["warnings"], str(nxt["warnings"]))
    nxt2 = S.parse_file(strip_cache(next_raw), G.NEXT_EXAMPLE)
    same = json.dumps(nxt["items"], sort_keys=True) == json.dumps(nxt2["items"], sort_keys=True)
    check("금년도 예산안: 저장 값을 지운 사본(웹판 그대로)도 같은 결과", nxt2["calc"] == "computed" and same and near(nxt2["total"], NEXT_TOTAL))
    bodies = [b for it in nxt["items"] for ln in it["lines"] for b in ln["mbodies"] if b]
    check("금년도 예산안: 월별 산식에 칸·시트 참조가 남지 않음", bodies and all(not re.search(r"[A-Z]{1,3}\$?\d|!", re.sub(r"ROUND(UP|DOWN)?|INT|ABS|MOD|MIN|MAX", "", b)) for b in bodies))
    check("금년도 예산안: 산식 값 = 칸 값", all(near(S.eval_expr(b), v) for it in nxt["items"] for ln in it["lines"] for b, v in zip(ln["mbodies"], ln["mvals"]) if b))
    lg = [S.parse_file(files[n], n) for n in names[1:]]
    a_rows, b_rows, carry = G.ledger_rows()
    check("사용내역: 두 파일 줄 수 · 이월잔액 제외", lg[0]["kind"] == "ledger" and lg[0]["n"] == len(a_rows) and lg[1]["n"] == len(b_rows)
          and lg[1]["skipped"]["carry"] == len(carry), f"{lg[0]['n']} {lg[1]['n']} {lg[1]['skipped']}")
    check("사용내역: 귀속코드", list(lg[0]["codes"]) == [G.CODE + "D"])
    check("사용내역: 주민등록번호·거래처코드 열은 읽지 않음", all(len(r) == len(S.LEDGER_FIELDS) for r in lg[0]["rows"]) and "000-00-00000" not in json.dumps(lg[0]))
    try:
        S.parse_file(open(TPL, "rb").read(), "틀.xlsx")
        check("알 수 없는 파일은 오류", False)
    except S.SiljeokError as e:
        check("알 수 없는 파일은 오류", "알 수 없는 파일" in str(e), str(e))

    print("■ 계획(양식에 넣을 값)")
    site = dict(name=G.SITE, code="", prev=prev, next=nxt, ledgers=lg, creators=None, accmap={}, inc={})
    state = json.loads(json.dumps(dict(year=None, sites=[site])))
    p = S.plan(state)
    s0 = p["sites"][0]
    check("기준 연도 2027 · 파일 이름", p["year"] == 2027 and p["form_name"].startswith("2026년_사업관리비_실적계획_양식_예시3"), p["form_name"])
    check("작성자를 고르기 전: 실적 0 · 경고", s0["total"]["f"] == 0 and s0["total"]["g"] == 0 and any("작성자를 고르지" in w for w in s0["warnings"]))
    check("작성자 목록 4명 · 현장코드 제안", {c["name"] for c in s0["creators"]} == set(G.STAFF + G.OTHERS) and s0["ledger"]["suggest_code"] == G.CODE)
    check("겹친 2줄 · 전년도 아닌 1줄 · 이월잔액 3줄 처리", s0["ledger"]["dup"] == 2 and s0["ledger"]["other_year"] == 1 and s0["ledger"]["carry"] == 3, str(s0["ledger"]))
    state["sites"][0]["creators"] = list(G.STAFF)
    state["sites"][0]["code"] = G.CODE
    p = S.plan(state)
    s0 = p["sites"][0]
    exp = G.expected_actual(G.STAFF)
    rows = {r["item"]: r for r in s0["rows"]}
    check("실적(현장 담당자 2명): 비목별 1~8월·9~12월", all(near(rows[k]["f"], v[0]) and near(rows[k]["g"], v[1]) for k, v in exp.items()) and
          near(s0["total"]["f"], sum(v[0] for v in exp.values())) and near(s0["total"]["g"], sum(v[1] for v in exp.values())),
          str({k: (rows.get(k, {}).get("f"), rows.get(k, {}).get("g"), v) for k, v in exp.items()}))
    check("재무·세무 전표는 빠짐(세금과공과 실적 0, 급료및수당 줄 없음)", rows["세금과공과"]["f"] == 0 and "급료및수당" not in rows)
    check("되돌린 전표(재무 대체) 알림 · 그 작성자까지 고르면 상계", any("되돌린 전표" in w and "판매촉진비 41,100,000원" in w for w in s0["warnings"]) and
          near(rows["판매촉진비"]["f"], 41100000) and
          S.plan(dict(state, sites=[dict(state["sites"][0], creators=G.STAFF + [G.OTHERS[0]])]))["sites"][0]["rows"][[r["item"] for r in s0["rows"]].index("판매촉진비")]["f"] == 0)
    check("입주관리비 계정은 관리비로", near(rows["관리비"]["f"], exp["관리비"][0]) and rows["관리비"]["accts"] == ["판매관리비(분양관리비)_입주관리비_D"])
    check("지급수수료 = 분양경비 + 기타 계정(산식)", rows["지급수수료"]["f_f"] and near(S.eval_expr(rows["지급수수료"]["f_f"]), rows["지급수수료"]["f"]), str(rows["지급수수료"]["f_f"]))
    check("전년도 예산(D) = 세목 금액 합 산식", near(S.eval_expr(rows["복리후생비"]["d_f"]), G.prev_budget_by_item()["복리후생비"]) and
          near(s0["total"]["d"], G.prev_budget_total()))
    check("금년도 예산(K) 합계 = 예산안 합계", near(s0["total"]["k"], NEXT_TOTAL), f"{s0['total']['k']:,.0f}")
    check("비목 순서는 양식 순서", [r["item"] for r in s0["rows"]] == [x for x in S.ORDER if x in rows] + [x for x in rows if x not in S.ORDER])
    check("제목", s0["title"] == f"[현장명 : {G.SITE} ({G.CODE})]", s0["title"])
    # 계정 옮기기·빼기, 증액, 작성자 추가
    st2 = json.loads(json.dumps(state))
    st2["sites"][0]["accmap"] = {"판매관리비(지급수수료)_기타_D": "-", "판매관리비(소모품비)_분양경비_D": "복리후생비"}
    st2["sites"][0]["inc"] = {"통신비": 1000000}
    st2["sites"][0]["creators"] = G.STAFF + [G.OTHERS[1]]
    st2["sites"][0]["note_diff"], st2["sites"][0]["note_change"] = "예시 분석", "예시 증감요인"
    r2 = {r["item"]: r for r in S.plan(st2)["sites"][0]["rows"]}
    check("계정 빼기: 지급수수료에서 기타 계정 제외", near(r2["지급수수료"]["f"], exp["지급수수료"][0] - 30000) and not r2["지급수수료"]["f_f"])
    check("계정 옮기기: 소모품비 → 복리후생비", near(r2["복리후생비"]["f"], exp["복리후생비"][0] + exp["소모품비"][0]) and r2["소모품비"]["f"] == 0)
    check("증액금액: 총예산 = 최초편성 + 증액", near(r2["통신비"]["c"], r2["통신비"]["d"] + 1000000))
    dup = S.plan(dict(st2, sites=[st2["sites"][0], st2["sites"][0]]))
    check("같은 현장명 두 번이면 오류", any("현장명이 같습니다" in e for e in dup["errors"]))
    # 내역을 시트 둘로 나눠 둔 파일(앞에 요약 시트): 머리글이 맞는 시트를 모두 읽는다
    w2 = openpyxl.Workbook()
    w2.active.title = "요약"
    w2.active.append(["G/L 계정", "G/L 계정명", "1-8월", "9-11월", "합계"])
    for title, part in (("내역_담당자", a_rows[:10]), ("내역_그외", a_rows[10:25])):
        src = openpyxl.load_workbook(io.BytesIO(G.ledger_xlsx(part))).active
        dst = w2.create_sheet(title)
        dst.append(["기간"] + [c.value for c in src[1]])
        for row in src.iter_rows(min_row=2, values_only=True):
            dst.append(["1-8월"] + list(row))
    bio = io.BytesIO()
    w2.save(bio)
    two = S.parse_file(bio.getvalue(), "두 시트.xlsx")
    check("사용내역이 여러 시트로 나뉜 파일도 모두 읽음", two["kind"] == "ledger" and two["n"] == 25, str(two["n"]))
    check("작성자 추가(세무): 세금과공과 실적에 재산세", near(r2["세금과공과"]["f"], 1629460) and near(r2["세금과공과"]["g"], 2351760))
    check("분석 문구: 첫 줄 + 상동", r2[list(r2)[0]]["note_j"] == " 예시 분석" and r2[list(r2)[1]]["note_j"] == " 상동" and r2[list(r2)[1]]["note_m"] == " 상동")

    print("■ 사내 양식")
    form = S.build_form(tpl, st2)
    wf = openpyxl.load_workbook(io.BytesIO(form))
    wv = openpyxl.load_workbook(io.BytesIO(form), data_only=True)
    wt = openpyxl.load_workbook(io.BytesIO(tpl))
    ws, wd, wt0 = wf.worksheets[0], wv.worksheets[0], wt.worksheets[0]
    n = len(r2)
    total = 3 + 2 * n + 1
    check("시트 하나 · 이름 · 외부 연결 없음", wf.sheetnames == [S.MAIN] and not getattr(wf, "_external_links", []))
    check("블록 배치: 제목·머리글·비목 줄+소 계·합 계", ws["A1"].value == f"[현장명 : {G.SITE} ({G.CODE})]" and ws["A2"].value == "비목" and
          ws["A5"].value == "소 계" and ws[f"A{total}"].value == "합 계" and ws.max_row == total, f"{ws.max_row} {total}")
    check("인쇄 영역 · 범위", ws.print_area in (f"'{S.MAIN}'!$A$1:$AA${total}", f"$A$1:$AA${total}"), str(ws.print_area))
    items = list(r2)
    ok_vals = ok_form = True
    for i, it in enumerate(items):
        dr, sr = 4 + 2 * i, 5 + 2 * i
        row = r2[it]
        ok_vals &= ws[f"A{dr}"].value == it and ws[f"B{dr}"].value == "분양경비" and near(wd[f"D{dr}"].value or 0, row["d"]) and near(wd[f"K{dr}"].value, row["k"]) \
            and near(wd[f"F{dr}"].value or 0, row["f"]) and near(wd[f"G{dr}"].value or 0, row["g"]) and near(wd[f"C{sr}"].value, row["c"])
        ok_form &= ws[f"C{dr}"].value == f"=D{dr}+E{dr}" and ws[f"H{dr}"].value == f"=F{dr}+G{dr}" and ws[f"K{dr}"].value == f"=SUM(O{dr}:Z{dr})" \
            and ws[f"I{dr}"].value == f"=IF(C{dr}=0,0,H{dr}/C{dr})" and ws[f"L{dr}"].value == f"=IF(OR(H{dr}=0,K{dr}=0),0,K{dr}/H{dr})" \
            and ws[f"D{sr}"].value == f"=SUM(D{dr}:D{dr})" and ws[f"Z{sr}"].value == f"=SUM(Z{dr}:Z{dr})"
    check("비목 줄 값(A·B·D·F·G·K)과 소 계", ok_vals)
    check("양식 수식(C·H·I·K·L, 소 계 SUM)", ok_form)
    subs = ",".join(f"D{5 + 2 * i}" for i in range(n))
    check("합 계 수식이 모든 소 계를 더함(D~G·월별)", ws[f"D{total}"].value == f"=SUM({subs})" and ws[f"E{total}"].value == f"=SUM({subs.replace('D', 'E')})"
          and ws[f"Z{total}"].value == f"=SUM({subs.replace('D', 'Z')})")
    n_f = sum(1 for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("="))
    n_v = sum(1 for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=") and wd[c.coordinate].value is not None)
    check("수식 칸마다 저장 값이 들어 있음(Excel 로 열지 않아도 값이 읽힘)", n_f > 0 and n_f == n_v, f"{n_v}/{n_f}")
    check("합계 값", near(wd[f"K{total}"].value, NEXT_TOTAL) and near(wd[f"D{total}"].value, G.prev_budget_total()))
    check("월별 칸은 산식", any(isinstance(ws.cell(4 + 2 * i, c).value, str) and ws.cell(4 + 2 * i, c).value.startswith("=") for i in range(n) for c in range(15, 27)))
    check("근거 칸: 예산안 산출근거에 ' · '", any(str(ws[f"N{4 + 2 * i}"].value or "").startswith(" · ") for i in range(n)))
    check("분석 문구 칸(J·M)", ws["J4"].value == " 예시 분석" and ws["M6"].value == " 상동")
    check("증액금액 칸", ws[f"E{4 + 2 * items.index('통신비')}"].value == 1000000)

    def sty(c):
        f, fl, b, a = c.font, c.fill, c.border, c.alignment
        col = lambda x: (x.rgb if x is not None and x.type == "rgb" else None)
        return (f.name, f.sz, bool(f.b), col(f.color), fl.fill_type, col(fl.fgColor) if fl.fill_type else None,
                tuple((s.style if s is not None else None) for s in (b.left, b.right, b.top, b.bottom)), a.horizontal, a.vertical, c.number_format)

    def same_style(r, seed, skip_shrink=True):
        for c in range(1, 28):
            a, b = ws.cell(r, c), wt0.cell(seed, c)
            if sty(a) != sty(b):
                return f"{a.coordinate}: {sty(a)} / {sty(b)}"
            if not a.alignment.shrink_to_fit and bool(a.alignment.wrap_text) != bool(b.alignment.wrap_text):
                return f"{a.coordinate}: 줄바꿈 다름"
        return ""

    bad = [same_style(1, 1), same_style(2, 2), same_style(3, 3), same_style(total, 6)] + [same_style(4 + 2 * i, 4) for i in range(n)] + \
          [same_style(5 + 2 * i, 5) for i in range(n)]
    check("서식이 틀의 씨앗 행과 같음(글꼴·채움·테두리·정렬·표시 형식)", not any(bad), str([b for b in bad if b][:3]))
    check("열 너비가 틀과 같음", {k: v.width for k, v in ws.column_dimensions.items()} == {k: v.width for k, v in wt0.column_dimensions.items()})
    mg = {str(m) for m in ws.merged_cells.ranges}
    check("병합(제목·머리글·소 계·합 계)", {"A1:N1", "O1:AA1", "A2:A3", "C2:E2", "O2:Z2", "A5:B5", f"A{total}:B{total}"} <= mg and len(mg) == 14 + n + 1, str(len(mg)))
    check("머리글 연도 26·27", ws["C2"].value == "26년 예산" and ws["K2"].value == "27년 예산" and ws["O2"].value == "27년 월별 예산")

    print("■ 연도 바꾸기 · 여러 현장 · 신규 현장")
    prev28 = S.parse_file(G.prev_budget_xlsx(year=2027), "prev27.xlsx")
    st3 = dict(year=2028, sites=[dict(name="가 현장", code="P0001", prev=prev28, next=None, ledgers=[], creators=None),
                                 dict(name="나 현장(신규)", code="", prev=None, next=dict(nxt, year=2028), ledgers=[], creators=None)])
    p3 = S.plan(st3)
    f3 = openpyxl.load_workbook(io.BytesIO(S.build_form(tpl, st3))).worksheets[0]
    n1 = len(p3["sites"][0]["rows"])
    t2 = 3 + 2 * n1 + 1 + 2
    check("2028년 기준: 머리글 27·28", f3["C2"].value == "27년 예산" and f3["F2"].value == "27년 실적" and f3["K2"].value == "28년 예산" and
          f3["J2"].value == "27년 예산대비 실적차이 분석(상세)" and f3["O2"].value == "28년 월별 예산")
    check("둘째 현장 블록이 빈 행 하나 뒤에", f3[f"A{t2 - 1}"].value is None and str(f3[f"A{t2}"].value).startswith("[현장명 : 나 현장(신규) (P    )]") and
          f3[f"A{t2 + 1}"].value == "비목" and f3[f"C{t2 + 1}"].value == "27년 예산", str(f3[f"A{t2}"].value))
    check("사용내역 없는 기존 현장: 실적 칸 빈칸 + 경고", f3["F4"].value is None and f3["G4"].value is None and any("비워" in w for w in p3["sites"][0]["warnings"]))
    check("신규 현장: 전년도 예산·실적 0", f3[f"D{t2 + 3}"].value == 0 and f3[f"F{t2 + 3}"].value == 0 and f3[f"G{t2 + 3}"].value == 0)
    check("현장코드 없으면 경고", any("현장코드" in w for w in p3["sites"][1]["warnings"]))
    try:
        S.build_form(tpl, dict(sites=[dict(name="", prev=prev)]))
        check("현장명 없으면 오류", False)
    except S.SiljeokError as e:
        check("현장명 없으면 오류", "현장명" in str(e), str(e))

    print("■ 근거자료")
    ev = S.build_evidence(st2)
    we = openpyxl.load_workbook(io.BytesIO(ev))
    check("시트 4개", we.sheetnames == ["비목별 집계", "계정별 집계", "반영 내역", "제외 내역"])
    inc_rows = [r for r in we["반영 내역"].iter_rows(min_row=2, values_only=True) if r[0]]
    exc_rows = [r for r in we["제외 내역"].iter_rows(min_row=2, values_only=True) if r[0]]
    exp3 = G.expected_actual(G.STAFF + [G.OTHERS[1]])
    tot_exp = sum(sum(v) for v in exp3.values()) - 60000          # 지급수수료_기타 계정(5월·10월 3만원씩)을 뺌
    check("반영 내역 합계 = 양식 실적 합계", near(sum(r[10] for r in inc_rows), tot_exp), f"{sum(r[10] for r in inc_rows):,.0f} vs {tot_exp:,.0f}")
    check("반영 내역 기간 표기(~ 없음)", {r[1] for r in inc_rows} == set(S.PERIODS))
    reasons = {r[1] for r in exc_rows}
    check("제외 내역: 사유 4가지", reasons == {"고르지 않은 작성자", "뺀 계정", "다른 파일과 겹침", f"전기일자가 {G.PREV_YEAR}년이 아님"}, str(reasons))
    check("반영 + 제외 = 사용내역 전체 줄", len(inc_rows) + len(exc_rows) == len(a_rows) + len(b_rows), f"{len(inc_rows)}+{len(exc_rows)} vs {len(a_rows) + len(b_rows)}")
    check("집계는 반영 내역을 더하는 수식", str(we["비목별 집계"]["C5"].value).startswith("=SUMPRODUCT(") and "주민" not in json.dumps([list(map(str, r)) for r in inc_rows], ensure_ascii=False))
    try:
        S.build_evidence(st3)
        check("사용내역이 없으면 근거자료 오류", False)
    except S.SiljeokError:
        check("사용내역이 없으면 근거자료 오류", True)

    if keep or use_excel:
        out = keep or os.path.join(os.environ.get("TEMP", SP), "siljeok_test")
        os.makedirs(out, exist_ok=True)
        fp, ep = os.path.join(out, "siljeok_form.xlsx"), os.path.join(out, "siljeok_evidence.xlsx")
        open(fp, "wb").write(form)
        open(ep, "wb").write(ev)
        print("  저장:", out)
        if use_excel:
            print("■ Excel 로 확인")
            excel_check(fp, "양식")
            excel_check(ep, "근거자료")
            import win32com.client as win32          # 근거자료: 다시 계산한 값으로 차이 칸이 0 인지
            xl = win32.DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            try:
                wb = xl.Workbooks.Open(ep, 0, True)
                ws1 = wb.Worksheets("비목별 집계")
                r, diffs, n_rows, total_c = 5, [], 0, 0.0
                while r < 200:
                    a, h = ws1.Cells(r, 1).Value, ws1.Cells(r, 8).Value2
                    if a == "작성 기준":
                        break
                    if a and isinstance(h, (int, float)):
                        n_rows += 1
                        if abs(h) > 0.5:
                            diffs.append((r, h))
                        if str(a).endswith(" 계"):
                            total_c = (ws1.Cells(r, 3).Value2 or 0) + (ws1.Cells(r, 4).Value2 or 0)
                    r += 1
                wb.Close(False)
            finally:
                xl.Quit()
            check(f"근거자료: 차이(집계 − 양식) 칸 {n_rows}개가 모두 0", n_rows > 3 and not diffs, str(diffs))
            check("근거자료: 집계 합계 = 반영 내역 합계", near(total_c, tot_exp), f"{total_c:,.0f} vs {tot_exp:,.0f}")

    print()
    print("실패 %d개: %s" % (len(fails), ", ".join(fails)) if fails else "모두 통과")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
