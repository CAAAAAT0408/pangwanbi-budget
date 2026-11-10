# -*- coding: utf-8 -*-
"""사업관리비 실적계획 자동작성 — 전년도 예산 + 금년도 예산안 + 전년도 사용내역 → 사내 업로드 양식(본사 서버용) + 실적 근거자료.

  parse_file(data, name) → dict          올린 엑셀을 알아보고 읽는다: 판관비 예산안(「NN년 예산」 시트) 또는 SAP 사용내역(G/L 계정명·차액·생성자명)
  plan(state) → dict                     현장별로 양식에 들어갈 값(비목 줄·합계·점검 메시지) — 화면 미리보기와 엑셀이 같은 계산을 쓴다
  build_form(template, state) → bytes    사내 양식 틀(회사양식/실적계획)을 현장·비목 수만큼 복제해 채운 xlsx
  build_evidence(state) → bytes          실적 칸(F·G)의 근거자료 xlsx (비목별·계정별 집계 + 반영/제외 전표 내역)

양식 한 현장 블록:  제목 [현장명 : 이름 (P코드)] · 머리글 2줄 · 비목마다 [자료 줄 + 소 계] · 합 계      (금액 단위 원)
  A 비목 · B 세목 · C 총예산(=D+E) · D 최초편성 · E 증액 · F 1~8월 실적 · G 9~12월 실적 · H 연간전망(=F+G) · I 예산比
  J 실적차이 분석 · K 금년도 예산(=SUM(O:Z)) · L 실적比 · M 증감요인 · N 예산 편성 근거 · O~Z 월별 예산 · AA 비고

규칙(사용자 결정, 2026-11):
  - 양식 틀(열·머리글·서식)은 그대로, 자료 칸만 채운다. 비목마다 한 줄, 세목은 모두 '분양경비'.
  - D = 전년도 예산안의 세목별 금액을 더한 수식(=a+b+c), K·월별 = 금년도 예산안의 월별 산식(다른 칸·시트 참조는 숫자로 풀어 씀),
    N = 금년도 예산안 산출근거(0원 줄 제외).
  - 실적(F·G) = 사용내역 가운데 고른 작성자(생성자명)의 전표만, 전기일자 1~8월 / 9~12월, 계정 이름 괄호 속 비목으로 모음
    ('분양관리비' 는 관리비). 이월잔액 줄·전년도가 아닌 줄·두 파일에 겹친 줄은 뺀다.
state(화면 → 엔진, JSON):
  {"year": 2027, "sites": [{"name", "code", "prev": <예산안>, "next": <예산안>, "ledgers": [<사용내역>…],
                            "creators": [이름…] | None, "accmap": {계정명: 비목 | "-"}, "inc": {비목: 증액금액},
                            "note_diff": "", "note_change": ""}]}
"""
import datetime
import io
import math
import re
import zipfile

import xlcalc

MAIN = "사업관리비 실적계획"
SEMOK = "분양경비"
ORDER = ["복리후생비", "통신비", "수도광열비", "세금과공과", "지급임차료", "보험료", "광고선전비", "판매촉진비", "운반비",
         "관리비", "도서인쇄비", "소모품비", "지급수수료", "여비교통비"]          # 양식의 비목 순서(G/L 계정 순)
ITEM_ALIAS = {"분양관리비": "관리비"}                                              # 계정 괄호 속 이름 → 양식 비목
EXCLUDE = "-"                                                                      # accmap 값: 이 계정은 실적에서 뺌
SPLIT_MONTH = 8                                                                    # F = 1~8월, G = 9~12월
PERIODS = ("1-8월", "9-12월")                                                      # 근거자료의 기간 표기('~' 는 엑셀 조건식에서 특수문자라 쓰지 않음)
LEDGER_FIELDS = ["date", "acct", "acct_name", "memo", "debit", "credit", "amt", "vendor", "doc", "doc_type", "creator"]
NO_CREATOR = "(작성자 없음)"
MONTH_COLS = ["O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
SEED = dict(title=1, h1=2, h2=3, data=4, sub=5, total=6)                           # 틀의 씨앗 행


class SiljeokError(Exception):
    """사용자에게 그대로 보여 줄 오류"""


def squash(v):
    return re.sub(r"\s+", "", str(v if v is not None else ""))


def num_text(v):
    """수식에 넣을 숫자 글자: 정수면 정수, 아니면 유효숫자 15자리"""
    f = float(v)
    if f == int(f) and abs(f) < 1e15:
        return str(int(f))
    return format(f, ".15g")


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _close(a, b):
    return abs(float(a) - float(b)) <= max(0.01, 1e-9 * max(abs(float(a)), abs(float(b))))


# ══════════════════════════════════════════════════════════════════════════════
# 산식 다듬기 — 숫자·사칙연산·ROUND 류만 든 식을 읽어 0 인 항을 지우고 군더더기 괄호를 벗긴다
# ══════════════════════════════════════════════════════════════════════════════
_TOK = re.compile(r"\s*(?:(?P<num>(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)(?P<pct>%?)|(?P<fn>[A-Za-z_][A-Za-z0-9_.]*)\s*\(|(?P<op>[-+*/^(),]))")
_FUNCS = {"ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "ABS", "MOD", "MIN", "MAX"}


class _Parser:
    def __init__(self, text):
        self.toks = []
        i, text = 0, text.strip()
        while i < len(text):
            m = _TOK.match(text, i)
            if not m or m.end() == i:
                raise ValueError("산식을 읽지 못했습니다: " + text[i:i + 20])
            if m.group("num") is not None:
                self.toks.append(("num", m.group("num") + (m.group("pct") or "")))
            elif m.group("fn") is not None:
                self.toks.append(("fn", m.group("fn").upper()))
            else:
                self.toks.append(("op", m.group("op")))
            i = m.end()
        self.i = 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self):
        t = self.peek()
        self.i += 1
        return t

    def parse(self):
        n = self.add()
        if self.i != len(self.toks):
            raise ValueError("산식 끝이 맞지 않습니다")
        return n

    def add(self):
        terms = [("+", self.mul())]
        while self.peek() in (("op", "+"), ("op", "-")):
            op = self.take()[1]
            terms.append((op, self.mul()))
        return terms[0][1] if len(terms) == 1 else ("add", terms)

    def mul(self):
        fs = [("*", self.unary())]
        while self.peek() in (("op", "*"), ("op", "/")):
            op = self.take()[1]
            fs.append((op, self.unary()))
        return fs[0][1] if len(fs) == 1 else ("mul", fs)

    def unary(self):
        if self.peek() == ("op", "-"):
            self.take()
            return ("neg", self.unary())
        if self.peek() == ("op", "+"):
            self.take()
            return self.unary()
        return self.power()

    def power(self):
        a = self.atom()
        if self.peek() == ("op", "^"):
            self.take()
            return ("pow", a, self.unary())
        return a

    def atom(self):
        k, v = self.take()
        if k == "num":
            return ("n", v)
        if k == "fn":
            if v not in _FUNCS:
                raise ValueError("지원하지 않는 함수: " + v)
            args = []
            if self.peek() != ("op", ")"):
                args.append(self.add())
                while self.peek() == ("op", ","):
                    self.take()
                    args.append(self.add())
            if self.take() != ("op", ")"):
                raise ValueError("괄호가 맞지 않습니다")
            return ("fn", v, args)
        if (k, v) == ("op", "("):
            n = self.add()
            if self.take() != ("op", ")"):
                raise ValueError("괄호가 맞지 않습니다")
            return n
        raise ValueError("산식을 읽지 못했습니다")


def _ev(n):
    k = n[0]
    if k == "n":
        t = n[1]
        return float(t[:-1]) / 100 if t.endswith("%") else float(t)
    if k == "neg":
        return -_ev(n[1])
    if k == "add":
        x = 0.0
        for op, c in n[1]:
            x = x + _ev(c) if op == "+" else x - _ev(c)
        return x
    if k == "mul":
        x = 1.0
        for op, c in n[1]:
            x = x * _ev(c) if op == "*" else x / _ev(c)
        return x
    if k == "pow":
        return _ev(n[1]) ** _ev(n[2])
    name, a = n[1], [_ev(x) for x in n[2]]
    if name == "ROUND":
        return xlcalc.xround(a[0], a[1] if len(a) > 1 else 0)
    if name == "ROUNDUP":
        return xlcalc.xround(a[0], a[1] if len(a) > 1 else 0, xlcalc.ROUND_UP)
    if name == "ROUNDDOWN":
        return xlcalc.xround(a[0], a[1] if len(a) > 1 else 0, xlcalc.ROUND_DOWN)
    if name == "INT":
        return float(xlcalc._dec(a[0]).to_integral_value(rounding=xlcalc.ROUND_FLOOR))
    if name == "ABS":
        return abs(a[0])
    if name == "MOD":
        return a[0] - a[1] * math.floor(a[0] / a[1])
    return min(a) if name == "MIN" else max(a)


def _simplify(n):
    k = n[0]
    if k == "n":
        return n
    if k == "neg":
        c = _simplify(n[1])
        return ("n", "0") if c == ("n", "0") else ("neg", c)
    if k == "pow":
        return ("pow", _simplify(n[1]), _simplify(n[2]))
    if k == "fn":
        return ("fn", n[1], [_simplify(a) for a in n[2]])
    if k == "add":
        terms = []
        for op, c in n[1]:
            c = _simplify(c)
            if c[0] == "add":                                   # 괄호 속 덧셈은 밖으로 편다
                for op2, c2 in c[1]:
                    terms.append((op2 if op == "+" else ("-" if op2 == "+" else "+"), c2))
            elif c[0] == "neg":
                terms.append(("-" if op == "+" else "+", c[1]))
            else:
                terms.append((op, c))
        kept = []
        for op, c in terms:
            try:
                zero = _ev(c) == 0
            except Exception:      # noqa: BLE001
                zero = False
            if not zero:
                kept.append((op, c))
        if not kept:
            return ("n", "0")
        kept.sort(key=lambda t: t[0] != "+")                    # 빼는 항은 뒤로(맨 앞이 '-' 로 시작하지 않게) — 순서는 유지
        if len(kept) == 1:
            return kept[0][1] if kept[0][0] == "+" else ("neg", kept[0][1])
        if kept[0][0] == "-":
            return ("add", [("+", ("neg", kept[0][1]))] + kept[1:])
        return ("add", kept)
    fs = []
    for op, c in n[1]:
        c = _simplify(c)
        if c[0] == "mul" and op == "*":                         # 곱셈 속 곱셈은 편다: a*(b*c/d) → a*b*c/d
            fs.extend(c[1])
        else:
            fs.append((op, c))
    return fs[0][1] if len(fs) == 1 else ("mul", fs)


def _prec(n):
    return {"add": 1, "mul": 2, "neg": 3, "pow": 4}.get(n[0], 9)


def _show(n, need=0):
    """식을 글자로. need = 이 자리에 괄호 없이 올 수 있는 가장 낮은 우선순위(덧셈 1 · 곱셈 2 · 부호 3 · 거듭제곱 4 · 숫자/함수 9)"""
    k = n[0]
    if k == "n":
        return n[1]
    if k == "fn":
        return n[1] + "(" + ",".join(_show(a) for a in n[2]) + ")"
    if k == "neg":
        s = "-" + _show(n[1], 9)                    # Excel 은 -2^2 를 (-2)^2 로 읽으므로 부호 뒤는 늘 괄호로 묶는다
    elif k == "pow":
        s = _show(n[1], 9) + "^" + _show(n[2], 9)
    elif k == "add":
        s = ""
        for i, (op, c) in enumerate(n[1]):
            t = _show(c, 2 if op == "-" else 1)     # 빼는 항이 덧셈이면 괄호
            s += t if (i == 0 and op == "+") else op + t
    else:
        s = ""
        for i, (op, c) in enumerate(n[1]):
            t = _show(c, 3 if op == "/" else 2)     # 나누는 수가 곱셈·덧셈이면 괄호
            s += t if i == 0 else op + t
    return "(" + s + ")" if _prec(n) < need else s


def eval_expr(text):
    """숫자 산식의 값(못 읽으면 ValueError)"""
    try:
        return _ev(_Parser(text).parse())
    except ZeroDivisionError:
        raise ValueError("0 으로 나눔")


def tidy(text):
    """산식 글자를 다듬는다: 0 인 항 삭제, 군더더기 괄호 삭제. 값이 달라지거나 못 읽으면 원래 글자 그대로."""
    try:
        ast = _Parser(text).parse()
        v0 = _ev(ast)
        out = _show(_simplify(ast))
        if _close(eval_expr(out), v0):
            return out
    except Exception:              # noqa: BLE001
        pass
    return text


def is_plain_number(text):
    return re.fullmatch(r"\d+(?:\.\d+)?", text or "") is not None


# ══════════════════════════════════════════════════════════════════════════════
# 1. 판관비 예산안(회사 「사업비 예산편성 양식」)의 「NN년 예산」 시트 읽기
# ══════════════════════════════════════════════════════════════════════════════
RE_BUDGET_SHEET = re.compile(r"^\s*(\d{4}|\d{2})\s*년\s*예산\s*$")


def _sheet_names(data):
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        wbx = z.read("xl/workbook.xml").decode("utf-8", "replace")
    except Exception:
        raise SiljeokError("엑셀(.xlsx) 파일이 아닙니다. .xls 파일이면 Excel 에서 .xlsx 로 저장한 뒤 올려 주세요.")
    out = []
    for m in re.finditer(r"<sheet\b[^>]*>", wbx):
        nm = re.search(r'\bname="([^"]*)"', m.group(0))
        if nm:
            out.append(nm.group(1).replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'"))
    return out


def _budget_sheet(names):
    best = None
    for n in names:
        m = RE_BUDGET_SHEET.match(n)
        if m:
            y = int(m.group(1))
            y = y if y >= 2000 else 2000 + y
            if best is None or y > best[0]:
                best = (y, n)
    return best


def basis_lines(lines):
    """예산안 산출근거 → 양식 '예산 편성 근거' 칸 글: 줄마다 ' · ', 앞줄에 딸린 '- …' 줄은 들여쓰기 유지"""
    out = []
    for ln in lines:
        if not ln.get("basis"):
            continue
        for raw in str(ln["basis"]).replace("\r", "").split("\n"):
            s = raw.strip()
            if not s:
                continue
            if s.startswith("-"):
                lead = len(raw) - len(raw.lstrip(" "))
                out.append(" " * (lead + 2) + re.sub(r" {2,}", " ", s))
            else:
                out.append(" · " + re.sub(r" {2,}", " ", s))
    return "\n".join(out)


def parse_budget(data, name=""):
    """예산안 xlsx(bytes) → {"kind":"budget","year","site","total","items":[{"item","lines":[{"semok","amount","basis","mvals","mbodies"}]}],…}
    값은 Excel 이 저장해 둔 값을 쓰고, 없으면(웹판에서 내려받아 Excel 로 열어 보지 않은 파일) 수식을 직접 계산한다."""
    import openpyxl
    names = _sheet_names(data)
    found = _budget_sheet(names)
    if not found:
        raise SiljeokError("「NN년 예산」 시트가 없습니다 — 판관비 예산안(사업비 예산편성) 파일이 맞는지 확인해 주세요.")
    year, sheet = found
    try:
        model = xlcalc.Model(data)
    except Exception as e:         # noqa: BLE001
        raise SiljeokError(f"예산안 파일을 읽지 못했습니다: {e}")
    model.INLINE_SHEETS = tuple(model.sheetnames)
    model.INLINE_DEPTH = 3                      # 세대수×단가 수준까지만 푼다(더 깊이 풀면 ROUNDUP(…) 같은 세대수 계산식까지 딸려 온다)
    raw = model.cells.get(sheet, {})
    last_row = model.maxrow.get(sheet, 0)
    cached = {}
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    try:
        ws = wb[sheet]
        for r, row in enumerate(ws.iter_rows(min_row=1, max_row=last_row or None, max_col=24, values_only=True), 1):
            for c, v in enumerate(row, 1):
                if v is not None:
                    cached[(r, c)] = v
    finally:
        wb.close()
    stat = dict(cached=0, computed=0)

    def is_formula(r, c):
        x = raw.get((r, c))
        return isinstance(x, tuple) and bool(x) and x[0] == "F"

    def val(r, c):
        v = cached.get((r, c))
        if v is None and is_formula(r, c):
            v = model.cell(sheet, r, c)
            if isinstance(v, xlcalc.XLError):
                raise SiljeokError(f"「{sheet}」 {xlcalc.num2col(c)}{r} 칸의 수식을 계산하지 못했습니다({v}). "
                                   "이 파일을 Excel 에서 한 번 열어 저장한 뒤 다시 올려 주세요.")
            stat["computed"] += 1
        elif is_formula(r, c):
            stat["cached"] += 1
        return v

    def num(r, c):
        v = val(r, c)
        return float(v) if _is_num(v) else 0.0

    def text(r, c):
        v = val(r, c)
        return v if isinstance(v, str) else ("" if v is None else str(v))

    # 머리글 줄: '항 목' 과 '세 목' 이 있는 줄
    hr = None
    for r in range(1, min(last_row, 20) + 1):
        cells = {c: squash(cached.get((r, c)) if not is_formula(r, c) else "") for c in range(1, 23)}
        if "세목" in cells.values() and "항목" in cells.values():
            hr = r
            head = cells
            break
    if hr is None:
        raise SiljeokError(f"「{sheet}」 시트에서 머리글(항목·세목) 줄을 찾지 못했습니다.")
    inv = {}
    for c, t in head.items():
        inv.setdefault(t, c)
    c_item, c_semok = inv["항목"], inv["세목"]
    yy = f"{year % 100:02d}"
    c_amt = inv.get(f"{yy}년예산") or next((c for c, t in head.items() if t.endswith("년예산") and c > c_semok), None) or 6
    c_basis = next((c for c, t in head.items() if "산출근거" in t), None) or 7
    c_months = [inv.get(f"{m}월") for m in range(1, 13)]
    if any(c is None for c in c_months):
        c_months = list(range(10, 22))
    site = ""
    for r in (hr - 1, 2, 1):
        m = re.match(r"^\s*\[\s*(.+?)\s*\]", str(cached.get((r, 1)) or ""))
        if m:
            site = m.group(1).strip()
            break

    items, order, warnings = {}, [], []
    cur, semok, total, sub_sum, group_sum = None, None, None, 0.0, {}
    for r in range(hr + 1, last_row + 1):
        a = squash(text(r, 1))
        if a.startswith("합계"):
            total = num(r, c_amt)
            break
        b = squash(text(r, c_item))
        d_raw = text(r, c_semok).strip()
        d = squash(d_raw)
        if d == "소계" or b == "소계":              # 소 계 줄: 위 묶음의 합(이 줄의 항목 칸 글자가 틀려 있어도 묶음을 바꾸지 않는다)
            if cur is not None:
                group_sum[cur] = num(r, c_amt)
            continue
        if b:
            cur, semok = b, None
            if cur not in items:
                items[cur] = []
                order.append(cur)
        if cur is None:
            continue
        if d:
            semok = d_raw
        amount = num(r, c_amt)
        mvals = [num(r, c) for c in c_months]
        if not amount and not any(mvals):
            continue
        basis = text(r, c_basis)
        mbodies = []
        for m, c in enumerate(c_months):
            v = mvals[m]
            body = None
            if v and is_formula(r, c):
                try:
                    e = model.explain(sheet, r, c)
                except Exception:      # noqa: BLE001
                    e = None
                if e:
                    e = tidy(e)
                    try:
                        if not is_plain_number(e) and len(e) <= 250 and _close(eval_expr(e), v):
                            body = e
                    except Exception:  # noqa: BLE001
                        body = None
            mbodies.append(body)
        if not _close(sum(mvals), amount):
            warnings.append(f"{cur} · {semok or ''} 줄({r}행): 예산 {amount:,.0f}원과 월별 합 {sum(mvals):,.0f}원이 다릅니다.")
        items[cur].append(dict(row=r, semok=semok or "", amount=amount, basis=basis, mvals=mvals, mbodies=mbodies))
    out_items = []
    for it in order:
        lines = items[it]
        if not lines:
            continue
        s = sum(ln["amount"] for ln in lines)
        if it in group_sum and not _close(s, group_sum[it]):
            warnings.append(f"{it}: 세목 금액의 합 {s:,.0f}원이 소계 {group_sum[it]:,.0f}원과 다릅니다.")
        out_items.append(dict(item=it, lines=lines))
    grand = sum(ln["amount"] for it in out_items for ln in it["lines"])
    if total is None:
        warnings.append("합계 줄을 찾지 못했습니다 — 세목 금액을 더한 값을 합계로 씁니다.")
        total = grand
    elif not _close(grand, total):
        warnings.append(f"세목 금액의 합 {grand:,.0f}원이 합계 줄 {total:,.0f}원과 다릅니다.")
    if out_items and 0 < abs(grand) < 1000000:
        warnings.append(f"합계가 {grand:,.0f}원으로 작습니다 — 천원 단위로 적힌 파일이 아닌지 확인해 주세요(양식은 원 단위).")
    return dict(kind="budget", file=name, year=year, sheet=sheet, site=site, total=grand, items=out_items,
                calc=("computed" if stat["computed"] and not stat["cached"] else ("mixed" if stat["computed"] else "cached")),
                warnings=warnings)


# ══════════════════════════════════════════════════════════════════════════════
# 2. 전년도 사용내역(SAP G/L 내역 엑셀) 읽기
# ══════════════════════════════════════════════════════════════════════════════
_HEAD = {"전기일자": "date", "적요": "memo", "G/L계정": "acct", "G/L계정명": "acct_name", "차변금액": "debit", "대변금액": "credit",
         "차액": "amt", "거래처명": "vendor", "전표번호": "doc", "전표유형명": "doc_type", "생성자명": "creator",
         "귀속코드": "code", "귀속명": "code_name"}


def _to_num(v):
    if _is_num(v):
        return float(v)
    if v is None:
        return 0.0
    s = str(v).strip().replace(",", "").replace(" ", "")
    if not s:
        return 0.0
    neg = False
    if s.endswith("-"):                    # SAP 글자 내보내기의 뒤에 붙는 음수 부호
        neg, s = True, s[:-1]
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1]
    try:
        f = float(s)
    except ValueError:
        return None
    return -f if neg else f


def _to_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if _is_num(v) and 20000 < float(v) < 80000:          # 엑셀 날짜 일련번호
        return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(v))).date()
    if isinstance(v, str):
        s = v.strip()
        m = re.match(r"^(\d{4})[-./]\s*(\d{1,2})[-./]\s*(\d{1,2})", s) or re.match(r"^(\d{4})(\d{2})(\d{2})(?!\d)", s)
        if m:
            try:
                return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                return None
    return None


def _ledger_header(row):
    """머리글 줄이면 {필드: 열 번호(0부터)}, 아니면 None"""
    cells = [squash(v) for v in row]
    if "G/L계정명" not in cells:
        return None
    col = {}
    n_attr = 0
    for i, t in enumerate(cells):
        if t == "귀속정보":                 # 같은 이름의 열이 둘: 앞이 코드, 뒤가 이름
            col.setdefault("code" if n_attr == 0 else "code_name", i)
            n_attr += 1
        elif t in _HEAD:
            col.setdefault(_HEAD[t], i)
    return col


def parse_ledger(data, name=""):
    """SAP 사용내역 xlsx(bytes) → {"kind":"ledger","rows":[[전기일자, 계정, 계정명, 적요, 차변, 대변, 차액, 거래처명, 전표번호, 전표유형명, 생성자명]…],…}
    전기일자가 없는 줄(이월잔액 등)은 빼고 개수만 적는다. 사업자번호(주민등록번호)·거래처코드 열은 읽지 않는다."""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    except Exception:
        raise SiljeokError("엑셀(.xlsx) 파일이 아닙니다. .xls 파일이면 Excel 에서 .xlsx 로 저장한 뒤 올려 주세요.")
    try:
        rows, codes = [], {}
        skipped = dict(carry=0, carry_sum=0.0, nodate=0, nodate_sum=0.0, bad=0)
        found, why, has_creator = 0, "", True
        for ws in wb.worksheets:                       # 머리글이 맞는 시트는 모두 읽는다(내역을 시트 둘로 나눠 둔 파일도 됨)
            try:
                ws.reset_dimensions()                  # 내보내기 프로그램이 적어 둔 범위(dimension)가 틀려도 끝까지 읽게
            except Exception:                          # noqa: BLE001
                pass
            it = ws.iter_rows(values_only=True)
            col = None
            for k, row in enumerate(it):
                if k >= 15:
                    break
                col = _ledger_header(row)
                if col:
                    break
            if not col:
                continue
            if "amt" not in col and not ("debit" in col and "credit" in col):
                why = why or "사용내역에 금액 열(차액 또는 차변금액·대변금액)이 없습니다."
                continue
            if "date" not in col:
                why = why or "사용내역에 전기일자 열이 없습니다."
                continue
            found += 1
            has_creator = has_creator and "creator" in col

            def get(row, key, col=col):
                i = col.get(key)
                return row[i] if i is not None and i < len(row) else None

            for row in it:
                if row is None or all(v is None or v == "" for v in row):
                    continue
                acct_name = str(get(row, "acct_name") or "").strip()
                if not acct_name:
                    continue
                debit, credit = _to_num(get(row, "debit")), _to_num(get(row, "credit"))
                amt = _to_num(get(row, "amt")) if "amt" in col else None
                if "amt" not in col or get(row, "amt") in (None, ""):
                    amt = (debit or 0.0) - (credit or 0.0) if debit is not None and credit is not None else None
                if amt is None:
                    skipped["bad"] += 1
                    continue
                memo = str(get(row, "memo") or "").strip()
                d = _to_date(get(row, "date"))
                if d is None:
                    if memo == "이월잔액":
                        skipped["carry"] += 1
                        skipped["carry_sum"] += amt
                    else:
                        skipped["nodate"] += 1
                        skipped["nodate_sum"] += amt
                    continue
                code = str(get(row, "code") or "").strip()
                if code:
                    codes.setdefault(code, str(get(row, "code_name") or "").strip())
                acct = get(row, "acct")
                acct = str(int(acct)) if _is_num(acct) and float(acct) == int(acct) else str(acct or "").strip()
                doc = get(row, "doc")
                doc = str(int(doc)) if _is_num(doc) and float(doc) == int(doc) else str(doc or "").strip()
                rows.append([d.isoformat(), acct, acct_name, memo, debit or 0.0, credit or 0.0, amt, str(get(row, "vendor") or "").strip(), doc,
                             str(get(row, "doc_type") or "").strip(), str(get(row, "creator") or "").strip()])
        if not found:
            raise SiljeokError(why or "사용내역 머리글(G/L 계정명)을 찾지 못했습니다.")
    finally:
        wb.close()
    if not rows and not skipped["carry"]:
        raise SiljeokError("사용내역에 읽을 줄이 없습니다.")
    dates = [r[0] for r in rows]
    return dict(kind="ledger", file=name, rows=rows, codes=codes, skipped=skipped, n=len(rows), total=sum(r[6] for r in rows),
                date_min=min(dates) if dates else None, date_max=max(dates) if dates else None, has_creator=has_creator)


def parse_file(data, name=""):
    """올린 파일이 예산안인지 사용내역인지 알아보고 읽는다."""
    names = _sheet_names(data)
    if _budget_sheet(names):
        return parse_budget(data, name)
    try:
        return parse_ledger(data, name)
    except SiljeokError as e:
        if "머리글" in str(e):
            raise SiljeokError("알 수 없는 파일입니다 — 판관비 예산안(「NN년 예산」 시트가 있는 사업비 예산편성 엑셀)이나 "
                               "SAP 사용내역(G/L 계정명·차액·생성자명 열이 있는 엑셀)을 올려 주세요.")
        raise


# ══════════════════════════════════════════════════════════════════════════════
# 3. 양식에 넣을 값
# ══════════════════════════════════════════════════════════════════════════════
def auto_item(acct_name):
    """G/L 계정명 '판매관리비(통신비)_분양경비_D' → 양식 비목 '통신비'"""
    m = re.search(r"\(([^()]+)\)", acct_name or "")
    it = squash(m.group(1)) if m else squash(acct_name)
    return ITEM_ALIAS.get(it, it)


def _guess_year(sites):
    for s in sites:
        if s.get("next"):
            return int(s["next"]["year"])
    for s in sites:
        if s.get("prev"):
            return int(s["prev"]["year"]) + 1
    years = {}
    for s in sites:
        for lg in s.get("ledgers") or []:
            for r in lg["rows"]:
                years[r[0][:4]] = years.get(r[0][:4], 0) + 1
    if years:
        return int(max(years, key=years.get)) + 1
    today = datetime.date.today()
    return today.year + (1 if today.month >= 7 else 0)


def site_code(ledgers):
    """사용내역의 귀속코드 'P0500D' → 현장코드 'P0500' (가장 많이 나온 것)"""
    cnt = {}
    for lg in ledgers or []:
        for c in (lg.get("codes") or {}):
            cnt[c] = cnt.get(c, 0) + 1
    if not cnt:
        return ""
    c = max(cnt, key=cnt.get)
    m = re.match(r"^([A-Za-z]\d{3,5})[A-Za-z]?$", c)
    return m.group(1).upper() if m else c


def _ledger_rows(ledgers, prev_year):
    """여러 사용내역 파일의 줄을 모은다 → (쓸 줄, 뺀 줄[(사유, 줄)], 요약). 줄 = dict(LEDGER_FIELDS + file, period)"""
    use, drop = [], []
    seen = {}                      # 겹침 검사: 앞 파일들에서 나온 (전표번호, 계정, 일자, 차변, 대변, 적요) 개수
    info = dict(files=[], dup=0, other_year=0, carry=0, carry_sum=0.0, nodate=0)
    for fi, lg in enumerate(ledgers or []):
        mine = {}
        n_use = 0
        for r in lg["rows"]:
            d = dict(zip(LEDGER_FIELDS, r))
            d["file"] = lg.get("file") or f"파일{fi + 1}"
            d["creator"] = d["creator"] or NO_CREATOR
            y, mth = int(d["date"][:4]), int(d["date"][5:7])
            d["period"] = PERIODS[0] if mth <= SPLIT_MONTH else PERIODS[1]
            d["first"] = mth <= SPLIT_MONTH
            if y != prev_year:
                info["other_year"] += 1
                drop.append((f"전기일자가 {prev_year}년이 아님", d))
                continue
            if d["doc"]:
                key = (d["doc"], d["acct"], d["date"], d["debit"], d["credit"], d["memo"])
                mine[key] = mine.get(key, 0) + 1
                if mine[key] <= seen.get(key, 0):             # 앞 파일에 이미 있는 만큼은 겹친 것
                    info["dup"] += 1
                    drop.append(("다른 파일과 겹침", d))
                    continue
            use.append(d)
            n_use += 1
        for k, v in mine.items():
            seen[k] = max(seen.get(k, 0), v)
        sk = lg.get("skipped") or {}
        info["carry"] += sk.get("carry", 0)
        info["carry_sum"] += sk.get("carry_sum", 0.0)
        info["nodate"] += sk.get("nodate", 0)
        info["files"].append(dict(file=lg.get("file"), n=lg.get("n", len(lg["rows"])), used=n_use, date_min=lg.get("date_min"),
                                  date_max=lg.get("date_max"), carry=sk.get("carry", 0), nodate=sk.get("nodate", 0)))
    return use, drop, info


def _reversals(included, others):
    """고른 작성자의 전표(included)와 금액이 반대이고 적요가 겹치는, 고르지 않은 작성자의 전표 짝 [(넣은 줄, 되돌린 줄)…]
    (예: 다른 현장 비용이 잘못 들어왔다가 '[귀속부서대체] …' 로 빠진 경우 — 실적에는 넣은 줄만 남아 부풀려진다)"""
    def core(memo):
        return re.sub(r"\s+", "", re.sub(r"^(\s*\[[^\]]*\])+", "", memo or ""))

    index = {}
    for d in included:
        index.setdefault((d["acct_name"], round(d["amt"])), []).append(d)
    out, used = [], set()
    for o in others:
        if not o["amt"] or not any(k in o["memo"] for k in ("대체", "취소", "정정", "역분개", "반제", "오류", "오기")):
            continue                      # 정산 청구처럼 금액만 우연히 같은 전표는 알리지 않는다
        for d in index.get((o["acct_name"], round(-o["amt"])), []):
            if id(d) in used:
                continue
            x, y = core(d["memo"]), core(o["memo"])
            if x and y and (x in y or y in x) and len(min(x, y, key=len)) >= 4:
                used.add(id(d))
                out.append((d, o))
                break
    return out


def _ordered(names):
    known = [n for n in ORDER if n in names]
    return known + [n for n in names if n not in ORDER]


def _plan_site(s, year, detail):
    prev_year = year - 1
    yy, py = f"{year % 100:02d}", f"{prev_year % 100:02d}"
    name = str(s.get("name") or "").strip()
    code = str(s.get("code") or "").strip()
    prev, nxt, ledgers = s.get("prev"), s.get("next"), s.get("ledgers") or []
    warnings, errors = [], []
    if not name:
        errors.append("현장명을 적어 주세요.")
    if nxt and int(nxt["year"]) != year:
        warnings.append(f"금년도 예산안 파일은 {nxt['year']}년 예산입니다(작성 기준 {year}년).")
    if prev and int(prev["year"]) != prev_year:
        warnings.append(f"전년도 예산 파일은 {prev['year']}년 예산입니다(작성 기준 전년도 {prev_year}년).")
    for tag, b in (("전년도 예산", prev), ("금년도 예산안", nxt)):
        for w in (b or {}).get("warnings") or []:
            warnings.append(f"{tag}: {w}")
    prev_items = {it["item"]: it["lines"] for it in (prev or {}).get("items") or []}
    next_items = {it["item"]: it["lines"] for it in (nxt or {}).get("items") or []}

    # ── 사용내역
    use, drop, linfo = _ledger_rows(ledgers, prev_year)
    has_ledger = bool(ledgers)
    chosen = s.get("creators")
    selected = set(chosen or [])
    creators = {}
    for d in use:
        c = creators.setdefault(d["creator"], dict(name=d["creator"], n=0, amt=0.0))
        c["n"] += 1
        c["amt"] += d["amt"]
    creator_list = sorted(creators.values(), key=lambda c: (-c["n"], c["name"]))
    for c in creator_list:
        c["on"] = c["name"] in selected
    accmap = s.get("accmap") or {}
    accounts, included = {}, []
    for d in use:
        if d["creator"] not in selected:
            drop.append(("고르지 않은 작성자", d))
            continue
        auto = auto_item(d["acct_name"])
        target = accmap.get(d["acct_name"]) or auto
        a = accounts.setdefault(d["acct_name"], dict(acct=d["acct"], name=d["acct_name"], auto=auto, item=target, n=0, f=0.0, g=0.0))
        a["n"] += 1
        a["f" if d["first"] else "g"] += d["amt"]
        if target == EXCLUDE:
            drop.append(("뺀 계정", d))
            continue
        d["item"] = squash(target)
        included.append(d)
    act = {}                       # 비목 → {"f": {계정명: 금액}, "g": {...}}
    for d in included:
        slot = act.setdefault(d["item"], dict(f={}, g={}))
        part = slot["f" if d["first"] else "g"]
        part[d["acct_name"]] = part.get(d["acct_name"], 0.0) + d["amt"]
    if has_ledger and not selected:
        warnings.append("사용내역에서 작성자를 고르지 않았습니다 — 실적 칸이 모두 0원입니다.")
    for hit in _reversals(included, [d for why, d in drop if why == "고르지 않은 작성자"])[:3]:
        a, b = hit
        warnings.append(f"고른 작성자의 전표를 되돌린 전표가 고르지 않은 작성자에게 있습니다: {auto_item(a['acct_name'])} {a['amt']:,.0f}원 "
                        f"({a['creator']} {a['date'][5:]} 「{a['memo'][:30]}」 ↔ {b['creator']} {b['date'][5:]} 「{b['memo'][:30]}」) — "
                        f"{b['creator']} 도 고르거나 계정별 반영 내용에서 빼면 실적에서 상계됩니다.")
    codes = sorted({c for lg in ledgers for c in (lg.get("codes") or {})})
    if len(codes) > 1:
        warnings.append("사용내역에 귀속코드가 여러 개입니다: " + ", ".join(codes) + " — 다른 현장 파일이 섞였는지 확인해 주세요.")
    elif codes and code and not codes[0].upper().startswith(code.upper()):
        warnings.append(f"사용내역의 귀속코드({codes[0]})가 현장코드({code})와 다릅니다.")
    if linfo["dup"]:
        warnings.append(f"사용내역 파일끼리 겹친 전표 {linfo['dup']}줄은 한 번만 넣었습니다.")
    if linfo["other_year"]:
        warnings.append(f"전기일자가 {prev_year}년이 아닌 {linfo['other_year']}줄은 뺐습니다.")
    if linfo["nodate"]:
        warnings.append(f"전기일자가 없는 {linfo['nodate']}줄(이월잔액 아님)은 뺐습니다.")
    if not code:
        warnings.append("현장코드(P0000)가 비어 있습니다 — 제목에 (P    ) 로 들어갑니다.")
    if not nxt:
        warnings.append(f"금년도 예산안이 없습니다 — {yy}년 예산 칸은 0 입니다.")
    if not prev and (has_ledger or not nxt):
        warnings.append(f"전년도 예산이 없습니다 — {py}년 예산 칸은 0 입니다.")
    if prev and not has_ledger:
        warnings.append(f"전년도 사용내역을 올리지 않아 {py}년 실적 칸(1~8월·9~12월)은 비워 두었습니다.")

    # ── 비목 줄
    names = []
    for src in (list(prev_items), list(next_items), [k for k, v in act.items() if any(v["f"].values()) or any(v["g"].values())]):
        for n in src:
            if n not in names:
                names.append(n)
    names = _ordered(names)
    inc = s.get("inc") or {}
    note_j, note_m = str(s.get("note_diff") or "").strip(), str(s.get("note_change") or "").strip()
    rows = []
    for idx, it in enumerate(names):
        pl, nl = prev_items.get(it) or [], next_items.get(it) or []
        d_parts = [ln["amount"] for ln in pl if ln["amount"]]
        d = sum(d_parts)
        e = float(inc.get(it) or 0)
        months = []
        for m in range(12):
            parts, total = [], 0.0
            for ln in nl:
                v = ln["mvals"][m]
                if not v:
                    continue
                total += v
                parts.append(ln["mbodies"][m] or num_text(v))
            if not parts:
                months.append(dict(v=0.0, f=None))
                continue
            f = None
            if not (len(parts) == 1 and is_plain_number(parts[0])):
                f = tidy("+".join(parts))
                try:
                    ok = _close(eval_expr(f), total) and len(f) <= 600
                except Exception:      # noqa: BLE001
                    ok = False
                if not ok:
                    f = "+".join(num_text(ln["mvals"][m]) for ln in nl if ln["mvals"][m])
                if is_plain_number(f):
                    f = None
            months.append(dict(v=total, f=f))
        k = sum(m["v"] for m in months)
        k_budget = sum(ln["amount"] for ln in nl)
        if nl and not _close(k, k_budget):
            warnings.append(f"{it}: 금년도 예산 {k_budget:,.0f}원 가운데 월별 칸에 있는 금액은 {k:,.0f}원입니다 — 양식의 {yy}년 예산은 월별 합계로 들어갑니다.")
        a = act.get(it)
        if has_ledger:
            fp = [(n, v) for n, v in sorted((a or dict(f={}))["f"].items()) if v]
            gp = [(n, v) for n, v in sorted((a or dict(g={}))["g"].items()) if v]
            f_val, g_val = sum(v for _, v in fp), sum(v for _, v in gp)
        elif not prev:
            fp, gp, f_val, g_val = [], [], 0.0, 0.0                     # 신규 현장: 전년도 실적 0
        else:
            fp, gp, f_val, g_val = [], [], None, None                   # 사용내역 없음: 비움
        c = d + e
        h = (f_val or 0.0) + (g_val or 0.0)
        row = dict(item=it, semok=SEMOK, c=c, d=d, e=e, f=f_val, g=g_val, h=h, i=(h / c if c else 0.0), k=k,
                   l=(k / h if h and k else 0.0), basis=basis_lines(nl),
                   d_f=("+".join(num_text(x) for x in d_parts) if len(d_parts) > 1 else None),
                   f_f=("+".join(num_text(v) for _, v in fp) if len(fp) > 1 else None),
                   g_f=("+".join(num_text(v) for _, v in gp) if len(gp) > 1 else None),
                   months=months, f_parts=fp, g_parts=gp,
                   accts=sorted(set((a or {}).get("f", {})) | set((a or {}).get("g", {}))),
                   note_j=((" " + note_j) if idx == 0 else " 상동") if note_j else "",
                   note_m=((" " + note_m) if idx == 0 else " 상동") if note_m else "",
                   in_prev=bool(pl), in_next=bool(nl), in_act=bool(a))
        rows.append(row)
    only_act = [r["item"] for r in rows if r["in_act"] and not r["in_prev"] and not r["in_next"]]
    if only_act:
        warnings.append("예산에는 없고 실적에만 있는 비목: " + ", ".join(only_act) + " — 줄을 따로 만들었습니다(계정별 반영 내용에서 다른 비목으로 옮기거나 뺄 수 있습니다).")
    if not rows and not errors:
        errors.append("양식에 넣을 내용이 없습니다 — 전년도 예산·금년도 예산안·사용내역 가운데 하나는 올려 주세요.")
    tot = dict(c=sum(r["c"] for r in rows), d=sum(r["d"] for r in rows), e=sum(r["e"] for r in rows),
               f=(sum(r["f"] or 0.0 for r in rows) if rows and rows[0]["f"] is not None else None),
               g=(sum(r["g"] or 0.0 for r in rows) if rows and rows[0]["g"] is not None else None),
               h=sum(r["h"] for r in rows), k=sum(r["k"] for r in rows),
               months=[sum(r["months"][m]["v"] for r in rows) for m in range(12)])
    tot["i"] = tot["h"] / tot["c"] if tot["c"] else 0.0
    tot["l"] = tot["k"] / tot["h"] if tot["h"] and tot["k"] else 0.0
    if prev and not _close(tot["d"], prev["total"]):
        warnings.append(f"전년도 예산 합계가 파일({prev['total']:,.0f}원)과 다릅니다({tot['d']:,.0f}원).")
    title = f"[현장명 : {name or '현장명'} ({code or 'P    '})]"
    excluded_creators = [c for c in creator_list if not c["on"]]
    out = dict(name=name, code=code, title=title, rows=rows, total=tot, warnings=warnings, errors=errors,
               has_ledger=has_ledger, creators=creator_list, accounts=sorted(accounts.values(), key=lambda a: (a["acct"], a["name"])),
               ledger=dict(files=linfo["files"], carry=linfo["carry"], carry_sum=linfo["carry_sum"], dup=linfo["dup"],
                           other_year=linfo["other_year"], nodate=linfo["nodate"], included=len(included), dropped=len(drop),
                           codes=codes, suggest_code=site_code(ledgers)),
               prev=(dict(file=prev.get("file"), year=prev["year"], total=prev["total"], site=prev.get("site"), calc=prev.get("calc")) if prev else None),
               next=(dict(file=nxt.get("file"), year=nxt["year"], total=nxt["total"], site=nxt.get("site"), calc=nxt.get("calc")) if nxt else None),
               items_all=_ordered(list(dict.fromkeys(ORDER + names))))
    if detail:
        out["included"] = included
        out["dropped"] = drop
        out["excluded_creators"] = excluded_creators
    return out


def plan(state, detail=False):
    """화면 입력(state) → 현장별 양식 값. detail=True 면 반영·제외 전표 줄도 담는다(근거자료용)."""
    sites = state.get("sites") or []
    try:
        year = int(state.get("year") or 0)
    except (TypeError, ValueError):
        year = 0
    if not year:
        year = _guess_year(sites)
    out = [_plan_site(s, year, detail) for s in sites]
    errors = []
    if not sites:
        errors.append("현장을 하나 이상 넣어 주세요.")
    for i, s in enumerate(out):
        for e in s["errors"]:
            errors.append(f"{s['name'] or f'현장 {i + 1}'}: {e}")
    seen = {}
    for i, s in enumerate(out):
        if s["name"] and s["name"] in seen:
            errors.append(f"현장 {seen[s['name']] + 1} 과 현장 {i + 1} 의 현장명이 같습니다({s['name']}) — 근거자료에서 가려낼 수 있게 이름을 다르게 적어 주세요.")
        seen.setdefault(s["name"], i)
    today = datetime.date.today().strftime("%y%m%d")
    tag = out[0]["name"] if len(out) == 1 and out[0]["name"] else "PM팀"
    tag = re.sub(r'[\\/:*?"<>|]', "_", tag)
    return dict(ok=True, year=year, prev_year=year - 1, sites=out, errors=errors,
                form_name=f"{year - 1}년_사업관리비_실적계획_양식_{tag}_{today}.xlsx",
                evidence_name=f"{year - 1}년_사업관리비_실적_근거자료_{tag}_{today}.xlsx")


def slim(p):
    """plan 결과에서 화면에 필요 없는 큰 목록을 뺀다."""
    for s in p["sites"]:
        for k in ("included", "dropped", "excluded_creators"):
            s.pop(k, None)
    return p


# ══════════════════════════════════════════════════════════════════════════════
# 4. 사내 양식 쓰기 (xlsx 내부 XML 을 그대로 고침 — 틀의 서식·열 너비·인쇄 설정 유지)
# ══════════════════════════════════════════════════════════════════════════════
def _text_units(s):
    """글자 너비 어림(글자 크기 = 1): 한글·전각 1, 영문 대문자 0.68, 소문자·숫자 0.56, 공백·쉼표·점 0.31"""
    w = 0.0
    for ch in s:
        o = ord(ch)
        if o >= 0x1100:
            w += 1.0
        elif ch in " .,:;'|!il()[]":
            w += 0.34
        elif ch.isdigit() or ch.islower():
            w += 0.57
        elif ch.isupper():
            w += 0.70
        else:
            w += 0.62
    return w


def wrapped_lines(text, col_width, font_pt=10.0):
    """줄바꿈 칸의 줄 수 어림: 열 너비(글자 수 단위, 기본 글꼴 11pt 숫자 폭 8px) 안에 font_pt 글자가 몇 줄로 접히는지"""
    if not text:
        return 1
    avail = max(8.0, col_width * 8.0 + 5 - 8)            # 칸 안쪽 너비(px)
    em = font_pt * 96 / 72
    n = 0
    for seg in str(text).split("\n"):
        n += max(1, int(math.ceil(_text_units(seg) * em / avail - 1e-9)))
    return n


FIT_MARGIN = 1.12        # 인쇄(축소 배율)에서는 화면보다 글자가 넓게 잡혀 화면에서 겨우 맞는 숫자가 #### 로 찍힌다 → 여유를 둔다


def _num_px(v, bold):
    """#,##0 으로 보이는 금액의 너비 어림(10pt 맑은 고딕, px)"""
    s = f"{abs(v):,.0f}"
    px = sum(3.6 if ch == "," else 7.5 for ch in s) + (5.0 if v < 0 else 0)
    return px * (1.08 if bold else 1.0)


def _pct_px(v, bold):
    """0.0% 로 보이는 비율의 너비 어림(0 은 '-')"""
    if not v:
        return 6.0
    s = f"{abs(v) * 100:.1f}%"
    px = sum(3.6 if ch == "." else (11.5 if ch == "%" else 7.5) for ch in s) + (5.0 if v < 0 else 0)
    return px * (1.08 if bold else 1.0)


def _template_years(sh):
    """틀 머리글의 연도(두 자리): 금년도 = 'NN년 월별 예산' 칸"""
    for ref, c in sh.cells.items():
        t = sh.value(ref)
        if isinstance(t, str):
            m = re.match(r"^\s*(\d{2})\s*년\s*월별\s*예산", t)
            if m:
                return int(m.group(1))
    return None


def template_year(template):
    """틀 머리글의 금년도(예: 2027) — 'NN년 월별 예산' 글자에서. 못 찾으면 None"""
    try:
        z = zipfile.ZipFile(io.BytesIO(template))
        sst = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
    except Exception:              # noqa: BLE001
        return None
    m = re.search(r"(\d{2})\s*년\s*월별\s*예산", sst)
    return 2000 + int(m.group(1)) if m else None


def build_form(template, state):
    """틀(bytes) + 화면 입력 → 채운 양식 xlsx(bytes)"""
    import xlsxio as X
    from lxml import etree
    q = X.q
    p = plan(state, detail=False)
    if p["errors"]:
        raise SiljeokError(p["errors"][0])
    book = X.Book(template)
    names = [s["name"] for s in book.sheet_list()]
    sh = book.sheet(MAIN if MAIN in names else names[0])
    chk = [squash(sh.value(f"A{r}")) for r in range(1, 7)]
    if not (chk[0].startswith("[현장명") and chk[1] == "비목" and chk[4] == "소계" and chk[5] == "합계"):
        raise SiljeokError("실적계획 양식 틀이 아닙니다(1행 제목 · 2~3행 머리글 · 4행 비목 · 5행 소 계 · 6행 합 계).")
    # 머리글 연도
    base = _template_years(sh)
    yy = p["year"] % 100
    if base is not None and base != yy:
        delta = yy - base
        for ref in list(sh.cells):
            r, _ = X.cell_key(ref)
            t = sh.value(ref)
            if r <= 3 and isinstance(t, str) and re.search(r"\d{2}\s*년", t):
                sh.set_text(ref, re.sub(r"(\d{2})(\s*년)", lambda m: f"{(int(m.group(1)) + delta) % 100:02d}{m.group(2)}", t))
    sh.unshare()
    snaps = {k: sh.row_copy(r) for k, r in SEED.items()}
    merges = sh.merged()
    mc = sh.root.find(q("mergeCells"))
    if mc is not None:
        for m in list(mc):
            mc.remove(m)
    widths = {}
    cols = sh.root.find(q("cols"))
    if cols is not None:
        for col in cols:
            for c in range(int(col.get("min")), min(int(col.get("max")), 60) + 1):
                if col.get("width"):
                    widths[c] = float(col.get("width"))

    def place(kind, dst):
        sh.put_row(snaps[kind], dst)
        src = SEED[kind]
        for (r1, c1), (r2, c2) in merges:
            if r1 == src:
                sh.merge(f"{X.num2col(c1)}{dst}:{X.num2col(c2)}{dst + (r2 - r1)}")

    # 배치: 현장 블록 사이에 빈 행 하나
    blocks, r = [], 1
    for s in p["sites"]:
        n = len(s["rows"])
        blocks.append(dict(site=s, title=r, pairs=[(r + 3 + 2 * i, r + 4 + 2 * i) for i in range(n)], total=r + 3 + 2 * n))
        r = blocks[-1]["total"] + 2
    for b in blocks:
        place("title", b["title"])
        place("h1", b["title"] + 1)
        place("h2", b["title"] + 2)
        for dr, sr in b["pairs"]:
            place("data", dr)
            place("sub", sr)
        place("total", b["total"])

    def setf(ref, formula, value):
        c = sh._cell(ref)
        sh._clear(c)
        etree.SubElement(c, q("f")).text = formula
        if value is not None:
            etree.SubElement(c, q("v")).text = num_text(value)
        book.touched.add(sh.path)

    def setv(ref, v, formula=None):
        if formula:
            setf(ref, formula, v)
        elif v is None:
            sh.clear(ref)
        else:
            sh.set_number(ref, v)

    def fit(ref, v, bold, pct=False):
        """숫자가 열 너비보다 길면 그 칸만 '셀에 맞춤'(#### 방지) — 열 너비는 건드리지 않는다"""
        if v is None:
            return
        col = X.cell_key(ref)[1]
        avail = widths.get(col, 8.38) * 8.0 + 5 - 7
        if (_pct_px(v, bold) if pct else _num_px(v, bold)) * FIT_MARGIN > avail:
            sh.shrink(ref)

    def fit_row(r, vals, months, bold):
        for col, key in (("C", "c"), ("D", "d"), ("E", "e"), ("F", "f"), ("G", "g"), ("H", "h")):
            fit(f"{col}{r}", vals[key], bold)
        fit(f"K{r}", vals["k"], True)                 # 금년도 예산 칸은 자료 줄도 굵은 글씨
        fit(f"I{r}", vals["i"], bold, True)
        fit(f"L{r}", vals["l"], bold, True)
        for col, v in zip(MONTH_COLS, months):
            fit(f"{col}{r}", v, bold)

    def height(r, lines):
        row = sh.rows[r]
        if lines <= 1:
            row.attrib.pop("ht", None)
            row.attrib.pop("customHeight", None)
        else:
            row.set("ht", num_text(34.8 if lines == 2 else round(15.6 * lines, 1)))
            row.set("customHeight", "1")

    N_COL, J_COL, M_COL = 14, 10, 13
    for b in blocks:
        s = b["site"]
        sh.set_text(f"A{b['title']}", s["title"])
        subs = []
        for row, (dr, sr) in zip(s["rows"], b["pairs"]):
            sh.set_text(f"A{dr}", row["item"])
            sh.set_text(f"B{dr}", row["semok"])
            setf(f"C{dr}", f"D{dr}+E{dr}", row["c"])
            setv(f"D{dr}", row["d"], row["d_f"])
            setv(f"E{dr}", row["e"])
            setv(f"F{dr}", row["f"], row["f_f"])
            setv(f"G{dr}", row["g"], row["g_f"])
            setf(f"H{dr}", f"F{dr}+G{dr}", row["h"])
            setf(f"I{dr}", f"IF(C{dr}=0,0,H{dr}/C{dr})", row["i"])
            setf(f"K{dr}", f"SUM(O{dr}:Z{dr})", row["k"])
            setf(f"L{dr}", f"IF(OR(H{dr}=0,K{dr}=0),0,K{dr}/H{dr})", row["l"])
            for ref, t in ((f"J{dr}", row["note_j"]), (f"M{dr}", row["note_m"]), (f"N{dr}", row["basis"])):
                if t:
                    sh.set_text(ref, t)
                else:
                    sh.clear(ref)
            for col, m in zip(MONTH_COLS, row["months"]):
                setv(f"{col}{dr}", m["v"], m["f"])
            sh.clear(f"AA{dr}")
            fit_row(dr, row, [m["v"] for m in row["months"]], False)
            height(dr, max(wrapped_lines(row["basis"], widths.get(N_COL, 74.6)), wrapped_lines(row["note_j"], widths.get(J_COL, 34.0)),
                           wrapped_lines(row["note_m"], widths.get(M_COL, 34.0))))
            # 소 계
            setf(f"C{sr}", f"D{sr}+E{sr}", row["c"])
            for col, key in (("D", "d"), ("E", "e"), ("F", "f"), ("G", "g")):
                setf(f"{col}{sr}", f"SUM({col}{dr}:{col}{dr})", row[key] or 0.0)
            setf(f"H{sr}", f"F{sr}+G{sr}", row["h"])
            setf(f"I{sr}", f"IF(C{sr}=0,0,H{sr}/C{sr})", row["i"])
            setf(f"K{sr}", f"SUM(O{sr}:Z{sr})", row["k"])
            setf(f"L{sr}", f"IF(OR(H{sr}=0,K{sr}=0),0,K{sr}/H{sr})", row["l"])
            for col, m in zip(MONTH_COLS, row["months"]):
                setf(f"{col}{sr}", f"SUM({col}{dr}:{col}{dr})", m["v"])
            fit_row(sr, row, [m["v"] for m in row["months"]], True)
            for col in ("J", "M", "N", "AA"):
                sh.clear(f"{col}{sr}")
            subs.append(sr)
        t, tot = b["total"], s["total"]
        plus = lambda col: "SUM(" + ",".join(f"{col}{x}" for x in subs) + ")"
        setf(f"C{t}", f"D{t}+E{t}", tot["c"])
        for col, key in (("D", "d"), ("E", "e"), ("F", "f"), ("G", "g")):
            setf(f"{col}{t}", plus(col), tot[key] or 0.0)
        setf(f"H{t}", f"F{t}+G{t}", tot["h"])
        setf(f"I{t}", f"IF(C{t}=0,0,H{t}/C{t})", tot["i"])
        setf(f"K{t}", f"SUM(O{t}:Z{t})", tot["k"])
        setf(f"L{t}", f"IF(OR(H{t}=0,K{t}=0),0,K{t}/H{t})", tot["l"])
        for col, v in zip(MONTH_COLS, tot["months"]):
            setf(f"{col}{t}", plus(col), v)
        fit_row(t, tot, tot["months"], True)
        for col in ("J", "M", "N", "AA"):
            sh.clear(f"{col}{t}")
    last = blocks[-1]["total"]
    dim = sh.root.find(q("dimension"))
    if dim is not None:
        dim.set("ref", f"A1:AA{last}")
    dn = book.wb.find(q("definedNames"))
    area = f"{X._qname(sh.name)}!$A$1:$AA${last}"
    done = False
    if dn is not None:
        for d in dn:
            if d.get("name") == "_xlnm.Print_Area":
                d.text = area
                done = True
    if not done:
        if dn is None:
            dn = etree.Element(q("definedNames"))
            book.wb.find(q("sheets")).addnext(dn)
        d = etree.SubElement(dn, q("definedName"))
        d.set("name", "_xlnm.Print_Area")
        d.set("localSheetId", "0")
        d.text = area
    for sv in sh.root.iter(q("sheetView")):
        for sel in sv.findall(q("selection")):
            sel.set("activeCell", "A1")
            sel.set("sqref", "A1")
    book.touched.add(sh.path)
    return book.save()


# ══════════════════════════════════════════════════════════════════════════════
# 5. 실적 근거자료
# ══════════════════════════════════════════════════════════════════════════════
def build_evidence(state):
    """실적 칸(F·G)에 넣은 값의 근거: 비목별 집계(양식 값과 대사) · 계정별 집계(작성자별) · 반영 내역 · 제외 내역"""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter as L
    p = plan(state, detail=True)
    if p["errors"]:
        raise SiljeokError(p["errors"][0])
    sites = [s for s in p["sites"] if s["has_ledger"]]
    if not sites:
        raise SiljeokError("사용내역을 올린 현장이 없어 근거자료를 만들 수 없습니다.")
    py = p["prev_year"]
    FONT = "맑은 고딕"
    f_title = Font(name=FONT, size=13, bold=True, color="1F3864")
    f_note = Font(name=FONT, size=9, color="595959")
    f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    f_body = Font(name=FONT, size=10)
    f_bold = Font(name=FONT, size=10, bold=True)
    fill_head = PatternFill("solid", fgColor="1F3864")
    fill_sub = PatternFill("solid", fgColor="2E5395")
    fill_tot = PatternFill("solid", fgColor="D9E2F3")
    fill_site = PatternFill("solid", fgColor="F2F3F5")
    side = Side(style="thin", color="B7C3D6")
    border = Border(left=side, right=side, top=side, bottom=side)
    NF = '#,##0;[Red]-#,##0;"-"'
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    mid = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")

    w_site = min(52.0, max(14.0, max(_text_units(s["name"] + " 계") for s in sites) * 10 * 96 / 72 / 8 + 3))
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "비목별 집계"
    ws2 = wb.create_sheet("계정별 집계")
    ws3 = wb.create_sheet("반영 내역")
    ws4 = wb.create_sheet("제외 내역")
    D3 = "'반영 내역'"

    def head(ws, r, c, text, fill=fill_head):
        cell = ws.cell(r, c, text)
        cell.font, cell.fill, cell.alignment, cell.border = f_head, fill, center, border
        return cell

    def body(ws, r, c, v=None, nf=None, al=None, font=None, fill=None):
        cell = ws.cell(r, c)
        if v is not None:
            cell.value = v
        cell.font, cell.border = font or f_body, border
        cell.alignment = al or (right if nf else left)
        if nf:
            cell.number_format = nf
        if fill:
            cell.fill = fill
        return cell

    # ── 반영 내역 / 제외 내역
    cols3 = ["현장", "기간", "비목", "전기일자", "G/L 계정", "G/L 계정명", "적요", "거래처명", "차변금액", "대변금액", "차액", "전표번호",
             "전표유형명", "생성자명", "사용내역 파일"]
    w3 = [w_site, 8, 12, 11, 10, 32, 40, 22, 13, 13, 13, 12, 18, 9, 22]
    for i, (t, w) in enumerate(zip(cols3, w3), 1):
        head(ws3, 1, i, t)
        ws3.column_dimensions[L(i)].width = w
    r3 = 2
    for s in sites:
        for d in sorted(s["included"], key=lambda d: (ORDER.index(d["item"]) if d["item"] in ORDER else 99, d["item"], d["date"], d["doc"])):
            vals = [s["name"], d["period"], d["item"], datetime.date.fromisoformat(d["date"]), d["acct"], d["acct_name"], d["memo"], d["vendor"],
                    d["debit"], d["credit"], d["amt"], d["doc"], d["doc_type"], d["creator"], d["file"]]
            for i, v in enumerate(vals, 1):
                cell = body(ws3, r3, i, v, NF if i in (9, 10, 11) else None)
                if i == 4:
                    cell.number_format = "yyyy-mm-dd"
                    cell.alignment = center
                elif i in (2, 3, 5, 12, 14):
                    cell.alignment = center
                else:
                    cell.alignment = Alignment(horizontal="right" if i in (9, 10, 11) else "left", vertical="center")
            r3 += 1
    last3 = max(r3 - 1, 2)
    ws3.freeze_panes = "A2"
    ws3.auto_filter.ref = f"A1:{L(len(cols3))}{last3}"
    cols4 = ["현장", "제외 사유"] + cols3[1:2] + cols3[3:]
    w4 = [w_site, 22, 8] + w3[3:]
    for i, (t, w) in enumerate(zip(cols4, w4), 1):
        head(ws4, 1, i, t)
        ws4.column_dimensions[L(i)].width = w
    r4 = 2
    for s in sites:
        for why, d in sorted(s["dropped"], key=lambda x: (x[0], x[1]["creator"], x[1]["date"], x[1]["doc"])):
            vals = [s["name"], why, d["period"], datetime.date.fromisoformat(d["date"]), d["acct"], d["acct_name"], d["memo"], d["vendor"],
                    d["debit"], d["credit"], d["amt"], d["doc"], d["doc_type"], d["creator"], d["file"]]
            for i, v in enumerate(vals, 1):
                cell = body(ws4, r4, i, v, NF if i in (9, 10, 11) else None)
                if i == 4:
                    cell.number_format = "yyyy-mm-dd"
                    cell.alignment = center
                elif i in (3, 5, 12, 14):
                    cell.alignment = center
                else:
                    cell.alignment = Alignment(horizontal="right" if i in (9, 10, 11) else "left", vertical="center")
            r4 += 1
    last4 = max(r4 - 1, 2)
    ws4.freeze_panes = "A2"
    ws4.auto_filter.ref = f"A1:{L(len(cols4))}{last4}"

    def sumifs(site_ref, extra):
        """반영 내역의 차액 합: 현장 + 조건들((열 글자, 값)…).
        SUMIFS 는 조건 글자의 ~ * ? 를 특수문자로 읽으므로(현장명에 들어갈 수 있음) 글자를 그대로 견주는 SUMPRODUCT 를 쓴다."""
        parts = [f"({D3}!$A$2:$A${last3}={site_ref})"]
        for col, crit in extra:
            parts.append(f"({D3}!${col}$2:${col}${last3}={crit})")
        return "=SUMPRODUCT(" + "*".join(parts) + f"*{D3}!$K$2:$K${last3})"

    # ── 비목별 집계
    ws1["A1"] = f"{py}년 실적 근거자료 — 사업관리비 실적계획 양식의 {py % 100:02d}년 실적(1~8월·9~12월) 칸"
    ws1["A1"].font = f_title
    ws1["A2"] = (f"작성 {datetime.date.today().isoformat()} · 금액 = 차변 − 대변(차액), 단위 원 · 전기일자 기준 1~8월 / 9~12월 · "
                 "집계 칸은 「반영 내역」 시트의 전표를 더하는 수식이고, '양식에 넣은 값'과 차이가 0 이면 맞습니다.")
    ws1["A2"].font = f_note
    heads1 = ["현장", "비목", "1~8월 실적", "9~12월 실적", "연간", "양식에 넣은 값\n(1~8월)", "양식에 넣은 값\n(9~12월)", "차이\n(집계 − 양식)", "모은 G/L 계정"]
    for i, (t, w) in enumerate(zip(heads1, [w_site, 14, 15, 15, 15, 15, 15, 13, 70]), 1):
        head(ws1, 4, i, t)
        ws1.column_dimensions[L(i)].width = w
    ws1.row_dimensions[4].height = 30
    r = 5
    for s in sites:
        first = r
        for row in s["rows"]:
            if not row["in_act"] and not (row["f"] or row["g"]):
                continue
            accts = row["accts"]
            body(ws1, r, 1, s["name"], al=mid)
            body(ws1, r, 2, row["item"], al=center)
            body(ws1, r, 3, sumifs(f"$A{r}", [("C", f"$B{r}"), ("B", f'"{PERIODS[0]}"')]), NF)
            body(ws1, r, 4, sumifs(f"$A{r}", [("C", f"$B{r}"), ("B", f'"{PERIODS[1]}"')]), NF)
            body(ws1, r, 5, f"=C{r}+D{r}", NF)
            body(ws1, r, 6, row["f"] or 0, NF)
            body(ws1, r, 7, row["g"] or 0, NF)
            body(ws1, r, 8, f"=(C{r}-F{r})+(D{r}-G{r})", NF)
            body(ws1, r, 9, ", ".join(accts))
            r += 1
        if r == first:
            body(ws1, r, 1, s["name"], al=mid)
            body(ws1, r, 2, "(반영한 전표 없음)", al=center)
            for c in range(3, 10):
                body(ws1, r, c, 0 if c < 9 else None, NF if c < 9 else None)
            r += 1
        body(ws1, r, 1, s["name"] + " 계", al=mid, font=f_bold, fill=fill_tot)
        body(ws1, r, 2, None, font=f_bold, fill=fill_tot)
        for c in range(3, 9):
            body(ws1, r, c, f"=SUM({L(c)}{first}:{L(c)}{r - 1})", NF, font=f_bold, fill=fill_tot)
        body(ws1, r, 9, None, fill=fill_tot)
        r += 2
    # 작성 기준
    ws1.cell(r, 1, "작성 기준").font = f_bold
    r += 1
    for s in sites:
        lg = s["ledger"]
        on = [c for c in s["creators"] if c["on"]]
        off = [c for c in s["creators"] if not c["on"]]
        moved = [a for a in s["accounts"] if a["item"] != a["auto"]]
        lines = [f"[{s['name']}]",
                 "사용내역 파일: " + " / ".join(f"{f['file']} ({f['date_min'] or '-'} ~ {f['date_max'] or '-'}, {f['n']:,}줄)" for f in lg["files"]),
                 f"실적에 넣은 작성자({len(on)}명): " + (", ".join(f"{c['name']} {c['amt']:,.0f}원({c['n']}건)" for c in on) or "없음"),
                 f"넣지 않은 작성자({len(off)}명): " + (", ".join(f"{c['name']} {c['amt']:,.0f}원({c['n']}건)" for c in off) or "없음")]
        extra = []
        if lg["carry"]:
            extra.append(f"이월잔액 {lg['carry']}줄({lg['carry_sum']:,.0f}원) 제외")
        if lg["other_year"]:
            extra.append(f"{py}년이 아닌 전표 {lg['other_year']}줄 제외")
        if lg["dup"]:
            extra.append(f"파일끼리 겹친 전표 {lg['dup']}줄은 한 번만 반영")
        if lg["nodate"]:
            extra.append(f"전기일자 없는 {lg['nodate']}줄 제외")
        if extra:
            lines.append("그 밖에: " + " · ".join(extra))
        if moved:
            lines.append("비목을 바꾸거나 뺀 계정: " + ", ".join(f"{a['name']} → {'제외' if a['item'] == EXCLUDE else a['item']}" for a in moved))
        for i, t in enumerate(lines):
            c = ws1.cell(r, 1, t)
            c.font = f_bold if i == 0 else f_body
            c.alignment = Alignment(vertical="center")
            r += 1
        r += 1
    ws1.freeze_panes = "A5"

    # ── 계정별 집계 (작성자별)
    names = []
    for s in sites:
        for c in s["creators"]:
            if c["on"] and c["name"] not in names:
                names.append(c["name"])
    ws2["A1"] = f"{py}년 실적 — G/L 계정별·작성자별 집계 (실적에 넣은 전표)"
    ws2["A1"].font = f_title
    ws2["A2"] = "금액 = 차변 − 대변(차액), 단위 원 · 「반영 내역」 시트를 더하는 수식 · 맨 아래 검산이 0 이면 작성자별 합이 합계와 같습니다."
    ws2["A2"].font = f_note
    heads2 = ["현장", "G/L 계정", "G/L 계정명", "양식 비목", "1~8월", "9~12월", "합계"] + names
    for i, t in enumerate(heads2, 1):
        head(ws2, 4, i, t, fill_head if i <= 7 else fill_sub)
        ws2.column_dimensions[L(i)].width = [w_site, 10, 36, 12, 15, 15, 15][i - 1] if i <= 7 else 13
    r = 5
    for s in sites:
        first = r
        for a in s["accounts"]:
            if a["item"] == EXCLUDE:
                continue
            body(ws2, r, 1, s["name"], al=mid)
            body(ws2, r, 2, a["acct"], al=center)
            body(ws2, r, 3, a["name"])
            body(ws2, r, 4, a["item"], al=center)
            body(ws2, r, 5, sumifs(f"$A{r}", [("F", f"$C{r}"), ("B", f'"{PERIODS[0]}"')]), NF)
            body(ws2, r, 6, sumifs(f"$A{r}", [("F", f"$C{r}"), ("B", f'"{PERIODS[1]}"')]), NF)
            body(ws2, r, 7, f"=E{r}+F{r}", NF)
            for j, nm in enumerate(names):
                body(ws2, r, 8 + j, sumifs(f"$A{r}", [("F", f"$C{r}"), ("N", f"{L(8 + j)}$4")]), NF)
            r += 1
        if r == first:
            continue
        body(ws2, r, 1, s["name"] + " 계", al=mid, font=f_bold, fill=fill_tot)
        for c in (2, 3, 4):
            body(ws2, r, c, None, fill=fill_tot)
        for c in range(5, 8 + len(names)):
            body(ws2, r, c, f"=SUM({L(c)}{first}:{L(c)}{r - 1})", NF, font=f_bold, fill=fill_tot)
        if names:
            r += 1
            body(ws2, r, 3, "검산 (작성자별 합 − 합계, 0 이면 맞음)", font=f_note)
            body(ws2, r, 7, f"=SUM({L(8)}{r - 1}:{L(7 + len(names))}{r - 1})-G{r - 1}", NF)
        r += 2
    ws2.freeze_panes = "E5"
    for ws in (ws1, ws2, ws3, ws4):
        ws.sheet_view.showGridLines = False
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = 9
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws3.print_title_rows = "1:1"
    ws4.print_title_rows = "1:1"
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()
