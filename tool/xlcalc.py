# -*- coding: utf-8 -*-
"""계산용 양식(xlsx)의 수식을 Excel 없이 계산한다 — 웹(브라우저 Pyodide)용.

  m = Model(경로 또는 바이트)        # openpyxl 로 한 번 읽어 수식을 파싱
  m.set("현장입력", "D5", 123)      # 입력값(덮어쓰기)
  m.reset()                         # 덮어쓴 값 모두 지움
  m.value("계산", "J68")            # 계산값 (숫자·글·bool·None·XLError)
  m.values("27년 예산", "A1:U300")  # 2차원 목록

지원 함수: IF AND OR NOT SUM MIN MAX AVERAGE ROUND ROUNDUP ROUNDDOWN INT MOD N TEXT COUNTIF COUNT COUNTA
          INDEX MATCH COLUMN ROW IFERROR ISERROR ISNUMBER ISBLANK LEFT RIGHT MID LEN TODAY ABS VALUE
연산자: + - * / ^ & = <> < > <= >= 단항 - , % 와 범위·이름·다른 시트 참조, 배열 비교(J11:U11<>0) 와 INDEX(…,0)·MATCH(TRUE,…).
Excel 과 같게: 빈 칸은 0/"", ROUND 는 .5 올림(0 에서 먼 쪽), 글자 비교는 대소문자 무시, 숫자<글자<논리.
"""
import datetime
import io
import math
import re
from decimal import Decimal, ROUND_HALF_UP, ROUND_UP, ROUND_DOWN

EPOCH = datetime.datetime(1899, 12, 30)


class XLError(Exception):
    def __init__(self, code):
        super().__init__(code)
        self.code = code

    def __repr__(self):
        return self.code

    __str__ = __repr__


ERR_DIV, ERR_VAL, ERR_NA, ERR_REF, ERR_NAME, ERR_NUM = "#DIV/0!", "#VALUE!", "#N/A", "#REF!", "#NAME?", "#NUM!"


def col2num(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def num2col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def to_serial(d):
    if isinstance(d, datetime.datetime):
        delta = d - EPOCH
        return delta.days + delta.seconds / 86400
    if isinstance(d, datetime.date):
        return (datetime.datetime(d.year, d.month, d.day) - EPOCH).days
    return d


# ───────────────────────────── 토큰·구문
_TOKEN = re.compile(r"""
   (?P<ws>\s+)
  |(?P<str>"(?:[^"]|"")*")
  |(?P<err>\#(?:REF!|N/A|DIV/0!|VALUE!|NAME\?|NUM!|NULL!))
  |(?P<ref>(?:(?:'(?:[^']|'')+'|[^\W\d][\w.]*)!)?
     (?:\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+))(?![\w(])
  |(?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?|\.\d+)
  |(?P<bool>TRUE|FALSE)(?![\w(])
  |(?P<func>[A-Za-z_][\w.]*)\s*\(
  |(?P<name>[^\W\d][\w.]*)
  |(?P<op><>|<=|>=|[-+*/^&=<>%,();])
""", re.X)


def tokenize(s):
    out, i = [], 0
    while i < len(s):
        m = _TOKEN.match(s, i)
        if not m:
            raise ValueError(f"수식을 읽을 수 없음: {s[i:i + 20]!r}")
        i = m.end()
        kind = m.lastgroup
        if kind == "ws":
            continue
        txt = m.group(kind)
        if kind == "func":
            txt = m.group(0)[:-1].strip()
        out.append((kind, txt))
    return out


def _parse_ref(txt):
    """'시트'!$A$1:$B$2 → (sheet|None, r1, c1, r2, c2)  (열 전체는 r=None, 행 전체는 c=None)"""
    sheet = None
    if "!" in txt:
        pre, txt = txt.rsplit("!", 1)
        sheet = pre[1:-1].replace("''", "'") if pre.startswith("'") else pre
    a, _, b = txt.partition(":")

    def one(p):
        m = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", p)
        if m:
            return int(m.group(2)), col2num(m.group(1))
        m = re.fullmatch(r"\$?([A-Z]{1,3})", p)
        if m:
            return None, col2num(m.group(1))
        m = re.fullmatch(r"\$?(\d+)", p)
        return int(m.group(1)), None
    r1, c1 = one(a)
    r2, c2 = one(b) if b else (r1, c1)
    return sheet, r1, c1, r2, c2


class Parser:
    """우선순위: 비교 < & < +- < */ < ^ < % < 단항- < 원소"""

    def __init__(self, s):
        self.t = tokenize(s)
        self.i = 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self):
        tok = self.t[self.i]
        self.i += 1
        return tok

    def parse(self):
        e = self.cmp()
        if self.i != len(self.t):
            raise ValueError(f"수식 끝에 남은 글자: {self.t[self.i:]}")
        return e

    def cmp(self):
        e = self.concat()
        while self.peek()[0] == "op" and self.peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self.take()[1]
            e = ("bin", op, e, self.concat())
        return e

    def concat(self):
        e = self.add()
        while self.peek() == ("op", "&"):
            self.take()
            e = ("bin", "&", e, self.add())
        return e

    def add(self):
        e = self.mul()
        while self.peek()[0] == "op" and self.peek()[1] in ("+", "-"):
            op = self.take()[1]
            e = ("bin", op, e, self.mul())
        return e

    def mul(self):
        e = self.pow()
        while self.peek()[0] == "op" and self.peek()[1] in ("*", "/"):
            op = self.take()[1]
            e = ("bin", op, e, self.pow())
        return e

    def pow(self):
        e = self.unary()
        while self.peek() == ("op", "^"):
            self.take()
            e = ("bin", "^", e, self.unary())
        return e

    def unary(self):
        if self.peek()[0] == "op" and self.peek()[1] in ("-", "+"):
            op = self.take()[1]
            e = self.unary()
            return ("un", op, e)
        return self.postfix()

    def postfix(self):
        e = self.atom()
        while self.peek() == ("op", "%"):
            self.take()
            e = ("pct", e)
        return e

    def atom(self):
        kind, txt = self.take()
        if kind == "num":
            return ("num", float(txt))
        if kind == "str":
            return ("str", txt[1:-1].replace('""', '"'))
        if kind == "bool":
            return ("bool", txt == "TRUE")
        if kind == "err":
            return ("err", txt)
        if kind == "ref":
            return ("ref",) + _parse_ref(txt)
        if kind == "name":
            return ("name", txt)
        if kind == "func":
            args = []
            if self.peek() == ("op", ")"):
                self.take()
                return ("func", txt.upper(), args)
            while True:
                if self.peek()[0] == "op" and self.peek()[1] in (",", ")"):
                    args.append(("empty",))
                else:
                    args.append(self.cmp())
                kind2, t2 = self.take()
                if t2 == ")":
                    return ("func", txt.upper(), args)
                if t2 != ",":
                    raise ValueError("함수 인수 구분 오류")
        if kind == "op" and txt == "(":
            e = self.cmp()
            if self.take() != ("op", ")"):
                raise ValueError("괄호 오류")
            return e
        raise ValueError(f"뜻밖의 토큰 {kind} {txt!r}")


# ───────────────────────────── 값
class Range:
    __slots__ = ("sheet", "r1", "c1", "r2", "c2")

    def __init__(self, sheet, r1, c1, r2, c2):
        self.sheet, self.r1, self.c1, self.r2, self.c2 = sheet, r1, c1, r2, c2


class Arr:
    __slots__ = ("rows",)

    def __init__(self, rows):
        self.rows = rows


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def num(v):
    """산술용 숫자: 빈칸 0, 논리 1/0, 숫자 글자는 숫자로, 그 밖의 글자는 #VALUE!"""
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if _is_num(v):
        return float(v)
    if isinstance(v, XLError):
        raise v
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        if s == "":
            raise XLError(ERR_VAL)
        try:
            if s.endswith("%"):
                return float(s[:-1]) / 100
            return float(s)
        except ValueError:
            raise XLError(ERR_VAL)
    raise XLError(ERR_VAL)


def general(v):
    """Excel '일반' 표시 글자"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if _is_num(v):
        f = float(v)
        if f == int(f) and abs(f) < 1e15:
            return str(int(f))
        s = f"{f:.10g}"
        if "e" in s:
            m, e = s.split("e")
            return f"{m}E{int(e):+03d}"
        return s
    if isinstance(v, XLError):
        raise v
    return str(v)


def txt(v):
    return general(v)


def truth(v):
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if _is_num(v):
        return v != 0
    if isinstance(v, XLError):
        raise v
    if isinstance(v, str):
        if v.upper() == "TRUE":
            return True
        if v.upper() == "FALSE":
            return False
        raise XLError(ERR_VAL)
    raise XLError(ERR_VAL)


def _rank(v):
    if v is None or _is_num(v):
        return 0
    if isinstance(v, str):
        return 1
    if isinstance(v, bool):
        return 2
    raise XLError(ERR_VAL)


def compare(a, b):
    """-1/0/1. 빈칸은 상대가 글자면 "", 아니면 0. 숫자 < 글자 < 논리. 글자는 대소문자 무시."""
    for v in (a, b):
        if isinstance(v, XLError):
            raise v
    if a is None:
        a = "" if isinstance(b, str) else (False if isinstance(b, bool) else 0.0)
    if b is None:
        b = "" if isinstance(a, str) else (False if isinstance(a, bool) else 0.0)
    ra, rb = _rank(a), _rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if ra == 1:
        a, b = a.upper(), b.upper()
    else:
        a, b = float(a), float(b)
    return -1 if a < b else (1 if a > b else 0)


def _dec(x):
    return Decimal(repr(float(x)))


def xround(x, n=0, mode=ROUND_HALF_UP):
    n = int(n)
    q = Decimal(1).scaleb(-n)
    return float(_dec(x).quantize(q, rounding=mode))


def excel_text(v, fmt):
    """TEXT(값, 형식) — 이 양식에서 쓰는 형식: 날짜(yyyy mm dd), 0%, #,##0, #,##0.0##, 00"""
    if isinstance(v, XLError):
        raise v
    if isinstance(v, str) and not re.fullmatch(r"\s*-?[\d,.]+\s*%?", v or "x"):
        return v
    x = num(v)
    f = fmt
    if re.search(r"[ymd]", f) and not re.search(r"[#0]", f):
        d = EPOCH + datetime.timedelta(days=int(math.floor(x)))
        out = f
        out = re.sub(r"yyyy", f"{d.year:04d}", out)
        out = re.sub(r"yy", f"{d.year % 100:02d}", out)
        out = re.sub(r"mm", f"{d.month:02d}", out)
        out = re.sub(r"(?<![m])m(?![m])", str(d.month), out)
        out = re.sub(r"dd", f"{d.day:02d}", out)
        out = re.sub(r"(?<![d])d(?![d])", str(d.day), out)
        return out
    pct = "%" in f
    if pct:
        x *= 100
        f = f.replace("%", "")
    neg = x < 0
    x = abs(x)
    ip, _, dp = f.partition(".")
    min_dec = dp.count("0")
    max_dec = len(re.sub(r"[^0#]", "", dp))
    r = xround(x, max_dec)
    s = f"{r:.{max_dec}f}" if max_dec else str(int(r))
    if max_dec:
        whole, frac = s.split(".")
        frac = frac.rstrip("0")
        if len(frac) < min_dec:
            frac = frac + "0" * (min_dec - len(frac))
    else:
        whole, frac = s, ""
    zeros = ip.count("0")
    if len(whole) < zeros:
        whole = "0" * (zeros - len(whole)) + whole
    if "," in ip:
        whole = f"{int(whole):,}" if whole.isdigit() else whole
        if len(whole) < zeros:
            whole = "0" * (zeros - len(whole)) + whole
    out = whole + ("." + frac if frac else "")
    if pct:
        out += "%"
    return ("-" + out) if neg and out.strip("0.,%") else out


def _crit(c):
    """COUNTIF 기준 → 검사 함수"""
    if isinstance(c, XLError):
        raise c
    s = general(c) if not isinstance(c, str) else c
    m = re.match(r"^(<>|<=|>=|=|<|>)?(.*)$", s, re.S)
    op, val = m.group(1) or "=", m.group(2)
    try:
        v = float(val.replace(",", "")) if val.strip() != "" else None
    except ValueError:
        v = None
    if v is not None and not re.search(r"[?*]", val):
        if op == "=":
            return lambda x: _is_num(x) and float(x) == v
        if op == "<>":
            return lambda x: not (_is_num(x) and float(x) == v)
        f = {"<": lambda x: x < v, ">": lambda x: x > v, "<=": lambda x: x <= v, ">=": lambda x: x >= v}[op]
        return lambda x: _is_num(x) and f(float(x))
    pat = "^" + "".join(".*" if ch == "*" else "." if ch == "?" else re.escape(ch) for ch in val) + "$"
    rx = re.compile(pat, re.I | re.S)
    if op == "=":
        if val == "":
            return lambda x: x is None or x == ""
        return lambda x: isinstance(x, str) and bool(rx.match(x))
    if op == "<>":
        if val == "":
            return lambda x: not (x is None or x == "")
        return lambda x: not (isinstance(x, str) and rx.match(x))
    raise XLError(ERR_VAL)


# ───────────────────────────── 모델
class Model:
    def __init__(self, src, log=None):
        import openpyxl
        if isinstance(src, (bytes, bytearray)):
            src = io.BytesIO(src)
        wb = openpyxl.load_workbook(src, read_only=True, data_only=False)
        self.cells = {}          # sheet -> {(r,c): value | ("F", text)}
        self.maxrow = {}
        self.maxcol = {}
        self.sheetnames = list(wb.sheetnames)
        for ws in wb.worksheets:
            d = {}
            mr = mc = 0
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if v is None:
                        continue
                    if isinstance(v, (datetime.datetime, datetime.date)):
                        v = to_serial(v)
                    elif isinstance(v, str) and v.startswith("="):
                        v = ("F", v[1:])
                    d[(cell.row, cell.column)] = v
                    if cell.row > mr:
                        mr = cell.row
                    if cell.column > mc:
                        mc = cell.column
            self.cells[ws.title] = d
            self.maxrow[ws.title] = mr
            self.maxcol[ws.title] = mc
        self.names = {}
        for name, dn in wb.defined_names.items():
            try:
                self.names[name] = _parse_ref(dn.attr_text.replace("$", "")) if "!" in dn.attr_text else None
            except Exception:
                self.names[name] = None
        wb.close()
        self.ast = {}
        self.over = {}
        self.cache = {}
        self.stack = set()
        self.today = to_serial(datetime.datetime.combine(datetime.date.today(), datetime.time()))

    # ---- 입력
    def set(self, sheet, ref, value):
        r, c = self._rc(ref)
        if isinstance(value, (datetime.datetime, datetime.date)):
            value = to_serial(value)
        self.over[(sheet, r, c)] = value
        self.cache.clear()

    def set_many(self, items):
        for sheet, ref, value in items:
            r, c = self._rc(ref)
            if isinstance(value, (datetime.datetime, datetime.date)):
                value = to_serial(value)
            self.over[(sheet, r, c)] = value
        self.cache.clear()

    def reset(self):
        self.over.clear()
        self.cache.clear()

    def name_ref(self, name):
        return self.names.get(name)

    @staticmethod
    def _rc(ref):
        m = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", ref)
        return int(m.group(2)), col2num(m.group(1))

    # ---- 읽기
    def value(self, sheet, ref):
        r, c = self._rc(ref)
        return self.cell(sheet, r, c)

    def values(self, sheet, ref):
        a, _, b = ref.partition(":")
        r1, c1 = self._rc(a)
        r2, c2 = self._rc(b) if b else (r1, c1)
        return [[self.cell(sheet, r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]

    def name_value(self, name):
        nr = self.names.get(name)
        if not nr:
            return XLError(ERR_NAME)
        sheet, r1, c1, r2, c2 = nr
        if (r1, c1) == (r2, c2):
            return self.cell(sheet, r1, c1)
        return [[self.cell(sheet, r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]

    def cell(self, sheet, r, c):
        key = (sheet, r, c)
        if key in self.over:
            return self.over[key]
        if key in self.cache:
            return self.cache[key]
        raw = self.cells.get(sheet, {}).get((r, c))
        if not (isinstance(raw, tuple) and raw and raw[0] == "F"):
            return raw
        if key in self.stack:
            v = XLError("#CIRC!")
        else:
            self.stack.add(key)
            try:
                v = self._eval_formula(raw[1], sheet, r, c)
            except XLError as e:
                v = e
            except RecursionError:
                v = XLError("#CIRC!")
            except Exception as e:                      # 지원하지 않는 수식 등
                v = XLError(f"#ERR:{type(e).__name__}:{str(e)[:40]}")
            finally:
                self.stack.discard(key)
        self.cache[key] = v
        return v

    def _eval_formula(self, text, sheet, r, c):
        ast = self.ast.get(text)
        if ast is None:
            ast = Parser(text).parse()
            self.ast[text] = ast
        v = self._ev(ast, sheet, r, c)
        if isinstance(v, Range):
            v = self._scalar(v, sheet, r, c)
        elif isinstance(v, Arr):
            v = v.rows[0][0] if v.rows and v.rows[0] else None
        if v is None:                      # 빈 칸을 돌려주는 수식(=A1, IF(…,A1,…))은 Excel 에서 0
            v = 0.0
        return v

    # ---- 평가
    def _scalar(self, v, sheet, r, c):
        """범위를 값 하나로(암시적 교차: 같은 행/열, 아니면 왼쪽 위)"""
        if isinstance(v, Range):
            if v.r1 == v.r2 and v.c1 == v.c2:
                return self.cell(v.sheet, v.r1, v.c1)
            if v.r1 == v.r2 and v.c1 <= c <= v.c2:
                return self.cell(v.sheet, v.r1, c)
            if v.c1 == v.c2 and v.r1 <= r <= v.r2:
                return self.cell(v.sheet, r, v.c1)
            raise XLError(ERR_VAL)
        if isinstance(v, Arr):
            return v.rows[0][0]
        return v

    def _rows_of(self, v):
        """범위/배열 → 2차원 값 목록"""
        if isinstance(v, Range):
            r1, r2 = v.r1, v.r2
            c1, c2 = v.c1, v.c2
            if r1 is None:
                r1, r2 = 1, self.maxrow.get(v.sheet, 0)
            if c1 is None:
                c1, c2 = 1, self.maxcol.get(v.sheet, 0)
            return [[self.cell(v.sheet, rr, cc) for cc in range(c1, c2 + 1)] for rr in range(r1, r2 + 1)]
        if isinstance(v, Arr):
            return v.rows
        return [[v]]

    def _ev(self, n, sheet, r, c):
        k = n[0]
        if k == "num":
            return n[1]
        if k == "str":
            return n[1]
        if k == "bool":
            return n[1]
        if k == "empty":
            return None
        if k == "err":
            raise XLError(n[1])
        if k == "ref":
            sh = n[1] or sheet
            if sh not in self.cells:
                raise XLError(ERR_REF)
            return Range(sh, n[2], n[3], n[4], n[5])
        if k == "name":
            nr = self.names.get(n[1])
            if nr is None:
                if n[1] in self.names:
                    raise XLError(ERR_NAME)
                raise XLError(ERR_NAME)
            return Range(*nr)
        if k == "pct":
            return num(self._scalar(self._ev(n[1], sheet, r, c), sheet, r, c)) / 100
        if k == "un":
            v = self._ev(n[2], sheet, r, c)
            if isinstance(v, (Range, Arr)):
                rows = self._rows_of(v)
                return Arr([[(-num(x) if n[1] == "-" else num(x)) for x in row] for row in rows])
            x = num(v)
            return -x if n[1] == "-" else x
        if k == "bin":
            return self._bin(n[1], self._ev(n[2], sheet, r, c), self._ev(n[3], sheet, r, c), sheet, r, c)
        if k == "func":
            return self._func(n[1], n[2], sheet, r, c)
        raise ValueError(k)

    def _bin(self, op, a, b, sheet, r, c):
        arr = isinstance(a, (Range, Arr)) or isinstance(b, (Range, Arr))
        if arr and not (isinstance(a, Range) and a.r1 == a.r2 and a.c1 == a.c2) \
                and not (isinstance(b, Range) and b.r1 == b.r2 and b.c1 == b.c2):
            A = self._rows_of(a) if isinstance(a, (Range, Arr)) else None
            B = self._rows_of(b) if isinstance(b, (Range, Arr)) else None
            if A is not None and B is not None and (len(A) != len(B) or len(A[0]) != len(B[0])):
                raise XLError(ERR_VAL)
            rows = A if A is not None else B
            out = []
            for i, row in enumerate(rows):
                line = []
                for j in range(len(row)):
                    x = A[i][j] if A is not None else a
                    y = B[i][j] if B is not None else b
                    try:
                        line.append(self._op(op, x, y))
                    except XLError as e:
                        line.append(e)
                out.append(line)
            return Arr(out)
        a = self._scalar(a, sheet, r, c)
        b = self._scalar(b, sheet, r, c)
        return self._op(op, a, b)

    @staticmethod
    def _op(op, a, b):
        if op in ("=", "<>", "<", ">", "<=", ">="):
            cmp = compare(a, b)
            return {"=": cmp == 0, "<>": cmp != 0, "<": cmp < 0, ">": cmp > 0, "<=": cmp <= 0, ">=": cmp >= 0}[op]
        if op == "&":
            return txt(a) + txt(b)
        x, y = num(a), num(b)
        if op == "+":
            return x + y
        if op == "-":
            return x - y
        if op == "*":
            return x * y
        if op == "/":
            if y == 0:
                raise XLError(ERR_DIV)
            return x / y
        if op == "^":
            try:
                return math.pow(x, y)
            except (ValueError, OverflowError):
                raise XLError(ERR_NUM)
        raise ValueError(op)

    def _nums(self, args, sheet, r, c, direct_text=True):
        """SUM/MIN/MAX 류: 범위 안의 숫자만, 직접 인수는 숫자로 바꿈"""
        out = []
        for a in args:
            v = self._ev(a, sheet, r, c)
            if isinstance(v, (Range, Arr)):
                for row in self._rows_of(v):
                    for x in row:
                        if isinstance(x, XLError):
                            raise x
                        if _is_num(x):
                            out.append(float(x))
            elif v is not None:
                out.append(num(v) if direct_text else v)
        return out

    def _func(self, name, args, sheet, r, c):
        ev = lambda i: self._ev(args[i], sheet, r, c)
        sc = lambda i: self._scalar(ev(i), sheet, r, c)
        if name == "IF":
            cond = truth(sc(0))
            if cond:
                return sc(1) if len(args) > 1 else True
            return sc(2) if len(args) > 2 else False
        if name == "IFERROR":
            try:
                v = sc(0)
            except XLError:
                return sc(1)
            return sc(1) if isinstance(v, XLError) else v
        if name in ("AND", "OR"):
            vals = []
            for a in args:
                v = self._ev(a, sheet, r, c)
                if isinstance(v, (Range, Arr)):
                    for row in self._rows_of(v):
                        for x in row:
                            if isinstance(x, (bool, int, float)) and x is not None:
                                vals.append(truth(x))
                else:
                    vals.append(truth(v))
            if not vals:
                raise XLError(ERR_VAL)
            return all(vals) if name == "AND" else any(vals)
        if name == "NOT":
            return not truth(sc(0))
        if name == "SUM":
            return sum(self._nums(args, sheet, r, c))
        if name in ("MIN", "MAX"):
            xs = self._nums(args, sheet, r, c)
            return (min(xs) if name == "MIN" else max(xs)) if xs else 0.0
        if name == "AVERAGE":
            xs = self._nums(args, sheet, r, c)
            if not xs:
                raise XLError(ERR_DIV)
            return sum(xs) / len(xs)
        if name == "ROUND":
            return xround(num(sc(0)), num(sc(1)) if len(args) > 1 else 0)
        if name == "ROUNDUP":
            return xround(num(sc(0)), num(sc(1)) if len(args) > 1 else 0, ROUND_UP)
        if name == "ROUNDDOWN":
            return xround(num(sc(0)), num(sc(1)) if len(args) > 1 else 0, ROUND_DOWN)
        if name == "INT":
            return float(math.floor(num(sc(0))))
        if name == "ABS":
            return abs(num(sc(0)))
        if name == "MOD":
            x, y = num(sc(0)), num(sc(1))
            if y == 0:
                raise XLError(ERR_DIV)
            return x - y * math.floor(x / y)
        if name == "N":
            v = sc(0)
            if isinstance(v, XLError):
                raise v
            if v is None or isinstance(v, str):
                return 0.0
            return num(v)
        if name == "VALUE":
            return num(sc(0))
        if name == "TEXT":
            return excel_text(sc(0), txt(sc(1)))
        if name == "LEFT":
            s = txt(sc(0))
            return s[:int(num(sc(1))) if len(args) > 1 else 1]
        if name == "RIGHT":
            s = txt(sc(0))
            k = int(num(sc(1))) if len(args) > 1 else 1
            return s[-k:] if k else ""
        if name == "MID":
            s = txt(sc(0))
            st = int(num(sc(1)))
            return s[st - 1:st - 1 + int(num(sc(2)))]
        if name == "LEN":
            return float(len(txt(sc(0))))
        if name == "TODAY":
            return float(self.today)
        if name == "COLUMN":
            if not args:
                return float(c)
            v = ev(0)
            return float(v.c1 if isinstance(v, Range) else c)
        if name == "ROW":
            if not args:
                return float(r)
            v = ev(0)
            return float(v.r1 if isinstance(v, Range) else r)
        if name in ("ISERROR", "ISNUMBER", "ISBLANK", "ISTEXT"):
            try:
                v = sc(0)
            except XLError as e:
                v = e
            if name == "ISERROR":
                return isinstance(v, XLError)
            if name == "ISNUMBER":
                return _is_num(v)
            if name == "ISBLANK":
                return v is None
            return isinstance(v, str)
        if name in ("COUNT", "COUNTA"):
            n = 0
            for a in args:
                v = self._ev(a, sheet, r, c)
                for row in self._rows_of(v):
                    for x in row:
                        if name == "COUNT" and _is_num(x):
                            n += 1
                        elif name == "COUNTA" and x is not None and x != "":
                            n += 1
            return float(n)
        if name == "COUNTIF":
            rows = self._rows_of(ev(0))
            test = _crit(sc(1))
            return float(sum(1 for row in rows for x in row if test(x)))
        if name == "INDEX":
            v = ev(0)
            rows = self._rows_of(v)
            ri = int(num(sc(1))) if len(args) > 1 and args[1][0] != "empty" else 0
            ci = int(num(sc(2))) if len(args) > 2 and args[2][0] != "empty" else 0
            if len(args) <= 2:
                if len(rows) == 1:
                    ri, ci = 1, ri
                elif rows and len(rows[0]) == 1:
                    ci = 1
            if ri == 0 and ci == 0:
                return Arr(rows)
            if ri == 0:
                return Arr([[row[ci - 1]] for row in rows])
            if ci == 0:
                if len(rows[ri - 1]) == 1:
                    return rows[ri - 1][0]
                return Arr([rows[ri - 1]])
            try:
                return rows[ri - 1][ci - 1]
            except IndexError:
                raise XLError(ERR_REF)
        if name == "MATCH":
            target = sc(0)
            rows = self._rows_of(ev(1))
            flat = [x for row in rows for x in row] if (len(rows) == 1 or (rows and len(rows[0]) == 1)) else None
            if flat is None:
                raise XLError(ERR_NA)
            mtype = int(num(sc(2))) if len(args) > 2 else 1
            if mtype == 0:
                for i, x in enumerate(flat):
                    if isinstance(x, XLError):
                        continue
                    if isinstance(target, bool) or isinstance(x, bool):
                        if isinstance(x, bool) and isinstance(target, bool) and x == target:
                            return float(i + 1)
                        continue
                    if _rank(x) == _rank(target) and compare(x, target) == 0 and not (x is None and target is None):
                        return float(i + 1)
                raise XLError(ERR_NA)
            best = None
            for i, x in enumerate(flat):
                if isinstance(x, XLError) or x is None or _rank(x) != _rank(target):
                    continue
                cmp = compare(x, target)
                if (mtype > 0 and cmp <= 0) or (mtype < 0 and cmp >= 0):
                    best = i + 1
                elif mtype > 0 and cmp > 0:
                    break
            if best is None:
                raise XLError(ERR_NA)
            return float(best)
        raise XLError(ERR_NAME)
