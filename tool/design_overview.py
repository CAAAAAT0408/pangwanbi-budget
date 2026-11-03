# -*- coding: utf-8 -*-
"""설계개요(엑셀·PDF) → 세대 구성(타입·세대수·공급면적)

  read_types(경로)  → dict(types=[dict(name, units, supply, excl, label, use, flags)], confidence, notes, …)

설계개요의 '평형별(단위세대) 면적표' — 머리글 유형(구분·타입)·세대수·전용면적·공급면적 — 를 찾아 그 아래 타입 줄을 읽는다.
  · 엑셀(xlsx·xls): 셀 위치로. 숨긴 시트보다 보이는 시트, 합계가 맞는 표를 고른다.
  · PDF 글자형: 단어 위치로(도면이 90° 돌아가 있으면 글자 방향에 맞춰 돌림).
  · PDF 그림형(도면을 그림으로 출력한 것 — 설계사무소 설계개요 PDF 대부분): Windows 내장 OCR(한국어)로 단어 위치를 얻는다.
    배율을 바꿔 세 번 읽어 칸마다 다수결, 소수점 빠짐(849343→84.9343)·B/8 혼동(948T→94BT)은 전용면적으로 바로잡는다.
  열은 숫자들의 가로 위치로 먼저 묶고, 머리글과 겹치는 열(없으면 머리글 순서를 지키는 가장 가까운 열)에 붙인다.
  끝으로 합계 줄(세대수 합·공급면적 합)과 맞는지, 공급 = 전용 + 주거공용 인지 확인해 알려 준다.
"""
import collections
import json
import os
import re
import subprocess
import sys
import tempfile
import time

Tok = collections.namedtuple("Tok", "x0 y0 x1 y1 s v")      # v = 엑셀 숫자 칸 값(그 밖엔 None)
HERE = os.path.dirname(os.path.abspath(__file__))
OCR_PS1 = os.path.join(HERE, "ocr_words.ps1")

RX_SUPPLY = re.compile(r"공\s*급(?!\s*비\s*율)(?!\s*대\s*비)")
RX_SALE = re.compile(r"분\s*양\s*면\s*적")      # '공급면적' 머리글이 없는 표는 분양면적(=전용+공용)이 공급면적
RX_UNITS = re.compile(r"세\s*대\s*수")
RX_UNITS2 = re.compile(r"세\s*대(?!\s*당)")
RX_EXCL = re.compile(r"전\s*용(?!\s*[률율])")
RX_TYPE = re.compile(r"유\s*형|주\s*택\s*형|타\s*입|TYPE|형\s*별|구\s*분", re.I)
RX_TOTAL = re.compile(r"(합|총)\s*계")
RX_SUB = re.compile(r"소\s*계")
RX_SCALE = re.compile(r"㎡|이하|초과|미만|이상")  # 규모 구분(60㎡이하 …)은 그룹 이름이 아님
RX_TITLE = re.compile(r"^[■□▣◆◇●○※▪◎]")
RX_HANGUL = re.compile(r"^[가-힣·,/()\s]+$")
UNIT_WORDS = re.compile(r"㎡|M²|M2|TYPE|타입|형", re.I)
OCR_UNIT_JUNK = re.compile(r"^(?:RN[I12]?|M[I12]?|RRF|W|N)(?=[A-Z0-9]|$)")   # ㎡ 오인식(m1·mI·rn2 …)
DIGIT_AS_LETTER = {"8": "B", "0": "D", "5": "S", "6": "G", "2": "Z", "4": "A", "7": "T"}
EXCL_MIN, EXCL_MAX = 14, 330


# ───────────────────────────── 단어(셀) 다루기
def _cx(t):
    return (t.x0 + t.x1) / 2


def _cy(t):
    return (t.y0 + t.y1) / 2


def _median(xs, default=1.0):
    xs = sorted(xs)
    return xs[len(xs) // 2] if xs else default


def _join_fragments(ts, h):
    """OCR 이 숫자를 두 토막 낸 것(31 | 9125m1 → 31.9125m1, 87, | 128.9 → 87,128.9) 잇기"""
    out = []
    for t in ts:
        if out:
            a = out[-1]
            if (t.x0 - a.x1) < 0.6 * h and re.fullmatch(r"\d{1,3},?", a.s) and re.match(r"\d{3}", t.s) \
                    and (a.s.endswith(",") or "." not in t.s):
                sep = "" if a.s.endswith(",") else "."
                out[-1] = Tok(a.x0, min(a.y0, t.y0), t.x1, max(a.y1, t.y1), a.s + sep + t.s, None)
                continue
        out.append(t)
    return out


def _lines(toks, join=False):
    """y 가운데가 가까운 단어끼리 한 줄로(줄 안은 왼→오)"""
    if not toks:
        return [], 1.0
    h = _median([t.y1 - t.y0 for t in toks]) or 1.0
    out = []
    for t in sorted(toks, key=_cy):
        y = _cy(t)
        if out and y - out[-1]["y"] <= h * 0.55:
            ln = out[-1]
            ln["t"].append(t)
            ln["y"] = sum(_cy(a) for a in ln["t"]) / len(ln["t"])
        else:
            out.append(dict(y=y, t=[t]))
    for ln in out:
        ln["t"].sort(key=lambda t: t.x0)
        if join:
            ln["t"] = _join_fragments(ln["t"], h)
    return out, h


def _is_title(line):
    """'■ 평형별 면적표' 같은 제목 줄(머리글로 쓰지 않음)"""
    ts = line["t"]
    if not ts:
        return False
    if RX_TITLE.match(ts[0].s):
        return True
    s = "".join(t.s for t in ts)
    return len(ts) <= 6 and bool(re.search(r"면\s*적\s*표|\(단\s*위", s))


def _phrases(line, rx):
    """줄 글자를 이어 붙여 rx 를 찾고, 걸친 단어들의 범위를 돌려줌('세 대 수'처럼 나뉜 글자도)"""
    if _is_title(line):
        return []
    s, spans = "", []
    for t in line["t"]:
        if s:
            s += " "
        spans.append((len(s), len(s) + len(t.s), t))
        s += t.s
    out = []
    for m in rx.finditer(s):
        ts = [t for a, b, t in spans if a < m.end() and b > m.start()]
        out.append(Tok(min(t.x0 for t in ts), min(t.y0 for t in ts), max(t.x1 for t in ts), max(t.y1 for t in ts),
                       m.group(0), None))
    return out


def _clean_num(s):
    s = s.replace(",", "").replace(" ", "")
    s = re.sub(r"^[(\[]+", "", s)
    s = re.sub(r"(?:㎡|[mM][²2lI1]?[\'’`′\"”:;.』\]]?|평|세대|세|호|대|%|ton|인|명|층|[)\]])+$", "", s)
    return re.sub(r"^[(\[]+", "", s)


def _num_parts(t):
    """숫자 칸 → (값, 숫자만(OCR 소수점 바로잡기용) 또는 None, 소수 여부). 숫자가 아니면 None"""
    if t.v is not None:
        return (float(t.v), None, float(t.v) != int(t.v))
    s = _clean_num(t.s)
    if re.fullmatch(r"\d+\.\d+|\d+", s):
        return (float(s), s.replace(".", ""), "." in s)
    return None


def _type_core(s, loose=False):
    u = UNIT_WORDS.sub("", s.upper().replace(" ", ""))
    u = re.sub(r"[()\[\]'’`]", "", u)
    if loose:
        u = re.sub(r"[^A-Z0-9\-]", "", u)
    return u


def _type_like(s, loose=False):
    """타입 이름처럼 생긴 글자: 84A, 94AT, 84㎡A, 101, 49A-1, 059.9800A, (OCR) 84m1A …"""
    return bool(re.fullmatch(r"0*\d{2,3}(?:\.\d+)?[A-Z][A-Z0-9\-]{0,5}|0*\d{2,3}", _type_core(s, loose)))


def _hangul_label(toks, xlim):
    ws = [t.s for t in toks if t.x1 <= xlim + 1e-6 and RX_HANGUL.match(t.s) and len(re.findall(r"[가-힣]", t.s)) >= 2]
    return " ".join(" ".join(ws).split())


# ───────────────────────────── 열 찾기
def _columns(lines):
    """본문 숫자들의 가로 범위를 묶어 열을 만든다(왼→오)"""
    toks = sorted((t for l in lines for t in l["t"] if _num_parts(t) is not None), key=_cx)
    cols = []
    for t in toks:
        best, bo = None, 0.0
        for c in cols:
            ov = min(t.x1, c["x1"]) - max(t.x0, c["x0"])
            r = ov / max(1e-9, min(t.x1 - t.x0, c["x1"] - c["x0"]))
            if r > bo:
                best, bo = c, r
        if best is not None and bo >= 0.5:
            best["n"] += 1
            best["x0"] += (t.x0 - best["x0"]) / best["n"]
            best["x1"] += (t.x1 - best["x1"]) / best["n"]
        else:
            cols.append(dict(x0=t.x0, x1=t.x1, n=1))
    cols.sort(key=lambda c: c["x0"] + c["x1"])
    return cols


def _match_cols(hdr, cols, h):
    """머리글(세대수·전용·공급) → 열: 겹치는 열 먼저, 없으면 머리글 순서를 지키는 가장 가까운 열"""
    keys = [k for k in ("units", "excl", "supply") if hdr.get(k) is not None]
    order = sorted(keys, key=lambda k: _cx(hdr[k]))

    def in_order(k, i, asg):
        return all((order.index(k) < order.index(k2)) == (i < j) for k2, j in asg.items())
    pairs = []
    for k in keys:
        p = hdr[k]
        for i, c in enumerate(cols):
            ov = min(p.x1, c["x1"]) - max(p.x0, c["x0"])
            if ov > 0:
                pairs.append((ov, k, i))
    pairs.sort(key=lambda x: (-x[0], x[2]))            # 겹침 큰 순, 같으면 왼쪽 열('공급면적(㎡/평)' → ㎡ 열)
    asg, used = {}, set()
    for ov, k, i in pairs:
        if k in asg or i in used or not in_order(k, i, asg):
            continue
        asg[k] = i
        used.add(i)
    for k in keys:
        if k in asg:
            continue
        cx = _cx(hdr[k])
        best = None
        for i, c in enumerate(cols):
            if i in used or not in_order(k, i, asg):
                continue
            d = abs((c["x0"] + c["x1"]) / 2 - cx)
            if d <= 12 * h and (best is None or d < best[0]):
                best = (d, i)
        if best:
            asg[k] = best[1]
            used.add(best[1])
    return asg


def _col_vals(line, col):
    """줄에서 그 열에 들어가는 숫자들(겹침이 큰 순)"""
    out = []
    for t in line["t"]:
        p = _num_parts(t)
        if p is None:
            continue
        ov = min(t.x1, col["x1"]) - max(t.x0, col["x0"])
        r = ov / max(1e-9, min(t.x1 - t.x0, col["x1"] - col["x0"]))
        if r >= 0.4:
            out.append((r, p))
    out.sort(key=lambda x: -x[0])
    return [p for r, p in out]


def _type_tok(toks, ucol, thdr, loose):
    """세대수 열 왼쪽에서 타입 이름 칸 고르기"""
    lim = ucol["x0"] + 0.25 * (ucol["x1"] - ucol["x0"])
    cands = [t for t in toks if t.x1 <= lim and _type_like(t.s, loose)]
    if not cands:
        return None
    if thdr is not None:
        ov = [t for t in cands if min(t.x1, thdr.x1) - max(t.x0, thdr.x0) > 0]
        if ov:
            return max(ov, key=lambda t: t.x1)
    lett = [t for t in cands if re.search(r"[A-Za-z]", t.s)]
    return max(lett or cands, key=lambda t: t.x1)


def _plain(vals):
    """타입 이름 없이도 면적표 줄로 볼 만한가: 세대수(정수) + 전용(14~330) + 공급(전용의 1.04~2배)"""
    u = [p for p in vals.get("units") or [] if not p[2] and p[0] >= 1]
    e = [p[0] for p in vals.get("excl") or [] if EXCL_MIN <= p[0] <= EXCL_MAX]
    if not u or not e:
        return False
    return any(e[0] * 1.04 <= p[0] <= e[0] * 2.0 for p in vals.get("supply") or [])


# ───────────────────────────── 표 찾기·줄 읽기(한 번 읽은 것)
def _near_phrase(band, rx, anchor, maxdx=None):
    best, bd = None, None
    for l in band:
        for p in _phrases(l, rx):
            dx = abs(_cx(p) - _cx(anchor))
            if maxdx is not None and dx > maxdx:
                continue
            d = dx + 4 * abs(_cy(p) - _cy(anchor))
            if bd is None or d < bd:
                best, bd = p, d
    return best


def _scan(lines, li, hdr, hdr_bottom, h, exact):
    """머리글 아래 줄들을 읽어 타입 줄·소계·합계를 모음"""
    loose = not exact
    body, prev = [], hdr_bottom
    for l in lines[li + 1:]:
        if l["y"] <= hdr_bottom:
            continue
        if body and (l["y"] - prev > (3.0 if exact else 4.0) * h or _is_title(l)):
            break
        body.append(l)
        prev = l["y"]
        if len(body) > 80:
            break
    cols = _columns(body)
    asg = _match_cols(hdr, cols, h)
    if "units" not in asg or "supply" not in asg:
        return None
    ucol = cols[asg["units"]]
    rows, labels, total, cand_tot = [], [], None, []
    seg, started, miss = 0, False, 0
    for l in body:
        toks = l["t"]
        vals = {k: _col_vals(l, cols[i]) for k, i in asg.items()}
        left = " ".join(t.s for t in toks if t.x1 <= ucol["x0"] + 0.25 * (ucol["x1"] - ucol["x0"]))
        if RX_SUB.search(left):
            if started:
                seg += 1
            continue
        if RX_TOTAL.search(left):
            if started:
                total = dict(y=l["y"], units=vals["units"], supply=vals["supply"])
                break
            continue
        tt = _type_tok(toks, ucol, hdr.get("type"), loose)
        if tt is None and _plain(vals):                      # 타입 칸이 번호뿐인 옛 양식
            tt = Tok(-1e9, l["y"], -1e9, l["y"], "", None)
        units, supply = vals["units"], vals["supply"]
        if exact and not any(not p[2] and p[0] >= 0 for p in units):
            units = supply = []                              # 엑셀: 세대수 칸이 빈 줄은 타입 줄이 아님(검토용 줄 등)
        if tt is None or (not units and not supply):
            if started:
                lab = _hangul_label(toks, ucol["x0"])
                if lab:
                    labels.append((l["y"], seg, lab))
                elif units and supply:
                    cand_tot.append(dict(y=l["y"], units=units, supply=supply))   # 이름 없는 합계 줄일 수 있음
            miss += 1
            if (started and miss >= 3) or (not started and miss >= 8):
                break
            continue
        started, miss = True, 0
        lab = _hangul_label(toks, tt.x0)
        if lab:
            labels.append((l["y"], seg, lab))
        nums = [p[0] for t in toks if t is not tt for p in [_num_parts(t)] if p is not None]
        rows.append(dict(y=l["y"], type=tt.s, seg=seg, units=units, supply=supply, excl=vals.get("excl", []), nums=nums))
    if not rows:
        return None
    scol = cols[asg["supply"]]
    return dict(anchors=dict(units=(ucol["x0"] + ucol["x1"]) / 2, supply=(scol["x0"] + scol["x1"]) / 2),
                rows=rows, labels=labels, total=total, cand_tot=cand_tot, h=h,
                top=hdr_bottom, bottom=(total or rows[-1])["y"])


def _tables(toks, exact=False):
    """단어(셀) 목록 → 면적표 후보들. exact=엑셀(값이 정확 — 세대수 없는 줄은 타입 줄이 아님)"""
    lines, h = _lines(toks, join=not exact)
    out, seen = [], set()
    for li, ln in enumerate(lines):
        band = [l for l in lines if abs(l["y"] - ln["y"]) <= 2.6 * h]
        ups = _phrases(ln, RX_UNITS)
        if not ups and _phrases(ln, RX_SUPPLY) and not any(_phrases(l, RX_UNITS) for l in band):
            ups = _phrases(ln, RX_UNITS2)              # 머리글이 '세대' 뿐인 표
        for units in ups:
            # 세대수 머리글마다 가장 가까운 '공급면적'과 '분양면적'(공급면적 머리글이 없는 표)을 둘 다 후보로 — 점수로 고름
            for sp in dict.fromkeys(p for p in (_near_phrase(band, RX_SUPPLY, units, 80 * h),
                                                _near_phrase(band, RX_SALE, units, 80 * h)) if p is not None):
                key = (round(units.x0, 1), round(units.y0, 1), round(sp.x0, 1), round(sp.y0, 1))
                if key in seen:
                    continue
                seen.add(key)
                excl = _near_phrase(band, RX_EXCL, sp)
                tys = [p for l in band for p in _phrases(l, RX_TYPE) if _cx(p) < _cx(units)]
                hdr = dict(units=units, supply=sp, excl=excl, type=max(tys, key=_cx) if tys else None)
                tab = _scan(lines, li, hdr, max(units.y1, sp.y1), h, exact)
                if tab:
                    tab["hdr"] = (min(units.y0, sp.y0), units, sp)
                    out.append(tab)
    return out


# ───────────────────────────── 값 바로잡기·다수결
def _fix_dec(p, ok):
    """OCR 이 소수점을 빠뜨린 숫자(849343) → ok 를 만족하는 자리에 점을 찍음"""
    v, digits, dot = p
    if dot or digits is None:
        return v if ok(v) else None
    for k in (2, 3, 1):
        if len(digits) > k:
            x = float(digits[:k] + "." + digits[k:])
            if ok(x):
                return x
    return v if ok(v) else None


def _lead_nums(s, loose):
    m = re.match(r"0*(\d+)", _type_core(s, loose))
    if not m:
        return []
    d = m.group(1)
    return [int(d[:k]) for k in (2, 3) if len(d) >= k]


def _fix_type(raw, excl, loose):
    """타입 이름 다듬기: '84㎡A'→'84A', '059.9800A'→'59A', (OCR) '84m1A'→'84A', '948T'→'94BT', '840'→'84D'"""
    u = _type_core(raw, loose)
    m = re.fullmatch(r"0*(\d+)(?:\.\d+)?([A-Z0-9\-]*)", u)
    if not m:
        return None
    digits, rest = m.group(1), m.group(2)
    n = None
    if excl is not None:
        for k in (2, 3):
            if len(digits) >= k and abs(int(digits[:k]) - excl) < 1.6:
                n, rest = digits[:k], digits[k:] + rest
                break
    if n is None:                                       # 전용면적을 모를 때: 100~330 이면 세 자리(101A), 아니면 두 자리
        k = 3 if len(digits) >= 3 and 100 <= int(digits[:3]) <= EXCL_MAX else 2
        n, rest = digits[:k], digits[k:] + rest
    if loose:
        rest = OCR_UNIT_JUNK.sub("", rest)
        if rest and rest[0].isdigit() and (len(rest) == 1 or not rest[1].isdigit()):
            rest = DIGIT_AS_LETTER.get(rest[0], rest[0]) + rest[1:]
    rest = rest.strip("-")
    return str(int(n)) + rest


def _vote(cands):
    """[(값, 무게)] → (가장 무거운 값, 그 비율)"""
    tally, first = collections.defaultdict(float), {}
    for v, w in cands:
        if v is None:
            continue
        key = round(v, 3) if isinstance(v, float) else v
        tally[key] += w
        if key not in first:
            first[key] = v
    if not tally:
        return None, 0.0
    key, w = max(tally.items(), key=lambda kv: kv[1])
    return first[key], w / sum(tally.values())


def _resolve(passes, ocr):
    """[(무게, 표)] — 같은 표를 여러 번 읽은 것 → 줄을 y 로 맞추고 칸마다 다수결"""
    allrows = [(r["y"], pi, w, r) for pi, (w, tab) in enumerate(passes) for r in tab["rows"]]
    ref = max((t for w, t in passes), key=lambda t: len(t["rows"]))
    ys = [r["y"] for r in ref["rows"]]
    pitch = _median([b - a for a, b in zip(ys, ys[1:])], 2 * ref["h"]) if len(ys) > 1 else 2 * ref["h"]
    tol = 0.45 * pitch
    clusters = []
    for y, pi, w, r in sorted(allrows, key=lambda x: x[0]):
        for c in reversed(clusters[-3:]):
            if abs(c["y"] - y) <= tol and pi not in c["p"]:
                c["m"].append((pi, w, r))
                c["p"].add(pi)
                c["y"] = sum(m[2]["y"] for m in c["m"]) / len(c["m"])
                break
        else:
            clusters.append(dict(y=y, m=[(pi, w, r)], p={pi}))
    rows = []
    for c in clusters:
        flags = []
        ex = []                                            # 전용면적: 타입 숫자에 맞는 값
        for pi, w, r in c["m"]:
            ns = _lead_nums(r["type"], ocr)
            ok = (lambda x, ns=ns: EXCL_MIN <= x <= EXCL_MAX and (not ns or any(abs(x - n) < 1.6 for n in ns)))
            for p in r["excl"][:2]:
                x = _fix_dec(p, ok) if ocr else (p[0] if ok(p[0]) else None)
                if x is not None:
                    ex.append((x, w))
                    break
        excl, ag_e = _vote(ex)
        if excl is None:                                   # 타입 숫자와 안 맞아도 범위 안이면 받음(송산 '33평'형 등)
            for pi, w, r in c["m"]:
                for p in r["excl"][:2]:
                    x = _fix_dec(p, lambda x: EXCL_MIN <= x <= EXCL_MAX)
                    if x is not None:
                        ex.append((x, w))
                        break
            excl, ag_e = _vote(ex)
        name, ag_t = _vote([(_fix_type(r["type"], excl, ocr), w) for pi, w, r in c["m"]])
        lo, hi = (excl * 1.04, excl * 2.0) if excl else (20, 450)
        nums = [v for pi, w, r in c["m"] for v in r["nums"]]
        sp = []
        for pi, w, r in c["m"]:
            for p in r["supply"]:
                x = _fix_dec(p, lambda x: lo <= x <= hi) if ocr else (p[0] if lo <= p[0] <= hi else None)
                if x is not None:
                    bonus = 0.6 if excl and any(abs(excl + v - x) <= 0.002 for v in nums) else 0
                    sp.append((x, w + bonus))
                    break
        supply, ag_s = _vote(sp)
        us = []
        for pi, w, r in c["m"]:
            for p in r["units"]:
                if not p[2] and abs(p[0] - round(p[0])) < 1e-9 and 0 <= p[0] < 100000:
                    us.append((int(round(p[0])), w))
                    break
        units, ag_u = _vote(us)
        if excl and supply and not any(abs(excl + v - supply) <= 0.002 for v in nums):
            flags.append("공급≠전용+공용")
        if ocr and len(passes) > 1 and min(a for a in (ag_t, ag_u, ag_s) if a) < 0.6:
            flags.append("판독 불일치")
        seg = collections.Counter(r["seg"] for pi, w, r in c["m"]).most_common(1)[0][0]
        rows.append(dict(y=c["y"], name=name, units=units, excl=excl, supply=supply, seg=seg, flags=flags,
                         reads=len(c["m"])))
    # 그룹 이름(이익공유·통합공공임대 …): 가까운 조각끼리 한 덩어리로(세로 병합 칸이 '이익'·'공유'로 갈라진 것),
    # 소계로 나뉜 구간에 덩어리가 하나면 구간 전체, 여럿이면 줄을 품은 덩어리 → 위쪽 → 아래쪽 순
    labs = sorted((y, seg, s) for w, t in passes[:1] for (y, seg, s) in t["labels"] if not RX_SCALE.search(s))
    blocks = []
    for y, seg, s in labs:
        if blocks and blocks[-1]["seg"] == seg and y - blocks[-1]["y1"] <= 1.5 * pitch:
            blocks[-1]["y1"] = y
            blocks[-1]["s"] += " " + s
        else:
            blocks.append(dict(seg=seg, y0=y, y1=y, s=s))
    for r in rows:
        bs = [b for b in blocks if b["seg"] == r["seg"]]
        pick = None
        if len(bs) == 1:
            pick = bs[0]
        elif bs:
            inside = [b for b in bs if b["y0"] - pitch <= r["y"] <= b["y1"] + pitch]
            above = [b for b in bs if b["y1"] < r["y"]]
            pick = min(inside, key=lambda b: abs((b["y0"] + b["y1"]) / 2 - r["y"])) if inside else \
                (max(above, key=lambda b: b["y1"]) if above else min(bs, key=lambda b: b["y0"]))
        r["label"] = " ".join(pick["s"].split()) if pick else ""
    # 합계 줄: 이름 있는 것 먼저, 없으면 세대수 합과 맞는 이름 없는 줄
    sum_u = sum(r["units"] or 0 for r in rows)
    sum_a = sum((r["units"] or 0) * (r["supply"] or 0) for r in rows)
    tu, ts = [], []
    for w, t in passes:
        tot = t.get("total")
        if not tot:
            tot = next((c for c in t.get("cand_tot") or []
                        if any(not p[2] and abs(p[0] - sum_u) <= max(2, 0.02 * sum_u) for p in c["units"])), None)
        if not tot:
            continue
        for p in tot.get("units") or []:
            if not p[2] and p[0] >= 1:
                tu.append((int(round(p[0])), w))
                break
        for p in tot.get("supply") or []:
            x = _fix_dec(p, lambda x: x > 100 and (not sum_a or abs(x - sum_a) <= 0.05 * sum_a)) if ocr else p[0]
            if x is not None and x > 100:
                ts.append((x, w))
                break
    return rows, _vote(tu)[0], _vote(ts)[0]


# ───────────────────────────── 마무리(검증·정리)
def _finalize(rows, tot_units, tot_supply, meta, max_types):
    notes = []
    rows = [r for r in rows if r["units"] != 0 and (r["name"] or r["supply"])]
    miss = [r for r in rows if r["units"] is None]
    if len(miss) == 1 and tot_units:
        rest = tot_units - sum(r["units"] for r in rows if r["units"] is not None)
        if rest > 0:
            miss[0]["units"] = rest
            miss[0]["flags"].append("세대수=합계에서 역산")
    for r in rows:                                        # 타입 이름이 없거나 전용면적과 안 맞으면 전용면적으로 이름 짓기
        if not r["excl"]:
            continue
        lead = int(re.match(r"\d+", r["name"]).group()) if r["name"] and re.match(r"\d+", r["name"]) else None
        if lead is not None and abs(lead - r["excl"]) < 1.6:
            continue
        base = str(int(r["excl"]))
        suffix = re.sub(r"^\d+", "", r["name"] or "")
        if not suffix:
            letters = [x["name"][len(base)] for x in rows if x is not r and x["name"] and x["name"].startswith(base)
                       and len(x["name"]) > len(base) and x["name"][len(base)].isalpha()]
            suffix = chr(max(map(ord, letters)) + 1) if letters else "A"
        r["name"] = base + suffix
        r["flags"].append("타입 이름 추정")
    for r in rows:
        r["use"] = "임대" not in (r.get("label") or "")
    sum_units = sum(r["units"] or 0 for r in rows)
    units_ok = (sum_units == tot_units) if tot_units else None
    sum_area = sum((r["units"] or 0) * (r["supply"] or 0) for r in rows)
    supply_ok = (abs(sum_area - tot_supply) <= max(1.0, tot_supply * 0.0005)) if tot_supply and sum_area else None
    if units_ok is False:
        notes.append(f"세대수 합 {sum_units:,} ≠ 표 합계 {tot_units:,}")
    if supply_ok is False:
        notes.append(f"공급면적×세대수 합 {sum_area:,.2f} ≠ 표 합계 {tot_supply:,.2f}")
    # 같은 이름·같은 면적(그룹만 다름)은 합침
    out = []
    for r in rows:
        same = next((x for x in out if x["use"] == r["use"] and x["name"] == r["name"]
                     and abs((x["supply"] or 0) - (r["supply"] or 0)) <= 0.05), None)
        if same:
            same["units"] = (same["units"] or 0) + (r["units"] or 0)
            same["label"] = " · ".join(dict.fromkeys(x for x in (same["label"], r["label"]) if x))
            same["flags"] = list(dict.fromkeys(same["flags"] + r["flags"]))
            continue
        out.append(dict(r))
    names = collections.Counter(x["name"] for x in out if x["use"])
    k = collections.Counter()
    for x in out:                                      # 이름이 겹치면: 숫자뿐인 이름(59·59·59)은 59A·59B·59C, 아니면 (2)
        if x["use"] and names[x["name"]] > 1:
            base = x["name"]
            k[base] += 1
            if base.isdigit():
                x["name"] = base + chr(64 + k[base])
            elif k[base] > 1:
                x["name"] = f"{base}({k[base]})"
    used = [x for x in out if x["use"]]
    if max_types and len(used) > max_types:
        notes.append(f"타입이 {len(used)}개라 세대 구성 칸({max_types}줄)에 맞게 같은 평형의 작은 타입을 합쳤습니다")
        while len(used) > max_types:
            grp = collections.defaultdict(list)
            for x in used:
                grp[re.match(r"\d+", x["name"]).group()].append(x)
            g = max((g for g in grp.values() if len(g) > 1), key=len, default=None)
            if g is None:
                break
            a, b = sorted(g, key=lambda x: x["units"] or 0)[:2]
            n = (a["units"] or 0) + (b["units"] or 0)
            b["supply"] = round(((a["supply"] or 0) * (a["units"] or 0) + (b["supply"] or 0) * (b["units"] or 0)) / n, 4) if n else b["supply"]
            b["name"] = b["name"] + "·" + re.sub(r"^\d+", "", a["name"])
            b["units"] = n
            b["flags"].append("작은 타입 합침")
            used.remove(a)
            out.remove(a)
    types = [dict(name=x["name"], units=x["units"], supply=round(x["supply"], 4) if x["supply"] else None,
                  excl=round(x["excl"], 4) if x["excl"] else None, label=x["label"], use=x["use"], flags=x["flags"])
             for x in out]
    nflag = sum(1 for t in types if t["flags"] and t["use"])
    if units_ok and (supply_ok or (supply_ok is None and not nflag)):
        conf = "high"
    elif units_ok or supply_ok or (units_ok is None and supply_ok is None and not nflag and len(types) >= 2):
        conf = "mid"
    else:
        conf = "low"
    return dict(types=types, total_units=tot_units, total_supply=tot_supply, sum_units=sum_units,
                units_ok=units_ok, supply_ok=supply_ok, confidence=conf, notes=notes, **meta)


def _score(res):
    good = [t for t in res["types"] if t["units"] and t["supply"]]
    s = 10 * len(good) - 4 * sum(1 for t in res["types"] if t["flags"])
    s += 60 if res["units_ok"] else (-20 if res["units_ok"] is False else 0)
    s += 30 if res["supply_ok"] else 0
    return s + res.get("bonus", 0)


# ───────────────────────────── 엑셀
def _nstr(v):
    try:
        return str(int(v)) if float(v) == int(v) else repr(float(v))
    except (ValueError, OverflowError):
        return repr(v)


def _excel_sheets(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        import xlrd
        wb = xlrd.open_workbook(path, formatting_info=True)
        for sh in wb.sheets():
            span = {(rl, cl): ch for rl, rh, cl, ch in sh.merged_cells}
            toks = []
            for r in range(sh.nrows):
                for c in range(sh.ncols):
                    v = sh.cell_value(r, c)
                    if v in (None, ""):
                        continue
                    c1 = span.get((r, c), c + 1)
                    if isinstance(v, float):
                        toks.append(Tok(c, r, c1, r + 1, _nstr(v), v))
                    else:
                        toks.append(Tok(c, r, c1, r + 1, " ".join(str(v).split()), None))
            yield sh.name, sh.visibility == 0, toks
        return
    import zipfile
    import openpyxl
    from openpyxl.utils import column_index_from_string as col_i
    # 값만 읽는 모드(서식까지 읽으면 5배 느림) + 병합 칸은 시트 XML 에서 직접
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    zf = zipfile.ZipFile(path)
    try:
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            if not any(isinstance(v, str) and "세" in v and "대" in v for row in rows for v in row):
                continue                                   # 세대수 머리글이 없는 시트는 건너뜀
            span = {}
            try:
                xml = zf.read(ws._worksheet_path.lstrip("/")).decode("utf-8", "ignore")
                for m in re.finditer(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', xml):
                    span[(int(m.group(2)), col_i(m.group(1)))] = col_i(m.group(3)) + 1
            except (KeyError, AttributeError):
                pass
            toks = []
            for r, row in enumerate(rows, 1):
                for c, v in enumerate(row, 1):
                    if v is None or v == "" or isinstance(v, bool) or hasattr(v, "year"):
                        continue
                    c1 = span.get((r, c), c + 1)
                    if isinstance(v, (int, float)):
                        toks.append(Tok(c, r, c1, r + 1, _nstr(v), float(v)))
                    else:
                        toks.append(Tok(c, r, c1, r + 1, " ".join(str(v).split()), None))
            yield ws.title, ws.sheet_state == "visible", toks
    finally:
        wb.close()
        zf.close()


def _excel_results(path, max_types):
    out = []
    for name, visible, toks in _excel_sheets(path):
        for tab in _tables(toks, exact=True):
            rows, tu, ts = _resolve([(1.0, tab)], ocr=False)
            res = _finalize(rows, tu, ts, dict(source="엑셀", where=f"시트 '{name}'"), max_types)
            res["bonus"] = (8 if visible else 0) + (4 if "개요" in name else 0)
            if res["types"]:
                out.append(res)
    return out


# ───────────────────────────── PDF
def _fitz():
    import site
    site.addsitedir(site.getusersitepackages())
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    return fitz


def _text_dir(page):
    """글자 방향(가장 많은 줄의 방향) — 90° 돌아간 도면 대응"""
    cnt = collections.Counter()
    for b in page.get_text("dict").get("blocks", []):
        for ln in b.get("lines", []):
            txt = "".join(s.get("text", "") for s in ln.get("spans", []))
            if re.search(r"[가-힣0-9]", txt):
                dx, dy = ln.get("dir", (1, 0))
                cnt[(round(dx), round(dy))] += len(txt)
    return cnt.most_common(1)[0][0] if cnt else (1, 0)


def _turn(x0, y0, x1, y1, d):
    """글자가 왼→오로 읽히게 좌표 돌리기: X = p·dir, Y = p·(dir 를 시계방향 90°)"""
    dx, dy = d
    pts = [(x0, y0), (x1, y0), (x0, y1), (x1, y1)]
    X = [x * dx + y * dy for x, y in pts]
    Y = [-x * dy + y * dx for x, y in pts]
    return min(X), min(Y), max(X), max(Y)


def _pdf_text_results(doc, max_types):
    out = []
    for p in list(doc)[:12]:
        words = p.get_text("words")
        txt = "".join(w[4] for w in words)
        if "세" not in txt or "대" not in txt:
            continue
        d = _text_dir(p)
        toks = [Tok(*_turn(w[0], w[1], w[2], w[3], d), w[4], None) for w in words if w[4].strip()]
        for tab in _tables(toks):
            rows, tu, ts = _resolve([(1.0, tab)], ocr=False)
            res = _finalize(rows, tu, ts, dict(source="PDF", where=f"{p.number + 1}쪽"), max_types)
            if res["types"]:
                res["bonus"] = 0
                out.append(res)
    return out


def _ocr_files(images, timeout=300):
    """그림 파일들 → [dict(w, h, words=[[x, y, w, h, 글자], …])]"""
    td = os.path.dirname(images[0])
    lst, out = os.path.join(td, "list.txt"), os.path.join(td, "ocr.jsonl")
    with open(lst, "w", encoding="utf-8") as f:
        f.write("\n".join(images) + "\n")
    try:
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", OCR_PS1, "-List", lst, "-Out", out],
                       check=True, timeout=timeout, capture_output=True,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (subprocess.SubprocessError, OSError) as e:
        err = getattr(e, "stderr", b"") or b""
        raise ValueError(f"그림 PDF 글자 인식(Windows OCR)에 실패했습니다: {err.decode('cp949', 'replace')[:300] or e}")
    res = []
    with open(out, encoding="utf-8-sig") as f:
        for line in f:
            if line.strip():
                res.append(json.loads(line))
    if res and res[0].get("error") == "no-ocr-language":
        raise ValueError("이 PC 의 Windows 에 한국어 글자 인식(OCR)이 없어 그림 PDF 를 읽을 수 없습니다. 엑셀 설계개요를 넣어 주세요.")
    return res


def _render(fitz, page, zoom, rot, clip, path):
    m = fitz.Matrix(zoom, zoom).prerotate(rot)
    pix = page.get_pixmap(matrix=m, clip=clip, colorspace=fitz.csGRAY, alpha=False)
    pix.save(path)
    return pix.x, pix.y


def _ocr_toks(res, ox, oy, zoom):
    return [Tok((ox + x) / zoom, (oy + y) / zoom, (ox + x + w) / zoom, (oy + y + h) / zoom, s, None)
            for x, y, w, h, s in res.get("words") or [] if str(s).strip()]


def _pdf_ocr_results(fitz, doc, max_types, log):
    pages = list(doc)[:12]
    pri = [p for p in pages if re.search(r"설\s*계\s*개\s*요|면\s*적\s*표", p.get_text())]
    pages = (pri or pages)[:4]
    base = {}
    for p in pages:                    # 쪽 회전(/Rotate)은 직접 돌려서 그림: 처음엔 보이는 방향, 못 찾으면 ±90°
        base[p.number] = p.rotation % 360
        if p.rotation:
            p.set_rotation(0)
    td = tempfile.mkdtemp(prefix="dov_")
    try:
        found = []
        for turns in ((0,), (90, 270)):
            jobs = []
            for p in pages:
                W = max(p.rect.width, p.rect.height)
                z = min(3.0, 9000 / W)
                for rot in ((base[p.number] + t) % 360 for t in turns):
                    fn = os.path.join(td, f"a_{p.number}_{rot}.png")
                    ox, oy = _render(fitz, p, z, rot, None, fn)
                    jobs.append((p, rot, z, ox, oy, fn))
            t = time.time()
            res = _ocr_files([j[-1] for j in jobs])
            log(f"OCR 1차 {len(jobs)}장 {time.time() - t:.1f}초")
            for (p, rot, z, ox, oy, fn), r in zip(jobs, res):
                for tab in _tables(_ocr_toks(r, ox, oy, z)):
                    rows, tu, ts = _resolve([(1.0, tab)], ocr=True)
                    fin = _finalize(rows, tu, ts, {}, max_types)
                    found.append((_score(fin), fin, p, rot, z, tab))
            if found:
                break
        if not found:
            return []
        found.sort(key=lambda x: -x[0])
        sc, fin1, p, rot, z0, tabA = found[0]
        meta = dict(source="PDF 그림(글자 인식)", where=f"{p.number + 1}쪽")
        if fin1["units_ok"] and fin1["supply_ok"]:          # 한 번에 합계까지 맞으면 그대로
            fin1.update(meta, reads=1, bonus=0)
            return [fin1]
        # 표 부분만 크게 다시 읽기(2번) — 좌표는 1차와 같은 기준(pt, 돌린 그림 기준)
        h = tabA["h"]
        top = tabA["hdr"][0] - 3 * h
        bot = tabA["bottom"] + 3 * h
        fr = fitz.Rect(-1e5, top, 1e5, bot) & (p.rect * fitz.Matrix(rot))
        clip = (fr * ~fitz.Matrix(rot)) & p.rect
        jobs = []
        for i, z in enumerate((4.2, 5.5)):
            z = min(z, 9500 / max(fr.width, fr.height))
            fn = os.path.join(td, f"b_{i}.png")
            ox, oy = _render(fitz, p, z, rot, clip, fn)
            jobs.append((z, ox, oy, fn))
        t = time.time()
        res = _ocr_files([j[-1] for j in jobs])
        log(f"OCR 2차 {len(jobs)}장 {time.time() - t:.1f}초")
        passes = [(1.0, tabA)]
        ha = tabA["hdr"]
        for (z, ox, oy, fn), r, w in zip(jobs, res, (1.15, 1.3)):
            tabs = _tables(_ocr_toks(r, ox, oy, z))
            near = [t for t in tabs if abs(t["hdr"][0] - ha[0]) < 3 * h
                    and abs(t["anchors"]["units"] - tabA["anchors"]["units"]) < 6 * h
                    and abs(t["anchors"]["supply"] - tabA["anchors"]["supply"]) < 6 * h]
            if near:
                passes.append((w, max(near, key=lambda t: len(t["rows"]))))
        rows, tu, ts = _resolve(passes, ocr=True)
        res = _finalize(rows, tu, ts, dict(meta, reads=len(passes)), max_types)
        res["bonus"] = 0
        return [res] if res["types"] else []
    finally:
        for f in os.listdir(td):
            try:
                os.remove(os.path.join(td, f))
            except OSError:
                pass
        try:
            os.rmdir(td)
        except OSError:
            pass


# ───────────────────────────── 밖에서 부르는 것
def read_types(path, max_types=20, ocr=True, log=None):
    """설계개요 파일 → 세대 구성. 못 찾으면 ValueError"""
    log = log or (lambda s: None)
    t0 = time.time()
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        cands = _excel_results(path, max_types)
    elif ext == ".pdf":
        fitz = _fitz()
        doc = fitz.open(path)
        cands = _pdf_text_results(doc, max_types)
        if not cands and ocr:
            cands = _pdf_ocr_results(fitz, doc, max_types, log)
    else:
        raise ValueError("설계개요는 엑셀(xlsx·xls) 또는 PDF 파일이어야 합니다.")
    if not cands:
        raise ValueError("설계개요에서 '세대수·공급면적' 표(평형별 면적표)를 찾지 못했습니다.")
    best = max(cands, key=_score)
    best.pop("bonus", None)
    best["sec"] = round(time.time() - t0, 1)
    return best


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    for path in sys.argv[1:]:
        try:
            r = read_types(path, log=print)
        except Exception as e:
            print(f"■ {os.path.basename(path)} → 실패: {e}")
            continue
        print(f"■ {os.path.basename(path)} → {r['source']} {r['where']} · {r['sec']}초 · 신뢰 {r['confidence']} · "
              f"합계 세대 {r['total_units']} (읽은 합 {r['sum_units']}, 일치 {r['units_ok']}) · 공급면적 합계 일치 {r['supply_ok']}")
        for t in r["types"]:
            print(f"   {'  ' if t['use'] else '× '}{t['name']:<8} {t['units'] if t['units'] is not None else '?':>5}세대  "
                  f"전용 {t['excl']}  공급 {t['supply']}  {t['label']}  {' / '.join(t['flags'])}")
        for n in r["notes"]:
            print("   ※", n)


if __name__ == "__main__":
    main()
