# -*- coding: utf-8 -*-
"""스캔 PDF 지출계획서 → Windows OCR → 표 복원(엑셀 파서와 같은 rec 구조). 열은 숫자 위치로 찾고, 합계 열·합계 줄로 검산."""
import io
import json
import os
import re
import sys
import warnings

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
SP = os.path.dirname(os.path.abspath(__file__))
ROOT = r"E:\1. 우미\1. 본사업무\★ 판관비 예산안\☆ 예산안 자동작성"
sys.path.insert(0, os.path.join(ROOT, "tool"))
import pymupdf                      # noqa: E402
import design_overview as dov       # noqa: E402

ZOOM = 3
CACHE = os.path.join(SP, "ocr_cache")
os.makedirs(CACHE, exist_ok=True)
GROUP_KEYS = [("복리후생비", ["복리", "후생", "복후"]), ("소모품비", ["소모", "모품", "소모꿈"]), ("지급임차료", ["임차", "임자", "차료"]),
              ("지급수수료", ["수수", "지수"]), ("도서인쇄비", ["도서", "인쇄"]), ("수도광열비", ["수도", "광열", "광옅"]), ("통신비", ["통신"]),
              ("운반비", ["운반"]), ("여비교통비", ["여비", "교통"]), ("관리비", ["관리비"]), ("보험료", ["보험"])]
SUB_ORDER = {"복리후생비": ["식대", "숙박비", "숙소관리비"], "소모품비": ["사무용가구", "소모품"], "지급임차료": ["집기비품"],
             "지급수수료": ["발급수수료", "보안경비용역", "용역-기타", "기타"]}
SUB_KEYS = [("식대", ["식대"]), ("숙박비", ["숙박", "숙소비", "숙소임차"]), ("숙소관리비", ["숙소관리", "관리비"]), ("사무용가구", ["가구"]),
            ("소모품", ["소모"]), ("집기비품", ["집기", "비품"]), ("발급수수료", ["발급"]), ("송금수수료", ["송금"]), ("보안경비용역", ["보안", "경비"]),
            ("용역-기타", ["용역"]), ("기타", ["기타"]), ("통신비", ["통신"]), ("우편료", ["우편"]), ("운반비", ["운반"]), ("여비교통비", ["여비", "교통"]),
            ("수도광열비", ["수도", "광열", "광옅"]), ("도서인쇄비", ["도서", "인쇄"]), ("관리비", ["관리"]), ("보험료", ["보험"]), ("전기료", ["전기"])]
ITEM_OF_SUB = {"식대": "복리후생비", "숙박비": "복리후생비", "숙소관리비": "복리후생비", "사무용가구": "소모품비", "소모품": "소모품비",
               "집기비품": "지급임차료", "발급수수료": "지급수수료", "송금수수료": "지급수수료", "보안경비용역": "지급수수료", "용역-기타": "지급수수료",
               "통신비": "통신비", "우편료": "통신비", "운반비": "운반비", "여비교통비": "여비교통비", "수도광열비": "수도광열비",
               "도서인쇄비": "도서인쇄비", "보험료": "보험료", "전기료": "수도광열비"}


def ocr_words(path, page_no=0):
    key = re.sub(r"[^\w가-힣]", "_", os.path.basename(path)) + f"_{page_no}.json"
    cp = os.path.join(CACHE, key)
    if os.path.exists(cp):
        return json.load(io.open(cp, encoding="utf-8"))
    doc = pymupdf.open(path)
    page = doc[page_no]
    pix = page.get_pixmap(matrix=pymupdf.Matrix(ZOOM, ZOOM), colorspace=pymupdf.csGRAY, alpha=False)
    res = dov.ocr_windows([pix.tobytes("png")])[0]
    words = [[float(x), float(y), float(w), float(h), str(t)] for x, y, w, h, t in (res.get("words") or [])]
    json.dump(words, io.open(cp, "w", encoding="utf-8"), ensure_ascii=False)
    return words


NUM_FIX = str.maketrans({"이": "0", "O": "0", "o": "0", "Q": "0", "D": "0", "l": "1", "I": "1", "|": "1", "ㅣ": "1", "，": ",", "．": ".", "기": "7", "S": "5"})


def numeric_like(t):
    s = t.strip().translate(NUM_FIX)
    return bool(s) and bool(re.fullmatch(r"[\d.,]+", s)) and any(ch.isdigit() for ch in s)


def as_number(t):
    """OCR 글 → 정수(원). 숫자가 아니면 None. '-' 는 0"""
    s = t.strip()
    if s in ("-", "—", "–", "ㅡ"):
        return 0
    s2 = s.translate(NUM_FIX)
    digits = re.sub(r"\D", "", s2)
    if len(digits) < 3:
        return None
    other = re.sub(r"[\d.,\s]", "", s2)
    if len(other) > max(1, len(s2) // 4):
        return None
    return int(digits)


def norm(s):
    return re.sub(r"\s+", "", s or "")


def group_name(raw):
    for name, keys in GROUP_KEYS:
        if any(k in raw for k in keys):
            return name
    return raw


def sub_name(raw, group, order_idx):
    for name, keys in SUB_KEYS:
        if any(k in raw for k in keys):
            if name == "숙소관리비" and group not in ("복리후생비", "") and "숙소" not in raw:
                continue
            return name
    if ("식" in raw and "대" in raw) or raw in ("식", "대"):
        return "식대"
    lst = SUB_ORDER.get(group) or []
    if order_idx < len(lst):
        return lst[order_idx]
    return raw or group


def clusters(xs, gap):
    xs = sorted(xs)
    out = []
    for x in xs:
        if out and x - out[-1][-1] <= gap:
            out[-1].append(x)
        else:
            out.append([x])
    return [(sum(c) / len(c), len(c)) for c in out]


def parse_pdf(path, page_no=0):
    words = ocr_words(path, page_no)
    if not words:
        return dict(err="OCR 단어 없음")
    W = [dict(x=x, y=y, w=w, h=h, cx=x + w / 2, cy=y + h / 2, t=t) for x, y, w, h, t in words]
    W.sort(key=lambda w: (round(w["cy"] / 14), w["x"]))
    merged = []
    for w in W:
        if merged:
            p = merged[-1]
            if abs(p["cy"] - w["cy"]) < 14 and 0 <= w["x"] - (p["x"] + p["w"]) < 22 and numeric_like(p["t"]) and numeric_like(w["t"]):
                p["t"] += w["t"]
                p["w"] = w["x"] + w["w"] - p["x"]
                p["cx"] = p["x"] + p["w"] / 2
                continue
        merged.append(dict(w))
    W = merged
    # 제목·작성일
    date_w = next((w for w in W if "작성일" in w["t"]), None)
    date_cy = date_w["cy"] if date_w else 260
    title_words = sorted([w for w in W if w["cy"] < date_cy - 15], key=lambda w: w["x"])
    title = " ".join(w["t"] for w in title_words if w["t"] not in ("(", ")"))
    date = None
    if date_w:
        line = " ".join(v["t"] for v in sorted(W, key=lambda v: v["x"]) if abs(v["cy"] - date_w["cy"]) < 20)
        m = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일", line)
        if m:
            date = f"( 작성일 : {m.group(1)}년 {int(m.group(2)):02d}월 {int(m.group(3)):02d}일 )"
    # 표 머리글 줄('예상비용 지출계획' 또는 '합계'/'비고')
    hd = [w for w in W if w["t"] in ("예", "예상", "예상비용") and w["cy"] > date_cy + 40 and 500 < w["cx"] < 900]
    hap = sorted([w for w in W if w["t"] in ("합", "합계") and w["cy"] > date_cy + 40 and 300 < w["cx"] < 700], key=lambda w: w["cy"])
    bigo_all = [w for w in W if w["t"] in ("비", "비고") and w["cy"] > date_cy + 40 and w["cx"] > 1000]
    if hd:
        hd_cy = min(w["cy"] for w in hd)
    elif hap:
        hd_cy = hap[0]["cy"] - 40
    elif bigo_all:
        hd_cy = min(w["cy"] for w in bigo_all) - 40
    else:
        return dict(err="표 머리글 없음")
    bigo = [w for w in bigo_all if hd_cy - 20 < w["cy"] < hd_cy + 120]
    bigo_cx = bigo[0]["cx"] if bigo else 1400
    # 열: 머리글 아래 숫자 단어의 x 묶음(왼쪽부터 합계, 달…)
    nums_all = [w for w in W if w["cy"] > hd_cy + 60 and w["cx"] < bigo_cx - 60 and w["cx"] > 330 and as_number(w["t"]) is not None]
    cl = [c for c in clusters([w["cx"] for w in nums_all], 70) if c[1] >= 2]
    if len(cl) < 2:
        return dict(err="숫자 열을 찾지 못함")
    centers = [c[0] for c in cl]
    gaps = [b - a for a, b in zip(centers, centers[1:])]
    step = sorted(gaps)[len(gaps) // 2]
    # 간격이 비정상(비고 안 숫자)인 열은 버림
    keep = [centers[0]]
    for c in centers[1:]:
        if 0.6 * step <= c - keep[-1] <= 1.5 * step:
            keep.append(c)
        else:
            break
    centers = keep
    if len(centers) < 2:
        return dict(err="달 열을 찾지 못함")
    hap_cx, mcols = centers[0], centers[1:]
    # 달 이름: 머리글 아래 'N월' 단어를 열에 붙임, 없으면 이웃 달로. 머리글 달 수가 2개 이상 읽히면 열 수를 그에 맞춤
    mon_words = [w for w in W if re.fullmatch(r"'?\d{1,2}월", w["t"]) and hd_cy + 10 < w["cy"] < hd_cy + 170 and w["cx"] < bigo_cx - 100]
    hdr_cnt = len({round(w["cx"] / 60) for w in mon_words if w["cx"] > hap_cx + step * 0.5})
    if hdr_cnt >= 2 and len(mcols) > hdr_cnt:
        mcols = mcols[:hdr_cnt]
    mval = [None] * len(mcols)
    for w in mon_words:
        j = min(range(len(mcols)), key=lambda k: abs(w["cx"] - mcols[k]))
        if abs(w["cx"] - mcols[j]) < step * 0.45 and mval[j] is None:
            mval[j] = int(re.sub(r"\D", "", w["t"]))
    if all(v is None for v in mval):
        return dict(err="월 머리글 없음")
    for _ in range(len(mcols)):
        for j in range(len(mcols)):
            if mval[j] is None:
                if j > 0 and mval[j - 1] is not None:
                    mval[j] = mval[j - 1] % 12 + 1
                elif j + 1 < len(mcols) and mval[j + 1] is not None:
                    mval[j] = (mval[j + 1] - 2) % 12 + 1
    hdr_y = max((w["cy"] for w in mon_words), default=hd_cy + 90)
    months = []
    for j, cx in enumerate(mcols):
        lab = " ".join(v["t"] for v in sorted(W, key=lambda v: v["x"]) if hdr_y - 5 < v["cy"] < hdr_y + 80 and abs(v["cx"] - cx) < step * 0.5
                       and not re.fullmatch(r"'?\d{1,2}월", v["t"]) and as_number(v["t"]) is None)
        months.append(dict(y=None, m=mval[j], label=norm(lab)))
    nm = len(months)
    cols = [hap_cx] + mcols
    note_x0 = mcols[-1] + step * 0.55
    # 숫자 단어 → 줄 묶기
    nums = [dict(v=as_number(w["t"]), **w) for w in W if w["cy"] > hdr_y + 40 and w["cx"] < note_x0 and w["cx"] > hap_cx - step * 0.5]
    nums = [n for n in nums if n["v"] is not None]
    nums.sort(key=lambda n: n["cy"])
    rows_y = []
    for n in nums:
        if rows_y and abs(n["cy"] - rows_y[-1]["cy"]) < 26:
            rows_y[-1]["items"].append(n)
        else:
            rows_y.append(dict(cy=n["cy"], items=[n]))
    out_rows = []
    for r in rows_y:
        r["cy"] = sum(n["cy"] for n in r["items"]) / len(r["items"])
        vals, raw = [None] * len(cols), [None] * len(cols)
        for n in r["items"]:
            j = min(range(len(cols)), key=lambda k: abs(n["cx"] - cols[k]))
            if abs(n["cx"] - cols[j]) < step * 0.55 and vals[j] is None:
                vals[j], raw[j] = n["v"], n["t"]
        if all(v is None for v in vals):
            continue
        r["vals"], r["raw"] = vals, raw
        out_rows.append(r)
    ys = [r["cy"] for r in out_rows]
    bands = []
    for i, y in enumerate(ys):
        top = (ys[i - 1] + y) / 2 if i else y - 70
        bot = (y + ys[i + 1]) / 2 if i + 1 < len(ys) else y + 70
        bands.append((top, bot))
    text_words = [w for w in W if w["cy"] > hdr_y + 40 and as_number(w["t"]) is None]
    left_a = sorted([w for w in text_words if w["cx"] < 215], key=lambda w: w["cy"])
    groups = []
    for w in left_a:
        if groups and w["cy"] - groups[-1]["y1"] < 45:
            groups[-1]["ws"].append(w)
            groups[-1]["y1"] = w["cy"]
        else:
            groups.append(dict(y0=w["cy"], y1=w["cy"], ws=[w]))
    for g in groups:
        g["raw"] = norm("".join(v["t"] for v in g["ws"]))
        g["name"] = group_name(g["raw"])
        g["cy"] = (g["y0"] + g["y1"]) / 2
    body = []
    for r, (top, bot) in zip(out_rows, bands):
        name_ws = [w for w in text_words if 60 < w["cx"] < hap_cx - step * 0.5 and abs(w["cy"] - r["cy"]) < 42]
        raw_name = norm("".join(w["t"] for w in sorted(name_ws, key=lambda w: w["x"])))
        grp = None
        if groups:
            inside = [g for g in groups if g["y0"] - 45 <= r["cy"] <= g["y1"] + 45]
            grp = inside[0] if inside else min(groups, key=lambda g: abs(g["cy"] - r["cy"]))
        note_ws = [w for w in text_words if w["cx"] >= note_x0 and top <= w["cy"] < bot]
        lines = {}
        for w in sorted(note_ws, key=lambda w: (round(w["cy"] / 22), w["x"])):
            lines.setdefault(round(w["cy"] / 22), []).append(w["t"])
        note = "\n".join(" ".join(v) for k, v in sorted(lines.items()))
        body.append(dict(cy=r["cy"], item=grp["name"] if grp else "", raw_name=raw_name, vals=r["vals"], raw=r["raw"], note=note))
    total_row = None
    for b in body:
        if "합계" in b["raw_name"] or ("합" in b["raw_name"] and "계" in b["raw_name"] and "소" not in b["raw_name"]):
            total_row = b
    if total_row is None and body:
        total_row = body[-1]
    body = [b for b in body if b is not total_row]
    rows, bad = [], []
    grp_acc = {}
    for b in body:
        vals = [v or 0 for v in b["vals"][1:1 + nm]]
        hv = b["vals"][0]
        rn = b["raw_name"]
        idx0 = len([r for r in rows if r["item"] == b["item"]])
        acc = grp_acc.setdefault(b["item"], [0.0] * nm)
        close = idx0 >= 1 and any(vals) and all(abs(acc[i] - vals[i]) <= max(1, 0.02 * max(acc[i], vals[i])) for i in range(nm))
        is_sub = ("소계" in rn or (len(rn) <= 2 and (rn in ("소", "계", "") or (("계" in rn or "가" in rn) and ("소" in rn or "人" in rn))))) and (close or rn != "") or close
        if is_sub:
            grp_acc[b["item"]] = [0.0] * nm
            continue
        ok = hv is not None and abs(hv - sum(vals)) <= 1
        if not ok and hv is not None:
            for j in range(nm):
                other = sum(v for k, v in enumerate(vals) if k != j)
                cand = hv - other
                rj = b["raw"][1 + j]
                if cand >= 0 and (rj is None or re.search(r"[^\d,]", rj) or cand == 0):
                    vals[j] = cand
                    ok = True
                    break
        sub = sub_name(rn, b["item"], idx0)
        item = ITEM_OF_SUB.get(sub, b["item"] or "?")
        if sub == "기타" and "중개" in b["note"]:
            item = "지급수수료"
        if not ok:
            bad.append(sub)
        for i in range(nm):
            acc[i] += vals[i]
        rows.append(dict(item=item, sub=sub, vals=[float(v) for v in vals], formulas=[None] * nm, note=b["note"], ok=ok, raw_name=rn))
    total = [sum(r["vals"][i] for r in rows) for i in range(nm)]
    grand_ok = None
    tv = None
    if total_row:
        tv = total_row["vals"][1:1 + nm]
        grand_ok = all(tv[i] is not None and abs(tv[i] - total[i]) <= 1 for i in range(nm))
    return dict(sheet="OCR", title=title, date=date, months=months, rows=rows, approvers=[], total=total, ocr=True,
                bad_rows=bad, grand_ok=grand_ok, total_row=tv)


if __name__ == "__main__":
    S = r"E:\1. 우미\2. 수도권남부"
    tests = sys.argv[1:] or [
        S + r"\14. 평택고덕 패키지형 공모사업\★ 2차(A33BL)\★ 지출계획서\(평택고덕2차) 오픈비용 지출계획서(26.10~12).pdf",
        S + r"\2. 과천 지식정보타운 S8\★★ 지출계획서 관련\★지출계획서(오픈, 운영비용)\(과천지식정보) MH운영비용 지출계획서 결재본(22.01월-03월).pdf",
        S + r"\종료현장\5. 위례 2차\위례신도시1차(A3-4B)\분양비용\(위례신도시 1차) MH운용비용 지출계획서(1~3월)_결재본.pdf",
        S + r"\3. 화성비봉\지출계획서\(화성비봉) 입주비용 지출계획서('24.10_12월).pdf",
    ]
    for p in tests:
        rec = parse_pdf(p)
        print("=====", os.path.basename(p))
        if "err" in rec:
            print("  ", rec["err"])
            continue
        print("  ", rec["title"][:60], "|", rec["date"], "|", [(m["m"], m["label"]) for m in rec["months"]])
        print("   검산 실패 줄:", rec["bad_rows"], "| 합계줄 일치:", rec["grand_ok"], rec["total_row"], "계산 합", [f"{t:,.0f}" for t in rec["total"]])
        for r in rec["rows"]:
            print(f"   {'OK' if r['ok'] else 'XX'} {r['item'][:6]:6s} {r['sub'][:7]:7s} ({r['raw_name'][:8]:8s}) {[int(v) for v in r['vals']]}  | {r['note'][:60].replace(chr(10), ' / ')}")
