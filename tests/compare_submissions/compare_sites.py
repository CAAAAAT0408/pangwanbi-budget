# -*- coding: utf-8 -*-
"""과거 제출본(profiles/*.json) → 현장 조건 추정 → 연도 양식으로 도구 계산 → 항목·줄 단위 비교.
결과: compare_results.json (요약·항목별·줄별·입력값) + 콘솔 요약"""
import sys, os, re, json, glob, warnings, time, math
warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
ROOT = r"E:\1. 우미\1. 본사업무\★ 판관비 예산안\☆ 예산안 자동작성"
SP = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "tool"))
import xlcalc, site_io, build_template as BT

TEMPLATES = {y: os.path.join(SP, f"양식_{y}.xlsx") for y in (2023, 2024, 2025, 2026)}
TEMPLATES[2027] = os.path.join(ROOT, "판관비 예산안 자동작성 양식_2027.xlsx")
MODELS = {}


def model_for(year):
    if year not in MODELS:
        MODELS[year] = xlcalc.Model(TEMPLATES[year])
    return MODELS[year]


# ─────────────────────────────── 줄 분류(실제 제출본 산출근거 ↔ 도구 산출근거 공통 키)
LINE_KEYS = [
    # (키, 정규식) — 먼저 맞는 것 사용
    ("대봉투(서식대)", r"대봉투|서식대"),
    ("보존등기", r"보존등기"),
    ("이전등기", r"이전등기"),
    ("건축물관리대장", r"건축물관리대장"),
    ("입주상담사", r"입주상담사"),
    ("입주매니저", r"입주매니저"),
    ("컨테이너·텐트·난방기·에어렉스", r"컨테이너|텐트|외부난방|에어렉스"),
    ("소모품", r"소모품"),
    ("식대", r"(?<!서)식대|식 +대|^잔여판촉"),
    ("피복비", r"피복"),
    ("숙소관리비", r"숙소관리비"),
    ("숙소비(입주)", r"입주 ?여직원|상주인력 ?숙소|입주 숙소|지원인력\) 숙소"),
    ("숙소비", r"숙소비"),
    ("여비", r"유류비|출장|여비|지원인력숙소"),
    ("우편·등기·내용증명", r"우편료\(우편 기본\)|우편 기본|^우편료"),
    ("통신요금", r"통신요금|MH ?운영기간 : 600|분양홍보관 운영기간|MH 기본|입주기간 : 500|기본 : 150천원|기본 150천원|전화"),
    ("우편·등기·내용증명", r"우편|등기|내용증명|어린이집|안내문|기본 : 200천원|기본 : 50천원|기본 : 4,000천원"),
    ("전기료", r"전기료|동·하절기|동하절기"),
    ("상하수도", r"상하수도"),
    ("난방비·전기미인입", r"난방|전기 ?미인입"),
    ("복합기", r"복합기"),
    ("프린터·파쇄기", r"프린터|파쇄기"),
    ("PC", r"\bPC\b|사무실PC|상담용PC"),
    ("주차장", r"주차장"),
    ("임차료(기타)", r"임차료|임차"),
    ("보험료", r"보험"),
    ("인지세", r"인지세|75천원|제세금|세금과공과"),
    ("갑티슈", r"갑티슈|각티슈"),
    ("물티슈", r"물티슈"),
    ("크린홈", r"크린홈"),
    ("종이컵", r"종이컵"),
    ("마스크", r"마스크"),
    ("실내화", r"실내화|신발주머니"),
    ("조직분양 현수막", r"현수막|사인물|조직분양 : 40|조직분양 : 30"),
    ("대형마트", r"대형마트|부스"),
    ("대납이자", r"대납이자|중도금 대납"),
    ("분양대행·조직 수수료", r"분양대행수수료|분양대행 수수료|조직분양 수수료"),
    ("상가 분양대행", r"상가 분양대행|매출액"),
    ("분양촉진(MGM)", r"분양촉진|MGM"),
    ("오픈 선물", r"오픈 ?선물|방문객 사은품|오픈 이벤트"),
    ("청약 선물", r"청약 ?선물|청약자 사은품"),
    ("당첨자 선물", r"당첨자"),
    ("계약 선물", r"계약 ?선물|계약자 사은품"),
    ("계약해지 대체", r"계약해지"),
    ("운반비", r"운반|포장이사|이사비용"),
    ("가구", r"가구비용"),
    ("소모품", r"소모품|오픈\(D\)|오픈/입주|입주\(D\)|입주 : 4,000|오픈 : 8,000|오픈 : 4,000|열화상|기본 100천원|기본 : 100천원|기본 : 300천원|기본 300천원"),
    ("발급수수료·확정일자", r"발급|확정일자"),
    ("송금수수료", r"송금"),
    ("무인경비·보안", r"무인경비|보안경비|기본 : 300천원/월"),
    ("특별공급·청약대행", r"특별공급|특공|청약대행|청약사이트"),
    ("무순위 전산추첨", r"무순위|잔여세대 전산추첨|전산추첨 대행 : 100"),
    ("계약취소 전산추첨", r"계약취소"),
    ("전기·소방·오수 안전관리", r"전기안전|소방|오수"),
    ("분양가 산출용역", r"분양가 ?산출"),
    ("카페테리아", r"카페테리아|까페테리아"),
    ("이벤트(네일아트 등)", r"네일아트|캐리커쳐|일반행사"),
    ("네트워크 공사", r"네트워크"),
    ("방역", r"방역"),
    ("M/H 중개수수료", r"중개수수료"),
    ("청소", r"청소"),
    ("아르바이트", r"아르바이트"),
    ("상담사", r"상담사|재분양"),
    ("사전점검행사", r"사전점검행사|사검점검"),
    ("입주지원센터", r"입주지원센터"),
    ("입주촉진책", r"입주촉진"),
    ("법무사", r"법무사"),
    ("문고", r"문고"),
    ("공가세대관리비", r"공가|관리비 : 6천원|평당"),
    ("분양사무실 관리비", r"일반/공용관리비|사무실 관리비|오피스텔 관리비|보유상가 관리비|관리비"),
]
_LK = [(k, re.compile(p)) for k, p in LINE_KEYS]


def line_key(text, item=""):
    t = (text or "").replace("\n", " ")
    t = re.sub(r"^\[분양사무실\]\s*", "", t)
    for k, rx in _LK:
        if rx.search(t):
            return k
    return "(기타)" if not item else f"(기타:{item})"


# ─────────────────────────────── 현장 조건 추정
def ym_of(year, m):
    return [year, m] if m else [None, None]


def first_nonzero(vals):
    for i, v in enumerate(vals):
        if v:
            return i + 1
    return None


def last_nonzero(vals):
    last = None
    for i, v in enumerate(vals):
        if v:
            last = i + 1
    return last


def rows_by_key(pr):
    out = {}
    for r in pr.get("rows") or []:
        if r["sub"] == "소계":
            continue
        if not (r.get("total") or any(r.get("months") or [])):      # 금액 0 인 양식 잔여 줄은 조건 추정에 쓰지 않음
            continue
        k = line_key((r.get("sub") or "") + " " + (r.get("g") or ""), r.get("item") or "")
        out.setdefault(k, []).append(r)
    return out


def months_of(rows, key):
    acc = [0] * 12
    for r in rows.get(key, []):
        for i, m in enumerate(r["months"]):
            acc[i] += m or 0
    return acc


def infer_inputs(pr, override=None):
    """profile → 도구 입력(data). 사람이 현장 조건을 읽어 넣는 것과 같은 수준(세부 조정은 기본값)"""
    year = pr["year"]
    stage = pr.get("stage") or [None] * 12
    rows = rows_by_key(pr)
    rates = pr.get("rates") or {}
    sale = (rates.get("분양세대") or {}).get("months") or [0] * 12
    sale_prev = (rates.get("분양세대") or {}).get("prev") or [0] * 4
    move = (rates.get("입주세대") or {}).get("months") or [0] * 12
    move_prev = (rates.get("입주세대") or {}).get("prev") or [0] * 4
    f = {}
    kinds = []
    notes = []
    f["현장명"] = pr.get("name") or pr["file"]
    f["지역"] = "수도권"
    f["사업구분"] = "자체사업"
    f["상품"] = "아파트"

    # 오픈월(D): 단계 머리글 '오픈(D+1)' = 오픈월(그 달에 식대 오픈단가 시작), 없으면 식대 오픈단가 첫 달, 없으면 분양 첫 달-1
    open_m = None
    for i, s in enumerate(stage):
        if isinstance(s, str) and s.startswith("오픈") and "D+2" not in s and open_m is None:
            open_m = i + 1
    if open_m is None:
        for i, s in enumerate(stage):
            if isinstance(s, str) and s.strip() == "오픈" and open_m is None:
                open_m = i + 1
    pre_cnt = sum(1 for s in stage if isinstance(s, str) and s.startswith("사전") and "점검" not in s)
    mh_tel = months_of(rows, "통신요금")
    mh_on = [1 if (v and v >= 550000) else 0 for v in mh_tel]      # 600천원 → M/H 운영 달(150천원 사무실과 구분)
    if open_m is None and any(sale):
        fs = first_nonzero(sale)
        if fs and fs > 1 and any(mh_on):
            open_m = fs - 1
    is_new = open_m is not None
    # 기분양(잔여분양): M/H 운영 중인데 올해 오픈이 아님
    remain = (not is_new) and any(mh_on)
    if is_new:
        kinds.append("new")
        f["오픈"] = [year, open_m]
        f["사전영업"] = max(1, pre_cnt) if pre_cnt else 1
        # M/H 종료: 통신 600 마지막 달
        last_mh = last_nonzero(mh_on)
        f["MH종료"] = [year, last_mh] if last_mh and last_mh < 12 else [None, None]
    elif remain:
        kinds.append("remain")
        f["오픈"] = [year - 1, 6]           # 전년에 오픈한 현장으로 보고(정확한 오픈월은 모름) → D+n 단가가 모두 '기타/잔여' 구간이 되게
        f["사전영업"] = 1
        last_mh = last_nonzero(mh_on)
        first_mh = first_nonzero(mh_on)
        f["MH종료"] = [year, last_mh] if last_mh and last_mh < 12 else [None, None]
        notes.append(f"기공급(잔여분양): M/H 운영 {first_mh}~{last_mh or 12}월로 보고 전년 6월 오픈(오픈 2개월·오픈후 3개월 구간이 올해에 없게)으로 설정")
    else:
        f["오픈"] = [None, None]
    # 전년 분양 누계·월별 분양
    f["전년분양"] = sale_prev[0] if sale_prev else 0
    monthly = dict(분양=[v or None for v in sale] if any(sale) else [None] * 12,
                   입주=[v or None for v in move] if any(move) else [None] * 12,
                   전년분양=[v or None for v in sale_prev[1:4]] if any(sale_prev[1:4]) else [None, None, None],
                   전년입주=[v or None for v in move_prev[1:4]] if any(move_prev[1:4]) else [None, None, None])
    # 분양대행 적용 여부
    agency_rows = rows.get("분양대행수수료", [])
    f["분양대행"] = "Y" if (is_new or remain) and any(r["total"] for r in agency_rows) else "N"
    # 조직분양: 분양대행 시트 '(조직)' 블록 세대 또는 현수막 줄
    org_blocks = [b for b in (pr.get("agency_blocks") or []) if "조직" in b["label"] and any(b["months_units"])]
    org_units = pr.get("agency_org_units") or [0] * 12
    sign = months_of(rows, "조직분양 현수막")
    if (is_new or remain) and (org_blocks or any(sign)):
        f["조직"] = "Y"
        src = org_units if any(org_units) else (org_blocks[0]["months_units"] if org_blocks else sign)
        s0, s1 = first_nonzero(src), last_nonzero(src)
        if any(sign) and (s0 is None or first_nonzero(sign) < s0):
            s0 = first_nonzero(sign)
        f["조직시작"] = [year, s0] if s0 else [None, None]
        f["조직종료"] = [year, s1] if s1 and s1 < 12 else [None, None]
        price = org_blocks[0]["price"] if org_blocks else None
        if price in (8000, 10000, 15000):
            f["급지"] = {8000: "B급", 10000: "C급", 15000: "D급"}[price]
        else:
            f["급지"] = "C급"
        tot_units = pr.get("total_units") or 0
        org_total = sum(org_units) if any(org_units) else None
        if org_total and tot_units:
            ratio = org_total / tot_units
            f["조직비율"] = round(ratio, 2) if ratio < 0.995 else None
        else:
            f["조직비율"] = None
    else:
        f["조직"] = "N"
        # 급지: 선물 줄로 추정
        if rows.get("청약 선물") or rows.get("당첨자 선물") or rows.get("계약 선물"):
            f["급지"] = "C급"
        elif rows.get("오픈 선물"):
            f["급지"] = "B급"
        else:
            f["급지"] = "C급" if is_new else "A급"
    # 입주
    chk_m = next((i + 1 for i, s in enumerate(stage) if isinstance(s, str) and "사전점검" in s), None)
    in_m = next((i + 1 for i, s in enumerate(stage) if isinstance(s, str) and s.startswith("입주")), None)
    if in_m is None and any(move) and (rows.get("입주지원센터") or rows.get("입주매니저") or rows.get("사전점검행사") or rows.get("입주상담사")):
        in_m = first_nonzero(move)
    if chk_m is None and rows.get("사전점검행사"):
        chk_m = first_nonzero(months_of(rows, "사전점검행사"))
    if in_m or chk_m:
        kinds.append("movein")
        f["사전점검"] = [year, chk_m] if chk_m else [None, None]
        f["입주"] = [year, in_m] if in_m else [None, None]
        f["전년입주"] = move_prev[0] if move_prev else 0
        # 입주지정기간: 입주상담사 개월 - 1 (기본 2)
        cons = months_of(rows, "입주상담사")
        n = sum(1 for v in cons if v)
        f["지정기간"] = max(1, n - 1) if n else 2
        f["입주촉진"] = "Y" if rows.get("입주촉진책") else "N"
        f["어린이집"] = "Y" if any("어린이집" in (r.get("g") or "") for r in rows.get("우편·등기·내용증명", [])) else "N"
    else:
        f["사전점검"] = [None, None]
        f["입주"] = [None, None]
        # 기입주 현장(공가관리비만 있는 경우): 입주개시월을 전년으로 두고 미입주세대만 반영
        if rows.get("공가세대관리비") and any(move_prev) or (rows.get("공가세대관리비") and move_prev and move_prev[0]):
            f["전년입주"] = move_prev[0]
    # 분양사무실: 통신요금 150(2023~2026)·200 등 — M/H 600 이 아닌 통신 달
    office = [1 if (v and v < 550000) else 0 for v in mh_tel]
    if any(office):
        kinds.append("office")
        o0, o1 = first_nonzero(office), last_nonzero(office)
        f["사무실시작"] = [year, o0]
        f["사무실종료"] = [year, o1] if o1 < 12 else [None, None]
    # 상가 호실(인지세 산출근거에 '상가' n실)
    for r in rows.get("인지세", []):
        m = re.search(r"(\d+)\s*(?:실|세대)\(상가\)|상가.*?(\d+)\s*실|(\d+)\s*실\s*\(상가", r.get("g") or "")
        if m:
            f["상가"] = int(next(x for x in m.groups() if x))
    # PC 대수
    for r in rows.get("PC", []):
        m = re.search(r"[×x]\s*(\d+)\s*대", r.get("g") or "")
        if m:
            f["PC"] = int(m.group(1))
    # 숙소 호실
    for r in rows.get("숙소관리비", []):
        m = re.search(r"[×x]\s*(\d+)\s*호", r.get("g") or "")
        if m:
            if "movein" in kinds and not is_new:
                f["숙소입주"] = int(m.group(1))
            else:
                f["숙소분양"] = int(m.group(1))
    # 피복비: 실제 줄이 있으면 1명(2026)/과거 30명×1,000 은 단가 구조가 달라 1명으로
    f["여직원"] = 1 if rows.get("피복비") else 0
    # 세대 구성(공급면적은 공가 시트에서 못 읽으면 84→112.5㎡ 같은 전용→공급 추정: 전용 ㎡ × 1.33)
    types = []
    for name, units in pr.get("types") or []:
        m = re.match(r"(\d+)", str(name))
        exclusive = int(m.group(1)) if m else 84
        supply = pr.get("supply", {}).get(str(name))
        if not supply or supply > 200 or supply < exclusive:      # 공가 시트 잔존값(다른 현장)·평 단위 오독은 버림
            supply = round(exclusive * 1.33, 1)
        types.append([str(name), units, supply])
    if override:
        for k, v in (override.get("fields") or {}).items():
            f[k] = v
        if "kinds" in override:
            kinds = override["kinds"]
        if "monthly" in override:
            monthly.update(override["monthly"])
        notes += override.get("notes") or []
    data = dict(fields=f, types=types, monthly=monthly, kinds=kinds)
    return data, notes


# ─────────────────────────────── 비교
GROUP_ALIAS = {"보험료": "보험료", "세금과공과": "세금과공과"}


def compare(pr, res):
    """항목(묶음)별 + 줄키별 실제 ↔ 도구"""
    actual_groups = {}
    actual_lines = {}
    for r in pr.get("rows") or []:
        item = (r.get("item") or "").replace(" ", "")
        if r["sub"] == "소계":
            continue
        if not item:
            continue
        # 실제 파일에서 세금과공과 소계 줄이 '보험료' 로 잘못 적힌 경우(2025 오산서2) 등: 항목은 줄의 B열 그대로
        msum = sum(r.get("months") or [])
        amt = r["total"] or 0
        if abs(amt - msum) > max(1000, 0.01 * max(abs(amt), abs(msum))) and msum:
            amt = msum
        actual_groups[item] = actual_groups.get(item, 0) + amt
        k = line_key((r.get("sub") or "") + " " + (r.get("g") or ""), item)
        actual_lines.setdefault((item, k), [0, []])
        actual_lines[(item, k)][0] += amt
        actual_lines[(item, k)][1].append((r.get("g") or "").replace("\n", " / ")[:90])
    tool_groups = {}
    tool_lines = {}
    for g in res["groups"]:
        item = (g["item"] or "").replace(" ", "")
        tool_groups[item] = tool_groups.get(item, 0) + (g["total"] or 0)
        for r in g["rows"]:
            if not r["total"]:
                continue
            k = line_key(r["g"], item)
            tool_lines.setdefault((item, k), [0, []])
            tool_lines[(item, k)][0] += r["total"] or 0
            tool_lines[(item, k)][1].append(r["g"][:90])
    items = []
    for it in sorted(set(actual_groups) | set(tool_groups), key=lambda x: (x not in actual_groups, x)):
        a, t = actual_groups.get(it, 0), tool_groups.get(it, 0)
        items.append(dict(item=it, actual=a, tool=t, diff=t - a))
    lines = []
    for key in sorted(set(actual_lines) | set(tool_lines)):
        a = actual_lines.get(key, [0, []])
        t = tool_lines.get(key, [0, []])
        lines.append(dict(item=key[0], key=key[1], actual=a[0], tool=t[0], diff=t[0] - a[0],
                          actual_g=" | ".join(a[1])[:160], tool_g=" | ".join(t[1])[:160]))
    return items, lines


def classify(line):
    a, t = line["actual"], line["tool"]
    if abs(a - t) <= max(1000, 0.005 * max(abs(a), abs(t))):
        return "일치"
    if a and not t:
        return "도구 미반영"
    if t and not a:
        return "실제 미반영(도구만)"
    r = t / a if a else 0
    return "도구 多" if r > 1 else "도구 少"


import infer_rules


def run_all(selected=None, overrides=None, mode="basic"):
    overrides = overrides or {}
    results = []
    for p in sorted(glob.glob(os.path.join(SP, "profiles", "*.json"))):
        pr = json.load(open(p, encoding="utf-8"))
        if selected and not any(s in pr["key"] for s in selected):
            continue
        year = pr["year"]
        ov = overrides.get(pr["key"]) or {}
        if "_자동작성입력" in (pr.get("sheets") or []):
            data = site_io.read_inputs(pr["path"])
            data.pop("year", None)
            notes = ["자동작성 도구로 만든 파일 — 숨긴 시트의 입력값을 그대로 다시 계산(회귀 확인)"]
        else:
            data, notes = infer_rules.infer(pr, rows_by_key, line_key, mode, ov)
        model = model_for(year)
        t = time.time()
        res = site_io.calc_model(model, data)
        sec = time.time() - t
        items, lines = compare(pr, res)
        for ln in lines:
            ln["class"] = classify(ln)
        a_tot_file = pr.get("total") or 0
        a_tot = sum(it["actual"] for it in items)          # 줄별(월별 합 기준) 합계
        inconsistent = abs(a_tot - a_tot_file) > max(1000, 0.005 * max(a_tot, a_tot_file))
        t_tot = res["total"]["total"] or 0
        results.append(dict(key=pr["key"], year=year, file=pr["file"], name=pr.get("name"), units=pr.get("total_units"), mode=mode,
                            kinds=data["kinds"], inputs=data, notes=notes, actual_total=a_tot, actual_total_file=a_tot_file, inconsistent=inconsistent, tool_total=t_tot,
                            diff=t_tot - a_tot, ratio=(t_tot / a_tot) if a_tot else None, nerrors=res["nerrors"],
                            errors=res["errors"][:5], stage=res["stage"], items=items, lines=lines, sec=sec))
        print(f"■ [{year}] {pr['file'][:48]:<48} 실제 {a_tot/1000:>13,.0f} 도구 {t_tot/1000:>13,.0f} 차이 {(t_tot-a_tot)/1000:>+13,.0f}천원 "
              f"({(t_tot/a_tot*100 if a_tot else 0):6.1f}%) 구분 {data['kinds']} 오류 {res['nerrors']} {sec:.1f}초")
    return results


if __name__ == "__main__":
    sel = [a for a in sys.argv[1:] if not a.startswith("--")]
    ovp = os.path.join(SP, "overrides.json")
    overrides = json.load(open(ovp, encoding="utf-8")) if os.path.exists(ovp) else {}
    mode = "adjusted" if "--adjusted" in sys.argv else "basic"
    results = run_all(sel or None, overrides, mode)
    out = os.path.join(SP, f"compare_results_{mode}.json")
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(results, fh, ensure_ascii=False, indent=1, default=str)
    print("저장:", out, "현장", len(results))
    if "--detail" in sys.argv:
        for r in results:
            print(f"\n===== [{r['year']}] {r['file']}  구분 {r['kinds']}  단계 {r['stage']}")
            print("  입력:", {k: v for k, v in r["inputs"]["fields"].items() if v not in (None, [None, None], "", 0)})
            for n in r["notes"]:
                print("  note:", n)
            for it in r["items"]:
                flag = "  " if abs(it["diff"]) <= 1000 else "**"
                print(f"  {flag} {it['item']:<8} 실제 {it['actual']/1000:>13,.0f} 도구 {it['tool']/1000:>13,.0f} 차이 {it['diff']/1000:>+13,.0f}")
            for ln in r["lines"]:
                if ln["class"] == "일치":
                    continue
                print(f"      · {ln['item'][:6]:<6} {ln['key']:<16} 실제 {ln['actual']/1000:>12,.0f} 도구 {ln['tool']/1000:>12,.0f} [{ln['class']}] {ln['actual_g'][:60]} ⇔ {ln['tool_g'][:60]}")
