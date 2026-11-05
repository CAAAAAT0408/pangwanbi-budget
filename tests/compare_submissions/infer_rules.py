# -*- coding: utf-8 -*-
"""현장 조건 추정(개선판) — compare_sites.infer_inputs 를 대체. mode='basic'(세부 조정은 기본값) / 'adjusted'(제출본에 적힌 수량·시점도 반영)"""
import re
import statistics


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


def months_of(rows, key):
    acc = [0] * 12
    for r in rows.get(key, []):
        for i, m in enumerate(r["months"]):
            acc[i] += m or 0
    return acc


def gtext(rows, key):
    return " ".join((r.get("g") or "").replace("\n", " ") for r in rows.get(key, []))


def infer(pr, rows_by_key, line_key, mode="basic", override=None):
    year = pr["year"]
    stage = pr.get("stage") or [None] * 12
    rows = rows_by_key(pr)
    rates = pr.get("rates") or {}
    sale = (rates.get("분양세대") or {}).get("months") or [0] * 12
    sale_prev = (rates.get("분양세대") or {}).get("prev") or [0] * 4
    move = (rates.get("입주세대") or {}).get("months") or [0] * 12
    move_prev = (rates.get("입주세대") or {}).get("prev") or [0] * 4
    total_units = pr.get("total_units") or 0
    is_mh_file = " MH)" in pr["file"]
    f = {"현장명": pr.get("name") or pr["file"], "지역": "수도권", "사업구분": "자체사업", "상품": "아파트"}
    kinds, notes, adjust_notes = [], [], []

    # ── M/H 운영 달: 통신 600 / 전기료 / 상하수도 / 무인경비(≥350) / 전기·소방 안전관리 / 청소
    tel = months_of(rows, "통신요금")
    mh_on = [0] * 12
    elec, water, clean = months_of(rows, "전기료"), months_of(rows, "상하수도"), months_of(rows, "청소")
    evidence = any(elec) or any(water) or any(clean)
    for i in range(12):
        if elec[i] or water[i] or (tel[i] >= 550000 and evidence):      # 전기료·상하수도(또는 M/H 증거가 있을 때 통신 600)
            mh_on[i] = 1
    office = [1 if (0 < tel[i] < 550000) else 0 for i in range(12)]

    # ── 오픈월(D): 단계 머리글 '오픈(D+1)'(=오픈월) 또는 '오픈', 없으면 식대 오픈 단가(≥2,800천원) 첫 달, 없으면 분양 첫 달-1(M/H 운영 중일 때)
    open_m = None
    for i, s in enumerate(stage):
        if isinstance(s, str) and s.startswith("오픈") and "D+2" not in s:
            open_m = i + 1
            break
    if open_m is None:
        meal = months_of(rows, "식대")
        for i, v in enumerate(meal):
            if v >= 2800000:
                open_m = i + 1
                break
    if open_m is None and any(sale) and any(mh_on):
        fs = first_nonzero(sale)
        if fs and fs > 1:
            open_m = fs - 1
    pre_cnt = sum(1 for s in stage if isinstance(s, str) and s.startswith("사전") and "점검" not in s)
    is_new = open_m is not None and not is_mh_file
    remain = (not is_new) and any(mh_on) and not is_mh_file

    if is_new:
        kinds.append("new")
        f["오픈"] = [year, open_m]
        f["사전영업"] = pre_cnt or 1
        last_mh = last_nonzero(mh_on)
        f["MH종료"] = [year, last_mh] if last_mh and last_mh < 12 else [None, None]
    elif remain:
        kinds.append("remain")
        f["오픈"] = [year - 1, 6]
        f["사전영업"] = 1
        first_mh, last_mh = first_nonzero(mh_on), last_nonzero(mh_on)
        f["MH종료"] = [year, last_mh] if last_mh and last_mh < 12 else [None, None]
        notes.append(f"기공급(잔여분양): M/H 운영 {first_mh}~{last_mh or 12}월. 오픈월은 전년 6월로 두어 올해는 모두 '기타/잔여판촉' 구간")
    else:
        f["오픈"] = [None, None]

    selling = is_new or remain
    # 월별 분양·입주: 분양 활동이 있을 때만 실제 월별 계약세대를 넘김(기입주 현장의 장부용 숫자는 무시)
    monthly = dict(분양=[None] * 12, 입주=[None] * 12, 전년분양=[None] * 3, 전년입주=[None] * 3)
    if remain and total_units and sale[0] >= 0.5 * total_units and not sale_prev[0]:
        f["전년분양_보정"] = sale[0]
        notes.append(f"1월 분양세대 {sale[0]:,}은 기분양 누계로 보아 전년 누계로 옮김")
        sale = [0] + sale[1:]
        sale_prev = [f["전년분양_보정"]] + sale_prev[1:]
    if selling and any(sale):
        monthly["분양"] = [v or None for v in sale]
        monthly["전년분양"] = [v or None for v in sale_prev[1:4]]
        f["전년분양"] = sale_prev[0] or 0
    else:
        f["전년분양"] = 0

    # ── 분양대행·조직분양
    f["분양대행"] = "Y" if selling and any(r["total"] for r in rows.get("분양대행·조직 수수료", [])) else "N"
    org_units = pr.get("agency_org_units") or [0] * 12
    sign = months_of(rows, "조직분양 현수막")
    org_blocks = [b for b in (pr.get("agency_blocks") or []) if "조직" in b["label"] and any(b["months_units"])]
    if selling and (any(org_units) or any(sign)):
        f["조직"] = "Y"
        src = org_units if any(org_units) else sign
        s0, s1 = first_nonzero(src), last_nonzero(src)
        if any(sign) and first_nonzero(sign) < (s0 or 13):
            s0 = first_nonzero(sign)
        f["조직시작"] = [year, s0] if s0 else [None, None]
        f["조직종료"] = [year, s1] if s1 and s1 < 12 else [None, None]
        price = org_blocks[0]["price"] if org_blocks else None
        f["급지"] = {8000: "B급", 10000: "C급", 15000: "D급"}.get(price, "C급")
        if price and price not in (8000, 10000, 15000):
            f["조직단가"] = price
        # 조직비율: 조직 세대 합 ÷ 조직 기간 분양 합이 98% 이상이면 상한 없음
        if any(org_units) and s0:
            period_sales = sum(sale[s0 - 1:(s1 or 12)])
            org_total = sum(org_units)
            if period_sales and org_total < 0.98 * period_sales and total_units:
                f["조직비율"] = round(org_total / total_units, 2)
                notes.append(f"조직분양 세대 {org_total:,.0f} / 기간 분양 {period_sales:,.0f} → 비율 상한 {f['조직비율']:.0%}")
            else:
                f["조직비율"] = None
        else:
            f["조직비율"] = None
        if any(sign):
            # 현수막 감축(2주/월): 월 금액이 앞 달의 절반이 되는 첫 달
            full = max(sign)
            for i, v in enumerate(sign):
                if v and full and abs(v - full / 2) < full * 0.05:
                    f["현수막감축"] = [year, i + 1]
                    f["현수막감축주수"] = 2
                    break
    else:
        f["조직"] = "N"
        if rows.get("청약 선물") or rows.get("당첨자 선물") or rows.get("계약 선물"):
            f["급지"] = "C급"
        elif rows.get("오픈 선물"):
            f["급지"] = "B급"
        else:
            f["급지"] = "C급" if is_new else "A급"
    if is_new and not rows.get("오픈 선물") and not rows.get("청약 선물") and not rows.get("당첨자 선물") and not rows.get("계약 선물") and f["조직"] == "N":
        f["급지"] = "A급"

    # ── 입주(사전점검·입주개시): 단계 머리글 또는 입주지원센터/사전점검행사 줄이 있을 때만
    chk_m = next((i + 1 for i, s in enumerate(stage) if isinstance(s, str) and "사전점검" in s), None)
    in_m = next((i + 1 for i, s in enumerate(stage) if isinstance(s, str) and s.startswith("입주")), None)
    strong = bool(rows.get("입주지원센터") or rows.get("사전점검행사") or (rows.get("입주상담사") and rows.get("입주매니저")))
    if is_mh_file:
        chk_m = in_m = None
        strong = False
    if (in_m or chk_m) and not strong and not any(move):
        chk_m = in_m = None
    if (in_m or chk_m) or strong:
        if in_m is None and any(move):
            in_m = first_nonzero(move)
        if chk_m is None and rows.get("사전점검행사"):
            chk_m = first_nonzero(months_of(rows, "사전점검행사"))
        if in_m is None and rows.get("입주지원센터"):
            in_m = first_nonzero(months_of(rows, "입주지원센터"))
        kinds.append("movein")
        f["사전점검"] = [year, chk_m] if chk_m else [None, None]
        f["입주"] = [year, in_m] if in_m else [None, None]
        f["전년입주"] = move_prev[0] or 0
        if any(move):
            monthly["입주"] = [v or None for v in move]
            monthly["전년입주"] = [v or None for v in move_prev[1:4]]
        cons = months_of(rows, "입주상담사")
        n = sum(1 for v in cons if v)
        f["지정기간"] = max(1, n - 1) if n else 2
        f["입주촉진"] = "Y" if rows.get("입주촉진책") else "N"
        f["어린이집"] = "Y" if "어린이집" in gtext(rows, "우편·등기·내용증명") else "N"
        if not rows.get("문고"):
            notes.append("실제 제출본에 문고 설치 없음(도구는 입주개시월 10,000 반영)")
    else:
        f["사전점검"] = [None, None]
        f["입주"] = [None, None]
        vac_rows = rows.get("공가세대관리비", [])
        if vac_rows and not is_mh_file:
            # 기입주 현장: 입주개시를 전년 12월로 두고, 미입주세대만큼 전년 입주누계를 뺀다
            g = gtext(rows, "공가세대관리비")
            m = re.search(r"(?:아파트\s*)?(\d+)\s*세대", g)
            vac = int(m.group(1)) if m else None
            if vac is None:
                amt = months_of(rows, "공가세대관리비")
                first_amt = next((v for v in amt if v), 0)
                avg_py = statistics.mean([(pr.get("supply", {}).get(t[0]) or 112) * 0.3025 for t in pr.get("types") or [["84A", 1]]])
                vac = max(1, round(first_amt / (6000 * avg_py))) if first_amt else 1
            f["입주"] = [year - 1, 6]
            f["전년입주"] = max(0, total_units - vac)
            kinds.append("movein")
            notes.append(f"기입주 현장: 공가세대관리비만 반영(미입주 {vac}세대로 추정, 입주개시 전년 6월·전년입주 {f['전년입주']}세대)")

    # ── 분양사무실
    if any(office):
        kinds.append("office")
        o0, o1 = first_nonzero(office), last_nonzero(office)
        f["사무실시작"] = [year, o0] if o0 > 1 else [year - 1, 12]      # 1월부터면 전년부터 운영 중(시작월 운반비 없음)
        f["사무실종료"] = [year, o1] if o1 < 12 else [None, None]
    # ── 상가 호실(인지세 산출근거)
    g_tax = gtext(rows, "인지세")
    m = re.search(r"(\d+)\s*(?:실|세대)\s*\(상가\)|상가\)?\s*[:：]?\s*(\d+)\s*실|(\d+)\s*실\s*\(상가", g_tax)
    if m:
        f["상가"] = int(next(x for x in m.groups() if x))
    # ── PC 대수·숙소 호실·여직원
    m = re.search(r"[×x]\s*(\d+)\s*대", gtext(rows, "PC"))
    if m:
        f["PC"] = int(m.group(1))
    m = re.search(r"[×x]\s*(\d+)\s*호", gtext(rows, "숙소관리비"))
    if m:
        if "movein" in kinds and not selling:
            f["숙소입주"] = int(m.group(1))
        else:
            f["숙소분양"] = int(m.group(1))
    if not rows.get("숙소비") and selling and is_new:
        pass
    f["여직원"] = 1 if rows.get("피복비") else 0

    # ── 조정 모드: 제출본에 적힌 수량·시점을 입력 칸에 반영(화면 '세부 조정'에서 넣을 수 있는 것만)
    if mode == "adjusted":
        amt = sum(months_of(rows, "대형마트"))
        unit = {2027: 13000000, 2026: 13000000, 2025: 12000000, 2024: 12000000, 2023: 6000000}.get(year, 13000000)
        if amt:
            f["대형마트"] = max(1, int(round(amt / unit)))
            adjust_notes.append(f"대형마트 {f['대형마트']}주(금액 {amt/1000:,.0f}천원 ÷ {unit/1000:,.0f})")
        m = re.search(r"[×x]\s*([\d,]+)\s*명", gtext(rows, "오픈 선물"))
        if m:
            f["오픈선물인원"] = int(m.group(1).replace(",", ""))
            adjust_notes.append(f"오픈선물 {m.group(1)}명")
        m = re.search(r"[×x]\s*([\d,]+)\s*명", gtext(rows, "청약 선물"))
        if m:
            f["청약선물인원"] = int(m.group(1).replace(",", ""))
        # 유니트(청소 오픈월 인원): 오픈월 청소 금액 ÷ 단가
        if is_new and open_m:
            cl = months_of(rows, "청소")
            v = cl[open_m - 1]
            if v >= 14000000 - 1 or (v and round(v / 3500000) >= 4) or (v and round(v / 3000000) == 4 and v % 3000000 == 0):
                f["유니트"] = 4
                adjust_notes.append("유니트 4개(청소 4명)")
        # 보험료 반영월
        ins = months_of(rows, "보험료")
        if any(ins):
            f["보험월"] = first_nonzero(ins)
            adjust_notes.append(f"보험료 {f['보험월']}월")
        # 사전영업 갑티슈 수량 등: 갑티슈 금액이 기본(15,000+15,000)과 다르면 수량 환산
        g = gtext(rows, "갑티슈")
        m = re.findall(r"([\d,]+)\s*개", g)
        if len(m) >= 2:
            try:
                f["갑티슈사전"], f["갑티슈오픈"] = int(m[0].replace(",", "")), int(m[1].replace(",", ""))
            except ValueError:
                pass
        # 예비입주자 내용증명 비율: '예비당첨 500%' 등
        m = re.search(r"(\d+)\s*%\)", gtext(rows, "우편·등기·내용증명"))
        if m:
            f["예비비율"] = int(m.group(1)) / 100
        # 전기 미인입(난방비 줄에 30,000)
        heat = months_of(rows, "난방비·전기미인입")
        if any(v >= 30000000 for v in heat):
            f["전기미인입"] = "Y"
            adjust_notes.append("전기미인입 Y")
        # 등기월: 우편 줄에서 기본·내용증명 외 큰 금액이 있는 달(최대 2개)
        post = months_of(rows, "우편·등기·내용증명")
        big = [(i + 1, v) for i, v in enumerate(post) if v >= 2000000]
        if is_new and open_m:
            big = [(mm, v) for mm, v in big if mm != open_m]      # 오픈월은 내용증명
        if big:
            big = sorted(big, key=lambda x: -x[1])[:2]
            ms = sorted(mm for mm, _ in big)
            f["등기월1"] = ms[0]
            f["등기월2"] = ms[1] if len(ms) > 1 else None
            adjust_notes.append(f"등기 {'·'.join(str(x) for x in ms)}월")
        # 상가 부동산MGM: '10,000천원 x 5건(상가)' / '부동산MGM(상가) : 10,000원 x 2건'
        g = gtext(rows, "분양촉진(MGM)")
        if "상가" in g:
            m = re.search(r"[×x]\s*(\d+)\s*건", g)
            if m:
                f["상가MGM"] = int(m.group(1))
                mm = first_nonzero(months_of(rows, "분양촉진(MGM)"))
                f["상가MGM월"] = [year, mm] if mm else [None, None]
                adjust_notes.append(f"상가MGM {m.group(1)}호실({mm}월)")
        # 잔여판촉 갑티슈(기공급): 갑티슈 금액 ÷ 1.5천원 → 수량, 첫 달
        if remain and rows.get("갑티슈"):
            amt = months_of(rows, "갑티슈")
            if any(amt):
                f["잔여갑티슈"] = int(round(sum(amt) / 1500))
                f["잔여판촉"] = [year, first_nonzero(amt)]
                adjust_notes.append(f"잔여판촉 갑티슈 {f['잔여갑티슈']:,}개({first_nonzero(amt)}월)")
        # 입주 안내 등기 횟수 / 어린이집은 그대로

    # ── 세대 구성(공급면적: 공가 시트 값, 없으면 전용 × 1.33)
    types = []
    for name, units in pr.get("types") or []:
        m = re.match(r"(\d+)", str(name))
        excl = int(m.group(1)) if m else 84
        sup = (pr.get("supply") or {}).get(str(name))
        if not sup or sup > 250 or sup < excl:
            sup = round(excl * 1.33, 1)
        types.append([str(name), units, sup])

    if override:
        for k, v in (override.get("fields") or {}).items():
            f[k] = v
        if "kinds" in override:
            kinds = override["kinds"]
        if "monthly" in override:
            monthly.update(override["monthly"])
        notes += override.get("notes") or []
    f.pop("전년분양_보정", None)
    if adjust_notes:
        notes.append("조정: " + ", ".join(adjust_notes))
    return dict(fields=f, types=types, monthly=monthly, kinds=kinds), notes
