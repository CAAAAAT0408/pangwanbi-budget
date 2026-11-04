# -*- coding: utf-8 -*-
"""지출계획서(견본주택 오픈·운영 / 입주 / 분양사무실 비용, 보통 3개월) 자동작성 엔진.

판관비 예산안 자동작성과 **같은 기준단가(기준_…, 천원)와 같은 단계 규칙**(사전영업·오픈·D+1·조직분양(잔여판촉)·
입주지원센터·분양사무실)으로 기간의 월별 비용을 만들고, 칸마다 산식(계산 수식)과 비고 글을 함께 돌려준다.
지출계획서에 들어가는 항목은 예산안 중 현장이 직접 지출하는 운영비(복리후생비·소모품비·지급임차료·지급수수료·
수도광열비(상하수도)·통신비·운반비·여비교통비, 분양사무실 관리비)이고, 광고·판촉·인지세·전기료·보험료는
따로 결재하므로 기본으로 넣지 않는다(전기료·난방비·보험료·송금수수료는 '반영 항목'으로 켤 수 있음).

  schema(K)                 화면 입력 칸 목록(구역·칸·기본값·구분 표시)
  calc(data, K)             입력값 → 결과(제목·월 머리글·항목별 월별 금액/산식/비고·소계·합계·점검)
  default_params()          기준단가 기본값(params.PARAMS의 2027 기준안 값) — 양식 '기준단가' 시트 값으로 덮어쓴다
  period_text(res)          파일 이름용 기간 글 ('26.10_12월)

data 구조(화면·저장 파일 공통)
  kinds     : ["mh", "run", "movein", "office"] 중 해당하는 것 (견본주택 오픈 / 견본주택 운영 / 입주 / 분양사무실)
              mh 는 오픈 일정(오픈월·사전영업)으로 단계를 잡고, run 은 오픈 일정 없이 기간 전체를 M/H 운영 기준으로(둘 다면 mh)
  fields    : {칸이름: 값}  (연월 칸은 [연, 월])
  months    : 기간 달마다 {"events": [행사…], "label": 머리글 직접입력, "식대인원": n, "청소인원": n}
  overrides : {줄키: {"0": 금액 또는 "=산식", …}}   화면에서 직접 고친 칸
  notes     : {줄키: 비고 글}                          직접 고친 비고
  extra     : [{"item": 항목, "name": 세목, "vals": [금액…], "note": 글}]   직접 추가한 줄
  title     : 제목 직접입력(비우면 자동)
  approvers : {"list": [{"role": 담당|검토|승인|협조, "pos": 직위, "name": 이름}], "width": 칸 너비, "height": 서명칸 높이}
              결재란 — 사람 수만큼 칸이 생기고 같은 역할이 이어지면 머리글을 합친다. 비우면 담당 과장 / 검토 부장·임원 / 승인 실장
"""
import datetime
import math
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import params as _P
except Exception:        # noqa: BLE001
    _P = None

INF = 10 ** 9

# ── 항목(묶음)과 줄 ───────────────────────────────────────────────────────
GROUPS = ["복리후생비", "소모품비", "지급임차료", "지급수수료", "보험료", "도서인쇄비", "수도광열비", "통신비", "운반비",
          "여비교통비", "관리비"]
# 줄키 → (항목, 세목 표시)
ROWS = [
    ("식대", "복리후생비", "식    대"), ("숙박비", "복리후생비", "숙 박 비"), ("숙소관리비", "복리후생비", "숙소관리비"),
    ("사무용가구", "소모품비", "사무용가구"), ("소모품", "소모품비", "소 모 품"),
    ("집기비품", "지급임차료", "집기비품"),
    ("발급수수료", "지급수수료", "발급수수료"), ("송금수수료", "지급수수료", "송금수수료"),
    ("보안경비용역", "지급수수료", "보안경비용역"), ("용역-기타", "지급수수료", "용역-기타"), ("숙소중개", "지급수수료", "기타"),
    ("보험료", "보험료", "화재보험"), ("도서인쇄비", "도서인쇄비", "도서인쇄비"),
    ("상하수도", "수도광열비", "상하수도"), ("전기료", "수도광열비", "전기료"), ("난방비", "수도광열비", "난방비"),
    ("통신비", "통신비", "통신비"), ("운반비", "운반비", "운반비"), ("여비교통비", "여비교통비", "여비교통비"),
    ("관리비", "관리비", "관리비"),
]
ROW_INFO = {k: (g, n) for k, g, n in ROWS}
# 줄이 하나뿐이면 항목 칸과 세목 칸을 합쳐 항목 이름만 보이는 묶음(과거 작성 양식과 같게)
SINGLE_MERGE = {"보험료", "도서인쇄비", "수도광열비", "통신비", "운반비", "여비교통비", "관리비"}

# 행사(달마다 고를 수 있음): 키 → (머리글 글, 설명)
EVENTS = [
    ("중도금대출", "중도금대출", "중도금대출 자서 — 은행 식대, 확정일자(대출확인서) 600원×세대, 대출 안내문 등기 4.5천원×세대"),
    ("유상옵션", "옵션", "유상옵션 행사 — 옵션 안내문 등기 4.5천원×세대('옵션 안내문' 체크 시), 식대 인원은 기타 기준 이상"),
    ("명의변경", "명변", "명의변경 — 식대 인원은 기타 기준 이상"),
    ("선착순", "선착순", "선착순(잔여세대) 판매 — 잔여판촉 기준: 식대 20천원×3명, 청소 2명"),
    ("전매", "전매", "전매 — 머리글 표시만"),
    ("사무실이전", "사무실 이전", "사무실 이전 — 운반비 2,500천원"),
]
EVENT_LABEL = {k: lb for k, lb, _ in EVENTS}

KIND_LABELS = [("mh", "견본주택 오픈", "사전영업·오픈·계약 분기 — 오픈월을 넣으면 단계별(사전/오픈/D+1) 기준"),
               ("run", "견본주택 운영", "이미 오픈한 M/H 의 운영 분기 — 오픈 일정 없이 운영 기준, 행사(선착순·중도금대출 등)는 달마다"),
               ("movein", "입주", "사전점검·입주지원센터 운영"),
               ("office", "분양사무실", "M/H 철거 뒤 분양사무실 운영")]

# 화면 입력 칸: (키, 표시, 종류, 기본값, 단위, 도움말, 선택지, 구분태그, 바로보임)
#   종류: text date ym int num pct yn list   /  기본값이 'K:이름' 이면 기준단가 값, 'K:식대인원:…' 은 계산값
INPUTS = [
    ("1. 기본 사항", None),
    ("현장명", "현장명", "text", None, "", "예) 평택고덕2차 — 제목 '[현장명] M/H 오픈비용 지출계획서'", None, "", True),
    ("작성일", "작성일", "date", None, "", "'( 작성일 : 2026년 10월 06일 )' 표시용. 비우면 오늘", None, "", True),
    ("시작", "지출 기간 시작", "ym", None, "연 / 월", "지출계획서 첫 달", None, "", True),
    ("개월수", "기간(개월)", "int", 3, "개월", "보통 3개월(분기). 2~6", None, "", True),
    ("지역", "지역", "list", "수도권", "", "여비교통비 단가 구분", "수도권,비수도권", "", True),
    ("세대수", "총세대수", "int", None, "세대", "내용증명·확정일자·안내문 등기 건수의 기준", None, "", True),
    ("제목", "제목(직접)", "text", None, "", "비우면 내용에 따라 'M/H 오픈비용 / M/H 운영비용 / 입주비용 / 운영비용 지출계획서'", None, "", False),

    ("2. 견본주택 오픈·운영", None),
    ("오픈", "오픈월(D)", "ym", None, "연 / 월", "사전영업 뒤 오픈하는 달 — 식대·소모품·청소·여비가 사전/오픈/계약(D+1) 단계별 기준으로. 이미 오픈해 운영 중이면 「견본주택 운영」 구분으로", None, "mh", True),
    ("사전영업", "사전영업 기간", "int", 1, "개월", "오픈 전 사전영업 개월수(기준 1개월) — M/H 운영 시작, 숙소비·무인경비 공사·운반비 시점", None, "mh", True),
    ("MH종료", "M/H 운영 종료월", "ym", None, "연 / 월", "비우면 기간 끝까지 운영", None, "mh run", True),
    ("조직시작", "조직분양(잔여판촉) 시작월", "ym", None, "연 / 월", "이 기간은 잔여판촉 기준(식대 20천원×3명, 청소 2명). 선착순 행사를 고른 달도 같음", None, "mh run", False),
    ("조직종료", "조직분양(잔여판촉) 종료월", "ym", None, "연 / 월", "비우면 M/H 종료까지", None, "mh run", False),
    ("유니트", "유니트 수", "int", 3, "개", "청소 인원: 오픈 1개월차 3개 이하 3명/4개 이상 4명, 2개월차 2명/3명", None, "mh", True),
    ("숙소호실", "상주인력 숙소 호실 수", "int", 3, "호실", "오픈: 1,000천원/월×3개월 일시불(사전영업 시작월) · 운영: 기존 숙소 1,000천원/월×호실 — 둘 다 숙소관리비 200천원/월×호실 추가", None, "mh run", True),
    ("숙소시작", "숙소 시작월", "ym", None, "연 / 월", "비우면 사전영업 시작월(운영 구분은 기간 첫 달부터 기존 숙소 월납). 넣으면 그 달에 3개월 일시불", None, "mh run", False),
    ("숙소종료", "숙소 종료월", "ym", None, "연 / 월", "비우면 M/H 종료월. 일시불 3개월 뒤에도 운영 중이면 월 1,000천원×호실", None, "mh run", False),
    ("숙소월납", "숙소비 월납(일시불 대신 매월)", "yn", "N", "", "체크하면 3개월 일시불 대신 매월 1,000천원×호실", None, "mh run", False),
    ("숙소중개", "숙소 부동산 중개수수료", "num", 0, "천원/호실", "기준안에 없음. 필요하면 호실당 금액(숙소 시작월에 지급수수료-기타 줄로)", None, "mh run", False),
    ("PC", "사무실 PC 대수", "int", 6, "대", "100천원/월·대, M/H 운영 기간", None, "mh run", True),
    ("복합기", "복합기 대수", "int", 1, "대", "300천원/월·대", None, "mh run", True),
    ("프린터", "계약용 프린터 대수", "int", "K:프린터_대수", "대", "200천원/월·대, 오픈월부터", None, "mh run", False),
    ("파쇄기", "파쇄기(세단기)", "yn", "Y", "", "300천원/월, 오픈월부터", None, "mh run", False),
    ("보안공사", "무인경비 공사비", "yn", "Y", "", "사전영업 시작월 2,000천원 (용역비 400천원/월은 M/H 운영 기간 자동)", None, "mh", False),
    ("오수관리", "오수관리 용역", "yn", "N", "", "450천원/월 — 기준안에는 있으나 과거 지출계획서는 넣지 않음", None, "mh run", False),
    ("가구", "사무용가구", "yn", "Y", "", "4,000천원, 오픈 전월(사전영업 시작월)", None, "mh", False),
    ("운반비", "운반비(집기이동·이전)", "yn", "Y", "", "2,500천원/회 — 사전영업 시작월·입주개시월·분양사무실 시작월·'사무실 이전' 행사 달", None, "mh run movein office", False),
    ("보험료", "견본주택 보험료", "yn", "N", "", "10,000천원/년 — 사전영업 시작월. 보통 별도 결재라 기본 제외", None, "mh", False),
    ("송금수수료", "송금수수료", "yn", "N", "", "30천원/월 — 예산안에는 있으나 지출계획서에는 보통 넣지 않음", None, "mh run movein office", False),
    ("전기료", "전기료", "yn", "N", "", "동·하절기 10,000 / 기타 6,000천원/월 — 자동이체라 보통 제외", None, "mh run", False),
    ("난방비", "난방비", "yn", "N", "", "11~3월 1,000천원/월 — 보통 제외", None, "mh run", False),
    ("예비비율", "예비입주자 내용증명 비율", "pct", 5, "", "오픈월 내용증명 건수 = 총세대 × 비율 (기준 500%)", None, "mh", False),
    ("식대단가오픈", "식대 단가 - 오픈 2개월·잔여판촉", "num", 20, "천원/일", "20천원/일 (기준안)", None, "mh run", False),
    ("식대단가기타", "식대 단가 - 기타", "num", 10, "천원/일", "10천원/일 (기준안)", None, "mh run movein", False),
    ("식대인원오픈", "식대 인원 - 오픈 2개월", "int", "K:식대인원:식대_오픈:식대단가오픈", "명", "비우면 기준 금액(오픈 2개월 3,000천원 = 20천원×5명×30일)", None, "mh", False),
    ("식대인원기타", "식대 인원 - 기타(사전영업·오픈 3개월~·입주)", "int", "K:식대인원:식대_기타:식대단가기타", "명", "비우면 기준 금액(1,200천원 = 10천원×4명×30일)", None, "mh run movein", False),
    ("식대인원잔여", "식대 인원 - 조직분양·선착순(잔여판촉)", "int", "K:식대인원:식대_잔여:식대단가오픈", "명", "비우면 기준 금액(1,800천원 = 20천원×3명×30일)", None, "mh run", False),
    ("은행식대단가", "은행 식대 단가(중도금대출 자서)", "num", 10, "천원/일", "중도금대출 행사 달: 단가×일수×인원", None, "mh run", False),
    ("은행식대일수", "은행 식대 일수", "int", 7, "일", "", None, "mh run", False),
    ("은행식대인원", "은행 식대 인원", "int", 20, "명", "", None, "mh run", False),
    ("옵션안내", "유상옵션 안내문 등기", "yn", "Y", "", "유상옵션 행사 달 4.5천원×세대", None, "mh run", False),
    ("도서인쇄", "도서인쇄비(대봉투 등)", "num", 0, "천원", "사전영업 시작월(없으면 첫 달)에 1회. 0이면 줄 없음", None, "mh run", False),

    ("3. 입주 일정", None),
    ("사전점검", "사전점검월", "ym", None, "연 / 월", "입주지원센터 운영 시작(소모품 1,000·식대·여비·복합기·PC)", None, "movein", True),
    ("입주", "입주개시월", "ym", None, "연 / 월", "입주 숙소비(3개월 일시불)·입주기간 통신비·운반비", None, "movein", True),
    ("지정기간", "입주지정기간", "int", 2, "개월", "식대·여비·통신: 입주개시~지정기간+1개월", None, "movein", True),
    ("입주숙소", "입주 지원인력 숙소 호실 수", "int", 2, "호실", "입주개시월 1,000천원×3개월 일시불 + 숙소관리비 200천원/월", None, "movein", False),
    ("입주PC", "입주지원센터 PC 대수", "int", 4, "대", "", None, "movein", False),
    ("입주안내", "입주 안내 등기 횟수", "int", 2, "회", "1 = 사전점검 안내(사전점검 전월), 2 = + 입주 안내(입주 전월) — 4.5천원×세대", None, "movein", False),
    ("어린이집", "국공립어린이집 동의서 발송", "yn", "Y", "", "입주 전월 4.5천원×세대", None, "movein", False),

    ("4. 분양사무실 운영", None),
    ("사무실시작", "분양사무실 운영 시작월", "ym", None, "연 / 월", "M/H 철거 뒤 분양사무실 운영 시작 — 운반비 2,500천원", None, "office", True),
    ("사무실종료", "분양사무실 운영 종료월", "ym", None, "연 / 월", "비우면 기간 끝까지", None, "office", True),
    ("통합센터", "통합센터 운영", "yn", "N", "", "관리비 600 → 1,000천원/월", None, "office", False),
    ("사무실복합기", "복합기 월 임차료", "num", "K:사무실_복합기", "천원/월", "300~400 탄력 적용", None, "office", False),
    ("사무실보안", "보안경비 월 용역비", "num", "K:사무실_보안", "천원/월", "250~350 탄력 적용", None, "office", False),
]


def default_params():
    """기준단가 기본값 {이름: 값(천원 등)} — params.PARAMS(2027 기준안)에서"""
    out = {}
    if _P is not None:
        for it in _P.PARAMS:
            if it[0] != "■":
                out[it[0]] = it[4]
    return out


def _n(v, default=0):
    if v in (None, ""):
        return default
    try:
        f = float(str(v).replace(",", ""))
    except ValueError:
        return default
    return int(f) if f == int(f) else f


def _ym(v):
    """[연, 월] → 월 번호(연×12+월). 없으면 None"""
    try:
        y, m = int(v[0]), int(v[1])
    except Exception:    # noqa: BLE001
        return None
    if y < 100:
        y += 2000
    if not (1 <= m <= 12):
        return None
    return y * 12 + m


def _split(idx):
    return (idx - 1) // 12, (idx - 1) % 12 + 1


def _won(x):
    """천원 → 원(정수 글)"""
    v = round(float(x) * 1000)
    return str(int(v))


def _num(x):
    """산식에 넣을 숫자 글"""
    try:
        f = float(x)
    except (TypeError, ValueError):
        return "0"
    return str(int(f)) if f == int(f) else repr(f)


def _c(x):
    """천원 금액 표시 — 1,000천원 / 20천원 / 4.5천원"""
    try:
        f = float(x)
    except (TypeError, ValueError):
        return str(x)
    if f == int(f):
        return f"{int(f):,}천원"
    return f"{f:,.1f}".rstrip("0").rstrip(".") + "천원"


_SAFE = re.compile(r"^[\d\s+\-*/().]+$")


def evaluate(expr):
    """산식 글(숫자·+-*/·괄호만) → 값. 다른 글자가 있으면 None"""
    if not expr or not _SAFE.match(expr):
        return None
    try:
        return float(eval(expr, {"__builtins__": {}}, {}))      # noqa: S307
    except Exception:    # noqa: BLE001
        return None


def resolve_default(default, K, fields=None):
    """INPUTS 기본값 → 실제 값 ('K:이름' 은 기준단가, 'K:식대인원:기준:단가칸' 은 기준 금액에서 인원 계산)"""
    if isinstance(default, str) and default.startswith("K:"):
        parts = default.split(":")
        if parts[1] == "식대인원":
            base = _n(K.get(parts[2]), 0)
            unit = _n((fields or {}).get(parts[3]), None)
            if unit is None:
                unit = 20 if parts[3] == "식대단가오픈" else 10
            if not unit:
                return 0
            return max(1, int(round(base / (unit * 30)))) if base else 0
        return K.get(parts[1])
    return default


def schema(K=None):
    K = dict(default_params(), **(K or {}))
    sections, cur = [], None
    for it in INPUTS:
        if it[1] is None:
            cur = dict(title=it[0], fields=[])
            sections.append(cur)
            continue
        key, label, kind, default, unit, help_, opts, tags, main = it
        f = dict(key=key, label=label, kind=kind, unit=unit, help=help_, options=opts.split(",") if opts else None,
                 tags=tags.split() or None, main=main, auto=isinstance(default, str) and str(default).startswith("K:"))
        f["default"] = resolve_default(default, K) if kind != "ym" else [None, None]
        cur["fields"].append(f)
    return dict(sections=sections, kinds=[dict(key=k, label=l, help=h) for k, l, h in KIND_LABELS],
                events=[dict(key=k, label=l, help=h) for k, l, h in EVENTS], groups=GROUPS)


# ── 계산 ───────────────────────────────────────────────────────────────────
class _Cell:
    """한 칸: 산식 항(expr, 값, 비고 조각) 여러 개"""

    def __init__(self):
        self.terms = []

    def add(self, expr, note=None):
        v = evaluate(expr)
        if v is None:
            return
        self.terms.append((expr, v, note))

    @property
    def value(self):
        return sum(t[1] for t in self.terms)

    @property
    def formula(self):
        if not self.terms:
            return None
        parts = []
        for e, _, _ in self.terms:
            parts.append(f"({e})" if ("+" in e or "-" in e) and len(self.terms) > 1 else e)
        return "+".join(parts)

    @property
    def note(self):
        return ", ".join(t[2] for t in self.terms if t[2])


def _runs(ms):
    """월 목록(연속이면 ~ 로). ms 는 월 번호(연×12+월) 목록 정렬됨"""
    out, i = [], 0
    while i < len(ms):
        j = i
        while j + 1 < len(ms) and ms[j + 1] == ms[j] + 1:
            j += 1
        a, b = _split(ms[i])[1], _split(ms[j])[1]
        out.append(f"{a}월" if i == j else f"{a}~{b}월")
        i = j + 1
    return ",".join(out)


def _join_notes(cells, idxs):
    """달마다 비고 조각 → 한 줄 글(모두 같으면 달 표시 없이)"""
    notes = [c.note for c in cells]
    nonblank = [(i, n) for i, n in zip(idxs, notes) if n]
    if not nonblank:
        return ""
    if len(set(n for _, n in nonblank)) == 1 and len(nonblank) == len(cells):
        return " " + nonblank[0][1]
    lines, seen = [], {}
    for i, n in nonblank:
        seen.setdefault(n, []).append(i)
    for n, ms in sorted(seen.items(), key=lambda x: x[1][0]):
        lines.append(f" {_runs(ms)} : {n}")
    return "\n".join(lines)


ROLES = ["담당", "검토", "승인", "협조"]
DEFAULT_APPROVERS = [dict(role="담당", pos="과장", name=""), dict(role="검토", pos="부장", name=""),
                     dict(role="검토", pos="임원", name=""), dict(role="승인", pos="실장", name="")]


def approver_box(ap):
    """결재란 입력 → dict(list=[{role,pos,name}], width, height). 역할이 비면 첫째 담당·마지막 승인·가운데 검토"""
    ap = ap if isinstance(ap, dict) else {}
    lst = []
    for a in ap.get("list") or []:
        if isinstance(a, str):
            a = dict(pos=a)
        if not isinstance(a, dict):
            continue
        pos, name = str(a.get("pos") or "").strip(), str(a.get("name") or "").strip()
        if not pos and not name:
            continue
        role = str(a.get("role") or "").strip()
        lst.append(dict(role=role if role in ROLES else "", pos=pos, name=name))
    if not lst:
        lst = [dict(x) for x in DEFAULT_APPROVERS]
    for i, a in enumerate(lst):
        if not a["role"]:
            a["role"] = "담당" if i == 0 else ("승인" if i == len(lst) - 1 else "검토")
    width = _n(ap.get("width"), None)
    height = _n(ap.get("height"), None)
    return dict(list=lst, width=width if width and width > 0 else None, height=height if height and height > 0 else None)


def calc(data, K=None):
    K = dict(default_params(), **(K or {}))
    f = dict(data.get("fields") or {})
    kinds = set(data.get("kinds") or [])
    if "mh" in kinds:
        kinds.discard("run")                      # 오픈 구분이 있으면 운영 구분은 뜻이 없음(같은 M/H)
    mhk = "mh" in kinds or "run" in kinds        # 견본주택(오픈 또는 운영)
    warnings = []

    def fv(key):
        """칸 값(비우면 기본값)"""
        v = f.get(key)
        if v in (None, ""):
            it = next((x for x in INPUTS if x[1] is not None and x[0] == key), None)
            if it is None:
                return None
            return resolve_default(it[3], K, f) if it[2] != "ym" else None
        return v

    def yn(key):
        return str(fv(key) or "N").upper().startswith("Y")

    start = _ym(f.get("시작"))
    n = int(_n(fv("개월수"), 3) or 3)
    n = max(1, min(n, 12))
    if start is None:
        today = datetime.date.today()
        start = today.year * 12 + today.month
        warnings.append("지출 기간 시작 연월을 넣어 주세요(지금은 이번 달부터로 계산).")
    idxs = [start + i for i in range(n)]
    units = int(_n(fv("세대수"), 0) or 0)
    metro = str(fv("지역") or "수도권") != "비수도권"

    D = _ym(f.get("오픈")) if "mh" in kinds else None          # 운영 구분(run)은 오픈 일정을 쓰지 않는다
    if "mh" in kinds and D is None:
        warnings.append("오픈월을 넣어 주세요. 이미 오픈해 운영 중인 현장이면 현장 구분을 「견본주택 운영」으로 고르세요(지금은 운영 기준으로 계산).")
    pre_n = int(_n(fv("사전영업"), 1) or 0)
    pre_start = (D - pre_n) if D else None
    mh_end = _ym(f.get("MH종료")) or INF
    org1, org2 = _ym(f.get("조직시작")), (_ym(f.get("조직종료")) or INF)
    chk, mv = (_ym(f.get("사전점검")), _ym(f.get("입주"))) if "movein" in kinds else (None, None)
    if "movein" in kinds and chk is None and mv is None:
        warnings.append("입주 현장은 사전점검월 또는 입주개시월을 넣어 주세요.")
    jg = int(_n(fv("지정기간"), 2) or 0)
    of1 = _ym(f.get("사무실시작")) if "office" in kinds else None
    of2 = _ym(f.get("사무실종료")) or INF
    if "office" in kinds and of1 is None:
        warnings.append("분양사무실 운영 시작월을 넣어 주세요.")
    if not kinds:
        warnings.append("현장 구분(견본주택 오픈 / 견본주택 운영 / 입주 / 분양사무실)을 고르세요.")

    mlist = data.get("months") or []

    def minfo(i):
        return mlist[i] if i < len(mlist) and isinstance(mlist[i], dict) else {}

    # 숙소
    lodging = int(_n(fv("숙소호실"), 0) or 0) if mhk else 0
    S = _ym(f.get("숙소시작")) or pre_start
    E = _ym(f.get("숙소종료")) or mh_end
    monthly = yn("숙소월납")
    lod_existing = False
    if lodging and S is None:                     # 운영 구분(오픈 일정 없음): 이미 쓰는 숙소 — 기간 첫 달부터 월납
        S, lod_existing = idxs[0], True
    lod_months = int(_n(K.get("숙소_개월"), 3) or 3)
    lod_price = _n(K.get("숙소비_월"), 1000)
    mv_lodging = int(_n(fv("입주숙소"), 0) or 0) if "movein" in kinds else 0

    # 단계 깃발
    stages = []
    for i, t in enumerate(idxs):
        ev = set(minfo(i).get("events") or [])
        k = (t - D) if D is not None else None
        mh = mhk and t <= mh_end and (D is None or pre_start <= t)
        st = dict(
            k=k, pre=(D is not None and pre_start <= t < D), open=(k == 0), open2=(k in (0, 1)),
            after3=(k is not None and 2 <= k <= 4), mh=mh, mh_open=(mh and (k is None or k >= 0)),
            org=(mh and org1 is not None and org1 <= t <= org2) or ("선착순" in ev and mh),
            chk=(chk is not None and t == chk),
            center=("movein" in kinds and ((chk if chk is not None else mv) is not None)
                    and (chk if chk is not None else mv) <= t <= ((mv if mv is not None else chk) + jg + 1)),
            mk=(t - mv) if mv is not None else None,
            mv_tel=(mv is not None and 0 <= t - mv <= jg),
            office=(of1 is not None and of1 <= t <= of2), events=ev, idx=t,
        )
        stages.append(st)

    # ── 줄 만들기 ──
    cells = {key: [_Cell() for _ in idxs] for key, _, _ in ROWS}
    p_open, p_etc = _n(fv("식대단가오픈"), 20), _n(fv("식대단가기타"), 10)
    n_open, n_etc, n_rem = (int(_n(fv("식대인원오픈"), 5) or 0), int(_n(fv("식대인원기타"), 4) or 0),
                            int(_n(fv("식대인원잔여"), 3) or 0))
    pcs, copiers = int(_n(fv("PC"), 6) or 0), int(_n(fv("복합기"), 1) or 0)
    printers = int(_n(fv("프린터"), _n(K.get("프린터_대수"), 5)) or 0)
    unit_cnt = int(_n(fv("유니트"), 3) or 3)
    bank = (_n(fv("은행식대단가"), 10), int(_n(fv("은행식대일수"), 7) or 0), int(_n(fv("은행식대인원"), 20) or 0))

    def clean_n(c):
        return c if c and c > 0 else 0

    meal_ppl, clean_ppl = [None] * n, [None] * n
    for i, (t, st) in enumerate(zip(idxs, stages)):
        mi = minfo(i)
        y, m = _split(t)
        ev = st["events"]
        # 식대
        c = cells["식대"][i]
        if st["mh"] or st["org"] or st["center"]:
            if st["open2"]:
                price, people, why = p_open, n_open, "오픈"
            elif st["org"]:
                price, people, why = p_open, n_rem, "잔여판촉"
            else:
                price, people, why = p_etc, n_etc, "기타"
            if any(e in ev for e in ("유상옵션", "명의변경")) and why == "기타":
                people = max(people, n_etc)
            ov = _n(mi.get("식대인원"), None)
            if ov is not None:
                people = int(ov)
            meal_ppl[i] = people
            if price and people:
                c.add(f"{_won(price)}*{_num(people)}*30", f"{_c(price)}/일×{people}명×30일")
        if "중도금대출" in ev and bank[0] and bank[1] and bank[2]:
            c.add(f"{_won(bank[0])}*{bank[1]}*{bank[2]}", f"은행식대({_c(bank[0])}/일×{bank[1]}일×{bank[2]}명)")
        # 숙박비
        c = cells["숙박비"][i]
        if lodging and S is not None and S <= t <= E:
            if lod_existing:
                c.add(f"{_won(lod_price)}*{lodging}", f"기존 숙소 : {_c(lod_price)}/월×{lodging}호")
            elif monthly:
                c.add(f"{_won(lod_price)}*{lodging}", f"상주인력 숙소 : {_c(lod_price)}/월×{lodging}호")
            elif t == S:
                c.add(f"{_won(lod_price)}*{lod_months}*{lodging}", f"상주인력 숙소 : {_c(lod_price)}/월×{lod_months}개월(일시불)×{lodging}호")
            elif t >= S + lod_months:
                c.add(f"{_won(lod_price)}*{lodging}", f"기존 숙소 : {_c(lod_price)}/월×{lodging}호")
        if mv_lodging and mv is not None and t == mv:
            pr = _n(K.get("숙소비_입주_월"), 1000)
            c.add(f"{_won(pr)}*{lod_months}*{mv_lodging}", f"입주 지원인력 숙소 : {_c(pr)}/월×{lod_months}개월(일시불)×{mv_lodging}호")
        # 숙소관리비
        c = cells["숙소관리비"][i]
        mg = _n(K.get("숙소관리비"), 200)
        if lodging and S is not None and S <= t <= E:
            c.add(f"{_won(mg)}*{lodging}", f"상주인력 숙소 관리비 : {_c(mg)}/월×{lodging}호")
        if mv_lodging and mv is not None and mv <= t < mv + lod_months:
            c.add(f"{_won(mg)}*{mv_lodging}", f"입주 숙소 관리비 : {_c(mg)}/월×{mv_lodging}호")
        # 사무용가구
        c = cells["사무용가구"][i]
        if yn("가구") and D is not None and t == (pre_start if pre_n >= 1 else D) and st["mh"]:
            c.add(_won(K.get("가구", 4000)), "책상, 수납장, 의자, 파티션, 이젤 등")
        # 소모품
        c = cells["소모품"][i]
        if st["mh"]:
            k = st["k"]
            if k == 0:
                c.add(_won(K.get("소모품_D", 4000)), "오픈(D) 소모품 — 사무용품, 다과, 생수, 청소비품 등")
            elif k == 1:
                c.add(_won(K.get("소모품_D1", 3000)), "D+1개월 소모품 — 사무용품, 다과, 생수 등")
            else:
                c.add(_won(K.get("소모품_D2", 1000)), "사무용품, 다과, 생수 등 M/H 운영 소모품 구입")
        elif st["center"]:
            c.add(_won(K.get("소모품_입주점검", 1000)), "사무용품, 소모품 구입(입주·사전점검)")
        elif st["office"]:
            c.add(_won(K.get("사무실_소모품", 300)), "분양사무실 소모품")
        # 집기비품
        c = cells["집기비품"][i]
        if st["mh"]:
            items = []
            if copiers:
                c.add(f"{_won(K.get('복합기', 300))}*{copiers}")
                items.append(f"복합기({copiers})")
            if pcs:
                c.add(f"{_won(K.get('PC', 100))}*{pcs}")
                items.append(f"PC({pcs})")
            if st["mh_open"] and printers:
                c.add(f"{_won(K.get('프린터', 200))}*{printers}")
                items.append(f"계약용 프린터({printers})")
            if st["mh_open"] and yn("파쇄기"):
                c.add(_won(K.get("파쇄기", 300)))
                items.append("파쇄기(1)")
            if c.terms:
                c.terms[-1] = (c.terms[-1][0], c.terms[-1][1], ", ".join(items))
        elif st["center"]:
            items = []
            if copiers:
                c.add(f"{_won(K.get('복합기', 300))}*{copiers}")
                items.append(f"복합기({copiers})")
            ipc = int(_n(fv("입주PC"), 0) or 0)
            if ipc:
                c.add(f"{_won(K.get('PC', 100))}*{ipc}")
                items.append(f"PC({ipc})")
            if c.terms:
                c.terms[-1] = (c.terms[-1][0], c.terms[-1][1], ", ".join(items))
        elif st["office"]:
            pr = _n(fv("사무실복합기"), _n(K.get("사무실_복합기"), 300))
            c.add(_won(pr), f"복합기(1) {_c(pr)}/월")
        # 발급수수료 · 송금수수료
        c = cells["발급수수료"][i]
        if st["mh_open"] or st["center"]:
            c.add(_won(K.get("발급_기본", 30)), "등본 및 제증명서류")
        elif st["office"]:
            c.add(_won(K.get("사무실_발급", 30)), "등본 및 제증명서류")
        if "중도금대출" in ev and units:
            c.add(f"{_num(K.get('확정일자', 600))}*{units}", f"확정일자(대출확인서 {_num(K.get('확정일자', 600))}원×{units:,}세대)")
        if yn("송금수수료"):
            c = cells["송금수수료"][i]
            if st["mh_open"] or st["center"] or st["office"]:
                c.add(_won(K.get("송금", 30)), "송금수수료")
        # 보안경비용역
        c = cells["보안경비용역"][i]
        if st["mh"]:
            if yn("보안공사") and pre_start is not None and t == pre_start:
                c.add(_won(K.get("무인경비_공사", 2000)), f"무인경비 공사비({_c(K.get('무인경비_공사', 2000))})")
            c.add(_won(K.get("무인경비_용역", 400)), f"무인경비 용역비({_c(K.get('무인경비_용역', 400))})")
        elif st["office"]:
            pr = _n(fv("사무실보안"), _n(K.get("사무실_보안"), 300))
            c.add(_won(pr), f"보안경비 용역비({_c(pr)})")
        # 용역-기타
        c = cells["용역-기타"][i]
        if st["mh"]:
            c.add(_won(K.get("전기안전", 350)), f"전기안전({_c(K.get('전기안전', 350))})")
            c.add(_won(K.get("소방안전", 350)), f"소방안전({_c(K.get('소방안전', 350))})")
            if yn("오수관리"):
                c.add(_won(K.get("오수관리", 450)), f"오수관리({_c(K.get('오수관리', 450))})")
            k = st["k"]
            if st["pre"]:
                ppl = K.get("청소_이후", 1)
            elif k == 0:
                ppl = K.get("청소_1개월_4이상", 4) if unit_cnt >= 4 else K.get("청소_1개월_3이하", 3)
            elif k == 1:
                ppl = K.get("청소_2개월_4이상", 3) if unit_cnt >= 4 else K.get("청소_2개월_3이하", 2)
            elif st["org"]:
                ppl = K.get("청소_잔여", 2)
            else:
                ppl = K.get("청소_이후", 1)
            ov = _n(mi.get("청소인원"), None)
            ppl = int(ov) if ov is not None else int(_n(ppl, 1))
            clean_ppl[i] = ppl
            if ppl > 0:
                c.add(f"{_won(K.get('청소단가', 3500))}*{ppl}", f"청소용역({ppl}명, {_c(K.get('청소단가', 3500))}/인)")
        # 숙소 중개수수료(기타)
        c = cells["숙소중개"][i]
        fee = _n(fv("숙소중개"), 0)
        if fee and lodging and S is not None and t == S:
            c.add(f"{_won(fee)}*{lodging}", f"숙소 부동산 중개수수료({lodging}개 호실, {_c(fee)}/호실)")
        # 보험료
        c = cells["보험료"][i]
        if yn("보험료") and st["mh"] and D is not None and t == (pre_start if pre_n >= 1 else D):
            c.add(_won(K.get("보험료", 10000)), "견본주택 재산종합보험(1년)")
        # 도서인쇄비
        c = cells["도서인쇄비"][i]
        amt = _n(fv("도서인쇄"), 0)
        if amt and t == ((pre_start if (pre_start is not None and pre_start in idxs) else idxs[0])):
            c.add(_won(amt), "대봉투, 영수증 등 인쇄비")
        # 수도광열비
        c = cells["상하수도"][i]
        if st["mh"]:
            c.add(_won(K.get("상하수도", 500)), f"M/H 상하수도 요금 {_c(K.get('상하수도', 500))}/월")
        if yn("전기료") and st["mh"]:
            c = cells["전기료"][i]
            if m in (6, 7, 8, 12, 1, 2):
                c.add(_won(K.get("전기_동하절기", 10000)), f"동·하절기 전기료 {_c(K.get('전기_동하절기', 10000))}/월")
            else:
                c.add(_won(K.get("전기_기타", 6000)), f"전기료 {_c(K.get('전기_기타', 6000))}/월")
        if yn("난방비") and st["mh"] and (m >= 11 or m <= 3):
            cells["난방비"][i].add(_won(K.get("난방비", 1000)), f"난방비 {_c(K.get('난방비', 1000))}/월")
        # 통신비
        c = cells["통신비"][i]
        if st["mh"]:
            c.add(_won(K.get("통신_MH", 600)), f"전화/인터넷({_c(K.get('통신_MH', 600))})")
            c.add(_won(K.get("우편_기본", 500)), f"우편·SMS({_c(K.get('우편_기본', 500))})")
        elif st["mv_tel"] or st["center"]:
            c.add(_won(K.get("통신_입주", 500)), f"전화/인터넷({_c(K.get('통신_입주', 500))})")
            c.add(_won(K.get("우편_기본", 500)), f"우편·SMS({_c(K.get('우편_기본', 500))})")
        elif st["office"]:
            c.add(_won(K.get("사무실_통신", 200)), f"전화/인터넷({_c(K.get('사무실_통신', 200))})")
            c.add(_won(K.get("사무실_우편", 50)), f"우편({_c(K.get('사무실_우편', 50))})")
        reg = K.get("등기", 4.5)
        if st["open"] and units:
            ratio = _n(fv("예비비율"), 5)
            cnt = int(round(units * ratio))
            if cnt:
                c.add(f"{_won(K.get('내용증명', 5))}*{cnt}", f"예비입주자 내용증명({_c(K.get('내용증명', 5))}×{cnt:,}건, 세대수의 {ratio * 100:.0f}%)")
        if "중도금대출" in ev and units:
            c.add(f"{_won(reg)}*{units}", f"중도금대출 안내문({_c(reg)}×{units:,}건)")
        if "유상옵션" in ev and units and yn("옵션안내"):
            c.add(f"{_won(reg)}*{units}", f"유상옵션 안내문({_c(reg)}×{units:,}건)")
        if units and "movein" in kinds:
            cnt_notice = int(_n(fv("입주안내"), 2) or 0)
            if chk is not None and t == chk - 1 and cnt_notice >= 1:
                c.add(f"{_won(reg)}*{units}", f"사전점검 안내문({_c(reg)}×{units:,}건)")
            if mv is not None and t == mv - 1 and cnt_notice >= 2:
                c.add(f"{_won(reg)}*{units}", f"입주 안내문({_c(reg)}×{units:,}건)")
            if mv is not None and t == mv - 1 and yn("어린이집"):
                c.add(f"{_won(K.get('어린이집', 4.5))}*{units}", f"국공립어린이집 동의서({_c(K.get('어린이집', 4.5))}×{units:,}건)")
        # 운반비
        c = cells["운반비"][i]
        if yn("운반비"):
            mv_fee = K.get("운반비", 2500)
            if st["mh"] and pre_start is not None and t == pre_start:
                c.add(_won(mv_fee), "M/H 집기 이동(포장이사)")
            if mv is not None and t == mv and "movein" in kinds:
                c.add(_won(mv_fee), "입주지원센터 이전")
            if of1 is not None and t == of1:
                c.add(_won(mv_fee), "분양사무실 이전")
            if "사무실이전" in ev and not c.terms:
                c.add(_won(mv_fee), "사무실 이전")
        # 여비교통비
        c = cells["여비교통비"][i]
        if st["open2"] and st["mh"]:
            v = K.get("여비_오픈_수도권" if metro else "여비_오픈_비수도권", 2000)
            c.add(_won(v), "유류대, 숙박비, 통행료 등")
        elif st["mh"] and (st["pre"] or st["after3"]):
            v = K.get("여비_사전_수도권" if metro else "여비_사전_비수도권", 1500)
            c.add(_won(v), "유류대, 숙박비, 통행료 등")
        elif st["mh"] or st["org"] or st["center"]:
            v = K.get("여비_기타_수도권" if metro else "여비_기타_비수도권", 1000)
            c.add(_won(v), "유류대, 숙박비, 통행료 등")
        # 관리비(분양사무실)
        c = cells["관리비"][i]
        if st["office"]:
            if yn("통합센터"):
                c.add(_won(K.get("사무실_관리비_통합", 1000)), "분양사무실 관리비(통합센터)")
            else:
                c.add(_won(K.get("사무실_관리비", 600)), "분양사무실 일반/공용관리비")

    # ── 직접 고친 칸·비고, 추가 줄 ──
    overrides = data.get("overrides") or {}
    notes_ov = data.get("notes") or {}
    rows_out = []
    for key, group, name in ROWS:
        cl = cells[key]
        ov = overrides.get(key) or {}
        out_cells, manual = [], False
        for i, c in enumerate(cl):
            o = ov.get(str(i), ov.get(i)) if isinstance(ov, dict) else None
            if o not in (None, ""):
                manual = True
                if isinstance(o, str) and o.strip().startswith("="):
                    expr = o.strip()[1:].replace(" ", "")
                    v = evaluate(expr)
                    out_cells.append(dict(v=v if v is not None else 0, f=expr if v is not None else None, note=c.note, manual=True,
                                          auto=c.value, autof=c.formula))
                else:
                    out_cells.append(dict(v=_n(o, 0), f=None, note=c.note, manual=True, auto=c.value, autof=c.formula))
            else:
                out_cells.append(dict(v=c.value, f=c.formula, note=c.note, manual=False))
        note = notes_ov.get(key)
        auto_note = _join_notes(cl, idxs)
        if any(x["v"] for x in out_cells) or manual:
            rows_out.append(dict(key=key, item=group, name=name, cells=out_cells, total=sum(x["v"] for x in out_cells),
                                 note=note if note not in (None, "") else auto_note, auto_note=auto_note,
                                 note_manual=note not in (None, ""), manual=manual))
    for j, ex in enumerate(data.get("extra") or []):
        if not isinstance(ex, dict):
            continue
        item = ex.get("item") if ex.get("item") in GROUPS else "지급수수료"
        vals = list(ex.get("vals") or [])
        out_cells = []
        for i in range(n):
            o = vals[i] if i < len(vals) else None
            if isinstance(o, str) and o.strip().startswith("="):
                expr = o.strip()[1:].replace(" ", "")
                v = evaluate(expr)
                out_cells.append(dict(v=v or 0, f=expr if v is not None else None, note="", manual=True))
            else:
                out_cells.append(dict(v=_n(o, 0), f=None, note="", manual=True))
        rows_out.append(dict(key=f"extra:{j}", item=item, name=ex.get("name") or "기타", cells=out_cells,
                             total=sum(x["v"] for x in out_cells), note=ex.get("note") or "", auto_note="", note_manual=True,
                             manual=True, extra=True))

    # ── 묶음·합계 ──
    groups = []
    for g in GROUPS:
        rs = [r for r in rows_out if r["item"] == g]
        if not rs:
            continue
        months_sum = [sum(r["cells"][i]["v"] for r in rs) for i in range(n)]
        groups.append(dict(item=g, rows=rs, months=months_sum, total=sum(months_sum), subtotal=len(rs) >= 2,
                           merge=(len(rs) == 1 and g in SINGLE_MERGE)))
    total_m = [sum(g["months"][i] for g in groups) for i in range(n)]

    # ── 머리글·제목 ──
    months_out = []
    for i, (t, st) in enumerate(zip(idxs, stages)):
        y, m = _split(t)
        parts = []
        if st["pre"]:
            parts.append("사전")
        if st["open"]:
            parts.append("오픈")
        if st["k"] == 1 and st["mh"]:
            parts.append("계약")
        if st["chk"]:
            parts.append("사전점검")
        if st["mk"] is not None and 0 <= st["mk"] <= jg and "movein" in kinds:
            parts.append("입주")
        if of1 is not None and t == of1:
            parts.append("분양사무실 이전")
        for e, lb, _ in EVENTS:
            if e in st["events"] and lb not in parts:
                parts.append(lb)
        auto = f"({','.join(parts)})" if parts else ""
        lab = minfo(i).get("label")
        stage_names = [k for k in ("pre", "open", "org", "mh", "center", "office") if st[k]]
        months_out.append(dict(y=y, m=m, idx=t, k=st["k"], label=lab if lab not in (None, "") else auto, auto_label=auto,
                               stage=stage_names, events=sorted(st["events"]), 식대인원=meal_ppl[i], 청소인원=clean_ppl[i]))
    any_pre_open = any(s["pre"] or (s["open2"] and s["mh"]) for s in stages)
    if any_pre_open:
        kind = "M/H 오픈비용"
    elif any(s["mh"] for s in stages):
        kind = "M/H 운영비용"
    elif any(s["center"] or s["mv_tel"] for s in stages):
        kind = "입주비용"
    else:
        kind = "운영비용"
    site = (f.get("현장명") or "현장명").strip()
    title = data.get("title") or f.get("제목") or f"[{site}] {kind} 지출계획서"
    d = f.get("작성일")
    try:
        dt = datetime.datetime.strptime(str(d)[:10], "%Y-%m-%d") if d else datetime.datetime.now()
    except ValueError:
        dt = datetime.datetime.now()
    date_txt = f"( 작성일 : {dt.year}년 {dt.month:02d}월 {dt.day:02d}일 )"
    approvers = approver_box(data.get("approvers"))
    if units == 0 and (mhk or "movein" in kinds):
        warnings.append("총세대수가 비어 있어 내용증명·확정일자·안내문 등기 건수는 0으로 계산했습니다.")
    return dict(title=title, kind=kind, site=site, date_txt=date_txt, date=dt.strftime("%Y-%m-%d"), approvers=approvers,
                months=months_out, groups=groups, total=total_m, total_all=sum(total_m), warnings=warnings,
                n=n, units=units)


def period_text(res):
    """('26.10_12월) — 해가 바뀌면 ('26.11_'27.01월)"""
    ms = res["months"]
    a, b = ms[0], ms[-1]
    if a["y"] == b["y"]:
        return f"('{a['y'] % 100:02d}.{a['m']:02d}_{b['m']:02d}월)"
    return f"('{a['y'] % 100:02d}.{a['m']:02d}_'{b['y'] % 100:02d}.{b['m']:02d}월)"


def file_name(res):
    return f"({res['site']}) {res['kind'].replace('M/H ', '')} 지출계획서{period_text(res)}.xlsx"
