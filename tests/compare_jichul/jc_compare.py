# -*- coding: utf-8 -*-
"""과거 지출계획서(현장×기간) 재현 검증: 원본 엑셀 → 현장 조건 추정 → jichul.calc(그 해 기준단가) → 칸 단위 비교 → 사례별 xlsx + 요약 엑셀.
  python jc_compare.py            (jc_parse.py 가 만든 jc_inventory.json, K_years.json 이 같은 폴더에 있어야 함)
"""
import datetime
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
import openpyxl                     # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill   # noqa: E402
from openpyxl.utils import get_column_letter as L          # noqa: E402
import jichul                       # noqa: E402
import jichul_xlsx                  # noqa: E402

OUT = os.path.join(ROOT, "tests", "compare_jichul", "out")
REPORT = os.path.join(ROOT, "tests", f"지출계획서비교_테스트결과_{datetime.date.today():%Y-%m-%d}.xlsx")
B = r"E:\1. 우미"
S = B + r"\2. 수도권남부"

# ── 사례 목록: (현장 표시, 기간 표시, 파일, 지역, 메모) — 같은 기간이 여러 판이면 가장 나중/결재본에 가까운 것
CASES = [
    ("오산 서2", "'25.04~06", S + r"\13. 오산 서2지구\지출계획서\(오산 서2) 오픈비용 지출계획서('25.04_06, 사전영업).xlsx", "수도권", ""),
    ("오산 서2", "'25.07~09", S + r"\13. 오산 서2지구\지출계획서\(오산 서2) 오픈비용 지출계획서('25.07_09월)_수정안.xlsx", "수도권", "수정안(7·8 사전, 9 오픈)"),
    ("오산 서2", "'25.10~12", S + r"\13. 오산 서2지구\지출계획서\(오산 서2) 오픈비용 지출계획서('25.10_12월)_250930.xlsx", "수도권", ""),
    ("오산 서2", "'26.01~03", S + r"\13. 오산 서2지구\지출계획서\(오산 서2) 운영비용 지출계획서('26.1_3월)_251209.xlsx", "수도권", ""),
    ("오산 서2", "'26.04~06", S + r"\13. 오산 서2지구\지출계획서\(오산 서2) 운영비용 지출계획서('26.4_6월)_260401.xlsx", "수도권", ""),
    ("오산 서2", "'26.07~09", S + r"\13. 오산 서2지구\지출계획서\(오산 서2) 운영비용 지출계획서('26.7~9월)_260720.xlsx", "수도권", "분양사무실 이전"),
    ("부산장안", "'25.02~04", S + r"\13. 오산 서2지구\지출계획서\부산\(부산장안) 오픈비용 지출계획서('25.02_04월)-결재전.xlsx", "비수도권", "결재전 판"),
    ("오산세교1차", "'24.08~10", S + r"\10. 오산세교 1차A14블럭(후분양)\★ 지출계획서\(오산세교1차) 오픈비용 지출계획서('24.08_10월) 5.xlsx", "수도권", ""),
    ("오산세교1차", "'24.11~'25.01", S + r"\10. 오산세교 1차A14블럭(후분양)\★ 지출계획서\(오산세교1차) 오픈비용 지출계획서('24.11월_'25년01월).xlsx", "수도권", ""),
    ("오산세교1차", "'25.01~03", S + r"\10. 오산세교 1차A14블럭(후분양)\★ 지출계획서\(오산세교1차) 오픈비용 지출계획서('25년01_03월) - 복사본.xlsx", "수도권", ""),
    ("오산세교 2차(사전청약)", "'21.11~'22.01", S + r"\10. 오산세교 1차A14블럭(후분양)\★ 오산세교2차A14블럭(후분양)\예산\(오산세교) MH오픈비용 지출계획서(11~1월)_20211122-DESKTOP-SSJR523.xlsx", "수도권", "사전청약 M/H(특수)"),
    ("평택고덕1차", "'26.05~07", S + r"\14. 평택고덕 패키지형 공모사업\★ 1차(A36BL)\★ 지출계획서\(평택고덕 1차) MH 오픈비용 지출계획서(안)_260330.xlsx", "수도권", "안"),
    ("평택고덕1차", "'26.08~10", S + r"\14. 평택고덕 패키지형 공모사업\★ 1차(A36BL)\★ 지출계획서\(평택고덕 1차) MH 운영비용 지출계획서(안)_260731.xlsx", "수도권", "안(0731)"),
    ("평택고덕2차", "'26.08~10(안)", S + r"\14. 평택고덕 패키지형 공모사업\★ 2차(A33BL)\★ 지출계획서\(평택고덕 2차) 오픈비용 지출계획서(안)_260803.xlsx", "수도권", "8월 안(결재본은 10~12월 PDF)"),
    ("과천지식정보", "'21.07~09", S + r"\2. 과천 지식정보타운 S8\★★ 지출계획서 관련\★지출계획서(오픈, 운영비용)\(과천지식정보) MH오픈비용 지출계획서(7~9월)_20210630 - 복사본 - 복사본.xlsx", "수도권", ""),
    ("과천지식정보", "'21.10~12", S + r"\2. 과천 지식정보타운 S8\★★ 지출계획서 관련\★지출계획서(오픈, 운영비용)\(과천지식정보) MH운영비용 지출계획서(10~12월)_20210913_최종(박과장님 수정).xlsx", "수도권", ""),
    ("과천지식정보", "'22.04~06", S + r"\2. 과천 지식정보타운 S8\★★ 지출계획서 관련\★지출계획서(오픈, 운영비용)\(과천지식정보) MH운영비용 지출계획서(4~6월)_20221212.xlsx", "수도권", ""),
    ("화성비봉", "'22.03~05", S + r"\3. 화성비봉\지출계획서\(화성비봉) MH오픈비용 지출계획서(3_5월).xlsx", "수도권", ""),
    ("화성비봉", "'22.08~10", S + r"\3. 화성비봉\지출계획서\(화성비봉) MH오픈비용 지출계획서(8_10월) - 수정본.xlsx", "수도권", "수정본"),
    ("화성비봉", "'24.09~11(입주)", S + r"\3. 화성비봉\지출계획서\(화성비봉) 입주비용 지출계획서('24.09_11월).xlsx", "수도권", ""),
    ("화성비봉", "'25.02~04(입주)", S + r"\3. 화성비봉\지출계획서\(화성비봉) 입주비용 지출계획서('25.02월_04월).xlsx", "수도권", ""),
    ("김포북변", "'24.09~11", S + r"\8. 김포북변(차토)\지출계획서\(김포북변) 운영비용 지출계획서('24.09_11월).xlsx", "수도권", ""),
    ("김포북변", "'24.10~12", S + r"\8. 김포북변(차토)\지출계획서\(김포북변) 운영비용 지출계획서('24.10_12월).xlsx", "수도권", ""),
    ("이천중리1차", "'23.06~08", S + r"\9. 이천 중리\1. 1차(B2블럭)\★ 지출계획서\(이천중리1차) MH오픈비용 지출계획서('23.06_08월)_수정중.xlsx", "수도권", "수정중 판"),
    ("이천중리1차", "'23.09~11", S + r"\9. 이천 중리\1. 1차(B2블럭)\★ 지출계획서\(이천중리1차) MH오픈비용 지출계획서('23.09_11월).xlsx", "수도권", ""),
    ("화성태안1차", "'21.01~03", S + r"\1. 태안1,2블럭\선급금\(화성태안 1차) MH 오픈비용지출계획서-시행사협의-20210125.xlsx", "수도권", "시행사협의 판"),
    ("화성태안1,2차", "'21.02~04", S + r"\1. 태안1,2블럭\선급금\(화성태안 1,2차) MH 오픈비용 지출계획서(안)_2~4월.xlsx", "수도권", ""),
    ("화성태안1,2차", "'21.05~07", S + r"\1. 태안1,2블럭\선급금\(화성태안 1,2차) MH 오픈비용 지출계획서(안)_5~7월-20210525.xlsx", "수도권", ""),
    ("화성태안1,2차", "'21.08~09", S + r"\1. 태안1,2블럭\선급금\(화성태안 1,2차) MH 운영비용 지출계획서(안)_8_9월-최종.xlsx", "수도권", "2개월"),
    ("화성태안2차", "'23.07~09(입주)", S + r"\1. 태안1,2블럭\선급금\(화성태안2차) 입주비용 지출계획서('23.07_08월).xlsx", "수도권", ""),
    ("인천검단2차", "'20.03~05", S + r"\이전\12. 검단\ㄱ검단2차\예산\★(인천검단 2차) MH 오픈비용 지출계획서(안)-(20.03.03)동(123).xlsx", "수도권", ""),
    ("인천검단2차", "'20.06~08", S + r"\이전\12. 검단\ㄱ검단2차\예산\지출계획서\(인천검단 2차) MH 운영비용 지출계획서(안)-200622.xlsx", "수도권", ""),
    ("인천검단2차", "'20.09~10", S + r"\이전\12. 검단\ㄱ검단2차\예산\지출계획서\(인천검단 2차) MH 운영비용 지출계획서(안)(20.09_11월)-200921.xlsx", "수도권", "2개월"),
    ("인천검단3,5차", "'20.11~'21.01", S + r"\이전\12. 검단\검단3,5차\예산(선급금)\지출계획서\(인천검단 3,5차) MH 운영비용 지출계획서(안) -201118.xlsx", "수도권", ""),
    ("인천검단3,5차", "'21.02~04", S + r"\이전\12. 검단\검단3,5차\예산(선급금)\지출계획서\(인천검단 3,5차) MH 오픈비용 지출계획서(안)_2~4월.xlsx", "수도권", ""),
    ("인천검단3,5차", "'21.05~07", S + r"\이전\12. 검단\검단3,5차\예산(선급금)\지출계획서\(인천검단 3,5차) MH 오픈비용 지출계획서(안)_5_7월-20210525.xlsx", "수도권", ""),
    ("인천검단3,5차", "'21.08~10", S + r"\이전\12. 검단\검단3,5차\예산(선급금)\지출계획서\(인천검단 3,5차) MH 운영비용 지출계획서(안)_8_10월-최종-20210827.xlsx", "수도권", ""),
    ("위례신도시2차", "'20.04~06", S + r"\종료현장\5. 위례 2차\지출계획서\(위례신도시 2차) MH오픈비용 지출계획서(4~6월).xlsx", "수도권", ""),
    ("위례신도시2차", "'20.07~09", S + r"\종료현장\5. 위례 2차\지출계획서\★(위례신도시 2차) MH운영비용 지출계획서(7~9월)-권.xlsx", "수도권", ""),
    ("위례신도시2차", "'20.10~12", S + r"\종료현장\5. 위례 2차\지출계획서\★(위례신도시 2차) MH 운영비용 지출계획서(안)(20.10)-권.xlsx", "수도권", ""),
    ("위례신도시2차", "'22.10~12(입주)", S + r"\종료현장\5. 위례 2차\지출계획서\(위례신도시 2차) 입주비용 지출계획서(10_12월).xlsx", "수도권", ""),
    ("위례신도시2차", "'23.01~03(입주)", S + r"\종료현장\5. 위례 2차\지출계획서\(위례신도시 2차) 입주비용 지출계획서('23.01_02월).xlsx", "수도권", ""),
    ("양주옥정(타팀)", "'21.03~05", S + r"\2. 과천 지식정보타운 S8\★★ 지출계획서 관련\★지출계획서(오픈, 운영비용)\타현장비교\(양주 옥정) MH 오픈비용 지출계획서(안)_21.03-05.xlsx", "수도권", "타팀 참고 파일"),
    ("양주옥정(타팀)", "'21.06~08", S + r"\2. 과천 지식정보타운 S8\★★ 지출계획서 관련\★지출계획서(오픈, 운영비용)\타현장비교\(양주옥정) MH 운영비용 지출계획서(안)_21.06.14.xlsx", "수도권", "타팀 참고 파일"),
    ("완주삼봉(타팀)", "'20.12~'21.02", B + r"\기타\문서\카카오톡 받은 파일\(완주 삼봉) MH 운영비용 지출계획서(안)_20.12월~21.02월.xlsx", "비수도권", "타팀 참고 파일"),
    ("경산하양(타팀)", "'21.05~08(입주)", S + r"\종료현장\5. 위례 2차\★ 사전점검\타현장\인천검단2차(사전점검)\인천검단2차 기안\1.샘플\★(경산하양) 입주비용 지출계획서(안)-권.xlsx", "비수도권", "타팀 참고 파일, 4개월"),
    ("동탄지식산업센터(홍보관)", "'20.08~10", B + r"\1. 본사업무\☆ 시장현황\★ 지식산업센터\당사분양\동탄우미뉴브\지출계획\200820 (동탄 지식산업센터) 오픈비용 지출계획서(20.08~10월).xlsx", "수도권", "지식산업센터 홍보관(아파트 M/H 아님)"),
    ("동탄지식산업센터(홍보관)", "'20.10~12", B + r"\1. 본사업무\☆ 시장현황\★ 지식산업센터\당사분양\동탄우미뉴브\지출계획\201015 (동탄 지식산업센터) 오픈비용 지출계획서(20.10~12월).xlsx", "수도권", "지식산업센터 홍보관"),
    ("동탄지식산업센터(홍보관)", "'21.01~03", B + r"\1. 본사업무\☆ 시장현황\★ 지식산업센터\당사분양\동탄우미뉴브\지출계획\210119 (동탄 지식산업센터) 오픈비용 지출계획서(21.1_3월).xlsx", "수도권", "지식산업센터 홍보관"),
    ("동탄지식산업센터(홍보관)", "'21.04~06", B + r"\1. 본사업무\☆ 시장현황\★ 지식산업센터\당사분양\동탄우미뉴브\지출계획\210412 (동탄 지식산업센터) 오픈비용 지출계획서(21.4_6월).xlsx", "수도권", "지식산업센터 홍보관"),
    ("동탄지식산업센터(홍보관)", "'21.07~09", B + r"\1. 본사업무\☆ 시장현황\★ 지식산업센터\당사분양\동탄우미뉴브\지출계획\210715 (동탄 지식산업센터) 오픈비용 지출계획서(21.7_9월).xlsx", "수도권", "지식산업센터 홍보관"),
    ("동탄지식산업센터(홍보관)", "'21.10~12", B + r"\1. 본사업무\☆ 시장현황\★ 지식산업센터\당사분양\동탄우미뉴브\지출계획\211018 (동탄 지식산업센터) 오픈비용 지출계획서(21.10~12월).xlsx", "수도권", "지식산업센터 홍보관"),
]
# 스캔 PDF 만 있는 사례(결재본) — Windows OCR 로 읽음(참고용, 검산 칸으로 신뢰도 표시)
CASES_PDF = [
    ("평택고덕2차", "'26.10~12(결재본)", S + r"\14. 평택고덕 패키지형 공모사업\★ 2차(A33BL)\★ 지출계획서\(평택고덕2차) 오픈비용 지출계획서(26.10~12).pdf", "수도권", "결재본 PDF(OCR)"),
    ("평택고덕1차", "'26.05~07(결재본)", S + r"\14. 평택고덕 패키지형 공모사업\★ 1차(A36BL)\★ 지출계획서\(평택고덕1차) 오픈비용 지출계획서_5~7월.pdf", "수도권", "결재본 PDF(OCR)"),
    ("평택고덕1차", "'26.08~10(결재본)", S + r"\14. 평택고덕 패키지형 공모사업\★ 1차(A36BL)\★ 지출계획서\(평택고덕1차) 운영비용 지출계획서_8~10월.pdf", "수도권", "결재본 PDF(OCR)"),
    ("오산세교1차", "'24.02~04", S + r"\10. 오산세교 1차A14블럭(후분양)\★ 지출계획서\[오산세교1차] 운영비용 지출계획서(02_04월).pdf", "수도권", "PDF(OCR)"),
    ("과천지식정보", "'22.01~03", S + r"\2. 과천 지식정보타운 S8\★★ 지출계획서 관련\★지출계획서(오픈, 운영비용)\(과천지식정보) MH운영비용 지출계획서 결재본(22.01월-03월).pdf", "수도권", "결재본 PDF(OCR)"),
    ("과천지식정보", "'24.02~04(입주)", S + r"\2. 과천 지식정보타운 S8\★★ 지출계획서 관련\★지출계획서(오픈, 운영비용)\[과천지식정보] 입주비용 지출계획서(02월_04월).pdf", "수도권", "PDF(OCR)"),
    ("과천지식정보", "'24.05~07(입주)", S + r"\2. 과천 지식정보타운 S8\★ 입주\(과천지식정보) 입주비용 지출계획서(5~7월).pdf", "수도권", "PDF(OCR)"),
    ("과천지식정보", "'24.08~10(입주)", S + r"\2. 과천 지식정보타운 S8\★★ 지출계획서 관련\★지출계획서(오픈, 운영비용)\[과천지식정보] 입주비용 지출계획서(08월_10월).pdf", "수도권", "PDF(OCR)"),
    ("화성비봉", "'22.06~08", S + r"\3. 화성비봉\지출계획서\(화성비봉) MH 운영비용 지출계획서(6-8월).pdf", "수도권", "결재본 PDF(OCR)"),
    ("화성비봉", "'23.11~'24.01(입주)", S + r"\3. 화성비봉\지출계획서\[화성비봉] 입주비용 지출계획서 23.11_24.01 1 1.pdf", "수도권", "PDF(OCR)"),
    ("화성비봉", "'24.02~04(입주)", S + r"\3. 화성비봉\지출계획서\[화성비봉] 입주비용 지출계획서(02월_04월).pdf", "수도권", "PDF(OCR)"),
    ("화성비봉", "'24.10~12(입주)", S + r"\3. 화성비봉\지출계획서\(화성비봉) 입주비용 지출계획서('24.10_12월).pdf", "수도권", "PDF(OCR)"),
    ("화성비봉", "'25.01~03(입주)", S + r"\3. 화성비봉\지출계획서\(화성비봉) 입주비용 지출계획서(25.01_03).pdf", "수도권", "PDF(OCR)"),
    ("이천중리1차", "'23.12~'24.02", S + r"\9. 이천 중리\1. 1차(B2블럭)\★ 지출계획서\(이천중리1차) MH오픈비용 지출계획서('23.12_02월) 2.pdf", "수도권", "PDF(OCR)"),
    ("이천중리1차", "'24.04~06", S + r"\9. 이천 중리\1. 1차(B2블럭)\★ 지출계획서\(이천중리1차) MH오픈비용 지출계획서 24년 4_6월.pdf", "수도권", "PDF(OCR)"),
    ("이천중리2차", "'23.10~12", S + r"\9. 이천 중리\2. 2차(B1블럭)\지출계획서\[이천중리2차] MH오픈비용 지출계획서(10_12월).pdf", "수도권", "PDF(OCR), 1쪽만"),
    ("이천중리2차", "'24.04~06", S + r"\9. 이천 중리\2. 2차(B1블럭)\지출계획서\(이천중리2차) MH오픈비용 지출계획서 24년 4_6월.pdf", "수도권", "PDF(OCR)"),
    ("화성태안2차", "'24.02~04(입주)", S + r"\1. 태안1,2블럭\★ 지출계획서\[화성태안2차] 입주비용 지출계획서(02_04월).pdf", "수도권", "PDF(OCR)"),
    ("위례신도시2차", "'20.07~09(결재본)", S + r"\종료현장\5. 위례 2차\지출계획서\(위례신도시 2차) MH오픈비용 지출계획서(20.07_09).pdf", "수도권", "결재본 PDF(OCR)"),
    ("위례신도시1차", "'19.07~09", S + r"\종료현장\5. 위례 2차\위례신도시1차(A3-4B)\분양비용\(위례신도시 1차) 운영비용 지출계획서(7월~9월).pdf", "수도권", "PDF(OCR)"),
    ("위례신도시1차", "'19.10~12", S + r"\종료현장\5. 위례 2차\위례신도시1차(A3-4B)\분양비용\(위례신도시 1차) MH운영비용 지출계획서(19.10~19.12).pdf", "수도권", "PDF(OCR)"),
    ("위례신도시1차", "'20.01~03", S + r"\종료현장\5. 위례 2차\위례신도시1차(A3-4B)\분양비용\(위례신도시 1차) MH운용비용 지출계획서(1~3월)_결재본.pdf", "수도권", "결재본 PDF(OCR)"),
    ("위례신도시1차", "'21.11~'22.01(입주)", S + r"\종료현장\5. 위례 2차\위례신도시1차(A3-4B)\★입주관련\입주비용\(위례신도시 1차) 입주비용 지출계획서(21.11_22.01월)(결재본).pdf", "수도권", "결재본 PDF(OCR)"),
    ("전주효천2차(타팀)", "'19.11~'20.01(입주)", S + r"\종료현장\5. 위례 2차\위례신도시2차(A3-2)\입주관련\입주비용\(전주효천 2차) 입주비용 지출계획서(19.11~20.01).pdf", "비수도권", "타팀 참고, PDF(OCR)"),
]
# 비고·산식에서 못 읽을 때 쓰는 현장 세대수(과거 파일에서 확인된 값)
SITE_UNITS = {"오산 서2": 1424, "오산세교1차": 1532, "김포북변": 1200, "위례신도시2차": 420, "인천검단2차": 437, "인천검단3,5차": 1200,
              "화성비봉": 798, "화성태안1,2차": 1300, "이천중리1차": 500, "평택고덕2차": 623, "양주옥정(타팀)": 2049}
MEAL_BY_YEAR = {2022: (16, 8), 2023: (16, 8), 2024: (18, 9), 2025: (20, 10), 2026: (20, 10), 2027: (20, 10)}

inv = {r["path"]: r for r in json.load(io.open(os.path.join(SP, "jc_inventory.json"), encoding="utf-8"))}
K_YEARS = {int(y): v for y, v in json.load(io.open(os.path.join(SP, "K_years.json"), encoding="utf-8")).items()}


def norm(s):
    return re.sub(r"\s+", "", str(s)) if s is not None else ""


def nums_in(s):
    return [int(x.replace(",", "")) for x in re.findall(r"\d[\d,]*", s or "")]


# ── 원본 줄 → 도구 줄키
def map_row(item, sub, note):
    s, n = sub or "", note or ""
    if item.startswith("복리후생"):
        if "식대" in s:
            return "식대"
        if "관리" in s:
            return "숙소관리비"
        if "숙박" in s or "숙소" in s:
            return "숙박비"
    if item.startswith("소모품"):
        return "사무용가구" if "가구" in s else "소모품"
    if "임차" in item:
        if "집기" in s or not s:
            return "집기비품"
        return "extra:임차료-" + (s or "기타")
    if "수수료" in item:
        if "발급" in s or "발급" in n:
            return "발급수수료"
        if "송금" in s:
            return "송금수수료"
        if "보안" in s or "경비" in s:
            return "보안경비용역"
        if "용역" in s:
            return "용역-기타"
        if "중개" in n or "중개" in s:
            return "숙소중개"
        return "extra:수수료-" + (s or "기타")
    if "보험" in item:
        return "보험료"
    if "도서" in item:
        return "도서인쇄비"
    if "수도" in item or "광열" in item:
        if "전기" in s or ("전기" in n and "상하수도" not in n):
            return "전기료"
        if "난방" in s or "난방" in n:
            return "난방비"
        return "상하수도"
    if "통신" in item:
        return "통신비"
    if "운반" in item:
        return "운반비"
    if "여비" in item:
        return "여비교통비"
    if "관리비" in item:
        return "관리비"
    return "extra:" + item + ("-" + s if s else "")


def merge_rows(rec):
    """원본 줄들을 도구 줄키로 묶음 → {key: dict(vals, formulas, note, names)}"""
    out = {}
    for r in rec["rows"]:
        key = map_row(r["item"], r["sub"], r["note"])
        o = out.setdefault(key, dict(vals=[0.0] * len(rec["months"]), formulas=[None] * len(rec["months"]), note="", names=[]))
        for i, v in enumerate(r["vals"]):
            o["vals"][i] += v or 0
            if r["formulas"][i] and not o["formulas"][i]:
                o["formulas"][i] = r["formulas"][i]
        o["note"] = (o["note"] + "\n" + r["note"]).strip()
        o["names"].append((r["item"], r["sub"]))
    return out


# ── 현장 조건 추정
def parse_date(s):
    m = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*월(?:\s*(\d{1,2})\s*일)?", s or "")
    if not m:
        return None
    return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3) or 1))


def month_years(rec, written):
    """월 머리글(연 없음) → [(연, 월)]: 작성일 기준, 달이 줄어들면 해 넘김"""
    ms = [m["m"] for m in rec["months"]]
    y = written.year if written else datetime.date.today().year
    if written and ms and ms[0] < written.month - 6:
        y += 1
    if written and ms and ms[0] > written.month + 6:
        y -= 1
    out, prev = [], None
    for m in ms:
        if prev is not None and m < prev:
            y += 1
        out.append((y, m))
        prev = m
    return out


def infer(site, region, rec):
    rows = merge_rows(rec)
    written = parse_date(rec["date"])
    if written is None:      # OCR 로 작성일을 못 읽으면 파일 이름('19.10, 20.01…)·파일 날짜로
        name = rec.get("name") or ""
        m = re.search(r"'?(\d{2})\.(\d{1,2})", name) or re.search(r"(20\d{2})", name)
        if m and len(m.group(1)) == 2:
            written = datetime.date(2000 + int(m.group(1)), int(m.group(2)), 1)
        elif m:
            written = datetime.date(int(m.group(1)), int(rec["months"][0]["m"]), 1)
        else:
            written = datetime.date.fromtimestamp(os.path.getmtime(rec["path"])) if rec.get("path") and os.path.exists(rec["path"]) else None
    yms = month_years(rec, written)
    n = len(yms)
    labels = [norm(m["label"]) for m in rec["months"]]
    title = rec["title"] or ""
    if rec.get("ocr") and "입주" in (rec.get("name") or "") and "입주비용" not in title:
        title = title.replace("운영비용", "입주비용") if "운영비용" in title else title + " 입주비용"
    notes_all = "\n".join(r["note"] for r in rec["rows"])
    year = yms[0][0]
    ky = min(max(year, 2022), 2027)
    K = K_YEARS[ky]
    fields, months, info = {}, [dict(events=[]) for _ in range(n)], []

    # 달별 표시 → 단계·행사
    pre_idx = [i for i, l in enumerate(labels) if "사전" in l and "사전점검" not in l]
    open_idx = [i for i, l in enumerate(labels) if "오픈" in l]
    con_idx = [i for i, l in enumerate(labels) if "계약" in l and "계약용" not in l]
    for i, l in enumerate(labels):
        ev = months[i]["events"]
        if any(k in l for k in ("중도금", "대출", "자서")):
            ev.append("중도금대출")
        if "옵션" in l:
            ev.append("유상옵션")
        if "명변" in l or "명의" in l:
            ev.append("명의변경")
        if "선착순" in l:
            ev.append("선착순")
        if "전매" in l:
            ev.append("전매")
        if "이전" in l:
            ev.append("사무실이전")
    # 비고의 은행식대 / 발급수수료의 확정일자 → 중도금대출(표시가 없을 때, 금액이 가장 큰 달)
    meal = rows.get("식대")
    if not any("중도금대출" in m["events"] for m in months):
        for key, word in (("식대", "은행"), ("발급수수료", "확정일자")):
            r = rows.get(key)
            if r and word in r["note"] and any(r["vals"]):
                i = max(range(n), key=lambda j: r["vals"][j])
                if r["vals"][i] > min(v for v in r["vals"]) and "중도금대출" not in months[i]["events"]:
                    months[i]["events"].append("중도금대출")
                    info.append(f"{yms[i][1]}월 중도금대출({word} 비고에서 추정)")
                break
    # 기간 끝의 빈 달 → M/H(또는 사무실) 종료월
    last_nonzero = max((i for i in range(n) if rec["total"][i]), default=n - 1)
    trailing_zero = last_nonzero < n - 1
    is_movein = "입주비용" in title
    is_office = ("사무실소모품" in norm(notes_all)) or any("분양사무실" in l for l in labels) and "입주" not in title
    kinds = []
    if is_movein:
        kinds = ["movein"]
        chk = [i for i, l in enumerate(labels) if "사전점검" in l]
        mv = [i for i, l in enumerate(labels) if "입주" in l]
        if chk:
            fields["사전점검"] = list(yms[chk[0]])
        if mv:
            fields["입주"] = list(yms[mv[0]])
            fields["지정기간"] = max(2, len(mv))
        elif not chk:
            fields["입주"] = list(yms[0])
            fields["지정기간"] = n
            info.append("입주 일정 표시 없음 → 첫 달 입주·지정기간 전체로 가정")
        lod = rows.get("숙박비")
        fields["입주숙소"] = _rooms(rows, K) if lod and any(lod["vals"]) else 0
        fields["입주PC"] = _count(rows, "집기비품", r"PC\s*\(?\s*(\d+)\s*대")
        fields["복합기"] = _count(rows, "집기비품", r"복합기\s*\(?\s*(\d+)\s*대") or 1
        fields["입주안내"] = 2
    elif is_office:
        kinds = ["office"]
        fields["사무실시작"] = list(yms[0])
        if trailing_zero:
            fields["사무실종료"] = list(yms[last_nonzero])
        fields["통합센터"] = "N"
        mvl = [i for i, l in enumerate(labels) if "이전" in l]
        info.append("분양사무실 운영(첫 달부터)" + (f", {yms[mvl[0]][1]}월 이전" if mvl else ""))
    else:
        if open_idx or pre_idx or con_idx:
            kinds = ["mh"]
            if open_idx:
                D = open_idx[0]
                fields["오픈"] = list(yms[D])
                fields["사전영업"] = D          # 기간 안에서 오픈 전 달은 모두 사전영업(준비) 기간
            elif con_idx:
                y, m = yms[con_idx[0]]
                idx = y * 12 + m - 1 - 1
                fields["오픈"] = [idx // 12, idx % 12 + 1]
                fields["사전영업"] = 1
                info.append("오픈월 = 계약월 전월로 추정")
            else:
                y, m = yms[-1]
                idx = y * 12 + m - 1 + 1
                fields["오픈"] = [idx // 12, idx % 12 + 1]
                fields["사전영업"] = len(pre_idx) + 0
                info.append("기간 전체 사전영업 → 오픈월 = 기간 다음 달로 추정")
        else:
            kinds = ["run"]
            info.append("단계 표시 없음 → 견본주택 운영 구분")
        if trailing_zero:
            fields["MH종료"] = list(yms[last_nonzero])
            info.append(f"{yms[last_nonzero][1]}월 뒤 빈 달 → M/H 종료월")
        fields["숙소호실"] = _rooms(rows, K)
        fields["PC"] = _count(rows, "집기비품", r"PC\s*\(?\s*(\d+)\s*대") or _count(rows, "집기비품", r"100000\*(\d+)", formulas=True) or 0
        fields["복합기"] = _count(rows, "집기비품", r"복합기\s*\(?\s*(\d+)\s*대") or (1 if rows.get("집기비품") and any(rows["집기비품"]["vals"]) else 0)
        pr = _count(rows, "집기비품", r"(?:프린터|복사기)\s*\(?\s*(\d+)\s*대")
        fields["프린터"] = pr
        # 예비입주자 내용증명: 파일에 있으면 건수/세대수, 없으면 0
        tel = rows.get("통신비")
        fields["예비비율"] = 0
        if tel and "내용증명" in tel["note"]:
            cnt = 0
            fs = " ".join(f or "" for f in tel["formulas"])
            for m in re.finditer(r"(?:3500|4000|4500|5000)\*(\d{2,5})\*(\d{2,4})%", fs):      # 5000*201*500% 꼴
                cnt = max(cnt, int(m.group(1)) * int(m.group(2)) / 100)
            if not cnt:
                for m in re.finditer(r"(?:3500|4000|4500|5000)\*(\d{2,5})|(\d{2,5})\*(?:3500|4000|4500|5000)", fs):
                    cnt = max(cnt, int(m.group(1) or m.group(2)))
            m = re.search(r"내용증명[^\n]*?([\d,]{2,6})\s*건", tel["note"])
            if m:
                cnt = max(cnt, int(m.group(1).replace(",", "")))
            if not cnt and "%" in tel["note"]:
                m = re.search(r"내용증명[^\n%]*?(\d{2,4})\s*%", tel["note"])
                if m:
                    cnt = -int(m.group(1))        # 비율만 아는 경우(음수로 표시 → 아래서 비율로)
            fields["예비비율"] = None if not cnt else cnt     # 세대수 정한 뒤 비율로 바꿈
            info.append(f"예비입주자 내용증명 {cnt or '?'}건")
        fields["파쇄기"] = "Y" if re.search(r"파쇄|세단", (rows.get("집기비품") or {}).get("note", "")) else "N"
        bo = rows.get("보안경비용역")
        fields["보안공사"] = "Y" if bo and re.search(r"공사|설치", bo["note"]) else "N"
        fields["가구"] = "Y" if rows.get("사무용가구") and any(rows["사무용가구"]["vals"]) else "N"
        fields["보험료"] = "Y" if rows.get("보험료") and any(rows["보험료"]["vals"]) else "N"
        fields["송금수수료"] = "Y" if rows.get("송금수수료") and any(rows["송금수수료"]["vals"]) else "N"
        fields["전기료"] = "Y" if rows.get("전기료") and any(rows["전기료"]["vals"]) else "N"
        fields["난방비"] = "Y" if rows.get("난방비") and any(rows["난방비"]["vals"]) else "N"
        yk = rows.get("용역-기타")
        fields["오수관리"] = "Y" if yk and re.search(r"오수|정화조", yk["note"]) else "N"
        ds = rows.get("도서인쇄비")
        fields["도서인쇄"] = round(sum(ds["vals"]) / 1000) if ds else 0
        jg = rows.get("숙소중개")
        if jg and any(jg["vals"]) and fields["숙소호실"]:
            fields["숙소중개"] = round(sum(jg["vals"]) / fields["숙소호실"] / 1000)
        # 식대 단가(파일 산식·비고에서) — 없으면 그 해 기준
        po, pe = MEAL_BY_YEAR[ky]
        prices = set()
        if meal:
            for f in meal["formulas"]:
                for m in re.finditer(r"(\d{4,5})\*(\d{1,2})\*(\d{1,2})", f or ""):
                    prices.add(int(m.group(1)) // 1000)
            for m in re.finditer(r"(\d{1,2})\s*천\s*원\s*/\s*일", meal["note"]):
                prices.add(int(m.group(1)))
        prices = sorted(p for p in prices if 5 <= p <= 30)
        if len(prices) >= 2:
            po, pe = prices[-1], prices[0]
        elif len(prices) == 1:
            if prices[0] >= 14:
                po, pe = prices[0], prices[0] // 2
            else:
                po, pe = prices[0] * 2, prices[0]
        fields["식대단가오픈"], fields["식대단가기타"] = po, pe
    rv = rows.get("운반비")
    fields["운반비"] = "Y" if rv and any(rv["vals"]) else "N"
    # 세대수: 통신·발급 산식(4500*N 등) → 현장 표
    units = 0
    for key in ("통신비", "발급수수료"):
        r = rows.get(key)
        if not r:
            continue
        for f in r["formulas"]:
            for m in re.finditer(r"(?:3500|4000|4500|5000|600)\*(\d{3,5})|(\d{3,5})\*(?:3500|4000|4500|5000|600)", f or ""):
                v = int(m.group(1) or m.group(2))
                if 100 <= v <= 5000:
                    units = max(units, v)
    for m in re.finditer(r"([\d,]{3,6})\s*세대", notes_all):
        v = int(m.group(1).replace(",", ""))
        if 100 <= v <= 5000:
            units = max(units, v)
    if not units:
        units = SITE_UNITS.get(site, 0)
        if units:
            info.append("세대수는 현장 표 값")
        else:
            info.append("세대수 미상(0)")
    if "예비비율" in fields:
        cnt = fields["예비비율"]
        if cnt is not None and cnt < 0:
            fields["예비비율"] = -cnt / 100
        else:
            fields["예비비율"] = (cnt / units) if (cnt and units) else (0.05 if cnt is None else 0)
    fields.update({"현장명": re.sub(r"^\[([^\]]+)\].*", r"\1", title).strip() or site, "작성일": written.isoformat() if written else None,
                   "시작": list(yms[0]), "개월수": n, "지역": region, "세대수": units})
    appr = {"list": [dict(pos=p) for p in rec.get("approvers") or []]} if rec.get("approvers") else {"list": None}
    data = dict(kinds=kinds, fields=fields, months=months, overrides={}, notes={}, extra=[], approvers=appr, title=None)
    return data, K, ky, rows, yms, info


def _count(rows, key, pat, formulas=False):
    r = rows.get(key)
    if not r:
        return 0
    src = " ".join(f or "" for f in r["formulas"]) if formulas else r["note"]
    vals = [int(m.group(1)) for m in re.finditer(pat, src)]
    return max(vals) if vals else 0


def _rooms(rows, K):
    if not any(any(rows[k]["vals"]) for k in ("숙박비", "숙소관리비") if k in rows):
        return 0
    for key in ("숙박비", "숙소관리비"):
        r = rows.get(key)
        if not r:
            continue
        m = re.search(r"[×x\*]\s*(\d+)\s*호", r["note"]) or re.search(r"(\d+)\s*호실", r["note"]) or re.search(r"(\d+)\s*개\s*호실", r["note"])
        if m:
            return int(m.group(1))
    lod = rows.get("숙박비")
    if lod and any(lod["vals"]):
        for f in lod["formulas"]:
            m = re.match(r"=(\d+)\*(\d)\*(\d+)$", f or "")
            if m:
                return int(m.group(3))
        mx = max(lod["vals"])
        price = float(K.get("숙소비_월") or 1000) * 1000
        return max(1, round(mx / (price * float(K.get("숙소_개월") or 3))))
    mg = rows.get("숙소관리비")
    if mg and any(mg["vals"]):
        mx = max(mg["vals"])
        return max(1, round(mx / (float(K.get("숙소관리비") or 200) * 1000)))
    return 0


# ── 비교
def tool_rows(res):
    out = {}
    for g in res["groups"]:
        for r in g["rows"]:
            out[r["key"]] = r
    return out


def near(a, b):
    return abs(a - b) <= max(10000, 0.005 * max(abs(a), abs(b)))


def compare(case_no, site, period, rec, data, res, rows, yms):
    n = len(yms)
    tr = tool_rows(res)
    cells, keys = [], sorted(set(rows) | set(tr), key=lambda k: (k.startswith("extra:"), k))
    stat = dict(match=0, diff=0, miss=0, only=0, extra_rows=0, cmp=0, struct=0, struct_n=0, near20=0, extra_names=[])
    row_sum = {}
    for key in keys:
        a_row, t_row = rows.get(key), tr.get(key)
        a_vals = a_row["vals"] if a_row else [0.0] * n
        t_vals = [c["v"] for c in t_row["cells"]] if t_row else [0.0] * n
        if key.startswith("extra:") and any(a_vals):
            stat["extra_rows"] += 1
            stat["extra_names"].append(key[6:] + f"({sum(a_vals):,.0f})")
        for i in range(n):
            a, t = a_vals[i], t_vals[i]
            if not a and not t:
                continue
            stat["cmp"] += 1
            stat["struct_n"] += 1
            if (a > 0) == (t > 0):
                stat["struct"] += 1
            if a and t and abs(a - t) <= 0.2 * max(a, t):
                stat["near20"] += 1
            if near(a, t):
                st = "일치"
                stat["match"] += 1
            elif a and not t:
                st = "도구 없음"
                stat["miss"] += 1
            elif t and not a:
                st = "도구만"
                stat["only"] += 1
            else:
                st = "금액 차이"
                stat["diff"] += 1
            tf = ("=" + t_row["cells"][i]["f"]) if t_row and t_row["cells"][i].get("f") else ""
            af = (a_row["formulas"][i] or "") if a_row else ""
            name = "/".join(sorted(set(f"{it}·{sb}" if sb else it for it, sb in a_row["names"]))) if a_row else (t_row["item"] + "·" + t_row["name"].replace(" ", ""))
            cells.append(dict(no=case_no, site=site, period=period, key=key, name=name, month=f"{yms[i][0]}.{yms[i][1]:02d}", a=a, af=af, t=t, tf=tf,
                              d=t - a, st=st, note=(a_row["note"][:80].replace("\n", " / ") if a_row else ""), tnote=(t_row["note"][:80].replace("\n", " / ") if t_row else "")))
            row_sum.setdefault(key, [name, 0.0, 0.0])
            row_sum[key][1] += a
            row_sum[key][2] += t
    return cells, stat, row_sum


def main():
    os.makedirs(OUT, exist_ok=True)
    summary, details, inputs, logs = [], [], [], []
    import jc_pdf
    for no, (site, period, path, region, memo) in enumerate(CASES + CASES_PDF, 1):
        if path.lower().endswith(".pdf"):
            try:
                rec = jc_pdf.parse_pdf(path)
            except Exception as e:   # noqa: BLE001
                rec = dict(err=f"{type(e).__name__}: {e}")
            rec["path"], rec["name"] = path, os.path.basename(path)
        else:
            rec = inv.get(path)
        if not rec or "rows" not in rec:
            logs.append((no, site, period, "원본을 읽지 못함: " + (rec.get("err") if rec else "목록에 없음")))
            print(f"[{no}] {site} {period}: 원본 없음/오류 {rec.get('err') if rec else ''}")
            continue
        try:
            data, K, ky, rows, yms, info = infer(site, region, rec)
            res = jichul.calc(data, K)
            cells, stat, row_sum = compare(no, site, period, rec, data, res, rows, yms)
        except Exception as e:   # noqa: BLE001
            import traceback
            logs.append((no, site, period, f"오류 {type(e).__name__}: {e}"))
            print(f"[{no}] {site} {period}: 오류", e)
            traceback.print_exc()
            continue
        # 사례별 xlsx
        folder = os.path.join(OUT, re.sub(r'[\\/:*?"<>|]', "_", site))
        os.makedirs(folder, exist_ok=True)
        fname = f"({site}) {res['kind'].replace('M/H ', '')} 지출계획서{jichul.period_text(res)}.xlsx"
        fpath = os.path.join(folder, fname)
        with open(fpath, "wb") as f:
            f.write(jichul_xlsx.build_xlsx(res, data, ky))
        a_total = sum(rec["total"])
        t_total = res["total_all"]
        top = sorted(((abs(v[2] - v[1]), k, v) for k, v in row_sum.items() if not near(v[1], v[2])), reverse=True)[:3]
        top_txt = "; ".join(f"{v[0]}: 원본 {v[1]:,.0f} → 도구 {v[2]:,.0f}" for _, k, v in top)
        f = data["fields"]
        cond = []
        if "mh" in data["kinds"]:
            cond.append(f"오픈 {f['오픈'][0]}.{f['오픈'][1]:02d} 사전 {f['사전영업']}개월")
        cond.append(f"세대 {f['세대수']:,}")
        if "숙소호실" in f or "입주숙소" in f:
            cond.append(f"숙소 {f.get('숙소호실', f.get('입주숙소'))}호")
        if "PC" in f or "입주PC" in f:
            cond.append(f"PC {f.get('PC', f.get('입주PC'))}·복합기 {f.get('복합기')}·프린터 {f.get('프린터', '-')}")
        if "식대단가오픈" in f:
            cond.append(f"식대 {f['식대단가오픈']}/{f['식대단가기타']}천원")
        evs = [f"{yms[i][1]}월 {','.join(m['events'])}" for i, m in enumerate(data["months"]) if m["events"]]
        if evs:
            cond.append("행사 " + " / ".join(evs))
        kind_txt = {"mh": "견본주택 오픈", "run": "견본주택 운영", "movein": "입주", "office": "분양사무실"}[data["kinds"][0]]
        rate = stat["match"] / stat["cmp"] if stat["cmp"] else 0
        ratio = t_total / a_total if a_total else 0
        if rec.get("ocr") and not rec.get("grand_ok"):
            verdict = "참고(OCR 검산 불일치)"
        elif 0.85 <= ratio <= 1.15:
            verdict = "비슷(±15%)"
        elif 0.7 <= ratio <= 1.3:
            verdict = "차이 있음(±30%)"
        elif ratio > 1.3:
            verdict = "도구가 큼"
        else:
            verdict = "도구가 작음"
        # 달 머리글 비교(원본 표시 vs 도구 자동 표시)
        lab_ok = sum(1 for m_a, m_t in zip(rec["months"], res["months"]) if norm(m_a["label"]).strip("()") == norm(m_t["auto_label"]).strip("()"))
        ocr_txt = ""
        if rec.get("ocr"):
            ocr_txt = f"OCR: 합계줄 {'일치' if rec.get('grand_ok') else '불일치'}, 검산 실패 {len(rec.get('bad_rows') or [])}줄" + (f"({', '.join(rec['bad_rows'])})" if rec.get("bad_rows") else "")
        summary.append(dict(no=no, site=site, kind=kind_txt, period=period, written=f.get("작성일") or "", ky=ky, a_total=a_total, t_total=t_total,
                            origin=("PDF(OCR)" if rec.get("ocr") else "엑셀"), ocr=ocr_txt, verdict=verdict,
                            ratio=(t_total / a_total if a_total else None), rows_a=len(rec["rows"]), rows_t=sum(len(g["rows"]) for g in res["groups"]),
                            cmp=stat["cmp"], match=stat["match"], rate=rate, near20=(stat["match"] + stat["near20"]) / stat["cmp"] if stat["cmp"] else 0,
                            diff=stat["diff"], miss=stat["miss"], only=stat["only"], extra_rows=stat["extra_rows"], extra_names=", ".join(stat["extra_names"]),
                            lab=f"{lab_ok}/{len(yms)}",
                            struct=(stat["struct"] / stat["struct_n"] if stat["struct_n"] else 0), top=top_txt, cond=" · ".join(cond), info=" · ".join(info),
                            labels_a=" ".join(f"{m['m']}월{norm(m['label'])}" for m in rec["months"]), labels_t=" ".join(f"{m['m']}월{m['label']}" for m in res["months"]),
                            title_a=rec["title"], title_t=res["title"], memo=memo, src=path, out=fpath, warnings="; ".join(res["warnings"])))
        details.extend(cells)
        inputs.append(dict(no=no, site=site, period=period, kinds=",".join(data["kinds"]), ky=ky, fields=json.dumps(f, ensure_ascii=False),
                           months=json.dumps(data["months"], ensure_ascii=False), approvers=",".join(rec.get("approvers") or [])))
        print(f"[{no:2d}] {site:14s} {period:14s} {kind_txt:8s} K{ky} 원본 {a_total:>12,.0f} 도구 {t_total:>12,.0f} ({(t_total / a_total * 100) if a_total else 0:5.1f}%) "
              f"칸 {stat['cmp']:3d} 일치 {rate * 100:5.1f}% 차이 {stat['diff']:2d} 누락 {stat['miss']:2d} 추가 {stat['only']:2d} 원본전용줄 {stat['extra_rows']}")
    write_report(summary, details, inputs, logs)
    print("\n요약:", REPORT)


FINDINGS = [
    "■ 검증 방법: 과거 지출계획서(엑셀 52건 + 스캔 PDF 23건, 2019~2026년, 18개 현장)마다 원본의 조건(기간·오픈월·세대수·숙소·PC·행사·식대 단가 등)을 읽어 도구(지출계획서 자동작성)에 넣고, "
    "그 해 기준안(2022~2027 기준단가, 2021년 이전은 2022 기준안)으로 계산한 결과를 원본과 줄·달 단위로 비교했다. 도구가 만든 엑셀은 tests\\compare_jichul\\out\\<현장>\\ 에 있다.",
    "■ PDF(OCR) 사례는 Windows 글자 인식으로 읽어 숫자·줄 이름에 오류가 남아 있을 수 있다(열 'OCR 검산' 참고). 결론은 엑셀 52건을 기준으로 한다.",
    "■ 결론 1 — 2026년 현장(평택고덕 1·2차, 오산서2 2026)은 합계가 원본의 82~109%, 칸 일치 43~67%로 '비슷한 값'이 나온다. 도구의 규칙이 2027 기준안·2026 실무(평택고덕 결재본)를 따라 만들어졌기 때문이다.",
    "■ 결론 2 — 2025년 이전 현장은 합계 비율이 0.7~4.5배로 벌어진다. 원인은 도구 버그가 아니라 (1) 그 해 기준안과 다른 실무 금액(통신비 300 vs 기준 800, 식대 인원 2~4명 vs 기준 4~6명, 발급수수료 50 vs 30, 보안 300 vs 400), "
    "(2) 운영 후반·잔여 소수 현장은 기준안 기본값(소모품 1,000·청소 1명 3,500·상하수도 500·통신 800)보다 훨씬 작게 편성, (3) 현장 특수 항목(주차장 임차료·코로나 방역·CCTV·홍보관 관리비·살수)이다.",
    "■ 결론 3 — 오픈 분기의 구조(사전/오픈/계약 머리글, 숙소 일시불, 사무용가구·무인경비 공사·운반비 1회, 소모품 4,000/3,000, 식대·여비 단계)는 대부분 재현된다. 다만 '오픈 전월 1회 항목'을 사전영업 시작월에 넣는 도구와 달리 2020~2021년 파일은 오픈월에 넣었고, "
    "2023·2025년 일부 현장은 사전영업월에 월 운영비(소모품·상하수도·용역·집기·통신)를 넣지 않았다(2026 평택고덕은 도구와 같음).",
    "■ 결론 4 — 금액이 크게 벌어지는 규칙 3가지: (a) 예비입주자 내용증명 = 세대수×500%(기준안) — 실무는 18~55%(오산서2 250건, 오산세교 850건)이고 2026 평택고덕만 200~500%. 세대수가 크면 수천만 원 차이. "
    "(b) 행사(중도금대출·옵션) 안내문 등기 4.5천원×세대를 행사 달에 넣는데 김포북변은 행사 전월에 발송했고 오산서2는 옵션변경 달에 넣지 않았다. (c) 운영 구분의 청소 1명 기본 — 오산서2·평택고덕 2026 운영기는 2명.",
    "■ 결론 5 — 입주 구분: 입주 숙소 3개월 일시불(기준안) vs 실무 월납·일부(화성비봉 650천원×2), 입주지원센터 비용을 지정기간+1개월까지 넣는 규칙 vs 실무는 입주 2개월째부터 줄임. 사례 수가 적어(엑셀 6건) 추가 확인 필요.",
    "■ 개선 후보(코드 미반영, 판단 요청): ① 운영 구분에 '축소 운영' 또는 항목별 반영 체크(소모품·청소·상하수도·통신·보안) ② 예비입주자 내용증명 비율 기본값 낮추기(또는 현장 입력 필수) ③ 행사 안내문 등기 '전월 발송' 선택 "
    "④ 발급수수료 기본 50천원(2024~2026 실무) ⑤ 사전영업월 월 운영비 반영 여부 선택 ⑥ 입주 숙소 월납 선택 ⑦ 2021년 이전 파일은 비교 참고용(기준안 없음).",
]


def write_report(summary, details, inputs, logs):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "요약"
    bold = Font(bold=True)
    ws.cell(1, 1, f"지출계획서 자동작성 — 과거 현장·기간별 재현 검증 ({datetime.date.today():%Y-%m-%d})").font = Font(bold=True, size=14)
    # 구분·연도별 집계
    r = 3
    ws.cell(r, 1, "구분별").font = bold
    ws.cell(r, 2, "사례").font = bold
    ws.cell(r, 3, "원본 합계").font = bold
    ws.cell(r, 4, "도구 합계").font = bold
    ws.cell(r, 5, "도구/원본").font = bold
    ws.cell(r, 6, "칸 일치율").font = bold
    ws.cell(r, 7, "비슷(±15%) 사례").font = bold
    xl = [s for s in summary if s["origin"] == "엑셀"]

    def agg(rowsel, label):
        nonlocal r
        r += 1
        a, t = sum(s["a_total"] for s in rowsel), sum(s["t_total"] for s in rowsel)
        c, m = sum(s["cmp"] for s in rowsel), sum(s["match"] for s in rowsel)
        ws.cell(r, 1, label)
        ws.cell(r, 2, len(rowsel))
        ws.cell(r, 3, a).number_format = "#,##0"
        ws.cell(r, 4, t).number_format = "#,##0"
        ws.cell(r, 5, t / a if a else None).number_format = "0.0%"
        ws.cell(r, 6, m / c if c else None).number_format = "0.0%"
        ws.cell(r, 7, sum(1 for s in rowsel if s["verdict"].startswith("비슷")))

    for k in ("견본주택 오픈", "견본주택 운영", "입주", "분양사무실"):
        sel = [s for s in xl if s["kind"] == k]
        if sel:
            agg(sel, f"엑셀 · {k}")
    for y in sorted({s["ky"] for s in xl}):
        agg([s for s in xl if s["ky"] == y], f"엑셀 · {y} 기준안(= {y}년 계획)")
    agg(xl, "엑셀 전체")
    pdfs = [s for s in summary if s["origin"] != "엑셀"]
    if pdfs:
        agg(pdfs, "PDF(OCR) 전체 — 참고")
    r += 2
    ws.cell(r, 1, "해석과 개선 후보").font = bold
    for line in FINDINGS:
        r += 1
        ws.cell(r, 1, line).alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=20)
        ws.row_dimensions[r].height = 46
    r += 2
    hdr_row = r
    head = ["번호", "현장", "도구 구분", "기간", "원본 작성일", "기준안(연도)", "원본 합계(원)", "도구 합계(원)", "도구/원본", "합계 판정", "원본 줄 수", "도구 줄 수",
            "비교 칸", "일치 칸", "일치율", "±20% 이내", "금액 차이", "도구 없음", "도구만 있음", "원본 전용 줄", "원본 전용 줄 이름(합계)", "구조 일치율", "달 머리글 일치",
            "큰 차이 3가지(줄: 원본 → 도구)", "추정한 조건", "추정 메모", "원본 월 표시", "도구 월 표시", "원본 제목", "도구 제목", "메모", "도구 경고", "원본 종류", "OCR 검산", "원본 파일", "도구가 만든 파일"]
    for i, h in enumerate(head, 1):
        c = ws.cell(hdr_row, i, h)
        c.font = bold
        c.fill = PatternFill("solid", fgColor="DDEBF7")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for s in summary:
        ws.append([s["no"], s["site"], s["kind"], s["period"], s["written"], s["ky"], s["a_total"], s["t_total"], s["ratio"], s["verdict"], s["rows_a"], s["rows_t"],
                   s["cmp"], s["match"], s["rate"], s["near20"], s["diff"], s["miss"], s["only"], s["extra_rows"], s["extra_names"], s["struct"], s["lab"],
                   s["top"], s["cond"], s["info"], s["labels_a"], s["labels_t"], s["title_a"], s["title_t"], s["memo"], s["warnings"], s["origin"], s["ocr"], s["src"], s["out"]])
    first, last = hdr_row + 1, hdr_row + len(summary)
    if summary:
        rr = last + 2
        ws.cell(rr, 2, "전체").font = bold
        ws.cell(rr, 7, f"=SUM(G{first}:G{last})")
        ws.cell(rr, 8, f"=SUM(H{first}:H{last})")
        ws.cell(rr, 9, f"=H{rr}/G{rr}")
        ws.cell(rr, 13, f"=SUM(M{first}:M{last})")
        ws.cell(rr, 14, f"=SUM(N{first}:N{last})")
        ws.cell(rr, 15, f"=N{rr}/M{rr}")
        for c in (17, 18, 19, 20):
            ws.cell(rr, c, f"=SUM({L(c)}{first}:{L(c)}{last})")
    widths = [5, 16, 11, 15, 11, 8, 14, 14, 8, 15, 7, 7, 7, 7, 7, 8, 7, 7, 8, 8, 30, 8, 8, 70, 60, 40, 36, 36, 30, 30, 22, 30, 10, 30, 60, 60]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    for row in ws.iter_rows(min_row=first, max_row=ws.max_row):
        for c in row:
            if c.column in (7, 8):
                c.number_format = "#,##0"
            if c.column in (9, 15, 16, 22):
                c.number_format = "0.0%"
            if c.column in (21, 24, 25, 26):
                c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = f"E{first}"
    ws.auto_filter.ref = f"A{hdr_row}:{L(len(head))}{last}"

    ws2 = wb.create_sheet("칸별 비교")
    head2 = ["번호", "현장", "기간", "줄키", "원본 줄(항목·세목)", "월", "원본 금액", "원본 산식", "도구 금액", "도구 산식", "차이(도구-원본)", "판정", "원본 비고", "도구 비고"]
    ws2.append(head2)
    for d in details:
        ws2.append([d["no"], d["site"], d["period"], d["key"], d["name"], d["month"], d["a"], d["af"], d["t"], d["tf"], d["d"], d["st"], d["note"], d["tnote"]])
    for i, w in enumerate([5, 16, 14, 12, 22, 8, 13, 28, 13, 28, 13, 9, 50, 50], 1):
        ws2.column_dimensions[L(i)].width = w
    for c in ws2[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row):
        for c in row:
            if c.column in (7, 9, 11):
                c.number_format = "#,##0"
    ws2.freeze_panes = "G2"
    ws2.auto_filter.ref = f"A1:{L(len(head2))}{len(details) + 1}"

    ws3 = wb.create_sheet("추정한 입력조건")
    ws3.append(["번호", "현장", "기간", "현장 구분", "기준안 연도", "입력 칸(JSON)", "월별 행사(JSON)", "원본 결재란"])
    for x in inputs:
        ws3.append([x["no"], x["site"], x["period"], x["kinds"], x["ky"], x["fields"], x["months"], x["approvers"]])
    for i, w in enumerate([5, 16, 14, 10, 8, 120, 60, 30], 1):
        ws3.column_dimensions[L(i)].width = w
    for c in ws3[1]:
        c.font = Font(bold=True)

    ws4 = wb.create_sheet("연도별 기준단가")
    keys = sorted({k for K in K_YEARS.values() for k in K})
    years = sorted(K_YEARS)
    ws4.append(["항목"] + [f"{y}년 기준안" for y in years])
    for k in keys:
        ws4.append([k] + [K_YEARS[y].get(k) for y in years])
    ws4.column_dimensions["A"].width = 24
    for c in ws4[1]:
        c.font = Font(bold=True)

    ws5 = wb.create_sheet("읽지 못한 원본")
    ws5.append(["번호", "현장", "기간", "사유"])
    for lg in logs:
        ws5.append(list(lg))
    wb.save(REPORT)


if __name__ == "__main__":
    main()
