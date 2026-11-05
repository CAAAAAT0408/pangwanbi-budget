# -*- coding: utf-8 -*-
"""연도별 기준안(2022~2027)으로 자동작성 양식을 만들어 '기준단가' 값을 K_years.json 으로 저장 (jc_compare.py 가 읽음).
   python tests\\compare_jichul\\build_k_years.py        — 양식 파일은 스크래치(임시) 폴더에 만들고 JSON 만 여기에 남긴다.

지출계획서 재현 검증 순서:  jc_parse.py(원본 엑셀 → jc_inventory.json) → build_k_years.py → jc_compare.py(→ out\\ 사례별 xlsx, tests\\지출계획서비교_테스트결과_<날짜>.xlsx)
스캔 PDF 는 jc_pdf.py(Windows OCR) 가 읽는다(ocr_cache\\ 에 단어 캐시)."""
import json
import os
import sys
import tempfile
import warnings

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
BUD = os.path.dirname(ROOT)
sys.path.insert(0, os.path.join(ROOT, "tool"))
import build_template   # noqa: E402
import site_io          # noqa: E402

KIJUN = {
    2027: os.path.join(BUD, r"★ 2027 예산안\PM팀 판관비 예산기준(안)_2026.10.15.xlsx"),
    2026: os.path.join(BUD, r"★ 2026 예산안\PM팀 판관비 예산기준(안)_2025.10.27.xlsx"),
    2025: os.path.join(BUD, r"★ 2025 예산안\PM팀 판관비 예산기준(안)_2024.11.18.xlsx"),
    2024: os.path.join(BUD, r"★ 2024 예산안\PM팀 판관비 예산기준(안)_2023.11.07.xlsx"),
    2023: os.path.join(BUD, r"마케팅팀 판관비 예산기준(안)_2022.10.31.xlsx"),
    2022: os.path.join(BUD, r"★ 2022 예산안\참고\마케팅팀 판관비 예산기준(안)_2021.10.15.xlsx"),
}
OUT = tempfile.mkdtemp(prefix="jc_k_")
table = {}
for y, kp in sorted(KIJUN.items()):
    out = os.path.join(OUT, f"양식_{y}.xlsx")
    build_template.build(out, kp)
    info = site_io.template_info(out)
    K = {p["key"]: p["value"] for p in info["params"] if "key" in p}
    chk = sum(1 for p in info["params"] if "key" in p and str(p.get("result") or "").startswith("확인"))
    table[y] = K
    print(f"{y}: {os.path.basename(kp)} 단가 {len(K)}개, 확인 필요 {chk}개")
with open(os.path.join(HERE, "K_years.json"), "w", encoding="utf-8") as f:
    json.dump(table, f, ensure_ascii=False, indent=1, default=str)
print("저장:", os.path.join(HERE, "K_years.json"))
