"""A33 PDF 의 1차(전체 3x)·2차(표 부분 4.2x·5.5x) PNG 를 뽑아 둔다 (브라우저 OCR 실험용)"""
import sys, os
sys.path.insert(0, r"F:\1. 우미\1. 본사업무\★ 판관비 예산안\☆ 예산안 자동작성\tool")
sys.stdout.reconfigure(encoding="utf-8")
import design_overview as D

out = sys.argv[1]
pdf = r"F:\1. 우미\2. 수도권남부\14. 평택고덕 패키지형 공모사업\★ 2차(A33BL)\설계\03. 2차 변경승인 완료 (26.09.08)\03. 2차 변경승인 완료 (26.09.08)\02. 도면\1. 건축\01 표제부\(변경후) AG-014~015 설계개요, 동별개요.pdf"
fitz = D._fitz()
doc = fitz.open(pdf)
gen = D._pdf_ocr_pipeline(fitz, doc, 20, print)
imgs1 = next(gen)
for i, png in enumerate(imgs1):
    open(os.path.join(out, f"full_{i}.png"), "wb").write(png)
res1 = D.ocr_windows(imgs1)
try:
    imgs2 = gen.send(res1)
    for i, png in enumerate(imgs2):
        open(os.path.join(out, f"crop_{i}.png"), "wb").write(png)
    print("2차 그림", len(imgs2))
    res2 = D.ocr_windows(imgs2)
    try:
        gen.send(res2)
    except StopIteration as e:
        print("결과:", [(t["name"], t["units"]) for t in e.value[0]["types"]] if e.value else None)
except StopIteration as e:
    print("1차에서 끝:", [(t["name"], t["units"]) for t in e.value[0]["types"]] if e.value else None)
