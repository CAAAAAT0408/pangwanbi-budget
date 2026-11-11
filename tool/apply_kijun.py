# -*- coding: utf-8 -*-
"""새 기준안을 자동작성 양식(또는 이미 작성한 현장 파일)에 반영한다.

사용
  python apply_kijun.py                       → 기준안 파일을 고르는 창, 폴더의 최신 양식에 반영해 새 연도 양식 저장
  python apply_kijun.py <기준안.xlsx>          → 최신 양식에 반영
  python apply_kijun.py <기준안.xlsx> <대상.xlsx> [--inplace]
        대상이 현장 파일이면 --inplace 로 그 파일의 기준단가만 갱신(원본은 백업 폴더에 복사)

결과 파일 이름: '판관비 예산안 자동작성 양식_<예산연도>.xlsx' (같은 연도 파일이 있으면 백업 후 덮어씀)
화면(app\\server.py)도 apply() 를 부른다.
"""
import datetime
import glob
import os
import re
import shutil
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import kijun_parser   # noqa: E402

TEMPLATE_GLOB = "판관비 예산안 자동작성 양식_*.xlsx"
BUDGET_DIR_NAME = "★ 판관비 예산안"


def find_budget_root(start=ROOT):
    """예산안 자료 폴더(1. 본사업무\\★ 판관비 예산안)를 위쪽 폴더들에서 찾는다.
    저장소를 AI자동화\\1. 자동화 시스템 같은 곳으로 옮겨도, 외장하드가 E:/F: 어느 쪽이어도 같은 폴더를 쓴다.
    못 찾으면 예전 규칙대로 저장소 바로 위 폴더(app\\server.py 의 _find_budget_root 와 같은 규칙)."""
    p = os.path.abspath(start)
    while True:
        parent = os.path.dirname(p)
        if not parent or parent == p:
            return os.path.dirname(os.path.abspath(start))
        if os.path.basename(parent) == BUDGET_DIR_NAME:
            return parent
        cand = os.path.join(parent, BUDGET_DIR_NAME)
        if os.path.isdir(cand):
            return cand
        p = parent


def templates():
    """[(연도, 경로)] 새 연도 순"""
    out = []
    for p in glob.glob(os.path.join(ROOT, TEMPLATE_GLOB)):
        m = re.search(r"_(\d{4})\.xlsx$", p)
        if m:
            out.append((int(m.group(1)), p))
    return sorted(out, reverse=True)


def latest_template():
    t = templates()
    return t[0][1] if t else None


def pick_file(title, initial):
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(title=title, initialdir=initial,
                                      filetypes=[("엑셀 파일", "*.xlsx"), ("모든 파일", "*.*")])
    root.destroy()
    return path or None


def budget_sheet(wb):
    for name in wb.sheetnames:
        if re.fullmatch(r"\d{2}년 예산", name):
            return name
    return None


def rename_year(wb, new_year):
    """'27년 예산' → 'YY년 예산' 시트 이름과 모든 수식 참조를 바꾼다"""
    old = budget_sheet(wb)
    new = f"{new_year % 100:02d}년 예산"
    if not old or old == new:
        return None
    wb[old].title = new
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and f"'{old}'!" in c.value:
                    c.value = c.value.replace(f"'{old}'!", f"'{new}'!")
    return old, new


def backup(path):
    d = os.path.join(ROOT, "백업")
    os.makedirs(d, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    dst = os.path.join(d, f"{stem}_{datetime.datetime.now():%Y%m%d_%H%M%S}{ext}")
    shutil.copy2(path, dst)
    return dst


class ApplyError(Exception):
    pass


def apply(kijun, target=None, inplace=False):
    """기준안 반영. 결과 dict(out, year, renamed, report, unused, written, backup)"""
    target = target or latest_template()
    if not target:
        raise ApplyError("양식 파일(" + TEMPLATE_GLOB + ")을 찾지 못했습니다.")
    lock = os.path.join(os.path.dirname(target), "~$" + os.path.basename(target))
    if inplace and os.path.exists(lock):
        raise ApplyError("대상 파일이 Excel 에 열려 있습니다. 닫고 다시 하세요.")
    try:
        wb = openpyxl.load_workbook(target)
    except Exception as e:
        raise ApplyError(f"양식을 열 수 없습니다: {e}")
    if "기준단가" not in wb.sheetnames:
        raise ApplyError("대상 파일에 '기준단가' 시트가 없습니다(자동작성 양식이 아님).")
    try:
        info = kijun_parser.apply_to_workbook(wb, kijun)
    except Exception as e:
        raise ApplyError(f"기준안을 읽지 못했습니다: {e}")
    year = info["written"].year + 1 if info["written"] else None
    renamed = rename_year(wb, year) if year else None
    if "기준안 반영결과" in wb.sheetnames:
        wb.move_sheet("기준안 반영결과", offset=wb.sheetnames.index("기준단가") + 1 - wb.sheetnames.index("기준안 반영결과"))
    wb.calculation.fullCalcOnLoad = True

    bak = None
    if inplace:
        bak = backup(target)
        out = target
    else:
        if not year:
            raise ApplyError("기준안 작성일(예: 작성일 : 2026년 10월 15일)을 찾지 못해 예산연도를 정할 수 없습니다.")
        out = os.path.join(ROOT, f"판관비 예산안 자동작성 양식_{year}.xlsx")
        if os.path.exists(out):
            bak = backup(out)
    wb.save(out)
    return dict(out=out, year=year, renamed=renamed, report=info["report"], unused=info["unused"],
                written=info["written"], backup=bak, target=target)


def main(argv):
    inplace = "--inplace" in argv
    args = [a for a in argv if not a.startswith("--")]
    kijun = args[0] if args else pick_file("새 PM팀 판관비 예산기준(안) 엑셀을 고르세요", find_budget_root())
    if not kijun:
        print("기준안 파일을 고르지 않아 끝냅니다.")
        return 1
    target = args[1] if len(args) > 1 else None
    print("기준안 :", kijun)
    print("대상   :", target or latest_template())
    try:
        r = apply(kijun, target, inplace)
    except ApplyError as e:
        print("✖", e)
        return 1
    rep = r["report"]
    cnt = {k: sum(1 for x in rep if x[0] == k) for k in ("변경", "동일", "확인", "세목")}
    print()
    print(f"기준안 작성일 {r['written']} → {r['year']}년 예산" if r["year"] else "기준안 작성일을 못 읽음(예산연도 그대로)")
    if r["renamed"]:
        print(f"시트 이름 변경: {r['renamed'][0]} → {r['renamed'][1]}")
    print(f"바뀐 값 {cnt['변경']}개, 같은 값 {cnt['동일']}개, 확인 필요 {cnt['확인']}개, 단가로 읽지 않은 기준안 줄 {len(r['unused'])}개")
    for kind, desc, old, new, why in rep:
        if kind in ("변경", "세목"):
            print(f"  [바뀜] {desc}: {kijun_parser.fmt(old) if not isinstance(old, str) else old} → "
                  f"{kijun_parser.fmt(new) if not isinstance(new, str) else new}")
    for kind, desc, old, new, why in rep:
        if kind == "확인":
            print(f"  [확인 필요] {desc}: {why} (기존 {kijun_parser.fmt(old)})")
    for ln in r["unused"]:
        print(f"  [읽지 않은 줄] {ln['row']}행 {' '.join(ln['item'].split())}: {' '.join(ln['text'].split())[:80]}")
    print()
    if r["backup"]:
        print("백업 :", r["backup"])
    print("저장 :", r["out"])
    print("→ 파일의 [기준안 반영결과] 시트에서 자세히 확인하세요.")
    return 0


if __name__ == "__main__":
    try:
        code = main(sys.argv[1:])
    except Exception:      # 창이 바로 닫히지 않게
        import traceback
        traceback.print_exc()
        code = 1
    sys.exit(code)
