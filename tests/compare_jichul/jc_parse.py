# -*- coding: utf-8 -*-
"""과거 지출계획서 엑셀 → 공통 구조(제목·작성일·월·줄별 월 금액/산식/비고). 모든 후보 파일을 훑어 JSON 으로 저장."""
import io
import json
import os
import re
import sys
import warnings

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl

SP = os.path.dirname(os.path.abspath(__file__))
BASE = r"E:\1. 우미"
SKIP = ("사은품", "경품", "판촉", "사전청약", "연말 자금", "전체일정", "~$")


def norm(s):
    return re.sub(r"\s+", "", str(s)) if s is not None else ""


def candidates():
    out = []
    for root, dirs, files in os.walk(BASE):
        for fn in files:
            if "지출계획" in fn and fn.lower().endswith((".xlsx", ".xls", ".pdf")) and not any(k in fn for k in SKIP):
                out.append(os.path.join(root, fn))
    return sorted(out)


def parse_xlsx(path):
    wbf = openpyxl.load_workbook(path, data_only=False)
    wbv = openpyxl.load_workbook(path, data_only=True)
    ws = None
    for name in wbf.sheetnames:
        n = name.replace(" ", "")
        if "지출계획" in n or "오픈" in n or "입주" in n or "운영" in n:
            ws = wbf[name]
            break
    ws = ws or wbf.worksheets[0]
    wv = wbv[ws.title]
    title = date = None
    hdr = None
    for row in ws.iter_rows(min_row=1, max_row=14):
        for c in row:
            v = c.value
            if isinstance(v, str):
                if "지출계획서" in v and title is None and c.column <= 3:
                    title = v.strip()
                if "작성일" in v and date is None:
                    date = v.strip()
                if norm(v) == "구분" and hdr is None:
                    hdr = c.row
    if hdr is None:
        return dict(err="머리글 없음")
    tcol = ncol = None
    for c in ws[hdr]:
        if isinstance(c.value, str):
            if norm(c.value) == "합계" and tcol is None:
                tcol = c.column
            if norm(c.value) == "비고" and ncol is None:
                ncol = c.column
    months, mcols = [], []
    for c in ws[hdr + 1]:
        v = c.value
        if isinstance(v, str) and re.search(r"\d{1,2}\s*월", v) and (tcol is None or c.column > tcol) and (ncol is None or c.column < ncol):
            if ws.column_dimensions[openpyxl.utils.get_column_letter(c.column)].hidden:
                continue
            m = re.search(r"(?:(\d{2,4})\s*년\s*)?'?(\d{1,2})\s*월", v)
            mm = int(m.group(2))
            if months and mm != months[-1]["m"] % 12 + 1:      # 달이 이어지지 않으면(옆에 붙은 다음 분기 메모 열) 끝
                break
            lab = re.sub(r"^.*?월", "", v, count=1, flags=re.S).strip()
            months.append(dict(y=int(m.group(1)) if m.group(1) else None, m=mm, label=lab))
            mcols.append(c.column)
    rows = []
    item = None
    for r in range(hdr + 2, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        an, bn = norm(a), norm(b)
        if an.startswith("합계"):
            break
        if an and not an.startswith("※"):
            item = str(a).strip()
        if bn.startswith("소계") or an.startswith("소계"):
            continue
        if not an and not bn:
            continue
        sub = str(b).strip() if bn else ""
        vals = []
        for c in mcols:
            v = wv.cell(r, c).value
            vals.append(float(v) if isinstance(v, (int, float)) else 0.0)
        fs = [ws.cell(r, c).value for c in mcols]
        fs = [str(f) if isinstance(f, str) and f.startswith("=") else None for f in fs]
        note = ws.cell(r, ncol).value if ncol else None
        if not any(vals) and not note:
            continue
        rows.append(dict(item=norm(item), sub=norm(sub), vals=vals, formulas=fs, note=(str(note).strip() if note else "")))
    # 결재란(직위)
    appr = []
    for row in ws.iter_rows(min_row=hdr + 2, max_row=ws.max_row):
        for c in row:
            if isinstance(c.value, str) and norm(c.value) == "결재":
                r0 = c.row
                for rr in range(r0 + 1, r0 + 5):
                    vals = [ws.cell(rr, cc).value for cc in range(c.column + 1, c.column + 9)]
                    vals = [norm(v) for v in vals if isinstance(v, str) and norm(v) and norm(v) != "/"]
                    if vals and all(re.fullmatch(r"[가-힣]{2,4}", v) for v in vals):
                        appr = vals
    return dict(sheet=ws.title, title=title, date=date, months=months, rows=rows, approvers=appr,
                total=[sum(r["vals"][i] for r in rows) for i in range(len(mcols))])


def main():
    out = []
    for p in candidates():
        rec = dict(path=p, name=os.path.basename(p), ext=os.path.splitext(p)[1].lower(), size=os.path.getsize(p))
        if rec["ext"] == ".xlsx":
            try:
                rec.update(parse_xlsx(p))
            except Exception as e:   # noqa: BLE001
                rec["err"] = f"{type(e).__name__}: {e}"
        out.append(rec)
    with io.open(os.path.join(SP, "jc_inventory.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    for rec in out:
        if rec["ext"] != ".xlsx":
            continue
        if "err" in rec:
            print(f"ERR  {rec['name']}  {rec['err']}")
            continue
        ms = " ".join(f"{m['m']}월{m['label'].replace(chr(10), '')}" for m in rec["months"])
        print(f"{rec['date'] or '':28s} | {rec['title'] or '':34s} | {ms:40s} | 줄 {len(rec['rows']):2d} | 합계 {sum(rec['total']):>13,.0f} | 결재 {rec['approvers']} | {rec['name']}")


if __name__ == "__main__":
    main()
