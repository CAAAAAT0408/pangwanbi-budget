# -*- coding: utf-8 -*-
"""웹판(docs/) 갱신: tool 의 엔진 파일과 계산용 양식·회사 양식을 docs/ 로 복사하고 VERSION 을 올린다.
   python tool\\sync_web.py            (엔진·양식을 고친 뒤 실행하고 커밋·push)"""
import datetime
import glob
import os
import re
import shutil

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOL = os.path.join(ROOT, "tool")
DOCS = os.path.join(ROOT, "docs")
PY = ["xlcalc.py", "params.py", "kijun_parser.py", "build_template.py", "site_io.py", "xlsxio.py",
      "report.py", "report_xl.py", "design_overview.py", "jichul.py", "jichul_xlsx.py", "siljeok.py"]
SILJEOK_FORM = os.path.join(ROOT, "회사양식", "실적계획", "사업관리비 실적계획 양식.xlsx")    # 사내 「사업관리비 실적계획」 양식 틀


def latest_template():
    ts = []
    for p in glob.glob(os.path.join(ROOT, "판관비 예산안 자동작성 양식_*.xlsx")):
        m = re.search(r"_(\d{4})\.xlsx$", p)
        if m:
            ts.append((int(m.group(1)), p))
    return max(ts)[1] if ts else None


def main():
    os.makedirs(os.path.join(DOCS, "py"), exist_ok=True)
    os.makedirs(os.path.join(DOCS, "data"), exist_ok=True)
    for f in PY:
        shutil.copyfile(os.path.join(TOOL, f), os.path.join(DOCS, "py", f))
    tpl = latest_template()
    year = int(re.search(r"_(\d{4})\.xlsx$", tpl).group(1))
    shutil.copyfile(tpl, os.path.join(DOCS, "data", "template.xlsx"))
    for old in glob.glob(os.path.join(DOCS, "data", "template_*.xlsx")):
        os.remove(old)
    forms = sorted(glob.glob(os.path.join(ROOT, "회사양식", "*.xlsx")))
    forms = [f for f in forms if not os.path.basename(f).startswith("~$")]
    shutil.copyfile(forms[-1], os.path.join(DOCS, "data", "form.xlsx"))
    if os.path.exists(SILJEOK_FORM):               # 실적계획 웹판이 쓰는 양식 틀(tool/build_siljeok_template.py 가 만든 것)
        shutil.copyfile(SILJEOK_FORM, os.path.join(DOCS, "data", "siljeok_form.xlsx"))
    os.makedirs(os.path.join(DOCS, "examples"), exist_ok=True)
    for ex in glob.glob(os.path.join(ROOT, "예시", "*.xlsx")):
        if not os.path.basename(ex).startswith("~$"):
            shutil.copyfile(ex, os.path.join(DOCS, "examples", os.path.basename(ex)))
    import json
    import sys
    sys.path.insert(0, TOOL)
    import site_io
    t = site_io.template_info(tpl)          # 지출계획서 웹판용 기준단가(양식의 '기준단가' 시트 → JSON)
    with open(os.path.join(DOCS, "data", "kijun.json"), "w", encoding="utf-8") as f:
        json.dump(dict(year=t["year"], kijun=t["kijun"], written=t["written"], template=os.path.basename(tpl),
                       synced=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), params=t["params"]),
                  f, ensure_ascii=False, indent=1, default=str)
    with open(os.path.join(DOCS, "data", "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(dict(template=os.path.basename(tpl), form=os.path.basename(forms[-1]), year=year,
                       synced=datetime.datetime.now().strftime("%Y-%m-%d %H:%M")), f, ensure_ascii=False, indent=1)
    ver = datetime.datetime.now().strftime("%Y.%m.%d.%H%M")
    for fn in ("worker.js", "index.html", os.path.join("jichul", "worker.js"), os.path.join("jichul", "index.html"),
               os.path.join("siljeok", "worker.js"), os.path.join("siljeok", "index.html")):
        p = os.path.join(DOCS, fn)
        s = open(p, encoding="utf-8").read()
        s2 = re.sub(r"const VERSION = '[^']*'", f"const VERSION = '{ver}'", s)
        if s2 != s:
            open(p, "w", encoding="utf-8").write(s2)
    print(f"docs 갱신: 엔진 {len(PY)}개, 양식 {os.path.basename(tpl)}, 회사양식 {os.path.basename(forms[-1])}, 기준단가 {len(t['params'])}줄, VERSION {ver}")


if __name__ == "__main__":
    main()
