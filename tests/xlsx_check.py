"""xlsx 구조 검사(행·칸 순서, 중복, 병합 겹침) + fill_budget 안쪽 단계 bisect"""
import sys, os, json, glob, site, re, zipfile, traceback
site.addsitedir(site.getusersitepackages())
sys.path.insert(0, r"F:\1. 우미\1. 본사업무\★ 판관비 예산안\☆ 예산안 자동작성\tool")
sys.stdout.reconfigure(encoding="utf-8")
import warnings
warnings.filterwarnings("ignore")
from lxml import etree
import win32com.client
import xlsxio as X
import report_xl as R

NS = X.NS
SP = os.path.dirname(os.path.abspath(__file__))


def check(path):
    z = zipfile.ZipFile(path)
    probs = []
    for n in z.namelist():
        if not re.match(r"xl/worksheets/sheet\d+\.xml$", n):
            continue
        root = etree.fromstring(z.read(n))
        prev_r = 0
        for row in root.iter(X.q("row")):
            r = int(row.get("r"))
            if r <= prev_r:
                probs.append(f"{n}: 행 순서 {prev_r}→{r}")
            prev_r = r
            prev_c = 0
            for c in row.findall(X.q("c")):
                ref = c.get("r")
                m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
                if not m or int(m.group(2)) != r:
                    probs.append(f"{n}: 칸 {ref} 가 행 {r} 에")
                    continue
                cn = X.col2num(m.group(1))
                if cn <= prev_c:
                    probs.append(f"{n}: 행 {r} 칸 순서 {prev_c}→{cn}({ref})")
                prev_c = cn
                kids = [X.q(k) for k in ("f", "v", "is")]
                for k in c:
                    if k.tag not in kids:
                        probs.append(f"{n}: {ref} 이상한 자식 {k.tag}")
                if c.get("t") == "s" and c.find(X.q("v")) is None:
                    probs.append(f"{n}: {ref} t=s 인데 v 없음")
        mc = root.find(X.q("mergeCells"))
        if mc is not None:
            seen = []
            for m in mc:
                a, _, b = m.get("ref").partition(":")
                (ra, ca), (rb, cb) = X.cell_key(a), X.cell_key(b or a)
                for (ra2, ca2, rb2, cb2, ref2) in seen:
                    if ra <= rb2 and rb >= ra2 and ca <= cb2 and cb >= ca2:
                        probs.append(f"{n}: 병합 겹침 {m.get('ref')} / {ref2}")
                seen.append((ra, ca, rb, cb, m.get("ref")))
            if int(mc.get("count") or -1) != len(mc):
                probs.append(f"{n}: mergeCells count {mc.get('count')} vs {len(mc)}")
        # 자식 순서
        order = [X.q(t) for t in R.WS_ORDER]
        idx = [order.index(ch.tag) for ch in root if ch.tag in order]
        if idx != sorted(idx):
            probs.append(f"{n}: worksheet 자식 순서 {[etree.QName(ch).localname for ch in root]}")
    st = etree.fromstring(z.read("xl/styles.xml"))
    for tag in ("borders", "cellXfs", "fonts", "fills"):
        e = st.find(X.q(tag))
        if e is not None and int(e.get("count") or -1) != len(e):
            probs.append(f"styles {tag} count {e.get('count')} vs {len(e)}")
    xfs = st.find(X.q("cellXfs"))
    nb = len(st.find(X.q("borders")))
    for i, xf in enumerate(xfs):
        if int(xf.get("borderId") or 0) >= nb:
            probs.append(f"styles xf {i} borderId {xf.get('borderId')} >= {nb}")
    return probs


for p in sys.argv[1:]:
    pr = check(p)
    print(os.path.basename(p), "문제", len(pr))
    for x in pr[:15]:
        print("   ", x)

# fill_budget 안쪽 bisect
ROOT = r"F:\1. 우미\1. 본사업무\★ 판관비 예산안\☆ 예산안 자동작성"
d = os.path.join(ROOT, "app", "_jobs", "bfa16d5fa55f")
res = json.load(open(os.path.join(d, "result.json"), encoding="utf-8"))
form = open(glob.glob(os.path.join(ROOT, "회사양식", "*.xlsx"))[0], "rb").read()
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False


def try_open(path):
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path))
        wb.Close(False)
        return "OK"
    except Exception as e:
        return "FAIL"


orig = dict(set_side=R.set_side, set_row_height=R.set_row_height, write_block=R.write_block, rebuild=R.rebuild)
try:
    for name, off in (("full", []), ("no_side", ["set_side"]), ("no_height", ["set_row_height"]),
                      ("no_side_height", ["set_side", "set_row_height"]), ("rebuild_only", ["set_side", "set_row_height", "write_block"])):
        for k in orig:
            setattr(R, k, orig[k])
        for k in off:
            setattr(R, k, lambda *a, **kw: None)
        book = X.Book(form)
        kit = R.StyleKit(book)
        try:
            R.fill_budget(book.sheet("26년 예산"), kit, res, "t", "d", "s")
        except Exception as e:
            print(name, "예외", e)
            continue
        out = os.path.join(SP, f"inner_{name}.xlsx")
        open(out, "wb").write(book.save())
        pr = check(out)
        print(f"{name:<16} Excel {try_open(out)}  구조문제 {len(pr)} {pr[:3]}")
finally:
    xl.Quit()
