# -*- coding: utf-8 -*-
"""사내 「사업관리비 실적계획」 양식(본사 서버 업로드용)에서 자동작성용 틀을 뽑는다 — 이 PC 의 Excel 사용(PC 전용, 한 번만).

  python tool\\build_siljeok_template.py "<사내 양식.xlsx>" ["<출력.xlsx>"] [--seed 현장명일부]

사내 양식에는 다른 현장 담당자가 적어 둔 현장 블록·보조 시트·옛 이름 정의·외부 연결이 들어 있다.
그 가운데 현장 블록 하나(씨앗)의 여섯 줄만 남긴다 — 서식·열 너비·병합·인쇄 설정은 사내 양식 그대로이고 자료는 비운다.

  1행 제목  [현장명 : … (P코드)]      2~3행 머리글      4행 비목 줄      5행 소 계      6행 합 계

실적계획 자동작성(tool/siljeok.py)이 이 여섯 줄을 현장·비목 수만큼 복제해 채운다.
기본 출력: 회사양식\\실적계획\\사업관리비 실적계획 양식.xlsx  (양식이 바뀌면 다시 만들고 python tool\\sync_web.py)
원본 양식은 읽기만 한다(사본을 고침). Excel 은 따로 띄운 인스턴스를 쓰고 클립보드는 건드리지 않는다.
"""
import io
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DEFAULT = os.path.join(ROOT, "회사양식", "실적계획", "사업관리비 실적계획 양식.xlsx")
MAIN = "사업관리비 실적계획"
LAST_COL = "AB"


def _retry(fn, tries=240, wait=0.5):
    import pywintypes
    for _ in range(tries):
        try:
            return fn()
        except pywintypes.com_error as e:           # Excel 이 바쁠 때(피호출자가 호출을 거부) 다시 시도
            if e.hresult in (-2147418111, -2147417846):
                time.sleep(wait)
                continue
            raise
    raise RuntimeError("Excel 이 응답하지 않습니다.")


def _squash(v):
    return re.sub(r"\s+", "", str(v or ""))


def find_blocks(ws, last_row):
    """[(제목행, 합계행, [(자료행들, 소계행), …])]"""
    R = _retry
    col_a = [None] + [R(lambda r=r: ws.Cells(r, 1).Value) for r in range(1, last_row + 1)]
    blocks, r = [], 1
    while r <= last_row:
        a = col_a[r]
        if isinstance(a, str) and a.strip().startswith("[현장명"):
            groups, cur, rr, total = [], [], r + 3, None
            while rr <= last_row:
                t = _squash(col_a[rr])
                if t == "합계":
                    total = rr
                    break
                if t == "소계":
                    groups.append((cur, rr))
                    cur = []
                elif t.startswith("[현장명"):
                    break
                else:
                    cur.append(rr)
                rr += 1
            if total:
                blocks.append((r, total, groups))
                r = total
        r += 1
    return blocks


def strip_workbook(src, dst, seed=None, log=print):
    import win32com.client as win32
    R = _retry
    work = os.path.join(tempfile.mkdtemp(prefix="siljeok_tpl_"), "work.xlsx")
    shutil.copyfile(src, work)
    os.chmod(work, 0o666)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    xl.EnableEvents = False
    xl.ScreenUpdating = False
    info = {}
    try:
        wb = R(lambda: xl.Workbooks.Open(work, 0, False))
        names = [R(lambda i=i: wb.Worksheets(i).Name) for i in range(1, R(lambda: wb.Worksheets.Count) + 1)]
        main = MAIN if MAIN in names else names[0]
        ws = R(lambda: wb.Worksheets(main))
        if R(lambda: ws.Shapes.Count) or R(lambda: ws.Comments.Count):
            raise SystemExit("양식 시트에 그림·메모가 있습니다 — 틀을 뽑기 전에 확인이 필요합니다.")
        R(lambda: setattr(xl, "Calculation", -4135))            # 수동 계산
        ur = R(lambda: ws.UsedRange)
        last = R(lambda: ur.Row) + R(lambda: ur.Rows.Count) - 1
        blocks = find_blocks(ws, last)
        if not blocks:
            raise SystemExit("'[현장명 : …]' 으로 시작하는 현장 블록을 찾지 못했습니다.")
        titles = [R(lambda b=b: ws.Cells(b[0], 1).Value) for b in blocks]
        if seed:
            cand = [b for b, t in zip(blocks, titles) if seed in t]
            if not cand:
                raise SystemExit(f"'{seed}' 이(가) 든 현장 블록이 없습니다: {titles}")
            blk = cand[0]
        else:
            blk = max(blocks, key=lambda b: len(b[2]))           # 비목이 가장 많은 블록(서식이 온전한 원형)
        t0, tot, groups = blk
        one = [(rows[0], sub) for rows, sub in groups if len(rows) == 1]
        if not one:
            raise SystemExit("비목 한 줄 + 소 계 로 된 묶음이 없는 블록입니다.")
        d0, s0 = one[0]
        info.update(seed_title=R(lambda: ws.Cells(t0, 1).Value), seed_rows=dict(title=t0, head=(t0 + 1, t0 + 2), data=d0, sub=s0, total=tot),
                    blocks=len(blocks))
        log(f"  씨앗 블록: {info['seed_title']}  (제목 {t0} · 머리글 {t0 + 1}~{t0 + 2} · 비목 {d0} · 소계 {s0} · 합계 {tot}행)")
        keep = [t0, t0 + 1, t0 + 2, d0, s0, tot]
        R(lambda: setattr(ws.Rows(f"1:{last + 5}"), "Hidden", False))
        # 아래에서 위로 지워 간다(남길 행 번호가 밀리지 않게)
        drop, prev = [], last + 50
        for k in sorted(keep, reverse=True):
            if k + 1 <= prev:
                drop.append((k + 1, prev))
            prev = k - 1
        if prev >= 1:
            drop.append((1, prev))
        for a, b in drop:
            if a <= b:
                R(lambda a=a, b=b: ws.Rows(f"{a}:{b}").Delete())
        chk = [_squash(R(lambda r=r: ws.Cells(r, 1).Value)) for r in range(1, 7)]
        if not (chk[0].startswith("[현장명") and chk[1] == "비목" and chk[4] == "소계" and chk[5] == "합계"):
            raise SystemExit(f"씨앗 행을 추리지 못했습니다: {chk}")
        # 자료 비우기(머리글 글자·'소 계'·'합 계'·'(금액단위: 원)' 은 둠)
        R(lambda: setattr(ws.Range("A1"), "Value", "[현장명 : 현장명 (P0000)]"))
        R(lambda: ws.Range(f"A4:{LAST_COL}4").ClearContents())
        R(lambda: ws.Range(f"C5:{LAST_COL}5").ClearContents())
        R(lambda: ws.Range(f"C6:{LAST_COL}6").ClearContents())
        R(lambda: ws.Range(f"AB1:AB6").ClearContents())
        R(lambda: setattr(ws.Rows(4), "UseStandardHeight", True))                        # 비목 줄 높이는 한 줄 기준(표준 높이)
        # 다른 시트
        gone = 0
        for nm in names:
            if nm != main:
                R(lambda nm=nm: wb.Worksheets(nm).Delete())
                gone += 1
        if main != MAIN:
            R(lambda: setattr(ws, "Name", MAIN))
        # 이름 정의(옛 파일에서 딸려 온 것) — 숨김 이름까지
        n_names = R(lambda: wb.Names.Count)
        left = 0
        for i in range(n_names, 0, -1):
            try:
                nm = R(lambda i=i: wb.Names(i))
                try:
                    nm.Visible = True
                except Exception:      # noqa: BLE001
                    pass
                R(lambda nm=nm: nm.Delete())
            except Exception:          # noqa: BLE001
                left += 1
        # 외부 연결
        links = R(lambda: wb.LinkSources(1)) or ()
        broken = 0
        for ln in links:
            try:
                R(lambda ln=ln: wb.BreakLink(ln, 1))
                broken += 1
            except Exception:          # noqa: BLE001
                pass
        # 쓰지 않는 셀 스타일(옛 파일의 이름 붙은 스타일 수천 개)
        used = set()
        for r in range(1, 7):
            for c in range(1, 29):
                try:
                    used.add(R(lambda r=r, c=c: ws.Cells(r, c).Style.Name))
                except Exception:      # noqa: BLE001
                    pass
        n_styles = R(lambda: wb.Styles.Count)
        removed = 0
        for i in range(n_styles, 0, -1):
            try:
                st = R(lambda i=i: wb.Styles(i))
                if not R(lambda st=st: st.BuiltIn) and R(lambda st=st: st.Name) not in used:
                    R(lambda st=st: st.Delete())
                    removed += 1
            except Exception:          # noqa: BLE001
                pass
        R(lambda: ws.ResetAllPageBreaks())
        R(lambda: setattr(ws.PageSetup, "PrintArea", "$A$1:$AA$6"))
        R(lambda: setattr(xl, "Calculation", -4105))
        R(lambda: ws.Activate())
        try:
            R(lambda: setattr(xl.ActiveWindow, "ScrollRow", 1))
            R(lambda: setattr(xl.ActiveWindow, "ScrollColumn", 1))
            R(lambda: ws.Range("A1").Select())
        except Exception:              # noqa: BLE001
            pass
        info.update(sheets_removed=gone, names=n_names, names_left=left, links=len(links), links_broken=broken,
                    styles=n_styles, styles_removed=removed, styles_used=sorted(used))
        R(lambda: wb.Save())
        R(lambda: wb.Close(False))
    finally:
        try:
            _retry(lambda: xl.Quit(), tries=40)
        except Exception:              # noqa: BLE001
            pass
    clean_package(work, dst, sheet_layout(src, main))
    shutil.rmtree(os.path.dirname(work), ignore_errors=True)
    return info


def sheet_layout(path, sheet_name):
    """원본 양식 시트 XML 의 열 너비(<cols>)·기본 행 높이(<sheetFormatPr>) 글자 그대로.
    Excel 은 저장할 때 이 값을 그 PC 글꼴 기준으로 다시 맞추므로(12.898 → 12.875), 틀에는 원본 값을 되돌려 넣는다."""
    try:
        z = zipfile.ZipFile(path)
        wbx = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        rid = None
        for m in re.finditer(r"<sheet\b[^>]*/>", wbx):
            tag = m.group(0)
            nm = re.search(r'name="([^"]*)"', tag)
            if nm and nm.group(1).replace("&amp;", "&") == sheet_name:
                rid = re.search(r'r:id="([^"]*)"', tag).group(1)
        if not rid:
            return {}
        target = None
        for m in re.finditer(r"<Relationship\b[^>]*/>", rels):
            tag = m.group(0)
            if re.search(r'Id="%s"' % re.escape(rid), tag):
                target = re.search(r'Target="([^"]*)"', tag).group(1)
        if not target:
            return {}
        part = target.lstrip("/") if target.startswith("/") else "xl/" + target
        x = z.read(part).decode("utf-8")
        out = {}
        m = re.search(r"<sheetFormatPr\b[^>]*/>", x)
        if m:
            out["sheetFormatPr"] = m.group(0)
        m = re.search(r"<cols>.*?</cols>", x, flags=re.S)
        if m:
            out["cols"] = m.group(0)
        return out
    except Exception:                  # noqa: BLE001
        return {}


def prune_styles(data):
    """Excel 이 못 지운 옛 이름 스타일(다른 공사 이름이 든 것)과 그것만 쓰던 글꼴·표시형식을 styles.xml 에서 뺀다.
    칸 서식(cellXfs)은 그대로 — 모두 '표준'(xfId 0)에 딸려 있을 때만 정리한다."""
    from lxml import etree
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    q = lambda t: "{%s}%s" % (ns, t)
    root = etree.fromstring(data)
    cell_xfs, style_xfs, styles = root.find(q("cellXfs")), root.find(q("cellStyleXfs")), root.find(q("cellStyles"))
    if cell_xfs is None or style_xfs is None or styles is None:
        return data
    if any((xf.get("xfId") or "0") != "0" for xf in cell_xfs):
        return data
    for xf in list(style_xfs)[1:]:
        style_xfs.remove(xf)
    style_xfs.set("count", "1")
    for st in list(styles):
        if (st.get("xfId") or "0") != "0":
            styles.remove(st)
    styles.set("count", str(len(styles)))
    xfs = list(cell_xfs) + list(style_xfs)
    fonts = root.find(q("fonts"))
    if fonts is not None:
        used = {int(xf.get("fontId") or 0) for xf in xfs}
        keep = max(max(used), 6) + 1                 # 0~6: 기본 글꼴과 칸 글꼴, 윗주(phoneticPr fontId) 글꼴
        fl = list(fonts)
        if all(i < keep for i in used):
            for f in fl[keep:]:
                fonts.remove(f)
            fonts.set("count", str(len(fonts)))
    nf = root.find(q("numFmts"))
    if nf is not None:
        used = {xf.get("numFmtId") for xf in xfs}
        for f in list(nf):
            if f.get("numFmtId") not in used:
                nf.remove(f)
        nf.set("count", str(len(nf)))
        if not len(nf):
            root.remove(nf)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def clean_package(src, dst, layout=None):
    """저장된 xlsx 에서 사람 이름이 남는 곳을 지운다: 문서 속성(만든 이·고친 이·회사), 남은 이름 정의(인쇄 영역만 둠), 외부 연결 부품.
    layout 이 있으면 시트의 열 너비·기본 행 높이를 원본 양식 값으로 되돌린다."""
    layout = layout or {}
    zin = zipfile.ZipFile(src)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in zin.infolist():
            name = item.filename
            data = zin.read(name)
            if name.startswith("xl/externalLinks/") or name == "xl/calcChain.xml":
                continue
            if name == "docProps/core.xml":
                s = data.decode("utf-8")
                s = re.sub(r"<dc:creator>.*?</dc:creator>", "<dc:creator></dc:creator>", s, flags=re.S)
                s = re.sub(r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", "<cp:lastModifiedBy></cp:lastModifiedBy>", s, flags=re.S)
                s = re.sub(r"<cp:lastPrinted>.*?</cp:lastPrinted>", "", s, flags=re.S)
                data = s.encode("utf-8")
            elif name == "docProps/app.xml":
                s = data.decode("utf-8")
                s = re.sub(r"<Company>.*?</Company>", "<Company></Company>", s, flags=re.S)
                s = re.sub(r"<Manager>.*?</Manager>", "", s, flags=re.S)
                data = s.encode("utf-8")
            elif name == "xl/workbook.xml":
                s = data.decode("utf-8")
                s = re.sub(r"<externalReferences>.*?</externalReferences>", "", s, flags=re.S)

                def keep_print(m):
                    kept = re.findall(r'<definedName [^>]*name="_xlnm\.Print_(?:Area|Titles)"[^>]*>.*?</definedName>', m.group(0), flags=re.S)
                    kept = [k for k in kept if "#REF!" not in k]
                    return "<definedNames>" + "".join(kept) + "</definedNames>" if kept else ""

                s = re.sub(r"<definedNames>.*?</definedNames>", keep_print, s, flags=re.S)
                data = s.encode("utf-8")
            elif name == "xl/_rels/workbook.xml.rels":
                s = data.decode("utf-8")
                s = re.sub(r'<Relationship [^>]*Type="[^"]*/(?:externalLink|calcChain)"[^>]*/>', "", s)
                data = s.encode("utf-8")
            elif name == "xl/styles.xml":
                data = prune_styles(data)
            elif name.startswith("xl/worksheets/sheet") and name.endswith(".xml") and layout:
                s = data.decode("utf-8")
                if layout.get("sheetFormatPr"):
                    s = re.sub(r"<sheetFormatPr\b[^>]*/>", lambda m: layout["sheetFormatPr"], s, count=1)
                if layout.get("cols"):
                    s = re.sub(r"<cols>.*?</cols>", lambda m: layout["cols"], s, count=1, flags=re.S)
                data = s.encode("utf-8")
            elif name == "[Content_Types].xml":
                s = data.decode("utf-8")
                s = re.sub(r'<Override [^>]*PartName="/xl/(?:externalLinks/[^"]*|calcChain\.xml)"[^>]*/>', "", s)
                data = s.encode("utf-8")
            z.writestr(name, data)
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    with open(dst, "wb") as f:
        f.write(out.getvalue())


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:                  # noqa: BLE001
        pass
    args = list(sys.argv[1:])
    seed = None
    if "--seed" in args:
        i = args.index("--seed")
        seed = args[i + 1]
        del args[i:i + 2]
    if not args:
        sys.exit(__doc__)
    src = os.path.abspath(args[0])
    dst = os.path.abspath(args[1]) if len(args) > 1 else OUT_DEFAULT
    if not os.path.exists(src):
        sys.exit("파일이 없습니다: " + src)
    print("사내 양식:", src)
    info = strip_workbook(src, dst, seed)
    print(f"  다른 시트 {info['sheets_removed']}개 · 이름 정의 {info['names']}개(남음 {info['names_left']}) · 외부 연결 {info['links']}개(끊음 {info['links_broken']})"
          f" · 셀 스타일 {info['styles']}개 중 {info['styles_removed']}개 정리")
    print("틀 저장:", dst, f"({os.path.getsize(dst):,} bytes)")


if __name__ == "__main__":
    main()
