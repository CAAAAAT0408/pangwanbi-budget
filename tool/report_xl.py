# -*- coding: utf-8 -*-
"""회사 양식(경영기획팀 '사업비 예산편성 양식')에 계산 결과를 채운다 — Excel 없이(xlsx 내부 XML 편집, 브라우저 Pyodide 용).

report.py(Excel COM 판)와 같은 규칙·같은 결과: 행이 더 필요하면 양식의 원래 행(첫 줄·중간 줄·소계·합계)을 복사해
끼우고 견본 행은 지운다(xlsxio 가 수식·병합·인쇄영역을 같이 옮김). 수식 값은 Excel 로 열 때 다시 계산된다.

  build_report(form_bytes, res, data) → (xlsx bytes, dict(budget_sheet=…))
"""
import copy
import datetime
import json
import math
import re

from lxml import etree

import xlsxio as X
from xlsxio import q, NS, RNS, PNS, CNS, col2num, num2col, cell_key
from report import DATA_SHEET, L, _txt, _find, _lines, xround, _groups_by_size, _distribute

WS_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
WS_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection',
            'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells',
            'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup',
            'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags',
            'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems',
            'tableParts', 'extLst']
STYLE_ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles',
               'colors', 'extLst']
BORDER_ORDER = ['start', 'end', 'left', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal']


# ── xlsx 편집 도우미 (분양개요 웹판 build_web.py 와 같은 방식) ───────────────
def cref(r, c):
    return f'{num2col(c)}{r}'


def localname(el):
    return etree.QName(el).localname if isinstance(el.tag, str) else ''


def _insert_in(parent, el, order):
    idx = order.index(localname(el))
    after = None
    for ch in parent:
        ln = localname(ch)
        if ln in order and order.index(ln) < idx:
            after = ch
    if after is None:
        parent.insert(0, el)
    else:
        after.addnext(el)


def set_formula(sh, ref, f):
    c = sh._cell(ref)
    sh._clear(c)
    e = etree.SubElement(c, q('f'))
    e.text = f
    sh.book.touched.add(sh.path)


def put(sh, r, c, v):
    """값 쓰기: '=…' 은 수식, 숫자는 숫자, '' / None 은 내용 지우기(서식은 둠) — COM 의 .Formula 와 같게"""
    ref = cref(r, c)
    if v is None or (isinstance(v, str) and v == ''):
        sh.clear(ref)
        return
    if isinstance(v, bool):
        v = int(v)
    if isinstance(v, (int, float)):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            sh.clear(ref)
            return
        sh.set_number(ref, v)
        return
    s = str(v)
    if s.startswith('=') and len(s) > 1:
        set_formula(sh, ref, s[1:])
        return
    if re.fullmatch(r'-?\d+(\.\d+)?', s.strip()):      # Excel 이 .Formula 로 받은 숫자 글자는 숫자로 두는 것과 같게
        sh.set_number(ref, float(s))
        return
    sh.set_text(ref, s)


def vals(sh, last_col='V', last_row=None):
    last_row = last_row or (sh.max_row() + 2)
    nc = col2num(last_col)
    return [[sh.value(cref(r, c)) for c in range(1, nc + 1)] for r in range(1, last_row + 1)]


def write_block(sh, r0, rows, first_col, last_col):
    c0 = col2num(first_col)
    ncol = col2num(last_col) - c0 + 1
    for i, row in enumerate(rows):
        row = list(row) + [''] * (ncol - len(row))
        for j, v in enumerate(row[:ncol]):
            put(sh, r0 + i, c0 + j, v)


def clear_range(sh, r1, c1, r2, c2):
    for ref, c in list(sh.cells.items()):
        r, cc = cell_key(ref)
        if r1 <= r <= r2 and c1 <= cc <= c2:
            sh._clear(c)
    sh.book.touched.add(sh.path)


def unmerge_range(sh, r1, c1, r2, c2):
    mc = sh.root.find(q('mergeCells'))
    if mc is None:
        return
    for m in list(mc):
        a, _, b = m.get('ref').partition(':')
        (ra, ca), (rb, cb) = cell_key(a), cell_key(b or a)
        if ra <= r2 and rb >= r1 and ca <= c2 and cb >= c1:
            mc.remove(m)
    mc.set('count', str(len(mc)))
    sh.book.touched.add(sh.path)


def merge_range(sh, r1, c1, r2, c2):
    unmerge_range(sh, r1, c1, r2, c2)
    if (r1, c1) == (r2, c2):
        return
    sh.merge(f'{cref(r1, c1)}:{cref(r2, c2)}')


def set_row_height(sh, r, h):
    if r not in sh.rows:
        sh._cell(cref(r, 1))
    row = sh.rows[r]
    row.set('ht', repr(float(h)).rstrip('0').rstrip('.'))
    row.set('customHeight', '1')
    sh.book.touched.add(sh.path)


def col_width(sh, c):
    cols = sh.root.find(q('cols'))
    if cols is not None:
        for col in cols:
            if int(col.get('min')) <= c <= int(col.get('max')) and col.get('width'):
                return float(col.get('width'))
    fp = sh.root.find(q('sheetFormatPr'))
    if fp is not None and fp.get('defaultColWidth'):
        return float(fp.get('defaultColWidth'))
    return 8.38


def set_col_width(sh, c1, c2, width):
    cols = sh.root.find(q('cols'))
    if cols is None:
        cols = etree.Element(q('cols'))
        _insert_in(sh.root, cols, WS_ORDER)
    keep = []
    for col in list(cols):
        mn, mx = int(col.get('min')), int(col.get('max'))
        cols.remove(col)
        if mx < c1 or mn > c2:
            keep.append(col)
            continue
        if mn < c1:
            a = copy.deepcopy(col)
            a.set('max', str(c1 - 1))
            keep.append(a)
        if mx > c2:
            b = copy.deepcopy(col)
            b.set('min', str(c2 + 1))
            keep.append(b)
    new = etree.Element(q('col'))
    new.set('min', str(c1))
    new.set('max', str(c2))
    new.set('width', str(width))
    new.set('customWidth', '1')
    keep.append(new)
    keep.sort(key=lambda e: int(e.get('min')))
    for e in keep:
        cols.append(e)
    sh.book.touched.add(sh.path)


def widen(sh, cols, max_value):
    """금액이 ##### 로 보이지 않게: 가장 큰 값의 글자 수만큼 넓힘(양식보다 좁히지는 않음) — COM AutoFit 대신"""
    need = len(f"{int(abs(max_value)):,}") * 1.15 + 1.5
    for col in cols:
        c = col2num(col)
        if col_width(sh, c) < need:
            set_col_width(sh, c, c, round(need, 2))


class StyleKit:
    """셀 서식(cellXfs) 파생: 테두리 한 변 바꾸기"""

    def __init__(self, book):
        if book._styles is None:
            book._styles = etree.fromstring(book.files['xl/styles.xml'])
        self.book = book
        self.root = book._styles
        self.cache = {}

    def _list(self, tag):
        e = self.root.find(q(tag))
        if e is None:
            e = etree.Element(q(tag))
            _insert_in(self.root, e, STYLE_ORDER)
        return e

    def border_of(self, s):
        xf = self._list('cellXfs')[s]
        return self._list('borders')[int(xf.get('borderId') or 0)]

    def side_el(self, s, side):
        return self.border_of(s).find(q(side))

    def with_side(self, s, side, side_el):
        key = (s, side, etree.tostring(side_el) if side_el is not None else None)
        if key in self.cache:
            return self.cache[key]
        xfs = self._list('cellXfs')
        nx = copy.deepcopy(xfs[s])
        be = copy.deepcopy(self.border_of(s))
        old = be.find(q(side))
        if old is not None:
            be.remove(old)
        if side_el is not None:
            new = copy.deepcopy(side_el)
            new.tag = q(side)                      # D 의 <top> 을 E 의 <bottom> 으로도 쓰므로 이름을 바꾼다
        else:
            new = etree.Element(q(side))
            new.set('style', 'hair')
            col = etree.SubElement(new, q('color'))
            col.set('indexed', '64')
        _insert_in(be, new, BORDER_ORDER)
        borders = self._list('borders')
        borders.append(be)
        borders.set('count', str(len(borders)))
        nx.set('borderId', str(len(borders) - 1))
        nx.set('applyBorder', '1')
        xfs.append(nx)
        xfs.set('count', str(len(xfs)))
        self.cache[key] = len(xfs) - 1
        return self.cache[key]


def set_side(sh, kit, ref, side, side_el):
    s = sh.style(ref)
    if s is None:
        sh._cell(ref)
        s = 0
    sh.set_style(ref, kit.with_side(s, side, side_el))
    sh.book.touched.add(sh.path)


def sheet_state(book, name, state):
    """'visible' | 'hidden' | 'veryHidden'"""
    for s in book.sheet_list():
        if s['name'] == name:
            if state == 'visible':
                s['el'].attrib.pop('state', None)
            else:
                s['el'].set('state', state)
            book.touched.add('xl/workbook.xml')
            book._structure_changed = True
            return True
    return False


def activate(book, name):
    names = [s['name'] for s in book.sheet_list()]
    if name in names:
        for wv in book.wb.iter(q('workbookView')):
            wv.set('activeTab', str(names.index(name)))
        book._structure_changed = True


def add_sheet(book, name, rows, hidden=True):
    """rows(줄 목록)로 새 시트를 맨 끝에 만든다 (글자는 공유문자열)."""
    nums = [int(m.group(1)) for n in list(book.files) + book.order for m in [re.match(r'xl/worksheets/sheet(\d+)\.xml$', n)] if m]
    path = f'xl/worksheets/sheet{max(nums + [0]) + 1}.xml'
    root = etree.Element(q('worksheet'), nsmap={None: NS, 'r': RNS})
    sv = etree.SubElement(etree.SubElement(root, q('sheetViews')), q('sheetView'))
    sv.set('workbookViewId', '0')
    fp = etree.SubElement(root, q('sheetFormatPr'))
    fp.set('defaultRowHeight', '16.5')
    data = etree.SubElement(root, q('sheetData'))
    for i, row in enumerate(rows, 1):
        re_ = etree.SubElement(data, q('row'))
        re_.set('r', str(i))
        for j, v in enumerate(row, 1):
            if v is None or v == '':
                continue
            c = etree.SubElement(re_, q('c'))
            c.set('r', cref(i, j))
            if isinstance(v, bool):
                v = int(v)
            if isinstance(v, (int, float)):
                e = etree.SubElement(c, q('v'))
                e.text = str(int(v)) if float(v).is_integer() else repr(float(v))
            else:
                c.set('t', 's')
                e = etree.SubElement(c, q('v'))
                e.text = str(book.add_string(str(v)))
    pm = etree.SubElement(root, q('pageMargins'))
    for k, v in (('left', '0.7'), ('right', '0.7'), ('top', '0.75'), ('bottom', '0.75'), ('header', '0.3'), ('footer', '0.3')):
        pm.set(k, v)
    book.files[path] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    book.order.append(path)
    ids = [int(re.sub(r'\D', '', r.get('Id') or '') or 0) for r in book.rels]
    rid = f'rId{max(ids + [0]) + 1}'
    rel = etree.SubElement(book.rels, q('Relationship', PNS))
    rel.set('Id', rid)
    rel.set('Type', X.WS_TYPE)
    rel.set('Target', path[len('xl/'):])
    ov = etree.SubElement(book.ct, q('Override', CNS))
    ov.set('PartName', '/' + path)
    ov.set('ContentType', WS_CT)
    sheets = book.wb.find(q('sheets'))
    sid = max(int(s.get('sheetId')) for s in sheets) + 1
    el = etree.SubElement(sheets, q('sheet'))
    el.set('name', name)
    el.set('sheetId', str(sid))
    if hidden:
        el.set('state', 'hidden')
    el.set(q('id', RNS), rid)
    book.touched.add('xl/workbook.xml')
    book._structure_changed = True
    return book.sheet(name)


def row_merges(sh, r):
    """r 행 안에서 끝나는(한 줄짜리) 병합의 열 범위들"""
    return [(a[1], b[1]) for a, b in sh.merged() if a[0] == r and b[0] == r]


def copy_row(sh, snap, merges, dst):
    """row_copy 로 떠 둔 행(서식·높이·수식)과 그 줄의 병합을 dst 에 놓는다"""
    sh.put_row(snap, dst)
    for c1, c2 in merges:
        sh.merge(f'{cref(dst, c1)}:{cref(dst, c2)}')


def rebuild(sh, start, end, layout):
    """start..end 행(견본)을 layout 대로 새로 짠다 — report.rebuild 와 같은 결과"""
    n = len(layout)
    sh.unshare()
    snaps = {src: (sh.row_copy(src), row_merges(sh, src)) for src in set(layout)}
    sh.insert_rows(end + 1, n)
    for i, src in enumerate(layout):
        copy_row(sh, snaps[src][0], snaps[src][1], end + 1 + i)
    sh.delete_rows(start, end - start + 1)
    return start, start + n - 1


def shift_years(book, fy, ty):
    """양식 연도(fy) 글자를 예산연도(ty)에 맞게: '25년 실적'→'26년 실적', '2026년'→'2027년' (상수 글만, 리치텍스트 유지)"""
    d = ty - fy
    if not d:
        return 0
    cnt = 0
    rx = re.compile(r"(?<!\d)(20\d{2}|\d{2})(?=\s*년)")
    for sh in book.all_sheets():
        for ref, c in list(sh.cells.items()):
            if c.get('t') not in ('s', 'inlineStr'):
                continue
            s = sh.value(ref)
            if not isinstance(s, str) or '년' not in s:
                continue
            pairs = set()
            for m in rx.finditer(s):
                y = int(m.group(1))
                full = y if y >= 2000 else 2000 + y
                if fy - 3 <= full <= fy + 1:
                    new = full + d
                    pairs.add((m.group(1), str(new) if y >= 2000 else f"{new % 100:02d}"))
            for old, new in sorted(pairs, key=lambda p: -int(p[0])):     # 큰 연도부터(겹침 방지)
                if sh.replace_text(ref, old + '년', new + '년') or sh.replace_text(ref, old, new):
                    cnt += 1
    return cnt


# ── 27년 예산 ─────────────────────────────────────────────────────
def fill_budget(sh, kit, res, title, date_txt, site, keep_interest=False):
    v = vals(sh, 'U')
    hr = _find(v, 3, lambda t: t == '세목')
    tot = _find(v, 0, lambda t: t.startswith('합'), hr)
    sub = _find(v, 3, lambda t: t == '소계', hr)
    T = dict(first=hr + 1, mid=hr + 2, sub=sub, first2=sub + 1, tot=tot)
    for g, gf in zip(res['groups'], res.get('formulas') or []):      # 월별 산식(계산 과정) — 줄마다 12개월, 없으면 None
        for r, rf in zip(g['rows'], gf or []):
            r['fx'] = rf
    groups = []
    for g in res['groups']:
        rows = [r for r in g['rows'] if r['total'] or (r['manual'] and r['g'])]
        if not rows and keep_interest and g['sub'] == '대납이자':
            rows = [r for r in g['rows'] if not r['manual']][:1]
        if rows:
            groups.append((g, rows))
    layout, kinds = [], []
    for gi, (g, rows) in enumerate(groups):
        for k, r in enumerate(rows):
            layout.append(T['first'] if (gi == 0 and k == 0) else (T['first2'] if k == 0 else T['mid']))
            kinds.append(('row', gi, k, r))
        layout.append(T['sub'])
        kinds.append(('sub', gi, None, None))
    a, b = rebuild(sh, T['first'], T['tot'] - 1, layout)
    tot_row = b + 1
    out, subs, sub_of, start_of = [], [], {}, {}
    for i, (kind, gi, k, r) in enumerate(kinds):
        rr = a + i
        g = groups[gi][0]
        if kind == 'row':
            start_of.setdefault(gi, rr)
            fx = r.get('fx') or [None] * 12
            months = [('=' + f) if f else ('' if not m else m) for m, f in zip(r['months'], fx)]   # 산식이 있으면 수식으로
            out.append([g['item'] if k == 0 else '', '', g['sub'] if k == 0 else '', '', f'=I{rr}', ' ' + r['g'].strip(),
                        '', f'=SUM(J{rr}:U{rr})'] + months)
            set_row_height(sh, rr, 19.5 if _lines(r['g']) == 1 else 15 * _lines(r['g']) + 6)
        else:
            s0 = start_of[gi]
            sums = [f'=SUM({L(10 + j)}{s0}:{L(10 + j)}{rr - 1})' for j in range(12)]
            out.append(['', '', '소  계', f'=SUM(E{s0}:E{rr - 1})', f'=SUM(F{s0}:F{rr - 1})', '', '', f'=SUM(J{rr}:U{rr})'] + sums)
            subs.append(rr)
            sub_of[gi] = rr
    write_block(sh, a, out, 'B', 'U')
    # 실적 열(E) 가운데 줄 윗선이 양식에 없음 → 묶음마다 세목 열(D)과 같은 점선을 다시 그림 (report.fill_budget 과 같게)
    for gi in start_of:
        s0, s1 = start_of[gi], sub_of[gi] - 1
        if s1 > s0:
            src = kit.side_el(sh.style(f'D{s0 + 1}') or 0, 'top')
            for rr in range(s0 + 1, s1 + 1):
                set_side(sh, kit, f'E{rr}', 'top', src)
            for rr in range(s0, s1):
                set_side(sh, kit, f'E{rr}', 'bottom', src)
    plus = lambda c: '=' + '+'.join(f'{c}{s}' for s in subs) if subs else ''
    write_block(sh, tot_row, [[plus('E'), plus('F'), '', '', f'=SUM(J{tot_row}:U{tot_row})']
                              + [plus(L(10 + j)) for j in range(12)]], 'E', 'U')
    put(sh, 2, 1, title)
    put(sh, 3, 1, date_txt)
    put(sh, hr - 1, 1, f'[ {site} ]')
    labels = res.get('stage') or []
    write_block(sh, hr - 1, [[x or '' for x in labels]], 'J', 'U')
    return dict(groups=[(groups[gi][0], start_of[gi], sub_of[gi]) for gi in range(len(groups))], total=tot_row, header=hr)


# ── 예산안 보고서 ─────────────────────────────────────────────────
def fill_report(sh, budget_name, binfo, title, date_txt, site, note_txt):
    v = vals(sh, 'P')
    hr = _find(v, 1, lambda t: t == '항목')
    first = hr + 2
    sub = _find(v, 3, lambda t: t == '소계', hr)
    tot = _find(v, 0, lambda t: t.startswith('합'), hr)
    T = dict(first=first, mid=first + 1, sub=sub, first2=sub + 1, tot=tot)
    items = []
    for g, s0, s1 in binfo['groups']:
        if items and items[-1][0] == g['item']:
            items[-1][1].append((g, s0, s1))
        else:
            items.append((g['item'], [(g, s0, s1)]))
    layout, kinds = [], []
    for ii, (item, gs) in enumerate(items):
        for k, x in enumerate(gs):
            layout.append(T['first'] if (ii == 0 and k == 0) else (T['first2'] if k == 0 else T['mid']))
            kinds.append(('row', ii, k, x))
        layout.append(T['sub'])
        kinds.append(('sub', ii, None, None))
    a, b = rebuild(sh, T['first'], T['tot'] - 1, layout)
    tot_row = b + 1
    out, subs, start = [], [], None
    bq = X._qname(budget_name)
    for i, (kind, ii, k, x) in enumerate(kinds):
        rr = a + i
        if kind == 'row':
            if k == 0:
                start = rr
            g, s0, s1 = x
            out.append([items[ii][0] if k == 0 else '', '', g['sub'], '', '', '', '', '', '',
                        f'=IF(H{rr}-E{rr}=0,"",IF(E{rr}=0,100%,(H{rr}-E{rr})/E{rr}))', '',
                        f"={bq}!F{s1}",
                        f'=IF(M{rr}-H{rr}=0,"",IF(H{rr}=0,100%,(M{rr}-H{rr})/H{rr}))', note_txt, ''])
        else:
            sm = lambda c: f'=SUM({c}{start}:{c}{rr - 1})'
            out.append(['', '', '소  계', sm('E'), sm('F'), sm('G'), sm('H'), sm('I'), sm('J'),
                        f'=IF(H{rr}-E{rr}=0,"",IF(E{rr}=0,100%,(H{rr}-E{rr})/E{rr}))', '', sm('M'),
                        f'=IF(M{rr}-H{rr}=0,"",IF(H{rr}=0,100%,(M{rr}-H{rr})/H{rr}))', '', ''])
            subs.append(rr)
    write_block(sh, a, out, 'B', 'P')
    plus = lambda c: '=' + '+'.join(f'{c}{s}' for s in subs) if subs else ''
    write_block(sh, tot_row, [[plus(c) for c in 'EFGHIJ'] +
                              [f'=IF(H{tot_row}-E{tot_row}=0,"",IF(E{tot_row}=0,100%,(H{tot_row}-E{tot_row})/E{tot_row}))', '',
                               plus('M'), f'=IF(M{tot_row}-H{tot_row}=0,"",IF(H{tot_row}=0,100%,(M{tot_row}-H{tot_row})/H{tot_row}))']],
                'E', 'N')
    put(sh, 2, 1, title)
    put(sh, 3, 1, date_txt)
    put(sh, hr - 1, 1, f'[ {site} ]')
    return dict(header=hr, total=tot_row)


# ── 기본정보 ─────────────────────────────────────────────────────
def fill_basic(sh, kit, fields, types):
    v = vals(sh, 'G', 40)
    rcnt = _find(v, 3, lambda t: t == '계')
    t_rows = [i + 1 for i in range(rcnt, len(v)) if _txt(v[i][3]) == '평형입력']
    if not t_rows:
        t_rows = [rcnt + 1, rcnt + 2]
    t1, tn = t_rows[0], t_rows[-1]
    m = len(types)
    if m > len(t_rows):
        k = m - len(t_rows)
        sh.unshare()
        snap, merges = sh.row_copy(t1), row_merges(sh, t1)
        sh.insert_rows(tn, k)
        for i in range(k):
            copy_row(sh, snap, merges, tn + i)
            for c in 'DEFG':
                set_side(sh, kit, f'{c}{tn + i}', 'top', None)
                set_side(sh, kit, f'{c}{tn + i}', 'bottom', None)
        tn += k
    rows = []
    for i in range(tn - t1 + 1):
        t = types[i] if i < len(types) else None
        rows.append([str(t[0]) if t else '', '', (t[1] if t and t[1] not in (None, '') else '')])
    write_block(sh, t1, rows, 'D', 'F')
    put(sh, rcnt, 6, f'=SUM(F{t1}:F{tn})')
    v = vals(sh, 'G', 40)
    fill = {'현장명': fields.get('현장명'), '담당소장': fields.get('담당소장'), '착공년월': fields.get('착공'),
            '준공년월': fields.get('준공'), '시행법인': fields.get('시행'),
            '시공법인': fields.get('시공') or fields.get('시행')}
    for i, row in enumerate(v):
        lab = _txt(row[2]) or _txt(row[3])
        if lab in fill and fill[lab] not in (None, ''):
            put(sh, i + 1, 6, str(fill[lab]))
    return dict(total=f'F{rcnt}')


# ── 분양율,입주율 ─────────────────────────────────────────────────
def fill_rates(sh, res, basic_total):
    v = vals(sh, 'U', 30)
    r_s = _find(v, 2, lambda t: t == '분양세대')
    r_i = _find(v, 2, lambda t: t == '입주세대')
    put(sh, 4, 3, f'=기본정보!{basic_total}')
    R = res['rates']

    def nums(xs):
        return ['' if not x else x for x in xs]
    if r_s:
        write_block(sh, r_s, [nums(R['전년분양']) + nums(R['분양'])], 'E', 'T')
    if r_i:
        write_block(sh, r_i, [nums(R['전년입주']) + nums(R['입주'])], 'E', 'T')


# ── 분양대행 및 촉진수수료 ─────────────────────────────────────────
def fill_agency(book, sh, res, types):
    c = res['calc']
    keys, units = _groups_by_size(types)
    if not keys:
        keys, units = ['전체'], {'전체': 1.0}
    series = {'대행': c.get('대행세대') or [0] * 12, '조직': c.get('조직세대') or [0] * 12, 'MGM': c.get('MGM') or [0] * 12}
    if not any(sum(x) for x in series.values()):
        sheet_state(book, sh.name, 'hidden')
        return False
    v = vals(sh, 'V', 40)
    r0 = _find(v, 0, lambda t: t == '분양대행수수료')
    rm = _find(v, 0, lambda t: t.startswith('분양촉진수수료('))
    if not (r0 and rm and _txt(v[r0 - 1][1]) == '세대수' and _txt(v[r0 + 3][1]) == '금액' and rm - r0 == 10):
        return False
    S = dict(u1=r0, um=r0 + 1, ul=r0 + 2, us=r0 + 3, a1=r0 + 4, am=r0 + 5, al=r0 + 6, as_=r0 + 7, cs=r0 + 8, tot=r0 + 9,
             m1=rm, mm=rm + 1, ml=rm + 2, ms=rm + 3, ma1=rm + 4, mam=rm + 5, mal=rm + 6, mas=rm + 7)
    n = len(keys)

    def pick(f, m, l):
        return [f] if n == 1 else [f] + [m] * (n - 2) + [l]
    lay, kinds = [], []

    def add(name, mgm):
        p = (S['m1'], S['mm'], S['ml'], S['ms'], S['ma1'], S['mam'], S['mal'], S['mas']) if mgm else \
            (S['u1'], S['um'], S['ul'], S['us'], S['a1'], S['am'], S['al'], S['as_'])
        lay.extend(pick(p[0], p[1], p[2]))
        kinds.extend((name, 'u', i) for i in range(n))
        lay.append(p[3])
        kinds.append((name, 'usum', None))
        lay.extend(pick(p[4], p[5], p[6]))
        kinds.extend((name, 'a', i) for i in range(n))
        lay.append(p[7])
        kinds.append((name, 'asum', None))
        if not mgm:
            lay.extend([S['cs'], S['tot']])
            kinds.extend([(name, 'consult', None), (name, 'total', None)])
    add('대행', False)
    if sum(series['조직']) > 0:
        add('조직', False)
        lay.append(S['tot'])
        kinds.append(('총계', 'grand', None))
    add('MGM', True)
    unmerge_range(sh, r0, 1, S['mas'], 3)
    a, _ = rebuild(sh, r0, S['mas'], lay)
    pos = {}
    for i, (name, kind, gi) in enumerate(kinds):
        pos.setdefault((name, kind), []).append(a + i)
    price = {'대행': res['price'].get('분양대행') or 0, '조직': res['price'].get('조직') or 0, 'MGM': res['price'].get('MGM') or 0}
    dist = {k: _distribute(series[k], keys, units) for k in series}
    cols = ['F', 'G', 'H'] + [L(10 + j) for j in range(12)]
    out = []
    for i, (name, kind, gi) in enumerate(kinds):
        rr = a + i
        row = {c_: '' for c_ in ['C', 'D', 'E', 'I'] + cols}
        if kind == 'u':
            row['C'] = int(keys[gi]) if keys[gi].isdigit() else keys[gi]
            row['D'] = price[name]
            for j in range(12):
                row[L(10 + j)] = dist[name][keys[gi]][j] or ''
        elif kind == 'a':
            u = pos[(name, 'u')][gi]
            row['C'], row['D'] = f'=+C{u}', f'=+D{u}'
            for col in cols:
                row[col] = f'=+{col}{u}*$D${rr}'
        elif kind in ('usum', 'asum'):
            src = pos[(name, 'u' if kind == 'usum' else 'a')]
            for col in cols:
                row[col] = f'=+SUM({col}{src[0]}:{col}{src[-1]})'
        elif kind == 'total':
            s_, c_r = pos[(name, 'asum')][0], pos[(name, 'consult')][0]
            for col in cols:
                row[col] = f'=+{col}{s_}+{col}{c_r}'
        elif kind == 'grand':
            t1, t2 = pos[('대행', 'total')][0], pos[('조직', 'total')][0]
            for col in cols:
                row[col] = f'=+{col}{t1}+{col}{t2}'
        row['E'] = f'=+SUM(F{rr}:H{rr})'
        row['I'] = f'=+SUM(J{rr}:U{rr})'
        out.append([row[c_] for c_ in ['C', 'D', 'E', 'F', 'G', 'H', 'I'] + [L(10 + j) for j in range(12)]])
    write_block(sh, a, out, 'C', 'U')
    labels = {'대행': '분양대행\n수수료', '조직': '분양대행\n수수료\n(조직)', 'MGM': '분양촉진\n수수료\n(MGM 등)'}
    for name in ('대행', '조직', 'MGM'):
        if (name, 'u') not in pos:
            continue
        u, am = pos[(name, 'u')], pos[(name, 'a')]
        last = pos[(name, 'total')][0] if (name, 'total') in pos else pos[(name, 'asum')][0]
        merge_range(sh, u[0], 1, last, 1)
        put(sh, u[0], 1, labels[name])
        merge_range(sh, u[0], 2, u[-1], 2)
        put(sh, u[0], 2, '세대수')
        merge_range(sh, am[0], 2, am[-1], 2)
        put(sh, am[0], 2, '금 액')
        for k in ('usum', 'asum'):
            put(sh, pos[(name, k)][0], 2, '')
        if (name, 'consult') in pos:
            cr, tr = pos[(name, 'consult')][0], pos[(name, 'total')][0]
            merge_range(sh, cr, 2, cr, 3)
            put(sh, cr, 2, '상담사 인건비')
            merge_range(sh, tr, 2, tr, 3)
            put(sh, tr, 2, '계')
    if ('총계', 'grand') in pos:
        gr = pos[('총계', 'grand')][0]
        merge_range(sh, gr, 1, gr, 3)
        put(sh, gr, 1, '총 계')
    biggest = max([sum(series[k]) * price[k] for k in series] + [0])
    widen(sh, [L(i) for i in range(5, 22)], biggest)
    return True


# ── 공가세대관리비 ───────────────────────────────────────────────
def fill_vacancy(book, sh, res, types):
    c = res['calc']
    vac = c.get('미입주', [0] * 12)
    if not any(vac):
        sheet_state(book, sh.name, 'hidden')
        return False
    tl = [t for t in types if t and t[1] not in (None, '')]
    keys = [str(t[0]) for t in tl]
    units = {str(t[0]): float(t[1]) for t in tl}
    py = {str(t[0]): (math.ceil(float(t[2]) * 0.3025) if t[2] not in (None, '') else 0) for t in tl}
    v = vals(sh, 'V', 40)
    r7 = _find(v, 0, lambda t: t.startswith('평당'))
    r0 = r7 + 1 if r7 else None
    if not r0 or _txt(v[r0 - 1][1]) != '세대수' or _txt(v[r0 + 3][1]) != '금액':
        return False
    S = dict(u1=r0, um=r0 + 1, ul=r0 + 2, us=r0 + 3, a1=r0 + 4, am=r0 + 5, al=r0 + 6, as_=r0 + 7)
    n = max(1, len(keys))
    pick = lambda f, m_, l_: [f] if n == 1 else [f] + [m_] * (n - 2) + [l_]
    lay = pick(S['u1'], S['um'], S['ul']) + [S['us']] + pick(S['a1'], S['am'], S['al']) + [S['as_']]
    unmerge_range(sh, r0, 1, S['as_'], 3)
    a, b = rebuild(sh, r0, S['as_'], lay)
    dist = _distribute(vac, keys, units) if keys else {}
    per = res['price'].get('공가') or 0
    months = [L(9 + j) for j in range(12)]
    rows = []
    us_row, as_row = a + n, a + 2 * n + 1
    for i in range(n):
        rr = a + i
        k = keys[i] if i < len(keys) else ''
        rows.append([k, f'=+SUM(E{rr}:G{rr})', '', '', '', f'=+SUM(I{rr}:T{rr})'] +
                    [(dist[k][j] or '') if k else '' for j in range(12)] + [f'{py.get(k, 0)}평' if k else ''])
    rows.append(['', f'=+SUM(E{us_row}:G{us_row})'] + [f'=+SUM({col}{a}:{col}{us_row - 1})' for col in 'EFG'] +
                [f'=+SUM(I{us_row}:T{us_row})'] + [f'=+SUM({m}{a}:{m}{us_row - 1})' for m in months] + [''])
    for i in range(n):
        rr = as_row - n + i
        u = a + i
        k = keys[i] if i < len(keys) else ''
        rows.append([f'=+C{u}', f'=+SUM(E{rr}:G{rr})', '', '', '', f'=+SUM(I{rr}:T{rr})'] +
                    [f'=+{m}$7*{m}{u}*{py.get(k, 0)}' for m in months] + [''])
    rows.append(['', f'=+SUM(E{as_row}:G{as_row})'] + [f'=+SUM({col}{as_row - n}:{col}{as_row - 1})' for col in 'EFG'] +
                [f'=+SUM(I{as_row}:T{as_row})'] + [f'=+SUM({m}{as_row - n}:{m}{as_row - 1})' for m in months] + [''])
    write_block(sh, a, rows, 'C', 'U')
    write_block(sh, r7, [[per if vac[j] else '' for j in range(12)]], 'I', 'T')
    merge_range(sh, a, 1, as_row, 1)
    put(sh, a, 1, '공가세대\n관리비')
    merge_range(sh, a, 2, us_row - 1, 2)
    put(sh, a, 2, '세대수')
    merge_range(sh, us_row + 1, 2, as_row - 1, 2)
    put(sh, us_row + 1, 2, '금 액')
    biggest = max(vac) * per * max(list(py.values()) + [0])
    widen(sh, [L(i) for i in range(4, 21)], biggest)
    return dict(total=as_row)


# ── 대납이자 ─────────────────────────────────────────────────────
def fill_interest(book, sh, res, fields, ty):
    v = vals(sh, 'W', 45)
    r1 = _find(v, 0, lambda t: t == '중도금대출')
    r2 = _find(v, 0, lambda t: t == '대출이자')
    if r1 and r2:
        clear_range(sh, r1, 5, r2 - 2, 9)
        clear_range(sh, r1, 11, r2 - 2, 22)
    rr = _find(v, 0, lambda t: t == '구분', (r2 or 30))
    if rr:
        rates = res.get('interest') or [None] * 4
        put(sh, rr + 1, 1, ty - 1)
        put(sh, rr + 2, 1, ty)
        for (r_, c_), val in zip(((rr + 1, 2), (rr + 1, 3), (rr + 2, 2), (rr + 2, 3)), rates):
            if val is not None and val != '':
                put(sh, r_, c_, val)
    if fields.get('무이자') != 'Y':
        sheet_state(book, sh.name, 'hidden')


def link_interest(bs, di, binfo):
    v = vals(di, 'W', 45)
    r2 = _find(v, 0, lambda t: t == '대출이자')
    rs = _find(v, 1, lambda t: t == '계', r2 or 0) if r2 else None
    hdr = _find(v, 10, lambda t: t == '1월')
    if not (rs and hdr):
        return
    c0 = next(j for j, x in enumerate(v[hdr - 1]) if _txt(x) == '1월') + 1
    dq = X._qname(di.name)
    for g, s0, s1 in binfo['groups']:
        if g['sub'] == '대납이자':
            write_block(bs, s0, [[f"={dq}!{L(c0 + j)}{rs}*1000000" for j in range(12)]], 'J', 'U')


# ── 입력값 보관(아주 숨김 시트) ─────────────────────────────────────
def store_inputs(book, data, year):
    s = json.dumps(dict(year=year, data=data), ensure_ascii=False)
    chunks = [s[i:i + 30000] for i in range(0, len(s), 30000)] or ['']
    rows = [['판관비 예산안 자동작성 입력값 — 화면에서 다시 불러올 때 씁니다(지우지 마세요).'], [len(chunks)]] + [[ch] for ch in chunks]
    add_sheet(book, DATA_SHEET, rows, hidden=True)
    sheet_state(book, DATA_SHEET, 'veryHidden')


def read_inputs(xlsx_bytes):
    """화면에서 만든 보고서 파일(아주 숨긴 '_자동작성입력' 시트) → dict(year, data). 없으면 None"""
    book = X.Book(xlsx_bytes)
    if DATA_SHEET not in [s['name'] for s in book.sheet_list()]:
        return None
    sh = book.sheet(DATA_SHEET)
    n = int(sh.value('A2') or 0)
    js = ''.join(str(sh.value(f'A{3 + i}') or '') for i in range(n))
    return json.loads(js)


# ── 전체 ─────────────────────────────────────────────────────────
def build_report(form_bytes, res, data):
    book = X.Book(form_bytes)
    kit = StyleKit(book)
    fields = data.get('fields') or {}
    types = [t for t in (data.get('types') or []) if t and (t[0] not in (None, '') or t[1] not in (None, ''))]
    ty = int(res['year'])
    names = [s['name'] for s in book.sheet_list()]
    bname = next(n for n in names if re.fullmatch(r'\d{2}년 예산', n))
    fy = 2000 + int(bname[:2])
    shift_years(book, fy, ty)
    new_name = f'{ty % 100:02d}년 예산'
    if bname != new_name:
        book.rename_sheet(bname, new_name)
    bs = book.sheet(new_name)
    site = fields.get('현장명') or '현장명'
    d = fields.get('작성일')
    try:
        dt = datetime.datetime.strptime(str(d)[:10], '%Y-%m-%d') if d else datetime.datetime.now()
    except ValueError:
        dt = datetime.datetime.now()
    date_txt = f'( {dt:%Y. %m. %d} )'
    title = f'[ {site} ] {ty}년 판관비 예산안'
    open_idx = None
    o = fields.get('오픈') or [None, None]
    if o[0] and o[1]:
        open_idx = int(o[0]) * 12 + int(o[1])
    note = ' 신규현장' if open_idx and ty * 12 + 1 <= open_idx <= ty * 12 + 12 else ''
    binfo = fill_budget(bs, kit, res, title, date_txt, site, keep_interest=fields.get('무이자') == 'Y')
    fill_report(book.sheet('예산안 보고서'), new_name, binfo, title, date_txt, site, note)
    basic = fill_basic(book.sheet('기본정보'), kit, fields, types)
    fill_rates(book.sheet('분양율,입주율'), res, basic['total'])
    fill_agency(book, book.sheet('분양대행 및 촉진수수료'), res, types)
    fill_vacancy(book, book.sheet('공가세대관리비'), res, types)
    fill_interest(book, book.sheet('대납이자'), res, fields, ty)
    if fields.get('무이자') == 'Y':
        link_interest(bs, book.sheet('대납이자'), binfo)
    for nm in ('프로젝트(수주)', '수주실패비'):
        if nm in names:
            sheet_state(book, nm, 'hidden')
    store_inputs(book, data, ty)
    activate(book, '예산안 보고서')
    return book.save(), dict(budget_sheet=new_name, budget_total=binfo['total'], groups=binfo['groups'])
