"""판관비 예산안 자동작성 — 사내망용 작은 웹서버.

  python server.py            # 사내망 공개 (접속키 필요)
  python server.py --local    # 이 PC 에서만

이 PC 에서는 http://localhost:8792/ 로, 다른 PC 에서는 http://<이 PC IP>:8792/?k=<접속키> 로 들어온다.
접속키는 app\\_key.txt. 계산은 이 PC 의 Excel 로 하므로 한 번에 하나씩 처리한다(전용 작업 줄).
"""
import json
import logging
import os
import queue
import random
import re
import shutil
import socket
import sys
import threading
import time
import traceback
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

APP = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(APP)                       # ☆ 예산안 자동작성
TOOL = os.path.join(ROOT, 'tool')
BUDGET_ROOT = os.path.dirname(ROOT)               # ★ 판관비 예산안
sys.path.insert(0, TOOL)

PORT = 8792
JOBS = os.path.join(APP, '_jobs')
KEYFILE = os.path.join(APP, '_key.txt')
URLFILE = os.path.join(APP, '_url.txt')
LOGFILE = os.path.join(APP, '_server.log')
SITES = os.path.join(APP, '_sites.json')
FORMS = os.path.join(ROOT, '회사양식')             # 경영기획팀 '사업비 예산편성 양식' (결과물이 이 양식 그대로 나감)
DOCS = os.path.join(ROOT, 'docs')                  # 지출계획서 화면(docs/jichul/index.html)은 웹판과 같은 파일
JICHUL_SITES = os.path.join(APP, '_jichul_sites.json')
JICHUL_ROOT = os.path.join(BUDGET_ROOT, '★ 지출계획서')   # PC 판 기본 저장 위치: ★ 지출계획서\\<현장명>\\
MAX_UPLOAD = 40 * 1024 * 1024
KEEP_HOURS = 24
IDLE_QUIT = 60            # 마지막 계산 뒤 이 초가 지나면 Excel 을 끈다

os.makedirs(JOBS, exist_ok=True)
logging.basicConfig(filename=LOGFILE, level=logging.INFO, encoding='utf-8',
                    format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger('server')

import openpyxl     # noqa: E402
import apply_kijun  # noqa: E402
import report       # noqa: E402
import site_io      # noqa: E402
import design_overview  # noqa: E402
import xlcalc       # noqa: E402
import jichul       # noqa: E402
import jichul_xlsx  # noqa: E402

_MODELS = {}
_MODEL_LOCK = threading.Lock()


def model_for(tpl):
    """계산용 양식의 수식 계산기(xlcalc) — 보고서 월별 칸에 넣는 산식(계산 과정)을 만들 때 쓴다. 금액 계산은 Excel 결과 그대로."""
    key = (tpl, os.path.getmtime(tpl))
    m = _MODELS.get(key)
    if m is None:
        _MODELS.clear()
        m = _MODELS[key] = xlcalc.Model(tpl)
    return m

os.makedirs(FORMS, exist_ok=True)


def form_info():
    p = report.latest_form(FORMS)
    if not p:
        return None
    year = None
    try:
        wb = openpyxl.load_workbook(p, read_only=True)
        m = re.search(r'(\d{2})년 예산', ' '.join(wb.sheetnames))
        wb.close()
        year = 2000 + int(m.group(1)) if m else None
    except Exception:
        pass
    return {'name': os.path.basename(p), 'path': p, 'year': year}


def ensure_report(jid):
    """작업(계산 결과)으로 회사 양식 보고서를 만든다(이미 있으면 그대로)"""
    d = job_dir(jid)
    out = os.path.join(d, 'report.xlsx')
    meta = rjson(os.path.join(d, 'report.json'))
    if meta and os.path.exists(out):
        return meta
    fi = form_info()
    if not fi:
        raise UserError('회사 양식 파일이 없습니다. "회사 양식 바꾸기"로 「사업비 예산편성 양식」을 넣어 주세요.')
    res = rjson(os.path.join(d, 'result.json'))
    inp = rjson(os.path.join(d, 'input.json'), {})
    if not res:
        raise UserError('계산 결과가 없습니다. 다시 계산해 주세요.')
    t = time.time()
    report.ERRORS.clear()
    r = WORKER.call(report.build_report, fi['path'], out, res, inp.get('data') or {}, os.path.join(d, 'pv'))
    for e in report.ERRORS:
        log.warning('report %s 미리보기: %s', jid, e)
    meta = {'previews': r['previews'], 'form': fi['name'], 'sec': round(time.time() - t, 1)}
    wjson(os.path.join(d, 'report.json'), meta)
    log.info('report %s %.1fs', jid, time.time() - t)
    return meta


class UserError(Exception):
    pass


def get_key():
    if os.path.exists(KEYFILE):
        k = open(KEYFILE, encoding='utf-8').read().strip()
        if k:
            return k
    k = ''.join(random.choice('abcdefghijkmnpqrstuvwxyz23456789') for _ in range(8))
    open(KEYFILE, 'w', encoding='utf-8').write(k)
    return k


KEY = get_key()


def lan_ip():
    ips = []
    try:
        ips = socket.gethostbyname_ex(socket.gethostname())[2]
    except Exception:
        pass
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('10.255.255.255', 1))
        ips.insert(0, s.getsockname()[0])
        s.close()
    except Exception:
        pass
    for ip in ips:
        if re.match(r'^(10\.|192\.168\.|172\.(1[6-9]|2\d|3[01])\.)', ip):
            return ip
    return ips[0] if ips else '127.0.0.1'


def rjson(path, default=None):
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return default


def wjson(path, obj):
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False, default=str, indent=1)
    os.replace(tmp, path)


def cleanup_jobs():
    now = time.time()
    for d in os.listdir(JOBS):
        p = os.path.join(JOBS, d)
        try:
            if os.path.isdir(p) and now - os.path.getmtime(p) > KEEP_HOURS * 3600:
                shutil.rmtree(p, ignore_errors=True)
        except Exception:
            pass


def new_job():
    cleanup_jobs()
    jid = os.urandom(6).hex()
    os.makedirs(os.path.join(JOBS, jid))
    return jid


def job_dir(job):
    if not re.fullmatch(r'[0-9a-f]{12}', job or ''):
        raise UserError('작업 번호가 올바르지 않습니다. 다시 계산해 주세요.')
    d = os.path.join(JOBS, job)
    if not os.path.isdir(d):
        raise UserError('작업이 만료되었습니다. 다시 계산해 주세요.')
    return d


def safe_name(name, default='현장'):
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', '_', str(name or '')).strip().strip('.')
    return name or default


def template_for(year):
    ts = apply_kijun.templates()
    if not ts:
        raise UserError('양식 파일(판관비 예산안 자동작성 양식_연도.xlsx)이 없습니다.')
    if year:
        for y, p in ts:
            if str(y) == str(year):
                return y, p
    return ts[0]


def parse_multipart(ctype, raw):
    m = re.search(r'boundary="?([^";]+)"?', ctype)
    if not m:
        raise UserError('업로드 형식이 올바르지 않습니다.')
    sep = b'--' + m.group(1).encode('latin-1')
    files, form = {}, {}
    for part in raw.split(sep):
        if not part or part.startswith(b'--'):
            continue
        part = part[2:] if part.startswith(b'\r\n') else part
        head, _, data = part.partition(b'\r\n\r\n')
        if data.endswith(b'\r\n'):
            data = data[:-2]
        hd = head.decode('utf-8', 'replace')
        nm = re.search(r'\bname="([^"]*)"', hd)
        fn = re.search(r"filename\*=UTF-8''([^;\r\n]+)", hd, re.I)
        fname = urllib.parse.unquote(fn.group(1)) if fn else None
        if fname is None:
            fn = re.search(r'filename="([^"]*)"', hd)
            fname = fn.group(1) if fn else None
        if not nm:
            continue
        if fname is not None:
            files[nm.group(1)] = (fname.replace('\\', '/').split('/')[-1], data)
        else:
            form[nm.group(1)] = data.decode('utf-8', 'replace')
    return files, form


# ------------------------------------------------------------------ Excel 전용 작업 줄

class ExcelWorker:
    """Excel COM 은 한 줄(스레드)에서만. 계산이 이어질 때 빠르게 하려고 잠깐 켜 둔다."""

    def __init__(self):
        self.q = queue.Queue()
        threading.Thread(target=self.run, daemon=True).start()

    def stop(self, timeout=20):
        """서버를 끌 때 Excel 도 끈다(숨은 Excel 이 남지 않게)"""
        box = {'ev': threading.Event()}
        self.q.put((None, (), box))
        box['ev'].wait(timeout)

    def call(self, fn, *args, timeout=180):
        box = {'ev': threading.Event()}
        self.q.put((fn, args, box))
        if not box['ev'].wait(timeout):
            raise UserError('Excel 계산이 너무 오래 걸립니다. 잠시 뒤 다시 해 주세요.')
        if 'error' in box:
            raise box['error']
        return box['result']

    def run(self):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        xl = None
        last = time.time()

        def release(app):
            """끄기. 사용자가 연 통합문서가 이 Excel 에 들어와 있으면 끄지 않고 화면에 넘겨 준다."""
            try:
                if app.Workbooks.Count > 0:
                    app.Visible = True
                    app.UserControl = True
                    log.warning('사용자 통합문서가 숨은 Excel 에 열려 있어 화면에 넘겨 줌')
                    return
                app.Quit()
            except Exception:
                pass

        while True:
            try:
                fn, args, box = self.q.get(timeout=5)
            except queue.Empty:
                if xl is not None and time.time() - last > IDLE_QUIT:
                    release(xl)
                    xl = None
                continue
            if fn is None:                      # 멈춤 신호
                if xl is not None:
                    release(xl)
                xl = None
                box['ev'].set()
                return
            try:
                if xl is not None:
                    try:
                        if xl.Workbooks.Count > 0:      # 남의 통합문서가 들어와 있으면 그 Excel 은 넘겨 주고 새로
                            release(xl)
                            xl = None
                    except Exception:
                        xl = None
                if xl is None:
                    t0 = time.time()
                    xl = win32com.client.DispatchEx('Excel.Application')
                    log.info('Excel 시작 %.1f초', time.time() - t0)
                    xl.Visible = False
                    xl.DisplayAlerts = False
                    xl.ScreenUpdating = False
                    xl.AskToUpdateLinks = False
                box['result'] = fn(xl, *args)
            except Exception as e:
                box['error'] = e
                log.error('excel\n%s', traceback.format_exc())
                try:
                    release(xl)
                except Exception:
                    pass
                xl = None
            finally:
                last = time.time()
                box['ev'].set()


WORKER = None


# ------------------------------------------------------------------ 저장 기록

def fix_drive(p):
    """외장하드가 PC 마다 다른 글자(E:/F:)로 잡혀도 저장한 현장 경로가 열리게, 없는 경로면 드라이브 글자를 이 폴더 기준으로 바꿔 본다."""
    if p and not os.path.exists(p):
        drive = os.path.splitdrive(ROOT)[0]
        old, rest = os.path.splitdrive(p)
        if drive and old and old != drive and os.path.exists(drive + rest):
            return drive + rest
    return p


def sites_list():
    items = rjson(SITES, []) or []
    out = []
    seen = set()
    for it in sorted(items, key=lambda x: x.get('saved', ''), reverse=True):
        p = fix_drive(it.get('path'))
        if p and p not in seen and os.path.exists(p):
            seen.add(p)
            out.append(dict(it, path=p))
    return out[:200]


def remember_site(name, path, year, total):
    items = [x for x in (rjson(SITES, []) or []) if x.get('path') != path]
    items.append(dict(name=name, path=path, year=year, total=total, saved=time.strftime('%Y-%m-%d %H:%M')))
    wjson(SITES, items[-500:])


def default_sub(fields, year, kinds=None):
    kinds = kinds or []
    for k, sub in (('new', '신규분양'), ('movein', '입주'), ('remain', '기공급')):
        if k in kinds:
            return sub

    def idx(v):
        try:
            return int(v[0]) * 12 + int(v[1])
        except Exception:
            return None
    o, i = idx(fields.get('오픈') or [None, None]), idx(fields.get('입주') or [None, None])
    y0, y1 = year * 12 + 1, year * 12 + 12
    if o is not None and o >= y0 - 1:
        return '신규분양'
    if i is not None and i <= y1:
        return '입주'
    return '기공급'


_KIJUN = {}


def kijun_for(tpl):
    """양식의 '기준단가' 시트 → ({이름: 값}, template_info) — 지출계획서 계산용(예산안과 같은 단가)"""
    key = (tpl, os.path.getmtime(tpl))
    if key not in _KIJUN:
        _KIJUN.clear()
        t = site_io.template_info(tpl)
        _KIJUN[key] = ({p['key']: p['value'] for p in t['params'] if 'key' in p}, t)
    return _KIJUN[key]


def jichul_sites_list():
    items = rjson(JICHUL_SITES, []) or []
    out, seen = [], set()
    for it in sorted(items, key=lambda x: x.get('saved', ''), reverse=True):
        p = fix_drive(it.get('path'))
        if p and p not in seen and os.path.exists(p):
            seen.add(p)
            out.append(dict(it, path=p))
    return out[:200]


def remember_jichul_site(name, path, year, total):
    items = [x for x in (rjson(JICHUL_SITES, []) or []) if x.get('path') != path]
    items.append(dict(name=name, path=path, year=year, total=total, saved=time.strftime('%Y-%m-%d %H:%M')))
    wjson(JICHUL_SITES, items[-500:])


# ------------------------------------------------------------------ HTTP

class Handler(BaseHTTPRequestHandler):
    server_version = 'PangwanbiBudget/1.0'

    def log_message(self, fmt, *args):
        log.info('%s %s', self.client_address[0], fmt % args)

    def is_local(self):
        return self.client_address[0] in ('127.0.0.1', '::1')

    def authorized(self):
        if self.is_local():
            return True
        q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        if q.get('k', [''])[0] == KEY or self.headers.get('X-Key') == KEY:
            return True
        return f'k={KEY}' in self.headers.get('Cookie', '')

    def send(self, code, body, ctype='application/json; charset=utf-8', extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False, default=str).encode('utf-8')
        elif isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def fail(self, msg, code=400):
        self.send(code, {'ok': False, 'error': msg})

    def body(self):
        n = int(self.headers.get('Content-Length') or 0)
        if n > MAX_UPLOAD:
            raise UserError('파일이 너무 큽니다 (40MB 이하).')
        return self.rfile.read(n) if n else b''

    def jbody(self):
        return json.loads(self.body() or b'{}')

    def q(self):
        return {k: v[0] for k, v in urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query).items()}

    def need_local(self):
        if not self.is_local():
            raise UserError('이 기능은 서버 PC 에서만 쓸 수 있습니다.')

    # ---------------------------------------------------------------- GET
    def do_GET(self):
        url = urllib.parse.urlparse(self.path)
        if not self.authorized():
            return self.send(403, '<meta charset="utf-8"><p style="font:16px sans-serif;padding:40px">'
                                  '접속키가 필요합니다. 담당자에게 받은 주소(…?k=…)로 들어와 주세요.</p>',
                             'text/html; charset=utf-8')
        try:
            q = self.q()
            if url.path in ('/', '/index.html'):
                extra = {}
                if q.get('k') == KEY:
                    extra['Set-Cookie'] = f'k={KEY}; Path=/; Max-Age=31536000; SameSite=Lax'
                with open(os.path.join(APP, 'static', 'index.html'), 'rb') as f:
                    return self.send(200, f.read(), 'text/html; charset=utf-8', extra)
            if url.path in ('/jichul', '/jichul/', '/jichul/index.html'):
                extra = {}
                if q.get('k') == KEY:
                    extra['Set-Cookie'] = f'k={KEY}; Path=/; Max-Age=31536000; SameSite=Lax'
                with open(os.path.join(DOCS, 'jichul', 'index.html'), 'rb') as f:
                    return self.send(200, f.read(), 'text/html; charset=utf-8', extra)
            if url.path.startswith('/api/jichul/'):
                return self.jichul_get(url.path[len('/api/jichul/'):], q)
            if url.path == '/api/ping':
                return self.send(200, {'ok': True, 'app': 'pangwanbi-budget', 'local': self.is_local()})
            if url.path == '/api/info':
                ts = apply_kijun.templates()
                info = {'ok': True, 'local': self.is_local(), 'budget_root': BUDGET_ROOT,
                        'templates': [{'year': y, 'name': os.path.basename(p)} for y, p in ts]}
                if ts:
                    y, p = template_for(q.get('year'))
                    t = site_io.template_info(p)
                    t.pop('params', None)
                    t.pop('path', None)
                    info['template'] = t
                fi = form_info()
                info['form'] = {'name': fi['name'], 'year': fi['year']} if fi else None
                if self.is_local():
                    info['share_url'] = f'http://{lan_ip()}:{PORT}/?k={KEY}'
                return self.send(200, info)
            if url.path == '/api/preview':
                d = job_dir(q.get('job'))
                fn = os.path.basename(q.get('f') or '')
                p = os.path.join(d, 'pv', fn)
                if not fn.endswith('.png') or not os.path.exists(p):
                    return self.fail('미리보기가 없습니다.', 404)
                with open(p, 'rb') as f:
                    return self.send(200, f.read(), 'image/png')
            if url.path == '/api/schema':
                y, p = template_for(q.get('year'))
                return self.send(200, {'ok': True, 'year': y, **site_io.schema(p)})
            if url.path == '/api/params':
                y, p = template_for(q.get('year'))
                t = site_io.template_info(p)
                t.pop('path', None)
                return self.send(200, {'ok': True, **t})
            if url.path == '/api/warm':          # 화면을 열 때 Excel 을 미리 깨워 첫 계산을 빠르게
                threading.Thread(target=lambda: WORKER.call(lambda xl: xl.Version), daemon=True).start()
                return self.send(200, {'ok': True})
            if url.path == '/api/recent':
                return self.send(200, {'ok': True, 'sites': sites_list() if self.is_local() else []})
            if url.path == '/api/download':
                d = job_dir(q.get('job'))
                if q.get('kind') == 'calc':          # 계산 근거(자동작성 양식) 파일
                    fn = os.path.join(d, 'work.xlsx')
                else:                                # 회사 양식 보고서
                    ensure_report(q.get('job'))
                    fn = os.path.join(d, 'report.xlsx')
                if not os.path.exists(fn):
                    return self.fail('아직 계산한 파일이 없습니다.', 404)
                name = safe_name(q.get('name') or '판관비 예산안.xlsx')
                if not name.lower().endswith('.xlsx'):
                    name += '.xlsx'
                with open(fn, 'rb') as f:
                    data = f.read()
                return self.send(200, data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                 {'Content-Disposition': "attachment; filename*=UTF-8''" + urllib.parse.quote(name)})
            return self.fail('없는 주소입니다.', 404)
        except UserError as e:
            return self.fail(str(e), 404)
        except Exception as e:
            log.error('GET %s\n%s', self.path, traceback.format_exc())
            return self.fail(f'오류: {e}', 500)

    # ---------------------------------------------------------------- POST
    def do_POST(self):
        url = urllib.parse.urlparse(self.path)
        if not self.authorized():
            return self.fail('접속키가 필요합니다.', 403)
        try:
            route = {'/api/calc': self.api_calc, '/api/save': self.api_save, '/api/load': self.api_load,
                     '/api/load_path': self.api_load_path, '/api/kijun': self.api_kijun, '/api/open': self.api_open,
                     '/api/report': self.api_report, '/api/form': self.api_form, '/api/types': self.api_types,
                     '/api/jichul/calc': self.jichul_calc, '/api/jichul/xlsx': self.jichul_xlsx, '/api/jichul/save': self.jichul_save,
                     '/api/jichul/load': self.jichul_load, '/api/jichul/load_path': self.jichul_load_path,
                     '/api/jichul/open': self.jichul_open}
            if url.path in route:
                return route[url.path]()
            if url.path == '/api/shutdown' and self.is_local():
                WORKER.stop()
                self.send(200, {'ok': True})
                threading.Thread(target=self.server.shutdown, daemon=True).start()
                return
            return self.fail('없는 주소입니다.', 404)
        except UserError as e:
            return self.fail(str(e))
        except apply_kijun.ApplyError as e:
            return self.fail(str(e))
        except Exception as e:
            log.error('POST %s\n%s', self.path, traceback.format_exc())
            return self.fail(f'처리 중 오류가 났습니다: {e}', 500)

    # ---------------------------------------------------------------- 지출계획서
    def jichul_get(self, cmd, q):
        y, tpl = template_for(q.get('year'))
        K, t = kijun_for(tpl)
        if cmd == 'info':
            ts = apply_kijun.templates()
            info = {'ok': True, 'local': self.is_local(), 'year': y, 'kijun': t.get('kijun'), 'written': t.get('written'),
                    'template': os.path.basename(tpl), 'count': len(K), 'jichul_root': JICHUL_ROOT,
                    'templates': [{'year': yy, 'name': os.path.basename(p)} for yy, p in ts]}
            if self.is_local():
                info['share_url'] = f'http://{lan_ip()}:{PORT}/jichul/?k={KEY}'
            return self.send(200, info)
        if cmd == 'schema':
            return self.send(200, {'ok': True, 'year': y, **jichul.schema(K)})
        if cmd == 'params':
            return self.send(200, {'ok': True, 'year': y, 'kijun': t.get('kijun'), 'written': t.get('written'), 'params': t['params']})
        if cmd == 'recent':
            return self.send(200, {'ok': True, 'sites': jichul_sites_list() if self.is_local() else []})
        return self.fail('없는 주소입니다.', 404)

    def _jichul_res(self, req):
        y, tpl = template_for(req.get('year'))
        K, _ = kijun_for(tpl)
        data = req.get('data') or {}
        res = jichul.calc(data, K)
        res['file_name'] = jichul.file_name(res)
        res['year'] = y
        return y, data, res

    def jichul_calc(self):
        y, data, res = self._jichul_res(self.jbody())
        log.info('jichul calc %s %s 합계 %s', res.get('site'), res.get('kind'), res.get('total_all'))
        self.send(200, {'ok': True, **res})

    def jichul_xlsx(self):
        req = self.jbody()
        y, data, res = self._jichul_res(req)
        out = jichul_xlsx.build_xlsx(res, data, y)
        name = safe_name(req.get('name') or res['file_name'])
        if not name.lower().endswith('.xlsx'):
            name += '.xlsx'
        self.send(200, out, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                  {'Content-Disposition': "attachment; filename*=UTF-8''" + urllib.parse.quote(name)})

    def jichul_save(self):
        self.need_local()
        req = self.jbody()
        y, data, res = self._jichul_res(req)
        site = safe_name(res.get('site'))
        folder = (req.get('folder') or '').strip() or os.path.join(JICHUL_ROOT, site)
        fname = safe_name(req.get('filename') or res['file_name'])
        if not fname.lower().endswith('.xlsx'):
            fname += '.xlsx'
        path = os.path.join(folder, fname)
        if os.path.exists(path) and not req.get('overwrite'):
            return self.send(200, {'ok': False, 'exists': True, 'path': path})
        if os.path.exists(os.path.join(folder, '~$' + fname)):
            raise UserError('같은 이름의 파일이 Excel 에 열려 있습니다. 닫고 다시 저장하세요.')
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError as e:
            raise UserError(f'폴더를 만들 수 없습니다: {folder} ({e})')
        with open(path, 'wb') as f:
            f.write(jichul_xlsx.build_xlsx(res, data, y))
        remember_jichul_site(f"{site} {res.get('kind')} {jichul.period_text(res)}", path, y, res.get('total_all'))
        log.info('jichul save %s', path)
        self.send(200, {'ok': True, 'path': path, 'folder': folder})

    def _jichul_loaded(self, path, name):
        obj = jichul_xlsx.read_inputs(path)
        if not obj:
            raise UserError('지출계획서 자동작성으로 만든 파일이 아닙니다(입력값이 들어 있지 않음).')
        return self.send(200, {'ok': True, 'data': obj.get('data') or {}, 'year': obj.get('year'), 'name': name})

    def jichul_load(self):
        files, form = parse_multipart(self.headers.get('Content-Type', ''), self.body())
        if 'file' not in files:
            raise UserError('엑셀 파일을 넣어 주세요.')
        name, data = files['file']
        if not name.lower().endswith('.xlsx'):
            raise UserError('엑셀(.xlsx) 파일이어야 합니다.')
        jid = new_job()
        p = os.path.join(JOBS, jid, 'load.xlsx')
        with open(p, 'wb') as f:
            f.write(data)
        return self._jichul_loaded(p, name)

    def jichul_load_path(self):
        self.need_local()
        p = fix_drive(self.jbody().get('path') or '')
        if not os.path.exists(p):
            raise UserError('파일이 없습니다: ' + p)
        return self._jichul_loaded(p, os.path.basename(p))

    def jichul_open(self):
        self.need_local()
        p = self.jbody().get('path') or ''
        known = {x['path'] for x in jichul_sites_list()} | {os.path.dirname(x['path']) for x in jichul_sites_list()}
        if not os.path.exists(p) or not (p.startswith(BUDGET_ROOT) or p in known):
            raise UserError('열 수 없는 경로입니다.')
        os.startfile(p)
        self.send(200, {'ok': True})

    def api_calc(self):
        req = self.jbody()
        data = req.get('data') or {}
        y, tpl = template_for(req.get('year'))
        jid = new_job()
        work = os.path.join(JOBS, jid, 'work.xlsx')
        t = time.time()
        site_io.write_inputs(tpl, work, data)
        wjson(os.path.join(JOBS, jid, 'input.json'), {'year': y, 'data': data})
        res = WORKER.call(site_io.excel_calc, work)
        try:                                           # 월별 산식(보고서 칸에 '=35*2000*1000' 처럼) — 실패해도 값은 그대로
            with _MODEL_LOCK:
                site_io.attach_formulas(model_for(tpl), data, res)
        except Exception:
            log.warning('산식 만들기 실패(값만 씀)\n%s', traceback.format_exc())
        wjson(os.path.join(JOBS, jid, 'result.json'), res)
        res['sec'] = round(time.time() - t, 1)
        res['job'] = jid
        res['template_year'] = y
        res['default_sub'] = default_sub(data.get('fields') or {}, y, data.get('kinds'))
        log.info('calc %s %s %.1fs total=%s', jid, (data.get('fields') or {}).get('현장명'), time.time() - t,
                 res['total']['total'])
        self.send(200, {'ok': True, **res})

    def api_save(self):
        self.need_local()
        req = self.jbody()
        d = job_dir(req.get('job'))
        ensure_report(req.get('job'))
        src = os.path.join(d, 'report.xlsx')          # 회사 양식 그대로인 보고서를 저장
        inp = rjson(os.path.join(d, 'input.json'), {})
        year = int(inp.get('year') or 0) or template_for(None)[0]
        site = safe_name(req.get('site') or ((inp.get('data') or {}).get('fields') or {}).get('현장명'))
        sub = safe_name(req.get('sub') or '', '')
        fname = safe_name(req.get('filename') or f'({site}) {year}년 사업비 예산편성.xlsx')
        if not fname.lower().endswith('.xlsx'):
            fname += '.xlsx'
        folder = os.path.join(BUDGET_ROOT, f'★ {year} 예산안', *(x for x in (sub, site) if x))
        path = os.path.join(folder, fname)
        if os.path.exists(path) and not req.get('overwrite'):
            return self.send(200, {'ok': False, 'exists': True, 'path': path})
        if os.path.exists(os.path.join(folder, '~$' + fname)):
            raise UserError('같은 이름의 파일이 Excel 에 열려 있습니다. 닫고 다시 저장하세요.')
        os.makedirs(folder, exist_ok=True)
        shutil.copyfile(src, path)
        total = req.get('total')
        remember_site(site, path, year, total)
        log.info('save %s', path)
        self.send(200, {'ok': True, 'path': path, 'folder': folder})

    def _loaded(self, path, name):
        try:
            data = site_io.read_inputs(path)
        except ValueError as e:
            raise UserError(str(e))
        year = data.pop('year', None)
        if not year:
            wb = openpyxl.load_workbook(path, read_only=True)
            m = re.search(r"(\d{2})년 예산", ' '.join(wb.sheetnames))
            wb.close()
            year = 2000 + int(m.group(1)) if m else None
        if not data.get('kinds'):
            data['kinds'] = site_io.guess_kinds(data.get('fields') or {}, year)
        return self.send(200, {'ok': True, 'data': data, 'year': year, 'name': name})

    def api_report(self):
        jid = self.jbody().get('job')
        meta = ensure_report(jid)
        return self.send(200, {'ok': True, **meta})

    def api_form(self):
        """회사 양식(경영기획팀 사업비 예산편성 양식) 바꾸기"""
        files, form = parse_multipart(self.headers.get('Content-Type', ''), self.body())
        if 'file' not in files:
            raise UserError('양식 엑셀 파일을 넣어 주세요.')
        name, data = files['file']
        if not name.lower().endswith('.xlsx'):
            raise UserError('엑셀(.xlsx) 파일이어야 합니다.')
        jid = new_job()
        p = os.path.join(JOBS, jid, 'form.xlsx')
        with open(p, 'wb') as f:
            f.write(data)
        try:
            wb = openpyxl.load_workbook(p, read_only=True)
            names = wb.sheetnames
            wb.close()
        except Exception as e:
            raise UserError(f'엑셀을 열 수 없습니다: {e}')
        if '예산안 보고서' not in names or not any(re.fullmatch(r'\d{2}년 예산', n) for n in names):
            raise UserError("회사 양식이 아닌 것 같습니다('예산안 보고서'와 'NN년 예산' 시트가 있어야 함).")
        dst = os.path.join(FORMS, safe_name(name))
        if os.path.exists(dst):
            os.replace(dst, dst[:-5] + time.strftime('_%Y%m%d_%H%M%S') + '.bak')
        shutil.copyfile(p, dst)
        os.utime(dst, None)
        log.info('form %s', dst)
        fi = form_info()
        return self.send(200, {'ok': True, 'form': {'name': fi['name'], 'year': fi['year']} if fi else None})

    def api_types(self):
        """설계개요(엑셀·PDF) → 세대 구성(타입·세대수·공급면적). 그림 PDF 는 Windows OCR 로 읽어 10초 안팎"""
        files, form = parse_multipart(self.headers.get('Content-Type', ''), self.body())
        if 'file' not in files:
            raise UserError('설계개요 파일(엑셀 또는 PDF)을 넣어 주세요.')
        name, data = files['file']
        ext = os.path.splitext(name)[1].lower()
        if ext not in ('.xlsx', '.xlsm', '.xls', '.pdf'):
            raise UserError('설계개요는 엑셀(xlsx·xls) 또는 PDF 파일이어야 합니다.')
        jid = new_job()
        p = os.path.join(JOBS, jid, 'overview' + ext)
        with open(p, 'wb') as f:
            f.write(data)
        t = time.time()
        try:
            r = design_overview.read_types(p, max_types=site_io.BT.NTYPES, log=log.info)
        except ValueError as e:
            raise UserError(str(e))
        log.info('types %s → %s %s %d개 %.1fs', name, r['source'], r['where'], len(r['types']), time.time() - t)
        self.send(200, {'ok': True, 'name': name, **r})

    def api_load(self):
        files, form = parse_multipart(self.headers.get('Content-Type', ''), self.body())
        if 'file' not in files:
            raise UserError('엑셀 파일을 넣어 주세요.')
        name, data = files['file']
        if not name.lower().endswith('.xlsx'):
            raise UserError('엑셀(.xlsx) 파일이어야 합니다.')
        jid = new_job()
        p = os.path.join(JOBS, jid, 'load.xlsx')
        with open(p, 'wb') as f:
            f.write(data)
        return self._loaded(p, name)

    def api_load_path(self):
        self.need_local()
        p = fix_drive(self.jbody().get('path') or '')
        if not os.path.exists(p):
            raise UserError('파일이 없습니다: ' + p)
        return self._loaded(p, os.path.basename(p))

    def api_kijun(self):
        files, form = parse_multipart(self.headers.get('Content-Type', ''), self.body())
        if 'file' not in files:
            raise UserError('기준안 엑셀 파일을 넣어 주세요.')
        name, data = files['file']
        if not name.lower().endswith('.xlsx'):
            raise UserError('엑셀(.xlsx) 파일이어야 합니다.')
        jid = new_job()
        p = os.path.join(JOBS, jid, name)
        with open(p, 'wb') as f:
            f.write(data)
        y, tpl = template_for(form.get('year'))
        r = apply_kijun.apply(p, tpl)
        rep = [dict(kind=k, desc=d, old=o, new=n, why=w) for k, d, o, n, w in r['report'] if k != '동일']
        same = sum(1 for x in r['report'] if x[0] == '동일')
        unused = [dict(row=ln['row'], item=' '.join(ln['item'].split()), text=' '.join(ln['text'].split()))
                  for ln in r['unused']]
        log.info('kijun %s → %s', name, r['out'])
        self.send(200, {'ok': True, 'year': r['year'], 'out': os.path.basename(r['out']), 'from': os.path.basename(tpl),
                        'renamed': r['renamed'], 'report': rep, 'same': same, 'unused': unused,
                        'written': str(r['written']) if r['written'] else None,
                        'backup': os.path.basename(r['backup']) if r['backup'] else None})

    def api_open(self):
        self.need_local()
        req = self.jbody()
        what = req.get('what')
        if what == 'template':
            y, p = template_for(req.get('year'))
            os.startfile(p)
        elif what == 'tooldir':
            os.startfile(ROOT)
        elif what == 'budgetdir':
            os.startfile(BUDGET_ROOT)
        else:
            p = req.get('path') or ''
            if not p.startswith(BUDGET_ROOT) or not os.path.exists(p):
                raise UserError('열 수 없는 경로입니다.')
            os.startfile(p)
        self.send(200, {'ok': True})


def main():
    global WORKER
    local_only = '--local' in sys.argv
    host = '127.0.0.1' if local_only else '0.0.0.0'
    try:
        srv = ThreadingHTTPServer((host, PORT), Handler)
    except OSError:
        print(f'{PORT} 번 포트를 이미 쓰고 있습니다. 서버가 이미 켜져 있을 수 있습니다.')
        log.warning('port %s busy', PORT)
        sys.exit(2)
    srv.daemon_threads = True
    WORKER = ExcelWorker()
    lines = [f'http://localhost:{PORT}/', '' if local_only else f'http://{lan_ip()}:{PORT}/?k={KEY}']
    with open(URLFILE, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    log.info('start %s', lines)
    try:
        print('판관비 예산안 자동작성 서버가 켜졌습니다.')
        print('  이 PC  :', lines[0])
        if lines[1]:
            print('  다른 PC:', lines[1])
    except Exception:
        pass
    try:
        srv.serve_forever()
    finally:
        WORKER.stop()
        try:
            os.remove(URLFILE)
        except OSError:
            pass
        log.info('stop')


if __name__ == '__main__':
    main()
