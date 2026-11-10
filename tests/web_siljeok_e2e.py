# -*- coding: utf-8 -*-
"""사업관리비 실적계획 화면 E2E: 헤드리스 Edge + CDP — 준비(웹판은 Pyodide) → 예시 불러오기 → 작성자·계정·증액 바꾸기 → 파일 칸으로 둘째 현장 올리기
   → 양식·근거자료 내려받아 PC 엔진(tool/siljeok.py)으로 만든 것과 칸마다 비교 → 화면 캡처

  웹판:  저장소 폴더에서  python -m http.server 8799 --bind 127.0.0.1  을 켠 뒤   python tests\\web_siljeok_e2e.py
  PC 판: 서버를 켠 뒤(실적계획 열기.bat)                                        python tests\\web_siljeok_e2e.py http://127.0.0.1:8792/siljeok/ 9337
"""
import base64
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request

sys.stdout.reconfigure(encoding="utf-8")
import site
site.addsitedir(site.getusersitepackages())
import warnings
warnings.filterwarnings("ignore")
import websocket          # noqa: E402
import openpyxl           # noqa: E402

SP = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(SP)
sys.path.insert(0, os.path.join(ROOT, "tool"))
sys.path.insert(0, SP)
import siljeok as S                # noqa: E402
import siljeok_samples as G        # noqa: E402

BASE = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8799/docs/siljeok/"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if not os.path.exists(EDGE):
    EDGE = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
port = int(sys.argv[2]) if len(sys.argv) > 2 else 9338
MODE = "pc" if "8792" in BASE else "web"
prof = tempfile.mkdtemp(prefix="edgeprof_siljeok_")          # 새 프로필(옛 프로필은 Edge 업데이트 뒤 디버그 포트가 안 뜨는 일이 있었음)
EXDIR = os.path.join(ROOT, "docs", "siljeok", "examples")
NEXT1 = os.path.join(ROOT, "docs", "examples", "(예시1) 신규분양 - 27.04 오픈 C급 조직분양.xlsx")
TPL = open(os.path.join(ROOT, "회사양식", "실적계획", "사업관리비 실적계획 양식.xlsx"), "rb").read()


def rd(p):
    with open(p, "rb") as f:
        return f.read()


def cells(xlsx):
    wb = openpyxl.load_workbook(io.BytesIO(xlsx))
    return {(ws.title, c.coordinate): c.value for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.value is not None}


def same_cells(a, b):
    ca, cb = cells(a), cells(b)
    diff = []
    for k in set(ca) | set(cb):
        x, y = ca.get(k), cb.get(k)
        if x == y:
            continue
        if isinstance(x, (int, float)) and isinstance(y, (int, float)) and abs(x - y) < 1e-6:
            continue
        if isinstance(x, str) and isinstance(y, str) and x[:4] == y[:4] == "작성 2":      # 근거자료의 작성일 줄
            continue
        diff.append((k, x, y))
    return len(ca), diff


proc = subprocess.Popen([EDGE, "--headless=new", "--disable-gpu", f"--remote-debugging-port={port}", f"--user-data-dir={prof}",
                         "--window-size=1600,1000", "--no-first-run", "about:blank"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
try:
    tabs = None
    for _ in range(100):
        try:
            tabs = json.loads(urllib.request.urlopen(f"http://127.0.0.1:{port}/json").read())
            break
        except Exception:
            time.sleep(0.2)
    tab = next(t for t in tabs if t["type"] == "page")
    ws = websocket.create_connection(tab["webSocketDebuggerUrl"], suppress_origin=True)
    ws.settimeout(600)
    mid = [0]
    logs = []

    def send(method, **params):
        mid[0] += 1
        ws.send(json.dumps({"id": mid[0], "method": method, "params": params}))
        while True:
            m = json.loads(ws.recv())
            if m.get("method") == "Runtime.consoleAPICalled":
                args = m["params"].get("args", [])
                logs.append(" ".join(str(a.get("value", a.get("description", ""))) for a in args)[:3000])
            if m.get("id") == mid[0]:
                if "error" in m:
                    raise RuntimeError(f"{method}: {m['error']}")
                return m.get("result", {})

    def js(expr, timeout=120):
        r = send("Runtime.evaluate", expression=expr, awaitPromise=True, returnByValue=True, timeout=timeout * 1000)
        if "exceptionDetails" in r:
            ex = r["exceptionDetails"]
            raise RuntimeError(ex.get("exception", {}).get("description") or ex.get("text"))
        return r.get("result", {}).get("value")

    def wait(expr, secs, step=0.4):
        t = time.time()
        while time.time() - t < secs:
            if js(expr):
                return time.time() - t
            time.sleep(step)
        raise TimeoutError(expr + " | 배너: " + str(js("document.querySelector('#banner').textContent")) + " | 로그: " + " / ".join(logs[-5:]))

    def settle(secs=180):
        """파일 읽기·계산이 끝날 때까지: 읽는 중인 현장 없음 + 마지막 계산 결과가 지금 입력과 같은 차례"""
        time.sleep(0.5)
        wait("S.sites.every(s => !s.busy) && !document.querySelector('.site .spin')", secs)
        seq = js("S.seq")
        time.sleep(0.6)
        wait(f"S.seq === {seq} || true", 5)
        time.sleep(0.5)
        wait("S.sites.every(s => !s.busy)", secs)

    def shot(name):
        h = js("document.documentElement.scrollHeight")
        send("Emulation.setDeviceMetricsOverride", width=1600, height=int(min(h, 7000)), deviceScaleFactor=1, mobile=False)
        r = send("Page.captureScreenshot", format="png")
        fn = os.path.join(SP, name)
        open(fn, "wb").write(base64.b64decode(r["data"]))
        print("캡처", fn)

    def upload(selector, paths):
        doc = send("DOM.getDocument", depth=0)
        node = send("DOM.querySelector", nodeId=doc["root"]["nodeId"], selector=selector)
        assert node.get("nodeId"), selector
        send("DOM.setFileInputFiles", files=paths, nodeId=node["nodeId"])

    def fetch_xlsx(kind):
        b64 = js("(async () => { const r = await route('" + kind + """', {state: collect(), name: 'x.xlsx'}); const u = r.xlsx; let s = '';
          for (let i = 0; i < u.length; i += 8192) s += String.fromCharCode.apply(null, u.subarray(i, i + 8192)); return btoa(s); })()""", 300)
        return base64.b64decode(b64)

    send("Page.enable")
    send("Runtime.enable")
    send("DOM.enable")
    send("Page.navigate", url=BASE)
    dt = wait("!!(typeof S !== 'undefined' && S.info)", 400, 1.0)
    print(f"준비 완료 {dt:.0f}초 · 모드 {js('S.mode')} · 틀 {js('S.info.template')} · 기준 {js('S.info.base_year')} · 제목 {js('document.title')}")
    assert js("S.mode") == MODE, js("S.mode")
    js("localStorage.clear()")
    assert js("document.querySelector('#bForm').disabled") is True

    # ── 예시 불러오기 → 로컬 엔진과 비교
    js("loadExample()")
    wait("S.sites[0].prev && S.sites[0].next && S.sites[0].ledgers.length === 2", 300)
    settle()
    wait("S.plan && S.plan.sites[0].rows.length > 5 && S.plan.sites[0].has_ledger && S.plan.sites[0].total.f > 0", 60)
    p = json.loads(js("JSON.stringify(S.plan)"))
    files = {n: rd(os.path.join(EXDIR, n)) for n in sorted(os.listdir(EXDIR))}
    prev = S.parse_file(files["(예시) 전년도 예산 - 26년 예산안.xlsx"], "(예시) 전년도 예산 - 26년 예산안.xlsx")
    nxt = S.parse_file(rd(os.path.join(ROOT, "docs", "examples", G.NEXT_EXAMPLE)), G.NEXT_EXAMPLE)
    lg = [S.parse_file(files[n], n) for n in ("(예시) 사용내역 1-8월.xlsx", "(예시) 사용내역 9-11월.xlsx")]
    st = dict(year=2027, sites=[dict(name=G.SITE, code=G.CODE, prev=prev, next=nxt, ledgers=lg, creators=list(G.STAFF), accmap={}, inc={},
                                     note_diff="", note_change="")])
    lp = S.slim(S.plan(json.loads(json.dumps(st))))
    s0, l0 = p["sites"][0], lp["sites"][0]
    print(f"예시: {s0['title']} · 비목 {len(s0['rows'])} · 26년 예산 {s0['total']['d']:,.0f} · 실적 {s0['total']['f']:,.0f}+{s0['total']['g']:,.0f} · 27년 {s0['total']['k']:,.0f}"
          f" · 작성자 {[c['name'] for c in s0['creators'] if c['on']]} · 경고 {len(s0['warnings'])}")
    assert p["year"] == 2027 and js("S.year") == 2027 and js("document.querySelector('#year').value") == "2027"
    def first_diff(a, b, path=""):
        if isinstance(a, dict) and isinstance(b, dict):
            for k in sorted(set(a) | set(b)):
                d = first_diff(a.get(k), b.get(k), f"{path}.{k}")
                if d:
                    return d
            return None
        if isinstance(a, (list, tuple)) and isinstance(b, (list, tuple)):
            if len(a) != len(b):
                return f"{path}: 길이 {len(a)} vs {len(b)}"
            for i, (x, y) in enumerate(zip(a, b)):
                d = first_diff(x, y, f"{path}[{i}]")
                if d:
                    return d
            return None
        if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) and not isinstance(b, bool):
            return None if abs(a - b) <= 1e-6 * max(1.0, abs(b)) else f"{path}: {a!r} vs {b!r}"
        return None if a == b else f"{path}: {a!r} vs {b!r}"

    print("  경고:", s0["warnings"])
    d = first_diff(s0["rows"], l0["rows"])
    assert not d, "예시 계획이 로컬 엔진과 다름 — " + str(d)
    assert s0["title"] == l0["title"] == f"[현장명 : {G.SITE} ({G.CODE})]", s0["title"]
    assert js("document.querySelector('.sname').value") == G.SITE and js("document.querySelector('.scode').value") == G.CODE
    assert js("document.querySelectorAll('.creators .chk').length") == 4 and js("document.querySelectorAll('.creators .chk.on').length") == 2
    assert js("document.querySelectorAll('table.out tr').length") == len(s0["rows"]) + 3
    assert js("document.querySelector('#bForm').disabled") is False and js("document.querySelector('#bEv').disabled") is False
    print("  화면 = 로컬 엔진 · 이름·코드 자동 · 작성자 2/4 · 파일 이름:", js("document.querySelector('#fname').value"))

    # ── 작성자 풀기/고르기, 계정 옮기기, 증액
    js("document.querySelector('.creators input[data-name=\"박세무\"]').click()")
    settle()
    wait("S.plan.sites[0].rows.find(r => r.item === '세금과공과').f > 0", 30)
    assert js("JSON.parse(localStorage.getItem('siljeok_creators'))['박세무']") is True
    js("document.querySelector('.creators input[data-name=\"박세무\"]').click()")
    settle()
    wait("S.plan.sites[0].rows.find(r => r.item === '세금과공과').f === 0", 30)
    for nm in G.STAFF:                       # 예시가 골라 준 이름은 기억에 없다 → 직접 풀었다 다시 고르면 기억됨
        for _ in range(2):
            js(f"document.querySelector('.creators input[data-name=\"{nm}\"]').click()")
            settle()
    mem = json.loads(js("localStorage.getItem('siljeok_creators')"))
    assert mem == {"박세무": False, G.STAFF[0]: True, G.STAFF[1]: True}, mem
    wait("S.plan.sites[0].creators.filter(c => c.on).length === 2", 30)
    js("""(() => { const sel = document.querySelector('.accts select[data-acct="판매관리비(지급수수료)_기타_D"]'); sel.value = '-'; sel.dispatchEvent(new Event('change')); })()""")
    settle()
    wait("S.plan.sites[0].accounts.find(a => a.name.includes('_기타_')).item === '-'", 30)
    js("""(() => { const i = document.querySelector('input.inc[data-item="통신비"]'); i.value = '1,000,000'; i.dispatchEvent(new Event('change')); })()""")
    settle()
    wait("S.plan.sites[0].rows.find(r => r.item === '통신비').e === 1000000", 30)
    js("""(() => { const a = document.querySelector('.note_diff'); a.value = '예시 분석'; a.dispatchEvent(new Event('input')); })()""")
    print("  작성자 고르기 기억 · 계정 빼기 · 증액 · 분석 문구 반영")

    # ── 둘째 현장: 파일 칸으로 올리기(금년도 예산안 = 예시1, 전년도 예산을 금년도 칸에 넣어도 연도로 알아서 감)
    js("document.querySelector('#bAdd').click()")
    settle()
    assert js("S.plan.sites.length") == 1 and js("document.querySelector('#bForm').disabled") is False, "빈 현장 카드는 계산·내려받기에서 빠져야 함"
    sid = js("S.sites[1].id")
    upload(f'.site[data-id="{sid}"] .slot[data-slot=next] input[type=file]', [NEXT1])
    settle()
    upload(f'.site[data-id="{sid}"] .slot[data-slot=next] input[type=file]', [os.path.join(EXDIR, "(예시) 전년도 예산 - 26년 예산안.xlsx")])
    settle()
    upload(f'.site[data-id="{sid}"] .slot[data-slot=ledger] input[type=file]', [os.path.join(EXDIR, n) for n in ("(예시) 사용내역 1-8월.xlsx", "(예시) 사용내역 9-11월.xlsx")])
    settle()
    wait("S.plan && S.plan.sites.length === 2 && S.plan.sites[1].rows.length > 5 && S.plan.sites[1].has_ledger", 120)
    s1 = json.loads(js("JSON.stringify(S.plan.sites[1])"))
    print(f"둘째 현장: {s1['title']} · 27년 {s1['total']['k']:,.0f} · 26년 예산 {s1['total']['d']:,.0f} · 작성자 {[c['name'] for c in s1['creators'] if c['on']]}"
          f" · 알림 {js('S.sites[1].msgs.map(m => m.text).join(\" | \")')}")
    assert js("!!S.sites[1].prev && !!S.sites[1].next && S.sites[1].next.year === 2027 && S.sites[1].prev.year === 2026")
    assert s1["name"].startswith("예시1") and s1["code"] == G.CODE and abs(s1["total"]["k"] - 6105846000) < 1, (s1["name"], s1["code"], s1["total"]["k"])
    assert [c["name"] for c in s1["creators"] if c["on"]] == sorted(G.STAFF, key=lambda n: [c["name"] for c in s1["creators"]].index(n)), "기억한 작성자가 미리 골라져야 함"
    assert "전년도 예산 칸에 넣었습니다" in js("S.sites[1].msgs.map(m => m.text).join(' | ')")
    # 틀린 파일
    upload(f'.site[data-id="{sid}"] .slot[data-slot=prev] input[type=file]', [os.path.join(ROOT, "회사양식", "실적계획", "사업관리비 실적계획 양식.xlsx")])
    settle()
    assert "알 수 없는 파일" in js("S.sites[1].msgs.map(m => m.text).join(' | ')")
    js("document.querySelectorAll('.bMon')[0].click()")
    shot("web_siljeok.png" if MODE == "web" else "pc_siljeok.png")

    # ── 내려받기: 화면 상태 그대로 로컬 엔진으로 만든 것과 칸마다 비교
    state = json.loads(js("JSON.stringify(collect())"))
    form = fetch_xlsx("form")
    n, diff = same_cells(form, S.build_form(TPL, state))
    print(f"양식 {len(form):,} bytes · 칸 {n} · 로컬 엔진과 다른 칸 {len(diff)} {diff[:3]}")
    assert not diff
    wf = openpyxl.load_workbook(io.BytesIO(form)).worksheets[0]
    t1 = 3 + 2 * len(state["sites"]) * 0 + 2 * len(json.loads(js("JSON.stringify(S.plan.sites[0].rows)"))) + 1
    assert wf["A1"].value == f"[현장명 : {G.SITE} ({G.CODE})]" and wf[f"A{t1}"].value == "합 계" and str(wf[f"A{t1 + 2}"].value).startswith("[현장명 : 예시1")
    assert wf["J4"].value == " 예시 분석" and wf["J6"].value == " 상동"
    ev = fetch_xlsx("evidence")
    n, diff = same_cells(ev, S.build_evidence(state))
    print(f"근거자료 {len(ev):,} bytes · 칸 {n} · 로컬 엔진과 다른 칸 {len(diff)} {diff[:3]}")
    assert not diff
    out = os.path.join(SP, "web_siljeok_form.xlsx" if MODE == "web" else "pc_siljeok_form.xlsx")
    open(out, "wb").write(form)
    # 단추로 내려받기(화면 메시지), 현장 삭제
    js("document.querySelector('#bForm').click()")
    wait("document.querySelector('#smsg').textContent.includes('내려받았습니다')", 120)
    js("window.confirm = () => true; document.querySelectorAll('.site .sdel')[1].click()")
    settle()
    assert js("S.sites.length") == 1 and js("S.plan.sites.length") == 1
    errs = [l for l in logs if "PAGEERR" in l or "PAGEREJ" in l or "Traceback" in l]
    print("콘솔 오류:", errs[:5] if errs else "없음")
    assert not errs
    print(("PC 판 화면" if MODE == "pc" else "웹판") + " E2E 통과")
    ws.close()
finally:
    proc.kill()
    time.sleep(0.5)
    shutil.rmtree(prof, ignore_errors=True)
