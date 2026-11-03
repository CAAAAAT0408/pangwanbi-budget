# -*- coding: utf-8 -*-
"""지출계획서 웹판 E2E: 헤드리스 Edge + CDP — Pyodide 준비 → 예시 3개 계산 → 칸 직접 수정 → xlsx 내려받기(PC 엔진과 값 비교) → 화면 캡처
  저장소 폴더에서  python -m http.server 8799 --bind 127.0.0.1  을 켠 뒤  python tests\\web_jichul_e2e.py [http://127.0.0.1:8799/docs/jichul/]"""
import base64
import io
import json
import os
import subprocess
import sys
import time
import urllib.request

sys.stdout.reconfigure(encoding="utf-8")
import site
site.addsitedir(site.getusersitepackages())
import websocket          # noqa: E402
import openpyxl           # noqa: E402

SP = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(SP)
sys.path.insert(0, os.path.join(ROOT, "tool"))
import jichul             # noqa: E402
import jichul_xlsx        # noqa: E402

BASE = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8799/docs/jichul/"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if not os.path.exists(EDGE):
    EDGE = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
port = int(sys.argv[2]) if len(sys.argv) > 2 else 9335
prof = os.path.join(SP, "edgeprof_web" if port == 9335 else f"edgeprof_{port}")
MODE = "pc" if "8792" in BASE else "web"
proc = subprocess.Popen([EDGE, "--headless=new", "--disable-gpu", f"--remote-debugging-port={port}", f"--user-data-dir={prof}",
                         "--window-size=1500,1000", "--no-first-run", "about:blank"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
try:
    for _ in range(50):
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
                return m.get("result", {})

    def js(expr, timeout=120):
        r = send("Runtime.evaluate", expression=expr, awaitPromise=True, returnByValue=True, timeout=timeout * 1000)
        if "exceptionDetails" in r:
            ex = r["exceptionDetails"]
            raise RuntimeError(ex.get("exception", {}).get("description") or ex.get("text"))
        return r.get("result", {}).get("value")

    def wait(expr, secs, step=0.5):
        t = time.time()
        while time.time() - t < secs:
            if js(expr):
                return time.time() - t
            time.sleep(step)
        raise TimeoutError(expr)

    def shot(name):
        h = js("document.documentElement.scrollHeight")
        send("Emulation.setDeviceMetricsOverride", width=1500, height=int(min(h, 6000)), deviceScaleFactor=1, mobile=False)
        r = send("Page.captureScreenshot", format="png")
        fn = os.path.join(SP, name)
        open(fn, "wb").write(base64.b64decode(r["data"]))
        print("캡처", fn)

    send("Page.enable")
    send("Runtime.enable")
    send("Page.navigate", url=BASE)
    dt = wait("!!(typeof S !== 'undefined' && S.schema)", 400, 1.0)
    print(f"준비 완료 {dt:.0f}초 · 모드 {js('S.mode')} · 배너 {js('document.querySelector(\"#banner\").textContent')!r} · 기준 {js('S.year')} {js('S.info && S.info.kijun')}")
    assert js("S.mode") == MODE, js("S.mode")
    js("localStorage.clear()")

    K = {p["key"]: p["value"] for p in json.load(open(os.path.join(ROOT, "docs", "data", "kijun.json"), encoding="utf-8"))["params"] if "key" in p}
    for ex in ("open", "run", "movein"):
        js("S.res = null")
        js(f"example('{ex}')")
        wait("!!S.res", 120)
        d = json.loads(js("JSON.stringify(collect())"))
        local = jichul.calc(d, K)
        web_total = js("S.res.total_all")
        print(f"예시 {ex}: {js('S.res.title')} {js('JSON.stringify(S.res.months.map(m => m.label))')} 합계 {web_total:,.0f} "
              f"| PC 엔진 {local['total_all']:,.0f} {'일치' if abs(web_total - local['total_all']) < 1 else '불일치!'} · 경고 {js('JSON.stringify(S.res.warnings)')}")
        assert abs(web_total - local["total_all"]) < 1

    # 칸 직접 수정(산식) + 비고 수정 + 결재자 추가 → 다시 계산
    js("example('open')")
    wait("!!S.res", 120)
    js("S.overrides = {'소모품': {'1': '=1500000*2'}}; S.notes = {'운반비': '웹에서 쓴 비고'}; S.approvers = {list: [{role:'', pos:'대리', name:'김웹'}, {role:'', pos:'과장', name:''}, {role:'', pos:'부장', name:''}, {role:'', pos:'상무', name:''}, {role:'', pos:'실장', name:''}], width: 12, height: 55}; renderApprovers(); S.res = null; calc()")
    wait("!!S.res", 120)
    man = js("[...document.querySelectorAll('#out td.cell.man')].length")
    print("직접 수정 뒤: 노란 칸", man, "· 소모품 11월", js("rowOf('소모품').cells[1].v"), js("rowOf('소모품').cells[1].f"),
          "· 결재 칸", js("S.res.approvers.list.length"), "· 결재 미리보기", js("document.querySelector('#apPv').textContent").strip()[:60])
    assert man == 1 and js("rowOf('소모품').cells[1].v") == 3000000 and js("S.res.approvers.list.length") == 5
    shot("web_jichul.png" if MODE == "web" else "pc_jichul.png")

    # xlsx 내려받기(워커) → PC 엔진으로 만든 것과 값·수식 비교
    b64 = js("""(async () => { const r = await route('xlsx', {year: S.year, data: collect()}); const u = r.xlsx; let s = '';
      for (let i = 0; i < u.length; i += 8192) s += String.fromCharCode.apply(null, u.subarray(i, i + 8192)); return btoa(s); })()""", 300)
    xb = base64.b64decode(b64)
    out = os.path.join(SP, "web_jichul.xlsx" if MODE == "web" else "pc_jichul.xlsx")
    open(out, "wb").write(xb)
    d = json.loads(js("JSON.stringify(collect())"))
    pc = jichul_xlsx.build_xlsx(jichul.calc(d, K), d, 2027)
    w1, w2 = openpyxl.load_workbook(io.BytesIO(xb)), openpyxl.load_workbook(io.BytesIO(pc))
    s1, s2 = w1[jichul_xlsx.SHEET], w2[jichul_xlsx.SHEET]
    diff = [(c.coordinate, c.value, s2[c.coordinate].value) for row in s1.iter_rows() for c in row if c.value != s2[c.coordinate].value]
    nf = sum(1 for row in s1.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("="))
    print(f"웹 xlsx {len(xb):,} bytes → {out} · 수식 칸 {nf} · PC 엔진 xlsx 와 다른 칸 {len(diff)} {diff[:3]}")
    assert not diff and jichul_xlsx.read_inputs(xb)["data"] == d
    # 되읽기(불러오기)
    js("S.res = null; (async () => { const b = new Blob([Uint8Array.from(atob('" + b64 + "'), c => c.charCodeAt(0))]); const f = new File([b], 'web.xlsx'); await loaded(await route('load', null, f)); })()")
    wait("!!S.res", 120)
    print("불러오기:", js("document.querySelector('#loadmsg').textContent"), "· 노란 칸", js("[...document.querySelectorAll('#out td.cell.man')].length"),
          "· 결재", js("S.approvers.list.length"), "· 파일 이름", js("document.querySelector('#fname').value"))
    assert js("[...document.querySelectorAll('#out td.cell.man')].length") == 1 and js("S.approvers.list.length") == 5
    errs = [l for l in logs if "PAGEERR" in l or "PAGEREJ" in l or "Traceback" in l]
    print("콘솔 오류:", errs[:5] if errs else "없음")
    assert not errs
    print(("PC 판 화면" if MODE == "pc" else "웹판") + " E2E 통과")
    ws.close()
finally:
    proc.kill()
