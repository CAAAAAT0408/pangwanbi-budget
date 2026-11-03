"""웹판 E2E: 헤드리스 Edge + CDP — Pyodide 준비 → 예시1 불러와 계산(합계 비교) → 설계개요 엑셀·PDF 읽기 → 보고서 xlsx 만들기(값 비교)"""
import sys, os, json, time, subprocess, base64, urllib.request, shutil, glob
sys.stdout.reconfigure(encoding="utf-8")
import site
site.addsitedir(site.getusersitepackages())
import websocket
import warnings
warnings.filterwarnings("ignore")

SP = os.path.dirname(os.path.abspath(__file__))
BASE = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8799/docs/"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if not os.path.exists(EDGE):
    EDGE = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
prof = os.path.join(SP, "edgeprof_web")          # CDN 내려받은 것을 다음 실행에도 쓰려고 고정 프로필
port = 9334
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
                logs.append(" ".join(str(a.get("value", a.get("description", ""))) for a in args)[:6000])
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
    t0 = time.time()
    send("Page.navigate", url=BASE)
    dt = wait("!!(typeof S !== 'undefined' && S.schema)", 400, 1.0)
    print(f"준비 완료 {dt:.0f}초 · 배너: {js('document.querySelector(\"#banner\").textContent')!r}")
    print("info:", js("JSON.stringify({year: S.year, tpl: S.info && S.info.template && S.info.template.kijun, form: S.info && S.info.form})"))
    js("localStorage.clear()")

    # 예시1 불러오기
    js("""(async () => { const b = await (await fetch('/예시/' + encodeURIComponent('(예시1) 신규분양 - 27.04 오픈 C급 조직분양.xlsx'))).blob();
      const f = new File([b], '(예시1) 신규분양.xlsx'); window.__err = null;
      try { await loaded(await upload('/api/load', f)); } catch (e) { window.__err = String(e); } })()""")
    wait("!!(S.res) || !!window.__err", 120)
    print("예시1 불러오기 오류:", js("window.__err"), "| 합계:", js("document.querySelector('#tot').textContent"),
          "| 오류 셀:", js("S.res && S.res.nerrors"), "| 계산 시간:", js("S.res && S.res.sec"))
    print("산식(월별 계산 수식) 칸:", js("S.res && S.res.formulas ? S.res.formulas.flat().filter(x => x).map(x => x.filter(y => y).length).reduce((a, b) => a + b, 0) : null"),
          "| 화면 산식 칸:", js("document.querySelectorAll('#sum td[title]').length"))
    shot("web_example1.png")

    # 설계개요 엑셀 읽기
    print("samples:", js("(async () => { const a = await fetch('/samples/a66.xlsx'); const b = await fetch('/samples/a33.pdf'); return a.status + '/' + (await a.blob()).size + ' ' + b.status + '/' + (await b.blob()).size; })()"))
    js("""(async () => { const b = await (await fetch('/samples/a66.xlsx')).blob(); const f = new File([b], 'A66 설계개요.xlsx');
      window.__done = false; await readOverview(f); window.__done = true; })()""")
    wait("window.__done === true", 120)
    print("설계개요(엑셀):", js("document.querySelector('#ovmsg').textContent").strip()[:160])
    print("   타입:", js("[...document.querySelectorAll('#types tr')].filter(t => t.querySelector('.t') && t.querySelector('.t').value).map(t => t.querySelector('.t').value + ':' + t.querySelector('.n').value).join(', ')"))

    # 설계개요 PDF(그림) 읽기 — Tesseract (OCR 단어 기록 켜기)
    js("worker.postMessage({id: 0, cmd: 'debug', args: {on: true}})")
    t = time.time()
    js("""(async () => { const b = await (await fetch('/samples/a33.pdf')).blob(); const f = new File([b], 'A33 설계개요.pdf');
      window.__done = false; await readOverview(f); window.__done = true; })()""")
    wait("window.__done === true", 600, 1.0)
    print(f"설계개요(PDF OCR) {time.time() - t:.0f}초:", js("document.querySelector('#ovmsg').textContent").strip()[:200])
    print("   타입:", js("[...document.querySelectorAll('#types tr')].filter(t => t.querySelector('.t') && t.querySelector('.t').value).map(t => t.querySelector('.t').value + ':' + t.querySelector('.n').value).join(', ')"))

    # 보고서 xlsx (예시1 다시 불러와 만들기)
    js("""(async () => { const b = await (await fetch('/예시/' + encodeURIComponent('(예시1) 신규분양 - 27.04 오픈 C급 조직분양.xlsx'))).blob();
      const f = new File([b], '(예시1) 신규분양.xlsx'); S.res = null; await loaded(await upload('/api/load', f)); })()""")
    wait("!!(S.res)", 120)
    b64 = js("""(async () => { const r = await post('/api/report', {job: S.job}); const u = r.xlsx; let s = '';
      for (let i = 0; i < u.length; i += 8192) s += String.fromCharCode.apply(null, u.subarray(i, i + 8192)); return btoa(s); })()""", 300)
    out = os.path.join(SP, "web_report.xlsx")
    open(out, "wb").write(base64.b64decode(b64))
    print("보고서 xlsx", len(b64) * 3 // 4, "bytes →", out)
    js("$$('.tabs button')[1].click()")
    time.sleep(0.5)
    shot("web_preview.png")
    print("콘솔(일부):", [l for l in logs if "error" in l.lower() or "Error" in l][:5])
    for l in logs:
        if "OCR words" in l:
            print(l[:3000])
            print("-----")
    ws.close()
finally:
    proc.kill()
