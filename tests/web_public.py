"""공개 주소(GitHub Pages) 확인: 준비 → 예시1 불러와 계산 → 보고서 xlsx 크기"""
import sys, os, json, time, subprocess, base64, urllib.request
sys.stdout.reconfigure(encoding="utf-8")
import site
site.addsitedir(site.getusersitepackages())
import websocket

SP = os.path.dirname(os.path.abspath(__file__))
BASE = sys.argv[1] if len(sys.argv) > 1 else "https://caaaaat0408.github.io/pangwanbi-budget/"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if not os.path.exists(EDGE):
    EDGE = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
prof = os.path.join(SP, "edgeprof_pub")
port = 9336
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
                logs.append(" ".join(str(a.get("value", a.get("description", ""))) for a in args)[:300])
            if m.get("id") == mid[0]:
                return m.get("result", {})

    def js(expr):
        r = send("Runtime.evaluate", expression=expr, awaitPromise=True, returnByValue=True)
        if "exceptionDetails" in r:
            raise RuntimeError(str(r["exceptionDetails"].get("exception", {}).get("description", ""))[:300])
        return r.get("result", {}).get("value")

    def wait(expr, secs):
        t = time.time()
        while time.time() - t < secs:
            if js(expr):
                return time.time() - t
            time.sleep(1)
        raise TimeoutError(expr)

    send("Page.enable")
    send("Runtime.enable")
    send("Page.navigate", url=BASE)
    dt = wait("!!(typeof S !== 'undefined' && S.schema)", 420)
    print(f"공개 주소 준비 {dt:.0f}초 · 기준안 {js('S.info.template.kijun')} · 회사양식 {js('S.info.form.name')}")
    print("저장한 현장 목록:", js("[...document.querySelectorAll('#recent option')].map(o => o.textContent).join(' | ')")[:200])
    js("""(async () => { window.__err = null; try { await loaded(await post('/api/load_path', {path: 'example:(예시1) 신규분양 - 27.04 오픈 C급 조직분양.xlsx'})); } catch (e) { window.__err = String(e); } })()""")
    wait("!!(S.res) || !!window.__err", 120)
    print("예시1:", js("window.__err"), "| 합계", js("document.querySelector('#tot').textContent"), "| 오류 셀", js("S.res && S.res.nerrors"), "| 계산", js("S.res && S.res.sec"), "초")
    n = js("(async () => { const r = await post('/api/report', {job: S.job}); return r.xlsx.length; })()")
    print("보고서 xlsx bytes:", n)
    r = send("Page.captureScreenshot", format="png")
    open(os.path.join(SP, "web_public.png"), "wb").write(base64.b64decode(r["data"]))
    print("오류 콘솔:", [l for l in logs if "rror" in l][:4])
    ws.close()
finally:
    proc.kill()
