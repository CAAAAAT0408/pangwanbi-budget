"""공개 주소(GitHub Pages) 확인: 준비 → 예시1 불러와 계산 → 보고서 xlsx 크기 → 실제 현장 조건(A36BL 2026 유형)으로 계산해 PC 엔진과 비교"""
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
prof = os.path.join(SP, "edgeprof_pub_new")
port = 9341
proc = subprocess.Popen([EDGE, "--headless=new", "--disable-gpu", f"--remote-debugging-port={port}", f"--user-data-dir={prof}",
                         "--window-size=1500,1000", "--no-first-run", "about:blank"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
try:
    tabs = None
    for _ in range(100):
        try:
            tabs = json.loads(urllib.request.urlopen(f"http://127.0.0.1:{port}/json").read())
            break
        except Exception:
            time.sleep(0.3)
    if not tabs:
        raise SystemExit("Edge 디버그 포트가 열리지 않음")
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
    t0 = time.time()
    send("Page.navigate", url=BASE)
    dt = wait("!!(typeof S !== 'undefined' && S.schema)", 420)
    print(f"공개 주소 준비 {dt:.0f}초 · 기준안 {js('S.info.template.kijun')} · 회사양식 {js('S.info.form.name')} · 연도 {js('S.year')}")
    print("저장한 현장 목록:", js("[...document.querySelectorAll('#recent option')].map(o => o.textContent).join(' | ')")[:300])
    js("""(async () => { window.__err = null; try { await loaded(await post('/api/load_path', {path: 'example:(예시1) 신규분양 - 27.04 오픈 C급 조직분양.xlsx'})); } catch (e) { window.__err = String(e); } })()""")
    wait("!!(S.res) || !!window.__err", 180)
    print("예시1:", js("window.__err"), "| 합계", js("document.querySelector('#tot').textContent"), "| 오류 셀", js("S.res && S.res.nerrors"), "| 계산", js("S.res && S.res.sec"), "초")
    t = time.time()
    n = js("(async () => { const r = await post('/api/report', {job: S.job}); return r.xlsx.length; })()")
    print(f"보고서 xlsx bytes: {n} ({time.time()-t:.1f}초)")
    # 실제 현장 조건(2026 평택고덕 A36BL 과 같은 조건을 2027 로) 으로 계산 — PC 엔진(xlcalc) 결과와 비교
    data = json.load(open(os.path.join(SP, "web_case.json"), encoding="utf-8"))
    js("window.__case = " + json.dumps(data, ensure_ascii=False))
    r = js("""(async () => { const r = await post('/api/calc', {year: S.year, data: window.__case}); return {total: r.total.total, nerrors: r.nerrors, groups: r.groups.map(g => [g.item, g.total])}; })()""")
    print("웹판 실제조건 계산: 합계", f"{r['total']/1000:,.0f}천원", "오류", r["nerrors"])
    json.dump(r, open(os.path.join(SP, "web_case_result.json"), "w", encoding="utf-8"), ensure_ascii=False)
    rr = send("Page.captureScreenshot", format="png")
    open(os.path.join(SP, "web_public.png"), "wb").write(base64.b64decode(rr["data"]))
    print("오류 콘솔:", [l for l in logs if "rror" in l][:4])
    ws.close()
finally:
    proc.kill()
