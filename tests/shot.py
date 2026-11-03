"""헤드리스 Edge + CDP 로 화면 캡처: 기본 화면, 예시1 불러온 화면(세부 조정 접힘 확인)"""
import sys, os, json, time, subprocess, base64, urllib.request, shutil, tempfile
sys.stdout.reconfigure(encoding="utf-8")
import site
site.addsitedir(site.getusersitepackages())
import websocket

SP = os.path.dirname(os.path.abspath(__file__))
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if not os.path.exists(EDGE):
    EDGE = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
prof = tempfile.mkdtemp(prefix="edgeprof_")
port = 9333
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
    mid = [0]

    def send(method, **params):
        mid[0] += 1
        ws.send(json.dumps({"id": mid[0], "method": method, "params": params}))
        while True:
            m = json.loads(ws.recv())
            if m.get("id") == mid[0]:
                return m.get("result", {})

    def js(expr):
        r = send("Runtime.evaluate", expression=expr, awaitPromise=True, returnByValue=True)
        return r.get("result", {}).get("value")

    def shot(name, full=True):
        if full:
            h = js("document.documentElement.scrollHeight")
            send("Emulation.setDeviceMetricsOverride", width=1500, height=int(h), deviceScaleFactor=1, mobile=False)
        r = send("Page.captureScreenshot", format="png")
        fn = os.path.join(SP, name)
        open(fn, "wb").write(base64.b64decode(r["data"]))
        print("캡처", fn)

    send("Page.enable")
    send("Runtime.enable")
    send("Page.navigate", url="http://127.0.0.1:8792/")
    for _ in range(60):
        if js("!!(window.S && S.schema)"):
            break
        time.sleep(0.3)
    js("localStorage.clear()")
    send("Page.navigate", url="http://127.0.0.1:8792/")
    for _ in range(60):
        if js("!!(window.S && S.schema)"):
            break
        time.sleep(0.3)
    time.sleep(0.5)
    print("기본 화면 보이는 입력 칸:", js("document.querySelectorAll('#form .f:not(.hidden) input, #form .f:not(.hidden) select').length"),
          "세부 접힘 블록:", js("document.querySelectorAll('details.more').length"))
    shot("ui_default.png")
    # 예시1 불러오기
    js("""(async () => { const j = await post('/api/load_path', {path: document.querySelector('#recent option[value*="예시1"]').value}); await loaded(j); })()""")
    for _ in range(80):
        if js("!!(S.res)"):
            break
        time.sleep(0.4)
    time.sleep(0.5)
    print("예시1 보이는 입력 칸:", js("document.querySelectorAll('#form .f:not(.hidden) input, #form .f:not(.hidden) select').length"),
          "펼친 세부:", js("[...document.querySelectorAll('details.more[open]')].map(d => d.dataset.i).join(',')"),
          "타입 줄 보임:", js("document.querySelectorAll('#types tr:not(.hidden) input.t').length"),
          "합계:", js("document.querySelector('#tot').textContent"))
    shot("ui_example1.png")
    ws.close()
finally:
    proc.kill()
    time.sleep(0.5)
    shutil.rmtree(prof, ignore_errors=True)
