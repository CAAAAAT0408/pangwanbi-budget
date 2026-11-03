"""웹판 준비 과정 진단: 배너·콘솔 메시지를 10초마다 출력"""
import sys, os, json, time, subprocess, urllib.request
sys.stdout.reconfigure(encoding="utf-8")
import site
site.addsitedir(site.getusersitepackages())
import websocket

SP = os.path.dirname(os.path.abspath(__file__))
BASE = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8799/docs/"
LIMIT = int(sys.argv[2]) if len(sys.argv) > 2 else 300
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if not os.path.exists(EDGE):
    EDGE = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
prof = os.path.join(SP, "edgeprof_web")
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
    ws.settimeout(60)
    mid = [0]
    logs = []

    def pump(msg):
        if msg.get("method") == "Runtime.consoleAPICalled":
            args = msg["params"].get("args", [])
            logs.append(msg["params"].get("type", "") + ": " + " ".join(str(a.get("value", a.get("description", ""))) for a in args)[:300])
        elif msg.get("method") == "Runtime.exceptionThrown":
            logs.append("EXC: " + json.dumps(msg["params"].get("exceptionDetails", {}).get("exception", {}).get("description", ""))[:300])
        elif msg.get("method") == "Log.entryAdded":
            e = msg["params"]["entry"]
            logs.append(f"LOG[{e.get('level')}]: {e.get('text', '')[:200]} {e.get('url', '')[:100]}")

    def send(method, **params):
        mid[0] += 1
        ws.send(json.dumps({"id": mid[0], "method": method, "params": params}))
        while True:
            m = json.loads(ws.recv())
            pump(m)
            if m.get("id") == mid[0]:
                return m.get("result", {})

    def js(expr):
        r = send("Runtime.evaluate", expression=expr, awaitPromise=True, returnByValue=True)
        if "exceptionDetails" in r:
            return "EXC " + str(r["exceptionDetails"].get("exception", {}).get("description", ""))[:200]
        return r.get("result", {}).get("value")

    send("Page.enable")
    send("Runtime.enable")
    send("Log.enable")
    send("Page.navigate", url=BASE)
    t0 = time.time()
    shown = 0
    while time.time() - t0 < LIMIT:
        time.sleep(10)
        ready = js("!!(window.S && S.schema)")
        banner = js("document.querySelector('#banner') ? document.querySelector('#banner').textContent : '(no banner)'")
        st = js("JSON.stringify({rseq: typeof rseq === 'undefined' ? null : rseq, pending: typeof pending === 'undefined' ? null : Object.keys(pending), info: !!(window.S && S.info), year: window.S && S.year, err: [...document.querySelectorAll('.msg.err')].map(e => e.textContent.slice(0, 120))})")
        print(f"[{time.time() - t0:5.0f}s] ready={ready} 배너={str(banner)[:80]!r} {st}")
        for l in logs[shown:]:
            print("     ", l)
        shown = len(logs)
        if ready:
            break
    ws.close()
finally:
    proc.kill()
