"""브라우저 Tesseract 설정 비교: 언어·PSM 별로 A33 표 그림을 읽어 기대값(세대수·전용·공급)이 몇 개 맞는지"""
import sys, os, json, time, subprocess, urllib.request, re
sys.stdout.reconfigure(encoding="utf-8")
import site
site.addsitedir(site.getusersitepackages())
import websocket

SP = os.path.dirname(os.path.abspath(__file__))
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if not os.path.exists(EDGE):
    EDGE = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
prof = os.path.join(SP, "edgeprof_web")
port = 9335
EXP_U = ["309", "61", "30", "49", "24", "6", "75", "54", "6", "9", "623"]
EXP_E = ["84.9343", "84.9075", "84.8505", "94.3143", "94.2875", "94.3805", "101.8311", "101.8531", "111.5911", "111.4331"]
EXP_S = ["110.2002", "110.5732", "111.0782", "121.5382", "121.9112", "122.5975", "131.0502", "131.6827", "142.8475", "143.2625"]
EXP_T = ["84A", "84B", "84D", "94AT", "94BT", "94DT", "101A", "101B", "111AT", "111BT"]
configs = [a.split(":") for a in sys.argv[1:]] or [["crop_0.png", "kor+eng", "11"], ["crop_0.png", "eng", "11"], ["crop_0.png", "eng", "6"],
                                                       ["crop_0.png", "kor+eng", "6"], ["crop_1.png", "eng", "11"], ["crop_0.png", "eng", "4"]]
proc = subprocess.Popen([EDGE, "--headless=new", "--disable-gpu", f"--remote-debugging-port={port}", f"--user-data-dir={prof}",
                         "--window-size=1200,800", "--no-first-run", "about:blank"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
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
            if m.get("method") == "Runtime.exceptionThrown":
                logs.append("EXC " + str(m["params"].get("exceptionDetails", {}).get("exception", {}).get("description", ""))[:300])
            if m.get("method") == "Runtime.consoleAPICalled":
                logs.append(m["params"].get("type", "") + ": " + " ".join(str(a.get("value", a.get("description", ""))) for a in m["params"].get("args", []))[:300])
            if m.get("id") == mid[0]:
                return m.get("result", {})

    def js(expr):
        r = send("Runtime.evaluate", expression=expr, awaitPromise=True, returnByValue=True)
        return r.get("result", {}).get("value")
    send("Page.enable")
    send("Runtime.enable")
    for img, lang, psm in configs:
        url = f"http://127.0.0.1:8799/ocr_test.html?img=/samples/{img}&lang={lang}&psm={psm}"
        send("Page.navigate", url=url)
        t = time.time()
        out = None
        while time.time() - t < 240:
            time.sleep(1)
            out = js("window.__out || null")
            if out:
                break
        if not out:
            print(f"{img} {lang} psm{psm}: 시간 초과 · 화면글 {js('document.body.innerText.slice(0, 200)')!r} · 로그 {logs[-5:]}")
            continue
        d = json.loads(out)
        toks = [re.sub(r"[,\s]", "", w[4]) for w in d["words"]]
        joined = " ".join(toks)
        hit = lambda exp: sum(1 for e in exp if e in toks)
        nums = [w for w in d["words"] if re.search(r"\d", w[4])]
        print(f"{img} {lang:<8} psm{psm}: 준비 {d['init']/1000:.0f}s OCR {d['ocr']/1000:.0f}s 단어 {len(d['words'])} | 세대수 {hit(EXP_U)}/{len(EXP_U)} 전용 {hit(EXP_E)}/10 공급 {hit(EXP_S)}/10 타입 {hit(EXP_T)}/10")
        print("    숫자 토큰 예:", [w[4] for w in nums[:40]])
    ws.close()
finally:
    proc.kill()
