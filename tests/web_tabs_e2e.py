# -*- coding: utf-8 -*-
"""예산안 페이지 위쪽 탭(「판관비 예산안」 / 「사업관리비 실적계획」 — 한 페이지) E2E: 헤드리스 Edge + CDP
   ① 처음엔 예산안 탭(실적계획 틀은 아직 안 불러옴) → 예산안에 예시1 불러와 계산, 조금 내려 둠
   ② 실적계획 탭 → 아래 틀(iframe)에 실적계획 화면: 머리글 이동 줄 숨김 · 탭 제목 · 주소 #siljeok · 바깥 스크롤 없음 → 틀 안에서 '예시로 해 보기'
   ③ 예산안 탭 → 계산 결과·스크롤 위치 그대로   ④ 다시 실적계획 → 틀을 다시 불러오지 않고 그대로
   ⑤ 틀 안 「판관비 예산안」 탭 글귀 → 예산안 탭   ⑥ 옛 주소 …/siljeok/ → …/#siljeok 로 옮겨 감   ⑦ …/#siljeok 로 바로 열기 → 실적계획 탭 · 화면 캡처 2장

  웹판:  저장소 폴더에서  python -m http.server 8799 --bind 127.0.0.1  을 켠 뒤   python tests\\web_tabs_e2e.py
  PC 판: 서버를 켠 뒤(페이지 열기.bat)                                          python tests\\web_tabs_e2e.py http://127.0.0.1:8792/ 9339
"""
import base64
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

SP = os.path.dirname(os.path.abspath(__file__))
BASE = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8799/docs/"
if not BASE.endswith("/"):
    BASE += "/"
port = int(sys.argv[2]) if len(sys.argv) > 2 else 9339
MODE = "pc" if "8792" in BASE else "web"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if not os.path.exists(EDGE):
    EDGE = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
prof = tempfile.mkdtemp(prefix="edgeprof_tabs_")          # 새 프로필(끝나면 지움)
EX1 = "(예시1) 신규분양 - 27.04 오픈 C급 조직분양.xlsx"
EXB = "/siljeok/examples/" if MODE == "pc" else "examples/"   # PC 서버는 예산안 예시도 /siljeok/examples/ 로 내줌
FREADY = "typeof S !== 'undefined' && !!S.info"             # 실적계획 화면(틀 안) 준비 끝

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
            if m.get("method") == "Runtime.exceptionThrown":
                logs.append("PAGEERR " + json.dumps(m["params"].get("exceptionDetails", {}), ensure_ascii=False)[:1000])
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

    def fjs(expr, timeout=120):
        """실적계획 틀(iframe) 안에서 식 계산 — 같은 출처라 contentWindow.eval 로 그 화면의 전역(S·loadExample 등)을 그대로 봄"""
        return js("(async () => { const w = document.getElementById('sjFrame').contentWindow; return await w.eval(" + json.dumps(expr) + "); })()", timeout)

    def wait(fn, expr, secs, step=0.5):
        t, last = time.time(), None
        while time.time() - t < secs:
            try:
                if fn(expr):
                    return time.time() - t
            except RuntimeError as e:          # 페이지를 옮기는 중(문맥 바뀜)이면 다시 봄
                last = e
            time.sleep(step)
        raise TimeoutError(f"{expr} | 마지막 오류: {last} | 로그: {' / '.join(logs[-5:])}")

    def state():
        return js("""(() => { const f = document.getElementById('sjFrame');
          return {on: document.querySelector('#ptabs a.pt.on').dataset.p, url: location.href, hash: location.hash, title: document.title,
                  src: f.getAttribute('src'), frameHidden: f.classList.contains('hidden'), mainHidden: document.querySelector('main').classList.contains('hidden'),
                  headerHidden: document.querySelector('header').classList.contains('hidden'), sjmode: document.documentElement.classList.contains('sjmode'),
                  y: Math.round(scrollY)}; })()""")

    def shot(name):
        r = send("Page.captureScreenshot", format="png")
        fn = os.path.join(SP, name)
        open(fn, "wb").write(base64.b64decode(r["data"]))
        print("  캡처", fn)

    send("Page.enable")
    send("Runtime.enable")
    # 창 크기를 정해 둠(헤드리스 Edge 가 --window-size 를 안 따르고 작은 창으로 뜬 적이 있음)
    send("Emulation.setDeviceMetricsOverride", width=1600, height=1000, deviceScaleFactor=1, mobile=False)

    # ① 예산안 탭으로 시작 · 예시1 계산
    send("Page.navigate", url=BASE)
    dt = wait(js, "!!(typeof S !== 'undefined' && S.schema)", 400, 1.0)
    st = state()
    print(f"① 준비 {dt:.0f}초 · {MODE} · 탭 {st['on']} · 제목 {st['title']} · 틀 src {st['src']}")
    assert st["on"] == "budget" and st["src"] is None and st["frameHidden"] and not st["mainHidden"] and st["hash"] == "" and not st["sjmode"], st
    budget_title = st["title"]
    assert budget_title.startswith("판관비 예산안 자동작성"), budget_title
    assert js("document.querySelectorAll('#ptabs a.pt').length") == 2 and js("!!document.querySelector('#ptabs a.ext[href$=\"jichul/\"]')")
    assert js("[...document.querySelectorAll('header a')].length") == 0          # 머리글의 옛 바로가기는 탭 줄로 옮김
    js("""(async () => { const b = await (await fetch('""" + EXB + """' + encodeURIComponent('""" + EX1 + """'))).blob();
      const f = new File([b], '""" + EX1 + """'); window.__err = null;
      try { await loaded(await upload('/api/load', f)); } catch (e) { window.__err = String(e); } })()""", 300)
    wait(js, "!!(S.res) || !!window.__err", 300)
    assert not js("window.__err"), js("window.__err")
    tot = js("document.querySelector('#tot').textContent")
    js("window.scrollTo(0, 700)")
    y0 = js("Math.round(scrollY)")
    print(f"  예시1 계산 합계 {tot} 천원 · 스크롤 {y0}")
    assert tot and y0 > 0

    # ② 실적계획 탭 → 틀에 실적계획 화면
    js("document.querySelector('#ptabs a.pt[data-p=siljeok]').click()")
    st = state()
    assert st["on"] == "siljeok" and st["hash"] == "#siljeok" and st["mainHidden"] and st["headerHidden"] and not st["frameHidden"] and st["sjmode"], st
    assert st["src"] and st["src"].rstrip("/").endswith("siljeok"), st["src"]
    dt = wait(fjs, FREADY, 400, 1.0)
    wait(js, "document.title.startsWith('사업관리비 실적계획 자동작성_v')", 30)
    lay = js("""(() => { const f = document.getElementById('sjFrame'), d = f.contentDocument;
      return {embed: d.documentElement.classList.contains('embed'), nav: getComputedStyle(d.querySelector('header nav')).display,
              over: document.documentElement.scrollHeight - innerHeight, fh: Math.round(f.getBoundingClientRect().height),
              rest: Math.round(innerHeight - document.getElementById('ptabs').getBoundingClientRect().height), furl: f.contentWindow.location.pathname,
              fh1: d.querySelector('header h1').textContent}; })()""")
    print(f"② 실적계획 틀 준비 {dt:.0f}초 · 제목 {js('document.title')} · 틀 {lay['furl']} · 높이 {lay['fh']}/{lay['rest']} · 바깥 넘침 {lay['over']}")
    assert lay["embed"] and lay["nav"] == "none" and lay["over"] <= 1 and abs(lay["fh"] - lay["rest"]) <= 2, lay
    assert fjs("S.mode") == MODE and lay["fh1"].startswith("사업관리비 실적계획 자동작성")
    fjs("loadExample()", 300)
    wait(fjs, "S.sites[0].prev && S.sites[0].next && S.sites[0].ledgers.length === 2", 300)
    wait(fjs, "S.sites.every(s => !s.busy) && !!S.plan && S.plan.sites[0].rows.length > 5 && S.plan.sites[0].total.f > 0", 300)
    rows = fjs("S.plan.sites[0].rows.length")
    title_sj = fjs("S.plan.sites[0].title")
    js("document.getElementById('sjFrame').contentWindow.__kept = 'kept'")
    print(f"  틀 안 예시: {title_sj} · 비목 {rows}")
    shot(f"{MODE}_tabs_siljeok.png")

    # ③ 예산안 탭 → 계산 결과·스크롤 그대로
    js("document.querySelector('#ptabs a.pt[data-p=budget]').click()")
    st = state()
    assert st["on"] == "budget" and st["hash"] == "" and st["title"] == budget_title and not st["mainHidden"] and st["frameHidden"] and not st["sjmode"], st
    assert js("document.querySelector('#tot').textContent") == tot and js("!!S.res") and abs(st["y"] - y0) <= 2, (st, y0)
    print(f"③ 예산안 탭: 합계 {tot} 그대로 · 스크롤 {st['y']} (전 {y0}) · 주소 {st['url']}")

    # ④ 다시 실적계획 → 틀 그대로(다시 불러오지 않음)
    js("document.querySelector('#ptabs a.pt[data-p=siljeok]').click()")
    assert fjs("window.__kept") == "kept" and fjs("S.plan.sites[0].rows.length") == rows and fjs("S.sites.length") >= 1
    assert js("document.title").startswith("사업관리비 실적계획 자동작성_v")
    print("④ 다시 실적계획: 틀 그대로(올린 예시·계획 유지)")

    # ⑤ 틀 안 「판관비 예산안」 탭 글귀 → 예산안 탭
    fjs("document.querySelector('a[data-ptab=budget]').click()")
    st = state()
    assert st["on"] == "budget" and not st["mainHidden"] and st["frameHidden"] and fjs("window.__kept") == "kept", st
    print("⑤ 틀 안 「판관비 예산안」 탭 글귀 → 예산안 탭")
    shot(f"{MODE}_tabs_budget.png")

    # ⑥ 옛 주소 …/siljeok/ → …/#siljeok
    send("Page.navigate", url=BASE + "siljeok/")
    wait(js, "location.hash === '#siljeok' && !!document.getElementById('sjFrame') && !!document.getElementById('sjFrame').getAttribute('src')", 60)
    wait(fjs, FREADY, 400, 1.0)
    st = state()
    print(f"⑥ 옛 주소 {BASE}siljeok/ → {st['url']} · 탭 {st['on']}")
    assert st["on"] == "siljeok" and st["url"] == BASE + "#siljeok" and st["mainHidden"], st

    # ⑦ …/#siljeok 로 바로 열기
    send("Page.navigate", url="about:blank")
    time.sleep(0.5)
    send("Page.navigate", url=BASE + "#siljeok")
    wait(js, "!!document.getElementById('sjFrame') && !!document.getElementById('sjFrame').getAttribute('src')", 60)
    wait(fjs, FREADY, 400, 1.0)
    st = state()
    print(f"⑦ {BASE}#siljeok 바로 열기 → 탭 {st['on']} · 제목 {st['title']}")
    assert st["on"] == "siljeok" and st["hash"] == "#siljeok" and not st["frameHidden"], st

    # ⑧ 좁은 창(휴대폰 너비): 탭 줄이 한 줄 · 틀이 남은 높이를 채우고 바깥 넘침 없음
    send("Emulation.setDeviceMetricsOverride", width=420, height=800, deviceScaleFactor=1, mobile=False)
    time.sleep(0.5)
    lay = js("""(() => { const f = document.getElementById('sjFrame'), p = document.getElementById('ptabs'), a = [...p.querySelectorAll('a.pt')];
      return {rows: new Set(a.map(x => Math.round(x.getBoundingClientRect().top))).size, ph: Math.round(p.getBoundingClientRect().height),
              over: document.documentElement.scrollHeight - innerHeight, wide: document.documentElement.scrollWidth - innerWidth,
              fh: Math.round(f.getBoundingClientRect().height), rest: Math.round(innerHeight - p.getBoundingClientRect().height)}; })()""")
    print(f"⑧ 좁은 창 420px: 탭 {lay['rows']}줄 · 탭 줄 높이 {lay['ph']} · 틀 {lay['fh']}/{lay['rest']} · 넘침 세로 {lay['over']} 가로 {lay['wide']}")
    assert lay["rows"] == 1 and lay["over"] <= 1 and lay["wide"] <= 1 and abs(lay["fh"] - lay["rest"]) <= 2, lay
    shot(f"{MODE}_tabs_narrow.png")

    errs = [l for l in logs if "PAGEERR" in l or "PAGEREJ" in l or "Traceback" in l]
    print("콘솔 오류:", errs[:5] if errs else "없음")
    assert not errs
    print(("PC 판" if MODE == "pc" else "웹판") + " 위쪽 탭 E2E 통과")
    ws.close()
finally:
    proc.kill()
    time.sleep(0.5)
    shutil.rmtree(prof, ignore_errors=True)
