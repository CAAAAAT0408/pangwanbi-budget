# -*- coding: utf-8 -*-
"""지출계획서 PC 판 API 시험(서버를 켠 뒤): /jichul/ 화면 · info · schema · calc(엔진과 같은 값) · xlsx · save(임시 폴더) · load_path · recent · open 제한"""
import json
import os
import sys
import tempfile
import urllib.parse
import urllib.request
import uuid

sys.stdout.reconfigure(encoding="utf-8")
B = "http://127.0.0.1:8792"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tool"))
import jichul            # noqa: E402
import jichul_xlsx       # noqa: E402


def get(p, raw=False):
    with urllib.request.urlopen(B + p, timeout=120) as r:
        return r.read() if raw else json.loads(r.read())


def post(p, obj, raw=False):
    req = urllib.request.Request(B + p, data=json.dumps(obj).encode(), headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=120) as r:
        return (r.read(), dict(r.headers)) if raw else json.loads(r.read())


def upload(p, path):
    bnd = uuid.uuid4().hex
    name = os.path.basename(path)
    body = (f"--{bnd}\r\nContent-Disposition: form-data; name=\"file\"; filename*=UTF-8''{urllib.parse.quote(name)}\r\n"
            f"Content-Type: application/octet-stream\r\n\r\n").encode() + open(path, "rb").read() + f"\r\n--{bnd}--\r\n".encode()
    req = urllib.request.Request(B + p, data=body, headers={"Content-Type": f"multipart/form-data; boundary={bnd}"}, method="POST")
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.loads(r.read())


html = get("/jichul/", raw=True).decode("utf-8")
assert "지출계획서 자동작성" in html and "worker.js" in html
print("/jichul/ 화면", len(html), "bytes")
info = get("/api/jichul/info")
print("info:", info["year"], info["kijun"], "단가", info["count"], "저장 위치", info["jichul_root"], "templates", [t["year"] for t in info["templates"]])
assert info["ok"] and info["count"] > 100
sc = get(f"/api/jichul/schema?year={info['year']}")
print("schema: 구역", [s["title"] for s in sc["sections"]], "구분", [k["key"] for k in sc["kinds"]], "행사", len(sc["events"]))
assert len(sc["sections"]) == 4 and len(sc["kinds"]) == 3 and len(sc["events"]) == 6
pr = get(f"/api/jichul/params?year={info['year']}")
assert pr["ok"] and len(pr["params"]) > 100

data = dict(kinds=["mh"], fields={"현장명": "API시험", "작성일": "2026-10-06", "시작": [2026, 10], "개월수": 3, "지역": "수도권",
            "세대수": 623, "오픈": [2026, 11], "사전영업": 1, "유니트": 4, "숙소호실": 4, "PC": 7}, months=[{}, {"events": ["중도금대출"]}, {}],
            approvers={"list": [{"pos": "과장"}, {"pos": "부장"}, {"pos": "실장"}], "width": 13})
r = post("/api/jichul/calc", {"year": info["year"], "data": data})
K = {p["key"]: p["value"] for p in pr["params"] if "key" in p}
local = jichul.calc(data, K)
print("calc:", r["title"], [m["label"] for m in r["months"]], f"합계 {r['total_all']:,.0f}", "파일", r["file_name"],
      "| 엔진 직접 계산과", "같음" if abs(r["total_all"] - local["total_all"]) < 1 else "다름!")
assert abs(r["total_all"] - local["total_all"]) < 1 and r["months"][1]["label"] == "(오픈,중도금대출)"

xb, hd = post("/api/jichul/xlsx", {"year": info["year"], "data": data}, raw=True)
print("xlsx:", len(xb), "bytes", hd.get("Content-Disposition", "")[:80])
assert len(xb) > 8000 and jichul_xlsx.read_inputs(xb)["data"] == data

folder = os.path.join(tempfile.gettempdir(), "jichul_api_test")
s = post("/api/jichul/save", {"year": info["year"], "data": data, "folder": folder, "filename": "api_test.xlsx"})
assert s.get("exists") or s["ok"], s
if s.get("exists"):
    s = post("/api/jichul/save", {"year": info["year"], "data": data, "folder": folder, "filename": "api_test.xlsx", "overwrite": True})
print("save:", s["path"])
ld = post("/api/jichul/load_path", {"path": s["path"]})
assert ld["data"] == data and ld["year"] == info["year"]
ld2 = upload("/api/jichul/load", s["path"])
assert ld2["data"] == data
print("load_path·load(업로드): 입력값 일치")
rc = get("/api/jichul/recent")
assert any(x["path"] == s["path"] for x in rc["sites"]), rc
print("recent:", [(x["name"], x["total"]) for x in rc["sites"][:3]])
try:
    post("/api/jichul/open", {"path": "C:\\Windows\\notepad.exe"})
    print("open 제한: 실패해야 하는데 통과함!")
    sys.exit(1)
except urllib.error.HTTPError as e:
    print("open 제한:", e.code, json.loads(e.read())["error"])
print("API 시험 통과")
