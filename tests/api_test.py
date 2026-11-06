"""서버 API 시험: schema(20줄·main), /api/types(엑셀·PDF), 예시 재계산(회귀), 타입 14개 계산+보고서"""
import json, sys, os, time, urllib.request, uuid, glob
sys.stdout.reconfigure(encoding="utf-8")
B = "http://127.0.0.1:8792"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # 이 저장소 폴더(외장하드 E: 또는 F:)
DRIVE = os.path.splitdrive(ROOT)[0]                                   # 외장하드 드라이브 글자 — 다른 현장 자료 경로에 씀


def get(p):
    with urllib.request.urlopen(B + p, timeout=300) as r:
        return json.loads(r.read())


def post(p, obj):
    req = urllib.request.Request(B + p, data=json.dumps(obj).encode(), headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=300) as r:
        return json.loads(r.read())


def upload(p, path):
    bnd = uuid.uuid4().hex
    name = os.path.basename(path)
    body = (f"--{bnd}\r\nContent-Disposition: form-data; name=\"file\"; filename*=UTF-8''{urllib.parse.quote(name)}\r\n"
            f"Content-Type: application/octet-stream\r\n\r\n").encode() + open(path, "rb").read() + f"\r\n--{bnd}--\r\n".encode()
    req = urllib.request.Request(B + p, data=body, headers={"Content-Type": f"multipart/form-data; boundary={bnd}"}, method="POST")
    with urllib.request.urlopen(req, timeout=300) as r:
        return json.loads(r.read())


import urllib.parse
sc = get("/api/schema?year=2027")
mains = [f["key"] for s in sc["sections"] for f in s["fields"] if f.get("main")]
print("schema ntypes", sc["ntypes"], "main 칸", len(mains), mains)

for f in [DRIVE + r"\1. 우미\2. 수도권남부\14. 평택고덕 패키지형 공모사업\★ 1차(A36BL)\설계\250707_설계개요\[A-66BL] AG-014~015 설계개요,동별개요.xlsx",
          DRIVE + r"\1. 우미\2. 수도권남부\14. 평택고덕 패키지형 공모사업\★ 2차(A33BL)\설계\03. 2차 변경승인 완료 (26.09.08)\03. 2차 변경승인 완료 (26.09.08)\02. 도면\1. 건축\01 표제부\(변경후) AG-014~015 설계개요, 동별개요.pdf"]:
    t = time.time()
    j = upload("/api/types", f)
    print(f"/api/types {os.path.basename(f)} → {j['source']} {j['where']} {j['sec']}초(왕복 {time.time()-t:.1f}) 신뢰 {j['confidence']} "
          f"타입 {len(j['types'])} 합계 {j['total_units']} 일치 {j['units_ok']}/{j['supply_ok']}")
    print("   ", [(x["name"], x["units"], x["supply"]) for x in j["types"]][:4], "…")

# 예시 회귀
exp = {"(예시1)": 6106656, "(예시2)": 1566721, "(예시3)": 3381545}
ex1 = None
for path in sorted(glob.glob(os.path.join(ROOT, "예시", "*.xlsx"))):
    ld = post("/api/load_path", {"path": path})
    r = post("/api/calc", {"year": 2027, "data": ld["data"]})
    tot = round(r["total"]["total"] / 1000)
    key = os.path.basename(path)[:5]
    print(f"예시 {os.path.basename(path)[:30]} 합계 {tot:,} (기대 {exp.get(key, '?'):,}) 오류 {r['nerrors']} {'OK' if exp.get(key) == tot else 'DIFF'}")
    if key == "(예시1)":
        ex1 = ld["data"]

# 타입 14개(용마터널) 계산 + 회사 양식 보고서
yj = upload("/api/types", DRIVE + r"\1. 우미\2. 수도권남부\18. 용마터널 공공주택 복합사업\설계\260804_설계개요.xlsx")
d = json.loads(json.dumps(ex1))
d["types"] = [[t["name"], t["units"], t["supply"]] for t in yj["types"] if t["use"]]
d["fields"]["현장명"] = "타입14개 시험"
r = post("/api/calc", {"year": 2027, "data": d})
print("타입", len(d["types"]), "개 계산: 오류", r["nerrors"], "합계", f"{r['total']['total']/1000:,.0f}천원", r["errors"][:3])
t = time.time()
m = post("/api/report", {"job": r["job"]})
print("보고서", m.get("previews"), f"{time.time()-t:.1f}초", "ERR" if not m.get("ok") else "")
