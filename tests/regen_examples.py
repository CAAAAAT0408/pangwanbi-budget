"""예시 3개를 회사 양식 결과물로 다시 만들기(서버 API 사용) + 저장 기능 시험"""
import sys, json, os, urllib.request, urllib.parse, shutil, glob
sys.stdout.reconfigure(encoding="utf-8")
B = "http://127.0.0.1:8792"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # 이 저장소 폴더(외장하드 E: 또는 F:)


def get(p, raw=False):
    with urllib.request.urlopen(B + p, timeout=300) as r:
        return r.read() if raw else json.loads(r.read())


def post(p, obj):
    req = urllib.request.Request(B + p, data=json.dumps(obj).encode(), headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=300) as r:
        return json.loads(r.read())


reg = []
for path in sorted(glob.glob(os.path.join(ROOT, "예시", "*.xlsx"))):
    name = os.path.basename(path)
    ld = post("/api/load_path", {"path": path})
    data = ld["data"]
    print(name, "구분", data.get("kinds"), "연도", ld["year"])
    r = post("/api/calc", {"year": 2027, "data": data})
    blob = get(f"/api/download?job={r['job']}&name=" + urllib.parse.quote(name), raw=True)
    tmp = path + ".new"
    open(tmp, "wb").write(blob)
    os.replace(tmp, path)
    total = round(r["total"]["total"] / 1000)
    print("   → 회사 양식으로 다시 만듦", len(blob), "bytes, 합계", f"{total:,}천원", "오류", r["nerrors"])
    reg.append(dict(name=(data.get("fields") or {}).get("현장명") or name, path=path, year=2027, total=total, saved="2026-11-02 00:00"))
    last_job = r["job"]

# 다시 불러오기(회사 양식 파일 → 입력값) 확인
ld = post("/api/load_path", {"path": reg[0]["path"]})
print("회사 양식 파일에서 다시 불러오기:", ld["data"]["fields"]["현장명"], ld["data"].get("kinds"), "타입", len(ld["data"]["types"]))

# 저장 시험(시험 폴더 만들고 지움)
sv = post("/api/save", {"job": last_job, "sub": "", "filename": "(저장시험_삭제예정) 2027년 사업비 예산편성.xlsx", "site": "저장시험_삭제예정", "total": 1})
print("저장:", sv.get("path"))
import site
site.addsitedir(site.getusersitepackages())
import openpyxl
wb = openpyxl.load_workbook(sv["path"], read_only=True)
print("   저장 파일 시트:", wb.sheetnames)
wb.close()
shutil.rmtree(sv["folder"])
print("   시험 폴더 지움:", not os.path.exists(sv["folder"]))
with open(os.path.join(ROOT, "app", "_sites.json"), "w", encoding="utf-8") as f:
    json.dump(reg, f, ensure_ascii=False, indent=1)
print("저장한 현장 목록 갱신:", [x["name"] for x in reg])
