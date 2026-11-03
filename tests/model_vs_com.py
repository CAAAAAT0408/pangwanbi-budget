"""calc_model(xlcalc) 결과 ↔ 작업 폴더의 result.json(Excel COM 결과) 비교"""
import sys, os, json, glob, time
sys.path.insert(0, r"F:\1. 우미\1. 본사업무\★ 판관비 예산안\☆ 예산안 자동작성\tool")
sys.stdout.reconfigure(encoding="utf-8")
import warnings
warnings.filterwarnings("ignore")
import xlcalc, site_io

ROOT = r"F:\1. 우미\1. 본사업무\★ 판관비 예산안\☆ 예산안 자동작성"
t = time.time()
model = xlcalc.Model(os.path.join(ROOT, "판관비 예산안 자동작성 양식_2027.xlsx"))
print(f"양식 읽기 {time.time() - t:.2f}초")


def same(a, b, path=""):
    if isinstance(a, dict) and isinstance(b, dict):
        out = []
        for k in set(a) | set(b):
            out += same(a.get(k), b.get(k), f"{path}.{k}")
        return out
    if isinstance(a, list) and isinstance(b, list):
        if len(a) != len(b):
            return [f"{path}: 길이 {len(a)} vs {len(b)}"]
        out = []
        for i, (x, y) in enumerate(zip(a, b)):
            out += same(x, y, f"{path}[{i}]")
        return out
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) and not isinstance(b, bool):
        return [] if abs(a - b) <= 1e-6 * max(1, abs(b)) else [f"{path}: {a} vs {b}"]
    if (a in (None, "", 0) and b in (None, "", 0)):
        return []
    return [] if a == b else [f"{path}: {a!r} vs {b!r}"]


jobs = sorted(glob.glob(os.path.join(ROOT, "app", "_jobs", "*", "result.json")), key=os.path.getmtime)
n = 0
for rj in jobs:
    d = os.path.dirname(rj)
    ij = os.path.join(d, "input.json")
    if not os.path.exists(ij):
        continue
    inp = json.load(open(ij, encoding="utf-8"))
    if int(inp.get("year") or 0) != 2027:
        continue
    exp = json.load(open(rj, encoding="utf-8"))
    t = time.time()
    got = site_io.calc_model(model, inp["data"])
    dt = time.time() - t
    exp.pop("sec", None); exp.pop("job", None); exp.pop("template_year", None); exp.pop("default_sub", None)
    diffs = same(got, exp)
    name = (inp["data"].get("fields") or {}).get("현장명")
    print(f"■ {os.path.basename(d)} {name}: 차이 {len(diffs)}개, {dt:.2f}초, 합계 {got['total']['total']:,.0f} vs {exp['total']['total']:,.0f}")
    for x in diffs[:8]:
        print("   ", x)
    n += 1
print("비교한 작업", n)
