"""월별 산식(xlcalc.Model.explain) 검사: 예시 3개 + 저장된 작업(_jobs, 2027) 입력으로 계산해
   0 이 아닌 월별 칸 중 산식이 만들어진 비율, 산식을 다시 계산한 값 ↔ 칸 값 일치, 줄별 산식 보기(--show)"""
import sys, os, json, glob
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # 이 저장소 폴더(외장하드 E: 또는 F:)
sys.path.insert(0, os.path.join(ROOT, "tool"))
sys.stdout.reconfigure(encoding="utf-8")
import warnings
warnings.filterwarnings("ignore")
import xlcalc, site_io

show = "--show" in sys.argv
model = xlcalc.Model(os.path.join(ROOT, "판관비 예산안 자동작성 양식_2027.xlsx"))

cases = []
for p in sorted(glob.glob(os.path.join(ROOT, "예시", "*.xlsx"))):
    d = site_io.read_inputs(p)
    d.pop("year", None)
    cases.append((os.path.basename(p)[:20], d))
seen = set()
for p in sorted(glob.glob(os.path.join(ROOT, "app", "_jobs", "*", "input.json")), key=os.path.getmtime, reverse=True):
    inp = json.load(open(p, encoding="utf-8"))
    if int(inp.get("year") or 0) != 2027:
        continue
    name = (inp["data"].get("fields") or {}).get("현장명") or "?"
    if name in seen:
        continue
    seen.add(name)
    cases.append((name, inp["data"]))

tot_cells = tot_fx = bad = 0
longest = ("", 0)
for name, data in cases:
    res = site_io.calc_model(model, data)
    fx = res.get("formulas") or []
    n_cells = n_fx = n_bad = 0
    for g, gf in zip(res["groups"], fx):
        for r, rf in zip(g["rows"], gf):
            if r["manual"]:
                continue
            for j, m in enumerate(r["months"]):
                if not m:
                    continue
                n_cells += 1
                f = rf[j] if rf else None
                if not f:
                    if show:
                        print(f"   (산식 없음) {r['g'][:50]} {j + 1}월 = {m:,.0f}")
                    continue
                n_fx += 1
                if len(f) > longest[1]:
                    longest = (f, len(f))
                v = xlcalc.Model._scalar(model, model._ev(xlcalc.Parser(f).parse(), res["sheet"], 1, 1), res["sheet"], 1, 1)
                if not isinstance(v, (int, float)) or abs(float(v) - float(m)) > 1e-6 * max(1, abs(float(m))):
                    n_bad += 1
                    print(f"   ✗ {r['g'][:50]} {j + 1}월: 산식 {f} = {v} ≠ {m}")
            if show and rf:
                sample = next(((j, f) for j, f in enumerate(rf) if f), None)
                if sample:
                    print(f"   {r['g'].strip()[:58]:<60} {sample[0] + 1:>2}월  ={sample[1]}")
    print(f"■ {name}: 0 아닌 월 칸 {n_cells}개 중 산식 {n_fx}개({100 * n_fx / max(1, n_cells):.0f}%), 불일치 {n_bad}개")
    tot_cells += n_cells
    tot_fx += n_fx
    bad += n_bad
print(f"합계: 칸 {tot_cells}, 산식 {tot_fx}({100 * tot_fx / max(1, tot_cells):.0f}%), 불일치 {bad}, 가장 긴 산식 {longest[1]}자: {longest[0][:120]}")
