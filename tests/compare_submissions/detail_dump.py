# -*- coding: utf-8 -*-
"""compare_results_<mode>.json → 현장별 항목·줄 차이 상세 출력"""
import sys, os, json
sys.stdout.reconfigure(encoding="utf-8")
SP = os.path.dirname(os.path.abspath(__file__))
mode = sys.argv[1] if len(sys.argv) > 1 else "adjusted"
sel = sys.argv[2:]
res = json.load(open(os.path.join(SP, f"compare_results_{mode}.json"), encoding="utf-8"))
for r in res:
    if sel and not any(s in r["key"] for s in sel):
        continue
    print(f"\n===== [{r['year']}] {r['file']}  구분 {r['kinds']}  실제 {r['actual_total']/1000:,.0f} 도구 {r['tool_total']/1000:,.0f} ({(r['ratio'] or 0)*100:.1f}%)")
    f = r["inputs"]["fields"]
    print("  입력:", {k: v for k, v in f.items() if v not in (None, [None, None], "", 0, "N") and k not in ("현장명", "지역", "사업구분", "상품")})
    print("  타입:", r["inputs"]["types"][:6], "분양:", r["inputs"]["monthly"].get("분양"), "입주:", r["inputs"]["monthly"].get("입주"))
    for n in r["notes"]:
        print("  note:", n)
    print("  단계:", r["stage"])
    for it in r["items"]:
        flag = "  " if abs(it["diff"]) <= 1000 else "**"
        print(f"  {flag} {it['item']:<8} 실제 {it['actual']/1000:>13,.0f} 도구 {it['tool']/1000:>13,.0f} 차이 {it['diff']/1000:>+13,.0f}")
    for ln in r["lines"]:
        if ln["class"] == "일치":
            continue
        print(f"      · {ln['item'][:6]:<6} {ln['key']:<18} 실제 {ln['actual']/1000:>12,.0f} 도구 {ln['tool']/1000:>12,.0f} [{ln['class']}] {ln['actual_g'][:70]} ⇔ {ln['tool_g'][:70]}")
