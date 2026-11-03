# -*- coding: utf-8 -*-
"""브라우저(Pyodide 워커) 안에서 worker.js 가 부르는 다리 — 서버(app/server.py)의 /api/jichul/… 와 같은 결과 구조.

  init(kijun_json_path)        기준단가(docs/data/kijun.json, tool/sync_web.py 가 양식에서 뽑아 둠) 읽기
  info() / schema() / params() JSON
  calc(data_json) → JSON(res)  입력값 → 지출계획서 계산 결과
  xlsx(data_json) → bytes      입력값 → 지출계획서 xlsx (다시 불러올 입력값 포함)
  load(path, name) → JSON      저장했던 xlsx → 입력값
"""
import json
import traceback

import jichul
import jichul_xlsx

STATE = dict(K={}, meta={}, log=print)


def set_logger(fn):
    STATE["log"] = fn


def _err(e):
    tb = traceback.format_exc()
    try:
        STATE["log"](tb[-600:])
    except Exception:    # noqa: BLE001
        pass
    return json.dumps(dict(ok=False, error=str(e)), ensure_ascii=False)


def init(kijun_json_path):
    with open(kijun_json_path, encoding="utf-8") as f:
        meta = json.load(f)
    STATE["meta"] = meta
    STATE["K"] = {p["key"]: p["value"] for p in meta.get("params") or [] if "key" in p}
    return json.dumps(dict(ok=True, year=meta.get("year")))


def info():
    try:
        m = STATE["meta"]
        return json.dumps(dict(ok=True, local=False, web=True, year=m.get("year"), kijun=m.get("kijun"), written=m.get("written"),
                               template=m.get("template"), count=len(STATE["K"]),
                               templates=[dict(year=m.get("year"), name=m.get("template"))]), ensure_ascii=False, default=str)
    except Exception as e:   # noqa: BLE001
        return _err(e)


def schema():
    try:
        return json.dumps(dict(ok=True, year=STATE["meta"].get("year"), **jichul.schema(STATE["K"])), ensure_ascii=False, default=str)
    except Exception as e:   # noqa: BLE001
        return _err(e)


def params():
    try:
        m = STATE["meta"]
        return json.dumps(dict(ok=True, year=m.get("year"), kijun=m.get("kijun"), written=m.get("written"), params=m.get("params") or []),
                          ensure_ascii=False, default=str)
    except Exception as e:   # noqa: BLE001
        return _err(e)


def calc(data_json):
    try:
        data = json.loads(data_json)
        res = jichul.calc(data, STATE["K"])
        res["file_name"] = jichul.file_name(res)
        res["year"] = STATE["meta"].get("year")
        return json.dumps(dict(ok=True, **res), ensure_ascii=False, default=str)
    except Exception as e:   # noqa: BLE001
        return _err(e)


def xlsx(data_json):
    data = json.loads(data_json)
    res = jichul.calc(data, STATE["K"])
    return jichul_xlsx.build_xlsx(res, data, STATE["meta"].get("year"))


def load(path, name):
    try:
        with open(path, "rb") as f:
            raw = f.read()
        obj = jichul_xlsx.read_inputs(raw)
        if not obj:
            raise ValueError("지출계획서 자동작성으로 만든 파일이 아닙니다(입력값이 들어 있지 않음).")
        return json.dumps(dict(ok=True, data=obj.get("data") or {}, year=obj.get("year"), name=name), ensure_ascii=False, default=str)
    except Exception as e:   # noqa: BLE001
        return _err(e)
