# -*- coding: utf-8 -*-
"""브라우저(Pyodide 워커) 안에서 worker.js 가 부르는 다리 — 서버(app/server.py)의 /api/siljeok/… 와 같은 결과 구조.

  init(template_path, template_name)   사내 양식 틀(docs/data/siljeok_form.xlsx, tool/sync_web.py 가 회사양식\\실적계획 에서 복사) 읽기
  info() → JSON
  parse(path, name) → JSON             올린 엑셀(예산안 또는 사용내역) 읽기
  plan(state_json) → JSON              현장별 양식 값(미리보기)
  form(state_json) → bytes             채운 사내 양식 xlsx      (오류면 빈 bytes, last_error() 로 사유)
  evidence(state_json) → bytes         실적 근거자료 xlsx       (오류면 빈 bytes, last_error() 로 사유)
"""
import json
import traceback

import siljeok

STATE = dict(tpl=b"", tpl_name="", year=None, error="", log=print)


def set_logger(fn):
    STATE["log"] = fn


def _fail(e):
    if not isinstance(e, siljeok.SiljeokError):
        tb = traceback.format_exc()
        try:
            STATE["log"](tb[-800:])
        except Exception:    # noqa: BLE001
            pass
        return "처리 중 오류가 났습니다: " + str(e)
    return str(e)


def init(template_path, template_name):
    with open(template_path, "rb") as f:
        STATE["tpl"] = f.read()
    STATE["tpl_name"] = template_name
    STATE["year"] = siljeok.template_year(STATE["tpl"])
    return json.dumps(dict(ok=True, year=STATE["year"]))


def info():
    return json.dumps(dict(ok=True, local=False, web=True, template=STATE["tpl_name"], base_year=STATE["year"]), ensure_ascii=False)


def parse(path, name):
    try:
        with open(path, "rb") as f:
            data = f.read()
        return json.dumps(dict(ok=True, **siljeok.parse_file(data, name)), ensure_ascii=False, default=str)
    except Exception as e:   # noqa: BLE001
        return json.dumps(dict(ok=False, error=_fail(e)), ensure_ascii=False)


def plan(state_json):
    try:
        return json.dumps(siljeok.slim(siljeok.plan(json.loads(state_json))), ensure_ascii=False, default=str)
    except Exception as e:   # noqa: BLE001
        return json.dumps(dict(ok=False, error=_fail(e)), ensure_ascii=False)


def form(state_json):
    STATE["error"] = ""
    try:
        return siljeok.build_form(STATE["tpl"], json.loads(state_json))
    except Exception as e:   # noqa: BLE001
        STATE["error"] = _fail(e)
        return b""


def evidence(state_json):
    STATE["error"] = ""
    try:
        return siljeok.build_evidence(json.loads(state_json))
    except Exception as e:   # noqa: BLE001
        STATE["error"] = _fail(e)
        return b""


def last_error():
    return STATE["error"]
