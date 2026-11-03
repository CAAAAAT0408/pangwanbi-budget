# -*- coding: utf-8 -*-
"""브라우저(Pyodide 워커) 안에서 worker.js 가 부르는 다리. 서버(app/server.py)의 /api/… 와 같은 결과 구조를 돌려준다.

  init(template_path, form_path)        계산용 양식(xlcalc.Model)·회사 양식 읽기
  info() / schema() / params()          JSON
  calc(data_json) → JSON(res)           입력값 → 계산 결과 (서버 /api/calc 와 같음)
  report() → bytes                      마지막 계산 결과를 회사 양식에 채운 xlsx
  report_info() → JSON                  마지막 보고서 정보(예산 시트 이름)
  load(path, name) → JSON               저장했던 보고서 xlsx → 입력값
  types_start(path) / types_step(json)  설계개요 읽기(그림 PDF 는 OCR 단계를 JS 와 주고받음)
  types_image(i) → bytes                OCR 할 PNG
"""
import json
import os
import sys
import time
import traceback

import xlcalc
import site_io
import design_overview
import report_xl

STATE = dict(model=None, form=None, form_name="", template_name="", res=None, data=None, year=None,
             session=None, images=[], log=print)


def set_logger(fn):
    STATE["log"] = fn


def set_debug(flag):
    STATE["debug_ocr"] = bool(flag)


def _log(msg):
    try:
        STATE["log"](str(msg))
    except Exception:
        pass


def _err(e):
    tb = traceback.format_exc()
    _log(tb[-600:])
    return json.dumps(dict(ok=False, error=str(e)), ensure_ascii=False)


def init(template_path, form_path, template_name=None, form_name=None):
    t = time.time()
    m = xlcalc.Model(template_path)
    m.name = template_name or os.path.basename(template_path)
    STATE["model"] = m
    STATE["template_name"] = m.name
    with open(form_path, "rb") as f:
        STATE["form"] = f.read()
    STATE["form_name"] = form_name or os.path.basename(form_path)
    info = site_io.template_info(m)
    STATE["year"] = int(info.get("year") or 0) or None
    _log(f"양식 읽기 {time.time() - t:.1f}초, 수식 {sum(1 for s in m.cells.values() for v in s.values() if isinstance(v, tuple))}개")
    return json.dumps(dict(ok=True, year=STATE["year"]))


def info():
    try:
        m = STATE["model"]
        t = site_io.template_info(m)
        t.pop("params", None)
        t.pop("path", None)
        fy = None
        import re
        mm = re.search(r"(20\d{2})", STATE["form_name"])
        if mm:
            fy = int(mm.group(1))
        return json.dumps(dict(ok=True, local=False, web=True, budget_root="",
                               templates=[dict(year=t["year"], name=STATE["template_name"])], template=t,
                               form=dict(name=STATE["form_name"], year=fy)), ensure_ascii=False, default=str)
    except Exception as e:
        return _err(e)


def schema():
    try:
        return json.dumps(dict(ok=True, year=STATE["year"], **site_io.schema(STATE["model"])), ensure_ascii=False, default=str)
    except Exception as e:
        return _err(e)


def params():
    try:
        t = site_io.template_info(STATE["model"])
        t.pop("path", None)
        return json.dumps(dict(ok=True, **t), ensure_ascii=False, default=str)
    except Exception as e:
        return _err(e)


def calc(data_json):
    try:
        t = time.time()
        data = json.loads(data_json)
        res = site_io.calc_model(STATE["model"], data)
        STATE["res"], STATE["data"] = res, data
        res = dict(res)
        res["sec"] = round(time.time() - t, 1)
        res["job"] = "web"
        res["template_year"] = STATE["year"]
        res["default_sub"] = ""
        return json.dumps(dict(ok=True, **res), ensure_ascii=False, default=str)
    except Exception as e:
        return _err(e)


def report():
    """마지막 계산 결과 → 회사 양식 xlsx (bytes)"""
    out, meta = report_xl.build_report(STATE["form"], STATE["res"], STATE["data"])
    STATE["report_meta"] = meta
    return out


def report_info():
    try:
        meta = STATE.get("report_meta") or {}
        return json.dumps(dict(ok=True, form=STATE["form_name"], budget_sheet=meta.get("budget_sheet"), previews=[]),
                          ensure_ascii=False)
    except Exception as e:
        return _err(e)


def load(path, name):
    try:
        with open(path, "rb") as f:
            raw = f.read()
        obj = report_xl.read_inputs(raw)
        if obj is None:
            data = site_io.read_inputs(path)          # 자동작성 계산용 양식 파일
            year = data.pop("year", None)
        else:
            data, year = obj.get("data") or {}, obj.get("year")
        if not data.get("kinds"):
            data["kinds"] = site_io.guess_kinds(data.get("fields") or {}, year or STATE["year"])
        return json.dumps(dict(ok=True, data=data, year=year, name=name), ensure_ascii=False, default=str)
    except Exception as e:
        return _err(e)


def _types_pack(kind, payload):
    if kind == "done":
        return json.dumps(dict(ok=True, kind="done", **payload), ensure_ascii=False, default=str)
    STATE["images"] = payload
    return json.dumps(dict(ok=True, kind="ocr", n=len(payload)))


def types_start(path, name):
    try:
        s = design_overview.TypesSession(path, max_types=site_io.BT.NTYPES, log=_log)
        STATE["session"] = s
        kind, payload = s.start()
        return _types_pack(kind, payload)
    except Exception as e:
        return _err(e)


def types_step(results_json):
    try:
        results = json.loads(results_json)
        if STATE.get("debug_ocr"):
            for i, r in enumerate(results):
                _log(f"OCR words[{i}] {r.get('w')}x{r.get('h')}: " + json.dumps(r.get("words", [])[:160], ensure_ascii=False))
        kind, payload = STATE["session"].step(results)
        return _types_pack(kind, payload)
    except Exception as e:
        return _err(e)


def types_image(i):
    return STATE["images"][int(i)]
