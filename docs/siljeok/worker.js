// 사업관리비 실적계획 자동작성 웹판 — 브라우저 안에서 Python(Pyodide)을 띄워 엔진(siljeok.py)을 돌린다(모듈 워커). 파일은 밖으로 나가지 않는다.
// 예산안 웹판(../worker.js)과 같은 Pyodide 묶음을 쓰되 PDF 읽기 부품은 받지 않는다(lxml 과 openpyxl 휠만).
import { loadPyodide } from 'https://cdn.jsdelivr.net/npm/@bentopdf/pymupdf-wasm@0.11.16/assets/pyodide.js';
const VERSION = '2026.11.10.1508';
const PYODIDE = 'https://cdn.jsdelivr.net/npm/@bentopdf/pymupdf-wasm@0.11.16/assets/';
const WHEELS = ['../wheels/et_xmlfile-2.0.0-py3-none-any.whl', '../wheels/openpyxl-3.1.5-py2.py3-none-any.whl'];
const PY_FILES = [['../py/siljeok.py', 'siljeok.py'], ['../py/xlcalc.py', 'xlcalc.py'], ['../py/xlsxio.py', 'xlsxio.py'], ['bridge.py', 'bridge.py']];
const TEMPLATE = '../data/siljeok_form.xlsx';
const TEMPLATE_NAME = '사업관리비 실적계획 양식.xlsx';
const APP = '/home/pyodide/app';
const WORK = '/home/pyodide/work';

let pyodide = null;
let seq = 0;
const status = msg => postMessage({status: msg});
const log = msg => postMessage({log: msg});
const abs = p => new URL(p, import.meta.url).href;

async function fetchBytes(url) {
  const r = await fetch(url);
  if (!r.ok) throw new Error(`${url} 을(를) 받지 못했습니다 (${r.status})`);
  return new Uint8Array(await r.arrayBuffer());
}
async function fetchText(url) {
  const r = await fetch(url);
  if (!r.ok) throw new Error(`${url} 을(를) 받지 못했습니다 (${r.status})`);
  return await r.text();
}

async function init() {
  status('Python 준비 중… (처음 한 번은 20~40초, 그 뒤로는 빨라집니다)');
  pyodide = await loadPyodide({indexURL: PYODIDE});
  status('엑셀 읽기·쓰기 부품 받는 중… (처음 한 번만)');
  await pyodide.loadPackage(['lxml', ...WHEELS.map(abs)]);
  status('엔진·사내 양식 받는 중…');
  pyodide.FS.mkdirTree(APP);
  pyodide.FS.mkdirTree(WORK);
  for (const [src, name] of PY_FILES) pyodide.FS.writeFile(`${APP}/${name}`, await fetchText(abs(`${src}?v=${VERSION}`)));
  pyodide.FS.writeFile(`${WORK}/form.xlsx`, await fetchBytes(abs(`${TEMPLATE}?v=${VERSION}`)));
  await pyodide.runPythonAsync(`
import sys
sys.path.insert(0, '${APP}')
import bridge
`);
  const bridge = pyodide.globals.get('bridge');
  bridge.set_logger(log);
  const r = JSON.parse(bridge.init(`${WORK}/form.xlsx`, TEMPLATE_NAME));
  bridge.destroy();
  status('');
  postMessage({ready: true, year: r.year});
}
const ready = init().catch(e => { postMessage({fatal: String(e && e.message || e)}); throw e; });

function call(fn, ...args) {
  const bridge = pyodide.globals.get('bridge');
  try {
    const f = bridge[fn];
    const out = f(...args);
    f.destroy && f.destroy();
    return out;
  } finally { bridge.destroy(); }
}
function bytes(fn, ...args) {
  const py = call(fn, ...args);
  const out = py.toJs();
  py.destroy && py.destroy();
  return out;
}
function xlsx(fn, state) {
  const out = bytes(fn, JSON.stringify(state));
  if (!out || !out.length) return {result: {ok: false, error: call('last_error') || '엑셀을 만들지 못했습니다.'}, transfer: []};
  return {result: {ok: true, xlsx: out}, transfer: [out.buffer]};
}

onmessage = async (e) => {
  const {id, cmd, args} = e.data;
  try {
    await ready;
    let result, transfer = [];
    if (cmd === 'info') result = JSON.parse(call('info'));
    else if (cmd === 'parse') {
      const p = `${WORK}/up_${++seq}.xlsx`;
      pyodide.FS.writeFile(p, new Uint8Array(args.file));
      try { result = JSON.parse(call('parse', p, args.name)); } finally { try { pyodide.FS.unlink(p); } catch (err) {} }
    } else if (cmd === 'plan') result = JSON.parse(call('plan', JSON.stringify(args.state)));
    else if (cmd === 'form' || cmd === 'evidence') ({result, transfer} = xlsx(cmd, args.state));
    else throw new Error('알 수 없는 명령: ' + cmd);
    postMessage({id, result}, transfer);
  } catch (err) {
    postMessage({id, result: {ok: false, error: '처리 중 오류가 났습니다: ' + String(err && err.message || err).split('\n').slice(-3).join(' ')}});
  }
};
