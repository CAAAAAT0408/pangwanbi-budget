// 지출계획서 자동작성 웹판 — 브라우저 안에서 Python(Pyodide)을 띄워 엔진(jichul.py)을 돌린다(모듈 워커). 파일은 밖으로 나가지 않는다.
// 예산안 웹판(../worker.js)과 같은 Pyodide 묶음을 쓰되 PDF 읽기 부품은 받지 않아 더 가볍다(openpyxl 휠만).
import { loadPyodide } from 'https://cdn.jsdelivr.net/npm/@bentopdf/pymupdf-wasm@0.11.16/assets/pyodide.js';
const VERSION = '2026.11.04.1422';
const PYODIDE = 'https://cdn.jsdelivr.net/npm/@bentopdf/pymupdf-wasm@0.11.16/assets/';
const WHEELS = ['../wheels/et_xmlfile-2.0.0-py3-none-any.whl', '../wheels/openpyxl-3.1.5-py2.py3-none-any.whl'];
const PY_FILES = [['../py/jichul.py', 'jichul.py'], ['../py/jichul_xlsx.py', 'jichul_xlsx.py'], ['../py/params.py', 'params.py'],
  ['bridge.py', 'bridge.py']];
const KIJUN = '../data/kijun.json';
const APP = '/home/pyodide/app';
const WORK = '/home/pyodide/work';

let pyodide = null;
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
  status('엑셀 작성 부품 받는 중… (처음 한 번만)');
  await pyodide.loadPackage(WHEELS.map(abs));
  status('엔진 받는 중…');
  pyodide.FS.mkdirTree(APP);
  pyodide.FS.mkdirTree(WORK);
  for (const [src, name] of PY_FILES) pyodide.FS.writeFile(`${APP}/${name}`, await fetchText(abs(`${src}?v=${VERSION}`)));
  pyodide.FS.writeFile(`${WORK}/kijun.json`, await fetchText(abs(`${KIJUN}?v=${VERSION}`)));
  await pyodide.runPythonAsync(`
import sys
sys.path.insert(0, '${APP}')
import bridge
`);
  const bridge = pyodide.globals.get('bridge');
  bridge.set_logger(log);
  const r = JSON.parse(bridge.init(`${WORK}/kijun.json`));
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

onmessage = async (e) => {
  const {id, cmd, args} = e.data;
  try {
    await ready;
    let result, transfer = [];
    if (cmd === 'info') result = JSON.parse(call('info'));
    else if (cmd === 'schema') result = JSON.parse(call('schema'));
    else if (cmd === 'params') result = JSON.parse(call('params'));
    else if (cmd === 'calc') result = JSON.parse(call('calc', JSON.stringify(args.data)));
    else if (cmd === 'xlsx') {
      const xlsx = bytes('xlsx', JSON.stringify(args.data));
      result = {ok: true, xlsx};
      transfer = [xlsx.buffer];
    } else if (cmd === 'load') {
      const p = `${WORK}/load_` + (args.name || 'x.xlsx').replace(/[^\w.가-힣-]/g, '_');
      pyodide.FS.writeFile(p, new Uint8Array(args.file));
      result = JSON.parse(call('load', p, args.name));
    } else {
      throw new Error('알 수 없는 명령: ' + cmd);
    }
    postMessage({id, result}, transfer);
  } catch (err) {
    postMessage({id, result: {ok: false, error: '처리 중 오류가 났습니다: ' + String(err && err.message || err).split('\n').slice(-3).join(' ')}});
  }
};
