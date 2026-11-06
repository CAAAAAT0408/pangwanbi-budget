// 브라우저 안에서 Python(Pyodide)을 띄워 판관비 예산안 엔진을 돌린다 (모듈 워커). 파일은 브라우저 밖으로 나가지 않는다.
import { loadPyodide } from 'https://cdn.jsdelivr.net/npm/@bentopdf/pymupdf-wasm@0.11.16/assets/pyodide.js';
const VERSION = '2026.11.06.1500';
// PyMuPDF(설계개요 PDF 읽기)가 들어 있는 Pyodide 묶음 (Python 3.13)
const PYODIDE = 'https://cdn.jsdelivr.net/npm/@bentopdf/pymupdf-wasm@0.11.16/assets/';
const PYMUPDF = PYODIDE + 'pymupdf-1.26.3-cp313-none-pyodide_2025_0_wasm32.whl';
// openpyxl·xlrd 는 묶음에 없거나(openpyxl) 묶음의 검증값이 틀려(xlrd) 직접 둔 휠을 쓴다
const WHEELS = ['wheels/et_xmlfile-2.0.0-py3-none-any.whl', 'wheels/openpyxl-3.1.5-py2.py3-none-any.whl',
  'wheels/xlrd-2.0.1-py2.py3-none-any.whl'];
const PY_FILES = ['xlcalc.py', 'params.py', 'kijun_parser.py', 'build_template.py', 'site_io.py', 'xlsxio.py',
  'report.py', 'report_xl.py', 'design_overview.py', 'bridge.py'];
const DATA = ['template.xlsx', 'form.xlsx'];      // 원래 파일 이름·연도는 data/manifest.json 에
const APP = '/home/pyodide/app';
const WORK = '/home/pyodide/work';
const TESS = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.esm.min.js';
const TESS_WORKER = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/worker.min.js';
const TESS_CORE = 'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.1.1';

let pyodide = null;
let tess = null;
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
  status('Python 준비 중… (처음 한 번은 30초~1분, 그 뒤로는 빨라집니다)');
  pyodide = await loadPyodide({indexURL: PYODIDE});
  status('엑셀·PDF 읽기 부품 받는 중… (약 25MB, 처음 한 번만)');
  await pyodide.loadPackage(['lxml', PYMUPDF, ...WHEELS.map(abs)]);
  status('엔진 받는 중…');
  pyodide.FS.mkdirTree(APP);
  pyodide.FS.mkdirTree(WORK);
  for (const f of PY_FILES) pyodide.FS.writeFile(`${APP}/${f}`, await fetchText(abs(`py/${f}?v=${VERSION}`)));
  for (const f of DATA) pyodide.FS.writeFile(`${WORK}/${f}`, await fetchBytes(abs(`data/${f}?v=${VERSION}`)));
  let man = {};
  try { man = JSON.parse(await fetchText(abs(`data/manifest.json?v=${VERSION}`))); } catch (e) {}
  await pyodide.runPythonAsync(`
import sys
sys.path.insert(0, '${APP}')
import bridge
`);
  const bridge = pyodide.globals.get('bridge');
  bridge.set_logger(log);
  status('계산용 양식 읽는 중…');
  const r = JSON.parse(bridge.init(`${WORK}/${DATA[0]}`, `${WORK}/${DATA[1]}`, man.template || DATA[0], man.form || DATA[1]));
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
function writeFile(name, data) {
  const p = `${WORK}/${name}`;
  pyodide.FS.writeFile(p, new Uint8Array(data));
  return p;
}

// ── 그림 PDF 글자 인식 (Tesseract, 한국어+영어) ────────────────────────
function pngSize(u8) {
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  return {w: dv.getUint32(16), h: dv.getUint32(20)};
}
async function ocrWorker() {
  if (tess) return tess;
  status('글자 인식 부품 받는 중… (처음 한 번, 약 15MB)');
  const mod = await import(TESS);
  const T = mod.default || mod;
  const w = await T.createWorker(['kor', 'eng'], 1, {workerPath: TESS_WORKER, corePath: TESS_CORE});
  await w.setParameters({tessedit_pageseg_mode: '11', preserve_interword_spaces: '1'});
  tess = w;
  return w;
}
async function ocrImages(n) {
  const w = await ocrWorker();
  const out = [];
  for (let i = 0; i < n; i++) {
    status(`글자 인식 중… (${i + 1}/${n})`);
    const png = bytes('types_image', i);
    const {w: W, h: H} = pngSize(png);
    const blob = new Blob([png], {type: 'image/png'});
    const r = await w.recognize(blob);
    const words = [];
    for (const wd of (r.data.words || [])) {
      const t = (wd.text || '').trim();
      if (!t) continue;
      const b = wd.bbox;
      words.push([b.x0, b.y0, b.x1 - b.x0, b.y1 - b.y0, t]);
    }
    out.push({w: W, h: H, words});
  }
  status('');
  return out;
}
async function readTypes(path, name) {
  let r = JSON.parse(call('types_start', path, name));
  while (r.ok && r.kind === 'ocr') {
    const results = await ocrImages(r.n);
    r = JSON.parse(call('types_step', JSON.stringify(results)));
  }
  return r;
}

onmessage = async (e) => {
  const {id, cmd, args} = e.data;
  try {
    await ready;
    let result, transfer = [];
    if (cmd === 'debug') { call('set_debug', !!(args && args.on)); result = {ok: true}; }
    else if (cmd === 'info') result = JSON.parse(call('info'));
    else if (cmd === 'schema') result = JSON.parse(call('schema'));
    else if (cmd === 'params') result = JSON.parse(call('params'));
    else if (cmd === 'calc') result = JSON.parse(call('calc', JSON.stringify(args.data)));
    else if (cmd === 'report') {
      const xlsx = bytes('report');
      result = JSON.parse(call('report_info'));
      result.xlsx = xlsx;
      transfer = [xlsx.buffer];
    } else if (cmd === 'load') {
      const p = writeFile('load_' + (args.name || 'x.xlsx').replace(/[^\w.가-힣-]/g, '_'), args.file);
      result = JSON.parse(call('load', p, args.name));
    } else if (cmd === 'types') {
      const p = writeFile('overview' + (args.name.match(/\.[^.]+$/) || ['.pdf'])[0].toLowerCase(), args.file);
      result = await readTypes(p, args.name);
    } else {
      throw new Error('알 수 없는 명령: ' + cmd);
    }
    postMessage({id, result}, transfer);
  } catch (err) {
    postMessage({id, result: {ok: false, error: '처리 중 오류가 났습니다: ' + String(err && err.message || err).split('\n').slice(-3).join(' ')}});
  }
};
