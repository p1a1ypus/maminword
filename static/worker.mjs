/* Конвертация DOCX <-> HTML прямо в браузере: тот же Python-код (docx2html.py / html2docx.py),
   запущенный в Pyodide. Используется, когда рядом нет сервера (например, на GitHub Pages). */
import { loadPyodide } from './vendor/pyodide/pyodide.mjs';

const ROOT = new URL('../', import.meta.url);
const PYODIDE = new URL('./vendor/pyodide/', import.meta.url).href;
const HOME = '/home/pyodide/mw';
let ready = null;
let booted = false;

function progress(text) { postMessage({ progress: text }); }

async function boot() {
  progress('Загружаю модуль Word — в первый раз это около 15 МБ…');
  const py = await loadPyodide({ indexURL: PYODIDE });
  progress('Подготавливаю модуль Word…');
  await py.loadPackage(['lxml', 'typing-extensions', PYODIDE + 'python_docx-1.2.0-py3-none-any.whl']);
  py.FS.mkdirTree(HOME + '/templates');
  for (const f of ['docx_common.py', 'docx2html.py', 'html2docx.py']) {
    const r = await fetch(new URL(f, ROOT));
    if (!r.ok) { throw new Error('не загрузился ' + f); }
    py.FS.writeFile(HOME + '/' + f, await r.text());
  }
  const tpl = await fetch(new URL('templates/blank.docx', ROOT));
  py.FS.writeFile(HOME + '/templates/blank.docx', new Uint8Array(await tpl.arrayBuffer()));
  py.runPython(`import sys\nsys.path.insert(0, '${HOME}')\nimport json, os, docx2html, html2docx`);
  progress('');
  booted = true;
  return py;
}

function shortError(err) {
  // у ошибок Python в сообщении весь traceback — показываем только последнюю строку
  const lines = String((err && err.message) || err).trim().split('\n');
  return lines[lines.length - 1].replace(/^[\w.]*(Error|Exception): /, '');
}

self.onmessage = async (e) => {
  const m = e.data;
  try {
    ready = ready || boot();
    const py = await ready;
    if (m.cmd === 'warmup') {
      postMessage({ id: m.id, result: true });
    } else if (m.cmd === 'open') {
      py.FS.writeFile('/tmp/in.docx', new Uint8Array(m.data));
      const out = py.runPython("json.dumps(docx2html.convert(open('/tmp/in.docx', 'rb').read()), ensure_ascii=False)");
      postMessage({ id: m.id, result: JSON.parse(out) });
    } else if (m.cmd === 'save') {
      py.FS.writeFile('/tmp/payload.json', JSON.stringify({ html: m.html, page: m.page }));
      if (m.original) { py.FS.writeFile('/tmp/orig.docx', new Uint8Array(m.original)); }
      else { try { py.FS.unlink('/tmp/orig.docx'); } catch (_) { /* не было */ } }
      const warns = py.runPython([
        "_p = json.load(open('/tmp/payload.json', encoding='utf-8'))",
        "_o = open('/tmp/orig.docx', 'rb').read() if os.path.exists('/tmp/orig.docx') else None",
        "_d, _w = html2docx.convert(_p, _o)",
        "open('/tmp/out.docx', 'wb').write(_d)",
        "json.dumps(_w, ensure_ascii=False)"
      ].join('\n'));
      const bytes = py.FS.readFile('/tmp/out.docx');
      const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
      postMessage({ id: m.id, result: { bytes: buf, warnings: JSON.parse(warns) } }, [buf]);
    }
  } catch (err) {
    if (!booted) { ready = null; } // загрузка не удалась (сеть) — следующая попытка начнёт заново
    postMessage({ id: m.id, error: shortError(err) });
  }
};
