// Check EDI Various Forms — รันโปรแกรม Python ตัวเดียวกับบนเครื่อง (scripts/bundle_edi.py + build_edi.py)
// ในเบราว์เซอร์ด้วย Pyodide; หน้าสแกน OCR ด้วย Tesseract.js (แทน OCR ของ Windows)
'use strict';
(() => {
  const $ = (id) => document.getElementById(id);
  const files = { mf: null, en: null };
  let py = null;          // Pyodide
  let reportBytes = null;

  const setStatus = (t, pct) => {
    $('status').textContent = t;
    if (pct !== undefined) $('bar').style.width = Math.max(0, Math.min(100, pct)) + '%';
  };
  const log = (t) => { const el = $('log'); el.classList.remove('hidden'); el.textContent += t + '\n'; el.scrollTop = el.scrollHeight; };

  function wireDrop(dropId, inputId, key, fnId) {
    const drop = $(dropId), input = $(inputId);
    const take = (f) => { if (!f) return; files[key] = f; $(fnId).textContent = f.name; ready(); };
    drop.addEventListener('click', () => input.click());
    drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    input.addEventListener('change', () => take(input.files[0]));
    drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('over'));
    drop.addEventListener('drop', (e) => {
      e.preventDefault(); drop.classList.remove('over');
      const list = [...e.dataTransfer.files];
      // วางทั้งสองไฟล์พร้อมกันก็ได้: .pdf ที่ชื่อมี ENTER = ENTER, .xls/.xlsx = MF
      if (list.length > 1) {
        for (const f of list) {
          const n = f.name.toLowerCase();
          if (/\.xlsx?$/.test(n) || (/\.pdf$/.test(n) && /(^|[^a-z])(mf|manifest)/.test(n))) { files.mf = f; $('fnMF').textContent = f.name; }
          else if (/\.pdf$/.test(n)) { files.en = f; $('fnEN').textContent = f.name; }
        }
        ready();
      } else take(list[0]);
    });
  }
  function ready() {
    reportBytes = null; $('dl').classList.add('hidden');
    $('run').disabled = !(files.mf && files.en);
    if (files.mf && files.en) setStatus('พร้อมตรวจ — กด "ตรวจสอบ"', 0);
  }

  const shedArgs = () => {
    const four = (v) => { const m = (v || '').match(/\d{3,4}/); return m ? m[0].padStart(4, '0') : ''; };
    const mode = document.querySelector('input[name=shedmode]:checked').value;
    if (mode === 'one') { const s = four($('shed').value); return s ? ['--shed', s] : []; }
    const a = four($('shedLcl').value), b = four($('shedCy').value);
    return (a || b) ? ['--shed-lcl', a, '--shed-other', b] : [];
  };

  async function boot() {
    if (py) return py;
    setStatus('กำลังโหลดตัวประมวลผล Python (ครั้งแรกใช้เวลาสักครู่)…', 5);
    py = await loadPyodide();
    setStatus('กำลังโหลดไลบรารีอ่าน PDF / Excel…', 15);
    await py.loadPackage(['pymupdf', 'xlrd', 'micropip']);
    await py.runPythonAsync("import micropip\nawait micropip.install(['openpyxl'])");
    setStatus('กำลังโหลดโปรแกรมตรวจ…', 25);
    py.FS.mkdirTree('/work/scripts');
    for (const f of ['bundle_edi.py', 'build_edi.py']) {
      const r = await fetch('scripts/' + f + '?v=' + Date.now(), { cache: 'no-cache' });
      if (!r.ok) throw new Error('โหลด ' + f + ' ไม่ได้');
      py.FS.writeFile('/work/scripts/' + f, new Uint8Array(await r.arrayBuffer()));
    }
    await py.runPythonAsync(`
import sys, io, os
sys.path.insert(0, '/work/scripts')
import bundle_edi as B
def _run(argv):
    buf = io.StringIO(); old = sys.stdout; sys.stdout = buf
    try:
        try:
            B.main(list(argv))
        except SystemExit as e:
            if e.code not in (None, 0): print(e)
    finally:
        sys.stdout = old
    return buf.getvalue()
def _png(path, i):
    import pymupdf
    d = pymupdf.open(path)
    return d[i].get_pixmap(dpi=300, colorspace=pymupdf.csGRAY).tobytes('png')
`);
    return py;
  }

  // Tesseract → รูปแบบเดียวกับ win_ocr.ps1: [{x,y,x1,y1,t}] พิกัดพิกเซลของภาพ 300 dpi
  // Tesseract รวมข้อความข้ามคอลัมน์เป็นบรรทัดเดียว ('MSC CO., LTD. : HASLK…') — OCR ของ Windows แยกเป็นชิ้นตามช่องว่าง
  // จึงตัดบรรทัดเมื่อช่องว่างระหว่างคำกว้างกว่าความสูงตัวอักษร และตัด ':' / '|' เดี่ยว (เส้นประแบ่งช่องของฟอร์ม)
  function tessLines(data) {
    const out = [];
    const clean = (t) => t.replace(/[©®]/g, '0').replace(/\s+/g, ' ').trim();
    const emit = (ws) => {
      const t = clean(ws.map((w) => w.text).join(' '));
      if (!t) return;
      out.push({ x: Math.min(...ws.map((w) => w.bbox.x0)), y: Math.min(...ws.map((w) => w.bbox.y0)),
                 x1: Math.max(...ws.map((w) => w.bbox.x1)), y1: Math.max(...ws.map((w) => w.bbox.y1)), t });
    };
    const line = (l) => {
      const words = (l.words || []).filter((w) => (w.text || '').trim());
      if (!words.length) { const t = clean(l.text || ''); if (t) out.push({ x: l.bbox.x0, y: l.bbox.y0, x1: l.bbox.x1, y1: l.bbox.y1, t }); return; }
      const h = Math.max(8, (l.bbox.y1 - l.bbox.y0));
      let seg = [];
      for (const w of words) {
        const sep = /^[:|¦!]+$/.test(w.text.trim());
        const prev = seg[seg.length - 1];
        if (sep || (prev && w.bbox.x0 - prev.bbox.x1 > 1.1 * h)) { if (seg.length) emit(seg); seg = []; }
        if (!sep) seg.push(w);
      }
      if (seg.length) emit(seg);
    };
    if (data.blocks) for (const b of data.blocks) for (const p of (b.paragraphs || [])) for (const l of (p.lines || [])) line(l);
    else if (data.lines) data.lines.forEach(line);
    return out;
  }

  async function check() {
    $('run').disabled = true; $('dl').classList.add('hidden'); $('log').textContent = ''; reportBytes = null;
    try {
      await boot();
      setStatus('กำลังอ่านไฟล์…', 30);
      const mfName = 'MF' + files.mf.name.slice(files.mf.name.lastIndexOf('.')).toLowerCase();
      for (const d of ['/work/input', '/work/output']) py.FS.mkdirTree(d);
      for (const f of py.FS.readdir('/work/output')) if (f !== '.' && f !== '..') py.FS.unlink('/work/output/' + f);
      py.FS.writeFile('/work/input/' + mfName, new Uint8Array(await files.mf.arrayBuffer()));
      py.FS.writeFile('/work/input/ENTER.pdf', new Uint8Array(await files.en.arrayBuffer()));
      const argv = ['--manifest', '/work/input/' + mfName, '--enter', '/work/input/ENTER.pdf',
                    '--outdir', '/work/output', '--name', 'EDI.xlsx', ...shedArgs()];
      const run = py.globals.get('_run'), png = py.globals.get('_png'), B = py.pyimport('bundle_edi');

      // รอบ 1: หาว่าหน้าไหนต้อง OCR
      setStatus('กำลังแยกชนิดหน้าและอ่านทุกฟอร์ม…', 35);
      await new Promise((r) => setTimeout(r, 30));
      B.BROWSER_OCR = py.toPy({}); B.OCR_WANTED.clear();
      let out = run(py.toPy(argv));
      const wanted = B.OCR_WANTED.toJs ? [...B.OCR_WANTED.toJs()].sort((a, b) => a - b) : [];

      if (wanted.length) {
        const worker = await Tesseract.createWorker('eng');
        const ocr = {};
        for (let k = 0; k < wanted.length; k++) {
          const i = wanted[k];
          setStatus(`OCR หน้าสแกน ${k + 1}/${wanted.length} (หน้า ${i + 1})…`, 40 + 50 * k / wanted.length);
          const bytes = png(py.toPy('/work/input/ENTER.pdf'), i).toJs();
          const { data } = await worker.recognize(new Blob([bytes], { type: 'image/png' }), {}, { blocks: true });
          ocr[i] = tessLines(data);
        }
        await worker.terminate();
        setStatus('กำลังเทียบกับ MF (รอบสุดท้าย)…', 92);
        await new Promise((r) => setTimeout(r, 30));
        // คีย์ต้องเป็นตัวเลข (เลขหน้า 0-based) — object ของ JS ให้คีย์เป็นข้อความ จึงส่งเป็น Map
        B.BROWSER_OCR = py.toPy(new Map(Object.entries(ocr).map(([k, v]) => [Number(k), v])));
        for (const f of py.FS.readdir('/work/output')) if (f !== '.' && f !== '..') py.FS.unlink('/work/output/' + f);
        out = run(py.toPy(argv));
      }
      log(out.trim());
      if (!py.FS.analyzePath('/work/output/EDI.xlsx').exists) throw new Error('ไม่ได้รายงาน — ดูข้อความด้านล่าง');
      reportBytes = py.FS.readFile('/work/output/EDI.xlsx');
      setStatus('เสร็จแล้ว — กดดาวน์โหลดรายงาน', 100);
      $('dl').classList.remove('hidden');
    } catch (e) {
      console.error(e);
      setStatus('ตรวจไม่สำเร็จ: ' + (e.message || e), 0);
      log(String(e && e.message || e));
    } finally {
      $('run').disabled = !(files.mf && files.en);
    }
  }

  function download() {
    if (!reportBytes) return;
    const base = files.en ? files.en.name.replace(/\.pdf$/i, '') : 'ENTER';
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([reportBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    a.download = 'EDI-' + base.replace(/[^\w\-ก-๙]+/g, '_').slice(0, 40) + '.xlsx';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  wireDrop('dropMF', 'fileMF', 'mf', 'fnMF');
  wireDrop('dropEN', 'fileEN', 'en', 'fnEN');
  $('run').addEventListener('click', check);
  $('dl').addEventListener('click', download);
  // ใช้ทดสอบอัตโนมัติ: window.CheckVarious.run(mfFile, enterFile, {shed:'0141'})
  window.CheckVarious = {
    py: () => py,
    boot: () => boot(),
    async run(mf, en, opt = {}) {
      files.mf = mf; files.en = en; $('fnMF').textContent = mf.name; $('fnEN').textContent = en.name;
      if (opt.shed) $('shed').value = opt.shed;
      await check(); return { ok: !!reportBytes, log: $('log').textContent, bytes: reportBytes };
    },
  };
})();
