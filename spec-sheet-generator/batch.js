/* ============================================================
   BATCH.JS — Artwork folder processing + batch PPTX export
   Depends on: app.js globals (state, buildPresentation, getSafeFilename, showToast)
   Libraries : pdfjs-dist, jszip (loaded via CDN in index.html)
   ============================================================ */

// ── PDF.js worker ─────────────────────────────────────────
if (typeof pdfjsLib !== 'undefined') {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
}

// ── Batch state ───────────────────────────────────────────
// Map: filename → { file, format, status, dataUrl, naturalSize, rowIdx }
const batchFiles = new Map();

// ── DOM refs ──────────────────────────────────────────────
const batchDropzone   = document.getElementById('batch-dropzone');
const batchFileInput  = document.getElementById('batch-file-input');
const batchFolderInput = document.getElementById('batch-folder-input');
const batchFileList   = document.getElementById('batch-file-list');
const batchActions    = document.getElementById('batch-actions');
const batchSummary    = document.getElementById('batch-summary');
const batchDiagnostics = document.getElementById('batch-diagnostics');
const btnBatchDl      = document.getElementById('btn-batch-download');

// ── Drag & Drop ───────────────────────────────────────────
batchDropzone.addEventListener('dragover', (e) => {
  e.preventDefault(); batchDropzone.classList.add('dragover');
});
batchDropzone.addEventListener('dragleave', () => batchDropzone.classList.remove('dragover'));
batchDropzone.addEventListener('drop', async (e) => {
  e.preventDefault(); batchDropzone.classList.remove('dragover');

  // Use DataTransferItem API to support recursive folder drops
  if (e.dataTransfer.items && e.dataTransfer.items.length) {
    const entries = [...e.dataTransfer.items]
      .map(item => item.webkitGetAsEntry ? item.webkitGetAsEntry() : null)
      .filter(Boolean);
    const files = await collectFilesFromEntries(entries);
    if (files.length) { handleBatchFiles(files); return; }
  }
  // Fallback for browsers without webkitGetAsEntry
  handleBatchFiles(e.dataTransfer.files);
});

// Recursively collect all files from an array of FileSystemEntry objects
async function collectFilesFromEntries(entries) {
  const files = [];
  async function processEntry(entry) {
    if (entry.isFile) {
      const file = await new Promise((res, rej) => entry.file(res, rej));
      files.push(file);
    } else if (entry.isDirectory) {
      // readEntries() returns at most 100 items per call — must loop until empty
      const reader = entry.createReader();
      let batch;
      do {
        batch = await new Promise((res, rej) => reader.readEntries(res, rej));
        for (const child of batch) await processEntry(child);
      } while (batch.length > 0);
    }
  }
  for (const entry of entries) await processEntry(entry);
  return files;
}

batchFileInput.addEventListener('change', () => {
  if (batchFileInput.files.length) handleBatchFiles(batchFileInput.files);
});
// Folder picker
batchFolderInput.addEventListener('change', () => {
  if (batchFolderInput.files.length) handleBatchFiles(batchFolderInput.files);
});
btnBatchDl.addEventListener('click', runBatchExport);

// ── Entry point ───────────────────────────────────────────
function handleBatchFiles(fileList) {
  // Filter to only artwork files (important for folder uploads which include system files)
  const files = [...fileList].filter(f => isArtworkFile(f));
  if (!files.length) { showToast('No supported artwork files found.', 'error'); return; }

  batchFileList.classList.remove('hidden');
  batchActions.classList.remove('hidden');

  // Process files concurrently (max 4 at a time)
  const queue = [...files];
  let running = 0;
  const MAX_CONCURRENT = 4;
  let pending = files.length;

  function next() {
    while (running < MAX_CONCURRENT && queue.length) {
      running++;
      const file = queue.shift();
      processBatchFile(file).finally(() => {
        running--;
        pending--;
        next();
        updateBatchSummary();
        // Run diagnostics once all files finish converting
        if (pending === 0) renderDiagnostics();
      });
    }
  }
  next();
}

function isArtworkFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  return ['pdf','svg','png','jpg','jpeg','webp','gif','ai','eps'].includes(ext);
}

// ── Per-file processing ───────────────────────────────────
async function processBatchFile(file) {
  const format = detectFormat(file);
  const entry  = { file, format, status: 'converting', dataUrl: null, naturalSize: null, rowIdx: null };
  batchFiles.set(file.name, entry);
  renderBatchItem(file.name);

  try {
    if (format === 'eps') {
      entry.status = 'eps-error';
    } else if (format === 'image') {
      entry.dataUrl = await fileToDataUrl(file);
      entry.naturalSize = await getImageSize(entry.dataUrl);
      entry.rowIdx = matchFileToRow(file.name);
      entry.status = entry.rowIdx !== null ? 'matched' : 'unmatched';
    } else if (format === 'svg') {
      const text = await file.text();
      entry.dataUrl = await svgToDataUrl(text);
      entry.naturalSize = await getImageSize(entry.dataUrl);
      entry.rowIdx = matchFileToRow(file.name);
      entry.status = entry.rowIdx !== null ? 'matched' : 'unmatched';
    } else if (format === 'pdf' || format === 'ai') {
      const buf = await file.arrayBuffer();
      const res = await pdfToDataUrl(buf);
      entry.dataUrl = res.dataUrl;
      entry.naturalSize = res.naturalSize;
      entry.rowIdx = matchFileToRow(file.name);
      entry.status = entry.rowIdx !== null ? 'matched' : 'unmatched';
    } else {
      entry.status = 'error';
    }
  } catch (err) {
    console.error('processBatchFile error:', file.name, err);
    entry.status = 'error';
  }

  renderBatchItem(file.name);
  updateBatchSummary();
}

// ── Format detection ──────────────────────────────────────
function detectFormat(file) {
  const ext  = file.name.split('.').pop().toLowerCase();
  const mime = file.type || '';
  if (['png','jpg','jpeg','webp','gif'].includes(ext) || mime.startsWith('image/')) return 'image';
  if (ext === 'svg' || mime === 'image/svg+xml')                                    return 'svg';
  if (ext === 'pdf' || mime === 'application/pdf')                                  return 'pdf';
  if (ext === 'ai')                                                                  return 'ai';  // treated as PDF
  if (ext === 'eps')                                                                 return 'eps';
  return 'unknown';
}

const FORMAT_LABELS = {
  image: 'IMG', svg: 'SVG', pdf: 'PDF', ai: 'AI', eps: 'EPS', unknown: '?',
};
const FORMAT_BADGE  = {
  image: 'badge-img', svg: 'badge-svg', pdf: 'badge-pdf', ai: 'badge-ai', eps: 'badge-eps', unknown: 'badge-error',
};

// ── Filename → Row matching ───────────────────────────────
function matchFileToRow(filename) {
  const fnCol = state.mapping['filename'];
  const itemCol = state.mapping['item'];
  if (!fnCol && !itemCol) return null;

  const norm = s => s.replace(/\.[^.]+$/, '').toLowerCase().replace(/[\s_\-]+/g, '');
  const uploadedNorm = norm(filename);
  if (!uploadedNorm) return null;

  let bestIdx = null, bestScore = 0;

  for (let i = 0; i < state.rows.length; i++) {
    const row = state.rows[i];
    const candidates = [
      fnCol   && row[fnCol]   ? String(row[fnCol])   : '',
      itemCol && row[itemCol] ? String(row[itemCol]) : '',
    ].filter(Boolean);

    for (const c of candidates) {
      const cn = norm(c);
      if (!cn) continue;
      // Exact
      if (uploadedNorm === cn) return i;
      // Inclusion score
      const longer = Math.max(uploadedNorm.length, cn.length);
      const shorter = Math.min(uploadedNorm.length, cn.length);
      if (uploadedNorm.includes(cn) || cn.includes(uploadedNorm)) {
        const score = shorter / longer;
        if (score > bestScore) { bestScore = score; bestIdx = i; }
      }
    }
  }
  return bestScore >= 0.4 ? bestIdx : null;
}

// ── Converters ────────────────────────────────────────────
function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => resolve(e.target.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function getImageSize(dataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => resolve({ w: 0, h: 0 });
    img.src = dataUrl;
  });
}

async function svgToDataUrl(svgText) {
  return new Promise((resolve, reject) => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(svgText, 'image/svg+xml');
    const svgEl = doc.documentElement;

    // Determine natural dimensions
    let nw = parseFloat(svgEl.getAttribute('width'))  || 0;
    let nh = parseFloat(svgEl.getAttribute('height')) || 0;
    const vbox = svgEl.getAttribute('viewBox');
    if ((!nw || !nh) && vbox) {
      const parts = vbox.trim().split(/\s+/);
      if (parts.length === 4) { nw = parseFloat(parts[2]); nh = parseFloat(parts[3]); }
    }
    if (!nw || isNaN(nw)) nw = 1600;
    if (!nh || isNaN(nh)) nh = 1200;

    // Scale up for PPTX quality (target ~1600px on longest side)
    const scale  = Math.min(4, 1600 / Math.max(nw, nh));
    const canvW  = Math.round(nw * scale);
    const canvH  = Math.round(nh * scale);

    svgEl.setAttribute('width',  canvW);
    svgEl.setAttribute('height', canvH);
    if (!svgEl.getAttribute('viewBox')) svgEl.setAttribute('viewBox', `0 0 ${nw} ${nh}`);

    const serialized = new XMLSerializer().serializeToString(svgEl);
    const blob = new Blob([serialized], { type: 'image/svg+xml' });
    const url  = URL.createObjectURL(blob);
    const img  = new Image();

    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width  = canvW;
      canvas.height = canvH;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvW, canvH);
      ctx.drawImage(img, 0, 0);
      URL.revokeObjectURL(url);
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = (e) => { URL.revokeObjectURL(url); reject(e); };
    img.src = url;
  });
}

async function pdfToDataUrl(arrayBuffer) {
  if (typeof pdfjsLib === 'undefined') throw new Error('PDF.js not loaded');
  const pdf  = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const page = await pdf.getPage(1);

  // Render at 2× screen density for good quality in PPTX
  const viewport = page.getViewport({ scale: 2.5 });
  const canvas   = document.createElement('canvas');
  canvas.width   = Math.round(viewport.width);
  canvas.height  = Math.round(viewport.height);
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  await page.render({ canvasContext: ctx, viewport }).promise;

  return {
    dataUrl:     canvas.toDataURL('image/png'),
    naturalSize: { w: canvas.width, h: canvas.height },
  };
}

// ── Render a single file list item ───────────────────────
function renderBatchItem(filename) {
  const entry = batchFiles.get(filename);
  if (!entry) return;

  let el = document.querySelector(`[data-batch-file="${CSS.escape(filename)}"]`);
  if (!el) {
    el = document.createElement('div');
    el.className = 'batch-file-item';
    el.setAttribute('data-batch-file', filename);
    batchFileList.appendChild(el);
  }

  el.setAttribute('data-status', entry.status);

  // Thumbnail
  let thumbHtml = '';
  if (entry.status === 'converting') {
    thumbHtml = `<div class="batch-thumb"><div class="batch-thumb-spinner"></div></div>`;
  } else if (entry.dataUrl) {
    thumbHtml = `<div class="batch-thumb"><img src="${entry.dataUrl}" alt="thumb"></div>`;
  } else {
    const icon = entry.status === 'eps-error' ? '🚫' : '⚠';
    thumbHtml = `<div class="batch-thumb" style="font-size:18px;">${icon}</div>`;
  }

  // Match info
  let matchText = '', matchClass = '';
  if (entry.status === 'matched') {
    const row = state.rows[entry.rowIdx];
    const itemCol = state.mapping['item'];
    const label = (itemCol && row[itemCol]) ? String(row[itemCol]) : `Row ${entry.rowIdx + 1}`;
    matchText = `→ ${label}`;
    matchClass = 'ok';
  } else if (entry.status === 'unmatched') {
    matchText = 'No matching line item found';
    matchClass = 'warn';
  } else if (entry.status === 'eps-error') {
    matchText = 'EPS cannot be converted in browser — save as PDF from Illustrator';
    matchClass = 'err';
  } else if (entry.status === 'error') {
    matchText = 'Conversion failed';
    matchClass = 'err';
  } else {
    matchText = 'Converting…';
  }

  // Format + status badges
  const fmtLabel  = FORMAT_LABELS[entry.format] || '?';
  const fmtBadge  = FORMAT_BADGE[entry.format]  || 'badge-error';
  const statusBadge = {
    matched:   '<span class="badge badge-ok">Matched</span>',
    unmatched: '<span class="badge badge-warn">Unmatched</span>',
    'eps-error': '<span class="badge badge-error">EPS — cannot convert</span>',
    error:     '<span class="badge badge-error">Error</span>',
    converting:'<span class="badge badge-warn">Converting…</span>',
  }[entry.status] || '';

  el.innerHTML = `
    ${thumbHtml}
    <div class="batch-file-info">
      <span class="batch-file-name" title="${filename}">${filename}</span>
      <span class="batch-file-match ${matchClass}">${matchText}</span>
    </div>
    <div class="batch-badges">
      <span class="badge ${fmtBadge}">${fmtLabel}</span>
      ${statusBadge}
    </div>`;
}

// ── Summary bar ───────────────────────────────────────────
function updateBatchSummary() {
  const entries = [...batchFiles.values()];
  const matched   = entries.filter(e => e.status === 'matched').length;
  const unmatched = entries.filter(e => e.status === 'unmatched').length;
  const errors    = entries.filter(e => ['eps-error','error'].includes(e.status)).length;
  const pending   = entries.filter(e => e.status === 'converting').length;

  batchSummary.textContent =
    `${matched} matched  ·  ${unmatched} unmatched  ·  ${errors} skipped` +
    (pending ? `  ·  ${pending} converting…` : '');

  btnBatchDl.disabled = matched === 0;
}

// ── Diagnostics: missing rows + duplicates ────────────────
function renderDiagnostics() {
  const entries = [...batchFiles.values()];

  // 1. Find spreadsheet rows with NO matched file
  const matchedRowIdxs = new Set(entries.filter(e => e.rowIdx !== null).map(e => e.rowIdx));
  const itemCol = state.mapping['item'];
  const fnCol   = state.mapping['filename'];

  const missingRows = state.rows
    .map((row, i) => ({ row, i }))
    .filter(({ i }) => !matchedRowIdxs.has(i))
    .map(({ row, i }) => {
      const label =
        (itemCol && row[itemCol] ? String(row[itemCol]).trim() : '') ||
        (fnCol   && row[fnCol]   ? String(row[fnCol]).trim()   : '') ||
        `Row ${i + 1}`;
      return label;
    });

  // 2. Find duplicate matches (multiple files → same row)
  const rowHits = {};
  entries.forEach(e => {
    if (e.rowIdx !== null) {
      if (!rowHits[e.rowIdx]) rowHits[e.rowIdx] = [];
      rowHits[e.rowIdx].push(e.file.name);
    }
  });
  const dupes = Object.entries(rowHits)
    .filter(([, files]) => files.length > 1)
    .map(([rowIdx, files]) => {
      const row   = state.rows[parseInt(rowIdx)];
      const label = (itemCol && row[itemCol] ? String(row[itemCol]).trim() : `Row ${parseInt(rowIdx) + 1}`);
      return { label, files };
    });

  // Render
  batchDiagnostics.innerHTML = '';

  if (missingRows.length === 0 && dupes.length === 0) {
    batchDiagnostics.classList.add('hidden');
    return;
  }
  batchDiagnostics.classList.remove('hidden');

  if (missingRows.length) {
    const card = document.createElement('div');
    card.className = 'diag-card diag-missing';
    card.innerHTML = `
      <div class="diag-title">⚠ ${missingRows.length} line item${missingRows.length !== 1 ? 's' : ''} have no matching artwork</div>
      <ul class="diag-list">${missingRows.map(n => `<li>${n}</li>`).join('')}</ul>`;
    batchDiagnostics.appendChild(card);
  }

  if (dupes.length) {
    const card = document.createElement('div');
    card.className = 'diag-card diag-dupes';
    card.innerHTML = `
      <div class="diag-title">✕ ${dupes.length} line item${dupes.length !== 1 ? 's' : ''} matched by multiple files (only first used)</div>
      <ul class="diag-list">${dupes.map(d =>
        `<li>${d.label}: ${d.files.join(', ')}</li>`
      ).join('')}</ul>`;
    batchDiagnostics.appendChild(card);
  }
}

// ── Batch export — chunked PPTX (max 5 per file) ─────────
async function runBatchExport() {
  const matched = [...batchFiles.values()].filter(e => e.status === 'matched' && e.dataUrl);
  if (!matched.length) { showToast('No matched files to export.', 'error'); return; }

  btnBatchDl.disabled = true;
  btnBatchDl.textContent = `Generating ${matched.length} slides…`;

  try {
    const CHUNK_SIZE = 5;
    const chunks = [];
    for (let i = 0; i < matched.length; i += CHUNK_SIZE) {
      chunks.push(matched.slice(i, i + CHUNK_SIZE));
    }

    if (chunks.length === 1) {
      // Single file export
      const pres = new PptxGenJS();
      pres.layout = 'LAYOUT_4x3';
      for (const entry of chunks[0]) {
        const row = state.rows[entry.rowIdx];
        addSlideToPres(pres, row, entry.dataUrl, entry.naturalSize);
      }
      await pres.writeFile({ fileName: 'spec-sheets.pptx' });
      showToast(`Downloaded ${matched.length} slide${matched.length !== 1 ? 's' : ''} as one PPTX!`, 'success');
    } else {
      // Multi-file ZIP export to avoid 100MB+ PPTX files breaking Google Slides
      const zip = new JSZip();
      for (let i = 0; i < chunks.length; i++) {
        const chunk = chunks[i];
        const pres = new PptxGenJS();
        pres.layout = 'LAYOUT_4x3';
        for (const entry of chunk) {
          const row = state.rows[entry.rowIdx];
          addSlideToPres(pres, row, entry.dataUrl, entry.naturalSize);
        }
        const bytes = await pres.write('arraybuffer');
        zip.file(`spec-sheets-part-${String(i + 1).padStart(2, '0')}.pptx`, bytes);
      }

      const blob = await zip.generateAsync({
        type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 5 },
      });

      const url = URL.createObjectURL(blob);
      const a = Object.assign(document.createElement('a'), { href: url, download: 'spec-sheets-batch.zip' });
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 5000);

      showToast(`Downloaded ${chunks.length} PPTX files in a ZIP!`, 'success');
    }
  } catch (err) {
    console.error('Batch export error:', err);
    showToast('Error generating export — see console for details.', 'error');
  } finally {
    btnBatchDl.disabled = false;
    btnBatchDl.innerHTML = `
      <svg width="13" height="13" viewBox="0 0 16 16" fill="none" style="display:inline;vertical-align:middle;margin-right:6px;">
        <path d="M8 2v9M8 11L5 8M8 11L11 8" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
        <path d="M2 14h12" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>
      </svg>
      Generate All &amp; Download`;
  }
}

