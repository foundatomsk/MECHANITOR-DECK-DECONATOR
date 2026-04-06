/* =======================================================
   PRINT SPEC GENERATOR — APP LOGIC
   ======================================================= */

// ── State ────────────────────────────────────────────────
const state = {
  rows:              [],
  headers:           [],
  mapping:           {},
  activeRow:         null,
  artworkDataUrl:    null,
  artworkNaturalSize: null,   // { w, h } in pixels
  projectInfo: {
    metaLine:    '',
    logoText:    '',
    logoDataUrl: null,
    brandColor:  '#E5007D',
  },
};

const SPEC_FIELDS = [
  { key: 'item',       label: 'ITEM'                   },
  { key: 'filename',   label: 'FILE NAME'               },
  { key: 'quantity',   label: 'QUANTITY'                },
  { key: 'dimensions', label: 'DIMENSIONS (TRIM SIZE)', special: 'dimensions' },
  { key: 'material',   label: 'MATERIAL'                },
  { key: 'colorMode',  label: 'COLOR MODE'              },
  { key: 'finish',     label: 'FINISH'                  },
  { key: 'scale',      label: 'SCALE'                   },
  { key: 'bleed',      label: 'BLEED'                   },
  { key: 'notes',      label: 'NOTES'                   },
];

const HEURISTICS = {
  item:      ['item'],
  filename:  ['file name', 'filename', 'file'],
  quantity:  ['qty', 'quantity', 'count', 'units', 'amount'],
  dimWidth:  ['width'],
  dimHeight: ['height'],
  material:  ['material', 'substrate', 'media', 'stock'],
  colorMode: ['color mode', 'colour mode', 'color', 'colour', 'ink'],
  finish:    ['finish', 'laminate', 'coating'],
  scale:     ['scale'],
  bleed:     ['bleed'],
  notes:     ['notes', 'note', 'comment', 'remarks'],
};

// ── DOM Refs ──────────────────────────────────────────────
const dropzone          = document.getElementById('dropzone');
const fileInput         = document.getElementById('file-input');
const fileBadge         = document.getElementById('file-badge');
const fileNameDisplay   = document.getElementById('file-name-display');
const btnClearFile      = document.getElementById('btn-clear-file');

const stepMapping       = document.getElementById('step-mapping');
const mappingGrid       = document.getElementById('mapping-grid');
const mappingToggle     = document.getElementById('mapping-toggle');
const mappingBody       = document.getElementById('mapping-body');
const mappingCaret      = document.getElementById('mapping-toggle-caret');
const mappingStatusText = document.getElementById('mapping-status-text');
const btnApplyMapping   = document.getElementById('btn-apply-mapping');

const stepSelect        = document.getElementById('step-select');
const lineItemSelect    = document.getElementById('line-item-select');

const stepPreview       = document.getElementById('step-preview');
const slidePreview      = document.getElementById('slide-preview');
const btnDownload       = document.getElementById('btn-download');

// ── Project Info inputs ───────────────────────────────────
const fMeta       = document.getElementById('f-meta');
const fLogoText   = document.getElementById('f-logo-text');
const fBrandColor = document.getElementById('f-brand-color');
const fBrandHex   = document.getElementById('f-brand-hex');

fMeta.addEventListener('input', () => {
  state.projectInfo.metaLine = fMeta.value;
  if (state.activeRow) renderSlide(state.activeRow);
});

fLogoText.addEventListener('input', () => {
  state.projectInfo.logoText = fLogoText.value;
  if (state.activeRow) renderSlide(state.activeRow);
});

fBrandColor.addEventListener('input', () => {
  state.projectInfo.brandColor = fBrandColor.value;
  fBrandHex.textContent = fBrandColor.value;
  if (state.activeRow) renderSlide(state.activeRow);
});

// Logo image upload
const logoUploadStrip = document.getElementById('logo-upload-strip');
const logoFileInput   = document.getElementById('logo-file-input');
logoUploadStrip.addEventListener('click', () => logoFileInput.click());
logoFileInput.addEventListener('change', () => {
  if (!logoFileInput.files.length) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    state.projectInfo.logoDataUrl = e.target.result;
    document.getElementById('logo-upload-text').textContent = logoFileInput.files[0].name;
    if (state.activeRow) renderSlide(state.activeRow);
  };
  reader.readAsDataURL(logoFileInput.files[0]);
});

// Artwork image upload — also stores natural pixel dimensions for PPTX
const artworkFileInput = document.getElementById('artwork-file-input');
artworkFileInput.addEventListener('click', function() { this.value = ''; });
artworkFileInput.addEventListener('change', async () => {
  if (!artworkFileInput.files.length) return;
  const file = artworkFileInput.files[0];
  
  const reader = new FileReader();
  reader.onload = (e) => {
    state.artworkDataUrl = e.target.result;
    // Measure natural size so PPTX can maintain aspect ratio correctly
    const probe = new Image();
    probe.onload = () => {
      state.artworkNaturalSize = { w: probe.naturalWidth, h: probe.naturalHeight };
    };
    probe.src = e.target.result;
    if (state.activeRow) renderSlide(state.activeRow);
  };
  reader.readAsDataURL(artworkFileInput.files[0]);
});

// ── Drag & Drop ───────────────────────────────────────────
dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', (e) => {
  e.preventDefault(); dropzone.classList.remove('dragover');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
dropzone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('click', function() { this.value = ''; });
fileInput.addEventListener('change', () => { if (fileInput.files.length) handleFile(fileInput.files[0]); });
btnClearFile.addEventListener('click', (e) => { e.stopPropagation(); resetAll(); });

// ── File Handling ─────────────────────────────────────────
function handleFile(file) {
  fileNameDisplay.textContent = file.name;
  fileBadge.style.display = 'flex';
  const reader = new FileReader();
  reader.onload = (e) => parseSpreadsheet(e.target.result);
  reader.readAsArrayBuffer(file);
}

// ── Parsing — scans for the real header row ───────────────
function parseSpreadsheet(arrayBuffer) {
  try {
    const workbook  = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet     = workbook.Sheets[sheetName];
    const allRows   = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    if (!allRows.length) { showToast('Spreadsheet appears to be empty.', 'error'); return; }

    // Find the row with the most non-empty string cells (skips title rows)
    let headerRowIdx = 0, maxCount = -1;
    for (let i = 0; i < Math.min(10, allRows.length); i++) {
      const count = allRows[i].filter(c => typeof c === 'string' && c.trim().length > 1).length;
      if (count > maxCount) { maxCount = count; headerRowIdx = i; }
    }

    const rawHdr = allRows[headerRowIdx];
    const normHdr = rawHdr.map(h => String(h == null ? '' : h).trim().replace(/\s+/g, ' '));

    state.rows = allRows.slice(headerRowIdx + 1)
      .filter(row => row.some(c => c !== '' && c != null))
      .map(row => {
        const obj = {};
        normHdr.forEach((h, i) => { if (h) obj[h] = row[i] !== undefined ? row[i] : ''; });
        return obj;
      });
    state.headers = normHdr.filter(h => h.length > 0);

    if (!state.rows.length || !state.headers.length) {
      showToast('No data found. Check that your file has a header row.', 'error'); return;
    }

    state.mapping = autoMapColumns(state.headers);
    buildMappingGrid();

    const mappedCount = Object.keys(state.mapping).length;
    if (mappingStatusText) {
      mappingStatusText.textContent = mappedCount > 0
        ? `${mappedCount} column${mappedCount !== 1 ? 's' : ''} mapped automatically`
        : 'No columns detected — please map manually';
    }

    stepMapping.classList.remove('hidden');
    stepSelect.classList.remove('hidden');
    stepPreview.classList.remove('hidden');
    document.getElementById('step-batch').classList.remove('hidden');

    buildLineItemDropdown();
    renderSlideEmpty();

    if (mappedCount === 0) {
      openAccordion();
      showToast('Could not detect columns — see mapping below.', 'error');
    } else {
      showToast(`Loaded ${state.rows.length} row${state.rows.length !== 1 ? 's' : ''}`, 'success');
    }
  } catch (err) {
    console.error('parseSpreadsheet error:', err);
    showToast('Could not parse this file. Please check the format.', 'error');
  }
}

// ── Auto-Mapping ──────────────────────────────────────────
function autoMapColumns(headers) {
  const result = {}, used = new Set();
  const norm = s => String(s).toLowerCase().replace(/\s+/g, ' ').trim();

  const tryMatch = (specKey) => {
    const hints = HEURISTICS[specKey] || [];
    for (const hint of hints) {
      for (const h of headers) {
        if (!used.has(h) && norm(h) === hint) { result[specKey] = h; used.add(h); return; }
      }
    }
    for (const hint of hints) {
      for (const h of headers) {
        if (!used.has(h) && norm(h).includes(hint)) { result[specKey] = h; used.add(h); return; }
      }
    }
  };

  tryMatch('dimWidth'); tryMatch('dimHeight');
  SPEC_FIELDS.forEach(({ key }) => { if (key !== 'dimensions') tryMatch(key); });
  return result;
}

// ── Mapping Grid ──────────────────────────────────────────
function buildMappingGrid() {
  mappingGrid.innerHTML = '';
  SPEC_FIELDS.forEach(({ key, label, special }) => {
    const wrapper = document.createElement('div');
    wrapper.className = 'mapping-row';
    const fl = document.createElement('div');
    fl.className = 'mapping-field-label'; fl.textContent = label;
    wrapper.appendChild(fl);
    if (special === 'dimensions') {
      const split = document.createElement('div');
      split.className = 'dim-split';
      ['dimWidth', 'dimHeight'].forEach((dk, i) => {
        const lbl = document.createElement('span');
        lbl.className = 'dim-sub-label'; lbl.textContent = i === 0 ? 'W' : 'H';
        split.appendChild(lbl);
        split.appendChild(buildColSelect(dk, state.mapping[dk]));
      });
      wrapper.appendChild(split);
    } else {
      wrapper.appendChild(buildColSelect(key, state.mapping[key]));
    }
    mappingGrid.appendChild(wrapper);
  });

  if (state.headers.length) {
    const hint = document.createElement('div');
    hint.className = 'detected-cols-hint';
    hint.innerHTML = `<strong>Detected columns:</strong> ${state.headers.join(', ')}`;
    mappingGrid.appendChild(hint);
  }
}

function buildColSelect(specKey, current) {
  const sel = document.createElement('select');
  sel.className = 'mapping-select'; sel.dataset.specKey = specKey;
  const blank = document.createElement('option');
  blank.value = ''; blank.textContent = '— skip —'; sel.appendChild(blank);
  state.headers.forEach(h => {
    const opt = document.createElement('option');
    opt.value = h; opt.textContent = h;
    if (current && h === current) opt.selected = true;
    sel.appendChild(opt);
  });
  return sel;
}

// ── Accordion ─────────────────────────────────────────────
function openAccordion() {
  mappingBody.classList.remove('hidden');
  mappingCaret.textContent = 'Close ▴';
  mappingToggle.setAttribute('aria-expanded', 'true');
}

mappingToggle.addEventListener('click', () => {
  const isOpen = !mappingBody.classList.contains('hidden');
  mappingBody.classList.toggle('hidden', isOpen);
  mappingCaret.textContent = isOpen ? 'Edit ▾' : 'Close ▴';
  mappingToggle.setAttribute('aria-expanded', String(!isOpen));
});

btnApplyMapping.addEventListener('click', () => {
  mappingGrid.querySelectorAll('.mapping-select').forEach(sel => {
    const key = sel.dataset.specKey;
    if (sel.value) state.mapping[key] = sel.value; else delete state.mapping[key];
  });
  buildLineItemDropdown(); renderSlideEmpty();
  mappingBody.classList.add('hidden');
  mappingCaret.textContent = 'Edit ▾';
  mappingToggle.setAttribute('aria-expanded', 'false');
  showToast('Mapping updated', 'success');
});

// ── Value Resolver ────────────────────────────────────────
function resolveValue(row, key) {
  if (key === 'dimensions') {
    const w = resolveRaw(row, 'dimWidth'), h = resolveRaw(row, 'dimHeight');
    if (w && h) return `${w}"W x ${h}"H`;
    if (w) return `${w}"W`; if (h) return `${h}"H`; return '';
  }
  return resolveRaw(row, key);
}

function resolveRaw(row, colKey) {
  const col = state.mapping[colKey];
  if (!col) return '';
  const val = row[col];
  return val !== undefined && val !== '' ? String(val).trim() : '';
}

// ── Line Item Dropdown ────────────────────────────────────
function buildLineItemDropdown() {
  lineItemSelect.innerHTML = '<option value="">— Choose a line item —</option>';
  const itemCol = state.mapping['item'], fnCol = state.mapping['filename'];
  state.rows.forEach((row, idx) => {
    const opt = document.createElement('option');
    opt.value = idx;
    const label =
      (itemCol && String(row[itemCol] || '').trim()) ||
      (fnCol   && String(row[fnCol]   || '').trim()) ||
      `Row ${idx + 1}`;
    opt.textContent = `${idx + 1}. ${label}`;
    lineItemSelect.appendChild(opt);
  });
}

lineItemSelect.addEventListener('change', () => {
  const idx = parseInt(lineItemSelect.value);
  if (isNaN(idx)) { renderSlideEmpty(); return; }
  state.activeRow = state.rows[idx];
  renderSlide(state.activeRow);
});

// ── Slide Preview ─────────────────────────────────────────
function renderSlideEmpty() {
  slidePreview.innerHTML = `
    <div class="slide-empty">
      <svg width="32" height="32" viewBox="0 0 32 32" fill="none">
        <rect x="4" y="6" width="24" height="20" rx="2" stroke="#ccc" stroke-width="1.5"/>
        <line x1="8" y1="12" x2="16" y2="12" stroke="#ccc" stroke-width="1.5"/>
        <line x1="8" y1="16" x2="20" y2="16" stroke="#ccc" stroke-width="1.5"/>
        <rect x="17" y="13" width="9" height="9" rx="1" stroke="#ccc" stroke-width="1.5"/>
      </svg>
      <span>Select a line item to preview</span>
    </div>`;
}

function renderSlide(row) {
  const info   = state.projectInfo;
  const accent = info.brandColor || '#E5007D';
  const itemName = resolveValue(row, 'item') || 'UNTITLED';

  slidePreview.innerHTML = '';

  // -- Outer frame
  const frame = document.createElement('div');
  frame.className = 'slide-frame';

  // -- Header top row: meta line + logo
  const hdrTop = document.createElement('div');
  hdrTop.className = 'slide-hdr-top';

  // Meta line with // separators in accent color
  const metaEl = document.createElement('div');
  metaEl.className = 'slide-meta-line';
  if (info.metaLine) {
    info.metaLine.split('//').forEach((part, i, arr) => {
      const s = document.createElement('span');
      s.className = 'meta-part'; s.textContent = part.trim();
      metaEl.appendChild(s);
      if (i < arr.length - 1) {
        const sep = document.createElement('span');
        sep.className = 'meta-sep'; sep.textContent = '//';
        sep.style.color = accent; metaEl.appendChild(sep);
      }
    });
  }
  hdrTop.appendChild(metaEl);

  // Logo
  const logoEl = document.createElement('div');
  logoEl.className = 'slide-logo';
  if (info.logoDataUrl) {
    const img = document.createElement('img');
    img.src = info.logoDataUrl; logoEl.appendChild(img);
  } else if (info.logoText) {
    logoEl.textContent = info.logoText;
    logoEl.style.color = accent;
  }
  hdrTop.appendChild(logoEl);
  frame.appendChild(hdrTop);

  // Bold title
  const titleEl = document.createElement('div');
  titleEl.className = 'slide-title';
  titleEl.textContent = itemName;
  frame.appendChild(titleEl);

  // Rule
  const rule = document.createElement('hr');
  rule.className = 'slide-rule';
  frame.appendChild(rule);

  // Body
  const body = document.createElement('div');
  body.className = 'slide-body';

  // LEFT: disclaimer + spec list
  const left = document.createElement('div');
  left.className = 'slide-col-left';

  const disc = document.createElement('p');
  disc.className = 'slide-disclaimer';
  disc.innerHTML = 'for visual reference only<br>please print from files provided';
  left.appendChild(disc);

  const specList = document.createElement('div');
  specList.className = 'slide-spec-list';
  let rowsDrawn = 0;

  SPEC_FIELDS.forEach(({ key, label }) => {
    if (key === 'item') return;
    const value = resolveValue(row, key);
    if (!value) return;
    const rowEl = document.createElement('div');
    rowEl.className = 'spec-row';
    rowEl.style.animationDelay = `${rowsDrawn * 0.03}s`;
    const kEl = document.createElement('span'); kEl.className = 'spec-key'; kEl.textContent = label;
    const sEl = document.createElement('span'); sEl.className = 'spec-sep'; sEl.textContent = ': ';
    const vEl = document.createElement('span'); vEl.className = 'spec-val'; vEl.textContent = value;
    rowEl.appendChild(kEl); rowEl.appendChild(sEl); rowEl.appendChild(vEl);
    specList.appendChild(rowEl); rowsDrawn++;
  });

  left.appendChild(specList);
  
  body.appendChild(left);

  // RIGHT: artwork flanked by measurement lines
  const wVal = resolveRaw(row, 'dimWidth');
  const hVal = resolveRaw(row, 'dimHeight');

  const right = document.createElement('div');
  right.className = 'slide-col-right';

  // Inner grid: [h-ann | artwork] / [spacer | w-ann]
  const artGrid = document.createElement('div');
  artGrid.className = 'art-grid';

  // -- Height annotation (left column) wraps an inner div that JS will position
  const hAnn = document.createElement('div');
  hAnn.className = 'dim-ann-h';
  if (hVal) {
    hAnn.style.color = accent;
    hAnn.innerHTML = `
      <div class="dim-ann-h-inner">
        <span class="dim-ann-text dim-ann-text-h">${hVal}"</span>
        <div class="dim-ruler-v">
          <div class="dim-tick dim-tick-h"></div>
          <div class="dim-shaft-v"></div>
          <div class="dim-tick dim-tick-h"></div>
        </div>
      </div>`;
  }
  artGrid.appendChild(hAnn);

  // -- Artwork frame
  const artFrame = document.createElement('div');
  artFrame.className = 'artwork-frame';
  if (state.artworkDataUrl) {
    const img = document.createElement('img');
    img.className = 'artwork-img';
    img.src = state.artworkDataUrl;
    // After load: adjust rulers to match actual rendered image dimensions
    img.onload = () => requestAnimationFrame(() => adjustRulers(img, artGrid));
    artFrame.appendChild(img);
  } else {
    const ph = document.createElement('div');
    ph.className = 'artwork-placeholder';
    ph.innerHTML = `
      <svg width="24" height="24" viewBox="0 0 28 28" fill="none">
        <rect x="2" y="2" width="24" height="24" rx="3" stroke="#ccc" stroke-width="1.5"/>
        <circle cx="9" cy="10" r="2" stroke="#ccc" stroke-width="1.5"/>
        <path d="M2 20l7-6 5 5 3-3 9 7" stroke="#ccc" stroke-width="1.5" stroke-linejoin="round"/>
      </svg>
      <span>Upload artwork for visual reference</span>`;
    artFrame.appendChild(ph);
  }
  artGrid.appendChild(artFrame);

  // -- Spacer (col 1, row 2) keeps h-ann from stretching into w-ann row
  artGrid.appendChild(document.createElement('div'));

  // -- Width annotation (col 2, row 2) — inner div JS will resize
  const wAnn = document.createElement('div');
  wAnn.className = 'dim-ann-w';
  if (wVal) {
    wAnn.style.color = accent;
    wAnn.innerHTML = `
      <div class="dim-ann-w-inner">
        <div class="dim-ruler-h">
          <div class="dim-tick dim-tick-v"></div>
          <div class="dim-shaft-h"></div>
          <div class="dim-tick dim-tick-v"></div>
        </div>
        <span class="dim-ann-text">${wVal}"</span>
      </div>`;
  }
  artGrid.appendChild(wAnn);

  right.appendChild(artGrid);
  body.appendChild(right);
  frame.appendChild(body);
  slidePreview.appendChild(frame);
}

// Adjusts measurement rulers to match the actual rendered image dimensions
// (image may be letterboxed inside the container with object-fit: contain)
function adjustRulers(img, artGrid) {
  const frame = img.closest('.artwork-frame');
  if (!frame) return;

  const cw = frame.clientWidth;
  const ch = frame.clientHeight;
  const nw = img.naturalWidth;
  const nh = img.naturalHeight;
  if (!cw || !ch || !nw || !nh) return;

  // Compute actual rendered size (object-fit: contain logic)
  const scale = Math.min(cw / nw, ch / nh);
  const rw = Math.round(nw * scale);
  const rh = Math.round(nh * scale);
  const ox = Math.round((cw - rw) / 2); // horizontal letterbox offset
  const oy = Math.round((ch - rh) / 2); // vertical letterbox offset

  // Adjust height ruler: position inner wrapper to span only the actual image height
  const hInner = artGrid.querySelector('.dim-ann-h-inner');
  if (hInner) {
    hInner.style.top    = `${oy}px`;
    hInner.style.height = `${rh}px`;
  }

  // Adjust width ruler: center the inner wrapper over the actual image width
  const wInner = artGrid.querySelector('.dim-ann-w-inner');
  if (wInner) {
    wInner.style.marginLeft = `${ox}px`;
    wInner.style.width      = `${rw}px`;
  }
}

// ── PPTX Core ─────────────────────────────────────────────
// Add one slide to an EXISTING presentation object.
// Called by both single-download and batch multi-slide export.
function addSlideToPres(pres, row, artDataUrl, artNatSize) {
  const info     = state.projectInfo;
  const accent   = (info.brandColor || '#E5007D').replace('#', '');
  const itemName = resolveValue(row, 'item') || 'UNTITLED';

  const slide = pres.addSlide();
  slide.background = { color: 'FFFFFF' };

  const W = 10, H = 7.5, PAD = 0.45, FONT = 'Arial';

  // Helper: thin filled rect instead of ShapeType.line — Google Slides compatible
  const LW_IN = 0.013; // ≈ 1pt at 72dpi, in inches
  const addH = (x, y, w) => slide.addShape(pres.ShapeType.rect,
    { x, y: y - LW_IN / 2, w, h: LW_IN, fill: { color: accent }, line: { color: accent, width: 0 } });
  const addV = (x, y, h) => slide.addShape(pres.ShapeType.rect,
    { x: x - LW_IN / 2, y, w: LW_IN, h, fill: { color: accent }, line: { color: accent, width: 0 } });
  const addHRule = (x, y, w) => slide.addShape(pres.ShapeType.rect,
    { x, y: y - 0.005, w, h: 0.008, fill: { color: 'cccccc' }, line: { color: 'cccccc', width: 0 } });

  // Meta line
  if (info.metaLine) {
    const parts = info.metaLine.split('//');
    const textArr = [];
    parts.forEach((p, i) => {
      textArr.push({ text: p.trim(), options: { color: '666666' } });
      if (i < parts.length - 1) textArr.push({ text: ' // ', options: { color: accent, bold: true } });
    });
    slide.addText(textArr, { x: PAD, y: 0.32, w: W - PAD * 2 - 1.3, h: 0.22, fontSize: 7, fontFace: FONT, valign: 'top' });
  }

  // Logo
  if (info.logoDataUrl) {
    slide.addImage({ data: info.logoDataUrl, x: W - PAD - 1.1, y: 0.26, w: 1.1, h: 0.38,
      sizing: { type: 'contain', align: 'right', valign: 'top' } });
  } else if (info.logoText) {
    slide.addText(info.logoText, { x: W - PAD - 1.1, y: 0.26, w: 1.1, h: 0.38,
      fontSize: 18, fontFace: FONT, bold: true, color: accent, align: 'right' });
  }

  // Title
  slide.addText(itemName, { x: PAD, y: 0.58, w: W - PAD * 2, h: 0.58,
    fontSize: 21, fontFace: FONT, bold: true, color: '111111', valign: 'middle' });

  // Rule (thin rect, more compatible than ShapeType.line)
  addHRule(PAD, 1.24, W - PAD * 2);

  // Disclaimer
  slide.addText('for visual reference only\nplease print from files provided', {
    x: PAD, y: 1.33, w: 3.8, h: 0.45, fontSize: 7.5, fontFace: FONT, italic: true, color: '888888', valign: 'top',
  });

  // Spec rows
  let y = 1.85;
  const LINE_H = 0.27;
  SPEC_FIELDS.forEach(({ key, label }) => {
    if (key === 'item') return;
    const value = resolveValue(row, key);
    if (!value) return;
    slide.addText([
      { text: `${label}: `, options: { bold: true, color: '111111' } },
      { text: value, options: { bold: false, color: '333333' } },
    ], { x: PAD, y, w: 3.9, h: LINE_H, fontSize: 8.5, fontFace: FONT, valign: 'top', wrap: true });
    y += LINE_H;
  });

  // Right column
  const wVal = resolveRaw(row, 'dimWidth');
  const hVal = resolveRaw(row, 'dimHeight');
  const H_ANN = hVal ? 0.38 : 0;
  const W_ANN = wVal ? 0.30 : 0;
  const TICK  = 0.07;
  const artX  = 4.55 + H_ANN, artY = 1.33;
  const artW  = W - artX - PAD, artH = H - artY - PAD - W_ANN;

  // Determine letterboxed image position.
  // We MUST scale solely by the native pixel dimensions of the image.
  // Using spec W/H forces stretching/distortion if the PDF page box doesn't exactly match the written spec ratio.
  let iw = artW, ih = artH, ix = artX, iy = artY;
  if (artDataUrl) {
    if (artNatSize && artNatSize.w > 0 && artNatSize.h > 0) {
      const scale = Math.min(artW / artNatSize.w, artH / artNatSize.h);
      iw = artNatSize.w * scale; ih = artNatSize.h * scale;
      ix = artX + (artW - iw) / 2;
      iy = artY + (artH - ih) / 2;
    }
    slide.addImage({ data: artDataUrl, x: ix, y: iy, w: iw, h: ih });
  } else {
    slide.addShape(pres.ShapeType.rect, { x: artX, y: artY, w: artW, h: artH,
      fill: { color: 'f5f5f5' }, line: { color: 'dddddd', width: 0.5 } });
    slide.addText('Artwork Placeholder', { x: artX, y: artY + artH / 2 - 0.18, w: artW, h: 0.36,
      fontSize: 9, color: 'bbbbbb', align: 'center' });
  }

  // Measurement lines — native PPTX line shapes with arrows on ends
  // This guarantees they are imported as exactly ONE shape piece in Google Slides that won't break apart.
  if (hVal) {
    const lx = artX - H_ANN * 0.45;
    slide.addShape(pres.ShapeType.line, {
      x: lx, y: iy, w: 0, h: ih,
      line: { color: accent, width: 1.0, beginArrowType: 'stealth', endArrowType: 'stealth' }
    });
    slide.addText(`${hVal}"`, {
      x: lx - 1.5, y: iy + ih / 2 - 0.15, w: 1.45, h: 0.4,
      fontSize: 10, fontFace: FONT, color: accent, bold: true, align: 'right', valign: 'middle'
    });
  }

  if (wVal) {
    const ly = iy + ih + W_ANN * 0.45;
    slide.addShape(pres.ShapeType.line, {
      x: ix, y: ly, w: iw, h: 0,
      line: { color: accent, width: 1.0, beginArrowType: 'stealth', endArrowType: 'stealth' }
    });
    slide.addText(`${wVal}"`, {
      x: ix + iw / 2 - 1.0, y: ly + 0.04, w: 2.0, h: 0.4,
      fontSize: 10, fontFace: FONT, color: accent, bold: true, align: 'center', valign: 'middle'
    });
  }

}

// Convenience wrapper for single-slide use
function buildPresentation(row, artDataUrl, artNatSize) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_4x3';
  addSlideToPres(pres, row, artDataUrl, artNatSize);
  return pres;
}



// Helper: safe filename from a row
function getSafeFilename(row) {
  const itemCol = state.mapping['item'], fnCol = state.mapping['filename'];
  const raw =
    (itemCol && row[itemCol] ? String(row[itemCol]) : '') ||
    (fnCol   && row[fnCol]   ? String(row[fnCol])   : '') ||
    'spec-sheet';
  return raw.replace(/[^a-z0-9_\-\.]/gi, '_');
}

// ── Single download ────────────────────────────────────────
btnDownload.addEventListener('click', generatePPTX);

function generatePPTX() {
  if (!state.activeRow) { showToast('Please select a line item first.', 'error'); return; }
  const pres = buildPresentation(state.activeRow, state.artworkDataUrl, state.artworkNaturalSize);
  pres.writeFile({ fileName: `${getSafeFilename(state.activeRow)}.pptx` })
    .then(() => showToast('PPTX downloaded!', 'success'))
    .catch(err => { console.error(err); showToast('Error generating PPTX.', 'error'); });
}

// ── Reset ─────────────────────────────────────────────────
function resetAll() {
  state.rows = []; state.headers = []; state.mapping = {};
  state.activeRow = null; state.artworkDataUrl = null; state.artworkNaturalSize = null;
  fileInput.value = ''; fileNameDisplay.textContent = '';
  fileBadge.style.display = 'none';
  ['step-mapping','step-select','step-preview','step-batch'].forEach(id =>
    document.getElementById(id).classList.add('hidden')
  );
  lineItemSelect.innerHTML = '<option value="">— Choose a line item —</option>';
  mappingGrid.innerHTML = '';
  renderSlideEmpty();
}


// ── Toast ─────────────────────────────────────────────────
function showToast(message, type = 'default') {
  const existing = document.getElementById('toast-notification');
  if (existing) existing.remove();
  const toast = document.createElement('div');
  toast.id = 'toast-notification';
  Object.assign(toast.style, {
    position: 'fixed', bottom: '28px', left: '50%',
    transform: 'translateX(-50%) translateY(20px)',
    background: type === 'error' ? '#ff4d6d' : type === 'success' ? '#22c55e' : '#6C63FF',
    color: '#fff', fontFamily: "'Inter', sans-serif",
    fontSize: '0.82rem', fontWeight: '600',
    padding: '12px 24px', borderRadius: '99px',
    boxShadow: '0 8px 32px rgba(0,0,0,0.35)',
    zIndex: '9999', opacity: '0',
    transition: 'opacity 0.22s ease, transform 0.22s ease',
    whiteSpace: 'nowrap', pointerEvents: 'none',
  });
  toast.textContent = message;
  document.body.appendChild(toast);
  requestAnimationFrame(() => requestAnimationFrame(() => {
    toast.style.opacity = '1'; toast.style.transform = 'translateX(-50%) translateY(0)';
  }));
  setTimeout(() => {
    toast.style.opacity = '0'; toast.style.transform = 'translateX(-50%) translateY(10px)';
    setTimeout(() => toast.remove(), 300);
  }, 2800);
}
