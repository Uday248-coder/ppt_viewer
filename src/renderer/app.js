'use strict';

/* PPT Viewer - renderer.
   Owns all view state and talks to main exclusively through window.pptv. */

const api = window.pptv;

const $ = (id) => document.getElementById(id);

const el = {
  viewLibrary: $('view-library'),
  viewViewer: $('view-viewer'),

  search: $('search'),
  sort: $('sort'),
  deckList: $('deck-list'),
  deckCount: $('deck-count'),
  rootList: $('root-list'),
  recentList: $('recent-list'),
  cacheInfo: $('cache-info'),
  emptyState: $('empty-state'),
  enginePill: $('engine-pill'),

  btnRescan: $('btn-rescan'),
  btnAddFolder: $('btn-add-folder'),
  btnAddFolder2: $('btn-add-folder-2'),
  btnPrune: $('btn-prune'),
  btnClearCache: $('btn-clear-cache'),

  viewerTitle: $('viewer-title'),
  viewerSub: $('viewer-sub'),
  btnBack: $('btn-back'),
  btnGrid: $('btn-grid'),
  btnNotes: $('btn-notes'),
  btnCache: $('btn-cache'),
  btnPdf: $('btn-pdf'),
  btnPresent: $('btn-present'),

  stage: $('stage'),
  slideWrap: $('slide-wrap'),
  slideImg: $('slide-img'),
  slideSpinner: $('slide-spinner'),
  zoomBadge: $('zoom-badge'),
  btnZoomReset: $('btn-zoom-reset'),

  notesPanel: $('notes-panel'),
  notesBody: $('notes-body'),
  notesSlideNo: $('notes-slide-no'),

  filmstrip: $('filmstrip'),
  btnPrev: $('btn-prev'),
  btnNext: $('btn-next'),
  slideNo: $('slide-no'),
  slideTotal: $('slide-total'),

  gridOverlay: $('grid-overlay'),
  gridBody: $('grid-body'),
  btnGridClose: $('btn-grid-close'),

  present: $('present'),
  presentImg: $('present-img'),
  presentStage: $('present-stage'),
  presentNotes: $('present-notes'),
  presentBar: $('present-bar'),
  presentTitle: $('present-title'),
  presentCounter: $('present-counter'),
  presentProgressFill: $('present-progress-fill'),
  pPrev: $('p-prev'),
  pNext: $('p-next'),
  pGrid: $('p-grid'),
  pNotes: $('p-notes'),
  pExit: $('p-exit'),

  toasts: $('toasts'),
  busy: $('busy'),
  busyText: $('busy-text'),
  busySub: $('busy-sub'),
  busyBar: $('busy-bar'),
  busyBarFill: $('busy-bar-fill'),
  dropVeil: $('drop-veil'),
};

const state = {
  view: 'library',
  roots: [],
  recents: [],
  decks: [],
  filtered: [],
  query: '',
  sortBy: 'name',
  deck: null, // active deck payload
  index: 1,
  showNotes: true,
  showGrid: false,
  presenting: false,
  presentNotes: false,
  zoom: { scale: 1, tx: 0, ty: 0, active: false },
  busyDepth: 0,
};

// Monotonic tokens that invalidate in-flight async work when the user acts
// faster than PowerPoint can render.
let openToken = 0;
let slideToken = 0;

/* ------------------------------------------------------------------ utils */

const fmtBytes = (b) => {
  if (!b) return '0 B';
  const u = ['B', 'KB', 'MB', 'GB'];
  const i = Math.min(u.length - 1, Math.floor(Math.log(b) / Math.log(1024)));
  return `${(b / Math.pow(1024, i)).toFixed(i === 0 ? 0 : 1)} ${u[i]}`;
};

const fmtDate = (ms) => {
  if (!ms) return '';
  const d = new Date(ms);
  return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
};

function toast(message, kind = '', ms = 2600) {
  const t = document.createElement('div');
  t.className = `toast${kind ? ` toast-${kind}` : ''}`;
  t.textContent = message;
  el.toasts.appendChild(t);
  setTimeout(() => {
    t.style.transition = 'opacity .2s';
    t.style.opacity = '0';
    setTimeout(() => t.remove(), 220);
  }, ms);
}

function busy(on, text = 'Working…', sub = '', progress = null) {
  state.busyDepth = Math.max(0, state.busyDepth + (on ? 1 : -1));
  const show = state.busyDepth > 0;
  el.busy.hidden = !show;
  if (!show) return;
  el.busyText.textContent = text;
  el.busySub.textContent = sub;
  el.busySub.style.display = sub ? '' : 'none';
  if (progress === null) {
    el.busyBar.hidden = true;
  } else {
    el.busyBar.hidden = false;
    const pct = Math.max(0, Math.min(100, progress.total ? (progress.done / progress.total) * 100 : 0));
    el.busyBarFill.style.width = `${pct}%`;
    el.busySub.textContent = `${progress.done} / ${progress.total}`;
  }
}

function shortPath(p) {
  const home = 'C:\\Users\\User';
  return p.startsWith(home) ? '~' + p.slice(home.length) : p;
}

/** IPC failures arrive wrapped as "Error invoking remote method 'x': Error: ...". */
function cleanError(e) {
  const m = (e && e.message) || String(e);
  const cleaned = m
    .replace(/^Error invoking remote method '[^']*':\s*/i, '')
    .replace(/^(Uncaught )?Error:\s*/i, '');
  return cleaned.trim() || m;
}

/* ------------------------------------------------------------------ views */

function setView(name) {
  state.view = name;
  el.viewLibrary.classList.toggle('is-active', name === 'library');
  el.viewViewer.classList.toggle('is-active', name === 'viewer');
  if (name === 'viewer') requestAnimationFrame(fitSlide);
  else fitSlide();
}

function showEmpty(open) {
  el.emptyState.classList.toggle('is-open', open);
}

/* ------------------------------------------------------------------ library */

function renderRoots() {
  el.rootList.innerHTML = '';
  for (const r of state.roots) {
    const li = document.createElement('li');
    li.className = 'root-item';
    li.title = r;

    const nm = document.createElement('span');
    nm.className = 'root-name';
    nm.textContent = r.split(/[\\/]/).filter(Boolean).pop() || r;
    li.appendChild(nm);

    const x = document.createElement('button');
    x.className = 'root-x';
    x.textContent = '×';
    x.title = 'Remove this folder';
    x.addEventListener('click', async (e) => {
      e.stopPropagation();
      state.roots = await api.library.removeRoot(r);
      renderRoots();
      toast('Folder removed');
    });
    li.appendChild(x);
    el.rootList.appendChild(li);
  }
}

function renderRecents() {
  el.recentList.innerHTML = '';
  if (!state.recents.length) {
    const li = document.createElement('li');
    li.className = 'deck-list-empty';
    li.style.padding = '10px';
    li.textContent = 'Nothing opened yet';
    el.recentList.appendChild(li);
    return;
  }
  for (const r of state.recents.slice(0, 24)) {
    const li = document.createElement('li');
    li.className = 'recent-item';
    li.title = r.path;
    li.textContent = r.name;
    const d = document.createElement('span');
    d.className = 'recent-dir';
    d.textContent = shortPath(r.path.replace(/[\\/][^\\/]+$/, ''));
    li.appendChild(d);
    li.addEventListener('click', () => openDeck(r.path));
    el.recentList.appendChild(li);
  }
}

function applyFilter() {
  const q = state.query.trim().toLowerCase();
  let list = state.decks;
  if (q) {
    list = list.filter(
      (d) => d.name.toLowerCase().includes(q) || d.path.toLowerCase().includes(q)
    );
  }
  const cmp = {
    name: (a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }),
    date: (a, b) => b.mtimeMs - a.mtimeMs,
    size: (a, b) => b.size - a.size,
  }[state.sortBy];
  state.filtered = list.slice().sort(cmp);
  renderDeckList();
}

function renderDeckList() {
  el.deckList.innerHTML = '';
  const list = state.filtered;

  if (state.roots.length) {
    el.deckCount.textContent = list.length
      ? `${list.length} deck${list.length === 1 ? '' : 's'}${state.query ? ' matching' : ''}`
      : state.query
      ? 'No matches'
      : 'No presentations found in these folders';
  }

  if (!list.length) {
    const d = document.createElement('div');
    d.className = 'deck-list-empty';
    d.textContent = state.roots.length
      ? state.query
        ? 'Nothing matches that search.'
        : 'No .ppt, .pptx or .ppsx files in these folders.'
      : 'Add a folder to get started.';
    el.deckList.appendChild(d);
    return;
  }

  for (const deck of list) {
    const row = document.createElement('div');
    row.className = 'deck-row';
    if (state.deck && state.deck.path === deck.path) row.classList.add('is-active');
    row.title = deck.path;

    const ic = document.createElement('div');
    ic.className = 'deck-icon';
    ic.textContent = 'P';
    row.appendChild(ic);

    const main = document.createElement('div');
    main.className = 'deck-main';
    const nm = document.createElement('div');
    nm.className = 'deck-name';
    nm.textContent = deck.name;
    main.appendChild(nm);
    const pth = document.createElement('div');
    pth.className = 'deck-path';
    pth.textContent = shortPath(deck.path.replace(/[\\/][^\\/]+$/, ''));
    main.appendChild(pth);
    row.appendChild(main);

    const meta = document.createElement('div');
    meta.className = 'deck-meta';
    const b1 = document.createElement('span');
    b1.className = 'badge';
    b1.textContent = deck.ext.replace('.', '').toUpperCase();
    meta.appendChild(b1);
    const b2 = document.createElement('span');
    b2.textContent = fmtBytes(deck.size);
    meta.appendChild(b2);
    const b3 = document.createElement('span');
    b3.textContent = fmtDate(deck.mtimeMs);
    meta.appendChild(b3);
    row.appendChild(meta);

    row.addEventListener('click', () => openDeck(deck.path));
    row.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      api.library.reveal(deck.path);
    });
    el.deckList.appendChild(row);
  }
}

async function refreshInfo() {
  const info = await api.app.info();
  state.roots = info.roots || [];
  state.recents = info.recents || [];
  renderRoots();
  renderRecents();
  el.cacheInfo.textContent = `Cache: ${fmtBytes(info.cacheBytes)}`;
  return info;
}

async function rescan({ quiet = false } = {}) {
  if (!state.roots.length) {
    showEmpty(true);
    return;
  }
  if (!quiet) busy(true, 'Scanning folders…');
  try {
    const res = await api.library.scan(state.roots);
    state.decks = res.decks;
    applyFilter();
    showEmpty(false);
    if (!quiet) {
      const secs = (res.ms / 1000).toFixed(1);
      toast(`Found ${res.decks.length} deck${res.decks.length === 1 ? '' : 's'} in ${secs}s`, 'ok');
      if (res.errors.length) toast(`Skipped ${res.errors.length} unreadable folder(s)`, 'err');
    }
  } catch (e) {
    toast(`Scan failed: ${cleanError(e)}`, 'err');
  } finally {
    if (!quiet) busy(false);
  }
}

async function addFolder() {
  const r = await api.library.chooseRoot();
  if (!r.added || !r.added.length) return;
  state.roots = r.roots;
  renderRoots();
  await rescan();
}

/* ------------------------------------------------------------------ deck open */

async function openDeck(p) {
  // Guards against a slow open being overtaken by a faster later one, which
  // would otherwise paint a previous deck's thumbnails into the new deck.
  const token = ++openToken;
  busy(true, 'Opening deck…', shortPath(p));
  el.slideSpinner.hidden = false;
  try {
    const deck = await api.deck.open(p);
    if (token !== openToken) return;

    state.deck = deck;
    state.index = 1;
    state.recents = (await api.app.info()).recents || [];
    if (token !== openToken) return;
    renderRecents();
    renderDeckList();

    el.viewerTitle.textContent = deck.title || deck.name;
    el.viewerSub.textContent = `${deck.slideCount} slides · ${deck.widthPt}×${deck.heightPt} pt${
      deck.fromCache ? ' · from cache' : ''
    }`;
    el.slideTotal.textContent = deck.slideCount;
    setView('viewer');
    el.slideSpinner.hidden = true;

    // Thumbnails first (very fast), so the filmstrip and grid are usable
    // immediately; full-resolution slides stream in afterwards.
    api.deck
      .ensureThumbs(p)
      .then(() => {
        if (token !== openToken) return;
        buildFilmstrip();
        if (state.showGrid) buildGrid();
      })
      .catch((e) => toast(`Thumbnails: ${cleanError(e)}`, 'err'));

    await goToSlide(1, { animate: false });
  } catch (e) {
    if (token === openToken) {
      el.slideSpinner.hidden = true;
      toast(`Could not open deck: ${cleanError(e)}`, 'err', 5200);
    }
  } finally {
    if (token === openToken) busy(false);
  }
}

async function ensureFull(indices) {
  if (!state.deck) return;
  try {
    await api.deck.ensureSlides(state.deck.path, indices);
  } catch (e) {
    toast(`Render failed: ${cleanError(e)}`, 'err', 4200);
  }
}

/** Full-resolution slide URL, with a cache-busting token while rendering. */
function fullUrl(i) {
  const s = state.deck.slides[i - 1];
  return s ? s.fullUrl : '';
}

/**
 * Maps PowerPoint's EntryEffect to a CSS animation.
 *
 * Codes were measured by injecting each OOXML transition type and reading back
 * the value PowerPoint reports (see the table in README). Modern decks use the
 * p14 range (3840+), which is not covered by the classic PpEntryEffect docs.
 */
const EFFECT_MAP = {
  258: 'fx-fade',   // cut through black
  513: 'fx-fade',   // random - pick a sane default
  1537: 'fx-fade',  // dissolve
  1538: 'fx-fade',  // fade through black
  1793: 'fx-push-l',
  1794: 'fx-push-r',
  1921: 'fx-push-r',
  1922: 'fx-push-l',
  2049: 'fx-wipe',
  2050: 'fx-wipe',
  2817: 'fx-wipe',  // wipe (p14)
  3073: 'fx-zoom',  // zoom (p14)
  3845: 'fx-zoom',  // circle
  3846: 'fx-zoom',  // diamond
  3849: 'fx-fade',  // fade (p14)
  3850: 'fx-fade',  // newsflash
  3851: 'fx-zoom',  // plus
  3853: 'fx-push-l',// push (p14)
  3856: 'fx-zoom',  // wedge
};

// Effects that mean "no animation": absent, none, and a plain cut.
const NO_ANIMATION = new Set([null, undefined, 0, 257]);

function effectClass(effect) {
  if (NO_ANIMATION.has(effect)) return null;
  return EFFECT_MAP[effect] || 'fx-fade';
}

const FX_CLASSES = ['fx', 'fx-fade', 'fx-fade-up', 'fx-push-l', 'fx-push-r', 'fx-zoom', 'fx-wipe'];

function applyTransition(img, effect) {
  const name = effectClass(effect);
  img.classList.remove(...FX_CLASSES);
  if (!name) return;
  void img.offsetWidth; // restart the animation
  img.classList.add('fx', name);
}

async function goToSlide(i, { animate = true } = {}) {
  if (!state.deck) return;
  const n = state.deck.slideCount;
  const idx = Math.max(1, Math.min(n, i));
  // Holding an arrow key queues several renders; only the newest may paint.
  const token = ++slideToken;
  state.index = idx;

  el.slideNo.textContent = idx;
  el.notesSlideNo.textContent = `Slide ${idx} of ${n}`;

  const s = state.deck.slides[idx - 1];
  el.notesBody.textContent = s && s.notes ? s.notes : 'No speaker notes for this slide.';
  el.notesBody.classList.toggle('is-empty', !(s && s.notes));

  updateFilmstripActive();
  resetZoom();

  if (state.presenting) {
    await ensureFull([idx]);
    if (token !== slideToken) return;
    if (animate) applyTransition(el.presentImg, s ? s.effect : 0);
    el.presentImg.src = fullUrl(idx);
    el.presentCounter.textContent = `${idx} / ${n}`;
    el.presentProgressFill.style.width = `${(idx / n) * 100}%`;
    el.presentTitle.textContent = state.deck.title || state.deck.name;
  } else {
    el.slideSpinner.hidden = false;
    await ensureFull([idx]);
    if (token !== slideToken) return;
    if (animate) applyTransition(el.slideImg, s ? s.effect : 0);
    el.slideImg.src = fullUrl(idx);
    el.slideSpinner.hidden = true;
  }

  // Warm the neighbours so paging never stalls on a render.
  ensureFull([idx + 1, idx + 2, idx - 1].filter((x) => x >= 1 && x <= n));
}

/* ------------------------------------------------------------------ zoom/pan */

// The slide is laid out at a fixed intrinsic size and scaled to fit, so zoom is
// a single scalar with no compounding layout maths.
const SLIDE_W = 1000;

function intrinsic() {
  const aspect = state.deck.aspect || 4 / 3;
  return { w: SLIDE_W, h: SLIDE_W / aspect };
}

function stageBox() {
  const r = el.stage.getBoundingClientRect();
  return { w: r.width, h: r.height, left: r.left, top: r.top };
}

function fitScale() {
  const b = stageBox();
  const { w, h } = intrinsic();
  const pad = 36;
  if (b.w <= pad * 2 || b.h <= pad * 2) return 1;
  return Math.min((b.w - pad * 2) / w, (b.h - pad * 2) / h);
}

function applyTransform() {
  if (!state.deck) return;
  const { w, h } = intrinsic();
  el.slideWrap.style.width = `${w}px`;
  el.slideWrap.style.height = `${h}px`;

  const s = fitScale() * state.zoom.scale;
  el.slideWrap.style.transform = `translate3d(${state.zoom.tx}px, ${state.zoom.ty}px, 0) scale(${s})`;

  const pct = Math.round(state.zoom.scale * 100);
  el.zoomBadge.textContent = `${pct}%`;
  el.zoomBadge.hidden = !state.zoom.active;
  el.btnZoomReset.hidden = !state.zoom.active;
  el.stage.classList.toggle('is-zoomed', state.zoom.scale > 1.02);
}

function fitSlide() {
  if (!state.deck) return;
  state.zoom = { scale: 1, tx: 0, ty: 0, active: false };
  const b = stageBox();
  const { w, h } = intrinsic();
  const s = fitScale();
  state.zoom.tx = (b.w - w * s) / 2;
  state.zoom.ty = (b.h - h * s) / 2;
  applyTransform();
}

function resetZoom() {
  fitSlide();
}

function zoomBy(factor, cx, cy) {
  if (!state.deck) return;
  const b = stageBox();
  const fit = fitScale();
  const cur = fit * state.zoom.scale;
  const nextZoom = Math.max(0.25, Math.min(9, state.zoom.scale * factor));
  if (nextZoom === state.zoom.scale) return;
  const next = fit * nextZoom;

  // Keep the point under the cursor pinned while scaling.
  const px = cx === undefined ? b.w / 2 : cx - b.left;
  const py = cy === undefined ? b.h / 2 : cy - b.top;
  state.zoom.tx = px - (px - state.zoom.tx) * (next / cur);
  state.zoom.ty = py - (py - state.zoom.ty) * (next / cur);
  state.zoom.scale = nextZoom;
  state.zoom.active = nextZoom > 1.02 || nextZoom < 0.98;
  applyTransform();
}

/* ------------------------------------------------------------------ filmstrip */

function buildFilmstrip() {
  if (!state.deck) return;
  el.filmstrip.innerHTML = '';
  const aspect = state.deck.aspect || 4 / 3;
  state.deck.slides.forEach((s) => {
    const d = document.createElement('div');
    d.className = 'thumb';
    d.style.width = `${Math.round(54 * aspect)}px`;
    d.dataset.index = s.index;
    d.title = s.title || `Slide ${s.index}`;

    const img = document.createElement('img');
    img.src = s.thumbUrl;
    img.alt = '';
    img.draggable = false;
    d.appendChild(img);

    if (s.notes) {
      const dot = document.createElement('div');
      dot.className = 'thumb-note';
      d.appendChild(dot);
    }
    const n = document.createElement('div');
    n.className = 'thumb-n';
    n.textContent = s.index;
    d.appendChild(n);

    d.addEventListener('click', () => goToSlide(s.index));
    el.filmstrip.appendChild(d);
  });
  updateFilmstripActive();
}

function updateFilmstripActive() {
  for (const t of el.filmstrip.children) {
    t.classList.toggle('is-active', Number(t.dataset.index) === state.index);
  }
  const active = el.filmstrip.querySelector('.is-active');
  if (active) {
    const fr = el.filmstrip;
    const l = active.offsetLeft - fr.clientWidth / 2 + active.clientWidth / 2;
    fr.scrollTo({ left: Math.max(0, l), behavior: 'smooth' });
  }
}

/* ------------------------------------------------------------------ grid */

function buildGrid() {
  if (!state.deck) return;
  el.gridBody.innerHTML = '';
  state.deck.slides.forEach((s) => {
    const cell = document.createElement('div');
    cell.className = 'gcell';
    if (s.index === state.index) cell.classList.add('is-active');
    cell.style.aspectRatio = String(state.deck.aspect || 1.777);

    const wrap = document.createElement('div');
    wrap.className = 'gcell-imgwrap';
    const img = document.createElement('img');
    img.src = s.thumbUrl;
    img.alt = '';
    img.draggable = false;
    wrap.appendChild(img);
    cell.appendChild(wrap);

    const cap = document.createElement('div');
    cap.className = 'gcell-cap';
    const no = document.createElement('span');
    no.className = 'gcell-no';
    no.textContent = s.index;
    cap.appendChild(no);
    const ti = document.createElement('span');
    ti.className = 'gcell-title';
    ti.textContent = s.title || '—';
    cap.appendChild(ti);
    if (s.notes) {
      const dot = document.createElement('span');
      dot.className = 'thumb-note';
      dot.style.position = 'static';
      cap.appendChild(dot);
    }
    cell.appendChild(cap);

    cell.addEventListener('click', () => {
      closeGrid();
      goToSlide(s.index);
    });
    el.gridBody.appendChild(cell);
  });
}

function openGrid() {
  state.showGrid = true;
  buildGrid();
  el.gridOverlay.hidden = false;
}

function closeGrid() {
  state.showGrid = false;
  el.gridOverlay.hidden = true;
}

function gridVisible() {
  return !el.gridOverlay.hidden;
}

/* ------------------------------------------------------------------ present */

async function enterPresent() {
  if (!state.deck) return;
  state.presenting = true;
  el.present.hidden = false;
  el.presentNotes.hidden = !state.presentNotes;
  el.pNotes.classList.toggle('is-on', state.presentNotes);
  el.presentBar.classList.add('is-pinned');
  setTimeout(() => el.presentBar.classList.remove('is-pinned'), 2200);
  await api.window.present();
  await goToSlide(state.index, { animate: false });
  // Render a generous window so a live presentation never stalls.
  busy(true, 'Preparing slideshow…', 'Rendering upcoming slides', { done: 0, total: 1 });
  const ahead = [];
  for (let k = 1; k <= 8; k++) ahead.push(state.index + k);
  try {
    await api.deck.ensureSlides(state.deck.path, ahead);
  } catch (e) {
    toast(`Preload: ${cleanError(e)}`, 'err');
  }
  busy(false);
}

function exitPresent() {
  state.presenting = false;
  el.present.hidden = true;
  api.window.exitPresent();
  fitSlide();
}

function presentVisible() {
  return !el.present.hidden;
}

function togglePresentNotes() {
  state.presentNotes = !state.presentNotes;
  el.presentNotes.hidden = !state.presentNotes;
  el.pNotes.classList.toggle('is-on', state.presentNotes);
  if (state.presentNotes) {
    const s = state.deck.slides[state.index - 1];
    el.presentNotes.textContent = s && s.notes ? s.notes : 'No speaker notes for this slide.';
  }
}

function toggleNotes() {
  state.showNotes = !state.showNotes;
  el.notesPanel.classList.toggle('is-open', state.showNotes);
  el.btnNotes.classList.toggle('is-on', state.showNotes);
  setTimeout(fitSlide, 30);
  api.app.setPref('showNotes', state.showNotes);
}

/* ------------------------------------------------------------------ actions */

async function cacheDeck() {
  if (!state.deck) return;
  const d = state.deck;
  busy(true, 'Caching deck…', d.name, { done: 0, total: d.slideCount });
  try {
    const r = await api.deck.cacheAll(d.path);
    d.cached = 'full';
    toast(
      r.rendered ? `Cached ${r.rendered} slide${r.rendered === 1 ? '' : 's'} — opens offline now` : 'Already fully cached',
      'ok',
      3400
    );
    await refreshInfo();
  } catch (e) {
    toast(`Caching failed: ${cleanError(e)}`, 'err', 4600);
  } finally {
    busy(false);
  }
}

async function exportPdf() {
  if (!state.deck) return;
  busy(true, 'Exporting PDF…', state.deck.name);
  try {
    const r = await api.deck.exportPdf(state.deck.path);
    if (!r.cancelled) toast('PDF exported', 'ok');
  } catch (e) {
    toast(`PDF export failed: ${cleanError(e)}`, 'err', 5000);
  } finally {
    busy(false);
  }
}

async function backToLibrary() {
  if (presentVisible()) return exitPresent();
  if (gridVisible()) return closeGrid();
  state.deck = null;
  setView('library');
  renderDeckList();
  await refreshInfo();
}

/* ------------------------------------------------------------------ events */

function bind() {
  el.btnAddFolder.addEventListener('click', addFolder);
  el.btnAddFolder2.addEventListener('click', addFolder);
  el.btnRescan.addEventListener('click', () => rescan());
  el.btnBack.addEventListener('click', backToLibrary);

  el.btnPrune.addEventListener('click', async () => {
    busy(true, 'Pruning cache…');
    try {
      const r = await api.library.prune();
      toast(`Removed ${r.removed} stale entr${r.removed === 1 ? 'y' : 'ies'} · freed ${fmtBytes(r.bytes)}`, 'ok');
      await refreshInfo();
    } finally {
      busy(false);
    }
  });

  el.btnClearCache.addEventListener('click', async () => {
    busy(true, 'Clearing cache…');
    try {
      await api.library.clearCache();
      state.deck = null;
      setView('library');
      toast('Cache cleared', 'ok');
      await refreshInfo();
    } finally {
      busy(false);
    }
  });

  el.search.addEventListener('input', (e) => {
    state.query = e.target.value;
    applyFilter();
  });
  el.sort.addEventListener('change', (e) => {
    state.sortBy = e.target.value;
    api.app.setPref('sortBy', state.sortBy);
    applyFilter();
  });

  el.btnGrid.addEventListener('click', () => (gridVisible() ? closeGrid() : openGrid()));
  el.btnGridClose.addEventListener('click', closeGrid);
  el.btnNotes.addEventListener('click', toggleNotes);
  el.btnCache.addEventListener('click', cacheDeck);
  el.btnPdf.addEventListener('click', exportPdf);
  el.btnPresent.addEventListener('click', enterPresent);
  el.btnZoomReset.addEventListener('click', resetZoom);
  el.btnPrev.addEventListener('click', () => goToSlide(state.index - 1));
  el.btnNext.addEventListener('click', () => goToSlide(state.index + 1));

  el.pPrev.addEventListener('click', () => goToSlide(state.index - 1));
  el.pNext.addEventListener('click', () => goToSlide(state.index + 1));
  el.pExit.addEventListener('click', exitPresent);
  el.pNotes.addEventListener('click', togglePresentNotes);
  el.pGrid.addEventListener('click', () => {
    if (presentVisible()) exitPresent();
    openGrid();
  });

  // Zoom / pan
  el.stage.addEventListener(
    'wheel',
    (e) => {
      if (!state.deck) return;
      e.preventDefault();
      zoomBy(e.deltaY < 0 ? 1.12 : 1 / 1.12, e.clientX, e.clientY);
    },
    { passive: false }
  );

  let dragging = false;
  let lastX = 0;
  let lastY = 0;
  el.stage.addEventListener('mousedown', (e) => {
    if (!state.deck || e.button !== 0) return;
    dragging = true;
    lastX = e.clientX;
    lastY = e.clientY;
    el.stage.classList.add('is-panning');
  });
  window.addEventListener('mousemove', (e) => {
    if (!dragging) return;
    state.zoom.tx += e.clientX - lastX;
    state.zoom.ty += e.clientY - lastY;
    lastX = e.clientX;
    lastY = e.clientY;
    state.zoom.active = true;
    applyTransform();
  });
  window.addEventListener('mouseup', () => {
    dragging = false;
    el.stage.classList.remove('is-panning');
  });

  el.stage.addEventListener('dblclick', (e) => {
    if (state.zoom.active) resetZoom();
    else zoomBy(2, e.clientX, e.clientY);
  });

  window.addEventListener('resize', () => {
    if (state.view === 'viewer' && !state.zoom.active) fitSlide();
  });

  // Keyboard
  window.addEventListener('keydown', (e) => {
    const typing = /^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement?.tagName || '');
    if (typing && e.key !== 'Escape') return;

    if (e.key === 'Escape') {
      if (presentVisible()) return exitPresent();
      if (gridVisible()) return closeGrid();
      if (state.view === 'viewer') return backToLibrary();
      return;
    }

    if (presentVisible()) {
      switch (e.key) {
        case 'ArrowRight': case 'PageDown': case ' ': case 'Enter':
          e.preventDefault(); goToSlide(state.index + 1); return;
        case 'ArrowLeft': case 'PageUp': case 'Backspace':
          e.preventDefault(); goToSlide(state.index - 1); return;
        case 'Home': e.preventDefault(); goToSlide(1); return;
        case 'End': e.preventDefault(); goToSlide(state.deck.slideCount); return;
        case 'g': case 'G': e.preventDefault(); exitPresent(); openGrid(); return;
        case 'n': case 'N': e.preventDefault(); togglePresentNotes(); return;
        case 'f': case 'F': e.preventDefault(); resetZoom(); return;
        default: return;
      }
    }

    if (state.view !== 'viewer' || !state.deck) return;

    switch (e.key) {
      case 'ArrowRight': case 'PageDown': case ' ': case 'Enter':
        e.preventDefault(); goToSlide(state.index + 1); break;
      case 'ArrowLeft': case 'PageUp': case 'Backspace':
        e.preventDefault(); goToSlide(state.index - 1); break;
      case 'ArrowDown': e.preventDefault(); goToSlide(state.index + 1); break;
      case 'ArrowUp': e.preventDefault(); goToSlide(state.index - 1); break;
      case 'Home': e.preventDefault(); goToSlide(1); break;
      case 'End': e.preventDefault(); goToSlide(state.deck.slideCount); break;
      case 'g': case 'G': e.preventDefault(); gridVisible() ? closeGrid() : openGrid(); break;
      case 'n': case 'N': e.preventDefault(); toggleNotes(); break;
      case 'F5': e.preventDefault(); enterPresent(); break;
      case '0': e.preventDefault(); resetZoom(); break;
      case '+': case '=': e.preventDefault(); zoomBy(1.25); break;
      case '-': case '_': e.preventDefault(); zoomBy(1 / 1.25); break;
      case 'p': case 'P': e.preventDefault(); cacheDeck(); break;
      default: break;
    }
  });

  // Drag & drop a folder or deck
  let dragDepth = 0;
  window.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dragDepth += 1;
    el.dropVeil.hidden = false;
  });
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dragDepth = Math.max(0, dragDepth - 1);
    if (!dragDepth) el.dropVeil.hidden = true;
  });
  window.addEventListener('drop', async (e) => {
    e.preventDefault();
    dragDepth = 0;
    el.dropVeil.hidden = true;
    const paths = [...(e.dataTransfer?.files || [])].map((f) => f.path).filter(Boolean);
    if (!paths.length) return;
    // A dropped deck opens straight away; a dropped folder joins the library.
    const decks = paths.filter((p) => /\.(pptx?|ppsx?|pptm|ppsm|potx|potm)$/i.test(p));
    const folders = paths.filter((p) => !/\.(pptx?|ppsx?|pptm|ppsm|potx|potm)$/i.test(p));
    for (const f of folders) await api.library.addRoot(f);
    state.roots = (await api.app.info()).roots;
    renderRoots();
    await rescan();
    if (decks.length === 1) await openDeck(decks[0]);
    else if (decks.length > 1) toast(`Added ${decks.length} decks to the library`);
  });

  api.app.onEngineStatus((s) => {
    if (s.restarting) {
      el.enginePill.className = 'pill pill-wait';
      el.enginePill.textContent = 'restarting';
      return;
    }
    if (s.comAvailable) {
      el.enginePill.className = 'pill pill-ok';
      el.enginePill.textContent = `engine ${s.version || 'ready'}`;
      el.enginePill.title = 'PowerPoint rendering engine ready';
    } else if (s.error) {
      el.enginePill.className = 'pill pill-err';
      el.enginePill.textContent = 'no engine';
      el.enginePill.title = s.error;
    } else {
      el.enginePill.className = 'pill pill-wait';
      el.enginePill.textContent = 'starting…';
    }
  });

  api.deck.onProgress((p) => {
    if (p.phase === 'done') busy(false);
  });
}

/* ------------------------------------------------------------------ boot */

(async function init() {
  bind();
  try {
    const info = await api.app.info();
    state.roots = info.roots || [];
    state.recents = info.recents || [];
    state.sortBy = info.prefs?.sortBy || 'name';
    el.sort.value = state.sortBy;
    if (info.prefs?.showNotes !== false) {
      state.showNotes = true;
      el.notesPanel.classList.add('is-open');
      el.btnNotes.classList.add('is-on');
    }
    renderRoots();
    renderRecents();
    el.cacheInfo.textContent = `Cache: ${fmtBytes(info.cacheBytes)}`;
    showEmpty(state.roots.length === 0);
    if (state.roots.length) await rescan({ quiet: true });
  } catch (e) {
    toast(`Startup failed: ${cleanError(e)}`, 'err', 6000);
  }
})();
