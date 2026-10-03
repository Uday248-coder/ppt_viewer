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
  crumbs: $('crumbs'),
  repoList: $('repo-list'),
  rootList: $('root-list'),
  recentList: $('recent-list'),
  cacheInfo: $('cache-info'),
  appVersion: $('app-version'),
  emptyVersion: $('empty-version'),
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
  btnPresenter: $('btn-presenter'),

  stage: $('stage'),
  slideWrap: $('slide-wrap'),
  slideImg: $('slide-img'),
  slideSpinner: $('slide-spinner'),
  zoomBadge: $('zoom-badge'),
  btnZoomReset: $('btn-zoom-reset'),
  mediaLayer: $('media-layer'),
  mediaBadge: $('media-badge'),
  btnMediaMute: $('btn-media-mute'),

  notesPanel: $('notes-panel'),
  notesBody: $('notes-body'),
  notesSlideNo: $('notes-slide-no'),
  notesTab: $('notes-tab'),
  btnNotesCollapse: $('btn-notes-collapse'),

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
  presentMedia: $('present-media'),
  presentNotes: $('present-notes'),
  presentBar: $('present-bar'),
  presentTitle: $('present-title'),
  presentCounter: $('present-counter'),
  presentProgressFill: $('present-progress-fill'),
  pPrev: $('p-prev'),
  pNext: $('p-next'),
  pGrid: $('p-grid'),
  pNotes: $('p-notes'),
  pMute: $('p-mute'),
  pTimer: $('p-timer'),
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
  repos: [],
  tree: [],
  scope: null,
  collapsed: new Set(),
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
  mediaMuted: false,
  prefillRunning: false,
  prefillToken: 0,
  mediaIndex: null,
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

/* The renderer runs with no Node integration, so it needs its own path maths.
   Both separators are accepted because Windows accepts both. */
function pathRelative(child, parent) {
  const c = String(child).replace(/[\\/]+$/, '').split(/[\\/]/);
  const p = String(parent).replace(/[\\/]+$/, '').split(/[\\/]/);
  if (!c[0] || !p[0] || c[0].toLowerCase() !== p[0].toLowerCase()) return null;
  let i = 0;
  while (i < c.length && i < p.length && c[i].toLowerCase() === p[i].toLowerCase()) i += 1;
  // Any parent segments not matched are levels to climb back up.
  const up = new Array(Math.max(0, p.length - i)).fill('..');
  return [...up, ...c.slice(i)].join('/');
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

/** Folder a deck sits in, relative to its repository, for grouping headers. */
function folderOf(deck) {
  if (!deck || !deck.rel) return (deck && deck.dirName) || '';
  const i = deck.rel.lastIndexOf('/');
  return i === -1 ? '' : deck.rel.slice(0, i);
}

/** "repo / folder" shown under the deck name, so its place is never a mystery. */
function deckPlace(deck) {
  const repo = deck.repoName || '';
  const folder = folderOf(deck);
  if (repo && folder) return `${repo}/${folder}`;
  return repo || folder || shortPath(deck.dir);
}

/**
 * Repositories found under the folders in the library, including sub-repos
 * checked out inside a master repo. Clicking one narrows the list to it.
 */
function renderRepos() {
  el.repoList.innerHTML = '';
  if (!state.repos.length) {
    const li = document.createElement('li');
    li.className = 'deck-list-empty repo-list-empty';
    li.textContent = 'No repositories found';
    el.repoList.appendChild(li);
    return;
  }

  el.repoList.appendChild(
    repoRow({ key: null, name: 'All folders', branch: null, deckCount: state.decks.length })
  );

  // Only outermost repositories get their own row; a sub-repo is rendered
  // inside the repository that contains it, so it appears exactly once.
  for (const repo of state.repos) {
    const contained = state.repos.some((r) => r.path !== repo.path && isInside(repo.path, r.path));
    if (contained) continue;

    const li = repoRow(repo);
    // A master repo contains its sub-repos; showing the containment makes a
    // deck's owner make sense at a glance. They need their own list to stack.
    const subs = state.repos.filter((c) => c.path !== repo.path && isInside(c.path, repo.path));
    if (subs.length) {
      const ul = document.createElement('ul');
      ul.className = 'repo-sublist';
      for (const child of subs) ul.appendChild(repoRow({ ...child, isSub: true }));
      li.appendChild(ul);
    }
    el.repoList.appendChild(li);
  }
}

function repoRow(repo) {
  const li = document.createElement('li');
  const active = (state.scope || null) === (repo.path ?? null);
  li.className = `repo-item${active ? ' is-active' : ''}${repo.isSub ? ' is-sub' : ''}`;
  li.title = repo.path || 'Every folder in the library';
  li.dataset.repoPath = repo.path || '';

  // Rows are appended to this holder, which may be the row itself or a nested
  // list inside it, so that adding a branch badge never disturbs the layout.
  const parts = document.createElement('span');
  parts.className = 'repo-row-main';
  li.appendChild(parts);

  const nm = document.createElement('span');
  nm.className = 'repo-name';
  nm.textContent = repo.name;
  parts.appendChild(nm);

  if (repo.branch) {
    const b = document.createElement('span');
    b.className = 'repo-branch';
    b.textContent = repo.branch;
    parts.appendChild(b);
  } else if (repo.linked) {
    const b = document.createElement('span');
    b.className = 'repo-branch repo-branch--muted';
    b.textContent = 'linked';
    parts.appendChild(b);
  }

  const c = document.createElement('span');
  c.className = 'repo-count';
  c.textContent = repo.deckCount;
  parts.appendChild(c);

  li.addEventListener('click', (e) => {
    // A sub-repo row is nested inside its parent's row, so without this a
    // click on the sub-repo would also select the master.
    e.stopPropagation();
    state.scope = repo.path || null;
    applyFilter();
    renderRepos();
  });
  return li;
}

function isInside(child, parent) {
  const rel = pathRelative(child, parent);
  return !!rel && !rel.startsWith('..');
}

/** Where the current view sits: library > repository. */
function renderCrumbs() {
  el.crumbs.innerHTML = '';
  const total = state.filtered.length;

  const mk = (label, onClick, current) => {
    const b = document.createElement('button');
    b.className = `crumb${current ? ' is-current' : ''}`;
    b.textContent = label;
    if (onClick) b.addEventListener('click', onClick);
    el.crumbs.appendChild(b);
    return b;
  };

  const scope = state.scope ? state.repos.find((r) => r.path === state.scope) : null;
  const scopeName = scope ? scope.name : null;

  mk(scopeName || 'All folders', () => {
    state.scope = null;
    applyFilter();
    renderRepos();
  }, !scopeName);

  if (scopeName) {
    const sep = document.createElement('span');
    sep.className = 'crumb-sep';
    sep.textContent = '/';
    el.crumbs.appendChild(sep);
    const folders = [...new Set(state.filtered.map((d) => folderOf(d)).filter(Boolean))];
    mk(folders.length === 1 ? folders[0] : 'any folder', null, true);
  }

  const count = document.createElement('span');
  count.className = 'crumb-count';
  count.textContent = state.query
    ? `${total} matching`
    : total
      ? `${total} deck${total === 1 ? '' : 's'}`
      : state.roots.length
        ? 'none here'
        : 'no folders added';
  el.crumbs.appendChild(count);
}

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
    el.repoList.appendChild;
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
  if (state.scope) {
    const scope = state.scope.toLowerCase();
    list = list.filter((d) => (d.repo || '').toLowerCase() === scope);
  }
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

/**
 * Groups the visible decks by repository.
 *
 * Grouping only appears when it actually says something: with a single
 * repository it would be pure noise, so the list stays flat instead.
 */
function groupDecks(list) {
  const repoKeys = [...new Set(list.map((d) => d.repo || ''))];
  if (repoKeys.length < 2) return null;
  const groups = [];
  for (const key of repoKeys) {
    const decks = list.filter((d) => (d.repo || '') === key);
    groups.push({ key, name: repoNameOf(decks[0], key), decks });
  }
  return groups;
}

function repoNameOf(deck, repoKey) {
  if (deck && deck.repoName) return deck.repoName;
  if (!repoKey) return 'Outside any repository';
  const bits = repoKey.split(/[\\/]/).filter(Boolean);
  return bits[bits.length - 1] || repoKey;
}

function renderDeckList() {
  el.deckList.innerHTML = '';
  const list = state.filtered;
  renderCrumbs();

  if (!list.length) {
    const d = document.createElement('div');
    d.className = 'deck-list-empty';
    d.textContent = state.roots.length
      ? state.query
        ? 'Nothing matches that search.'
        : state.scope
          ? 'No presentations in this repository.'
          : 'No .ppt, .pptx or .ppsx files in these folders.'
      : 'Add a folder to get started.';
    el.deckList.appendChild(d);
    return;
  }

  // Built off-document and appended in one go: a folder of a few thousand
  // decks would otherwise trigger a layout pass per row.
  const frag = document.createDocumentFragment();
  const groups = groupDecks(list);

  const addDecks = (decks) => {
    for (const deck of decks) frag.appendChild(deckRow(deck));
  };

  if (!groups) {
    addDecks(list);
  } else {
    for (const g of groups) {
      frag.appendChild(groupHeader(g.name, g.decks.length, g.key));
      const repo = state.repos.find((r) => r.path === g.key);
      // Second level: folder inside the repository, but only where it varies.
      const byFolder = new Map();
      for (const deck of g.decks) {
        const f = folderOf(deck);
        if (!byFolder.has(f)) byFolder.set(f, []);
        byFolder.get(f).push(deck);
      }
      if (byFolder.size > 1) {
        for (const [folder, decks] of byFolder) {
          if (!folder) continue;
          frag.appendChild(folderHeader(folder, decks.length));
          addDecks(decks);
        }
      } else {
        addDecks(g.decks);
      }
      if (repo && !repo.branch && repo.linked) {
        frag.appendChild(folderHeader('linked checkout', 0, true));
      }
    }
  }

  el.deckList.appendChild(frag);
}

function groupHeader(text, count, repoPath) {
  const h = document.createElement('div');
  h.className = 'group-head group-head--repo';
  const isCollapsed = state.collapsed.has(repoPath);
  const caret = document.createElement('span');
  caret.className = 'group-caret';
  caret.textContent = isCollapsed ? '▸' : '▾';
  h.appendChild(caret);
  const label = document.createElement('span');
  label.className = 'group-label';
  label.textContent = text;
  h.appendChild(label);
  const c = document.createElement('span');
  c.className = 'group-count';
  c.textContent = count;
  h.appendChild(c);
  h.addEventListener('click', () => {
    if (state.collapsed.has(repoPath)) state.collapsed.delete(repoPath);
    else state.collapsed.add(repoPath);
    renderDeckList();
  });
  return h;
}

function folderHeader(text, count, muted = false) {
  const h = document.createElement('div');
  h.className = `group-head group-head--folder${muted ? ' is-muted' : ''}`;
  const label = document.createElement('span');
  label.className = 'group-label';
  label.textContent = text;
  h.appendChild(label);
  if (count) {
    const c = document.createElement('span');
    c.className = 'group-count';
    c.textContent = count;
    h.appendChild(c);
  }
  return h;
}

function deckRow(deck) {
  const row = document.createElement('div');
  row.className = 'deck-row';
  if (state.collapsed.has(deck.repo || '')) row.hidden = true;
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
  pth.textContent = deckPlace(deck);
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
  return row;
}

function showAppVersion(v) {
  el.appVersion.textContent = v ? `v${v}` : 'v—';
  el.appVersion.title = v ? `PPT Viewer ${v}` : '';
  el.emptyVersion.textContent = v || '—';
}

async function refreshInfo() {
  const info = await api.app.info();
  state.roots = info.roots || [];
  state.recents = info.recents || [];
  showAppVersion(info.version);
  renderRepos();
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
    state.repos = res.repos || [];
    state.tree = res.tree || [];
    // A repository that no longer exists must not keep filtering the list.
    if (state.scope && !state.repos.some((r) => r.path === state.scope)) state.scope = null;
    applyFilter();
    renderRepos();
    showEmpty(false);
    if (!quiet) {
      const secs = (res.ms / 1000).toFixed(1);
      const n = res.decks.length;
      let msg = `Found ${n} deck${n === 1 ? '' : 's'} in ${secs}s`;
      if (state.repos.length > 1) {
        msg += ` · ${state.repos.length} repositor${state.repos.length === 1 ? 'y' : 'ies'}`;
      }
      toast(msg, 'ok');
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
  stopPrefill();
  inflight.clear();
  slideRetries.clear();
  busy(true, 'Opening deck…', shortPath(p));
  el.slideSpinner.hidden = false;
  try {
    const deck = await api.deck.open(p);
    if (token !== openToken) return;

    state.deck = deck;
    state.index = 1;
    // Where this deck sits in its repository, so the viewer keeps the library's
    // sense of place instead of becoming a dead end.
    const libDeck = state.decks.find((d) => d.path === p);
    const place = libDeck && libDeck.repo ? `${libDeck.repoName}/${folderOf(libDeck)}` : '';

    state.recents = (await api.app.info()).recents || [];
    if (token !== openToken) return;
    renderRecents();
    renderDeckList();

    el.viewerTitle.textContent = deck.title || deck.name;
    el.viewerSub.textContent =
      `${deck.slideCount} slides · ${deck.widthPt}×${deck.heightPt} pt` +
      `${deck.fromCache ? ' · from cache' : ''}` +
      (place ? ` · ${place.replace(/\//g, ' › ')}` : '');
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

    // Only fill the rest of the deck in the background if it is not already
    // cached; a fully cached deck has nothing left to render.
    if (!deck.fromCache && deck.cached !== 'full') {
      setTimeout(() => {
        if (token === openToken) startPrefill();
      }, 900);
    }
  } catch (e) {
    if (token === openToken) {
      el.slideSpinner.hidden = true;
      toast(`Could not open deck: ${cleanError(e)}`, 'err', 5200);
    }
  } finally {
    if (token === openToken) busy(false);
  }
}

// Slide index -> the promise that will render it. Kept so that navigating to a
// slide whose render is already running waits for that render instead of
// painting a URL for a file that does not exist yet (which 404s and, because
// browsers do not retry a cached 404, leaves the slide blank forever).
const inflight = new Map();

/** Renders a group of not-yet-requested slides as one worker round trip. */
function startRenderGroup(indices) {
  const fresh = indices.filter((i) => !inflight.has(i));
  if (!fresh.length) return Promise.resolve();
  const p = api.deck
    .ensureSlides(state.deck.path, fresh)
    .catch((e) => toast(`Render failed: ${cleanError(e)}`, 'err', 4200))
    .finally(() => fresh.forEach((i) => inflight.delete(i)));
  fresh.forEach((i) => inflight.set(i, p));
  return p;
}

/**
 * Requests full-resolution renders and resolves once they are on disk.
 *
 * Superseded work is never started: holding the right arrow key queues renders
 * for slides the user passed long ago, and those renders are serialised behind
 * PowerPoint, so they delay the slide actually being looked at.
 */
function ensureFull(indices, token) {
  if (!state.deck) return Promise.resolve();
  const valid = (indices || []).filter((i) => i >= 1 && i <= state.deck.slideCount);
  if (!valid.length) return Promise.resolve();
  if (token !== undefined && token !== slideToken) return Promise.resolve();

  const waiting = valid.filter((i) => inflight.has(i)).map((i) => inflight.get(i));
  const fresh = valid.filter((i) => !inflight.has(i));
  if (!waiting.length && !fresh.length) return Promise.resolve();
  return Promise.all([...waiting, startRenderGroup(fresh)]);
}

/**
 * Last resort for a slide image that failed to load even though it was asked
 * for. Re-renders it and busts the URL, because a cached 404 is never retried
 * by the browser on its own.
 */
const slideRetries = new Map();

function retrySlide(img, index) {
  if (!index || !state.deck) return;
  const n = slideRetries.get(index) || 0;
  if (n >= 3) {
    toast(`Slide ${index} could not be rendered.`, 'err', 4200);
    return;
  }
  slideRetries.set(index, n + 1);
  inflight.delete(index);
  api.deck
    .ensureSlides(state.deck.path, [index])
    .then(() => {
      const u = fullUrl(index);
      if (u) img.setAttribute('src', `${u}?r=${Date.now()}`);
    })
    .catch(() => {});
}

/** Full-resolution slide URL. */
function fullUrl(i) {
  const s = state.deck.slides[i - 1];
  return s ? s.fullUrl : '';
}

/* ------------------------------------------------------------------ media */

/**
 * Draws the current slide's media on top of the rendered still.
 *
 * PowerPoint bakes a video's poster frame into the exported PNG, so the media
 * element is positioned over that same frame using the shape's own geometry
 * read from the package - which means it lines up exactly, including when the
 * shape is cropped or rotated.
 */
function syncMedia(index) {
  const items = (state.deck && state.deck.media && state.deck.media[index]) || [];
  // Re-entering a slide has to start its media again; leaving it should not
  // leave a video sitting at its last frame for the next visit.
  const fresh = index !== state.mediaIndex;
  state.mediaIndex = index;
  for (const layer of [el.mediaLayer, el.presentMedia]) {
    if (!layer) continue;
    // Reusing the elements keeps a video from reloading on every slide change.
    const existing = [...layer.children];
    while (existing.length > items.length) {
      const gone = existing.pop();
      const node = gone.firstElementChild;
      if (node && node.tagName.toLowerCase() !== 'img' && typeof node.pause === 'function') {
        node.pause();
      }
      gone.remove();
    }
    while (existing.length < items.length) {
      const node = document.createElement('div');
      node.className = 'media-item';
      layer.appendChild(node);
      existing.push(node);
    }
    items.forEach((item, i) => placeMedia(existing[i], item, layer === el.presentMedia, fresh));
  }
  updateMediaBadge(items);
}

function placeMedia(host, item, isPresent, fresh) {
  const tag = item.kind === 'gif' ? 'img' : item.kind === 'audio' ? 'audio' : 'video';
  if (host.dataset.kind !== tag) {
    host.textContent = '';
    host.dataset.kind = tag;
  }
  const rect = item.rect || {};
  host.style.left = `${(rect.left || 0) * 100}%`;
  host.style.top = `${(rect.top || 0) * 100}%`;
  host.style.width = `${(rect.width == null ? 1 : rect.width) * 100}%`;
  host.style.height = `${(rect.height == null ? 1 : rect.height) * 100}%`;
  host.style.transform = item.rotation ? `rotate(${item.rotation}deg)` : '';
  host.style.zIndex = item.background ? '0' : '2';

  let node = host.firstElementChild;
  if (!node || node.tagName.toLowerCase() !== tag) {
    host.textContent = '';
    node = document.createElement(tag);
    if (tag !== 'img') {
      node.playsInline = true;
      node.preload = 'auto';
      node.controls = !isPresent;
    }
    host.appendChild(node);
  }
  const media = node;
  media.className = 'media-node';

  if (media.getAttribute('src') !== item.url) media.setAttribute('src', item.url);

  if (tag === 'audio') {
    host.classList.add('is-audio');
    if (!host.querySelector('.audio-chip')) {
      const chip = document.createElement('div');
      chip.className = 'audio-chip';
      chip.textContent = item.autoplay ? '♪' : '♪ click to play';
      host.appendChild(chip);
    }
  } else {
    host.classList.remove('is-audio');
  }
  if (tag === 'video') {
    media.loop = !!item.loop;
    media.muted = !!state.mediaMuted;
  }

  if (item.autoplay) {
    if (tag === 'img') {
      // A GIF animates on its own; nothing to start.
    } else if (fresh || media.paused) {
      // Re-entering the slide rewinds, so a video plays from the top every
      // time rather than resuming from wherever it stopped.
      try {
        media.currentTime = (item.startMs || 0) / 1000;
      } catch {
        /* seeking before metadata arrives is not worth surfacing */
      }
      host.classList.remove('needs-click');
      const attempt = tag === 'video' ? media.play() : null;
      // Autoplay can still be refused, e.g. a decode the first frame cannot
      // reach. The click affordance below is the fallback.
      if (attempt && attempt.catch) attempt.catch(() => host.classList.add('needs-click'));
    }
  }

  if (item.endMs && tag === 'video' && !media.dataset.endBound) {
    // PowerPoint's play command carries a duration; honouring it means the clip
    // stops where the author said it should.
    const stopAt = item.endMs / 1000;
    media.dataset.endBound = '1';
    media.addEventListener('timeupdate', () => {
      if (media.currentTime >= stopAt) media.pause();
    });
  }

  if (!host.dataset.bound) {
    host.dataset.bound = '1';
    host.addEventListener('click', (e) => {
      const m = host.firstElementChild;
      if (!m || m.tagName.toLowerCase() === 'img') return;
      e.stopPropagation();
      if (m.paused) m.play().catch(() => {});
      else m.pause();
    });
  }
}

function stopAllMedia() {
  for (const layer of [el.mediaLayer, el.presentMedia]) {
    if (!layer) continue;
    for (const host of layer.children) {
      const node = host.firstElementChild;
      if (node && node.tagName.toLowerCase() !== 'img' && typeof node.pause === 'function') {
        node.pause();
      }
    }
  }
}

function toggleMute() {
  state.mediaMuted = !state.mediaMuted;
  for (const layer of [el.mediaLayer, el.presentMedia]) {
    if (!layer) continue;
    for (const host of layer.children) {
      const node = host.firstElementChild;
      if (node && node.tagName.toLowerCase() === 'video') {
        node.muted = state.mediaMuted;
        if (!state.mediaMuted && node.paused) node.play().catch(() => {});
      }
    }
  }
  el.btnMediaMute.classList.toggle('is-on', state.mediaMuted);
  el.pMute.classList.toggle('is-on', state.mediaMuted);
  el.btnMediaMute.title = state.mediaMuted ? 'Media muted - click to unmute' : 'Mute media';
  toast(state.mediaMuted ? 'Media muted' : 'Media unmuted');
}

function updateMediaBadge(items) {
  const videos = items.filter((i) => i.kind === 'video');
  const sounds = items.filter((i) => i.kind === 'audio');
  const gifs = items.filter((i) => i.kind === 'gif');
  const bits = [];
  if (videos.length) bits.push(`${videos.length} video${videos.length === 1 ? '' : 's'}`);
  if (sounds.length) bits.push(`${sounds.length} audio`);
  if (gifs.length) bits.push(`${gifs.length} animated image${gifs.length === 1 ? '' : 's'}`);
  el.mediaBadge.textContent = bits.join(' · ');
  el.mediaBadge.hidden = !bits.length;
  el.btnMediaMute.hidden = !videos.length && !sounds.length;
}

/* ------------------------------------------------------------------ prefill */

/**
 * Quietly renders the rest of the deck so any later jump is instant.
 *
 * Runs after the first slide is on screen and after thumbnails, and stops as
 * soon as the user leaves the deck. Main puts it on a low-priority lane so it
 * can never delay a render the user is waiting on.
 */
async function startPrefill() {
  const token = ++state.prefillToken;
  if (!state.deck) return;
  const deckPath = state.deck.path;
  state.prefillRunning = true;
  try {
    const r = await api.deck.prefill(deckPath);
    if (token !== state.prefillToken) return;
    state.prefillRunning = false;
    if (r.complete && state.deck && state.deck.path === deckPath) state.deck.cached = 'full';
  } catch {
    if (token === state.prefillToken) state.prefillRunning = false;
  }
}

function stopPrefill() {
  state.prefillToken += 1;
  state.prefillRunning = false;
  api.deck.stopPrefill().catch(() => {});
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

  // Whatever the previous slide was playing has to stop, or its audio follows
  // you onto the next slide.
  stopAllMedia();
  syncMedia(idx);

  if (state.presenting) {
    await ensureFull([idx], token);
    if (token !== slideToken) return;
    if (animate) applyTransition(el.presentImg, s ? s.effect : 0);
    el.presentImg.src = fullUrl(idx);
    el.presentCounter.textContent = `${idx} / ${n}`;
    el.presentProgressFill.style.width = `${(idx / n) * 100}%`;
    el.presentTitle.textContent = state.deck.title || state.deck.name;
    if (state.presentNotes) {
      el.presentNotes.textContent = s && s.notes ? s.notes : 'No speaker notes for this slide.';
    }
  } else {
    el.slideSpinner.hidden = false;
    await ensureFull([idx], token);
    if (token !== slideToken) return;
    if (animate) applyTransition(el.slideImg, s ? s.effect : 0);
    el.slideImg.src = fullUrl(idx);
    el.slideSpinner.hidden = true;
  }

  // Warm the neighbours so paging never stalls on a render. Fire-and-forget:
  // this must not delay the slide the user is looking at.
  ensureFull([idx + 1, idx + 2, idx - 1], token);

  pushPresenterState();
}

/** Keeps the presenter window in step with whatever the audience window shows. */
function pushPresenterState() {
  if (!state.deck) return;
  const idx = state.index;
  const cur = state.deck.slides[idx - 1];
  const next = state.deck.slides[idx];
  api.presenter.state({
    deckName: state.deck.name,
    deckTitle: state.deck.title || state.deck.name,
    index: idx,
    slideCount: state.deck.slideCount,
    slideTitle: cur ? cur.title : '',
    notes: cur ? cur.notes : '',
    curUrl: cur ? cur.fullUrl : '',
    nextUrl: next ? next.fullUrl : '',
    hasNext: !!next,
    aspect: state.deck.aspect || 4 / 3,
    muted: state.mediaMuted,
  });
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
  // The present frame is sized from the same ratio so its media lines up.
  if (el.presentStage) el.presentStage.style.setProperty('--ar', String(state.deck.aspect || 4 / 3));

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
  const frag = document.createDocumentFragment();
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
    // A long deck should not decode every thumbnail at once.
    img.loading = 'lazy';
    img.decoding = 'async';
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
    frag.appendChild(d);
  });
  el.filmstrip.appendChild(frag);
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
  const frag = document.createDocumentFragment();
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
    img.loading = 'lazy';
    img.decoding = 'async';
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
    frag.appendChild(cell);
  });
  el.gridBody.appendChild(frag);
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

// Elapsed talk time, shared by the slideshow bar and the presenter window.
let timerHandle = null;
let timerStartedAt = 0;
let presenterWanted = false;

function fmtElapsed(ms) {
  const total = Math.max(0, Math.floor(ms / 1000));
  const h = Math.floor(total / 3600);
  const m = Math.floor((total % 3600) / 60);
  const s = total % 60;
  const mm = h ? String(m).padStart(2, '0') : String(m);
  return `${h ? `${h}:` : ''}${mm}:${String(s).padStart(2, '0')}`;
}

function paintTimer() {
  const elapsed = Date.now() - timerStartedAt;
  if (el.pTimer) el.pTimer.textContent = fmtElapsed(elapsed);
  if (presenterWanted) api.presenter.tick(elapsed);
}

function startTimer() {
  if (timerHandle) return;
  timerStartedAt = Date.now();
  timerHandle = setInterval(paintTimer, 500);
  paintTimer();
}

function resetTimer() {
  timerStartedAt = Date.now();
  paintTimer();
}

function stopTimer() {
  if (timerHandle) clearInterval(timerHandle);
  timerHandle = null;
}

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
  // Render a generous window so a live presentation never stalls. One batched
  // call rather than a request per slide.
  busy(true, 'Preparing slideshow…', 'Rendering upcoming slides', { done: 0, total: 1 });
  const ahead = [];
  for (let k = 1; k <= 8; k++) ahead.push(state.index + k);
  try {
    await api.deck.ensureSlides(state.deck.path, ahead);
  } catch (e) {
    toast(`Preload: ${cleanError(e)}`, 'err');
  }
  busy(false);
  startTimer();
}

function exitPresent() {
  state.presenting = false;
  el.present.hidden = true;
  stopAllMedia();
  stopTimer();
  presenterWanted = false;
  el.btnPresenter.classList.remove('is-on');
  api.presenter.close();
  api.window.exitPresent();
  fitSlide();
}

function presentVisible() {
  return !el.present.hidden;
}

function togglePresenter() {
  if (presenterWanted) {
    presenterWanted = false;
    el.btnPresenter.classList.remove('is-on');
    api.presenter.close();
    toast('Presenter window closed');
    return;
  }
  presenterWanted = true;
  el.btnPresenter.classList.add('is-on');
  // State can only be pushed once the window exists, so it goes after the open
  // resolves - otherwise the presenter would greet you with an empty deck.
  api.presenter
    .open()
    .then(() => {
      pushPresenterState();
      paintTimer();
    })
    .catch((e) => toast(`Presenter window: ${cleanError(e)}`, 'err', 4600));
  if (!timerHandle) startTimer();
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

/* ------------------------------------------------------------------ notes */

/**
 * Speaker notes are the widest thing in the viewer, so hiding them has to feel
 * like a deliberate control rather than a keyboard shortcut you have to
 * remember. Closing leaves a labelled tab behind, which keeps the panel
 * discoverable instead of making the stage silently wider.
 */
function setNotesOpen(open, { remember = true } = {}) {
  state.showNotes = open;
  el.notesPanel.classList.toggle('is-open', open);
  el.notesTab.hidden = open;
  el.btnNotes.classList.toggle('is-on', open);
  el.btnNotesCollapse.textContent = open ? '›' : '‹';
  el.btnNotesCollapse.title = open ? 'Hide notes (N)' : 'Show notes (N)';
  // The stage changes width, so the fit has to be recomputed.
  requestAnimationFrame(() => fitSlide());
  if (remember) api.app.setPref('showNotes', open);
}

function toggleNotes() {
  setNotesOpen(!state.showNotes);
}

/* ------------------------------------------------------------------ actions */

async function cacheDeck() {
  if (!state.deck) return;
  const d = state.deck;
  stopPrefill();
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
  stopPrefill();
  stopAllMedia();
  inflight.clear();
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
  el.btnNotesCollapse.addEventListener('click', toggleNotes);
  el.notesTab.addEventListener('click', () => setNotesOpen(true));
  el.btnCache.addEventListener('click', cacheDeck);
  el.btnPdf.addEventListener('click', exportPdf);
  el.btnPresent.addEventListener('click', enterPresent);
  el.btnPresenter.addEventListener('click', togglePresenter);
  el.btnZoomReset.addEventListener('click', resetZoom);
  el.btnPrev.addEventListener('click', () => goToSlide(state.index - 1));
  el.btnNext.addEventListener('click', () => goToSlide(state.index + 1));

  el.pPrev.addEventListener('click', () => goToSlide(state.index - 1));
  el.pNext.addEventListener('click', () => goToSlide(state.index + 1));
  el.pExit.addEventListener('click', exitPresent);
  el.pNotes.addEventListener('click', togglePresentNotes);
  el.pMute.addEventListener('click', toggleMute);
  el.btnMediaMute.addEventListener('click', toggleMute);

  el.slideImg.addEventListener('error', () => retrySlide(el.slideImg, state.index));
  el.presentImg.addEventListener('error', () => retrySlide(el.presentImg, state.index));

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
        case 'm': case 'M': e.preventDefault(); toggleMute(); return;
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
      case 'm': case 'M': e.preventDefault(); toggleMute(); break;
      case 'F5': e.preventDefault(); enterPresent(); break;
      case 'S': e.preventDefault(); togglePresenter(); break;
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
    const isDeck = (p) => /\.(pptx?|ppsx?|pptm|ppsm|potx|potm)$/i.test(p);
    const decks = paths.filter(isDeck);
    const folders = paths.filter((p) => !isDeck(p));
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

  // The presenter window asks the audience window to move, so that advancing
  // from either place cannot desynchronise the two.
  api.app.onNavigate(({ delta }) => {
    if (state.deck) goToSlide(state.index + delta);
  });

  api.app.onResetTimer(() => {
    if (timerHandle) resetTimer();
  });

  // Closed with its own title-bar button: stop pretending it is open, or the
  // toolbar keeps claiming a window that is not there.
  api.presenter.onExit(() => {
    presenterWanted = false;
    el.btnPresenter.classList.remove('is-on');
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
    el.notesTab.hidden = state.showNotes;
    el.btnNotesCollapse.textContent = state.showNotes ? '›' : '‹';
    renderRepos();
    renderRoots();
    renderRecents();
    showAppVersion(info.version);
    el.cacheInfo.textContent = `Cache: ${fmtBytes(info.cacheBytes)}`;
    showEmpty(state.roots.length === 0);
    if (state.roots.length) await rescan({ quiet: true });
  } catch (e) {
    toast(`Startup failed: ${cleanError(e)}`, 'err', 6000);
  }
})();
