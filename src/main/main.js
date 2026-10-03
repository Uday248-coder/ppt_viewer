'use strict';

const { app, BrowserWindow, ipcMain, dialog, shell, protocol, net, nativeTheme } = require('electron');
const path = require('path');
const fs = require('fs');
const fsp = require('fs/promises');
const { Readable } = require('stream');

const { ComBridge } = require('./com-bridge');
const { SlideCache, KINDS, RENDER_VERSION, THUMB_W, FULL_W } = require('./cache');
const { LibraryScanner } = require('./library');
const { Settings } = require('./settings');
const { DeckService } = require('./deck-service');
const presenter = require('./presenter');

const IMAGE_SCHEME = 'pptv';

// Cached slide images are served through a private protocol rather than file://
// so the renderer can only ever reach files inside the slide cache.
protocol.registerSchemesAsPrivileged([
  {
    scheme: IMAGE_SCHEME,
    privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true, bypassCSP: false },
  },
]);

// Chromium's own mapping is unreliable for a custom scheme, and a media element
// refuses to play a response it cannot classify.
const MIME = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.apng': 'image/apng',
  '.mp4': 'video/mp4',
  '.m4v': 'video/mp4',
  '.mov': 'video/quicktime',
  '.webm': 'video/webm',
  '.mkv': 'video/x-matroska',
  '.mp3': 'audio/mpeg',
  '.m4a': 'audio/mp4',
  '.aac': 'audio/aac',
  '.wav': 'audio/wav',
  '.ogg': 'audio/ogg',
  '.flac': 'audio/flac',
};

let win = null;
let settings = null;
let cache = null;
let scanner = null;
let com = null;
let decks = null;
let engineState = { ready: false, comAvailable: false, version: null, error: null };

function slideUrl(key, kind, index) {
  return `${IMAGE_SCHEME}://slide/${key}/${kind}/slide-${String(index).padStart(4, '0')}.png`;
}

function mediaUrl(key, name) {
  return `${IMAGE_SCHEME}://media/${key}/media/${encodeURIComponent(name)}`;
}

function deckPayload(res) {
  const info = res.info;
  const aspect = info.heightPt ? info.widthPt / info.heightPt : 4 / 3;
  const mediaBySlide = {};
  for (const [index, items] of Object.entries(info.mediaSlides || {})) {
    mediaBySlide[index] = items.map((it) => ({ ...it, url: mediaUrl(res.key, it.file) }));
  }
  return {
    path: res.path,
    key: res.key,
    fromCache: res.fromCache,
    name: info.name || path.basename(res.path),
    title: info.title || '',
    slideCount: info.slideCount,
    widthPt: info.widthPt,
    heightPt: info.heightPt,
    aspect,
    cached: res.cached ? res.cached.level : 'none',
    thumbSize: info.thumbSize || { long: THUMB_W },
    fullSize: info.fullSize || { long: FULL_W },
    media: mediaBySlide,
    mediaFiles: info.mediaFiles || [],
    slides: (info.slides || []).map((s) => ({
      index: s.index,
      title: s.title || '',
      notes: s.notes || '',
      effect: s.effect,
      thumbUrl: slideUrl(res.key, KINDS.THUMB, s.index),
      fullUrl: slideUrl(res.key, KINDS.FULL, s.index),
    })),
  };
}

async function createWindow() {
  const w = settings.data.window;
  win = new BrowserWindow({
    width: w.width || 1440,
    height: w.height || 900,
    x: w.x ?? undefined,
    y: w.y ?? undefined,
    minWidth: 900,
    minHeight: 600,
    show: false,
    backgroundColor: '#0f1115',
    title: 'PPT Viewer',
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, '..', 'preload', 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      spellcheck: false,
      // A presenter deck is expected to make noise without being asked first,
      // exactly as it would in PowerPoint's own slideshow.
      autoplayPolicy: 'no-user-gesture-required',
    },
  });

  if (w.maximized) win.maximize();

  win.once('ready-to-show', () => win.show());

  const persistBounds = () => {
    if (!win || win.isDestroyed()) return;
    const b = win.getNormalBounds();
    settings.data.window = { ...settings.data.window, ...b, maximized: win.isMaximized() };
    settings.save().catch(() => {});
  };
  win.on('resize', persistBounds);
  win.on('move', persistBounds);
  win.on('close', persistBounds);
  win.on('closed', () => {
    win = null;
  });

  await win.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));
}

function send(channel, payload) {
  if (win && !win.isDestroyed()) win.webContents.send(channel, payload);
}

/** A byte range of a file as a web stream, so large media is never buffered. */
function fileStream(file, start, length) {
  return Readable.toWeb(fs.createReadStream(file, { start, end: start + length - 1 }));
}

function wireIpc() {
  ipcMain.handle('app:info', async () => {
    return {
      ...engineState,
      cacheDir: cache.root,
      cacheBytes: await cache.totalBytes(),
      roots: settings.data.libraryRoots,
      recents: settings.data.recents,
      prefs: settings.data.prefs,
      version: app.getVersion(),
      electron: process.versions.electron,
      node: process.versions.node,
    };
  });

  // --- library -----------------------------------------------------------
  ipcMain.handle('library:scan', async (_e, roots) => {
    const list = (Array.isArray(roots) && roots.length ? roots : settings.data.libraryRoots).filter(Boolean);
    const res = await scanner.scan(list);
    return {
      roots: list,
      ms: res.ms,
      errors: res.errors.slice(0, 20),
      cancelled: res.cancelled,
      repos: res.repos,
      tree: res.tree,
      decks: res.files.map((f) => ({
        ...f,
        key: SlideCache.keyFor(f.path, { size: f.size, mtimeMs: f.mtimeMs }),
      })),
    };
  });

  ipcMain.handle('library:addRoot', async (_e, dir) => {
    await settings.addRoot(dir);
    return settings.data.libraryRoots;
  });

  ipcMain.handle('library:removeRoot', async (_e, dir) => {
    await settings.removeRoot(dir);
    return settings.data.libraryRoots;
  });

  ipcMain.handle('library:chooseRoot', async () => {
    const r = await dialog.showOpenDialog(win, {
      title: 'Choose a folder to scan for presentations',
      properties: ['openDirectory', 'multiSelections', 'createDirectory'],
    });
    if (r.canceled || !r.filePaths.length) return { roots: settings.data.libraryRoots, added: [] };
    for (const p of r.filePaths) await settings.addRoot(p);
    return { roots: settings.data.libraryRoots, added: r.filePaths };
  });

  ipcMain.handle('library:reveal', async (_e, target) => {
    shell.showItemInFolder(target);
    return true;
  });

  ipcMain.handle('library:prune', async () => {
    const res = await scanner.scan(settings.data.libraryRoots, { batchSize: 10_000 });
    const live = res.files.map((d) => SlideCache.keyFor(d.path, { size: d.size, mtimeMs: d.mtimeMs }));
    return cache.prune(live);
  });

  ipcMain.handle('cache:clear', async () => {
    await cache.clearAll();
    return { ok: true, bytes: 0 };
  });

  // --- decks -------------------------------------------------------------
  ipcMain.handle('deck:open', async (_e, p) => {
    const res = await decks.openDeck(p);
    await settings.touchRecent({ path: res.path, name: path.basename(res.path), at: Date.now() });
    const payload = deckPayload(res);
    return payload;
  });

  ipcMain.handle('deck:ensureThumbs', async (_e, p) => decks.ensureThumbs(p));

  ipcMain.handle('deck:ensureSlides', async (_e, { path: p, indices }) => decks.ensureSlides(p, indices || []));

  ipcMain.handle('deck:cacheAll', async (_e, p) => {
    send('deck:progress', { path: p, phase: 'cache' });
    const r = await decks.cacheDeck(p);
    send('deck:progress', { path: p, phase: 'done', ...r });
    return r;
  });

  ipcMain.handle('deck:prefill', async (_e, p) => decks.prefill(p));

  ipcMain.handle('deck:stopPrefill', async () => ({ wasRunning: decks.stopPrefill() }));

  ipcMain.handle('deck:status', async (_e, p) => {
    const st = await fsp.stat(p);
    const key = SlideCache.keyFor(p, st);
    const status = await cache.status(key);
    return { key, level: status.level, slideCount: status.meta?.slideCount || 0 };
  });

  ipcMain.handle('deck:exportPdf', async (_e, p) => {
    const r = await dialog.showSaveDialog(win, {
      title: 'Export deck as PDF',
      defaultPath: path.join(app.getPath('documents'), `${path.basename(p, path.extname(p))}.pdf`),
      filters: [{ name: 'PDF', extensions: ['pdf'] }],
    });
    if (r.canceled || !r.filePath) return { cancelled: true };
    const out = await decks.exportPdf(p, r.filePath);
    return { cancelled: false, file: out.file };
  });

  ipcMain.handle('recents:clear', async () => {
    await settings.clearRecents();
    return [];
  });

  ipcMain.handle('prefs:set', async (_e, { key, value }) => {
    await settings.setPref(key, value);
    return settings.data.prefs;
  });

  // --- window ------------------------------------------------------------
  ipcMain.handle('win:present', async () => {
    if (!win) return false;
    win.setFullScreen(true);
    return true;
  });

  ipcMain.handle('win:exitPresent', async () => {
    if (!win) return false;
    win.setFullScreen(false);
    return true;
  });

  ipcMain.handle('win:isFullScreen', async () => (win ? win.isFullScreen() : false));

  // --- presenter ----------------------------------------------------------
  ipcMain.handle('presenter:open', async () => presenter.openPresenter());

  ipcMain.handle('presenter:close', async () => {
    presenter.closePresenter();
    return true;
  });

  ipcMain.handle('presenter:isOpen', async () => presenter.isOpen());

  ipcMain.handle('presenter:state', async (_e, s) => {
    presenter.pushState(s);
    return true;
  });

  ipcMain.handle('presenter:tick', async (_e, payload) => {
    presenter.pushTick(payload);
    return true;
  });

  // Navigation requested from the presenter window. It is forwarded rather than
  // applied so there is exactly one slide cursor for the whole app.
  ipcMain.handle('presenter:nav', async (_e, delta) => {
    send('deck:navigate', { delta: Number(delta) || 0 });
    return true;
  });

  ipcMain.handle('presenter:resetTimer', async () => {
    send('presenter:reset-timer', {});
    return true;
  });

  ipcMain.handle('win:openExternal', async (_e, url) => {
    // Never navigate to arbitrary URLs; only hand http(s) to the OS browser.
    if (/^https?:\/\//i.test(url)) await shell.openExternal(url);
    return true;
  });
}

async function probeEngine() {
  try {
    const d = await com.request('diagnose', {}, { timeout: 60_000 });
    engineState = { ready: true, comAvailable: !!d.comAvailable, version: d.version, error: null };
  } catch (e) {
    engineState = { ready: false, comAvailable: false, version: null, error: e.message };
  }
  return engineState;
}

app.whenReady().then(async () => {
  nativeTheme.themeSource = 'dark';

  settings = new Settings(app.getPath('userData'));
  cache = new SlideCache(SlideCache.defaultRoot(app.getPath('userData')));
  scanner = new LibraryScanner();
  com = new ComBridge();
  decks = new DeckService(com, cache, settings);

  await fsp.mkdir(cache.root, { recursive: true });

  com.on('worker-restarting', () => send('engine:status', { restarting: true }));
  com.on('worker-exit', (info) => send('engine:status', { restarting: false, exit: info }));
  decks.on('deck-opening', (d) => send('deck:opening', d));
  decks.on('render-progress', (d) => send('deck:progress', { path: decks.current?.path, phase: 'render', ...d }));

  // Serve cached slide images and media, refusing anything outside the cache
  // root. Range requests are honoured because a media element will not scrub a
  // video it cannot seek in.
  protocol.handle(IMAGE_SCHEME, async (request) => {
    let target = null;
    try {
      const url = new URL(request.url);
      const rel = decodeURIComponent(url.pathname).replace(/^\/+/, '');
      target = path.resolve(cache.root, rel);
      const rootResolved = path.resolve(cache.root);
      if (target !== rootResolved && !target.startsWith(rootResolved + path.sep)) {
        return new Response('Forbidden', { status: 403 });
      }

      const stat = await fsp.stat(target);
      if (!stat.isFile()) return new Response('Not found', { status: 404 });

      const type = MIME[path.extname(target).toLowerCase()] || 'application/octet-stream';
      const base = { 'Content-Type': type, 'Accept-Ranges': 'bytes' };

      if (request.method === 'HEAD') {
        return new Response(null, { headers: { ...base, 'Content-Length': String(stat.size) } });
      }

      const range = request.headers.get('Range');
      const match = range ? /^bytes=(\d*)-(\d*)$/.exec(range.trim()) : null;
      if (match) {
        let start = match[1] === '' ? null : Number(match[1]);
        let end = match[2] === '' ? null : Number(match[2]);
        if (start === null) {
          // A suffix range asks for the last N bytes.
          start = end === null ? 0 : Math.max(0, stat.size - end);
          end = stat.size - 1;
        } else {
          if (end === null || end >= stat.size) end = stat.size - 1;
        }
        if (start > end || start >= stat.size) {
          return new Response(null, {
            status: 416,
            headers: { ...base, 'Content-Range': `bytes */${stat.size}` },
          });
        }
        return new Response(fileStream(target, start, end - start + 1), {
          status: 206,
          headers: {
            ...base,
            'Content-Length': String(end - start + 1),
            'Content-Range': `bytes ${start}-${end}/${stat.size}`,
          },
        });
      }

      return new Response(fileStream(target, 0, stat.size), {
        headers: { ...base, 'Content-Length': String(stat.size) },
      });
    } catch (e) {
      return new Response(`Not found: ${e.message}`, { status: 404 });
    }
  });

  wireIpc();
  await createWindow();

  // Warm PowerPoint in the background so the first deck opens without delay.
  probeEngine().then((s) => send('engine:status', s));

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', async () => {
  try { presenter.disposePresenter(); } catch { }
  if (process.platform !== 'darwin') app.quit();
});

/**
 * Quitting has to release PowerPoint properly.
 *
 * The render worker holds the PowerPoint COM object, so simply exiting leaves
 * POWERPNT.EXE running with a deck open. Quitting therefore pauses once, asks
 * the worker to shut down, and only then exits - with a short fuse so a wedged
 * COM call can never make the app unquittable.
 */
let quitting = false;
app.on('before-quit', (event) => {
  if (quitting) return;
  quitting = true;
  event.preventDefault();

  const fuse = setTimeout(() => {
    try { com?.killNow(); } catch { }
    app.exit(0);
  }, 2500);
  (async () => {
    try { await decks?.shutdown(); } catch { }
    try { await settings?.save(); } catch { }
    clearTimeout(fuse);
    app.exit(0);
  })();
});
