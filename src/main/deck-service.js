'use strict';

const fsp = require('fs/promises');
const path = require('path');
const os = require('os');
const { EventEmitter } = require('events');
const { SlideCache, KINDS, RENDER_VERSION, THUMB_W, FULL_W } = require('./cache');
const mediaReader = require('./media');

const HRESULT_TAIL = /\s*\|?\s*(exception from )?hresult:?\s*0x[0-9a-f]+\s*$/i;

/**
 * PowerPoint reports many failures as bare HRESULTs, which mean nothing to a
 * user. Where it does give real prose ("its file extension has changed"), that
 * wording is more specific than anything we would invent - so keep it.
 */
function humanizeError(err) {
  const raw = String((err && err.message) || '').trim();
  if (!raw) return 'PowerPoint could not complete the request.';

  // Decide whether this is a bare HRESULT by looking for real prose in front of
  // it, rather than guessing from string length.
  const prose = raw.replace(HRESULT_TAIL, '').replace(/[\s,|]+/g, ' ').trim();
  const hasProse = prose.split(' ').filter(Boolean).length >= 3;

  if (!hasProse) {
    if (/0x80cb4002|0x800a03ec/i.test(raw)) {
      return 'This file could not be read by PowerPoint — it may be corrupt, truncated, or not actually a presentation.';
    }
    if (/0x80040154/i.test(raw)) return 'PowerPoint is not installed on this machine.';
    if (/0x800a175d|0x80070bc9/i.test(raw)) {
      return 'PowerPoint automation is unavailable — Office may need to be activated or repaired. See the README.';
    }
    // Unknown code: stay vague rather than assert a cause we cannot confirm.
    return 'PowerPoint could not complete the request.';
  }

  const h = raw.toLowerCase();
  if (/password|encrypted/.test(h)) {
    return 'This deck is password-protected. Remove the password in PowerPoint, then reopen it.';
  }
  if (/0x800a175d|not licensed|deactivat|reduced functionality|automation is (blocked|unavailable)/.test(h)) {
    return 'PowerPoint automation is blocked. Reactivate Office, then reopen the viewer.';
  }
  if (/timed out|timeout/.test(h)) {
    return 'PowerPoint stopped responding (it may be showing a hidden dialog). It has been restarted — try again.';
  }
  if (/class not registered|no application is associated/.test(h)) {
    return 'PowerPoint is not installed on this machine.';
  }
  // Preserve PowerPoint's own explanation, dropping only the HRESULT tail.
  const cleaned = raw.replace(HRESULT_TAIL, '').trim();
  return cleaned || 'PowerPoint could not complete the request.';
}

function fail(err) {
  const e = new Error(humanizeError(err));
  e.cause = err;
  return e;
}

/**
 * Owns everything about "a deck is open": its identity, its cache state, and
 * the render pipeline that fills that cache.
 *
 * The COM worker holds exactly one PowerPoint presentation, so every operation
 * funnels through _ensureOpen(), which reopens the deck if the worker's current
 * presentation has drifted to a different file.
 *
 * Work is also split into two lanes. Anything the user is waiting on runs
 * interactively; filling the rest of the deck in the background runs at low
 * priority and is dropped the moment an interactive request arrives. Without
 * that split, a background prefill of a 300-slide deck would put a minute of
 * queued renders in front of a simple arrow-key press.
 */
class DeckService extends EventEmitter {
  constructor(comBridge, cache, settings) {
    super();
    this.com = comBridge;
    this.cache = cache;
    this.settings = settings;
    this.current = null; // { path, key, info }
    this._interactive = Promise.resolve();
    this._background = Promise.resolve();
    this._backgroundRunning = false;
    this._backgroundAbort = null;
    this.com.on('worker-event', (ev) => {
      if (ev && ev.phase === 'slides') this.emit('render-progress', ev);
    });
  }

  /** Runs ahead of background work; nothing prefill-related can delay it. */
  _serial(fn) {
    const run = () => fn().catch((e) => { throw fail(e); });
    const chained = this._interactive.then(run, run);
    this._interactive = chained.then(
      () => {},
      () => {}
    );
    return chained;
  }

  /** Yields to the interactive lane between every unit of work. */
  _backgroundSerial(fn) {
    const run = () => this._serial(fn);
    const chained = this._background.then(run, run);
    this._background = chained.then(
      () => {},
      () => {}
    );
    return chained;
  }

  async _statOf(absPath) {
    try {
      const st = await fsp.stat(absPath);
      if (!st.isFile()) throw new Error('Not a file');
      return st;
    } catch {
      throw new Error(`Deck not found: ${absPath}`);
    }
  }

  async _ensureOpen(absPath, key) {
    if (this.current && this.current.path === absPath) return this.current;
    const info = await this.com.request('open', { path: absPath }, { timeout: 180_000 });
    this.current = { path: absPath, key, info };
    return this.current;
  }

  /**
   * Opens a deck for viewing.
   *
   * If a complete cache already exists the deck opens with zero COM involvement,
   * which is what makes repeat launches feel instant. Otherwise the deck is
   * opened through PowerPoint and its structure is cached.
   */
  async openDeck(absPath, { preferCache = true } = {}) {
    return this._serial(async () => {
      const abs = path.resolve(absPath);
      const st = await this._statOf(abs);
      const key = SlideCache.keyFor(abs, st);
      const status = await this.cache.status(key);

      // Fully cached: serve from disk, never touch PowerPoint.
      if (preferCache && status.level === 'full' && status.meta) {
        const mediaOk = await this.cache.mediaReady(key, status.meta.mediaFiles);
        if (mediaOk) {
          this.current = null;
          return {
            path: abs,
            key,
            fromCache: true,
            info: status.meta,
            cached: status,
          };
        }
      }

      this.emit('deck-opening', { path: abs, name: path.basename(abs) });

      // Extracted before the deck is opened: PowerPoint locks any file it has
      // open, and the OOXML package cannot be read through that lock. Doing it
      // in this order also means the media is ready by the time the first slide
      // is drawn, so a video starts on the right frame.
      const previous = status.meta;
      const media = await this._extractMedia(abs, key, {
        slides: previous ? previous.mediaSlides : null,
        files: previous ? previous.mediaFiles : null,
        error: previous ? previous.mediaError : null,
      });

      const opened = await this._ensureOpen(abs, key);
      const info = opened.info;

      const meta = {
        path: abs,
        name: path.basename(abs),
        slideCount: info.slideCount,
        widthPt: info.widthPt,
        heightPt: info.heightPt,
        title: info.title || '',
        slides: info.slides || [],
        cachedAt: Date.now(),
        renderVersion: RENDER_VERSION,
        sourceMtimeMs: Math.floor(st.mtimeMs),
        sourceSize: st.size,
        // Recorded per deck: the long edge is fixed but the other dimension
        // follows the deck's own shape, so these are not global constants.
        thumbSize: previous?.thumbSize || { long: THUMB_W },
        fullSize: previous?.fullSize || { long: FULL_W },
        mediaSlides: media.slides,
        mediaFiles: media.files,
        mediaError: media.error || null,
      };
      await this.cache.writeMeta(key, meta);

      return { path: abs, key, fromCache: false, info: meta, cached: await this.cache.status(key) };
    });
  }

  /**
   * Unpacks embedded media and works out which slide each file belongs to.
   *
   * Runs through the worker but needs no PowerPoint, and any earlier result is
   * kept when the media is already unpacked and intact.
   */
  async _extractMedia(absPath, key, previous) {
    const alreadyHave = previous && previous.slides && previous.files;
    if (alreadyHave && (await this.cache.mediaReady(key, previous.files))) {
      return { slides: previous.slides, files: previous.files, error: previous.error || null };
    }

    const scratch = path.join(os.tmpdir(), `pptv-media-${process.pid}-${key}`);
    const copy = path.join(scratch, 'source.pptx');
    try {
      // PowerPoint locks any deck it has open, and the package cannot be read
      // through that lock. Extracting before the deck is opened is the normal
      // path; the throwaway copy is the fallback for when it is already open.
      let res;
      try {
        res = await this.com.request(
          'unpackMedia',
          { file: absPath, outDir: scratch },
          { timeout: 300_000 }
        );
      } catch (direct) {
        res = await this.com.request(
          'unpackMedia',
          { file: absPath, outDir: scratch, viaCopy: copy },
          { timeout: 300_000 }
        );
      }
      const parsed = await mediaReader.readUnpacked(scratch);

      const dest = this.cache.mediaDir(key);
      await fsp.mkdir(dest, { recursive: true });
      const files = [];
      for (const m of parsed.media) {
        await fsp.copyFile(path.join(scratch, 'media', m.name), this.cache.mediaPath(key, m.name));
        files.push({ name: m.name, kind: m.kind, bytes: m.bytes });
      }

      const slides = {};
      for (const [index, items] of parsed.slides) {
        slides[index] = items.map((it) => ({
          kind: it.kind,
          file: it.file,
          name: it.name || '',
          background: !!it.background,
          rect: it.rect,
          rotation: it.rotation,
          autoplay: !!it.autoplay,
          loop: !!it.loop,
          startMs: it.startMs || 0,
          endMs: it.endMs || null,
        }));
      }
      return { slides, files, error: null };
    } catch (e) {
      // A deck whose media cannot be unpacked is still perfectly viewable as
      // stills, so this is never fatal - but it is worth saying out loud,
      // because "the video did not play" with no explanation is maddening.
      const message = humanizeError(e);
      this.emit('media-warning', { path: absPath, message });
      return { slides: {}, files: [], error: message };
    } finally {
      await fsp.rm(scratch, { recursive: true, force: true }).catch(() => {});
    }
  }

  /** Renders any missing thumbnails. Cheap, so this is done eagerly. */
  async ensureThumbs(absPath, { onProgress } = {}) {
    return this._serial(async () => {
      const abs = path.resolve(absPath);
      const st = await this._statOf(abs);
      const key = SlideCache.keyFor(abs, st);
      const status = await this.cache.status(key);
      const meta = status.meta;
      if (!meta) throw new Error('Deck metadata is not cached; open the deck first.');

      if (status.level !== 'none') {
        return { key, thumbs: status.thumbs, slideCount: meta.slideCount, rendered: 0 };
      }

      const cur = await this._ensureOpen(abs, key);
      const tmp = path.join(os.tmpdir(), `pptv-thumb-${process.pid}-${key}`);
      const res = await this.com.request(
        'exportDir',
        { dir: tmp, long: THUMB_W },
        { timeout: 300_000 }
      );
      await this.cache.ingest(KINDS.THUMB, res.files, key);
      await fsp.rm(tmp, { recursive: true, force: true });
      onProgress?.({ done: meta.slideCount, total: meta.slideCount });

      return {
        key,
        thumbs: meta.slideCount,
        slideCount: meta.slideCount,
        rendered: meta.slideCount,
        info: cur.info,
        thumbSize: { width: res.width, height: res.height },
      };
    });
  }

  /**
   * Renders specific slides at full resolution. Only missing slides are rendered,
   * so revisiting a deck costs nothing.
   *
   * One worker call covers the whole burst: asking for slides one at a time
   * puts a protocol round trip in front of the user for each one.
   */
  async ensureSlides(absPath, indices, { onProgress } = {}) {
    const wanted = [...new Set((indices || []).map(Number).filter((n) => n >= 1))];
    if (!wanted.length) return { key: null, rendered: 0, alreadyHave: 0 };

    return this._serial(async () => {
      const abs = path.resolve(absPath);
      const st = await this._statOf(abs);
      const key = SlideCache.keyFor(abs, st);
      const status = await this.cache.status(key);
      if (!status.meta) throw new Error('Deck metadata is not cached; open the deck first.');

      const inRange = wanted.filter((n) => n <= status.meta.slideCount);
      const missing = [];
      for (const i of inRange) {
        if (!(await this.cache.exists(KINDS.FULL, key, i))) missing.push(i);
      }
      if (!missing.length) {
        return { key, rendered: 0, alreadyHave: inRange.length };
      }

      await this._ensureOpen(abs, key);
      const tmpDir = path.join(os.tmpdir(), `pptv-full-${process.pid}-${key}`);
      const res = await this.com.request(
        'exportSlides',
        { dir: tmpDir, long: FULL_W, indices: missing },
        { timeout: 600_000 }
      );
      const written = await this.cache.ingest(KINDS.FULL, res.files, key);
      await fsp.rm(tmpDir, { recursive: true, force: true }).catch(() => {});
      onProgress?.({ done: missing.length, total: missing.length });

      return {
        key,
        rendered: written.length,
        alreadyHave: inRange.length - missing.length,
        fullSize: { width: res.width, height: res.height },
      };
    });
  }

  /** Pre-renders the whole deck so it later opens with no PowerPoint at all. */
  async cacheDeck(absPath, { onProgress } = {}) {
    return this._serial(async () => {
      const abs = path.resolve(absPath);
      const st = await this._statOf(abs);
      const key = SlideCache.keyFor(abs, st);
      const status = await this.cache.status(key);
      if (!status.meta) throw new Error('Deck metadata is not cached; open the deck first.');

      const total = status.meta.slideCount;
      const missing = [];
      for (let i = 1; i <= total; i++) {
        if (!(await this.cache.exists(KINDS.FULL, key, i))) missing.push(i);
      }
      if (!missing.length) return { key, rendered: 0, total, complete: true };

      await this._ensureOpen(abs, key);
      const tmp = path.join(os.tmpdir(), `pptv-bulk-${process.pid}-${key}`);
      const res = await this.com.request('exportDir', { dir: tmp, long: FULL_W }, { timeout: 900_000 });
      // exportDir renders everything; ignore any already-cached slides when counting.
      const written = await this.cache.ingest(KINDS.FULL, res.files, key);
      await fsp.rm(tmp, { recursive: true, force: true }).catch(() => {});
      onProgress?.({ done: total, total });

      return { key, rendered: written.length, total, complete: true };
    });
  }

  /**
   * Quietly renders the rest of the deck in the background, so jumping to any
   * slide later is instant and the deck becomes fully offline-capable without
   * the user asking for it.
   *
   * Runs one small batch at a time and re-checks the abort flag between them, so
   * closing the deck or pressing an arrow key stops it promptly.
   */
  async prefill(absPath, { batchSize = 12, onProgress } = {}) {
    const abs = path.resolve(absPath);
    const st = await this._statOf(abs);
    const key = SlideCache.keyFor(abs, st);
    const status = await this.cache.status(key);
    if (!status.meta) return { key, rendered: 0, total: 0, stopped: true };

    const total = status.meta.slideCount;
    const missing = [];
    for (let i = 1; i <= total; i++) {
      if (!(await this.cache.exists(KINDS.FULL, key, i))) missing.push(i);
    }
    if (!missing.length) return { key, rendered: 0, total, complete: true };

    const abort = { stopped: false };
    this._backgroundAbort = abort;
    this._backgroundRunning = true;
    let done = 0;

    try {
      for (let at = 0; at < missing.length && !abort.stopped; at += batchSize) {
        if (abort.stopped || this._backgroundAbort !== abort) break;
        const batch = missing.slice(at, at + batchSize);
        const r = await this._backgroundSerial(async () => {
          // The user may have opened a different deck while this was queued.
          if (abort.stopped) return 0;
          if (this.current && this.current.path !== abs) return 0;
          return (await this._renderBatch(abs, key, batch)).length;
        });
        done += r;
        onProgress?.({ done, total: missing.length, running: !abort.stopped });
      }
    } catch {
      // A background failure is not worth interrupting the user over; the
      // slides it would have produced simply stay uncached.
    }

    if (this._backgroundAbort === abort) {
      this._backgroundRunning = false;
      this._backgroundAbort = null;
    }
    return { key, rendered: done, total: missing.length, stopped: abort.stopped, complete: done >= missing.length };
  }

  /** Renders one batch of full-resolution slides, returning what was written. */
  async _renderBatch(absPath, key, indices) {
    await this._ensureOpen(absPath, key);
    const tmpDir = path.join(os.tmpdir(), `pptv-prefill-${process.pid}-${key}`);
    const res = await this.com.request(
      'exportSlides',
      { dir: tmpDir, long: FULL_W, indices },
      { timeout: 600_000 }
    );
    const written = await this.cache.ingest(KINDS.FULL, res.files, key);
    await fsp.rm(tmpDir, { recursive: true, force: true }).catch(() => {});
    return written;
  }

  stopPrefill() {
    if (this._backgroundAbort) this._backgroundAbort.stopped = true;
    return this._backgroundRunning;
  }

  async exportPdf(absPath, outFile) {
    return this._serial(async () => {
      const abs = path.resolve(absPath);
      const st = await this._statOf(abs);
      const key = SlideCache.keyFor(abs, st);
      const out = path.resolve(outFile);
      if (!out.toLowerCase().endsWith('.pdf')) out += '.pdf';
      await this._ensureOpen(abs, key);
      const res = await this.com.request('exportPdf', { file: out }, { timeout: 300_000 });
      return { file: res.file };
    });
  }

  async close() {
    this.stopPrefill();
    return this._serial(async () => {
      if (this.current) {
        try { await this.com.request('close'); } catch { }
      }
      this.current = null;
    });
  }

  /** Releases the PowerPoint process - used when the app goes idle. */
  async shutdown() {
    this.stopPrefill();
    await this.close();
    await this.com.stop();
  }
}

module.exports = { DeckService, humanizeError };
