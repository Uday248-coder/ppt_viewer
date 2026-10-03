'use strict';

const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');
const crypto = require('crypto');

// Only the long edge is fixed. The other dimension follows each deck's own
// aspect ratio, so a 4:3 deck renders as 1920x1440 instead of being squashed.
const THUMB_W = 480;
const FULL_W = 1920;

const KINDS = { THUMB: 'thumb', FULL: 'full' };

/**
 * Bumped whenever a change makes previously rendered slides wrong rather than
 * merely stale. v2 raised it because renders used to force a 1920x1080 export
 * regardless of the deck's own shape, which silently squashed every 4:3 deck.
 * Folding this into the key orphans the old entries, so Prune can drop them and
 * each deck re-renders itself correctly the next time it is opened.
 */
const RENDER_VERSION = 2;

/**
 * Content-addressed slide cache.
 *
 * A deck's identity is derived from its render version, path, byte size and
 * mtime, so editing a deck in PowerPoint automatically invalidates its cache
 * without any explicit "refresh" step and without ever hashing multi-megabyte
 * files.
 */
class SlideCache {
  constructor(root) {
    this.root = root;
    this.kindDirs = { [KINDS.THUMB]: 'thumb', [KINDS.FULL]: 'full' };
  }

  static defaultRoot(userDataDir) {
    return path.join(userDataDir, 'slide-cache');
  }

  static keyFor(absPath, stat) {
    const h = crypto.createHash('sha1');
    h.update(`v${RENDER_VERSION}|`);
    h.update(path.resolve(absPath).toLowerCase());
    h.update('|');
    h.update(String(stat.size));
    h.update('|');
    h.update(String(Math.floor(stat.mtimeMs)));
    return h.digest('hex').slice(0, 16);
  }

  deckDir(key) {
    return path.join(this.root, key);
  }

  metaPath(key) {
    return path.join(this.deckDir(key), 'meta.json');
  }

  slidePath(key, kind, index) {
    return path.join(this.deckDir(key), this.kindDirs[kind], `slide-${String(index).padStart(4, '0')}.png`);
  }

  /** Playable media lives beside the renders so it is covered by the same key. */
  mediaDir(key) {
    return path.join(this.deckDir(key), 'media');
  }

  mediaPath(key, name) {
    return path.join(this.mediaDir(key), path.basename(name));
  }

  async ensureDeckDir(key) {
    await fsp.mkdir(this.deckDir(key), { recursive: true });
  }

  async readMeta(key) {
    try {
      const raw = await fsp.readFile(this.metaPath(key), 'utf8');
      return JSON.parse(raw);
    } catch {
      return null;
    }
  }

  async writeMeta(key, meta) {
    await this.ensureDeckDir(key);
    const tmp = this.metaPath(key) + '.tmp';
    await fsp.writeFile(tmp, JSON.stringify(meta), 'utf8');
    await fsp.rename(tmp, this.metaPath(key));
  }

  /**
   * Reports what is already on disk for a deck, verifying file counts so a
   * half-finished or externally deleted cache never claims to be complete.
   */
  async status(key) {
    const meta = await this.readMeta(key);
    if (!meta || !meta.slideCount) return { level: 'none', meta: null, missing: 0 };

    const count = async (kind) => {
      const dir = path.join(this.deckDir(key), this.kindDirs[kind]);
      try {
        const files = await fsp.readdir(dir);
        return files.filter((f) => /^slide-\d+\.png$/i.test(f)).length;
      } catch {
        return 0;
      }
    };

    const [thumbs, fulls] = await Promise.all([count(KINDS.THUMB), count(KINDS.FULL)]);
    const n = meta.slideCount;
    const level = fulls >= n ? 'full' : thumbs >= n ? 'thumbs' : 'none';
    return {
      level,
      meta,
      thumbs: Math.min(thumbs, n),
      fulls: Math.min(fulls, n),
      missing: n - Math.min(fulls, n),
    };
  }

  /**
   * Moves PowerPoint's unpadded "Slide1.PNG" output into our own deterministic
   * zero-padded layout, so slide order never depends on lexical sorting.
   */
  async ingest(kind, exportedFiles, key) {
    const dest = path.join(this.deckDir(key), this.kindDirs[kind]);
    await fsp.mkdir(dest, { recursive: true });
    const written = [];
    for (const f of exportedFiles) {
      if (!f || !f.index) continue;
      const target = this.slidePath(key, kind, f.index);
      try {
        await fsp.copyFile(f.file, target);
        written.push(target);
      } catch { /* skip unreadable slide rather than failing the deck */ }
    }
    return written;
  }

  async writeSlide(kind, key, index, srcFile) {
    const target = this.slidePath(key, kind, index);
    await fsp.mkdir(path.dirname(target), { recursive: true });
    await fsp.copyFile(srcFile, target);
    return target;
  }

  async exists(kind, key, index) {
    try {
      await fsp.access(this.slidePath(key, kind, index));
      return true;
    } catch {
      return false;
    }
  }

  /**
   * True when every media file the deck references is still on disk. A deck
   * whose media was half-extracted must not advertise video it cannot play.
   */
  async mediaReady(key, files) {
    if (!files || !files.length) return true;
    for (const f of files) {
      try {
        const st = await fsp.stat(this.mediaPath(key, f.name));
        if (st.size !== f.bytes) return false;
      } catch {
        return false;
      }
    }
    return true;
  }

  async removeDeck(key) {
    await fsp.rm(this.deckDir(key), { recursive: true, force: true });
  }

  /**
   * Drops cache entries whose source file has changed or disappeared. Without
   * this the cache directory grows without bound as decks are edited.
   */
  async prune(liveKeys) {
    const live = new Set(liveKeys);
    let removed = 0;
    let bytes = 0;
    let entries = 0;
    let dirs;
    try {
      dirs = await fsp.readdir(this.root, { withFileTypes: true });
    } catch {
      return { removed: 0, bytes: 0, entries: 0 };
    }
    for (const d of dirs) {
      if (!d.isDirectory()) continue;
      entries += 1;
      const size = await dirSize(path.join(this.root, d.name));
      bytes += size;
      if (!live.has(d.name)) {
        await fsp.rm(path.join(this.root, d.name), { recursive: true, force: true });
        removed += 1;
      }
    }
    return { removed, bytes, entries };
  }

  async clearAll() {
    await fsp.rm(this.root, { recursive: true, force: true });
    await fsp.mkdir(this.root, { recursive: true });
  }

  async totalBytes() {
    try {
      const dirs = await fsp.readdir(this.root, { withFileTypes: true });
      let total = 0;
      for (const d of dirs) {
        if (d.isDirectory()) total += await dirSize(path.join(this.root, d.name));
      }
      return total;
    } catch {
      return 0;
    }
  }
}

async function dirSize(dir) {
  let total = 0;
  let entries;
  try {
    entries = await fsp.readdir(dir, { withFileTypes: true });
  } catch {
    return 0;
  }
  for (const e of entries) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) total += await dirSize(p);
    else {
      try {
        total += (await fsp.stat(p)).size;
      } catch { }
    }
  }
  return total;
}

module.exports = {
  SlideCache,
  RENDER_VERSION,
  THUMB_W,
  FULL_W,
  KINDS,
  dirSize,
};
