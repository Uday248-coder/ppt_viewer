'use strict';

const fsp = require('fs/promises');
const path = require('path');
const { EventEmitter } = require('events');

const EXTENSIONS = new Set(['.ppt', '.pptx', '.pps', '.ppsx', '.pptm', '.ppsm', '.potx', '.potm']);

// Noise directories that would otherwise dominate a recursive scan.
const SKIP_DIRS = new Set([
  'node_modules', '.git', '.svn', '$recycle.bin', 'system volume information',
  'appdata', 'windows', 'program files', 'program files (x86)',
  'ppt viewer', 'slide-cache', 'dist', '.next', '__pycache__',
]);

const isHidden = (name) => name.startsWith('.') || name.startsWith('~$');

/**
 * Recursively discovers presentation files under one or more roots.
 *
 * Scans are cancellable and reported in batches, so a deep tree with thousands
 * of decks still renders results progressively instead of blocking the UI.
 */
class LibraryScanner extends EventEmitter {
  constructor() {
    super();
    this._cancel = false;
    this._scanning = false;
  }

  cancel() {
    this._cancel = true;
  }

  get scanning() {
    return this._scanning;
  }

  async scan(roots, { batchSize = 40, maxDepth = 12 } = {}) {
    this._cancel = false;
    this._scanning = true;
    const started = Date.now();
    const found = [];
    const errors = [];
    const seenDirs = new Set();
    let batch = [];

    const emit = () => {
      if (batch.length) {
        this.emit('batch', batch.slice());
        batch = [];
      }
    };

    const push = async (abs, stat) => {
      const rec = {
        path: abs,
        name: path.basename(abs),
        dir: path.dirname(abs),
        dirName: path.basename(path.dirname(abs)),
        size: stat.size,
        mtimeMs: stat.mtimeMs,
        ext: path.extname(abs).toLowerCase(),
      };
      found.push(rec);
      batch.push(rec);
      if (batch.length >= batchSize) emit();
    };

    const walk = async (dir, depth) => {
      if (this._cancel || depth > maxDepth) return;
      let real;
      try {
        real = await fsp.realpath(dir);
      } catch {
        return;
      }
      const key = real.toLowerCase();
      if (seenDirs.has(key)) return; // symlink/junction loop guard
      seenDirs.add(key);

      let entries;
      try {
        entries = await fsp.readdir(dir, { withFileTypes: true });
      } catch (e) {
        errors.push({ dir, message: e.message });
        return;
      }

      const dirs = [];
      for (const entry of entries) {
        if (this._cancel) return;
        const name = entry.name;
        if (isHidden(name)) continue;

        const abs = path.join(dir, name);
        if (entry.isDirectory()) {
          if (SKIP_DIRS.has(name.toLowerCase())) continue;
          dirs.push(abs);
        } else if (entry.isFile() || entry.isSymbolicLink()) {
          if (!EXTENSIONS.has(path.extname(name).toLowerCase())) continue;
          try {
            const st = await fsp.stat(abs);
            if (st.isFile()) await push(abs, st);
          } catch { /* file vanished mid-scan */ }
        }
      }
      // Breadth-ish: recurse after collecting this level's files so shallow
      // decks show up first instead of the deepest tree winning the race.
      for (const d of dirs) await walk(d, depth + 1);
    };

    for (const root of roots) {
      if (this._cancel) break;
      try {
        const st = await fsp.stat(root);
        if (st.isFile()) {
          if (EXTENSIONS.has(path.extname(root).toLowerCase())) await push(root, st);
          continue;
        }
      } catch {
        errors.push({ dir: root, message: 'Path does not exist' });
        continue;
      }
      await walk(root, 0);
    }

    emit();
    this._scanning = false;
    found.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
    return { files: found, errors, ms: Date.now() - started, cancelled: this._cancel };
  }
}

module.exports = { LibraryScanner, EXTENSIONS };
