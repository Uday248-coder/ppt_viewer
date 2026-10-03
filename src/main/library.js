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

// Tree depth beyond which folders stop nesting. A deeply nested source tree
// should not turn the sidebar into a column of single-letter rows; past this
// point a folder is shown as one flat level named with its relative path.
const MAX_TREE_DEPTH = 5;

async function statOf(p) {
  try {
    return await fsp.stat(p);
  } catch {
    return null;
  }
}

/**
 * Works out whether a directory is a git repository, and if so which branch it
 * is on.
 *
 * Two shapes matter: a normal checkout has a `.git` directory, while a
 * submodule or `git worktree` has a `.git` *file* pointing elsewhere. Both are
 * real repositories as far as a user is concerned - the whole point of pointing
 * the viewer at a master repo is to have its sub-repos found too - so both are
 * reported, and only the branch is unavailable for the linked form until the
 * git directory is resolved.
 */
async function detectRepo(dir) {
  const git = await statOf(path.join(dir, '.git'));
  if (!git) return null;

  let branch = null;
  let detached = false;

  // A submodule or `git worktree` has a `.git` file whose only content is a
  // pointer to the real git directory; HEAD lives over there.
  let gitDir = path.join(dir, '.git');
  if (!git.isDirectory()) {
    try {
      const pointer = (await fsp.readFile(path.join(dir, '.git'), 'utf8')).trim();
      const target = /^gitdir:\s*(.+)$/.exec(pointer);
      if (target) gitDir = path.resolve(dir, target[1].trim());
    } catch {
      /* leave gitDir pointing at the .git file; the read below will just fail */
    }
  }

  try {
    const head = await fsp.readFile(path.join(gitDir, 'HEAD'), 'utf8');
    const ref = /^ref:\s*refs\/heads\/(.+)$/.exec(head.trim());
    if (ref) branch = ref[1];
    else if (head.trim()) detached = true;
  } catch {
    /* a repository with no readable HEAD is still a repository */
  }

  return { path: dir, name: path.basename(dir), branch, linked: !git.isDirectory(), detached };
}

/** Submodules declared in .gitmodules, whether or not they are checked out. */
async function readGitmodules(dir) {
  let raw;
  try {
    raw = await fsp.readFile(path.join(dir, '.gitmodules'), 'utf8');
  } catch {
    return [];
  }
  const out = [];
  let current = null;
  for (const line of raw.split(/\r?\n/)) {
    const trimmed = line.trim();
    const sub = /^\[submodule\s+"(.+)"\]$/.exec(trimmed);
    if (sub) {
      current = { name: sub[1], path: '', url: '' };
      out.push(current);
      continue;
    }
    if (!current) continue;
    const kv = /^(path|url)\s*=\s*(.+)$/.exec(trimmed);
    if (kv) current[kv[1]] = kv[2].trim();
  }
  return out.filter((m) => m.path);
}

function node(type, path_, name, extra = {}) {
  return { type, path: path_, name, children: [], deckCount: 0, ...extra };
}

/**
 * Recursively discovers presentation files under one or more roots.
 *
 * Scans are cancellable and reported in batches, so a deep tree with thousands
 * of decks still renders results progressively instead of blocking the UI.
 *
 * As well as the flat deck list, the scan reports the shape of what it found:
 * which roots are repositories, which repositories are nested inside them, and
 * how the folders and decks are arranged. That is what lets the viewer be
 * pointed at a master repository and still make sense of the sub-repositories
 * checked out inside it.
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
    const repos = [];
    const seenRepos = new Set();
    let batch = [];

    const emit = () => {
      if (batch.length) {
        this.emit('batch', batch.slice());
        batch = [];
      }
    };

    const push = async (abs, stat, owner) => {
      const rec = {
        path: abs,
        name: path.basename(abs),
        dir: path.dirname(abs),
        dirName: path.basename(path.dirname(abs)),
        size: stat.size,
        mtimeMs: stat.mtimeMs,
        ext: path.extname(abs).toLowerCase(),
        // Which repository this deck lives in, and where inside it.
        repo: owner ? owner.path : null,
        repoName: owner ? owner.name : null,
        rel: owner ? path.relative(owner.path, abs).split(path.sep).join('/') : null,
      };
      found.push(rec);
      batch.push(rec);
      if (batch.length >= batchSize) emit();
    };

    const walk = async (dir, depth, treeNode, owner) => {
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

      // A nested repository ends the previous repository's ownership: decks
      // below this point belong to the sub-repo, not the one containing it.
      let currentOwner = owner;
      if (depth > 0) {
        const repo = await detectRepo(dir);
        if (repo) {
          const repoKey = repo.path.toLowerCase();
          if (!seenRepos.has(repoKey)) {
            seenRepos.add(repoKey);
            repos.push(repo);
            const child = node('repo', repo.path, repo.name, {
              branch: repo.branch,
              linked: repo.linked,
              depth,
            });
            if (treeNode) treeNode.children.push(child);
            treeNode = child;
          }
          currentOwner = repo;
        }
      }

      let entries;
      try {
        entries = await fsp.readdir(dir, { withFileTypes: true });
      } catch (e) {
        errors.push({ dir, message: e.message });
        return;
      }

      const dirs = [];
      const filesHere = [];
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
            if (st.isFile()) {
              await push(abs, st, currentOwner);
              filesHere.push({
                path: abs,
                name,
                size: st.size,
                mtimeMs: st.mtimeMs,
                ext: path.extname(abs).toLowerCase(),
              });
            }
          } catch { /* file vanished mid-scan */ }
        }
      }

      // Folders with nothing in them are not worth showing in a sidebar.
      const ownCount = filesHere.length;
      if (ownCount && treeNode) {
        const folder = node('folder', dir, path.basename(dir), {
          depth,
          rel: currentOwner ? path.relative(currentOwner.path, dir).split(path.sep).join('/') : null,
        });
        folder.children = filesHere.map((f) =>
          node('deck', f.path, f.name, {
            size: f.size,
            mtimeMs: f.mtimeMs,
            ext: f.ext,
            rel: currentOwner ? path.relative(currentOwner.path, f.path).split(path.sep).join('/') : null,
          })
        );
        folder.deckCount = ownCount;
        treeNode.children.push(folder);
      }
      treeNode.deckCount += ownCount;

      // Breadth-ish: recurse after collecting this level's files so shallow
      // decks show up first instead of the deepest tree winning the race.
      for (const d of dirs) {
        if (this._cancel) return;
        let childNode = treeNode;
        // Past the depth limit a folder is shown flat rather than nested, so
        // the sidebar stays usable in a deep source tree.
        if (treeNode && depth + 1 > MAX_TREE_DEPTH) {
          const tail = node('folder', d, path.basename(d), { depth: depth + 1, flat: true });
          childNode = tail;
          treeNode.children.push(tail);
        }
        const before = childNode.deckCount;
        await walk(d, depth + 1, childNode, currentOwner);
        if (childNode.flat && childNode.deckCount === before) childNode.remove?.();
      }
    };

    const tree = [];
    for (const root of roots) {
      if (this._cancel) break;
      try {
        const st = await fsp.stat(root);
        if (st.isFile()) {
          if (EXTENSIONS.has(path.extname(root).toLowerCase())) await push(root, st, null);
          continue;
        }
      } catch {
        errors.push({ dir: root, message: 'Path does not exist' });
        continue;
      }

      const rootRepo = await detectRepo(root);
      const submodules = await readGitmodules(root);
      const rootNode = node('root', root, path.basename(root) || root, {
        isRepo: !!rootRepo,
        branch: rootRepo ? rootRepo.branch : null,
        linked: rootRepo ? rootRepo.linked : false,
        submodules: submodules.map((m) => ({
          ...m,
          present: !!seenDirs.has(path.join(root, m.path).toLowerCase()),
        })),
      });
      tree.push(rootNode);

      if (rootRepo) {
        const repoKey = rootRepo.path.toLowerCase();
        if (!seenRepos.has(repoKey)) {
          seenRepos.add(repoKey);
          repos.push(rootRepo);
        }
      }

      await walk(root, 0, rootNode, rootRepo);
    }

    emit();
    this._scanning = false;
    found.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));

    // Count decks per repository, including the root, for the sidebar summary.
    const counts = new Map();
    for (const d of found) {
      const key = d.repo ? d.repo.toLowerCase() : '';
      counts.set(key, (counts.get(key) || 0) + 1);
    }
    for (const r of repos) r.deckCount = counts.get(r.path.toLowerCase()) || 0;

    pruneEmpty(tree);

    return {
      files: found,
      errors,
      ms: Date.now() - started,
      cancelled: this._cancel,
      repos,
      tree,
    };
  }
}

/** Drops branches of the tree that lead nowhere, so the sidebar stays short. */
function pruneEmpty(nodes) {
  for (const n of nodes) {
    n.children = (n.children || []).filter((c) => c.type !== 'folder' || c.deckCount > 0 || c.children.length);
    pruneEmpty(n.children);
    if (n.type === 'folder' && n.deckCount === 0 && !n.children.length) n.children = [];
  }
}

module.exports = { LibraryScanner, EXTENSIONS, detectRepo, readGitmodules };