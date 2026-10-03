'use strict';

const { spawn } = require('child_process');
const { EventEmitter } = require('events');
const path = require('path');
const fs = require('fs');
const os = require('os');

const SENTINEL = '@@PPV@@';

const PS_CANDIDATES = [
  path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe'),
  'powershell.exe',
  'pwsh.exe',
];

const DEFAULT_TIMEOUT = 90_000;

/**
 * Owns the long-lived PowerShell process that holds the PowerPoint COM object.
 *
 * PowerPoint instantiation costs ~1.7s, so the process is kept warm and reused
 * for every render. If it dies (or hangs on a modal dialog) it is transparently
 * restarted and in-flight requests are rejected rather than left dangling.
 */
class ComBridge extends EventEmitter {
  constructor(opts = {}) {
    super();
    this.logDir = opts.logDir || null;
    this.proc = null;
    this.shell = null;
    this.nextId = 1;
    this.pending = new Map();
    this.queue = Promise.resolve();
    this.buffer = '';
    this.stopping = false;
    this.restarts = 0;
    this.noise = [];
  }

  get alive() {
    return !!(this.proc && this.proc.exitCode === null && !this.proc.killed);
  }

  _resolveShell() {
    for (const c of PS_CANDIDATES) {
      if (path.isAbsolute(c) && fs.existsSync(c)) return c;
    }
    return 'powershell.exe';
  }

  _scriptPath() {
    // Inside a packaged app the script lives in app.asar.unpacked, because an
    // external powershell.exe cannot read through the asar virtual filesystem.
    let p = path.join(__dirname, '..', 'worker', 'render-worker.ps1');
    if (p.includes(`app.asar${path.sep}`)) {
      p = p.replace(`app.asar${path.sep}`, `app.asar.unpacked${path.sep}`);
    }
    return p;
  }

  start() {
    if (this.alive) return;

    this.shell = this._resolveShell();
    const script = this._scriptPath();
    if (!fs.existsSync(script)) {
      throw new Error(`Render worker not found at ${script}`);
    }

    this.stopping = false;
    this.buffer = '';
    this.noise = [];

    // -NoProfile/-NonInteractive keep boot fast and stop PowerShell from blocking
    // on its own prompts. -STA matches PowerPoint's COM threading expectation.
    const args = ['-NoProfile', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass', '-File', script];

    this.proc = spawn(this.shell, args, {
      windowsHide: true,
      stdio: ['pipe', 'pipe', 'pipe'],
      cwd: os.tmpdir(),
    });

    this.proc.stdout.setEncoding('utf8');
    this.proc.stdout.on('data', (chunk) => this._onStdout(chunk));

    this.proc.stderr.setEncoding('utf8');
    this.proc.stderr.on('data', (chunk) => {
      const t = String(chunk).trim();
      if (t) this.noise.push(t.slice(0, 800));
      this.emit('worker-stderr', t);
    });

    this.proc.on('error', (err) => {
      this.emit('worker-error', err);
      this._failAllPending(err.message);
    });

    this.proc.on('exit', (code, signal) => {
      const wasStopping = this.stopping;
      this.proc = null;
      this.emit('worker-exit', { code, signal });
      this._failAllPending(`Render worker exited (code=${code} signal=${signal})`);
      if (!wasStopping) {
        this.restarts += 1;
        this.emit('worker-restarting', { restarts: this.restarts });
        const delay = Math.min(5000, 250 * Math.pow(2, Math.min(this.restarts, 4)));
        setTimeout(() => {
          if (!this.alive && !this.stopping) {
            try { this.start(); } catch { /* surfaced via request failures */ }
          }
        }, delay);
      }
    });
  }

  _onStdout(chunk) {
    this.buffer += chunk;
    let nl;
    while ((nl = this.buffer.indexOf('\n')) !== -1) {
      const line = this.buffer.slice(0, nl).trim();
      this.buffer = this.buffer.slice(nl + 1);
      if (!line) continue;

      if (!line.startsWith(SENTINEL)) {
        // Stray output from PowerPoint/PowerShell - ignore but keep for diagnostics.
        this.noise.push(line.slice(0, 800));
        if (this.noise.length > 50) this.noise.shift();
        continue;
      }

      let msg;
      try { msg = JSON.parse(line.slice(SENTINEL.length)); } catch { continue; }
      const entry = this.pending.get(msg.id);
      if (!entry) {
        // id 0 is the worker's unsolicited progress channel; it deliberately
        // never matches a request, so anything arriving here is an event.
        if (msg.id === 0 && msg.event) this.emit('worker-event', msg.event);
        continue;
      }
      this.pending.delete(msg.id);
      clearTimeout(entry.timer);
      if (msg.ok) entry.resolve(msg.result);
      else entry.reject(new Error(msg.error || 'Unknown worker error'));
    }
  }

  _failAllPending(reason) {
    for (const [, entry] of this.pending) {
      clearTimeout(entry.timer);
      entry.reject(new Error(reason));
    }
    this.pending.clear();
  }

  _write(payload) {
    if (!this.proc || !this.proc.stdin.writable) {
      throw new Error('Render worker is not running');
    }
    this.proc.stdin.write(JSON.stringify(payload) + '\n', 'utf8');
  }

  /**
   * Sends a command and resolves with its result. Requests are serialised
   * because the worker holds a single PowerPoint presentation at a time.
   */
  request(cmd, data = {}, { timeout = DEFAULT_TIMEOUT, allowRestart = true } = {}) {
    const run = () =>
      new Promise((resolve, reject) => {
        if (!this.alive) {
          if (!allowRestart) {
            reject(new Error('Render worker is not running'));
            return;
          }
          try { this.start(); } catch (e) { reject(e); return; }
        }

        const id = this.nextId++;
        const timer = setTimeout(() => {
          this.pending.delete(id);
          // A timeout almost always means PowerPoint is blocked on a dialog we
          // cannot see. Recycling the worker is the only reliable recovery.
          try { this.hardRestart(); } catch { }
          reject(new Error(`Render worker timed out after ${timeout}ms (cmd: ${cmd})`));
        }, timeout);

        this.pending.set(id, { resolve, reject, timer });
        try {
          this._write({ id, cmd, data });
        } catch (err) {
          clearTimeout(timer);
          this.pending.delete(id);
          reject(err);
        }
      });

    const chained = this.queue.then(run, run);
    this.queue = chained.then(() => {}, () => {});
    return chained;
  }

  hardRestart() {
    if (this.proc) {
      try { this.proc.kill(); } catch { }
      this.proc = null;
    }
    this.restarts += 1;
    this.start();
  }

  async stop() {
    this.stopping = true;
    if (!this.proc) return;
    const p = this.proc;
    try {
      await Promise.race([
        this.request('shutdown', {}, { timeout: 4000, allowRestart: false }),
        new Promise((r) => setTimeout(r, 4000)),
      ]);
    } catch { /* worker will be killed below regardless */ }
    try { p.kill(); } catch { }
    this.proc = null;
    this._failAllPending('Render worker stopped');
  }

  /**
   * Synchronous last resort, for when the app is being torn down and there is
   * no longer time to negotiate a graceful shutdown.
   */
  killNow() {
    this.stopping = true;
    if (this.proc) {
      try { this.proc.kill(); } catch { }
      this.proc = null;
    }
    this._failAllPending('Render worker stopped');
  }
}

module.exports = { ComBridge };
