/**
 * mcp-word-bridge entrypoint
 *
 * Smart CLI that MCP hosts invoke via { "command": "node dist/server.js" }.
 * Ensures exactly one daemon is running, then proxies stdio ↔ HTTP.
 *
 * Startup protocol:
 *   1. Try exclusive lock on LOCK_FILE
 *   2. Got lock   → spawn daemon, wait for port, release lock, enter proxy mode
 *   3. Lock busy  → wait for port (another process is starting daemon), enter proxy mode
 *
 * Proxy mode:
 *   - Reads JSON-RPC from stdin, POSTs to daemon /mcp, writes responses to stdout
 *   - Maintains SSE stream for server-initiated messages
 *   - On stdin EOF → sends DELETE to close session, then exits
 */
import https from 'https';
import { spawn } from 'child_process';
import fs from 'fs';
import path from 'path';

const PORT = parseInt(process.env.MCP_WORD_BRIDGE_PORT || '3000', 10);
const LOCK_FILE = process.env.MCP_WORD_BRIDGE_LOCK
  || path.join(process.env.HOME || '/tmp', '.mcp-word-bridge.lock');
const DAEMON_PATH = path.join(__dirname, 'daemon.js');
const ENDPOINT = `https://127.0.0.1:${PORT}/mcp`;

// ─── Lock & Daemon Management ───────────────────────────────────────────────────

// An empty lock held for longer than this (starter died before the daemon wrote
// its PID) is treated as stale. Kept above waitForPort()'s timeout so we never
// reclaim a lock from a starter that is still legitimately waiting.
const STALE_LOCK_MS = 30000;

function pidAlive(pid: number): boolean {
  try {
    // Signal 0 performs error checking without sending a signal.
    process.kill(pid, 0);
    return true;
  } catch (e) {
    // ESRCH → no such process. EPERM → process exists but is owned by another
    // user (still alive). Anything else, assume alive to stay conservative.
    return (e as NodeJS.ErrnoException).code !== 'ESRCH';
  }
}

// Remove the lock if it was orphaned by a dead starter/daemon. Returns true when
// a stale lock was cleared so the caller can retry acquiring it.
function reclaimStaleLock(): boolean {
  let contents: string;
  let mtimeMs: number;
  try {
    contents = fs.readFileSync(LOCK_FILE, 'utf8').trim();
    mtimeMs = fs.statSync(LOCK_FILE).mtimeMs;
  } catch {
    // Lock vanished (e.g. daemon just cleaned up) — treat as reclaimable.
    return true;
  }

  const pid = Number(contents);
  const stale = contents === ''
    ? (Date.now() - mtimeMs) > STALE_LOCK_MS   // empty: starter never spawned a daemon
    : !Number.isInteger(pid) || !pidAlive(pid); // PID present: daemon died uncleanly

  if (!stale) return false;

  try {
    fs.unlinkSync(LOCK_FILE);
    process.stderr.write(`[mcp-word-bridge] Removed stale lock (${contents || 'empty'}).\n`);
  } catch {
    // Someone else may have removed or replaced it concurrently; let the retry decide.
  }
  return true;
}

function tryLock(): boolean {
  try {
    const fd = fs.openSync(LOCK_FILE, fs.constants.O_CREAT | fs.constants.O_EXCL | fs.constants.O_WRONLY);
    fs.closeSync(fd);
    return true;
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === 'EEXIST' && reclaimStaleLock()) {
      // Stale lock cleared — attempt to acquire it exactly once more.
      try {
        const fd = fs.openSync(LOCK_FILE, fs.constants.O_CREAT | fs.constants.O_EXCL | fs.constants.O_WRONLY);
        fs.closeSync(fd);
        return true;
      } catch {
        // Lost the race to another starter that grabbed it first — fall through.
      }
    }
    return false;
  }
}

function portReady(): Promise<boolean> {
  return new Promise(resolve => {
    const req = https.get(ENDPOINT, { rejectUnauthorized: false }, res => {
      res.resume();
      resolve(true);
    });
    req.on('error', () => resolve(false));
    req.setTimeout(1500, () => { req.destroy(); resolve(false); });
  });
}

function waitForPort(timeoutMs = 15000): Promise<void> {
  const start = Date.now();
  return new Promise((resolve, reject) => {
    const poll = () => {
      portReady().then(ok => {
        if (ok) return resolve();
        if (Date.now() - start > timeoutMs) return reject(new Error('Daemon did not start'));
        setTimeout(poll, 200);
      });
    };
    poll();
  });
}

function spawnDaemon(): void {
  const child = spawn('node', [DAEMON_PATH], {
    env: { ...process.env, MCP_WORD_BRIDGE_PORT: String(PORT), MCP_WORD_BRIDGE_LOCK: LOCK_FILE },
    detached: true,
    stdio: 'ignore',
  });
  child.unref();
}

// ─── Stdio-to-HTTP Proxy ────────────────────────────────────────────────────────

function runProxy(): void {
  let sessionId: string | undefined;
  let sseReconnects = 0;
  const MAX_RETRIES = 5;
  const retryCount = new Map<string, number>();

  // Stdin → HTTP POST
  let buffer = '';
  process.stdin.setEncoding('utf8');
  process.stdin.on('data', (chunk: string) => {
    buffer += chunk;
    let idx: number;
    while ((idx = buffer.indexOf('\n')) !== -1) {
      const line = buffer.slice(0, idx).trim();
      buffer = buffer.slice(idx + 1);
      if (line) postMessage(line);
    }
  });

  process.stdin.on('end', () => {
    sendDelete().then(() => process.exit(0));
  });

  function postMessage(json: string): void {
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'Accept': 'application/json, text/event-stream',
    };
    if (sessionId) headers['mcp-session-id'] = sessionId;
    const body = Buffer.from(json, 'utf8');

    const req = https.request({
      hostname: '127.0.0.1', port: PORT, path: '/mcp', method: 'POST',
      headers: { ...headers, 'Content-Length': body.length },
      rejectUnauthorized: false,
    }, (res) => {
      // Capture session ID
      const sid = res.headers['mcp-session-id'];
      if (sid && typeof sid === 'string' && !sessionId) {
        sessionId = sid;
        openSSE();
      }
      // Read response
      const ct = res.headers['content-type'] || '';
      if (ct.includes('text/event-stream')) {
        readSSE(res);
      } else {
        let data = '';
        res.on('data', (d: Buffer) => { data += d.toString(); });
        res.on('end', () => { if (data.trim()) process.stdout.write(data.trim() + '\n'); });
      }
    });

    req.on('error', (e) => {
      const code = (e as any).code;
      if (code === 'ECONNREFUSED' || code === 'ECONNRESET') {
        const n = (retryCount.get(json) || 0) + 1;
        if (n <= MAX_RETRIES) { retryCount.set(json, n); setTimeout(() => postMessage(json), 300); return; }
      }
      retryCount.delete(json);
      process.stderr.write(`[proxy] POST failed: ${e.message}\n`);
    });

    req.write(body);
    req.end();
  }

  function readSSE(stream: NodeJS.ReadableStream): void {
    let buf = '';
    stream.on('data', (chunk: Buffer) => {
      buf += chunk.toString();
      let idx: number;
      while ((idx = buf.indexOf('\n\n')) !== -1) {
        const event = buf.slice(0, idx);
        buf = buf.slice(idx + 2);
        const lines = event.split('\n').filter((l: string) => l.startsWith('data: ')).map((l: string) => l.slice(6));
        if (lines.length) process.stdout.write(lines.join('\n') + '\n');
      }
    });
  }

  function openSSE(): void {
    if (!sessionId) return;
    const req = https.request({
      hostname: '127.0.0.1', port: PORT, path: '/mcp', method: 'GET',
      headers: { 'Accept': 'text/event-stream', 'mcp-session-id': sessionId },
      rejectUnauthorized: false,
    }, (res) => {
      sseReconnects = 0;
      readSSE(res);
      res.on('end', () => { setTimeout(() => { if (sessionId) openSSE(); }, 1000); });
    });
    req.on('error', () => {
      sseReconnects++;
      if (sseReconnects < 5) setTimeout(() => openSSE(), 1000);
    });
    req.end();
  }

  function sendDelete(): Promise<void> {
    if (!sessionId) return Promise.resolve();
    return new Promise(resolve => {
      const req = https.request({
        hostname: '127.0.0.1', port: PORT, path: '/mcp', method: 'DELETE',
        headers: { 'mcp-session-id': sessionId! },
        rejectUnauthorized: false,
      }, () => resolve());
      req.on('error', () => resolve());
      req.setTimeout(2000, () => { req.destroy(); resolve(); });
      req.end();
    });
  }
}

// ─── Main ───────────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  // Fast path: daemon already running
  if (await portReady()) {
    process.stderr.write('[mcp-word-bridge] Connected to running daemon.\n');
    runProxy();
    return;
  }

  // Try to become the starter
  if (tryLock()) {
    process.stderr.write('[mcp-word-bridge] Starting daemon...\n');
    spawnDaemon();
    // Lock file will be overwritten by daemon with its PID once it's ready
  } else {
    process.stderr.write('[mcp-word-bridge] Another process is starting daemon, waiting...\n');
  }

  // Either way, wait for port
  await waitForPort();
  process.stderr.write('[mcp-word-bridge] Daemon ready.\n');
  runProxy();
}

main().catch(e => {
  process.stderr.write(`[mcp-word-bridge] Fatal: ${(e as Error).message}\n`);
  process.exit(1);
});
