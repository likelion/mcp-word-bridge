/**
 * mcp-word-bridge daemon
 *
 * HTTP-only MCP server that serves multiple clients via Streamable HTTP.
 * Manages the Word taskpane WebSocket bridge, TLS certs, and static assets.
 * Auto-shuts down after a configurable grace period with no active sessions.
 *
 * Never started directly by users — the entrypoint (server.ts) spawns this.
 */
import https from 'https';
import http from 'http';
import crypto from 'crypto';
import { execSync } from 'child_process';
import fs from 'fs';
import path from 'path';
import { WebSocketServer, WebSocket } from 'ws';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { Bridge } from './bridge';
import { createMcpServer } from './mcp';
import { MAX_PAYLOAD } from '../shared/constants';

const PORT = parseInt(process.env.MCP_WORD_BRIDGE_PORT || '3000', 10);
const GRACE_MS = parseInt(process.env.MCP_WORD_BRIDGE_GRACE || '5000', 10);
const LOCK_FILE = process.env.MCP_WORD_BRIDGE_LOCK
  || path.join(process.env.HOME || '/tmp', '.mcp-word-bridge.lock');
const CERTS_DIR = path.join(__dirname, '..', 'certs');

// ─── TLS Certificates ───────────────────────────────────────────────────────────

function ensureCerts(): void {
  const certPath = path.join(CERTS_DIR, 'cert.pem');
  const keyPath = path.join(CERTS_DIR, 'key.pem');
  if (fs.existsSync(certPath) && fs.existsSync(keyPath)) return;

  process.stderr.write('[daemon] Generating self-signed TLS certificate...\n');
  if (!fs.existsSync(CERTS_DIR)) fs.mkdirSync(CERTS_DIR, { recursive: true });

  const confPath = path.join(CERTS_DIR, 'cert.conf');
  if (!fs.existsSync(confPath)) {
    fs.writeFileSync(confPath, [
      '[req]', 'default_bits = 2048', 'prompt = no', 'default_md = sha256',
      'distinguished_name = dn', 'x509_extensions = v3_req', '',
      '[dn]', 'CN = localhost', '',
      '[v3_req]', 'basicConstraints = CA:TRUE', 'subjectAltName = @alt_names', '',
      '[alt_names]', 'DNS.1 = localhost', 'IP.1 = 127.0.0.1', '',
    ].join('\n'));
  }
  execSync(
    `openssl req -x509 -newkey rsa:2048 -keyout "${keyPath}" -out "${certPath}" -days 3650 -nodes -config "${confPath}"`,
    { stdio: 'pipe' },
  );
  process.stderr.write('[daemon] ✓ Certificate generated.\n');
  if (process.platform === 'darwin') {
    process.stderr.write(`[daemon] Trust: security add-trusted-cert -r trustRoot -k ~/Library/Keychains/login.keychain-db "${certPath}"\n`);
  }
}

// ─── Session Tracker & Auto-Shutdown ────────────────────────────────────────────

const sessions = new Map<string, { transport: StreamableHTTPServerTransport; server: ReturnType<typeof createMcpServer> }>();
let graceTimer: ReturnType<typeof setTimeout> | null = null;
let hadSession = false; // Only auto-shutdown after at least one session has connected

function resetGraceTimer(): void {
  if (graceTimer) { clearTimeout(graceTimer); graceTimer = null; }
  if (sessions.size === 0 && hadSession) {
    graceTimer = setTimeout(() => {
      if (sessions.size === 0) {
        process.stderr.write(`[daemon] No sessions for ${GRACE_MS / 1000}s — shutting down.\n`);
        shutdown();
      }
    }, GRACE_MS);
  }
}

function removeSession(sid: string): void {
  const s = sessions.get(sid);
  if (!s) return;
  sessions.delete(sid);
  process.stderr.write(`[daemon] Session ${sid.slice(0, 8)}… closed (${sessions.size} active)\n`);
  resetGraceTimer();
}

// ─── HTTPS Server ───────────────────────────────────────────────────────────────

const MIME: Record<string, string> = {
  '.html': 'text/html', '.js': 'application/javascript', '.css': 'text/css',
  '.png': 'image/png', '.json': 'application/json', '.xml': 'application/xml',
};

function createServer(bridge: Bridge) {
  ensureCerts();
  const tlsOpts = {
    key: fs.readFileSync(path.join(CERTS_DIR, 'key.pem')),
    cert: fs.readFileSync(path.join(CERTS_DIR, 'cert.pem')),
  };
  const rootDir = path.join(__dirname, '..');

  return https.createServer(tlsOpts, async (req, res) => {
    try {
      const urlPath = (req.url || '/').split('?')[0]!;

      if (urlPath === '/mcp') {
        await handleMcp(req, res, bridge);
        return;
      }

      // Static file serving (taskpane assets)
      serveStatic(rootDir, urlPath, res);
    } catch (e) {
      process.stderr.write(`[daemon] Request error: ${(e as Error).message}\n`);
      if (!res.headersSent) { res.writeHead(500); res.end('Internal error'); }
    }
  });
}

async function handleMcp(req: http.IncomingMessage, res: http.ServerResponse, bridge: Bridge): Promise<void> {
  const sessionId = req.headers['mcp-session-id'] as string | undefined;

  // DELETE — explicit session teardown
  if (req.method === 'DELETE' && sessionId) {
    const s = sessions.get(sessionId);
    if (s) { await s.transport.close(); }
    res.writeHead(200); res.end();
    return;
  }

  // Existing session — forward
  if (sessionId && sessions.has(sessionId)) {
    const s = sessions.get(sessionId)!;
    await s.transport.handleRequest(req, res);
    // Detect SSE disconnect as backup cleanup
    if (req.method === 'GET') {
      res.on('close', () => { const entry = sessions.get(sessionId); if (entry) entry.transport.close(); });
    }
    return;
  }

  // New session (POST without session ID)
  if (req.method === 'POST' && !sessionId) {
    if (graceTimer) { clearTimeout(graceTimer); graceTimer = null; }

    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => crypto.randomUUID(),
    });
    transport.onclose = () => { const sid = transport.sessionId; if (sid) removeSession(sid); };
    transport.onerror = (err) => { process.stderr.write(`[daemon] Transport error: ${err.message}\n`); };

    const mcpServer = createMcpServer(bridge);
    await mcpServer.connect(transport);
    await transport.handleRequest(req, res);

    const sid = transport.sessionId;
    if (sid) {
      hadSession = true;
      sessions.set(sid, { transport, server: mcpServer });
      process.stderr.write(`[daemon] New session ${sid.slice(0, 8)}… (${sessions.size} active)\n`);
    }
    return;
  }

  res.writeHead(400, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify({ error: 'Bad request' }));
}

function serveStatic(rootDir: string, urlPath: string, res: http.ServerResponse): void {
  let filePath = urlPath === '/' ? '/taskpane.html' : urlPath;
  filePath = path.resolve(rootDir, '.' + filePath);
  if (!filePath.startsWith(rootDir + path.sep) && filePath !== rootDir) {
    res.writeHead(403); res.end('Forbidden'); return;
  }
  const ext = path.extname(filePath);
  const contentType = MIME[ext] || 'application/octet-stream';
  fs.readFile(filePath, (err, data) => {
    if (err) {
      if (ext === '.png') {
        const pixel = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', 'base64');
        res.writeHead(200, { 'Content-Type': 'image/png' }); res.end(pixel); return;
      }
      res.writeHead(404); res.end('Not found'); return;
    }
    res.writeHead(200, { 'Content-Type': contentType }); res.end(data);
  });
}

// ─── Lifecycle ──────────────────────────────────────────────────────────────────

let shuttingDown = false;
let httpsServer: https.Server;
let wss: WebSocketServer;

function shutdown(): void {
  if (shuttingDown) return;
  shuttingDown = true;
  process.stderr.write('[daemon] Shutting down...\n');
  wss.clients.forEach(ws => ws.terminate());
  httpsServer.close();
  httpsServer.closeAllConnections();
  try { fs.unlinkSync(LOCK_FILE); } catch { /* ignore */ }
  setTimeout(() => process.exit(0), 500);
}

process.on('SIGTERM', shutdown);
process.on('SIGINT', shutdown);
process.on('unhandledRejection', (reason) => {
  process.stderr.write(`[daemon] Unhandled rejection: ${reason}\n`);
});

async function main(): Promise<void> {
  const bridge = new Bridge();
  httpsServer = createServer(bridge);
  wss = new WebSocketServer({ server: httpsServer, maxPayload: MAX_PAYLOAD });

  wss.on('connection', (ws: WebSocket, req) => {
    if (req.url === '/taskpane') {
      bridge.attach(ws);
      process.stderr.write('[daemon] Taskpane connected\n');
      ws.on('close', () => process.stderr.write('[daemon] Taskpane disconnected\n'));
    }
  });

  await new Promise<void>((resolve, reject) => {
    httpsServer.on('error', (err: NodeJS.ErrnoException) => {
      if (err.code === 'EADDRINUSE') reject(new Error(`Port ${PORT} in use`));
      else reject(err);
    });
    httpsServer.listen(PORT, '127.0.0.1', () => resolve());
  });

  // Write PID to lock file so entrypoint can find us
  fs.writeFileSync(LOCK_FILE, String(process.pid));
  process.stderr.write(`[daemon] Listening on https://127.0.0.1:${PORT}/mcp (grace=${GRACE_MS / 1000}s)\n`);
}

main().catch(e => {
  process.stderr.write(`[daemon] Fatal: ${(e as Error).message}\n`);
  process.exit(1);
});
