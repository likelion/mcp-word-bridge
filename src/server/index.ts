import https from 'https';
import fs from 'fs';
import path from 'path';
import { WebSocketServer, WebSocket } from 'ws';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { Bridge } from './bridge';
import { createMcpServer } from './mcp';
import { MAX_PAYLOAD } from '../shared/constants';

const PORT = parseInt(process.env.MCP_WORD_BRIDGE_PORT || '3000', 10);
const CERTS_DIR = path.join(__dirname, '..', 'certs');

// ─── TLS Certificate Management ────────────────────────────────────────────────

function ensureCerts(): void {
  const certPath = path.join(CERTS_DIR, 'cert.pem');
  const keyPath = path.join(CERTS_DIR, 'key.pem');
  if (fs.existsSync(certPath) && fs.existsSync(keyPath)) return;

  process.stderr.write('[mcp-word-bridge] Generating self-signed TLS certificate...\n');
  if (!fs.existsSync(CERTS_DIR)) fs.mkdirSync(CERTS_DIR, { recursive: true });

  const { execSync } = require('child_process');
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
  process.stderr.write('[mcp-word-bridge] ✓ Certificate generated.\n');

  if (process.platform === 'darwin') {
    process.stderr.write(`[mcp-word-bridge] Trust it: security add-trusted-cert -r trustRoot -k ~/Library/Keychains/login.keychain-db "${certPath}"\n`);
  } else if (process.platform === 'win32') {
    process.stderr.write(`[mcp-word-bridge] Trust it: certutil -user -addstore Root "${certPath}"\n`);
  }
}

// ─── HTTPS Static Server ────────────────────────────────────────────────────────

const MIME: Record<string, string> = {
  '.html': 'text/html', '.js': 'application/javascript', '.css': 'text/css',
  '.png': 'image/png', '.json': 'application/json', '.xml': 'application/xml',
};

function createHttpsServer() {
  ensureCerts();
  const options = {
    key: fs.readFileSync(path.join(CERTS_DIR, 'key.pem')),
    cert: fs.readFileSync(path.join(CERTS_DIR, 'cert.pem')),
  };

  const rootDir = path.join(__dirname, '..');

  return https.createServer(options, (req, res) => {
    let urlPath = (req.url || '/').split('?')[0]!;
    let filePath = urlPath === '/' ? '/taskpane.html' : urlPath;
    filePath = path.resolve(rootDir, '.' + filePath);

    // Prevent path traversal
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
  });
}

// ─── Main ───────────────────────────────────────────────────────────────────────

let shuttingDown = false;
function shutdown(): void {
  if (shuttingDown) return;
  shuttingDown = true;
  process.stderr.write('[mcp-word-bridge] Shutting down...\n');
  wss.clients.forEach(ws => ws.terminate());
  httpsServer.close();
  httpsServer.closeAllConnections();
  setTimeout(() => process.exit(0), 500);
}

const httpsServer = createHttpsServer();
const wss = new WebSocketServer({ server: httpsServer, maxPayload: MAX_PAYLOAD });
wss.on('error', () => { /* handled via httpsServer error event */ });
const bridge = new Bridge();

wss.on('connection', (ws: WebSocket, req) => {
  if (req.url === '/taskpane') {
    bridge.attach(ws);
    process.stderr.write('[bridge] Taskpane connected\n');
    ws.on('close', () => {
      process.stderr.write('[bridge] Taskpane disconnected\n');
    });
  }
});

process.on('SIGTERM', shutdown);
process.on('SIGINT', shutdown);
process.stdin.on('end', shutdown);
process.stdin.on('close', shutdown);

async function main(): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    httpsServer.on('error', (err: NodeJS.ErrnoException) => {
      if (err.code === 'EADDRINUSE') {
        reject(new Error(`Port ${PORT} is already in use. Stop the other process or set MCP_WORD_BRIDGE_PORT.`));
      } else { reject(err); }
    });
    httpsServer.listen(PORT, () => {
      process.stderr.write(`[mcp-word-bridge] Bridge server on https://localhost:${PORT}\n`);
      resolve();
    });
  });

  const mcpServer = createMcpServer(bridge);
  const transport = new StdioServerTransport();
  await mcpServer.connect(transport);
  process.stderr.write(`[mcp-word-bridge] MCP server ready\n`);
}

main().catch(e => {
  process.stderr.write('[mcp-word-bridge] Fatal: ' + (e as Error).message + '\n');
  process.exit(1);
});
