/**
 * Multi-client integration tests for mcp-word-bridge daemon architecture.
 * Tests all combinations: simultaneous start, sequential start, disconnect order.
 *
 * Run: node test/integration/multi-client.test.js
 */
const { spawn } = require('child_process');
const https = require('https');
const fs = require('fs');
const path = require('path');

const LOCK_FILE = path.join(process.env.HOME, '.mcp-word-bridge.lock');
const PORT = 3000;
const ENV = { ...process.env, MCP_WORD_BRIDGE_GRACE: '2000' };
const INIT_MSG = JSON.stringify({
  jsonrpc: '2.0', method: 'initialize',
  params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'test', version: '1' } },
  id: 1,
});

let testsPassed = 0;
let testsFailed = 0;

function cleanup() {
  try { fs.unlinkSync(LOCK_FILE); } catch {}
  return killDaemon();
}

function killDaemon() {
  return new Promise(resolve => {
    if (fs.existsSync(LOCK_FILE)) {
      try {
        const pid = parseInt(fs.readFileSync(LOCK_FILE, 'utf8'));
        process.kill(pid, 'SIGTERM');
      } catch {}
    }
    setTimeout(resolve, 1000);
  });
}

function daemonAlive() {
  return new Promise(resolve => {
    const req = https.get(`https://127.0.0.1:${PORT}/mcp`, { rejectUnauthorized: false }, (res) => {
      res.resume();
      resolve(true);
    });
    req.on('error', () => resolve(false));
    req.setTimeout(2000, () => { req.destroy(); resolve(false); });
  });
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function startClient() {
  const c = spawn('node', ['dist/server.js'], { env: ENV, stdio: ['pipe', 'pipe', 'pipe'] });
  let stderr = '';
  let stdout = '';
  c.stderr.on('data', d => { stderr += d.toString(); });
  c.stdout.on('data', d => { stdout += d.toString(); });
  return {
    proc: c, get stderr() { return stderr; }, get stdout() { return stdout; },
    send(msg) { c.stdin.write(msg + '\n'); },
    disconnect() { c.stdin.end(); },
    kill() { c.kill(); },
    waitForOutput(match, timeoutMs = 10000) {
      return new Promise((resolve, reject) => {
        const start = Date.now();
        const check = () => {
          if (stdout.includes(match)) return resolve();
          if (Date.now() - start > timeoutMs) return reject(new Error(`Timeout waiting for "${match}". Got stdout: ${stdout.slice(0, 200)}`));
          setTimeout(check, 100);
        };
        check();
      });
    },
  };
}

async function assert(condition, msg) {
  if (condition) { testsPassed++; console.log(`  ✓ ${msg}`); }
  else { testsFailed++; console.log(`  ✗ ${msg}`); }
}

// ─── Test 1: Simultaneous start, both initialize ────────────────────────────────

async function testSimultaneousStart() {
  console.log('\nTest 1: Two clients start simultaneously');
  await cleanup();

  const c1 = startClient();
  const c2 = startClient();

  await sleep(10000); // wait for both to connect to daemon

  c1.send(INIT_MSG);
  c2.send(INIT_MSG);

  await c1.waitForOutput('"protocolVersion"');
  await c2.waitForOutput('"protocolVersion"');

  await assert(c1.stdout.includes('mcp-word-bridge'), 'C1 got init response');
  await assert(c2.stdout.includes('mcp-word-bridge'), 'C2 got init response');
  await assert(await daemonAlive(), 'Daemon is alive');

  c1.kill(); c2.kill();
  await cleanup();
}

// ─── Test 2: Sequential start (C1 first, then C2 joins) ─────────────────────────

async function testSequentialStart() {
  console.log('\nTest 2: C1 starts first, C2 joins later');
  await cleanup();

  const c1 = startClient();
  await sleep(8000); // C1 fully starts daemon

  c1.send(INIT_MSG);
  await c1.waitForOutput('"protocolVersion"');
  await assert(c1.stdout.includes('mcp-word-bridge'), 'C1 initialized');

  const c2 = startClient();
  await sleep(5000); // C2 detects existing daemon

  c2.send(INIT_MSG);
  await c2.waitForOutput('"protocolVersion"');
  await assert(c2.stdout.includes('mcp-word-bridge'), 'C2 initialized (joined existing daemon)');
  await assert(c2.stderr.includes('Connected to running daemon'), 'C2 detected existing daemon');

  c1.kill(); c2.kill();
  await cleanup();
}

// ─── Test 3: First client leaves, daemon stays ───────────────────────────────────

async function testFirstLeavesDaemonStays() {
  console.log('\nTest 3: C1 leaves, daemon stays alive for C2');
  await cleanup();

  const c1 = startClient();
  const c2 = startClient();
  await sleep(10000);

  c1.send(INIT_MSG);
  c2.send(INIT_MSG);
  await c1.waitForOutput('"protocolVersion"');
  await c2.waitForOutput('"protocolVersion"');

  c1.disconnect();
  await sleep(2000);

  await assert(await daemonAlive(), 'Daemon still alive after C1 left');

  c2.kill();
  await cleanup();
}

// ─── Test 4: Last client leaves, daemon shuts down ───────────────────────────────

async function testLastLeavesShutdown() {
  console.log('\nTest 4: Both leave, daemon auto-shuts down');
  await cleanup();

  const c1 = startClient();
  const c2 = startClient();
  await sleep(10000);

  c1.send(INIT_MSG);
  c2.send(INIT_MSG);
  await c1.waitForOutput('"protocolVersion"');
  await c2.waitForOutput('"protocolVersion"');

  c1.disconnect();
  await sleep(1000);
  c2.disconnect();

  // Grace period is 2s, wait 4s total
  await sleep(4000);

  await assert(!(await daemonAlive()), 'Daemon shut down after both left');
  await assert(!fs.existsSync(LOCK_FILE), 'Lock file cleaned up');
}

// ─── Test 5: Single client lifecycle ─────────────────────────────────────────────

async function testSingleClient() {
  console.log('\nTest 5: Single client start → init → disconnect → daemon shuts down');
  await cleanup();

  const c1 = startClient();
  await sleep(8000);

  c1.send(INIT_MSG);
  await c1.waitForOutput('"protocolVersion"');
  await assert(c1.stdout.includes('mcp-word-bridge'), 'C1 initialized');

  c1.disconnect();
  await sleep(4000);

  await assert(!(await daemonAlive()), 'Daemon shut down after single client left');
}

// ─── Test 6: Client reconnects within grace period ───────────────────────────────

async function testReconnectDuringGrace() {
  console.log('\nTest 6: Client leaves, new client connects during grace period');
  await cleanup();

  const c1 = startClient();
  await sleep(8000);

  c1.send(INIT_MSG);
  await c1.waitForOutput('"protocolVersion"');

  c1.disconnect();
  await sleep(1000); // within 2s grace

  // New client arrives before grace expires
  const c2 = startClient();
  await sleep(5000);

  c2.send(INIT_MSG);
  await c2.waitForOutput('"protocolVersion"');
  await assert(c2.stdout.includes('mcp-word-bridge'), 'C2 connected during grace period');
  await assert(await daemonAlive(), 'Daemon stayed alive');

  c2.kill();
  await cleanup();
}

// ─── Run all tests ───────────────────────────────────────────────────────────────

async function run() {
  console.log('Multi-client integration tests\n==============================');
  await testSimultaneousStart();
  await testSequentialStart();
  await testFirstLeavesDaemonStays();
  await testLastLeavesShutdown();
  await testSingleClient();
  await testReconnectDuringGrace();

  console.log(`\n${testsPassed + testsFailed} assertions: ${testsPassed} passed, ${testsFailed} failed`);
  process.exit(testsFailed > 0 ? 1 : 0);
}

run().catch(e => {
  console.error('Test runner error:', e.message);
  cleanup().then(() => process.exit(1));
});
