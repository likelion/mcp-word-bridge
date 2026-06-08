/**
 * MCP Word Bridge — Taskpane Application
 * Runs inside the Word add-in webview. Receives commands via WebSocket,
 * executes them against the Word JavaScript API, and returns results.
 */

import { commandRegistry } from './commands';
import { log } from './log';
import type { BridgeRequest, BridgeResponseMessage } from '../shared/protocol';

declare const Office: any;
declare const Word: any;

// ─── UI Elements ────────────────────────────────────────────────────────────────

const wordStatus = document.getElementById('word-status')!;
const wsStatus = document.getElementById('ws-status')!;

// ─── Word Initialization ────────────────────────────────────────────────────────

// eslint-disable-next-line @typescript-eslint/no-unused-vars
let _wordReady = false;

Office.onReady((info: any) => {
  if (info.host === Office.HostType.Word) {
    _wordReady = true;
    wordStatus.textContent = 'Word: connected ✓';
    wordStatus.className = 'status ok';
    log('Office.js ready', 'log-ok');
    connectWebSocket();
  } else {
    wordStatus.textContent = 'Word: wrong host';
    wordStatus.className = 'status err';
  }
});

// ─── WebSocket Connection ───────────────────────────────────────────────────────

let ws: WebSocket | null = null;

function connectWebSocket(): void {
  ws = new WebSocket('wss://' + window.location.host + '/taskpane');
  ws.onopen = () => {
    wsStatus.textContent = 'WebSocket: connected ✓';
    wsStatus.className = 'status ok';
    log('WebSocket connected', 'log-ok');
  };
  ws.onclose = () => {
    wsStatus.textContent = 'WebSocket: disconnected';
    wsStatus.className = 'status warn';
    setTimeout(connectWebSocket, 3000);
  };
  ws.onerror = () => log('WebSocket error', 'log-err');
  ws.onmessage = (event: MessageEvent) => handleCommand(JSON.parse(event.data));
}

function sendResponse(id: string, result: unknown): void {
  if (ws?.readyState === WebSocket.OPEN) {
    const msg: BridgeResponseMessage = { type: 'response', id, result };
    ws.send(JSON.stringify(msg));
  }
}

function sendError(id: string, error: string): void {
  if (ws?.readyState === WebSocket.OPEN) {
    const msg: BridgeResponseMessage = { type: 'response', id, error };
    ws.send(JSON.stringify(msg));
  }
}

// ─── Error Formatting ───────────────────────────────────────────────────────────

/**
 * Extract useful details from Office.js OfficeExtension.Error objects.
 * Falls back to e.message for non-Office errors.
 */
function formatOfficeError(e: any): string {
  // If it's already a well-formed error message (from our own throw), return as-is
  if (!e.debugInfo && !e.code) return e.message || String(e);

  const parts: string[] = [];
  const code = e.code || '';
  const msg = e.message || '';

  // Start with the code + message
  if (code && msg && !msg.includes(code)) {
    parts.push(`${code}: ${msg}`);
  } else {
    parts.push(msg || code);
  }

  // Append debugInfo details if available
  if (e.debugInfo) {
    if (e.debugInfo.message && e.debugInfo.message !== msg) {
      parts.push(e.debugInfo.message);
    }
    if (e.debugInfo.errorLocation) {
      parts.push(`Location: ${e.debugInfo.errorLocation}`);
    }
  }

  return parts.join(' | ');
}

// ─── Command Dispatcher ─────────────────────────────────────────────────────────

async function handleCommand(cmd: BridgeRequest): Promise<void> {
  log('← ' + cmd.action, 'log-cmd');
  try {
    const handler = commandRegistry[cmd.action];
    if (!handler) throw new Error('Unknown action: ' + cmd.action);

    const result = await Word.run(async (ctx: any) => {
      return handler(ctx, cmd.params || {});
    });

    log('→ ok', 'log-ok');
    sendResponse(cmd.id, result);
  } catch (e: any) {
    const errMsg = formatOfficeError(e);
    log('→ ERR: ' + errMsg, 'log-err');
    sendError(cmd.id, errMsg);
  }
}
