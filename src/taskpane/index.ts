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
    log('→ ERR: ' + e.message, 'log-err');
    sendError(cmd.id, e.message);
  }
}
