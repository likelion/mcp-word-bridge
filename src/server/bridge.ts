import { WebSocket } from 'ws';
import { HEAVY_OPERATIONS, TIMEOUT_DEFAULT, TIMEOUT_HEAVY } from '../shared/constants';
import type { BridgeRequest, BridgeResponseMessage } from '../shared/protocol';

interface PendingRequest {
  resolve: (result: unknown) => void;
  reject: (error: Error) => void;
}

/**
 * Bridge to the Word taskpane via WebSocket.
 * Sends commands and awaits responses with timeout.
 */
export class Bridge {
  private socket: WebSocket | null = null;
  private pending = new Map<string, PendingRequest>();
  private requestId = 0;

  get connected(): boolean {
    return this.socket !== null && this.socket.readyState === WebSocket.OPEN;
  }

  attach(ws: WebSocket): void {
    this.socket = ws;
    ws.on('message', (data: Buffer) => {
      try {
        const msg = JSON.parse(data.toString()) as BridgeResponseMessage;
        if (msg.type === 'response' && msg.id) {
          const pending = this.pending.get(msg.id);
          if (pending) {
            this.pending.delete(msg.id);
            if (msg.error) {
              pending.reject(new Error(msg.error));
            } else {
              pending.resolve(msg.result);
            }
          }
        }
      } catch {
        // Ignore malformed messages
      }
    });
    ws.on('close', () => {
      this.socket = null;
      this.rejectAll('Word taskpane disconnected. Reopen the MCP Word Bridge add-in in Word.');
    });
  }

  detach(): void {
    this.socket = null;
    this.rejectAll('Word taskpane disconnected.');
  }

  /**
   * Send a command to the taskpane and await its response.
   */
  send<T = unknown>(action: string, params: Record<string, unknown> = {}): Promise<T> {
    return new Promise((resolve, reject) => {
      if (!this.connected) {
        reject(new Error('Word taskpane not connected. Open the MCP Word Bridge add-in in Word.'));
        return;
      }

      const id = String(++this.requestId);
      const timeoutMs = HEAVY_OPERATIONS.has(action) ? TIMEOUT_HEAVY : TIMEOUT_DEFAULT;

      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`Operation timed out after ${timeoutMs / 1000}s. The document may be too large or Word is busy.`));
      }, timeoutMs);

      this.pending.set(id, {
        resolve: (result) => {
          clearTimeout(timer);
          resolve(result as T);
        },
        reject: (err) => {
          clearTimeout(timer);
          reject(err);
        },
      });

      const msg: BridgeRequest = { id, action, params };
      this.socket!.send(JSON.stringify(msg));
    });
  }

  private rejectAll(message: string): void {
    for (const [, pending] of this.pending) {
      pending.reject(new Error(message));
    }
    this.pending.clear();
  }
}
