/**
 * Live test client — spawns an MCP server and calls tools via the MCP SDK.
 * Requires Word with the add-in to connect to the server's WebSocket.
 */
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import path from 'path';

const PORT = process.env.MCP_WORD_BRIDGE_PORT || '3000';

export class LiveTestClient {
  private client: Client | null = null;
  private transport: StdioClientTransport | null = null;

  async connect(): Promise<void> {
    this.transport = new StdioClientTransport({
      command: 'node',
      args: [path.join(__dirname, '../../dist/server.js')],
      env: { ...process.env, MCP_WORD_BRIDGE_PORT: PORT },
    });
    this.client = new Client({ name: 'live-test', version: '1.0.0' }, {});
    await this.client.connect(this.transport);
  }

  async waitForWord(timeoutSec = 60): Promise<boolean> {
    const deadline = Date.now() + timeoutSec * 1000;
    while (Date.now() < deadline) {
      try {
        const result = await this.call('word_get_text');
        if (result) return true;
      } catch (e: any) {
        if (e.message?.includes('not connected')) {
          await new Promise(r => setTimeout(r, 2000));
          continue;
        }
        throw e;
      }
    }
    return false;
  }

  async call<T = any>(toolName: string, args: Record<string, any> = {}): Promise<T> {
    const result = await this.client!.callTool({ name: toolName, arguments: args });
    if (result.isError) {
      const msg = (result.content as any)[0]?.text || 'Unknown error';
      throw new Error(msg);
    }
    return JSON.parse((result.content as any)[0]?.text || '{}');
  }

  async expectError(toolName: string, args: Record<string, any> = {}): Promise<string> {
    const result = await this.client!.callTool({ name: toolName, arguments: args });
    if (!result.isError) {
      throw new Error('Expected error but got: ' + JSON.stringify(result.content));
    }
    return (result.content as any)[0]?.text || '';
  }

  async disconnect(): Promise<void> {
    if (this.client) {
      try { await this.client.close(); } catch { /* ignore */ }
    }
  }

  async resetDocument(): Promise<void> {
    await this.call('word_clear');
    try { await this.call('word_set_change_tracking', { mode: 'Off' }); } catch { /* ignore */ }
    try { await this.call('word_accept_all_tracked_changes'); } catch { /* ignore */ }
    try {
      await this.call('word_set_page_layout', {
        orientation: 'Portrait', topMargin: 72, bottomMargin: 72, leftMargin: 72, rightMargin: 72,
      });
    } catch { /* ignore */ }
    try { await this.call('word_set_header_footer', { type: 'header', text: '' }); } catch { /* ignore */ }
    try { await this.call('word_set_header_footer', { type: 'footer', text: '' }); } catch { /* ignore */ }
    try {
      const props = await this.call('word_get_custom_properties');
      for (const p of props.properties || []) {
        await this.call('word_delete_custom_property', { key: p.key });
      }
    } catch { /* ignore */ }
  }
}
