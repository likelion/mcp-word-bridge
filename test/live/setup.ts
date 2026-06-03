/**
 * Shared setup for live tests.
 * Manages a single client connection across all test files.
 */
import { LiveTestClient } from './client';

let client: LiveTestClient;
let available = false;
let initialized = false;

export async function getClient(): Promise<LiveTestClient> {
  if (!initialized) {
    initialized = true;
    client = new LiveTestClient();
    try {
      await client.connect();
      available = await client.waitForWord(30);
    } catch (e: any) {
      console.error('Could not connect:', e.message);
    }
    if (!available) {
      console.warn(
        '\n⚠️  Word not connected. Skipping live tests.\n' +
        '   Ensure Word is open with MCP Word Bridge add-in loaded.\n' +
        '   Stop any other MCP client using the bridge first.\n',
      );
    }
  }
  return client;
}

export function skip(): boolean {
  return !available;
}

export async function cleanup(): Promise<void> {
  if (client) await client.disconnect();
}
