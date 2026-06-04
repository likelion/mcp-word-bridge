import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { CallToolRequestSchema, ListToolsRequestSchema, ListResourcesRequestSchema, ReadResourceRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import type { Bridge } from './bridge';
import { buildToolRegistry } from './tools';
import { ToolError } from './types';
import { createMutex } from './mutex';
import { usageGuide } from './usage-guide';

// Read version from package.json at build time (inlined by esbuild)
import pkg from '../../package.json';

export function createMcpServer(bridge: Bridge): Server {
  const server = new Server(
    { name: 'mcp-word-bridge', version: pkg.version },
    { capabilities: { tools: {}, resources: {} } },
  );

  const { tools, handlers } = buildToolRegistry();

  // Serialize tool calls so concurrent MCP requests execute one at a time.
  // This prevents race conditions when multiple tools modify the document.
  const toolMutex = createMutex();

  // List tools
  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: tools.map(t => ({
      name: t.name,
      description: t.description,
      inputSchema: {
        type: 'object' as const,
        properties: t.schema.properties,
        ...(t.schema.required ? { required: t.schema.required } : {}),
      },
    })),
  }));

  // Call tool — serialized via mutex to guarantee ordering
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    return toolMutex.run(async () => {
      const { name, arguments: args } = request.params;
      const handler = handlers.get(name);
      if (!handler) {
        return { content: [{ type: 'text' as const, text: 'Unknown tool: ' + name }], isError: true };
      }
      try {
        const result = await handler(args || {}, bridge);
        return result as any;
      } catch (e) {
        const msg = e instanceof ToolError ? e.message : 'Error: ' + (e as Error).message;
        return { content: [{ type: 'text' as const, text: msg }], isError: true };
      }
    });
  });

  // Resources — usage guide
  server.setRequestHandler(ListResourcesRequestSchema, async () => ({
    resources: [{
      uri: 'word-bridge://usage-guide',
      name: 'Word Bridge Usage Guide',
      description: 'Patterns, workflows, and pitfalls for LLMs operating on Word documents',
      mimeType: 'text/markdown',
    }],
  }));

  server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
    if (request.params.uri === 'word-bridge://usage-guide') {
      return { contents: [{ uri: request.params.uri, mimeType: 'text/markdown', text: usageGuide }] };
    }
    throw new Error('Resource not found: ' + request.params.uri);
  });

  return server;
}
