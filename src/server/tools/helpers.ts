import type { ToolDefinition, ToolResult } from '../types';

/**
 * Wrap a handler result as MCP text content.
 */
export function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Convenience: define a tool that simply forwards to a bridge action
 * and returns the result as JSON.
 * The `bridgeAction` property enables batch optimization.
 */
export function forwardTool(
  name: string,
  description: string,
  schema: ToolDefinition['schema'],
  action: string,
): ToolDefinition & { bridgeAction: string } {
  return {
    name,
    description,
    schema,
    bridgeAction: action,
    async handler(args, bridge) {
      const result = await bridge.send(action, args);
      return jsonResult(result);
    },
  };
}
