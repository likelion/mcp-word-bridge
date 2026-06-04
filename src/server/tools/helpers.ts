import type { ToolDefinition, ToolResult } from '../types';

/**
 * Wrap a handler result as MCP text content.
 */
export function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/** Optional server-side validation function run before forwarding to taskpane. */
export type ForwardValidator = (args: Record<string, unknown>) => void;

/**
 * Convenience: define a tool that simply forwards to a bridge action
 * and returns the result as JSON.
 * The `bridgeAction` property enables batch optimization.
 * An optional `validate` function runs server-side before forwarding,
 * providing a single authoritative validation layer.
 */
export function forwardTool(
  name: string,
  description: string,
  schema: ToolDefinition['schema'],
  action: string,
  validate?: ForwardValidator,
): ToolDefinition & { bridgeAction: string } {
  return {
    name,
    description,
    schema,
    bridgeAction: action,
    async handler(args, bridge) {
      if (validate) validate(args);
      const result = await bridge.send(action, args);
      return jsonResult(result);
    },
  };
}
