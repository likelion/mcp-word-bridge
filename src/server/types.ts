import type { Bridge } from './bridge';

/**
 * JSON Schema property definition for tool input schemas.
 */
export interface SchemaProperty {
  type: 'string' | 'number' | 'boolean' | 'array' | 'object';
  description?: string;
  enum?: string[];
  items?: SchemaProperty;
}

/**
 * A tool definition: name, description, schema, and handler.
 * This is the single source of truth — no separate action map.
 */
export interface ToolDefinition {
  name: string;
  description: string;
  schema: {
    properties: Record<string, SchemaProperty>;
    required?: string[];
  };
  handler: ToolHandler;
}

/**
 * Tool handler function signature.
 */
export type ToolHandler = (
  args: Record<string, unknown>,
  bridge: Bridge,
) => Promise<ToolResult>;

/**
 * Result returned by a tool handler.
 */
export interface ToolResult {
  content: Array<{ type: 'text'; text: string }>;
  isError?: boolean;
}

/**
 * Error thrown by tools — caught and returned as isError response.
 */
export class ToolError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ToolError';
  }
}
