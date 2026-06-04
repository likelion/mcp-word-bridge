import type { ToolDefinition, ToolHandler } from '../types';
import { ToolError } from '../types';
import { MAX_BATCH_OPERATIONS } from '../../shared/constants';
import { jsonResult } from './helpers';
import type { ForwardValidator } from './helpers';
import type { BatchResult } from '../../shared/protocol';

interface OpEntry { tool: string; args: Record<string, unknown>; originalIndex: number }
interface ResultEntry { index: number; tool: string; success: boolean; result?: unknown; error?: string }

/**
 * The batch tool buffers consecutive "native" (forwarded) operations and sends them
 * to the taskpane in a single batchExecute message. Server-composed tools flush the
 * buffer and execute individually.
 *
 * Validators are run server-side before buffering native ops, ensuring the same
 * validation applies whether a tool is called directly or via batch.
 */
export function createBatchTool(
  registry: Map<string, ToolHandler>,
  actionMap: Map<string, string>,
  validators: Map<string, ForwardValidator>,
): ToolDefinition {
  return {
    name: 'word_batch',
    description: '[Batch] Execute multiple operations in a single call. Operations execute sequentially — if one fails, subsequent are skipped. Note: paragraph indices are NOT auto-adjusted between operations. Multiple inserts at the same index will produce reversed order (last inserted appears first).',
    schema: {
      properties: {
        operations: {
          type: 'array',
          items: { type: 'object', description: 'Array of {tool, args} objects' },
          description: 'Array of {tool, args} objects to execute sequentially',
        },
      },
      required: ['operations'],
    },
    async handler(args, bridge) {
      const operations = args.operations as Array<{ tool?: string; args?: Record<string, unknown> }>;

      if (!operations || !Array.isArray(operations) || operations.length === 0) {
        throw new ToolError('operations must be a non-empty array');
      }
      if (operations.length > MAX_BATCH_OPERATIONS) {
        throw new ToolError(`maximum ${MAX_BATCH_OPERATIONS} operations per batch`);
      }

      // Reject nested batch calls to prevent unbounded recursion
      for (const op of operations) {
        if (op.tool === 'word_batch') {
          throw new ToolError('word_batch cannot be nested inside another batch. Flatten your operations into a single batch call.');
        }
      }

      const results: ResultEntry[] = [];
      let nativeBuf: OpEntry[] = [];
      let stopped = false;

      const flushNative = async (): Promise<boolean> => {
        if (nativeBuf.length === 0) return true;
        const batchOps = nativeBuf.map(item => ({
          action: actionMap.get(item.tool)!,
          params: item.args,
        }));
        const batchResult = await bridge.send<BatchResult>('batchExecute', { operations: batchOps });
        for (const r of batchResult.results) {
          const item = nativeBuf[r.index]!;
          if (r.success) {
            results.push({ index: item.originalIndex, tool: item.tool, success: true, result: r.result });
          } else {
            results.push({ index: item.originalIndex, tool: item.tool, success: false, error: r.error });
            nativeBuf = [];
            return false;
          }
        }
        nativeBuf = [];
        return true;
      };

      for (let i = 0; i < operations.length && !stopped; i++) {
        const op = operations[i]!;
        if (!op.tool) {
          results.push({ index: i, tool: '', success: false, error: 'Missing tool name' });
          stopped = true;
          break;
        }

        const handler = registry.get(op.tool);
        if (!handler) {
          const ok = await flushNative();
          if (!ok) { stopped = true; break; }
          results.push({ index: i, tool: op.tool, success: false, error: 'Unknown tool: ' + op.tool });
          stopped = true;
          break;
        }

        // If this tool has a bridge action mapping, it's a native op — buffer it
        if (actionMap.has(op.tool)) {
          // Run server-side validate callback before buffering (same validation
          // that runs on direct calls — prevents batch from bypassing it)
          const validate = validators.get(op.tool);
          if (validate) {
            try {
              validate(op.args || {});
            } catch (e) {
              results.push({ index: i, tool: op.tool, success: false, error: (e as Error).message });
              stopped = true;
              break;
            }
          }
          nativeBuf.push({ tool: op.tool, args: op.args || {}, originalIndex: i });
        } else {
          // Server-composed tool — flush buffer first, then execute individually
          const ok = await flushNative();
          if (!ok) { stopped = true; break; }
          try {
            const toolResult = await handler(op.args || {}, bridge);
            if (toolResult.isError) {
              results.push({ index: i, tool: op.tool, success: false, error: toolResult.content[0]?.text });
              stopped = true;
            } else {
              const parsed = JSON.parse(toolResult.content[0]?.text || '{}');
              results.push({ index: i, tool: op.tool, success: true, result: parsed });
            }
          } catch (e) {
            results.push({ index: i, tool: op.tool, success: false, error: (e as Error).message });
            stopped = true;
          }
        }
      }

      if (!stopped) await flushNative();

      const succeeded = results.filter(r => r.success).length;
      return jsonResult({
        completed: succeeded,
        failed: results.length - succeeded,
        total: operations.length,
        results,
      });
    },
  };
}
