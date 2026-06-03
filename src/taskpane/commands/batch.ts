import type { CommandHandler } from './index';
import { log } from '../log';

export const batchCommands: Record<string, CommandHandler> = {
  async batchExecute(ctx, p) {
    if (!p.operations || !Array.isArray(p.operations) || p.operations.length === 0)
      throw new Error('operations must be a non-empty array');
    log('  batch: ' + p.operations.length + ' ops', 'log-batch');
    // Late-binding import to avoid circular dependency
    const { commandRegistry } = await import('./index');
    const results: any[] = [];
    for (let i = 0; i < p.operations.length; i++) {
      const op = p.operations[i];
      try {
        const handler = commandRegistry[op.action];
        if (!handler) throw new Error('Unknown action: ' + op.action);
        log('  [' + (i + 1) + '/' + p.operations.length + '] ' + op.action, 'log-batch');
        const result = await handler(ctx, op.params || {});
        results.push({ index: i, success: true, result });
      } catch (e: any) {
        log('  [' + (i + 1) + '/' + p.operations.length + '] ' + (op.action || '?') + ' ✗ ' + e.message, 'log-err');
        results.push({ index: i, success: false, error: e.message });
        break;
      }
    }
    return { results };
  },
};
