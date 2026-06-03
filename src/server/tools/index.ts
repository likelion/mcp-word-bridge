import type { ToolDefinition, ToolHandler } from '../types';
import { documentTools } from './document';
import { paragraphTools } from './paragraphs';
import { searchTools } from './search';
import { formattingTools } from './formatting';
import { tableTools } from './tables';
import { listTools } from './lists';
import { commentTools } from './comments';
import { footnoteTools } from './footnotes';
import { bookmarkTools } from './bookmarks';
import { hyperlinkTools } from './hyperlinks';
import { contentControlTools } from './content-controls';
import { headerFooterTools } from './headers-footers';
import { imageTools } from './images';
import { layoutTools } from './layout';
import { fieldTools } from './fields';
import { trackingTools } from './tracking';
import { propertyTools } from './properties';
import { advancedTools } from './advanced';
import { equationTools } from './equations';
import { createBatchTool } from './batch';

/**
 * All tool definitions, assembled in a single registry.
 * The batch tool is added last since it needs access to all other handlers.
 */
export function buildToolRegistry(): {
  tools: ToolDefinition[];
  handlers: Map<string, ToolHandler>;
} {
  const allTools: ToolDefinition[] = [
    ...documentTools,
    ...paragraphTools,
    ...searchTools,
    ...formattingTools,
    ...tableTools,
    ...listTools,
    ...commentTools,
    ...footnoteTools,
    ...bookmarkTools,
    ...hyperlinkTools,
    ...contentControlTools,
    ...headerFooterTools,
    ...imageTools,
    ...layoutTools,
    ...fieldTools,
    ...trackingTools,
    ...propertyTools,
    ...advancedTools,
    ...equationTools,
  ];

  // Build handler map and action map from all non-batch tools
  const handlers = new Map<string, ToolHandler>();
  const actionMap = new Map<string, string>();
  for (const tool of allTools) {
    handlers.set(tool.name, tool.handler);
    // Tools created with forwardTool have a bridgeAction property
    if ('bridgeAction' in tool && typeof tool.bridgeAction === 'string') {
      actionMap.set(tool.name, tool.bridgeAction);
    }
  }

  // Create batch tool with access to all handlers and the action map
  const batchTool = createBatchTool(handlers, actionMap);
  allTools.push(batchTool);
  handlers.set(batchTool.name, batchTool.handler);

  return { tools: allTools, handlers };
}
