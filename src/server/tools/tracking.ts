import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';

export const getTrackedChanges = forwardTool(
  'word_get_tracked_changes',
  '[Track Changes] Get all tracked changes with index, type (Added/Deleted), author, date, and text.',
  { properties: {} },
  'getTrackedChanges',
);

export const acceptTrackedChange = forwardTool(
  'word_accept_tracked_change',
  '[Track Changes] Accept a single tracked change by its index.',
  {
    properties: {
      index: { type: 'number', description: 'Change index (0-based)' },
    },
    required: ['index'],
  },
  'acceptTrackedChange',
);

export const rejectTrackedChange = forwardTool(
  'word_reject_tracked_change',
  '[Track Changes] Reject a single tracked change by its index.',
  {
    properties: {
      index: { type: 'number', description: 'Change index (0-based)' },
    },
    required: ['index'],
  },
  'rejectTrackedChange',
);

export const acceptAllTrackedChanges = forwardTool(
  'word_accept_all_tracked_changes',
  '[Track Changes] Accept all tracked changes at once.',
  { properties: {} },
  'acceptAllTrackedChanges',
);

export const rejectAllTrackedChanges = forwardTool(
  'word_reject_all_tracked_changes',
  '[Track Changes] Reject all tracked changes at once.',
  { properties: {} },
  'rejectAllTrackedChanges',
);

export const trackingTools: ToolDefinition[] = [
  getTrackedChanges,
  acceptTrackedChange,
  rejectTrackedChange,
  acceptAllTrackedChanges,
  rejectAllTrackedChanges,
];
