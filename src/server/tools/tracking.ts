import type { ToolDefinition } from '../types';
import { forwardTool } from './helpers';
import { checkNonNegative } from '../validation';
import { ToolError } from '../types';

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
  (args) => { checkNonNegative(args.index, 'index'); },
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
  (args) => { checkNonNegative(args.index, 'index'); },
);

export const acceptAllTrackedChanges = forwardTool(
  'word_accept_all_tracked_changes',
  '[Track Changes] Accept all tracked changes at once.',
  { properties: {} },
  'acceptAllTrackedChanges',
);

export const acceptTrackedChangesInRange = forwardTool(
  'word_accept_tracked_changes_in_range',
  '[Track Changes] Accept tracked changes within a paragraph index range. Useful for accepting changes in a specific section without affecting the rest of the document.',
  {
    properties: {
      startIndex: { type: 'number', description: 'First paragraph index (0-based, inclusive)' },
      endIndex: { type: 'number', description: 'Last paragraph index (0-based, exclusive). Omit to accept through end of document.' },
    },
    required: ['startIndex'],
  },
  'acceptTrackedChangesInRange',
  (args) => {
    checkNonNegative(args.startIndex, 'startIndex');
    if (args.endIndex !== undefined) {
      checkNonNegative(args.endIndex, 'endIndex');
      if ((args.endIndex as number) <= (args.startIndex as number)) {
        throw new ToolError('endIndex must be greater than startIndex.');
      }
    }
  },
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
  acceptTrackedChangesInRange,
  rejectAllTrackedChanges,
];
