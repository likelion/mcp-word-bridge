import type { ToolDefinition } from '../types';
import { ToolError } from '../types';
import { jsonResult } from './helpers';
import { checkOccurrence } from '../validation';
import type { SearchResult } from '../../shared/protocol';

// Lazy-loaded dependencies (heavy, only needed for equations)
let latexToOmml: typeof import('../../lib/equations').latexToOmml;
let buildEquationOoxml: typeof import('../../lib/equations').buildEquationOoxml;

async function loadEquationLib() {
  if (!latexToOmml) {
    const lib = await import('../../lib/equations');
    latexToOmml = lib.latexToOmml;
    buildEquationOoxml = lib.buildEquationOoxml;
  }
}

const MARKER = '\u200B\uFEFF\u200B';

export const insertEquation: ToolDefinition = {
  name: 'word_insert_equation',
  description: '[Equations] Insert a LaTeX math equation as a native editable Word equation. Display mode (default) inserts a centered block. Inline mode inserts after a search match (provide anchorText) or at cursor.',
  schema: {
    properties: {
      latex: { type: 'string', description: 'LaTeX math expression' },
      displayMode: { type: 'boolean', description: 'true (default) = block, false = inline' },
      location: { type: 'string', enum: ['Start', 'End'], description: 'For display mode. Default: End' },
      anchorText: { type: 'string', description: 'For inline: search for this text and insert after it' },
      occurrence: { type: 'number', description: '0=first, 1=second, etc. Default: 0' },
      matchCase: { type: 'boolean', description: 'Default: false' },
    },
    required: ['latex'],
  },
  async handler(args, bridge) {
    await loadEquationLib();
    const { mml2omml } = await import('mathml2omml');

    const latex = args.latex as string;
    const displayMode = args.displayMode !== false;

    let result;
    try {
      result = latexToOmml(latex, displayMode, mml2omml);
    } catch (e) {
      const msg = (e as Error).message;
      const friendly = msg.startsWith('LaTeX parse error') || msg.startsWith('"latex"') ? msg : 'Error: ' + msg;
      return { content: [{ type: 'text', text: friendly }], isError: true };
    }

    const ooxml = buildEquationOoxml(result.omml, displayMode);

    if (displayMode) {
      await bridge.send('insertOoxml', { ooxml, location: args.location || 'End' });
    } else if (args.anchorText) {
      const anchorText = args.anchorText as string;
      const matchCase = (args.matchCase as boolean) || false;

      const searchResult = await bridge.send<SearchResult>('search', { query: anchorText, matchCase });
      if (!searchResult || searchResult.count === 0) throw new ToolError('Anchor not found: ' + anchorText);
      const occurrence = checkOccurrence(args.occurrence as number | undefined, searchResult.count);

      await bridge.send('insertText', { text: ' ' + MARKER, after: anchorText, occurrence, matchCase });
      try {
        await bridge.send('insertOoxmlAtSelection', { ooxml });
      } finally {
        await bridge.send('searchAndReplace', { find: MARKER, replace: '' });
      }
    } else {
      await bridge.send('insertOoxmlAtSelection', { ooxml });
    }

    return jsonResult({ success: true, displayMode, latex });
  },
};

export const equationTools: ToolDefinition[] = [
  insertEquation,
];
