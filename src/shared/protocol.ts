/**
 * WebSocket protocol between MCP server and Word taskpane.
 * Both sides must agree on these message shapes.
 */

// --- Outbound: Server → Taskpane ---

export interface BridgeRequest {
  id: string;
  action: string;
  params: Record<string, unknown>;
}

// --- Inbound: Taskpane → Server ---

export interface BridgeResponseMessage {
  type: 'response';
  id: string;
  result?: unknown;
  error?: string;
}

// --- Paragraph data returned by getParagraphs ---

export interface ParagraphInfo {
  index: number;
  text: string;
  style: string;
  alignment: string;
  isListItem: boolean;
  inTable: boolean;
  isTocEntry: boolean;
  outlineLevel: number;
}

export interface GetParagraphsResult {
  total: number;
  count: number;
  paragraphs: ParagraphInfo[];
  warning?: string;
}

// --- Paragraph detail from getParagraphByIndex ---

export interface ParagraphDetail {
  text: string;
  style: string;
  alignment: string;
  firstLineIndent: number;
  leftIndent: number;
  rightIndent: number;
  lineSpacing: number;
  spaceBefore: number;
  spaceAfter: number;
  outlineLevel: number;
  isListItem: boolean;
  font: {
    name: string;
    size: number;
    bold: boolean;
    italic: boolean;
    color: string;
    underline: string;
  };
}

// --- Search result ---

export interface SearchResult {
  count: number;
  matches: Array<{ index: number; text: string }>;
}

// --- OOXML result ---

export interface OoxmlResult {
  ooxml: string;
}

// --- Batch execution ---

export interface BatchOperation {
  action: string;
  params: Record<string, unknown>;
}

export interface BatchResult {
  results: Array<{
    index: number;
    success: boolean;
    result?: unknown;
    error?: string;
  }>;
}

// --- Generic success ---

export interface SuccessResult {
  success: true;
  warning?: string;
}
