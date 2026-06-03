import { MAX_SEARCH_LENGTH } from '../shared/constants';
import { ToolError } from './types';

/** Validate search text is within Word's limit */
export function checkSearchLength(text: string): void {
  if (text.length > MAX_SEARCH_LENGTH) {
    throw new ToolError('Search text is too long (max ~255 characters). Shorten the text or use a substring.');
  }
}

/** Validate a non-empty string parameter */
export function checkNonEmpty(value: unknown, name: string): asserts value is string {
  if (!value || typeof value !== 'string' || value.trim() === '') {
    throw new ToolError(`${name} must be a non-empty string.`);
  }
}

/** Validate a non-negative integer index */
export function checkNonNegative(value: unknown, name: string): asserts value is number {
  if (typeof value !== 'number' || value < 0) {
    throw new ToolError(`${name} must be non-negative.`);
  }
}

/** Validate index is within bounds */
export function checkBounds(index: number, count: number, name: string): void {
  if (index >= count) {
    throw new ToolError(
      `${name} ${index} out of range. Document has ${count} paragraphs (valid indices: 0-${count - 1}).`,
    );
  }
}

/** Validate occurrence parameter */
export function checkOccurrence(occurrence: number | undefined, count: number): number {
  const idx = occurrence ?? 0;
  if (idx < 0) {
    throw new ToolError('occurrence must be non-negative (0-indexed).');
  }
  if (idx >= count) {
    throw new ToolError(
      `Occurrence ${idx} not found (only ${count} match${count === 1 ? '' : 'es'}).`,
    );
  }
  return idx;
}

/** Normalize alignment aliases to Word API values */
const ALIGNMENT_MAP: Record<string, string> = {
  Left: 'Left',
  Center: 'Centered',
  Centered: 'Centered',
  Right: 'Right',
  Justify: 'Justified',
  Justified: 'Justified',
};

export function normalizeAlignment(value: string | undefined): string | null {
  if (!value) return null;
  const mapped = ALIGNMENT_MAP[value];
  if (!mapped) {
    throw new ToolError(`Invalid alignment: "${value}". Valid values: Left, Center, Right, Justified.`);
  }
  return mapped;
}

/** Validate hex color */
export function checkHexColor(color: string, name: string): void {
  if (!/^#[0-9A-Fa-f]{6}$/.test(color)) {
    throw new ToolError(`${name} must be a valid hex color (e.g. #FF0000).`);
  }
}

/** Validate URL format */
export function checkUrl(url: string): void {
  if (!/^https?:\/\/.+/i.test(url)) {
    throw new ToolError('URL must be a valid HTTP or HTTPS URL (e.g. https://example.com).');
  }
  if (/[<>"{}|\\^`]/.test(url)) {
    throw new ToolError(`Malformed URL: "${url}". URL contains invalid characters that must be percent-encoded.`);
  }
}
