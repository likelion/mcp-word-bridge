import { MAX_SEARCH_LENGTH } from '../shared/constants';
import { ToolError } from './types';

/** Word special find/replace codes that manipulate document structure */
const WORD_SPECIAL_CODES = /\^(p|w|t|l|m|b|n|s|d|a|e|f|g|v|~|\^|\-|13|11|14|12|07|09|30|31|32|34|36|37|38|39|40|41|42|43|44|45|46|47|92|94|127|129|130|131|132|133|134|135|136|137|138|139|140|141|142|143|144|145|146|147|148|149|150|151|152|153|154|155|156|157|158|159|160|161|162|163|164|165|166|167|168|169|170|171|172|173|174|175|176|177|178|179|180|181|182|183|184|185|186|187|188|189|190|191|192|193|194|195|196|197|198|199|200|201|202|203|204|205|206|207|208|209|210|211|212|213|214|215|216|217|218|219|220|221|222|223|224|225|226|227|228|229|230|231|232|233|234|235|236|237|238|239|240|241|242|243|244|245|246|247|248|249|250|251|252|253|254|255)/;

/** Validate that search/replace strings do not contain Word special codes */
export function checkNoSpecialCodes(text: string, paramName: string): void {
  const match = text.match(WORD_SPECIAL_CODES);
  if (match) {
    throw new ToolError(
      `${paramName} contains Word special code "${match[0]}" which can corrupt document structure. ` +
      `Use literal text only. Common special codes: ^p (paragraph mark), ^t (tab), ^w (whitespace), ^13 (paragraph mark).`,
    );
  }
}

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
  if (typeof value !== 'number' || value < 0 || !Number.isInteger(value)) {
    throw new ToolError(`${name} must be a non-negative integer.`);
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
  left: 'Left',
  center: 'Centered',
  centered: 'Centered',
  right: 'Right',
  justify: 'Justified',
  justified: 'Justified',
};

export function normalizeAlignment(value: string | undefined): string | null {
  if (!value) return null;
  const mapped = ALIGNMENT_MAP[value.toLowerCase()];
  if (!mapped) {
    throw new ToolError(`Invalid alignment: "${value}". Valid values: Left, Center, Right, Justified.`);
  }
  return mapped;
}

/** Maximum allowed spacing/indent value in points (22 inches = 1584pt, matching Word's page width limit) */
export const MAX_SPACING_POINTS = 1584;

/** Maximum key length for custom document properties */
export const MAX_CUSTOM_PROPERTY_KEY_LENGTH = 255;

/** Validate spacing/indent values are within bounds */
export function checkSpacingBounds(value: number, name: string): void {
  if (value > MAX_SPACING_POINTS) {
    throw new ToolError(
      `${name} value ${value} exceeds maximum (${MAX_SPACING_POINTS} points = 22 inches). ` +
      `Use a value between 0 and ${MAX_SPACING_POINTS}.`,
    );
  }
}

/**
 * Validate indent values.
 * - firstLineIndent: allows negative (hanging indent), bounded to ±1584pt.
 * - leftIndent / rightIndent: must be non-negative, bounded to 1584pt.
 */
export function checkIndentBounds(value: number, name: string): void {
  if (name === 'firstLineIndent') {
    // Hanging indent: negative is valid, but bounded
    if (value < -MAX_SPACING_POINTS || value > MAX_SPACING_POINTS) {
      throw new ToolError(
        `${name} value ${value} is out of range. Valid range: -${MAX_SPACING_POINTS} to ${MAX_SPACING_POINTS} points.`,
      );
    }
  } else {
    // leftIndent, rightIndent: non-negative
    if (value < 0) {
      throw new ToolError(`${name} must be non-negative (in points).`);
    }
    if (value > MAX_SPACING_POINTS) {
      throw new ToolError(
        `${name} value ${value} exceeds maximum (${MAX_SPACING_POINTS} points = 22 inches).`,
      );
    }
  }
}

/** Validate custom property key length */
export function checkPropertyKeyLength(key: string): void {
  if (key.length > MAX_CUSTOM_PROPERTY_KEY_LENGTH) {
    throw new ToolError(
      `key must be ${MAX_CUSTOM_PROPERTY_KEY_LENGTH} characters or fewer (got ${key.length}).`,
    );
  }
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
