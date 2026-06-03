import type { CommandHandler } from './index';
import { MAX_SEARCH_LENGTH } from '../../shared/constants';

// =============================================================================
// Shared validation utilities (imported by other command modules)
// =============================================================================

/** Validate search string length before calling Word API */
export function checkSearchLength(text: string): void {
  if (text && text.length > MAX_SEARCH_LENGTH) {
    throw new Error('Search text is too long (max ~255 characters). Shorten the text or use a substring.');
  }
}

/** Strip Word internal control characters from text before returning to MCP client */
export function sanitizeText(text: string): string {
  return text.replace(/[\u0002\u0003\u0005\u0013\u0014\u0015]/g, '');
}

/** Validate a non-negative integer index parameter */
export function checkIndex(value: unknown, name: string): asserts value is number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0) {
    throw new Error(`${name} must be a non-negative integer.`);
  }
}

/** Validate a hex color string (e.g. #FF0000) */
export function checkHexColor(color: string, name: string = 'color'): void {
  if (!/^#[0-9A-Fa-f]{6}$/.test(color)) {
    throw new Error(`${name} must be a valid hex color (e.g. #FF0000).`);
  }
}

/** Validate occurrence parameter is non-negative */
export function checkOccurrence(occurrence: unknown): void {
  if (occurrence !== undefined && (typeof occurrence !== 'number' || occurrence < 0)) {
    throw new Error('occurrence must be non-negative (0-indexed).');
  }
}

/** Validate a required non-empty string parameter */
export function checkNonEmptyString(value: unknown, name: string): asserts value is string {
  if (!value || typeof value !== 'string' || value.trim() === '') {
    throw new Error(`${name} cannot be empty. Provide a non-empty search string.`);
  }
}

/** Validate anchor text (non-empty + length check) */
export function checkAnchorText(value: unknown): asserts value is string {
  checkNonEmptyString(value, 'anchorText');
  checkSearchLength(value);
}

const ALIGNMENT_MAP: Record<string, string> = {
  left: 'Left', center: 'Centered', centered: 'Centered',
  right: 'Right', justify: 'Justified', justified: 'Justified',
};

/** Normalize and validate alignment value. Returns null if undefined. */
export function checkAlignment(value: string | undefined): string | null {
  if (!value) return null;
  const mapped = ALIGNMENT_MAP[value.toLowerCase()];
  if (!mapped) throw new Error(`Invalid alignment: "${value}". Valid values: Left, Center, Right, Justified.`);
  return mapped;
}

/** Validate style exists in the document. Returns once confirmed or throws. */
export async function checkStyleExists(ctx: any, styleName: string): Promise<void> {
  const styleObj = ctx.document.getStyles().getByNameOrNullObject(styleName);
  styleObj.load('nameLocal');
  await ctx.sync();
  if (styleObj.isNullObject) {
    throw new Error(`Style not found: "${styleName}". Use word_get_styles to see available styles.`);
  }
}

/** Valid Word highlight color names */
export const HIGHLIGHT_COLORS = ['Yellow', 'Green', 'Cyan', 'Magenta', 'Blue', 'Red', 'DarkBlue', 'DarkCyan', 'DarkGreen', 'DarkMagenta', 'DarkRed', 'DarkYellow', 'Gray25', 'Gray50', 'Black', 'White', 'NoHighlight'] as const;

/** Validate highlight color (named or hex) */
export function checkHighlightColor(color: string): void {
  const isNamed = HIGHLIGHT_COLORS.some(c => c.toLowerCase() === color.toLowerCase());
  const isHex = /^#[0-9A-Fa-f]{6}$/.test(color);
  if (!isNamed && !isHex) {
    throw new Error(`Invalid highlightColor: "${color}". Use a named color (${HIGHLIGHT_COLORS.join(', ')}) or a hex color (e.g. #FFFF00).`);
  }
}

export const documentCommands: Record<string, CommandHandler> = {
  async getDocumentText(ctx) {
    const body = ctx.document.body;
    body.load('text');
    await ctx.sync();
    return { text: sanitizeText(body.text) };
  },

  async getDocumentProperties(ctx) {
    const props = ctx.document.properties;
    props.load('title,subject,author,keywords,comments,lastAuthor,revisionNumber,creationDate,lastSaveTime,category,company,manager,format,applicationName,template,lastPrintDate,security');
    ctx.document.load('changeTrackingMode');
    await ctx.sync();
    let path: string | null = null;
    try {
      const fileProps: any = await new Promise((resolve) => {
        Office.context.document.getFilePropertiesAsync(resolve);
      });
      if (fileProps.status === Office.AsyncResultStatus.Succeeded && fileProps.value.url) {
        const url = fileProps.value.url;
        if (/\.(docx?|docm|dotx?|dotm)$/i.test(url)) path = url;
      }
    } catch { /* ignore */ }
    return {
      title: props.title, subject: props.subject, author: props.author,
      keywords: props.keywords, comments: props.comments, category: props.category,
      company: props.company, manager: props.manager, format: props.format,
      applicationName: props.applicationName, template: props.template,
      lastAuthor: props.lastAuthor, revisionNumber: props.revisionNumber,
      creationDate: props.creationDate, lastSaveTime: props.lastSaveTime,
      lastPrintDate: props.lastPrintDate, security: props.security,
      changeTrackingMode: ctx.document.changeTrackingMode, path,
    };
  },

  async setDocumentProperties(ctx, p) {
    const props = ctx.document.properties;
    if (p.title !== undefined) props.title = p.title;
    if (p.subject !== undefined) props.subject = p.subject;
    if (p.author !== undefined) props.author = p.author;
    if (p.keywords !== undefined) props.keywords = p.keywords;
    if (p.comments !== undefined) props.comments = p.comments;
    if (p.category !== undefined) props.category = p.category;
    if (p.company !== undefined) props.company = p.company;
    if (p.manager !== undefined) props.manager = p.manager;
    if (p.format !== undefined) props.format = p.format;
    await ctx.sync();
    return { success: true };
  },

  async setChangeTracking(ctx, p) {
    const validModes = ['TrackAll', 'TrackMineOnly', 'Off'];
    const mode = p.mode || 'TrackAll';
    if (!validModes.includes(mode)) throw new Error(`Invalid mode: "${mode}". Valid values: TrackAll, TrackMineOnly, Off`);
    ctx.document.changeTrackingMode = mode;
    await ctx.sync();
    return { success: true, mode };
  },

  async saveDocument(ctx) {
    ctx.document.save();
    await ctx.sync();
    let path: string | null = null;
    try {
      const fileProps: any = await new Promise((resolve) => {
        Office.context.document.getFilePropertiesAsync(resolve);
      });
      if (fileProps.status === Office.AsyncResultStatus.Succeeded && fileProps.value.url) {
        if (/\.(docx?|docm|dotx?|dotm)$/i.test(fileProps.value.url)) path = fileProps.value.url;
      }
    } catch { /* ignore */ }
    const result: any = { success: true };
    if (!path) result.warning = 'Document has no file path — use Save As in Word to set a location.';
    return result;
  },

  async clearDocument(ctx) {
    ctx.document.body.clear();
    await ctx.sync();
    const paras = ctx.document.body.paragraphs;
    paras.load('text');
    await ctx.sync();
    if (paras.items.length > 0) {
      paras.items[0].style = 'Normal';
      paras.items[0].getRange('Start').select();
      await ctx.sync();
    }
    return { success: true };
  },

  async getWordCount(ctx) {
    const body = ctx.document.body;
    body.load('text');
    const paragraphs = body.paragraphs;
    paragraphs.load('text');
    await ctx.sync();
    const text = sanitizeText(body.text);
    const words = text.split(/\s+/).filter((w: string) => w.length > 0).length;
    return { words, characters: text.length, charactersNoSpaces: text.replace(/\s/g, '').length, paragraphs: paragraphs.items.length };
  },

  async getStyles(ctx) {
    const styles = ctx.document.getStyles();
    styles.load('nameLocal,type,builtIn');
    await ctx.sync();
    const items = styles.items.map((s: any) => ({ name: s.nameLocal, type: s.type, builtIn: s.builtIn }));
    const returned = items.slice(0, 80);
    return { count: returned.length, total: items.length, styles: returned };
  },

  async getCoauthors(ctx) {
    const coauthoring = ctx.document.coauthoring;
    coauthoring.load('canCoauthor,canMerge,pendingUpdates');
    const authors = coauthoring.authors;
    authors.load('name,email');
    await ctx.sync();
    return {
      canCoauthor: coauthoring.canCoauthor, canMerge: coauthoring.canMerge,
      pendingUpdates: coauthoring.pendingUpdates,
      authors: authors.items.map((a: any) => ({ name: a.name, email: a.email })),
    };
  },

  async getApiVersion() {
    const sets = [
      'WordApi 1.9', 'WordApi 1.8', 'WordApi 1.7', 'WordApi 1.6', 'WordApi 1.5',
      'WordApi 1.4', 'WordApi 1.3', 'WordApi 1.2', 'WordApi 1.1',
      'WordApiDesktop 1.5', 'WordApiDesktop 1.4', 'WordApiDesktop 1.3',
      'WordApiDesktop 1.2', 'WordApiDesktop 1.1',
    ];
    const supported: Record<string, boolean> = {};
    sets.forEach(s => {
      const [name, ver] = s.split(' ');
      supported[s] = Office.context.requirements.isSetSupported(name, ver);
    });
    return { platform: Office.context.platform, host: Office.context.host, diagnostics: Office.context.diagnostics, supportedSets: supported };
  },

  async ping() {
    return { status: 'ok', wordReady: true };
  },
};

// Re-export for use in other command files
declare const Office: any;
