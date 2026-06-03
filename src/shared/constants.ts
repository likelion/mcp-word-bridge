/** Maximum search string length accepted by Word API */
export const MAX_SEARCH_LENGTH = 255;

/** Maximum operations per batch call */
export const MAX_BATCH_OPERATIONS = 50;

/** Default timeout for standard bridge operations (ms) */
export const TIMEOUT_DEFAULT = 30_000;

/** Extended timeout for heavy operations (ms) */
export const TIMEOUT_HEAVY = 60_000;

/** Heavy operations that get extended timeout */
export const HEAVY_OPERATIONS = new Set([
  'insertHtml',
  'insertOoxml',
  'getStyles',
  'insertTableOfContents',
  'batchExecute',
]);

/** Maximum WebSocket payload size (bytes) */
export const MAX_PAYLOAD = 10 * 1024 * 1024;
