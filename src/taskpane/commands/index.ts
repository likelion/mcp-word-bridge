/**
 * Command registry — maps action names to handler functions.
 * Each handler receives (ctx: Word.RequestContext, params: Record<string, any>)
 * and returns a result object.
 */

import { paragraphCommands } from './paragraphs';
import { searchCommands } from './search';
import { formattingCommands } from './formatting';
import { tableCommands } from './tables';
import { listCommands } from './lists';
import { commentCommands } from './comments';
import { footnoteCommands } from './footnotes';
import { bookmarkCommands } from './bookmarks';
import { hyperlinkCommands } from './hyperlinks';
import { contentControlCommands } from './content-controls';
import { headerFooterCommands } from './headers-footers';
import { imageCommands } from './images';
import { layoutCommands } from './layout';
import { fieldCommands } from './fields';
import { trackingCommands } from './tracking';
import { propertyCommands } from './properties';
import { documentCommands } from './document';
import { ooxmlCommands } from './ooxml';
import { batchCommands } from './batch';

export type CommandHandler = (ctx: any, params: Record<string, any>) => Promise<any>;

export const commandRegistry: Record<string, CommandHandler> = {
  ...documentCommands,
  ...paragraphCommands,
  ...searchCommands,
  ...formattingCommands,
  ...tableCommands,
  ...listCommands,
  ...commentCommands,
  ...footnoteCommands,
  ...bookmarkCommands,
  ...hyperlinkCommands,
  ...contentControlCommands,
  ...headerFooterCommands,
  ...imageCommands,
  ...layoutCommands,
  ...fieldCommands,
  ...trackingCommands,
  ...propertyCommands,
  ...ooxmlCommands,
  ...batchCommands,
};
