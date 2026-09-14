/**
 * Markdown → `.md` export (wave1 follow-up).
 *
 * Previously called `window.electronAPI.saveFile(markdown, name)`
 * *positionally* — `saveFile`'s real signature takes a single
 * `{ content, suggestedName, filters?, existingPath? }` request object, so
 * `name` was silently passed as... nothing `saveFile` reads, and the browser
 * download fallback never ran because the (wrong) call still "succeeded".
 * Now routed through the shared `saveTextOutput` helper like every other
 * export function (UX-11).
 */
import { sanitizeFileName, saveTextOutput, type SaveFilter } from './download';
import { toFriendlyError } from '../friendlyLibraryError';

const MARKDOWN_MIME = 'text/markdown;charset=utf-8';
const MARKDOWN_FILTERS: readonly SaveFilter[] = [{ name: 'Markdown', extensions: ['md'] }];

/** Downloads the raw markdown source as a `.md` file. */
export async function exportMarkdown(markdown: string, fileName: string): Promise<void> {
  try {
    const name = sanitizeFileName(fileName, 'md');
    await saveTextOutput(markdown, name, MARKDOWN_MIME, MARKDOWN_FILTERS);
  } catch (err) {
    console.error('[export] exportMarkdown failed:', err);
    throw toFriendlyError(err, 'Markdown export failed');
  }
}
