/**
 * The large-document half of `MarkdownRenderer`: parses `markdown` off the
 * main thread via `useMarkdownHastTree` (backed by `markdownParse.worker.ts`)
 * and renders the resulting hast tree with `renderHastTree`, which mirrors
 * `<ReactMarkdown>`'s own rendering step exactly — see that file's header.
 *
 * Renders into the same `.markdown-body#markdown-content` container
 * `MarkdownRenderer`'s inline (`<ReactMarkdown>`) branch does, and forwards
 * the same ref, so search (`useSearch`'s `TreeWalker`), the table of
 * contents, export, and print all keep working against it exactly as they
 * do for a small document — none of them know or care which branch rendered
 * the content.
 */
import { forwardRef } from 'react';
import { Loader2 } from 'lucide-react';
import type { Components } from 'react-markdown';
import { useTranslate } from '../../i18n';
import { useMarkdownHastTree } from './useMarkdownHastTree';
import { renderHastTree } from './renderHastTree';

interface WorkerMarkdownBodyProps {
  readonly markdown: string;
  readonly components: Partial<Components>;
}

export const WorkerMarkdownBody = forwardRef<HTMLDivElement, WorkerMarkdownBodyProps>(
  function WorkerMarkdownBody({ markdown, components }, ref) {
    const t = useTranslate();
    const state = useMarkdownHastTree(markdown);

    if (state.status === 'loading') {
      return (
        <div ref={ref} className="markdown-body" id="markdown-content">
          <div className="viewer-fallback viewer-loading" role="status" aria-live="polite">
            <Loader2 size={28} className="viewer-fallback__icon viewer-loading__spinner" aria-hidden="true" />
            <p className="viewer-fallback__title">{t('markdownWorker.parsing')}</p>
          </div>
        </div>
      );
    }

    if (state.status === 'error') {
      return (
        <div ref={ref} className="markdown-body" id="markdown-content">
          <div className="viewer-fallback" role="alert">
            <p className="viewer-fallback__title">{t('markdownWorker.error', { error: state.error })}</p>
          </div>
        </div>
      );
    }

    return (
      <div ref={ref} className="markdown-body" id="markdown-content">
        {renderHastTree(state.tree, components)}
      </div>
    );
  }
);
