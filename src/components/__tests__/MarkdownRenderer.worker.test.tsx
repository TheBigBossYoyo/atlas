/**
 * Exercises `MarkdownRenderer`'s Worker branch end-to-end within jsdom: a
 * mock `Worker` stands in for the real `markdownParse.worker.ts` (jsdom has
 * no real `Worker` global — see `useMarkdownHastTree.test.ts`'s header) but
 * computes its response with the SAME `parseMarkdownToHast` the real worker
 * uses, so this still proves the branch wires `MarkdownRenderer` ->
 * `useMarkdownHastTree` -> `WorkerMarkdownBody` -> `renderHastTree` together
 * correctly: a loading spinner while the (fake, but real-parse-backed)
 * worker "runs," then the fully rendered document once it replies.
 */
import { afterEach, describe, expect, it, vi } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';

import { MarkdownRenderer } from '../MarkdownRenderer';
import { MARKDOWN_WORKER_BYTE_THRESHOLD } from '../markdown/sizeThresholds';
import { parseMarkdownToHast } from '../markdown/markdownPipeline';
import type { MarkdownWorkerRequest, MarkdownWorkerResponse } from '../markdown/markdownParse.worker';

function buildLargeMarkdown(): string {
  const paragraph = '## Section\n\nSome **bold** text with a [link](https://example.com).\n\n';
  return paragraph.repeat(Math.ceil(MARKDOWN_WORKER_BYTE_THRESHOLD / paragraph.length) + 1);
}

describe('MarkdownRenderer — Worker branch', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('shows a loading state, then renders the worker-parsed tree, for a document past the Worker threshold', async () => {
    class FakeWorker {
      onmessage: ((event: MessageEvent) => void) | null = null;
      onerror: ((event: ErrorEvent) => void) | null = null;
      postMessage(message: MarkdownWorkerRequest) {
        // Mirrors markdownParse.worker.ts's own onmessage handler, just
        // invoked synchronously-via-microtask instead of in a real Worker.
        queueMicrotask(() => {
          const response: MarkdownWorkerResponse = { ok: true, tree: parseMarkdownToHast(message.markdown) };
          this.onmessage?.({ data: response } as MessageEvent);
        });
      }
      terminate() {}
    }
    vi.stubGlobal('Worker', FakeWorker);

    const markdown = buildLargeMarkdown();
    expect(markdown.length).toBeGreaterThanOrEqual(MARKDOWN_WORKER_BYTE_THRESHOLD);

    const { container } = render(<MarkdownRenderer markdown={markdown} />);

    // Two loading stages now: `WorkerMarkdownBody` is itself `lazy()`-loaded
    // (see MarkdownRenderer.tsx's PERF comment), so the very first render is
    // briefly a plain Suspense fallback with no `role="status"` yet — only
    // once that chunk resolves does `WorkerMarkdownBody` mount and show its
    // own status spinner while the (fake) worker "parses."
    await screen.findByRole('status');

    await waitFor(() => {
      expect(container.querySelectorAll('h2')).not.toHaveLength(0);
    });

    expect(container.querySelector('#markdown-content')).not.toBeNull();
    expect(container.querySelector('a[href="https://example.com"]')).not.toBeNull();
    expect(container.querySelector('strong')?.textContent).toBe('bold');
  });

  it('shows an error message when the worker reports a failure', async () => {
    class FailingFakeWorker {
      onmessage: ((event: MessageEvent) => void) | null = null;
      onerror: ((event: ErrorEvent) => void) | null = null;
      postMessage() {
        queueMicrotask(() => {
          const response: MarkdownWorkerResponse = { ok: false, error: 'out of memory' };
          this.onmessage?.({ data: response } as MessageEvent);
        });
      }
      terminate() {}
    }
    vi.stubGlobal('Worker', FailingFakeWorker);

    const markdown = buildLargeMarkdown();
    render(<MarkdownRenderer markdown={markdown} />);

    await waitFor(() => {
      expect(screen.getByRole('alert')).toHaveTextContent('out of memory');
    });
  });
});
