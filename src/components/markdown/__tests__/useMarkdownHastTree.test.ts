/**
 * `useMarkdownHastTree` — Worker lifecycle tests, mirroring
 * `useSpreadsheetWorkbook.test.ts`'s own pattern (see that file's header):
 * jsdom (this project's test environment) has no `Worker` global, so a
 * minimal mock `Worker` stands in to verify the request/response contract
 * and cleanup, rather than exercising a real Worker end-to-end (that's what
 * `tests/e2e/perf.spec.ts`-style Playwright/Electron coverage is for).
 */
import { renderHook, waitFor } from '@testing-library/react';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { useMarkdownHastTree } from '../useMarkdownHastTree';
import type { MarkdownWorkerRequest } from '../markdownParse.worker';

describe('useMarkdownHastTree', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('posts the markdown to a Worker and starts in the loading state', async () => {
    const postedMessages: unknown[] = [];

    class FakeWorker {
      onmessage: ((event: MessageEvent) => void) | null = null;
      onerror: ((event: ErrorEvent) => void) | null = null;
      postMessage(message: unknown) {
        postedMessages.push(message);
      }
      terminate() {}
    }
    vi.stubGlobal('Worker', FakeWorker);

    const { result } = renderHook(() => useMarkdownHastTree('# Hello'));

    expect(result.current.status).toBe('loading');
    await waitFor(() => expect(postedMessages).toHaveLength(1));
    expect((postedMessages[0] as MarkdownWorkerRequest).markdown).toBe('# Hello');
  });

  it('transitions to ready when the worker replies with a tree', async () => {
    const captured: { onMessage: ((event: MessageEvent) => void) | null } = { onMessage: null };
    const fakeTree = { type: 'root', children: [] };

    class FakeWorker {
      onmessage: ((event: MessageEvent) => void) | null = null;
      onerror: ((event: ErrorEvent) => void) | null = null;
      postMessage() {
        captured.onMessage = (event) => this.onmessage?.(event);
      }
      terminate() {}
    }
    vi.stubGlobal('Worker', FakeWorker);

    const { result } = renderHook(() => useMarkdownHastTree('# Hello'));
    await waitFor(() => expect(captured.onMessage).not.toBeNull());

    captured.onMessage?.({ data: { ok: true, tree: fakeTree } } as MessageEvent);
    await waitFor(() => expect(result.current.status).toBe('ready'));
    if (result.current.status !== 'ready') throw new Error('expected ready');
    expect(result.current.tree).toBe(fakeTree);
  });

  it('transitions to error when the worker reports a parse failure', async () => {
    const captured: { onMessage: ((event: MessageEvent) => void) | null } = { onMessage: null };

    class FakeWorker {
      onmessage: ((event: MessageEvent) => void) | null = null;
      onerror: ((event: ErrorEvent) => void) | null = null;
      postMessage() {
        captured.onMessage = (event) => this.onmessage?.(event);
      }
      terminate() {}
    }
    vi.stubGlobal('Worker', FakeWorker);

    const { result } = renderHook(() => useMarkdownHastTree('# Hello'));
    await waitFor(() => expect(captured.onMessage).not.toBeNull());

    captured.onMessage?.({ data: { ok: false, error: 'boom' } } as MessageEvent);
    await waitFor(() => expect(result.current.status).toBe('error'));
    if (result.current.status !== 'error') throw new Error('expected error');
    expect(result.current.error).toBe('boom');
  });

  it('transitions to error when the worker itself errors', async () => {
    const captured: { onError: ((event: ErrorEvent) => void) | null } = { onError: null };

    class FakeWorker {
      onmessage: ((event: MessageEvent) => void) | null = null;
      onerror: ((event: ErrorEvent) => void) | null = null;
      postMessage() {
        captured.onError = (event) => this.onerror?.(event);
      }
      terminate() {}
    }
    vi.stubGlobal('Worker', FakeWorker);

    const { result } = renderHook(() => useMarkdownHastTree('# Hello'));
    await waitFor(() => expect(captured.onError).not.toBeNull());

    captured.onError?.({ message: 'worker crashed' } as ErrorEvent);
    await waitFor(() => expect(result.current.status).toBe('error'));
    if (result.current.status !== 'error') throw new Error('expected error');
    expect(result.current.error).toBe('worker crashed');
  });

  it('terminates the previous worker and re-posts when markdown changes', async () => {
    const terminate = vi.fn();
    const postMessage = vi.fn();

    class FakeWorker {
      onmessage: ((event: MessageEvent) => void) | null = null;
      onerror: ((event: ErrorEvent) => void) | null = null;
      postMessage(message: unknown) {
        postMessage(message);
      }
      terminate() {
        terminate();
      }
    }
    vi.stubGlobal('Worker', FakeWorker);

    const { rerender, unmount } = renderHook(({ markdown }) => useMarkdownHastTree(markdown), {
      initialProps: { markdown: '# One' },
    });
    await waitFor(() => expect(postMessage).toHaveBeenCalledTimes(1));

    rerender({ markdown: '# Two' });
    await waitFor(() => expect(postMessage).toHaveBeenCalledTimes(2));
    expect(terminate).toHaveBeenCalledTimes(1);
    expect((postMessage.mock.calls[1]![0] as MarkdownWorkerRequest).markdown).toBe('# Two');

    unmount();
    expect(terminate).toHaveBeenCalledTimes(2);
  });
});
