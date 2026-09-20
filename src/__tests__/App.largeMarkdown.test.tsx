/**
 * Markdown parsing (micromark, under react-markdown) degrades superlinearly:
 * measured on this project's corpus, 0.5 MB ~2 s, 1 MB ~9 s, 2 MB ~44 s, and
 * a 5 MB document takes minutes at a ~2.6 GB peak. `MarkdownRenderer` now
 * parses documents past `components/markdown/sizeThresholds.ts`'s (much
 * lower) `MARKDOWN_WORKER_BYTE_THRESHOLD` in a Worker, so that cost no
 * longer blocks the main thread — but it doesn't go away, and a document
 * large enough can still take minutes or exhaust a Worker's heap. Past
 * `LARGE_MARKDOWN_PREVIEW_BYTES` (this file's own threshold, well above the
 * Worker one) the preview therefore still waits to be asked for, while the
 * editor pane — unaffected either way — opens normally.
 *
 * Mirrors App.emptyFile.test.tsx's harness (mocked `formats/registry`: this
 * suite cares about the shell's branching, not how a viewer renders). Also
 * mocks `MarkdownRenderer` itself wholesale, so it never actually reaches
 * `MARKDOWN_WORKER_BYTE_THRESHOLD` or a Worker — that combination is
 * `MarkdownRenderer.worker.test.tsx`'s job.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';

import { LARGE_MARKDOWN_PREVIEW_BYTES } from '../components/LargeMarkdownNotice';

function FakeViewer() {
  return <div data-testid="fake-viewer">fake viewer content</div>;
}

vi.mock('../formats/registry', () => ({
  viewerRegistry: new Proxy({} as Record<string, () => Promise<unknown>>, {
    get: () => () => Promise.resolve(FakeViewer),
  }),
}));

vi.mock('../components/MarkdownRenderer', () => ({
  MarkdownRenderer: ({ markdown }: { markdown: string }) => (
    <div data-testid="markdown-preview">{markdown.slice(0, 20)}</div>
  ),
}));

import App from '../App';

const HUGE = `# Heading\n\n${'word '.repeat(LARGE_MARKDOWN_PREVIEW_BYTES / 5)}`;
const SMALL = '# Small document\n\nJust a paragraph.\n';

function buildElectronAPI(content: string, path: string): typeof window.electronAPI {
  return {
    getInitialFile: vi.fn().mockResolvedValue(null),
    openFileDialog: vi.fn().mockResolvedValue(null),
    openFileByPath: vi.fn().mockResolvedValue({ content, name: 'doc.md', path }),
    saveFile: vi.fn().mockResolvedValue({ saved: true }),
    onFileOpened: vi.fn().mockReturnValue(() => {}),
    setTheme: vi.fn(),
    openFileBinary: vi.fn().mockResolvedValue({
      canceled: false,
      path,
      buffer: new TextEncoder().encode(content).buffer,
    }),
    readBinaryByPath: vi.fn().mockResolvedValue({ path, buffer: new TextEncoder().encode(content).buffer }),
    onFileOpenedPath: vi.fn().mockReturnValue(() => {}),
    notifyDirtyState: vi.fn(),
    onRequestSaveBeforeClose: vi.fn().mockReturnValue(() => {}),
    reportSaveBeforeCloseResult: vi.fn(),
  } as unknown as typeof window.electronAPI;
}

beforeEach(() => {
  localStorage.clear();
  document.title = '';
  Object.defineProperty(window, 'matchMedia', {
    writable: true,
    configurable: true,
    value: vi.fn().mockImplementation((query: string) => ({
      matches: false,
      media: query,
      onchange: null,
      addListener: vi.fn(),
      removeListener: vi.fn(),
      addEventListener: vi.fn(),
      removeEventListener: vi.fn(),
      dispatchEvent: vi.fn(),
    })),
  });
  window.electronAPI = buildElectronAPI(HUGE, '/abs/huge.md');
});

afterEach(() => {
  vi.restoreAllMocks();
  delete (window as { electronAPI?: unknown }).electronAPI;
});

describe('App — very large markdown documents', () => {
  it('opens the editor but holds the preview back, and renders it on request', async () => {
    render(<App />);
    fireEvent.click(screen.getByRole('button', { name: 'Open' }));

    await waitFor(() => {
      expect(screen.getByText('This document is too large to preview quickly')).toBeInTheDocument();
    });
    // The preview never mounted, but the editor did.
    expect(screen.queryByTestId('markdown-preview')).not.toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Editor' })).toBeInTheDocument();

    fireEvent.click(screen.getByRole('button', { name: 'Render the preview anyway' }));

    await waitFor(() => {
      expect(screen.getByTestId('markdown-preview')).toBeInTheDocument();
    });
    expect(screen.queryByText('This document is too large to preview quickly')).not.toBeInTheDocument();
  });

  it('leaves an ordinary document alone', async () => {
    window.electronAPI = buildElectronAPI(SMALL, '/abs/small.md');
    render(<App />);
    fireEvent.click(screen.getByRole('button', { name: 'Open' }));

    await waitFor(() => {
      expect(screen.getByTestId('markdown-preview')).toBeInTheDocument();
    });
    expect(screen.queryByText('This document is too large to preview quickly')).not.toBeInTheDocument();
  });
});
