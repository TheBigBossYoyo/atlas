/**
 * NEW-01 — a 0-byte file used to either fall through to the Welcome screen
 * (markdown — `isMarkdownDocument`/`hasContent` derived from content
 * *length*, so an empty-but-loaded file looked indistinguishable from "no
 * file at all") or reach a format's parser as a genuinely empty buffer and
 * throw a low-level error (every binary format). This covers the renderer
 * side of both fixes:
 *   - a 0-byte markdown file opens as an (empty) editor bound to its path,
 *     not the Welcome screen (App.tsx's `hasContent`).
 *   - a 0-byte file of a format Atlas has no blank template for (main.cjs
 *     only substitutes one for docx/xlsx/ods/pptx/odp — see
 *     `main.ipc.newDocument.test.ts` for that main-process substitution)
 *     shows a friendly "This file is empty" notice instead of reaching the
 *     viewer at all (App.tsx's `EmptyFileNotice` branch).
 *
 * Mirrors App.shellSession.test.tsx's own harness (mocked `formats/registry`
 * — this suite cares about the shell's branching, never how a real viewer
 * renders).
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';

function FakeViewer() {
  return <div data-testid="fake-viewer">fake viewer content</div>;
}

vi.mock('../formats/registry', () => ({
  viewerRegistry: new Proxy(
    {} as Record<string, () => Promise<unknown>>,
    {
      get: () => () => Promise.resolve(FakeViewer),
    },
  ),
}));

import App from '../App';

function buildElectronAPI(overrides: Partial<typeof window.electronAPI> = {}): typeof window.electronAPI {
  return {
    getInitialFile: vi.fn().mockResolvedValue(null),
    openFileDialog: vi.fn().mockResolvedValue(null),
    openFileByPath: vi.fn().mockResolvedValue({ content: '', name: 'file.md', path: '' }),
    saveFile: vi.fn().mockResolvedValue({ saved: true }),
    onFileOpened: vi.fn().mockReturnValue(() => {}),
    setTheme: vi.fn(),
    openFileBinary: vi.fn().mockResolvedValue({ canceled: true, path: '', buffer: new ArrayBuffer(0) }),
    readBinaryByPath: vi.fn().mockResolvedValue({ path: '', buffer: new ArrayBuffer(0) }),
    onFileOpenedPath: vi.fn().mockReturnValue(() => {}),
    notifyDirtyState: vi.fn(),
    onRequestSaveBeforeClose: vi.fn().mockReturnValue(() => {}),
    reportSaveBeforeCloseResult: vi.fn(),
    ...overrides,
  } as typeof window.electronAPI;
}

function openViaToolbar() {
  fireEvent.click(screen.getByRole('button', { name: 'Open' }));
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

  window.electronAPI = buildElectronAPI();
});

afterEach(() => {
  vi.restoreAllMocks();
  delete (window as { electronAPI?: unknown }).electronAPI;
});

describe('App — 0-byte file handling (NEW-01)', () => {
  it('a 0-byte .md file opens as an empty editor, not the Welcome screen', async () => {
    render(<App />);
    window.electronAPI!.openFileBinary = vi.fn().mockResolvedValue({
      canceled: false,
      path: '/abs/empty.md',
      buffer: new ArrayBuffer(0),
    });

    await openViaToolbar();

    await waitFor(() => {
      expect(screen.queryByRole('button', { name: 'Load Sample Document' })).not.toBeInTheDocument();
    });
    // The markdown editor/preview chrome is showing (view-mode toggle only
    // renders once a markdown document is active).
    expect(screen.getByRole('button', { name: 'Editor' })).toBeInTheDocument();
  });

  it('a 0-byte .pdf file shows a friendly "this file is empty" notice, not the viewer', async () => {
    render(<App />);
    window.electronAPI!.openFileBinary = vi.fn().mockResolvedValue({
      canceled: false,
      path: '/abs/empty.pdf',
      buffer: new ArrayBuffer(0),
    });

    await openViaToolbar();

    await waitFor(() => {
      expect(screen.getByText('This file is empty')).toBeInTheDocument();
    });
    expect(screen.queryByTestId('fake-viewer')).not.toBeInTheDocument();
  });

  it('a non-empty .pdf file still reaches the real viewer (no false-positive empty notice)', async () => {
    render(<App />);
    window.electronAPI!.openFileBinary = vi.fn().mockResolvedValue({
      canceled: false,
      path: '/abs/real.pdf',
      buffer: new ArrayBuffer(8),
    });

    await openViaToolbar();

    await waitFor(() => {
      expect(screen.getByTestId('fake-viewer')).toBeInTheDocument();
    });
    expect(screen.queryByText('This file is empty')).not.toBeInTheDocument();
  });
});

describe('App — New document menu (NEW-01)', () => {
  it('clicking New, then a format, asks main to create it and opens the result', async () => {
    render(<App />);
    const newDocument = vi.fn().mockResolvedValue({ created: true, path: '/abs/Document.docx' });
    window.electronAPI!.newDocument = newDocument;
    window.electronAPI!.readBinaryByPath = vi.fn().mockResolvedValue({
      path: '/abs/Document.docx',
      buffer: new ArrayBuffer(8),
    });

    fireEvent.click(screen.getByRole('button', { name: 'New' }));
    fireEvent.click(screen.getByRole('button', { name: 'Word document (.docx)' }));

    expect(newDocument).toHaveBeenCalledWith('docx');
    await waitFor(() => {
      expect(screen.getByTestId('fake-viewer')).toBeInTheDocument();
    });
  });

  it('Ctrl+N opens the New document menu', () => {
    render(<App />);
    expect(screen.queryByRole('button', { name: 'Word document (.docx)' })).not.toBeInTheDocument();

    fireEvent.keyDown(window, { key: 'n', ctrlKey: true });

    expect(screen.getByRole('button', { name: 'Word document (.docx)' })).toBeInTheDocument();
  });

  it('a cancelled New (dialog dismissed in main) leaves the app untouched', async () => {
    render(<App />);
    const newDocument = vi.fn().mockResolvedValue({ created: false });
    window.electronAPI!.newDocument = newDocument;

    fireEvent.click(screen.getByRole('button', { name: 'New' }));
    fireEvent.click(screen.getByRole('button', { name: 'Markdown (.md)' }));

    expect(newDocument).toHaveBeenCalledWith('markdown');
    await waitFor(() => expect(newDocument).toHaveBeenCalled());
    expect(screen.getByRole('button', { name: 'Load Sample Document' })).toBeInTheDocument();
  });
});
