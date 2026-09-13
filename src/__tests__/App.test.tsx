/**
 * App — integration tests for the P1.1/P2.3/P2.4/P2.10 shell wiring.
 *
 * These exercise App.tsx's own wiring (the unsaved-changes guard, the
 * combined dirty signal, the single lifted useRecentFiles instance, the
 * file-identity-keyed dirty reset, and the loading/error banner) rather than
 * re-testing what's already covered at the hook/context/viewer level in
 * useFileHandler.test.ts, ViewerContext.test.tsx, and DocxViewer.editor.test.tsx.
 *
 * Markdown content used here is deliberately trivial (no code fences/math) so
 * MarkdownRenderer's brief preview-mode mount doesn't exercise the heavier
 * Shiki/KaTeX/Mermaid paths; the loading/error-banner tests use a `.pdf` path
 * so they never touch the markdown render path at all.
 */

import { fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import App from '../App';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function buildElectronAPI(overrides: Partial<typeof window.electronAPI> = {}): typeof window.electronAPI {
  return {
    getInitialFile: vi.fn().mockResolvedValue(null),
    openFileDialog: vi.fn().mockResolvedValue(null),
    openFileByPath: vi.fn().mockResolvedValue({ content: '', name: 'file.md', path: '' }),
    saveFile: vi.fn().mockResolvedValue({ saved: false }),
    onFileOpened: vi.fn().mockReturnValue(() => {}),
    setTheme: vi.fn(),
    openFileBinary: vi.fn().mockResolvedValue({ canceled: true, path: '', buffer: new ArrayBuffer(0) }),
    readBinaryByPath: vi.fn().mockResolvedValue({ path: '', buffer: new ArrayBuffer(0) }),
    onFileOpenedPath: vi.fn().mockReturnValue(() => {}),
    ...overrides,
  } as typeof window.electronAPI;
}

function mockOpenMarkdownFile(path: string, content: string) {
  window.electronAPI!.openFileBinary = vi.fn().mockResolvedValue({
    canceled: false,
    path,
    buffer: new TextEncoder().encode(content).buffer,
  });
  window.electronAPI!.openFileByPath = vi.fn().mockResolvedValue({
    content,
    name: path.split('/').pop(),
    path,
  });
}

function openViaToolbar() {
  fireEvent.click(screen.getByRole('button', { name: 'Open' }));
}

// The loaded filename and its dirty dot render in both the Toolbar and the
// StatusBar (same text/aria-label in each) — query the Toolbar's copy
// directly rather than an ambiguous screen.getByText/getByLabelText.
function toolbarFilenameText(): string {
  const el = document.querySelector('.toolbar__filename');
  return el?.firstChild?.textContent?.trim() ?? '';
}

function toolbarIsDirty(): boolean {
  return document.querySelector('.toolbar__dirty') !== null;
}

beforeEach(() => {
  localStorage.clear();

  // jsdom doesn't implement matchMedia — useTheme() reads it on mount.
  Object.defineProperty(window, 'matchMedia', {
    writable: true,
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

describe('App — unsaved-changes guard (P1.1/SHELL-03/04/LOAD-01/RUN-01)', () => {
  async function loadFirstFileAndMakeItDirty() {
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    // Switch to Editor view so RawEditor (a plain textarea) is mounted, and
    // edit it — this is the real user-facing path to a dirty markdown doc.
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# A (edited)' },
    });

    await waitFor(() => {
      expect(toolbarIsDirty()).toBe(true);
    });
  }

  it('Cancel leaves the dialog closed with the edit intact and does not open the new file', async () => {
    render(<App />);
    await loadFirstFileAndMakeItDirty();

    mockOpenMarkdownFile('/abs/b.md', '# B');
    await openViaToolbar();

    const dialog = await screen.findByRole('alertdialog');
    expect(dialog).toHaveTextContent('Unsaved changes');

    fireEvent.click(within(dialog).getByRole('button', { name: 'Cancel' }));

    expect(screen.queryByRole('alertdialog')).not.toBeInTheDocument();
    expect(window.electronAPI!.openFileByPath).not.toHaveBeenCalledWith('/abs/b.md');
    // The edit — and the dirty dot — survive the cancelled open.
    expect(toolbarIsDirty()).toBe(true);
    expect(screen.getByPlaceholderText('Type or paste markdown here...')).toHaveValue('# A (edited)');
  });

  it('Discard proceeds with the open and drops the unsaved edit', async () => {
    render(<App />);
    await loadFirstFileAndMakeItDirty();

    mockOpenMarkdownFile('/abs/b.md', '# B');
    await openViaToolbar();

    const dialog = await screen.findByRole('alertdialog');
    fireEvent.click(within(dialog).getByRole('button', { name: 'Discard' }));

    await waitFor(() => {
      expect(toolbarFilenameText()).toBe('b.md');
    });
    expect(window.electronAPI!.openFileByPath).toHaveBeenCalledWith('/abs/b.md');
    expect(toolbarIsDirty()).toBe(false);
  });

  it('Save saves the active document first and only then proceeds with the open', async () => {
    window.electronAPI!.saveFile = vi.fn().mockResolvedValue({ saved: true });

    render(<App />);
    await loadFirstFileAndMakeItDirty();

    mockOpenMarkdownFile('/abs/b.md', '# B');
    await openViaToolbar();

    const dialog = await screen.findByRole('alertdialog');
    fireEvent.click(within(dialog).getByRole('button', { name: 'Save' }));

    await waitFor(() => {
      expect(window.electronAPI!.saveFile).toHaveBeenCalledTimes(1);
    });
    await waitFor(() => {
      expect(window.electronAPI!.openFileByPath).toHaveBeenCalledWith('/abs/b.md');
    });
    expect(screen.queryByRole('alertdialog')).not.toBeInTheDocument();
  });

  it('Save shows an inline error and keeps the dialog open when the save fails', async () => {
    window.electronAPI!.saveFile = vi.fn().mockResolvedValue({ saved: false });

    render(<App />);
    await loadFirstFileAndMakeItDirty();

    mockOpenMarkdownFile('/abs/b.md', '# B');
    await openViaToolbar();

    const dialog = await screen.findByRole('alertdialog');
    fireEvent.click(within(dialog).getByRole('button', { name: 'Save' }));

    await waitFor(() => {
      expect(within(dialog).getByRole('alert')).toHaveTextContent(/save failed/i);
    });
    expect(window.electronAPI!.openFileByPath).not.toHaveBeenCalledWith('/abs/b.md');
  });
});

describe('App — P2.10: dirty reset keyed on file identity, not content equality (SHELL-05/LOAD-07)', () => {
  it('opening a second file with byte-identical content still clears the dirty flag and loads the new content', async () => {
    render(<App />);

    mockOpenMarkdownFile('/abs/a.md', '# Same');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    const textarea = screen.getByPlaceholderText('Type or paste markdown here...');
    fireEvent.change(textarea, { target: { value: '# Same (edited)' } });
    await waitFor(() => expect(toolbarIsDirty()).toBe(true));

    // b.md's on-disk content is identical to a.md's ORIGINAL (pre-edit)
    // content — the old string-equality guard compared the incoming
    // markdown against that stale snapshot and silently failed to reset.
    mockOpenMarkdownFile('/abs/b.md', '# Same');
    await openViaToolbar();
    const dialog = await screen.findByRole('alertdialog');
    fireEvent.click(within(dialog).getByRole('button', { name: 'Discard' }));

    await waitFor(() => expect(toolbarFilenameText()).toBe('b.md'));
    expect(toolbarIsDirty()).toBe(false);
    expect(textarea).toHaveValue('# Same');
  });
});

describe('App — P2.4: single useRecentFiles instance (SHELL-06/LOAD-16/RUN-16)', () => {
  it('removing a recent entry is not resurrected by a load-triggered addRecent call', async () => {
    localStorage.setItem(
      'atlas-recent',
      JSON.stringify([
        { path: '/abs/keep.md', name: 'keep.md', openedAt: Date.now() - 1000 },
        { path: '/abs/remove-me.md', name: 'remove-me.md', openedAt: Date.now() - 500 },
      ]),
    );

    render(<App />);

    fireEvent.click(screen.getByRole('button', { name: /remove remove-me\.md from recent/i }));

    mockOpenMarkdownFile('/abs/new.md', '# New');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('new.md'));

    const persisted: Array<{ path: string }> = JSON.parse(localStorage.getItem('atlas-recent') ?? '[]');
    const paths = persisted.map((entry) => entry.path);
    expect(paths).toContain('/abs/keep.md');
    expect(paths).toContain('/abs/new.md');
    expect(paths).not.toContain('/abs/remove-me.md');
  });
});

describe('App — P2.3: loading/error state surfaced app-wide (SHELL-07/LOAD-06/RUN-17)', () => {
  it('shows a loading indicator while a file load is in flight', async () => {
    window.electronAPI!.openFileBinary = vi.fn().mockResolvedValue({
      canceled: false,
      path: '/abs/big.pdf',
      buffer: new ArrayBuffer(0),
    });
    window.electronAPI!.readBinaryByPath = vi.fn().mockImplementation(() => new Promise(() => {}));

    render(<App />);
    await openViaToolbar();

    await waitFor(() => {
      expect(screen.getByRole('status')).toHaveTextContent(/opening file/i);
    });
  });

  it('shows a dismissible error banner when a load fails, without touching the currently-open file', async () => {
    window.electronAPI!.openFileBinary = vi.fn().mockResolvedValue({
      canceled: false,
      path: '/abs/locked.pdf',
      buffer: new ArrayBuffer(0),
    });
    window.electronAPI!.readBinaryByPath = vi.fn().mockRejectedValue(new Error('EBUSY: file is locked'));

    render(<App />);
    await openViaToolbar();

    await waitFor(() => {
      expect(screen.getByRole('alert')).toHaveTextContent(/EBUSY/i);
    });

    fireEvent.click(screen.getByRole('button', { name: 'Dismiss error' }));

    expect(screen.queryByRole('alert')).not.toBeInTheDocument();
    // Nothing was ever successfully loaded, so the Welcome screen is still up.
    expect(screen.getByRole('button', { name: 'Load Sample Document' })).toBeInTheDocument();
  });
});
