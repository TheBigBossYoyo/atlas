/**
 * App — P2.5/P2.6/P2.8/P2.13/X4/SHELL-27/QA-05 integration tests.
 *
 * Covers the shell-session wiring not already exercised by App.test.tsx
 * (P1.1/P2.3/P2.4/P2.10) or App.dirtyState.characterization.test.tsx (P0.4):
 * close-file + Ctrl+W, the dynamic window title, in-app search's format
 * gating (X4), draft recovery (P2.6), the close-confirmation IPC round-trip
 * (P2.5), and the shortcut dispatcher's active-viewer-wins-over-shell-global
 * precedence (P2.1) exercised through the real App tree rather than in
 * isolation (useShortcutManager.test.tsx covers the dispatcher itself).
 *
 * Like App.dirtyState.characterization.test.tsx, non-markdown viewers are
 * lazy-loaded via ViewerRouter + formats/registry, so `formats/registry` is
 * mocked to a trivial stand-in — real Text/Code/RTF/ODT viewers pull in
 * heavy, irrelevant dependencies, and this suite cares only about the shell
 * wiring around whichever viewer is active, never how that viewer renders.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import { useViewerShortcuts } from '../hooks/useShortcutManager';

// Mirrors DocxViewer's own SHELL-08/09/UX-03 registration (unit-tested
// separately in DocxViewer.editor.test.tsx / useShortcutManager.test.tsx):
// claims Ctrl+B at the active-viewer precedence tier unconditionally, so any
// test in this file that opens a non-markdown file can assert the shell's
// own Ctrl+B (sidebar toggle) never ran while a viewer is active — proving
// the dispatcher's precedence holds through the real App tree.
function FakeViewer() {
  useViewerShortcuts((event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'b') {
      event.preventDefault();
      return true;
    }
    return false;
  });
  return <div data-testid="fake-viewer">fake viewer content, findable text here</div>;
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

// ---------------------------------------------------------------------------
// Helpers (mirrors App.test.tsx's own — no shared factory exists yet, P4.8)
// ---------------------------------------------------------------------------

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

function mockOpenBinaryFile(path: string) {
  window.electronAPI!.openFileBinary = vi.fn().mockResolvedValue({
    canceled: false,
    path,
    buffer: new ArrayBuffer(8),
  });
  window.electronAPI!.readBinaryByPath = vi.fn().mockResolvedValue({
    path,
    buffer: new ArrayBuffer(8),
  });
}

function openViaToolbar() {
  fireEvent.click(screen.getByRole('button', { name: 'Open' }));
}

function toolbarFilenameText(): string {
  const el = document.querySelector('.toolbar__filename');
  return el?.firstChild?.textContent?.trim() ?? '';
}

function toolbarIsDirty(): boolean {
  return document.querySelector('.toolbar__dirty') !== null;
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

describe('App — close-file action (P2.8/SHELL-16)', () => {
  it('Ctrl+W on a clean document returns to the Welcome screen', async () => {
    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    fireEvent.keyDown(window, { key: 'w', ctrlKey: true });

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Load Sample Document' })).toBeInTheDocument();
    });
  });

  it('the toolbar close-file button is gated behind the unsaved-changes confirmation while dirty', async () => {
    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# A (edited)' },
    });

    fireEvent.click(screen.getByRole('button', { name: 'Close file' }));

    const dialog = await screen.findByRole('alertdialog');
    fireEvent.click(within(dialog).getByRole('button', { name: 'Discard' }));

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Load Sample Document' })).toBeInTheDocument();
    });
  });

  it('has no close-file button on the Welcome screen (nothing open to close)', () => {
    render(<App />);
    expect(screen.queryByRole('button', { name: 'Close file' })).not.toBeInTheDocument();
  });
});

describe('App — dynamic window title (P2.8/SHELL-18)', () => {
  it('shows plain "Atlas" with nothing open', () => {
    render(<App />);
    expect(document.title).toBe('Atlas');
  });

  it('reflects the open file name, and a dirty dot once edited', async () => {
    render(<App />);
    mockOpenMarkdownFile('/abs/report.md', '# Report');
    await openViaToolbar();

    await waitFor(() => expect(document.title).toBe('report.md — Atlas'));

    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# Report (edited)' },
    });

    await waitFor(() => expect(document.title).toBe('● report.md — Atlas'));
  });

  it('reverts to "Atlas" after closing the file', async () => {
    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(document.title).toBe('a.md — Atlas'));

    fireEvent.keyDown(window, { key: 'w', ctrlKey: true });

    await waitFor(() => expect(document.title).toBe('Atlas'));
  });
});

describe('App — in-app search format gating (P2.1/X4/DAT-15/RUN-08)', () => {
  it('Ctrl+F opens the search overlay for a text-class format', async () => {
    render(<App />);
    mockOpenBinaryFile('/abs/notes.txt');
    await openViaToolbar();

    await screen.findByTestId('fake-viewer');
    expect(screen.getByTitle('Search (Ctrl+F)')).toBeInTheDocument();

    fireEvent.keyDown(window, { key: 'f', ctrlKey: true });

    expect(await screen.findByPlaceholderText('Search in document...')).toBeInTheDocument();
  });

  it('Ctrl+F does not open (or register) search for a format with no in-app search (e.g. PDF)', async () => {
    render(<App />);
    mockOpenBinaryFile('/abs/report.pdf');
    await openViaToolbar();

    await screen.findByTestId('fake-viewer');
    expect(screen.queryByTitle('Search (Ctrl+F)')).not.toBeInTheDocument();

    fireEvent.keyDown(window, { key: 'f', ctrlKey: true });

    expect(screen.queryByPlaceholderText('Search in document...')).not.toBeInTheDocument();
  });
});

describe('App — shortcut precedence: active viewer wins over shell-global (P2.1/SHELL-08/SHELL-09/UX-03)', () => {
  it('a viewer-claimed combo never reaches the shell-global handler for the same key', async () => {
    // Ctrl+B is shell-global's sidebar toggle; the mocked FakeViewer above
    // claims it itself (mirroring DocxViewer's own registration). Confirm
    // the sidebar never reacts — proving the dispatcher's precedence holds
    // through the real App tree, not just useShortcutManager.test.tsx's
    // isolated harness.
    render(<App />);
    mockOpenBinaryFile('/abs/notes.txt');
    await openViaToolbar();
    await screen.findByTestId('fake-viewer');

    expect(document.querySelector('.sidebar')).toBeInTheDocument();

    fireEvent.keyDown(window, { key: 'b', ctrlKey: true });

    // Still open — shell-global's toggleSidebar never ran.
    expect(document.querySelector('.sidebar')).toBeInTheDocument();
  });

  it('shell-global still handles the combo once the claiming viewer unmounts (Welcome screen)', async () => {
    render(<App />);
    mockOpenBinaryFile('/abs/notes.txt');
    await openViaToolbar();
    await screen.findByTestId('fake-viewer');

    fireEvent.keyDown(window, { key: 'w', ctrlKey: true });
    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Load Sample Document' })).toBeInTheDocument();
    });

    // No content/sidebar at all on the Welcome screen — Ctrl+B is simply a
    // no-op here rather than an error, since the claiming viewer is gone.
    expect(document.querySelector('.sidebar')).not.toBeInTheDocument();
    expect(() => fireEvent.keyDown(window, { key: 'b', ctrlKey: true })).not.toThrow();
  });
});

describe('App — draft recovery (P2.6/SHELL-11/LOAD-20)', () => {
  function seedDraft(markdown: string, savedAt = Date.now()) {
    localStorage.setItem('atlas-draft', JSON.stringify({ markdown, fileName: null, savedAt }));
  }

  it('shows a restore prompt on launch when a draft exists, and restoring loads it as a dirty untitled document', async () => {
    seedDraft('# Recovered content');
    render(<App />);

    expect(await screen.findByRole('alert')).toHaveTextContent(/unsaved draft/i);

    fireEvent.click(screen.getByRole('button', { name: 'Restore' }));

    // Restoring loads it as an untitled document (no real file/path exists
    // for a draft, matching "Load Sample"'s pre-existing behavior) — the
    // status bar shows "Untitled" plus the dirty dot rather than the
    // toolbar's filename-gated one.
    await waitFor(() => {
      expect(document.querySelector('.statusbar__dirty')).not.toBeNull();
    });
    expect(screen.getByText('Untitled')).toBeInTheDocument();
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    expect(screen.getByPlaceholderText('Type or paste markdown here...')).toHaveValue('# Recovered content');
  });

  it('Discard clears the draft and does not restore anything', async () => {
    seedDraft('# Should not appear');
    render(<App />);

    fireEvent.click(await screen.findByRole('button', { name: 'Discard' }));

    expect(screen.queryByRole('alert')).not.toBeInTheDocument();
    expect(localStorage.getItem('atlas-draft')).toBeNull();
    expect(screen.getByRole('button', { name: 'Load Sample Document' })).toBeInTheDocument();
  });

  it('does not show a restore prompt when no draft exists', () => {
    render(<App />);
    expect(screen.queryByText(/unsaved draft/i)).not.toBeInTheDocument();
  });

  it('Load Sample dismisses a pending restore prompt instead of leaving a stale offer visible', async () => {
    // Regression: loadSample() used to leave `pendingDraft` (and thus the
    // banner) untouched, even though it immediately makes `isMarkdownDocument`
    // true and re-enables autosave — which would then start overwriting the
    // crashed session's own draft under the same storage key within 800ms,
    // while the banner kept advertising a "Restore" for content already
    // being evicted underneath it.
    seedDraft('# Should not appear');
    render(<App />);

    expect(await screen.findByRole('alert')).toHaveTextContent(/unsaved draft/i);

    fireEvent.click(screen.getByRole('button', { name: 'Load Sample Document' }));

    await waitFor(() => {
      expect(screen.queryByRole('alert')).not.toBeInTheDocument();
    });
  });
});

describe('App — main-process close confirmation round-trip (P2.5/SHELL-02/ELEC-06)', () => {
  it('pushes the combined dirty state to main whenever it changes', async () => {
    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    expect(window.electronAPI!.notifyDirtyState).toHaveBeenCalledWith(false);

    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# A (edited)' },
    });

    await waitFor(() => {
      expect(window.electronAPI!.notifyDirtyState).toHaveBeenCalledWith(true);
    });
  });

  it('saves and reports success when main requests a save-before-close', async () => {
    let requestSave: (() => void) | null = null;
    window.electronAPI!.onRequestSaveBeforeClose = vi.fn().mockImplementation((cb: () => void) => {
      requestSave = cb;
      return () => {};
    });
    window.electronAPI!.saveFile = vi.fn().mockResolvedValue({ saved: true });

    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    await act(async () => {
      requestSave?.();
    });

    await waitFor(() => {
      expect(window.electronAPI!.reportSaveBeforeCloseResult).toHaveBeenCalledWith({ saved: true });
    });
  });
});

describe('App — markdown save-failure surfacing (a wave1 follow-up: saveFile errors used to be silent)', () => {
  async function editSample() {
    render(<App />);
    fireEvent.click(screen.getByRole('button', { name: 'Load Sample Document' }));
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# edited' },
    });
  }

  it('shows a specific reason (e.g. the file is locked) via FileStatusBanner on a direct re-save failure', async () => {
    window.electronAPI!.saveFile = vi.fn().mockResolvedValue({
      saved: false,
      error: 'This file appears to be open in another program.',
    });

    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# A (edited)' },
    });

    fireEvent.click(screen.getByRole('button', { name: 'Save' }));

    expect(await screen.findByRole('alert')).toHaveTextContent(/open in another program/i);

    fireEvent.click(screen.getByRole('button', { name: 'Dismiss error' }));
    expect(screen.queryByRole('alert')).not.toBeInTheDocument();
  });

  it('falls back to a generic friendly message when main reports no specific reason for a direct re-save failure', async () => {
    window.electronAPI!.saveFile = vi.fn().mockResolvedValue({ saved: false });

    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# A (edited)' },
    });

    fireEvent.click(screen.getByRole('button', { name: 'Save' }));

    expect(await screen.findByRole('alert')).toHaveTextContent(/failed to save/i);
  });

  it('does not show an error banner when a Save-As-style dialog is simply cancelled (no existing path, no reason given)', async () => {
    window.electronAPI!.saveFile = vi.fn().mockResolvedValue({ saved: false });

    await editSample();

    fireEvent.click(screen.getByRole('button', { name: 'Save' }));

    // No existingPath on an untitled document — main would have shown a
    // Save dialog; a plain `{ saved: false }` here is an ordinary Cancel.
    await waitFor(() => expect(window.electronAPI!.saveFile).toHaveBeenCalled());
    expect(screen.queryByRole('alert')).not.toBeInTheDocument();
  });

  it('clears a prior save error once a subsequent save succeeds', async () => {
    window.electronAPI!.saveFile = vi
      .fn()
      .mockResolvedValueOnce({ saved: false, error: 'Disk is full.' })
      .mockResolvedValueOnce({ saved: true });

    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# A (edited)' },
    });

    fireEvent.click(screen.getByRole('button', { name: 'Save' }));
    expect(await screen.findByRole('alert')).toHaveTextContent(/disk is full/i);

    fireEvent.click(screen.getByRole('button', { name: 'Save' }));
    await waitFor(() => expect(screen.queryByRole('alert')).not.toBeInTheDocument());
  });
});

describe('App — a save still in flight while the tab it belongs to is switched away (review fix)', () => {
  // `saveFile` closes over the tab active when it *started*: the awaited
  // `window.electronAPI.saveFile(...)` can resolve well after the user has
  // switched (via Discard, through the unsaved-changes guard) to a different
  // tab. Before this fix, the continuation ran unconditionally against
  // whichever tab was showing by then — clearing that tab's dirty dot and
  // wiping the single shared autosave draft even though neither belonged to
  // the document actually written to disk.
  it('leaves the newly active tab dirty and its autosave draft intact once the old tab\'s save resolves', async () => {
    let resolveSave: ((result: { saved: boolean; path?: string }) => void) | undefined;

    render(<App />);

    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    // A second, clean document — opening it while a.md is clean needs no
    // confirmation, and leaves a.md as a background tab.
    mockOpenMarkdownFile('/abs/b.md', '# B');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('b.md'));

    // Switch back to a.md and dirty it.
    fireEvent.click(screen.getByRole('tab', { name: 'a.md' }));
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# A (edited)' },
    });
    await waitFor(() => expect(toolbarIsDirty()).toBe(true));

    // Start a save that never resolves on its own.
    window.electronAPI!.saveFile = vi.fn().mockImplementation(
      () => new Promise<{ saved: boolean; path?: string }>((resolve) => { resolveSave = resolve; }),
    );
    fireEvent.click(screen.getByRole('button', { name: 'Save' }));
    await waitFor(() => expect(window.electronAPI!.saveFile).toHaveBeenCalledTimes(1));

    // While that save is still pending, switch to b.md — a.md is still
    // (from the app's perspective) dirty, so this goes through the
    // unsaved-changes guard; Discard proceeds with the switch.
    fireEvent.click(screen.getByRole('tab', { name: 'b.md' }));
    const dialog = await screen.findByRole('alertdialog');
    fireEvent.click(within(dialog).getByRole('button', { name: 'Discard' }));
    await waitFor(() => expect(toolbarFilenameText()).toBe('b.md'));

    // Dirty b.md and let its own autosave draft actually land — the real
    // 800ms debounce, not a seeded fixture, so this proves the live hook's
    // draft (not just a hand-placed one) survives the stale continuation.
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# B (edited)' },
    });
    await waitFor(() => expect(toolbarIsDirty()).toBe(true));
    await waitFor(
      () => expect(localStorage.getItem('atlas-draft')).not.toBeNull(),
      { timeout: 2000 },
    );
    const draftBeforeResolve = localStorage.getItem('atlas-draft');
    expect(draftBeforeResolve).toContain('B (edited)');

    // The stale a.md save now finishes successfully.
    await act(async () => {
      resolveSave?.({ saved: true });
    });

    // b.md is still showing, still dirty, and its draft was never touched.
    expect(toolbarFilenameText()).toBe('b.md');
    expect(toolbarIsDirty()).toBe(true);
    expect(screen.getByPlaceholderText('Type or paste markdown here...')).toHaveValue('# B (edited)');
    expect(localStorage.getItem('atlas-draft')).toBe(draftBeforeResolve);
  });
});
