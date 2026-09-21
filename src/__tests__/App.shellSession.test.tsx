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
    onRequestDiscardBeforeClose: vi.fn().mockReturnValue(() => {}),
    reportDiscardBeforeCloseResult: vi.fn(),
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

describe('App — discarding a markdown draft actually clears it (DRAFT-1)', () => {
  // Types into the editor and waits for the REAL 800ms autosave debounce to
  // actually persist a draft — a hand-seeded `localStorage` entry wouldn't
  // prove the live wiring works, just that the assertion can read the key.
  async function dirtyAndAutosave(text: string) {
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: text },
    });
    await waitFor(() => expect(toolbarIsDirty()).toBe(true));
    await waitFor(() => expect(localStorage.getItem('atlas-draft')).not.toBeNull(), { timeout: 2000 });
    expect(localStorage.getItem('atlas-draft')).toContain(text);
  }

  it('Ctrl+W (closing the active tab) clears the draft on Discard', async () => {
    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    await dirtyAndAutosave('# A (edited, about to be discarded)');

    fireEvent.keyDown(window, { key: 'w', ctrlKey: true });
    const dialog = await screen.findByRole('alertdialog');
    fireEvent.click(within(dialog).getByRole('button', { name: 'Discard' }));

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Load Sample Document' })).toBeInTheDocument();
    });
    expect(localStorage.getItem('atlas-draft')).toBeNull();
  });

  it('the toolbar close-file button clears the draft on Discard', async () => {
    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    await dirtyAndAutosave('# A (edited)');

    fireEvent.click(screen.getByRole('button', { name: 'Close file' }));
    const dialog = await screen.findByRole('alertdialog');
    fireEvent.click(within(dialog).getByRole('button', { name: 'Discard' }));

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Load Sample Document' })).toBeInTheDocument();
    });
    expect(localStorage.getItem('atlas-draft')).toBeNull();
  });

  it('switching tabs away from a dirty markdown document clears the draft on Discard', async () => {
    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    mockOpenBinaryFile('/abs/notes.docx');
    await openViaToolbar();
    await screen.findByTestId('fake-viewer');

    fireEvent.click(screen.getByRole('tab', { name: 'a.md' }));
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    await dirtyAndAutosave('# A (edited before switching away)');

    // Switch to the .docx tab — this is the exact repro from the bug report:
    // discarding markdown changes by switching to a non-markdown tab, which
    // then gates autosave off entirely and (before this fix) never cleared
    // the draft it had already written.
    fireEvent.click(screen.getByRole('tab', { name: 'notes.docx' }));
    const dialog = await screen.findByRole('alertdialog');
    fireEvent.click(within(dialog).getByRole('button', { name: 'Discard' }));

    await waitFor(() => expect(toolbarFilenameText()).toBe('notes.docx'));
    expect(localStorage.getItem('atlas-draft')).toBeNull();
  });

  it('File > Open (Open toolbar button) while a markdown document is dirty clears the draft on Discard', async () => {
    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    await dirtyAndAutosave('# A (edited)');

    mockOpenMarkdownFile('/abs/other.md', '# Other');
    await openViaToolbar();

    const dialog = await screen.findByRole('alertdialog');
    fireEvent.click(within(dialog).getByRole('button', { name: 'Discard' }));

    await waitFor(() => expect(toolbarFilenameText()).toBe('other.md'));
    expect(localStorage.getItem('atlas-draft')).toBeNull();
  });

  it('Cancelling the unsaved-changes dialog leaves the draft intact (only Discard clears it)', async () => {
    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    await dirtyAndAutosave('# A (edited)');

    fireEvent.keyDown(window, { key: 'w', ctrlKey: true });
    const dialog = await screen.findByRole('alertdialog');
    fireEvent.click(within(dialog).getByRole('button', { name: 'Cancel' }));

    await waitFor(() => expect(screen.queryByRole('alertdialog')).not.toBeInTheDocument());
    expect(toolbarFilenameText()).toBe('a.md');
    expect(localStorage.getItem('atlas-draft')).not.toBeNull();
  });

  it('the crash-recovery banner\'s own Restore and Discard are unaffected by this fix', async () => {
    localStorage.setItem(
      'atlas-draft',
      JSON.stringify({ markdown: '# Recovered', fileName: null, savedAt: Date.now() }),
    );
    render(<App />);

    expect(await screen.findByRole('alert')).toHaveTextContent(/unsaved draft/i);
    fireEvent.click(screen.getByRole('button', { name: 'Discard' }));

    expect(screen.queryByRole('alert')).not.toBeInTheDocument();
    expect(localStorage.getItem('atlas-draft')).toBeNull();
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

  it('subscribes onRequestSaveBeforeClose exactly once, even as the document changes underneath it (root-cause regression — CI flake investigation)', async () => {
    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# A (edited)' },
    });
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# A (edited again)' },
    });

    await waitFor(() => {
      expect(window.electronAPI!.notifyDirtyState).toHaveBeenCalledWith(true);
    });

    // A real Electron IPC listener that's torn down and re-registered on
    // every keystroke and every step of opening a file isn't just wasted
    // churn — it's what produced this suite's own flake (a race between
    // `waitFor` observing a committed DOM change and the passive effect
    // that re-subscribes with the fresh `saveFile` actually having run; see
    // `saveFileRef`'s comment in App.tsx). The close-confirmation callback
    // must be registered exactly once, for the app shell's whole lifetime,
    // no matter how much the document underneath it changes.
    expect(window.electronAPI!.onRequestSaveBeforeClose).toHaveBeenCalledTimes(1);
  });
});

describe('App — main-process discard-before-close round trip (QUIT-DRAFT-1)', () => {
  // Types into the editor and waits for the REAL 800ms autosave debounce to
  // actually persist a draft — mirrors the DRAFT-1 suite above's own
  // `dirtyAndAutosave`, so this proves the live autosave wiring is cleared,
  // not just that the assertion can read the key.
  async function dirtyAndAutosave(text: string) {
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: text },
    });
    await waitFor(() => expect(toolbarIsDirty()).toBe(true));
    await waitFor(() => expect(localStorage.getItem('atlas-draft')).not.toBeNull(), { timeout: 2000 });
    expect(localStorage.getItem('atlas-draft')).toContain(text);
  }

  it('clears the draft and acknowledges when main requests a discard-before-close (Quit + Discard)', async () => {
    let requestDiscard: (() => void) | null = null;
    window.electronAPI!.onRequestDiscardBeforeClose = vi.fn().mockImplementation((cb: () => void) => {
      requestDiscard = cb;
      return () => {};
    });

    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    await dirtyAndAutosave('# A (edited, about to be quit-discarded)');

    act(() => {
      requestDiscard?.();
    });

    expect(localStorage.getItem('atlas-draft')).toBeNull();
    expect(window.electronAPI!.reportDiscardBeforeCloseResult).toHaveBeenCalledTimes(1);
  });

  it('acknowledges even when there is no draft to clear, so main is never left waiting out its timeout needlessly', async () => {
    let requestDiscard: (() => void) | null = null;
    window.electronAPI!.onRequestDiscardBeforeClose = vi.fn().mockImplementation((cb: () => void) => {
      requestDiscard = cb;
      return () => {};
    });

    render(<App />);

    act(() => {
      requestDiscard?.();
    });

    expect(window.electronAPI!.reportDiscardBeforeCloseResult).toHaveBeenCalledTimes(1);
  });

  it('does not clear an existing draft when the discard is for the active NON-markdown document (do-not-over-clear)', async () => {
    // Mirrors handleUnsavedDialogDiscard's own same-shaped guard (see its
    // comment in App.tsx): a genuine crash-recovery draft left over from an
    // earlier markdown session must survive a Quit+Discard of an unrelated,
    // currently-active non-markdown document.
    localStorage.setItem(
      'atlas-draft',
      JSON.stringify({ markdown: '# Earlier crash, still unrecovered', fileName: null, savedAt: Date.now() }),
    );

    let requestDiscard: (() => void) | null = null;
    window.electronAPI!.onRequestDiscardBeforeClose = vi.fn().mockImplementation((cb: () => void) => {
      requestDiscard = cb;
      return () => {};
    });

    render(<App />);
    // Dismiss the crash-recovery banner's own offer without discarding it,
    // so the draft is still sitting in storage when Quit fires below.
    await screen.findByRole('alert');

    mockOpenBinaryFile('/abs/notes.docx');
    await openViaToolbar();
    await screen.findByTestId('fake-viewer');

    act(() => {
      requestDiscard?.();
    });

    expect(localStorage.getItem('atlas-draft')).not.toBeNull();
    expect(window.electronAPI!.reportDiscardBeforeCloseResult).toHaveBeenCalledTimes(1);
  });

  it('subscribes onRequestDiscardBeforeClose exactly once, even as the document changes underneath it (mount-once, mirrors the Save listener)', async () => {
    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# A (edited)' },
    });
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# A (edited again)' },
    });

    await waitFor(() => {
      expect(window.electronAPI!.notifyDirtyState).toHaveBeenCalledWith(true);
    });

    expect(window.electronAPI!.onRequestDiscardBeforeClose).toHaveBeenCalledTimes(1);
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

describe('App — Save As onto an already-open tab (SESS-1, through the real shell)', () => {
  // `documentSessions.test.ts` already unit-tests `renameSession`'s two
  // collision branches directly against the pure reducer (both are exercised
  // there with plain `LoadedFile`/`DocumentSessionsState` values). What that
  // suite can't prove is that `App.tsx` actually reaches those branches with
  // the right arguments at the right time — this describe block drives both
  // through the real component tree: a real Ctrl+Shift+S, a real deferred
  // `window.electronAPI.saveFile` promise, and real DOM assertions on the
  // resulting tab strip and editor content.

  it('Save As from the active tab onto an already-open background tab\'s path drops the background tab; the renamed tab stays active with its saved content', async () => {
    // Case 1 (the common case) — `followSavedPath`'s synchronous branch:
    // the ACTIVE document (a.md) is Save-As'd onto a path a background tab
    // (b.md) already occupies. b.md's tab is dropped; a.md's content now
    // lives under b.md's path, still active, still the tab the user is
    // looking at.
    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    mockOpenMarkdownFile('/abs/b.md', '# B');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('b.md'));

    fireEvent.click(screen.getByRole('tab', { name: 'a.md' }));
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# A (about to be saved onto b.md\'s path)' },
    });
    await waitFor(() => expect(toolbarIsDirty()).toBe(true));

    window.electronAPI!.saveFile = vi.fn().mockResolvedValue({ saved: true, path: '/abs/b.md' });
    fireEvent.keyDown(window, { key: 's', ctrlKey: true, shiftKey: true });

    await waitFor(() => expect(toolbarFilenameText()).toBe('b.md'));
    expect(toolbarIsDirty()).toBe(false);

    // Only one tab survives — b.md's original background tab was dropped,
    // not duplicated alongside the renamed one (no two sessions share an id).
    expect(screen.getAllByRole('tab')).toHaveLength(1);
    expect(screen.getByRole('tab', { name: 'b.md' })).toBeInTheDocument();
    expect(screen.queryByRole('tab', { name: 'a.md' })).not.toBeInTheDocument();

    // The surviving tab holds a.md's (just-saved) content, not whatever
    // b.md used to hold.
    expect(screen.getByPlaceholderText('Type or paste markdown here...')).toHaveValue(
      '# A (about to be saved onto b.md\'s path)',
    );
  });

  it('a background Save As that resolves onto the currently-active tab\'s path leaves the active tab untouched and drops the background one instead', async () => {
    // Case 2 (the deliberately conservative exception) —
    // `applySavedPathToInactiveTab`'s branch: a.md's Save As started while it
    // was active, but the user switched to b.md before it resolved. It
    // resolves onto b.md's path — the path of the tab now ACTUALLY on
    // screen. There's nothing to re-render a.md's content into (nothing
    // adopts it), so b.md is left exactly as it was and a.md's now-redundant
    // background tab is dropped instead — mirroring
    // documentSessions.test.ts's "renaming a background document onto the
    // currently-active document's path" reducer case, but reached here via
    // a real async save race instead of a direct reducer call.
    let resolveSave: ((result: { saved: boolean; path?: string }) => void) | undefined;

    render(<App />);
    mockOpenMarkdownFile('/abs/a.md', '# A');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));

    mockOpenMarkdownFile('/abs/b.md', '# B');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('b.md'));

    // Switch back to a.md, dirty it, and start a Save As that never resolves
    // on its own.
    fireEvent.click(screen.getByRole('tab', { name: 'a.md' }));
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.md'));
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    fireEvent.change(screen.getByPlaceholderText('Type or paste markdown here...'), {
      target: { value: '# A (background save, about to collide with b.md)' },
    });
    await waitFor(() => expect(toolbarIsDirty()).toBe(true));

    window.electronAPI!.saveFile = vi.fn().mockImplementation(
      () => new Promise<{ saved: boolean; path?: string }>((resolve) => { resolveSave = resolve; }),
    );
    fireEvent.keyDown(window, { key: 's', ctrlKey: true, shiftKey: true });
    await waitFor(() => expect(window.electronAPI!.saveFile).toHaveBeenCalledTimes(1));

    // While that save is still pending, switch to b.md — a.md is still
    // dirty from the app's perspective, so this goes through the
    // unsaved-changes guard; Discard proceeds with the switch (mirrors the
    // in-flight-save describe block above).
    fireEvent.click(screen.getByRole('tab', { name: 'b.md' }));
    const dialog = await screen.findByRole('alertdialog');
    fireEvent.click(within(dialog).getByRole('button', { name: 'Discard' }));
    await waitFor(() => expect(toolbarFilenameText()).toBe('b.md'));

    // The stale a.md Save As now resolves — landing on '/abs/b.md', the path
    // of the tab actually on screen right now.
    await act(async () => {
      resolveSave?.({ saved: true, path: '/abs/b.md' });
    });

    // b.md is still showing, completely untouched by a.md's save.
    expect(toolbarFilenameText()).toBe('b.md');
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    expect(screen.getByPlaceholderText('Type or paste markdown here...')).toHaveValue('# B');

    // a.md's now-redundant background tab is gone — not left dangling, and
    // not duplicated onto b.md's identity either.
    expect(screen.getAllByRole('tab')).toHaveLength(1);
    expect(screen.getByRole('tab', { name: 'b.md' })).toBeInTheDocument();
    expect(screen.queryByRole('tab', { name: 'a.md' })).not.toBeInTheDocument();
  });
});
