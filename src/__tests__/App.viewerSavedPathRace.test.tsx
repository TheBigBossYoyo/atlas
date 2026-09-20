/**
 * App — SAVE-1 regression: a non-markdown Save As (DOCX/slides/spreadsheet)
 * that resolves after the user has already switched to a different tab.
 *
 * `handleViewerSavedPath` (App.tsx) used to trust the ambient `file`/`filePath`
 * closure to mean "the document that was just saved" — true only when nothing
 * raced it. Since `reportSavedPath` fires from an `await` a viewer started
 * possibly long before, the user is free to switch tabs first; the callback
 * would then rename **the tab now showing** to the saved document's new path
 * and re-read that path over whatever was on screen, discarding the user's
 * work and pointing later saves at the wrong file.
 *
 * This uses a fake viewer (mirroring App.shellSession.test.tsx's own
 * `formats/registry` mock) that registers a `saveAs()` which never resolves
 * on its own — the test resolves it explicitly, strictly after the tab
 * switch has already landed, so the race is deterministic instead of relying
 * on timing.
 */
import { useEffect } from 'react';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, fireEvent, render, screen, waitFor } from '@testing-library/react';
import type { LoadedFile } from '../formats/types';
import {
  useRegisterViewerSave,
  useRegisterViewerSaveAs,
  useReportSavedPath,
} from '../viewers/shared/useViewerContext';

// Set by the fake viewer's saveAs() the moment it's invoked; resolved
// explicitly by each test, after whatever tab-switching it wants to do.
let pendingSaveAs: { readonly startedFromPath: string; readonly resolve: (newPath: string) => void } | null = null;

function FakeSaveableViewer({ file }: { file: LoadedFile }) {
  const registerSave = useRegisterViewerSave();
  const registerSaveAs = useRegisterViewerSaveAs();
  const reportSavedPath = useReportSavedPath();

  useEffect(() => {
    const saveAs = () =>
      new Promise<boolean>((resolve) => {
        pendingSaveAs = {
          startedFromPath: file.path,
          resolve: (newPath: string) => {
            // Mirrors DocxViewer/useSlideEditorCore/useSpreadsheetEditor's own
            // call site: `startedFromPath` (this document's identity when the
            // save began) alongside the path it landed on.
            reportSavedPath(file.path, newPath);
            resolve(true);
          },
        };
      });
    registerSave(async () => true);
    registerSaveAs(saveAs);
    return () => {
      registerSave(null);
      registerSaveAs(null);
    };
  }, [file.path, registerSave, registerSaveAs, reportSavedPath]);

  return <div data-testid="fake-viewer">{file.path}</div>;
}

vi.mock('../formats/registry', () => ({
  viewerRegistry: new Proxy(
    {} as Record<string, () => Promise<unknown>>,
    {
      get: () => () => Promise.resolve(FakeSaveableViewer),
    },
  ),
}));

import App from '../App';

// ---------------------------------------------------------------------------
// Helpers (mirrors App.shellSession.test.tsx's own — no shared factory exists
// yet, P4.8)
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

function triggerSaveAsShortcut() {
  fireEvent.keyDown(window, { key: 'S', ctrlKey: true, shiftKey: true });
}

beforeEach(() => {
  pendingSaveAs = null;
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
  pendingSaveAs = null;
});

describe('App — a non-markdown Save As that resolves after the tab it belongs to was switched away (SAVE-1)', () => {
  it('renames the SAVED document\'s own (background) tab and leaves the showing tab completely untouched', async () => {
    render(<App />);

    mockOpenBinaryFile('/abs/a.docx');
    await openViaToolbar();
    await screen.findByTestId('fake-viewer');
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.docx'));

    // Start a Save As on a.docx — it will not resolve until this test says so.
    triggerSaveAsShortcut();
    await waitFor(() => expect(pendingSaveAs).not.toBeNull());
    expect(pendingSaveAs!.startedFromPath).toBe('/abs/a.docx');

    // Switch to a second document while that write is still in flight.
    mockOpenBinaryFile('/abs/b.pptx');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('b.pptx'));
    await waitFor(() => expect(screen.getByTestId('fake-viewer')).toHaveTextContent('/abs/b.pptx'));

    const readBinaryByPathCallsBeforeResolve = (window.electronAPI!.readBinaryByPath as ReturnType<typeof vi.fn>).mock
      .calls.length;

    // The stale a.docx Save As now resolves, reporting a path elsewhere.
    await act(async () => {
      pendingSaveAs!.resolve('/abs/a-renamed.docx');
    });

    // b.pptx is still showing — no rename, no re-read, no adoptFile onto it.
    expect(toolbarFilenameText()).toBe('b.pptx');
    expect(screen.getByTestId('fake-viewer')).toHaveTextContent('/abs/b.pptx');
    expect(window.electronAPI!.readBinaryByPath).toHaveBeenCalledTimes(readBinaryByPathCallsBeforeResolve);
    expect(document.querySelectorAll('[data-testid="fake-viewer"]')).toHaveLength(1);

    // a.docx's own background tab followed its new path.
    expect(await screen.findByRole('tab', { name: 'a-renamed.docx' })).toBeInTheDocument();
    expect(screen.queryByRole('tab', { name: 'a.docx' })).not.toBeInTheDocument();
    expect(screen.getByRole('tab', { name: 'b.pptx' })).toBeInTheDocument();
  });

  it('switching to a-renamed.docx afterwards shows the saved document at its new path, not the original', async () => {
    render(<App />);

    mockOpenBinaryFile('/abs/a.docx');
    await openViaToolbar();
    await screen.findByTestId('fake-viewer');
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.docx'));

    triggerSaveAsShortcut();
    await waitFor(() => expect(pendingSaveAs).not.toBeNull());

    mockOpenBinaryFile('/abs/b.pptx');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('b.pptx'));

    await act(async () => {
      pendingSaveAs!.resolve('/abs/a-renamed.docx');
    });
    const renamedTab = await screen.findByRole('tab', { name: 'a-renamed.docx' });

    // Reactivating the saved document re-reads it from ITS new path, not the
    // original one it was opened from.
    fireEvent.click(renamedTab);
    await waitFor(() => expect(toolbarFilenameText()).toBe('a-renamed.docx'));
    await waitFor(() =>
      expect(window.electronAPI!.readBinaryByPath).toHaveBeenCalledWith('/abs/a-renamed.docx'),
    );
  });

  it('a Save As that resolves before any tab switch still follows the showing tab, exactly as before', async () => {
    render(<App />);

    mockOpenBinaryFile('/abs/a.docx');
    await openViaToolbar();
    await screen.findByTestId('fake-viewer');
    await waitFor(() => expect(toolbarFilenameText()).toBe('a.docx'));

    triggerSaveAsShortcut();
    await waitFor(() => expect(pendingSaveAs).not.toBeNull());

    await act(async () => {
      pendingSaveAs!.resolve('/abs/a-renamed.docx');
    });

    await waitFor(() => expect(toolbarFilenameText()).toBe('a-renamed.docx'));
    expect(await screen.findByRole('tab', { name: 'a-renamed.docx' })).toBeInTheDocument();
  });
});
