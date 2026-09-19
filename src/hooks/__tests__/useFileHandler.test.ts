/**
 * useFileHandler — unit tests (W1.5, extended for P1.1/P2.4/P2.10/P2.12)
 *
 * All tests mock window.electronAPI before rendering the hook so no real IPC
 * calls are made. localStorage is available via jsdom.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { renderHook, act, waitFor } from '@testing-library/react';
import { useFileHandler, type UseFileHandlerOptions } from '../useFileHandler';
import { createMockElectronAPI } from '../../__tests__/mocks/electronAPI';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makePdfBuffer(): ArrayBuffer {
  // Starts with %PDF magic bytes: 0x25 0x50 0x44 0x46
  const buf = new ArrayBuffer(8);
  const view = new Uint8Array(buf);
  view[0] = 0x25; view[1] = 0x50; view[2] = 0x44; view[3] = 0x46;
  return buf;
}

function makeTextBuffer(text: string): ArrayBuffer {
  return new TextEncoder().encode(text).buffer;
}

/** Every test needs `addRecent`; most don't care about the recorded calls. */
function renderFileHandler(overrides: Partial<UseFileHandlerOptions> = {}) {
  const addRecent = overrides.addRecent ?? vi.fn();
  return renderHook(() => useFileHandler({ addRecent, ...overrides }));
}

// ---------------------------------------------------------------------------
// Default mock factory — override per test as needed
//
// P4.8/QA-25 — this used to be a hand-rolled literal local to this file;
// it's now the shared `createMockElectronAPI()` fixture (this file is where
// that fixture's shape was lifted from, being the most-representative
// consumer). Kept as a thin local alias so the many call sites below didn't
// all need renaming.
// ---------------------------------------------------------------------------

const buildElectronAPI = createMockElectronAPI;

// ---------------------------------------------------------------------------
// Setup / teardown
// ---------------------------------------------------------------------------

beforeEach(() => {
  localStorage.clear();
});

afterEach(() => {
  vi.restoreAllMocks();
  // Remove electronAPI so tests don't bleed into each other
  delete (window as { electronAPI?: unknown }).electronAPI;
});

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('useFileHandler', () => {
  // 1. loadFromPath — markdown (text class)
  it('loadFromPath("foo.md") calls openFileByPath and sets text/markdown file', async () => {
    window.electronAPI = buildElectronAPI({
      openFileByPath: vi.fn().mockResolvedValue({
        content: '# Hello',
        name: 'foo.md',
        path: '/abs/foo.md',
      }),
    });

    const { result } = renderFileHandler();

    await act(async () => {
      await result.current.loadFromPath('/abs/foo.md');
    });

    expect(window.electronAPI!.openFileByPath).toHaveBeenCalledWith('/abs/foo.md');
    expect(result.current.file).toMatchObject({
      kind: 'text',
      format: 'markdown',
      content: '# Hello',
      path: '/abs/foo.md',
    });
    expect(result.current.error).toBeNull();
    expect(result.current.loading).toBe(false);
  });

  // 2. loadFromPath — PDF (binary class)
  it('loadFromPath("doc.pdf") calls readBinaryByPath and sets binary/pdf file', async () => {
    const buf = makePdfBuffer();
    window.electronAPI = buildElectronAPI({
      readBinaryByPath: vi.fn().mockResolvedValue({ path: '/abs/doc.pdf', buffer: buf }),
    });

    const { result } = renderFileHandler();

    await act(async () => {
      await result.current.loadFromPath('/abs/doc.pdf');
    });

    expect(window.electronAPI!.readBinaryByPath).toHaveBeenCalledWith('/abs/doc.pdf');
    expect(result.current.file).toMatchObject({
      kind: 'binary',
      format: 'pdf',
      content: buf,
      path: '/abs/doc.pdf',
    });
  });

  // 3. loadFromPath — unknown extension with %PDF magic → resolves to pdf/binary
  it('loadFromPath("mystery.xyz") with PDF magic bytes resolves format to pdf', async () => {
    const buf = makePdfBuffer();
    window.electronAPI = buildElectronAPI({
      readBinaryByPath: vi.fn().mockResolvedValue({ path: '/abs/mystery.xyz', buffer: buf }),
    });

    const { result } = renderFileHandler();

    await act(async () => {
      await result.current.loadFromPath('/abs/mystery.xyz');
    });

    expect(result.current.file).toMatchObject({
      kind: 'binary',
      format: 'pdf',
    });
  });

  // 4. loadFromPath — notes.txt → text class, format:'text'
  it('loadFromPath("notes.txt") sets kind:text, format:text', async () => {
    window.electronAPI = buildElectronAPI({
      openFileByPath: vi.fn().mockResolvedValue({
        content: 'plain text',
        name: 'notes.txt',
        path: '/abs/notes.txt',
      }),
    });

    const { result } = renderFileHandler();

    await act(async () => {
      await result.current.loadFromPath('/abs/notes.txt');
    });

    expect(result.current.file).toMatchObject({
      kind: 'text',
      format: 'text',
      content: 'plain text',
    });
  });

  // 5. openDialog — canceled → no state change
  it('openDialog() when canceled leaves file as null', async () => {
    window.electronAPI = buildElectronAPI({
      openFileBinary: vi.fn().mockResolvedValue({ canceled: true, path: '', buffer: new ArrayBuffer(0) }),
    });

    const { result } = renderFileHandler();

    await act(async () => {
      await result.current.openDialog();
    });

    expect(result.current.file).toBeNull();
  });

  // 6. openDialog — success → routes through loadFromPath, file set
  it('openDialog() success routes through loadFromPath and sets file', async () => {
    const openFileByPath = vi.fn();
    window.electronAPI = buildElectronAPI({
      openFileBinary: vi.fn().mockResolvedValue({
        canceled: false,
        path: '/abs/report.md',
        buffer: makeTextBuffer('# Report'),
      }),
      openFileByPath,
    });

    const { result } = renderFileHandler();

    await act(async () => {
      await result.current.openDialog();
    });

    expect(result.current.file).toMatchObject({
      kind: 'text',
      format: 'markdown',
      content: '# Report',
    });
  });

  // ---------------------------------------------------------------------------
  // P4.10/LOAD-13 — openDialog's prefetched buffer skips the second read
  // ---------------------------------------------------------------------------

  it('P4.10/LOAD-13: openDialog() does not re-read the file — no second IPC call for a text-class format', async () => {
    const openFileByPath = vi.fn();
    window.electronAPI = buildElectronAPI({
      openFileBinary: vi.fn().mockResolvedValue({
        canceled: false,
        path: '/abs/report.md',
        buffer: makeTextBuffer('# Report'),
      }),
      openFileByPath,
    });

    const { result } = renderFileHandler();

    await act(async () => {
      await result.current.openDialog();
    });

    expect(openFileByPath).not.toHaveBeenCalled();
    expect(result.current.file).toMatchObject({ kind: 'text', format: 'markdown', content: '# Report' });
  });

  it('P4.10/LOAD-13: openDialog() does not re-read the file — no second IPC call for a binary-class format', async () => {
    const readBinaryByPath = vi.fn();
    const buf = makePdfBuffer();
    window.electronAPI = buildElectronAPI({
      openFileBinary: vi.fn().mockResolvedValue({ canceled: false, path: '/abs/doc.pdf', buffer: buf }),
      readBinaryByPath,
    });

    const { result } = renderFileHandler();

    await act(async () => {
      await result.current.openDialog();
    });

    expect(readBinaryByPath).not.toHaveBeenCalled();
    expect(result.current.file).toMatchObject({ kind: 'binary', format: 'pdf', content: buf });
  });

  it('P4.10/LOAD-13: loadFromPath still reads via IPC when no buffer is prefetched (drag-drop/recent/OS-open paths)', async () => {
    const openFileByPath = vi.fn().mockResolvedValue({ content: '# Hi', name: 'a.md', path: '/abs/a.md' });
    window.electronAPI = buildElectronAPI({ openFileByPath });

    const { result } = renderFileHandler();

    await act(async () => {
      await result.current.loadFromPath('/abs/a.md');
    });

    expect(openFileByPath).toHaveBeenCalledWith('/abs/a.md');
    expect(result.current.file).toMatchObject({ kind: 'text', format: 'markdown', content: '# Hi' });
  });

  // ---------------------------------------------------------------------------
  // P2.14/LOAD-04 — magic-byte verification on the fast path
  // ---------------------------------------------------------------------------

  describe('P2.14/LOAD-04: magic-byte verification on the recognized-extension fast path', () => {
    afterEach(() => {
      vi.restoreAllMocks();
    });

    it('opens normally when the magic bytes agree with the extension', async () => {
      const confirmSpy = vi.spyOn(window, 'confirm');
      window.electronAPI = buildElectronAPI({
        readBinaryByPath: vi.fn().mockResolvedValue({ path: '/abs/doc.pdf', buffer: makePdfBuffer() }),
      });

      const { result } = renderFileHandler();
      await act(async () => {
        await result.current.loadFromPath('/abs/doc.pdf');
      });

      expect(confirmSpy).not.toHaveBeenCalled();
      expect(result.current.file).toMatchObject({ kind: 'binary', format: 'pdf' });
    });

    it('prompts and proceeds when the user confirms a mismatched file', async () => {
      // .docx extension, but the bytes are actually a PDF.
      const confirmSpy = vi.spyOn(window, 'confirm').mockReturnValue(true);
      window.electronAPI = buildElectronAPI({
        readBinaryByPath: vi.fn().mockResolvedValue({ path: '/abs/fake.docx', buffer: makePdfBuffer() }),
      });

      const { result } = renderFileHandler();
      await act(async () => {
        await result.current.loadFromPath('/abs/fake.docx');
      });

      expect(confirmSpy).toHaveBeenCalledTimes(1);
      expect(confirmSpy.mock.calls[0]?.[0]).toMatch(/doesn't look like a valid/i);
      expect(result.current.file).toMatchObject({ kind: 'binary', format: 'docx' });
    });

    it('aborts silently when the user declines a mismatched file', async () => {
      const confirmSpy = vi.spyOn(window, 'confirm').mockReturnValue(false);
      window.electronAPI = buildElectronAPI({
        readBinaryByPath: vi.fn().mockResolvedValue({ path: '/abs/fake.docx', buffer: makePdfBuffer() }),
      });

      const { result } = renderFileHandler();
      await act(async () => {
        await result.current.loadFromPath('/abs/fake.docx');
      });

      expect(confirmSpy).toHaveBeenCalledTimes(1);
      expect(result.current.file).toBeNull();
      expect(result.current.error).toBeNull();
      expect(result.current.loading).toBe(false);
    });

    it('still prompts for a ZIP-family mismatch where magic is confidently a different supported binary format', async () => {
      // .pdf extension, but the bytes are actually a ZIP-based xlsx — still
      // a real mismatch worth confirming.
      const confirmSpy = vi.spyOn(window, 'confirm').mockReturnValue(true);
      const zipXlsxBuffer = (() => {
        const marker = new TextEncoder().encode('xl/workbook.xml');
        const bytes = new Uint8Array(4 + marker.length);
        bytes.set([0x50, 0x4b, 0x03, 0x04], 0);
        bytes.set(marker, 4);
        return bytes.buffer;
      })();
      window.electronAPI = buildElectronAPI({
        readBinaryByPath: vi.fn().mockResolvedValue({ path: '/abs/report.pdf', buffer: zipXlsxBuffer }),
      });

      const { result } = renderFileHandler();
      await act(async () => {
        await result.current.loadFromPath('/abs/report.pdf');
      });

      expect(confirmSpy).toHaveBeenCalledTimes(1);
      expect(result.current.file).toMatchObject({ kind: 'binary', format: 'pdf' });
    });
  });

  // 7. onFileOpenedPath boot subscription → simulating event triggers loadFromPath
  it('onFileOpenedPath subscription triggers loadFromPath on event', async () => {
    let capturedCallback: ((path: string) => void) | null = null;

    window.electronAPI = buildElectronAPI({
      onFileOpenedPath: vi.fn().mockImplementation((cb: (path: string) => void) => {
        capturedCallback = cb;
        return () => {};
      }),
      openFileByPath: vi.fn().mockResolvedValue({
        content: '# Via event',
        name: 'event.md',
        path: '/abs/event.md',
      }),
    });

    const { result } = renderFileHandler();

    // Simulate the OS event firing
    await act(async () => {
      capturedCallback?.('/abs/event.md');
      // Allow the async loadFromPath to settle
      await new Promise(r => setTimeout(r, 0));
    });

    await waitFor(() => {
      expect(result.current.file).toMatchObject({
        kind: 'text',
        format: 'markdown',
        content: '# Via event',
      });
    });
  });

  // 8. getInitialFile returning {path} at boot → file populated
  it('getInitialFile returning {path} at boot populates file', async () => {
    window.electronAPI = buildElectronAPI({
      getInitialFile: vi.fn().mockResolvedValue({ path: '/abs/init.md' }),
      openFileByPath: vi.fn().mockResolvedValue({
        content: '# Init',
        name: 'init.md',
        path: '/abs/init.md',
      }),
    });

    const { result } = renderFileHandler();

    await waitFor(() => {
      expect(result.current.file).toMatchObject({
        kind: 'text',
        format: 'markdown',
        content: '# Init',
      });
    });
  });

  // 9. Error in readBinaryByPath → error populated, file stays null, loading false
  it('error in readBinaryByPath populates error and resets loading', async () => {
    window.electronAPI = buildElectronAPI({
      readBinaryByPath: vi.fn().mockRejectedValue(new Error('IPC failure')),
    });

    const { result } = renderFileHandler();

    await act(async () => {
      await result.current.loadFromPath('/abs/doc.pdf');
    });

    expect(result.current.file).toBeNull();
    expect(result.current.error).toBe('IPC failure');
    expect(result.current.loading).toBe(false);
  });

  // ---------------------------------------------------------------------------
  // P2.4 — addRecent lifted in as a parameter (single useRecentFiles instance)
  // ---------------------------------------------------------------------------

  it('P2.4: calls the injected addRecent with the loaded file, not an internal instance', async () => {
    window.electronAPI = buildElectronAPI({
      openFileByPath: vi.fn().mockResolvedValue({
        content: '# Hello',
        name: 'foo.md',
        path: '/abs/foo.md',
      }),
    });

    const addRecent = vi.fn();
    const { result } = renderFileHandler({ addRecent });

    await act(async () => {
      await result.current.loadFromPath('/abs/foo.md');
    });

    expect(addRecent).toHaveBeenCalledWith({ path: '/abs/foo.md', name: 'foo.md' });
  });

  // ---------------------------------------------------------------------------
  // P2.3 — clearError()
  // ---------------------------------------------------------------------------

  it('P2.3: clearError() dismisses the error without touching the loaded file', async () => {
    window.electronAPI = buildElectronAPI({
      openFileByPath: vi.fn().mockResolvedValue({ content: '# A', name: 'a.md', path: '/abs/a.md' }),
      readBinaryByPath: vi.fn().mockRejectedValue(new Error('boom')),
    });

    const { result } = renderFileHandler();

    await act(async () => {
      await result.current.loadFromPath('/abs/a.md');
    });
    await act(async () => {
      await result.current.loadFromPath('/abs/b.pdf');
    });

    expect(result.current.error).toBe('boom');
    // The previously-loaded file must not have been discarded by the failed load.
    expect(result.current.file).toMatchObject({ path: '/abs/a.md' });

    act(() => {
      result.current.clearError();
    });

    expect(result.current.error).toBeNull();
    expect(result.current.file).toMatchObject({ path: '/abs/a.md' });
  });

  // ---------------------------------------------------------------------------
  // P1.1 — confirmDiscardChanges guard
  // ---------------------------------------------------------------------------

  it('P1.1: loadFromPath aborts with no state change when confirmDiscardChanges resolves false', async () => {
    const readBinaryByPath = vi.fn();
    window.electronAPI = buildElectronAPI({ readBinaryByPath });

    const confirmDiscardChanges = vi.fn().mockResolvedValue(false);
    const { result } = renderFileHandler({ confirmDiscardChanges });

    await act(async () => {
      await result.current.loadFromPath('/abs/doc.pdf');
    });

    expect(confirmDiscardChanges).toHaveBeenCalledTimes(1);
    expect(readBinaryByPath).not.toHaveBeenCalled();
    expect(result.current.file).toBeNull();
    expect(result.current.loading).toBe(false);
  });

  it('P1.1: loadFromPath proceeds when confirmDiscardChanges resolves true', async () => {
    window.electronAPI = buildElectronAPI({
      openFileByPath: vi.fn().mockResolvedValue({ content: '# Ok', name: 'ok.md', path: '/abs/ok.md' }),
    });

    const confirmDiscardChanges = vi.fn().mockResolvedValue(true);
    const { result } = renderFileHandler({ confirmDiscardChanges });

    await act(async () => {
      await result.current.loadFromPath('/abs/ok.md');
    });

    expect(confirmDiscardChanges).toHaveBeenCalledTimes(1);
    expect(result.current.file).toMatchObject({ path: '/abs/ok.md' });
  });

  it('P1.1: openDialog also runs through the confirmDiscardChanges guard (via loadFromPath)', async () => {
    window.electronAPI = buildElectronAPI({
      openFileBinary: vi.fn().mockResolvedValue({
        canceled: false,
        path: '/abs/report.md',
        buffer: makeTextBuffer('# Report'),
      }),
      openFileByPath: vi.fn().mockResolvedValue({ content: '# Report', name: 'report.md', path: '/abs/report.md' }),
    });

    const confirmDiscardChanges = vi.fn().mockResolvedValue(false);
    const { result } = renderFileHandler({ confirmDiscardChanges });

    await act(async () => {
      await result.current.openDialog();
    });

    expect(confirmDiscardChanges).toHaveBeenCalledTimes(1);
    expect(result.current.file).toBeNull();
  });

  // ---------------------------------------------------------------------------
  // P2.10 — request-token race: a slower, earlier load can't overwrite a
  // faster, newer one.
  // ---------------------------------------------------------------------------

  it('P2.10: a slower earlier load does not overwrite a faster later one', async () => {
    let resolveSlow: ((value: { path: string; buffer: ArrayBuffer }) => void) | null = null;
    const readBinaryByPath = vi.fn().mockImplementation((path: string) => {
      if (path === '/abs/slow.pdf') {
        return new Promise((resolve) => {
          resolveSlow = resolve;
        });
      }
      return Promise.resolve({ path, buffer: makePdfBuffer() });
    });

    window.electronAPI = buildElectronAPI({ readBinaryByPath });

    const { result } = renderFileHandler();

    let slowLoad!: Promise<void>;
    act(() => {
      slowLoad = result.current.loadFromPath('/abs/slow.pdf');
    });

    // Start (and finish) a second, newer load before the first resolves.
    await act(async () => {
      await result.current.loadFromPath('/abs/fast.pdf');
    });

    expect(result.current.file).toMatchObject({ path: '/abs/fast.pdf' });

    // Now let the slow (stale) load resolve — it must not clobber the fast one.
    await act(async () => {
      resolveSlow?.({ path: '/abs/slow.pdf', buffer: makePdfBuffer() });
      await slowLoad;
    });

    expect(result.current.file).toMatchObject({ path: '/abs/fast.pdf' });
  });

  it('P2.10: loadGeneration increments on every successful load, even reopening the same path', async () => {
    window.electronAPI = buildElectronAPI({
      openFileByPath: vi.fn().mockResolvedValue({ content: '# A', name: 'a.md', path: '/abs/a.md' }),
    });

    const { result } = renderFileHandler();

    expect(result.current.loadGeneration).toBe(0);

    await act(async () => {
      await result.current.loadFromPath('/abs/a.md');
    });
    expect(result.current.loadGeneration).toBe(1);

    await act(async () => {
      await result.current.loadFromPath('/abs/a.md');
    });
    expect(result.current.loadGeneration).toBe(2);
  });

  // ---------------------------------------------------------------------------
  // P2.12/SHELL-20/ELEC-19/QA-26 — browser-mode guard (no window.electronAPI)
  // ---------------------------------------------------------------------------

  it('P2.12: loadFromPath surfaces a friendly error instead of throwing when electronAPI is missing', async () => {
    delete (window as { electronAPI?: unknown }).electronAPI;

    const { result } = renderFileHandler();

    await act(async () => {
      await expect(result.current.loadFromPath('/abs/anything.md')).resolves.toBeUndefined();
    });

    expect(result.current.error).toMatch(/desktop app/i);
    expect(result.current.file).toBeNull();
    expect(result.current.loading).toBe(false);
  });

  it('P2.12: openDialog surfaces a friendly error instead of throwing when electronAPI is missing', async () => {
    delete (window as { electronAPI?: unknown }).electronAPI;

    const { result } = renderFileHandler();

    await act(async () => {
      await expect(result.current.openDialog()).resolves.toBeUndefined();
    });

    expect(result.current.error).toMatch(/desktop app/i);
  });

  // UX — a markdown/text file whose path was valid when the load started
  // (e.g. a "Recent" entry, or reopening via the OS) but no longer exists on
  // disk — deleted, renamed, or moved out from under Atlas. `openFileByPath`
  // resolves to `null` for this case (electron/main.cjs's `open-file-by-path`
  // handler) rather than rejecting; previously this surfaced as the raw,
  // path-leaking `Failed to read file: /abs/gone.md` instead of a message a
  // non-technical user could act on.
  it('loadFromPath surfaces a friendly "not found" error, without leaking the raw path, when the file no longer exists', async () => {
    window.electronAPI = buildElectronAPI({
      openFileByPath: vi.fn().mockResolvedValue(null),
    });

    const { result } = renderFileHandler();

    await act(async () => {
      await result.current.loadFromPath('/abs/gone.md');
    });

    expect(result.current.error).toMatch(/could not be found/i);
    expect(result.current.error).not.toContain('/abs/gone.md');
    expect(result.current.file).toBeNull();
    expect(result.current.loading).toBe(false);
  });

  // ---------------------------------------------------------------------------
  // P2.11/LOAD-10 — extensionless text sniff
  // ---------------------------------------------------------------------------

  describe('P2.11/LOAD-10: extensionless plain-text sniff', () => {
    it('routes an extensionless plain-text file (e.g. README) to the text viewer', async () => {
      const buffer = makeTextBuffer('This is a README with no extension.\n');
      window.electronAPI = buildElectronAPI({
        readBinaryByPath: vi.fn().mockResolvedValue({ path: '/abs/README', buffer }),
      });

      const { result } = renderFileHandler();
      await act(async () => {
        await result.current.loadFromPath('/abs/README');
      });

      expect(result.current.file).toMatchObject({
        kind: 'text',
        format: 'text',
        content: 'This is a README with no extension.\n',
      });
    });

    it('keeps a genuinely unrecognizable extensionless binary as unknown/binary', async () => {
      const bytes = new Uint8Array([0x00, 0x01, 0x02, 0x03, 0xff, 0xfe]);
      window.electronAPI = buildElectronAPI({
        readBinaryByPath: vi.fn().mockResolvedValue({ path: '/abs/mystery', buffer: bytes.buffer }),
      });

      const { result } = renderFileHandler();
      await act(async () => {
        await result.current.loadFromPath('/abs/mystery');
      });

      expect(result.current.file).toMatchObject({ kind: 'binary', format: 'unknown' });
    });
  });
});
