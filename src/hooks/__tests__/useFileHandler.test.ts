/**
 * useFileHandler — unit tests (W1.5)
 *
 * All tests mock window.electronAPI before rendering the hook so no real IPC
 * calls are made. localStorage is available via jsdom.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { renderHook, act, waitFor } from '@testing-library/react';
import { useFileHandler } from '../useFileHandler';

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

// ---------------------------------------------------------------------------
// Default mock factory — override per test as needed
// ---------------------------------------------------------------------------

function buildElectronAPI(overrides: Partial<typeof window.electronAPI> = {}): typeof window.electronAPI {
  return {
    getInitialFile: vi.fn().mockResolvedValue(null),
    openFileDialog: vi.fn().mockResolvedValue(null),
    openFileByPath: vi.fn().mockResolvedValue({ content: '', name: 'file', path: '' }),
    saveFile: vi.fn().mockResolvedValue({ saved: false }),
    onFileOpened: vi.fn().mockReturnValue(() => {}),
    setTheme: vi.fn(),
    openFileBinary: vi.fn().mockResolvedValue({ canceled: true, path: '', buffer: new ArrayBuffer(0) }),
    readBinaryByPath: vi.fn().mockResolvedValue({ path: '', buffer: new ArrayBuffer(0) }),
    onFileOpenedPath: vi.fn().mockReturnValue(() => {}),
    ...overrides,
  } as typeof window.electronAPI;
}

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

    const { result } = renderHook(() => useFileHandler());

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

    const { result } = renderHook(() => useFileHandler());

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

    const { result } = renderHook(() => useFileHandler());

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

    const { result } = renderHook(() => useFileHandler());

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

    const { result } = renderHook(() => useFileHandler());

    await act(async () => {
      await result.current.openDialog();
    });

    expect(result.current.file).toBeNull();
  });

  // 6. openDialog — success → routes through loadFromPath, file set
  it('openDialog() success routes through loadFromPath and sets file', async () => {
    window.electronAPI = buildElectronAPI({
      openFileBinary: vi.fn().mockResolvedValue({
        canceled: false,
        path: '/abs/report.md',
        buffer: makeTextBuffer('# Report'),
      }),
      openFileByPath: vi.fn().mockResolvedValue({
        content: '# Report',
        name: 'report.md',
        path: '/abs/report.md',
      }),
    });

    const { result } = renderHook(() => useFileHandler());

    await act(async () => {
      await result.current.openDialog();
    });

    expect(window.electronAPI!.openFileByPath).toHaveBeenCalledWith('/abs/report.md');
    expect(result.current.file).toMatchObject({
      kind: 'text',
      format: 'markdown',
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

    const { result } = renderHook(() => useFileHandler());

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

    const { result } = renderHook(() => useFileHandler());

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

    const { result } = renderHook(() => useFileHandler());

    await act(async () => {
      await result.current.loadFromPath('/abs/doc.pdf');
    });

    expect(result.current.file).toBeNull();
    expect(result.current.error).toBe('IPC failure');
    expect(result.current.loading).toBe(false);
  });
});
