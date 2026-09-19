/**
 * Shared `window.electronAPI` test double (P4.8 / QA-25).
 *
 * Before this file, every test that needed `window.electronAPI` hand-rolled
 * its own literal matching `src/electron.d.ts`'s `ElectronAPI` shape (see
 * e.g. `src/hooks/__tests__/useFileHandler.test.ts`'s original
 * `buildElectronAPI`) — harmless on its own, but it meant N copies to keep
 * in sync by hand every time the real interface grew a field, with no
 * compiler help pointing at the other N-1 copies. `createMockElectronAPI()`
 * is the one place that mapping lives now.
 *
 * Usage:
 * ```ts
 * import { createMockElectronAPI } from '../../__tests__/mocks/electronAPI';
 *
 * window.electronAPI = createMockElectronAPI({
 *   // override only what this test cares about
 *   readBinaryByPath: vi.fn().mockResolvedValue({ path: '/abs/doc.pdf', buffer }),
 * });
 * ```
 *
 * Every field is a `vi.fn()` with a reasonable default (resolves to "no
 * file" / "cancelled" / a no-op unsubscribe) so a test that doesn't care
 * about IPC at all can still assign the return value straight to
 * `window.electronAPI` and have the app boot without throwing. Pass
 * `overrides` for the handful of calls a given test actually asserts on or
 * needs to resolve differently.
 *
 * Remember to restore the ambient global afterwards, e.g. in `afterEach`:
 * ```ts
 * delete (window as { electronAPI?: unknown }).electronAPI;
 * ```
 * so a mock assigned in one test can't leak into the next.
 */
import { vi } from 'vitest';

export type MockElectronAPI = NonNullable<Window['electronAPI']>;

export function createMockElectronAPI(overrides: Partial<MockElectronAPI> = {}): MockElectronAPI {
  const base: MockElectronAPI = {
    getInitialFile: vi.fn().mockResolvedValue(null),
    openFileDialog: vi.fn().mockResolvedValue(null),
    openFileByPath: vi.fn().mockResolvedValue({ content: '', name: 'file', path: '' }),
    saveFile: vi.fn().mockResolvedValue({ saved: false }),
    saveBinaryFile: vi.fn().mockResolvedValue({ saved: false }),
    onFileOpened: vi.fn().mockReturnValue(() => {}),
    setTheme: vi.fn(),
    notifyDirtyState: vi.fn(),
    onRequestSaveBeforeClose: vi.fn().mockReturnValue(() => {}),
    reportSaveBeforeCloseResult: vi.fn(),
    openFileBinary: vi.fn().mockResolvedValue({ canceled: true, path: '', buffer: new ArrayBuffer(0) }),
    newDocument: vi.fn().mockResolvedValue({ created: false }),
    readBinaryByPath: vi.fn().mockResolvedValue({ path: '', buffer: new ArrayBuffer(0) }),
    onFileOpenedPath: vi.fn().mockReturnValue(() => {}),
    getPathForFile: vi.fn().mockReturnValue(''),
    registerDroppedPath: vi.fn().mockResolvedValue({ ok: true }),
    requestOpenRecent: vi.fn().mockResolvedValue({ ok: true }),
    revealInFolder: vi.fn().mockResolvedValue({ ok: true }),
    image: { pick: vi.fn().mockResolvedValue({ cancelled: true }) },
    spellcheck: {
      onContextMenu: vi.fn().mockReturnValue(() => {}),
      replaceMisspelling: vi.fn().mockResolvedValue({}),
      addWord: vi.fn().mockResolvedValue({ added: true }),
      getLanguages: vi.fn().mockResolvedValue({ available: [], enabled: [] }),
      setLanguages: vi.fn().mockResolvedValue({ ok: true }),
    },
  };

  return { ...base, ...overrides };
}
