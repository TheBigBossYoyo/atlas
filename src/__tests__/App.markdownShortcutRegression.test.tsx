/**
 * App — markdown shortcut regression gate (P2.1).
 *
 * The improvement plan requires this specific regression test to gate
 * P2.1's own merge (not deferred to a later e2e pass): verifying Ctrl+1/2/3
 * view-mode switching and RawEditor's Ctrl+B/I still behave identically
 * after replacing six independent window-level keydown listeners with the
 * centralized dispatcher. A markdown document has no active-viewer
 * registration at all (that tier is currently only used by DocxViewer), so
 * every shortcut below reaches the shell-global tier exactly as it always
 * has — this suite exists to prove that migration didn't change markdown's
 * observable shortcut behavior, per the owner's "markdown is perfect, don't
 * change its behavior" constraint.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import App from '../App';

function buildElectronAPI(): typeof window.electronAPI {
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
  } as typeof window.electronAPI;
}

beforeEach(() => {
  localStorage.clear();
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

function loadSample() {
  fireEvent.click(screen.getByRole('button', { name: 'Load Sample Document' }));
}

describe('App — markdown shortcut regression gate (P2.1)', () => {
  it('Ctrl+1/2/3 still switch between Preview/Split/Editor view modes', async () => {
    render(<App />);
    loadSample();
    await waitFor(() => expect(document.querySelector('.content--preview')).toBeInTheDocument());

    fireEvent.keyDown(window, { key: '2', ctrlKey: true });
    await waitFor(() => expect(document.querySelector('.content--split')).toBeInTheDocument());

    fireEvent.keyDown(window, { key: '3', ctrlKey: true });
    await waitFor(() => expect(document.querySelector('.content--editor')).toBeInTheDocument());

    fireEvent.keyDown(window, { key: '1', ctrlKey: true });
    await waitFor(() => expect(document.querySelector('.content--preview')).toBeInTheDocument());
  });

  it('RawEditor: Ctrl+B while typing does not toggle the sidebar (textarea was already a plain field pre-dispatcher)', async () => {
    render(<App />);
    loadSample();
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));

    const textarea = screen.getByPlaceholderText('Type or paste markdown here...');
    expect(document.querySelector('.sidebar')).toBeInTheDocument();

    fireEvent.keyDown(textarea, { key: 'b', ctrlKey: true });

    expect(document.querySelector('.sidebar')).toBeInTheDocument();
  });

  it('RawEditor: Ctrl+I while typing is inert (no shell-global binding, no crash)', async () => {
    render(<App />);
    loadSample();
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));

    const textarea = screen.getByPlaceholderText('Type or paste markdown here...');
    fireEvent.change(textarea, { target: { value: 'hello' } });

    expect(() => fireEvent.keyDown(textarea, { key: 'i', ctrlKey: true })).not.toThrow();
    expect(textarea).toHaveValue('hello');
  });

  it('Ctrl+B toggles the sidebar from outside a text field, exactly as before the dispatcher', async () => {
    render(<App />);
    loadSample();
    expect(document.querySelector('.sidebar')).toBeInTheDocument();

    fireEvent.keyDown(window, { key: 'b', ctrlKey: true });
    expect(document.querySelector('.sidebar')).not.toBeInTheDocument();

    fireEvent.keyDown(window, { key: 'b', ctrlKey: true });
    expect(document.querySelector('.sidebar')).toBeInTheDocument();
  });
});
