/**
 * App — export wiring integration tests (X2/UX-05, UX-11, UX-12, UX-18).
 *
 * Exercises what changed in App.tsx's `handleExport`/export-menu wiring
 * specifically: the markdown HTML export now serializes the live-rendered
 * DOM (not the raw markdown string), CSV export routes csv/tsv file content
 * through `exportCsv`, and an export failure surfaces as a toast instead of
 * a blocking native `alert()`. The exportX() functions' own internals
 * (Electron routing, friendly-error wrapping, DOM serialization details) are
 * covered at the unit level in `src/utils/export/__tests__/` and
 * `src/utils/__tests__/export.markdown.test.tsx` — these tests only check
 * that App.tsx calls them with the right thing.
 */
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import App from '../App';

// PERF-01 — `App.tsx`'s export handlers now `await import('./utils/export')`
// on demand instead of a top-level `import { exportXxx, ... }` (so cold
// start / a non-export session never pays for bundling `docx`, `xlsx`,
// `html2canvas-pro`, and `react-dom/server` — see App.tsx's own PERF-01
// comment for the measured effect). In the real app that dynamic import is
// just a `fetch()` of an already-built, pre-transformed chunk — effectively
// instant. Under Vitest, though, a dynamic `import()` is transformed
// on-demand the FIRST time it actually runs (unlike a statically-imported
// module, which Vite's SSR module runner pre-transforms ahead of the test),
// and that one-time esbuild pass over `utils/export`'s whole dependency
// graph measurably exceeded `waitFor`'s default 1000ms budget on this
// (shared, concurrently-loaded) machine — a test-environment artifact, not
// a real regression: every test below still asserts the exact same
// behavior, just with room for that one-time cost.
const EXPORT_WAIT_OPTIONS = { timeout: 15_000 } as const;

function buildElectronAPI(overrides: Partial<typeof window.electronAPI> = {}): typeof window.electronAPI {
  return {
    getInitialFile: vi.fn().mockResolvedValue(null),
    openFileDialog: vi.fn().mockResolvedValue(null),
    openFileByPath: vi.fn().mockResolvedValue({ content: '', name: 'file.md', path: '' }),
    saveFile: vi.fn().mockResolvedValue({ saved: true }),
    saveBinaryFile: vi.fn().mockResolvedValue({ saved: true }),
    onFileOpened: vi.fn().mockReturnValue(() => {}),
    setTheme: vi.fn(),
    openFileBinary: vi.fn().mockResolvedValue({ canceled: true, path: '', buffer: new ArrayBuffer(0) }),
    readBinaryByPath: vi.fn().mockResolvedValue({ path: '', buffer: new ArrayBuffer(0) }),
    onFileOpenedPath: vi.fn().mockReturnValue(() => {}),
    ...overrides,
  } as typeof window.electronAPI;
}

function mockOpenTextFile(path: string, content: string) {
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

// The loaded filename renders in both the Toolbar and the StatusBar (same
// text) — query the Toolbar's copy directly, matching App.test.tsx's own
// helper, rather than an ambiguous screen.getByText.
function toolbarFilenameText(): string {
  const el = document.querySelector('.toolbar__filename');
  return el?.firstChild?.textContent?.trim() ?? '';
}

beforeEach(() => {
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

describe('App — markdown HTML export serializes the live DOM (X2/UX-05)', () => {
  it('exports rendered KaTeX markup, not the raw "$..$" source', async () => {
    render(<App />);
    mockOpenTextFile('/abs/math.md', '# Title\n\nInline math: $E = mc^2$ done.');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('math.md'));

    fireEvent.click(screen.getByRole('button', { name: 'Export' }));
    fireEvent.click(screen.getByRole('button', { name: 'HTML' }));

    await waitFor(() => expect(window.electronAPI!.saveFile).toHaveBeenCalled(), EXPORT_WAIT_OPTIONS);
    const call = (window.electronAPI!.saveFile as ReturnType<typeof vi.fn>).mock.calls[0]![0];
    expect(call.content).toContain('class="katex"');
    expect(call.content).not.toContain('$E = mc^2$');
    expect(call.suggestedName).toBe('math.html');
  });
});

describe('App — CSV export for csv/tsv files (UX-12)', () => {
  it('exports the raw csv file content unchanged', async () => {
    render(<App />);
    mockOpenTextFile('/abs/data.csv', 'name,age\nAda,36');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('data.csv'));

    fireEvent.click(screen.getByRole('button', { name: 'Export' }));
    fireEvent.click(screen.getByRole('button', { name: 'Export to CSV' }));

    await waitFor(() => expect(window.electronAPI!.saveFile).toHaveBeenCalled(), EXPORT_WAIT_OPTIONS);
    const call = (window.electronAPI!.saveFile as ReturnType<typeof vi.fn>).mock.calls[0]![0];
    expect(call.content).toBe('name,age\nAda,36');
    expect(call.suggestedName).toBe('data.csv');
  });

  it('converts tab-delimited tsv file content to comma-delimited csv', async () => {
    render(<App />);
    mockOpenTextFile('/abs/data.tsv', 'name\tage\nAda\t36');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('data.tsv'));

    fireEvent.click(screen.getByRole('button', { name: 'Export' }));
    fireEvent.click(screen.getByRole('button', { name: 'Export to CSV' }));

    await waitFor(() => expect(window.electronAPI!.saveFile).toHaveBeenCalled(), EXPORT_WAIT_OPTIONS);
    const call = (window.electronAPI!.saveFile as ReturnType<typeof vi.fn>).mock.calls[0]![0];
    expect(call.content).toBe('name,age\r\nAda,36');
  });
});

describe('App — export failure surfaces as a toast, not a blocking alert() (UX-18)', () => {
  it('shows a themed toast and never calls window.alert', async () => {
    const alertSpy = vi.spyOn(window, 'alert').mockImplementation(() => {});
    window.electronAPI!.saveFile = vi.fn().mockResolvedValue({ saved: false, error: 'Disk is full' });

    render(<App />);
    mockOpenTextFile('/abs/notes.md', '# Notes');
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('notes.md'));

    fireEvent.click(screen.getByRole('button', { name: 'Export' }));
    fireEvent.click(screen.getByRole('button', { name: 'Markdown' }));

    const toast = await screen.findByRole('alert', {}, EXPORT_WAIT_OPTIONS);
    expect(toast).toHaveTextContent(/disk is full/i);
    expect(alertSpy).not.toHaveBeenCalled();
  });
});

describe('App — PDF "Save a copy" (X1 — this used to bypass the native save dialog entirely with a raw browser download, even inside Electron)', () => {
  it('copies the original bytes through window.electronAPI.saveBinaryFile, not a synthetic <a download> click', async () => {
    const pdfBytes = new TextEncoder().encode('%PDF-1.4 fake pdf bytes').buffer;
    window.electronAPI!.openFileBinary = vi.fn().mockResolvedValue({
      canceled: false,
      path: '/abs/report.pdf',
      buffer: pdfBytes,
    });
    const saveBinaryFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI!.saveBinaryFile = saveBinaryFileMock;
    const anchorClickSpy = vi.spyOn(HTMLAnchorElement.prototype, 'click');

    render(<App />);
    await openViaToolbar();
    await waitFor(() => expect(toolbarFilenameText()).toBe('report.pdf'));

    fireEvent.click(screen.getByRole('button', { name: 'Export' }));
    fireEvent.click(screen.getByRole('button', { name: 'Save a copy' }));

    await waitFor(() => expect(saveBinaryFileMock).toHaveBeenCalled(), EXPORT_WAIT_OPTIONS);
    expect(saveBinaryFileMock).toHaveBeenCalledWith(
      expect.objectContaining({ suggestedName: 'report.pdf' }),
    );
    expect(anchorClickSpy).not.toHaveBeenCalled();
  });
});
