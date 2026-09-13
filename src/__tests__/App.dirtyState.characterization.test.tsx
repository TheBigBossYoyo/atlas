/**
 * App — dirty-state characterization suite (P0.4 / QA-07).
 *
 * Freezes today's exact isDirty / Save-button / active-file transition
 * sequence through markdown's own load → edit → save → open-another-file →
 * open-a-non-markdown-file flow (App.tsx's dirty-tracking state machine,
 * roughly lines 80-193). Per the improvement plan's Section 7 guardrail,
 * P1.1 (capability contract) and P2.10 (dirty-reset keying) both modify
 * code this test exercises directly and must leave this sequence
 * unchanged for the markdown-to-markdown case (P1.1 explicitly adds a
 * discard-changes confirm dialog on this flow for OTHER formats — this
 * suite's baseline, captured against current `main`, has none; a later
 * wave adding one for non-markdown files should not change this sequence,
 * since every step below already saves before switching files).
 *
 * Note (CSS blind spot): assertions below read DOM structure/attributes
 * (an element's presence, a `disabled` attribute, text content) — jsdom
 * loads no stylesheet, so this cannot see how the dirty dot or a disabled
 * Save button actually *look*, only whether/what they are.
 *
 * Note (P2.7 behavior change, wave1/viewer-quickfixes): the global Save
 * button used to render disabled-but-visible for non-markdown formats;
 * it is now omitted from the DOM entirely (Toolbar's `canSave` gate).
 * `saveDisabled` is `null` at step 5 to reflect that — a real, intended
 * behavior change for the non-markdown path, not a markdown regression.
 * Steps 1-4 (all markdown) are unaffected and unchanged from baseline.
 *
 * Non-markdown viewers are lazy-loaded via `ViewerRouter` + `formats/
 * registry` and pull in real, heavy dependencies (react-window, shiki,
 * SheetJS, ...) irrelevant to dirty-state tracking, so `formats/registry`
 * is mocked to a trivial stand-in — this suite cares which format loaded,
 * never how that format renders.
 */

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, fireEvent, render, screen, waitFor } from '@testing-library/react';

vi.mock('../formats/registry', () => ({
  viewerRegistry: new Proxy(
    {} as Record<string, () => Promise<unknown>>,
    {
      get: () => () =>
        Promise.resolve(function FakeViewer() {
          return <div data-testid="fake-viewer">FAKE VIEWER</div>;
        }),
    },
  ),
}));

import App from '../App';

// ---------------------------------------------------------------------------
// window.electronAPI mock — mirrors src/hooks/__tests__/useFileHandler.test.ts's
// buildElectronAPI, extended with the handlers ElectronAPI additionally
// requires (onSpellCheckMenu/replaceMisspelling/addWordToDictionary) so
// assigning the result to `window.electronAPI` type-checks.
// ---------------------------------------------------------------------------

type MockElectronAPI = NonNullable<typeof window.electronAPI>;

function buildElectronAPI(overrides: Partial<MockElectronAPI> = {}): MockElectronAPI {
  return {
    getInitialFile: vi.fn().mockResolvedValue(null),
    openFileDialog: vi.fn().mockResolvedValue(null),
    openFileByPath: vi.fn().mockResolvedValue({ content: '', name: 'file', path: '' }),
    saveFile: vi.fn().mockResolvedValue({ saved: true }),
    onFileOpened: vi.fn().mockReturnValue(() => {}),
    setTheme: vi.fn(),
    openFileBinary: vi.fn().mockResolvedValue({ canceled: true, path: '', buffer: new ArrayBuffer(0) }),
    readBinaryByPath: vi.fn().mockResolvedValue({ path: '', buffer: new ArrayBuffer(0) }),
    onFileOpenedPath: vi.fn().mockReturnValue(() => {}),
    onSpellCheckMenu: vi.fn().mockReturnValue(() => {}),
    replaceMisspelling: vi.fn().mockResolvedValue({}),
    addWordToDictionary: vi.fn().mockResolvedValue({}),
    ...overrides,
  } as MockElectronAPI;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

interface CapturedState {
  readonly step: string;
  readonly fileName: string | null;
  /**
   * The Toolbar's own dirty dot (`.toolbar__filename .toolbar__dirty`) is
   * nested INSIDE the filename badge, which App.tsx only renders at all
   * when `fileName` is truthy — so it reads `false` whenever no real file
   * is open yet (e.g. right after "Load Sample Document"), even while the
   * StatusBar's independent dirty dot (gated only on `hasContent`, not on
   * `fileName`) reports the same `isDirty` state as `true`. Both are
   * captured separately below because that divergence between two UI
   * indicators of the *same* boolean is itself real, current behavior this
   * suite exists to freeze — not a test bug to paper over.
   */
  readonly toolbarDirtyDot: boolean;
  readonly statusBarDirtyDot: boolean;
  /** `null` when the Save button isn't rendered at all (P2.7: `canSave`
   * gates its presence, not just its `disabled` attribute). */
  readonly saveDisabled: boolean | null;
  /** Which top-level content element is on screen right now: 'markdown'
   * (`.markdown-body`), a non-markdown FormatId (`ViewerRouter`'s
   * `data-viewer`), or null — including while viewMode='editor' shows only
   * RawEditor's textarea and neither element is mounted. */
  readonly activeFormat: string | null;
}

function captureState(step: string): CapturedState {
  const saveButton = screen.queryByTitle('Save (Ctrl+S)') as HTMLButtonElement | null;
  const filenameEl = document.querySelector('.toolbar__filename');
  const viewerEl = document.querySelector('[data-viewer]');
  const isMarkdownActive = document.querySelector('.markdown-body') !== null;

  return {
    step,
    fileName: filenameEl ? (filenameEl.childNodes[0]?.textContent ?? null) : null,
    toolbarDirtyDot: document.querySelector('.toolbar__dirty') !== null,
    statusBarDirtyDot: document.querySelector('.statusbar__dirty') !== null,
    saveDisabled: saveButton ? saveButton.disabled : null,
    activeFormat: viewerEl ? viewerEl.getAttribute('data-viewer') : isMarkdownActive ? 'markdown' : null,
  };
}

function makeTextBuffer(text: string): ArrayBuffer {
  return new TextEncoder().encode(text).buffer;
}

// ---------------------------------------------------------------------------
// Setup / teardown
// ---------------------------------------------------------------------------

beforeEach(() => {
  localStorage.clear();
  // useTheme() reads matchMedia when no theme is stored yet; jsdom has no
  // implementation of its own.
  window.matchMedia = vi.fn().mockImplementation((query: string) => ({
    matches: false,
    media: query,
    onchange: null,
    addListener: vi.fn(),
    removeListener: vi.fn(),
    addEventListener: vi.fn(),
    removeEventListener: vi.fn(),
    dispatchEvent: vi.fn(),
  }));
});

afterEach(() => {
  vi.restoreAllMocks();
  localStorage.clear();
  delete (window as { electronAPI?: unknown }).electronAPI;
  delete (window as { matchMedia?: unknown }).matchMedia;
});

// ---------------------------------------------------------------------------
// Test
// ---------------------------------------------------------------------------

describe('App dirty-state characterization', () => {
  it('records isDirty / Save-button / active-file at each step of load sample -> edit -> save -> open another .md -> open a non-.md file', async () => {
    const openFileBinaryMock = vi
      .fn()
      .mockResolvedValueOnce({ canceled: false, path: '/abs/other.md', buffer: makeTextBuffer('') })
      .mockResolvedValueOnce({ canceled: false, path: '/abs/notes.txt', buffer: makeTextBuffer('') });
    const openFileByPathMock = vi
      .fn()
      .mockResolvedValueOnce({ content: '# Other File\n\nDifferent content.', name: 'other.md', path: '/abs/other.md' })
      .mockResolvedValueOnce({ content: 'Plain text notes, not markdown.', name: 'notes.txt', path: '/abs/notes.txt' });

    window.electronAPI = buildElectronAPI({
      openFileBinary: openFileBinaryMock,
      openFileByPath: openFileByPathMock,
    });

    render(<App />);
    const transitions: CapturedState[] = [];

    // 1. Load sample --------------------------------------------------------
    fireEvent.click(screen.getByRole('button', { name: 'Load Sample Document' }));
    transitions.push(captureState('1. load sample'));

    // 2. Edit -----------------------------------------------------------------
    fireEvent.click(screen.getByRole('button', { name: 'Editor' }));
    const textarea = screen.getByPlaceholderText('Type or paste markdown here...') as HTMLTextAreaElement;
    fireEvent.change(textarea, { target: { value: `${textarea.value}\n\nEdited by the characterization test.` } });
    // Switch back to Preview so `activeFormat` reads '.markdown-body' again
    // instead of null (in 'editor' view mode only RawEditor's textarea is
    // mounted — MarkdownRenderer isn't rendered at all until 'preview'/'split').
    fireEvent.click(screen.getByRole('button', { name: 'Preview' }));
    transitions.push(captureState('2. edit'));

    // 3. Save -------------------------------------------------------------
    await act(async () => {
      fireEvent.click(screen.getByTitle('Save (Ctrl+S)'));
    });
    await waitFor(() => expect(document.querySelector('.statusbar__dirty')).toBeNull());
    transitions.push(captureState('3. save'));

    // 4. Open another markdown file -----------------------------------------
    fireEvent.click(screen.getByTitle('Open file (Ctrl+O)'));
    await waitFor(() => {
      const filenameEl = document.querySelector('.toolbar__filename');
      expect(filenameEl?.childNodes[0]?.textContent).toBe('other.md');
    });
    transitions.push(captureState('4. open another markdown file'));

    // 5. Open a non-markdown file --------------------------------------------
    fireEvent.click(screen.getByTitle('Open file (Ctrl+O)'));
    await waitFor(() => {
      expect(screen.getByTestId('fake-viewer')).toBeInTheDocument();
    });
    transitions.push(captureState('5. open a non-markdown file'));

    expect(openFileBinaryMock).toHaveBeenCalledTimes(2);
    expect(openFileByPathMock).toHaveBeenCalledTimes(2);

    expect(transitions).toMatchInlineSnapshot(`
      [
        {
          "activeFormat": "markdown",
          "fileName": null,
          "saveDisabled": false,
          "statusBarDirtyDot": true,
          "step": "1. load sample",
          "toolbarDirtyDot": false,
        },
        {
          "activeFormat": "markdown",
          "fileName": null,
          "saveDisabled": false,
          "statusBarDirtyDot": true,
          "step": "2. edit",
          "toolbarDirtyDot": false,
        },
        {
          "activeFormat": "markdown",
          "fileName": null,
          "saveDisabled": false,
          "statusBarDirtyDot": false,
          "step": "3. save",
          "toolbarDirtyDot": false,
        },
        {
          "activeFormat": "markdown",
          "fileName": "other.md",
          "saveDisabled": false,
          "statusBarDirtyDot": false,
          "step": "4. open another markdown file",
          "toolbarDirtyDot": false,
        },
        {
          "activeFormat": "text",
          "fileName": "notes.txt",
          "saveDisabled": null,
          "statusBarDirtyDot": false,
          "step": "5. open a non-markdown file",
          "toolbarDirtyDot": false,
        },
      ]
    `);
  });
});
