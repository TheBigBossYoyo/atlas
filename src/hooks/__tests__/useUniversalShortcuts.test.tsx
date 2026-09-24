/**
 * useUniversalShortcuts — FIELD-01 regression coverage for the
 * `ctx.inPlainField` gate: Ctrl+W must reach `closeFile` even while a plain
 * text field/contentEditable has focus (mirroring Ctrl+S/Ctrl+P's existing
 * carve-out), while the gate must still do its original job for a shortcut
 * that has no business firing out from under a focused field (Ctrl+B, which
 * a contentEditable rich-text surface treats natively as "toggle bold").
 */
import { fireEvent, render } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';

import { ShortcutManagerProvider } from '../ShortcutManagerProvider';
import { useUniversalShortcuts } from '../useUniversalShortcuts';

function baseProps(overrides: Partial<Parameters<typeof useUniversalShortcuts>[0]> = {}) {
  return {
    openFile: vi.fn(),
    saveFile: vi.fn(),
    saveFileAs: vi.fn(),
    exportPrimary: vi.fn(),
    openExportMenu: vi.fn(),
    openNewMenu: vi.fn(),
    cycleTheme: vi.fn(),
    toggleSidebar: vi.fn(),
    setViewMode: vi.fn(),
    toggleShortcuts: vi.fn(),
    closeFile: vi.fn(),
    isMarkdown: true,
    ...overrides,
  };
}

function Probe(props: ReturnType<typeof baseProps>) {
  useUniversalShortcuts(props);
  return (
    <div>
      <textarea data-testid="plain-field" />
      <div data-testid="not-a-field" tabIndex={-1} />
    </div>
  );
}

describe('useUniversalShortcuts — inPlainField gate', () => {
  it('FIELD-01: Ctrl+W reaches closeFile even with a plain textarea focused (was silently swallowed)', () => {
    const props = baseProps();
    render(
      <ShortcutManagerProvider>
        <Probe {...props} />
      </ShortcutManagerProvider>,
    );

    fireEvent.keyDown(document.querySelector('[data-testid="plain-field"]')!, {
      key: 'w',
      ctrlKey: true,
    });

    expect(props.closeFile).toHaveBeenCalledTimes(1);
  });

  it('still calls closeFile for Ctrl+W outside any field (no regression)', () => {
    const props = baseProps();
    render(
      <ShortcutManagerProvider>
        <Probe {...props} />
      </ShortcutManagerProvider>,
    );

    fireEvent.keyDown(document.querySelector('[data-testid="not-a-field"]')!, {
      key: 'w',
      ctrlKey: true,
    });

    expect(props.closeFile).toHaveBeenCalledTimes(1);
  });

  it('the gate still does its original job: Ctrl+B inside a plain field does not toggle the sidebar', () => {
    const props = baseProps();
    render(
      <ShortcutManagerProvider>
        <Probe {...props} />
      </ShortcutManagerProvider>,
    );

    fireEvent.keyDown(document.querySelector('[data-testid="plain-field"]')!, {
      key: 'b',
      ctrlKey: true,
    });

    expect(props.toggleSidebar).not.toHaveBeenCalled();
  });

  it('Ctrl+B outside any field still toggles the sidebar', () => {
    const props = baseProps();
    render(
      <ShortcutManagerProvider>
        <Probe {...props} />
      </ShortcutManagerProvider>,
    );

    fireEvent.keyDown(document.querySelector('[data-testid="not-a-field"]')!, {
      key: 'b',
      ctrlKey: true,
    });

    expect(props.toggleSidebar).toHaveBeenCalledTimes(1);
  });

  // SHORTCUT-FIELD-1 — found by the 3.7.0 installer smoke test and the shell
  // sweep: Ctrl+O did nothing with the caret in a Word document, the markdown
  // editor or the code editor. Like Ctrl+S/Ctrl+W, open/new are file-level
  // commands with no in-field meaning to protect.
  it.each([
    ['o', 'openFile'],
    ['n', 'openNewMenu'],
  ] as const)('SHORTCUT-FIELD-1: Ctrl+%s reaches %s even with a plain field focused', (key, action) => {
    const props = baseProps();
    render(
      <ShortcutManagerProvider>
        <Probe {...props} />
      </ShortcutManagerProvider>,
    );

    fireEvent.keyDown(document.querySelector('[data-testid="plain-field"]')!, { key, ctrlKey: true });

    expect(props[action]).toHaveBeenCalledTimes(1);
  });

  it('SHORTCUT-FIELD-1: Ctrl+E and Ctrl+T stay gated inside a field (a rich-text surface may claim them)', () => {
    const props = baseProps();
    render(
      <ShortcutManagerProvider>
        <Probe {...props} />
      </ShortcutManagerProvider>,
    );

    fireEvent.keyDown(document.querySelector('[data-testid="plain-field"]')!, { key: 'e', ctrlKey: true });
    fireEvent.keyDown(document.querySelector('[data-testid="plain-field"]')!, { key: 't', ctrlKey: true });

    expect(props.openExportMenu).not.toHaveBeenCalled();
    expect(props.cycleTheme).not.toHaveBeenCalled();
  });

  it('a plain letter with no modifier is never touched regardless of focus (sanity: handler bails before inPlainField is even read)', () => {
    const props = baseProps();
    render(
      <ShortcutManagerProvider>
        <Probe {...props} />
      </ShortcutManagerProvider>,
    );

    fireEvent.keyDown(document.querySelector('[data-testid="plain-field"]')!, { key: 'w' });
    fireEvent.keyDown(document.querySelector('[data-testid="plain-field"]')!, { key: 'b' });

    expect(props.closeFile).not.toHaveBeenCalled();
    expect(props.toggleSidebar).not.toHaveBeenCalled();
  });
});
