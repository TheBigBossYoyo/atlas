/**
 * useShortcutManager / ShortcutManagerProvider — P2.1/SHELL-19/D5 regression
 * tests for the centralized shortcut dispatcher's precedence rules, in
 * isolation from the full App tree (App.test.tsx covers the real
 * SHELL-08/09/UX-03 collision scenarios end to end).
 */
import { fireEvent, render } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';

import { ShortcutManagerProvider } from '../ShortcutManagerProvider';
import { useShellShortcut, useViewerShortcuts } from '../useShortcutManager';
import type { ShortcutHandler } from '../shortcutManagerContext';

function pressCtrlB() {
  fireEvent.keyDown(window, { key: 'b', ctrlKey: true });
}

function Probe({
  shell,
  viewer,
  shellEnabled = true,
  viewerEnabled = true,
}: {
  shell?: ShortcutHandler;
  viewer?: ShortcutHandler;
  shellEnabled?: boolean;
  viewerEnabled?: boolean;
}) {
  useShellShortcut(shell ?? (() => false), shellEnabled);
  useViewerShortcuts(viewer ?? (() => false), viewerEnabled);
  return null;
}

describe('ShortcutManagerProvider — precedence', () => {
  it('calls the shell-global handler when no viewer handler is registered', () => {
    const shell = vi.fn().mockReturnValue(true);
    render(
      <ShortcutManagerProvider>
        <Probe shell={shell} />
      </ShortcutManagerProvider>,
    );

    pressCtrlB();

    expect(shell).toHaveBeenCalledTimes(1);
  });

  it('a claiming viewer handler wins over shell-global (SHELL-08/RUN-02: Ctrl+B must not also toggle the sidebar)', () => {
    const shell = vi.fn().mockReturnValue(true);
    const viewer = vi.fn().mockReturnValue(true);
    render(
      <ShortcutManagerProvider>
        <Probe shell={shell} viewer={viewer} />
      </ShortcutManagerProvider>,
    );

    pressCtrlB();

    expect(viewer).toHaveBeenCalledTimes(1);
    expect(shell).not.toHaveBeenCalled();
  });

  it('falls through to shell-global when the viewer handler does not claim the event', () => {
    const shell = vi.fn().mockReturnValue(true);
    const viewer = vi.fn().mockReturnValue(false);
    render(
      <ShortcutManagerProvider>
        <Probe shell={shell} viewer={viewer} />
      </ShortcutManagerProvider>,
    );

    pressCtrlB();

    expect(viewer).toHaveBeenCalledTimes(1);
    expect(shell).toHaveBeenCalledTimes(1);
  });

  it('reports inPlainField for an input, a textarea, and a contentEditable div alike (SHELL-08 fix: widened past input/textarea)', () => {
    const seen: boolean[] = [];
    const shell: ShortcutHandler = (_event, ctx) => {
      seen.push(ctx.inPlainField);
      return true;
    };

    render(
      <ShortcutManagerProvider>
        <Probe shell={shell} />
        <input data-testid="plain-input" />
        <textarea data-testid="plain-textarea" />
        <div data-testid="rich-text" contentEditable suppressContentEditableWarning />
        <div data-testid="plain-div" />
      </ShortcutManagerProvider>,
    );

    fireEvent.keyDown(document.querySelector('[data-testid="plain-input"]')!, { key: 'b', ctrlKey: true });
    fireEvent.keyDown(document.querySelector('[data-testid="plain-textarea"]')!, { key: 'b', ctrlKey: true });
    fireEvent.keyDown(document.querySelector('[data-testid="rich-text"]')!, { key: 'b', ctrlKey: true });
    fireEvent.keyDown(document.querySelector('[data-testid="plain-div"]')!, { key: 'b', ctrlKey: true });

    expect(seen).toEqual([true, true, true, false]);
  });

  it('does not register (and does not call) a handler while enabled=false', () => {
    const shell = vi.fn().mockReturnValue(true);
    const { rerender } = render(
      <ShortcutManagerProvider>
        <Probe shell={shell} shellEnabled={false} />
      </ShortcutManagerProvider>,
    );

    pressCtrlB();
    expect(shell).not.toHaveBeenCalled();

    rerender(
      <ShortcutManagerProvider>
        <Probe shell={shell} shellEnabled={true} />
      </ShortcutManagerProvider>,
    );
    pressCtrlB();
    expect(shell).toHaveBeenCalledTimes(1);
  });

  it('always calls the latest handler closure without re-registering on every render', () => {
    const calls: number[] = [];
    function Counting({ n }: { n: number }) {
      useShellShortcut(() => {
        calls.push(n);
        return true;
      });
      return null;
    }

    const { rerender } = render(
      <ShortcutManagerProvider>
        <Counting n={1} />
      </ShortcutManagerProvider>,
    );
    rerender(
      <ShortcutManagerProvider>
        <Counting n={2} />
      </ShortcutManagerProvider>,
    );

    pressCtrlB();

    expect(calls).toEqual([2]);
  });

  it('within a tier, the most-recently-registered handler is tried first', () => {
    const order: string[] = [];
    function Handler({ name, claim }: { name: string; claim: boolean }) {
      useShellShortcut(() => {
        order.push(name);
        return claim;
      });
      return null;
    }

    render(
      <ShortcutManagerProvider>
        <Handler name="first" claim={false} />
        <Handler name="second" claim={true} />
      </ShortcutManagerProvider>,
    );

    pressCtrlB();

    expect(order).toEqual(['second']);
  });
});

describe('useShellShortcut / useViewerShortcuts — no provider', () => {
  it('silently registers nothing when rendered outside a ShortcutManagerProvider (existing component tests do this)', () => {
    const shell = vi.fn();
    expect(() => render(<Probe shell={shell} />)).not.toThrow();

    pressCtrlB();
    expect(shell).not.toHaveBeenCalled();
  });
});
