/**
 * useTheme — SHELL-25 regression test: a live OS dark/light toggle should
 * update the app's theme immediately while the user hasn't made an explicit
 * choice, and should be ignored once they have.
 */
import { act, renderHook } from '@testing-library/react';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { useTheme } from '../useTheme';

type ChangeListener = (event: { matches: boolean }) => void;

function installMatchMediaMock(initialMatches: boolean) {
  let matches = initialMatches;
  const listeners = new Set<ChangeListener>();

  Object.defineProperty(window, 'matchMedia', {
    writable: true,
    configurable: true,
    value: (query: string) => ({
      get matches() {
        return matches;
      },
      media: query,
      onchange: null,
      addListener: () => {},
      removeListener: () => {},
      addEventListener: (_type: string, listener: ChangeListener) => {
        listeners.add(listener);
      },
      removeEventListener: (_type: string, listener: ChangeListener) => {
        listeners.delete(listener);
      },
      dispatchEvent: () => true,
    }),
  });

  return {
    setSystemDark(next: boolean) {
      matches = next;
      for (const listener of listeners) {
        listener({ matches: next });
      }
    },
    listenerCount: () => listeners.size,
  };
}

beforeEach(() => {
  localStorage.clear();
  delete (window as { electronAPI?: unknown }).electronAPI;
});

afterEach(() => {
  localStorage.clear();
});

describe('useTheme — live system theme changes (SHELL-25)', () => {
  it('follows a live OS dark/light toggle when the user has never made an explicit choice', () => {
    const mql = installMatchMediaMock(false);
    const { result } = renderHook(() => useTheme());

    expect(result.current.theme).toBe('light');

    act(() => {
      mql.setSystemDark(true);
    });

    expect(result.current.theme).toBe('dark');
  });

  it('stops following the OS once the user has explicitly picked a theme via setTheme', () => {
    const mql = installMatchMediaMock(false);
    const { result } = renderHook(() => useTheme());

    act(() => {
      result.current.setTheme('sepia');
    });
    expect(result.current.theme).toBe('sepia');

    act(() => {
      mql.setSystemDark(true);
    });

    // The explicit choice sticks — a live OS toggle no longer overrides it.
    expect(result.current.theme).toBe('sepia');
  });

  it('stops following the OS once the user has explicitly picked a theme via cycleTheme', () => {
    const mql = installMatchMediaMock(false);
    const { result } = renderHook(() => useTheme());

    act(() => {
      result.current.cycleTheme();
    });
    const afterCycle = result.current.theme;

    act(() => {
      mql.setSystemDark(true);
    });

    expect(result.current.theme).toBe(afterCycle);
  });

  it('persists the explicit-choice flag across remounts (a relaunch must not resume following the OS)', () => {
    const mql = installMatchMediaMock(false);
    const first = renderHook(() => useTheme());
    act(() => {
      first.result.current.setTheme('dark');
    });
    first.unmount();

    const second = renderHook(() => useTheme());
    expect(second.result.current.theme).toBe('dark');

    act(() => {
      mql.setSystemDark(true);
    });
    // Still ignores the OS after remounting — the earlier explicit choice
    // was recorded in localStorage, not just an in-memory flag.
    expect(second.result.current.theme).toBe('dark');
  });

  it('migration: a pre-existing stored theme with no explicit flag is treated as explicit, so an OS toggle does not override it', () => {
    // Simulates an install from before the explicit-choice flag existed:
    // `atlas-theme` is already on disk (from ordinary persistence, not
    // necessarily a deliberate ThemeMenu pick) but `atlas-theme-explicit`
    // has never been written.
    localStorage.setItem('atlas-theme', 'nord');
    const mql = installMatchMediaMock(false);

    const { result } = renderHook(() => useTheme());
    expect(result.current.theme).toBe('nord');

    act(() => {
      mql.setSystemDark(true);
    });

    expect(result.current.theme).toBe('nord');
    expect(localStorage.getItem('atlas-theme-explicit')).toBe('1');
  });

  it('a brand-new install (no stored theme at all) still follows the OS live, unaffected by the migration', () => {
    const mql = installMatchMediaMock(false);
    const { result } = renderHook(() => useTheme());

    expect(result.current.theme).toBe('light');

    act(() => {
      mql.setSystemDark(true);
    });

    expect(result.current.theme).toBe('dark');
  });

  it('regression (wave-3 shell-polish review): a brand-new install does not persist an explicit-choice flag just from a re-render', () => {
    // The mount effect writes `atlas-theme` to localStorage right after the
    // first render. If `hasExplicitChoice()`'s persisting side effect were
    // re-evaluated on every render (a non-lazy `useRef(hasExplicitChoice())`
    // initializer does this — it discards the return value after mount but
    // still re-runs the side effect), any later re-render — including one
    // triggered by a live OS toggle, not a user action — would wrongly and
    // permanently mark the install as having made an explicit choice.
    const mql = installMatchMediaMock(false)
    const { rerender } = renderHook(() => useTheme())

    // Force at least one extra render after mount, exactly as an OS-driven
    // update or any unrelated parent re-render would.
    act(() => {
      mql.setSystemDark(true)
    })
    rerender()

    expect(localStorage.getItem('atlas-theme-explicit')).toBeNull()
  })

  it('still follows further live OS toggles within the same mounted instance after several re-renders', () => {
    // Multiple re-renders in a row (each one, pre-fix, would re-run
    // hasExplicitChoice()'s persisting side effect) must not eventually
    // "trip" the in-memory explicitRef either — it was set once at mount
    // and must stay false across any number of later renders.
    const mql = installMatchMediaMock(false)
    const { result, rerender } = renderHook(() => useTheme())

    act(() => {
      mql.setSystemDark(true)
    })
    rerender()
    rerender()
    act(() => {
      mql.setSystemDark(false)
    })

    expect(result.current.theme).toBe('light')
    expect(localStorage.getItem('atlas-theme-explicit')).toBeNull()
  })

  it('cleans up its matchMedia listener on unmount', () => {
    const mql = installMatchMediaMock(false);
    const { unmount } = renderHook(() => useTheme());

    expect(mql.listenerCount()).toBe(1);
    unmount();
    expect(mql.listenerCount()).toBe(0);
  });
});
