/**
 * useRecentFiles — unit tests (P4.8 / QA-27).
 *
 * QA-27: this hook previously had zero dedicated tests — only incidental
 * coverage via another hook's test (which never exercised dedup, ordering,
 * truncation, or the corrupt-storage fallback directly). Covers: add
 * (including dedup-by-path/name and MAX-length truncation), remove, clear,
 * `localStorage` persistence/rehydration, and defensive parsing of
 * malformed stored data.
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { renderHook, act } from '@testing-library/react';
import { useRecentFiles } from '../useRecentFiles';
import type { RecentFile } from '../../types';

const STORAGE_KEY = 'atlas-recent';
// Matches the hook's own internal MAX; asserted indirectly below (the hook
// doesn't export the constant, so this test pins the *observed* limit — if
// it's ever deliberately changed, this is the one place to update).
const MAX_RECENT = 8;

function readStoredRecent(): unknown {
  const raw = localStorage.getItem(STORAGE_KEY);
  return raw === null ? null : JSON.parse(raw);
}

beforeEach(() => {
  localStorage.clear();
});

afterEach(() => {
  localStorage.clear();
});

describe('useRecentFiles', () => {
  it('starts empty when localStorage has nothing stored yet', () => {
    const { result } = renderHook(() => useRecentFiles());
    expect(result.current.recent).toEqual([]);
  });

  it('addRecent prepends a new entry with an openedAt timestamp', () => {
    const { result } = renderHook(() => useRecentFiles());

    act(() => {
      result.current.addRecent({ path: '/abs/a.md', name: 'a.md' });
    });

    expect(result.current.recent).toHaveLength(1);
    expect(result.current.recent[0]).toMatchObject({ path: '/abs/a.md', name: 'a.md' });
    expect(typeof result.current.recent[0]?.openedAt).toBe('number');
  });

  it('orders entries newest-first', () => {
    const { result } = renderHook(() => useRecentFiles());

    act(() => {
      result.current.addRecent({ path: '/abs/a.md', name: 'a.md' });
    });
    act(() => {
      result.current.addRecent({ path: '/abs/b.md', name: 'b.md' });
    });
    act(() => {
      result.current.addRecent({ path: '/abs/c.md', name: 'c.md' });
    });

    expect(result.current.recent.map((r) => r.name)).toEqual(['c.md', 'b.md', 'a.md']);
  });

  it('dedups by path: re-adding the same path moves it to the front instead of duplicating', () => {
    const { result } = renderHook(() => useRecentFiles());

    act(() => {
      result.current.addRecent({ path: '/abs/a.md', name: 'a.md' });
    });
    act(() => {
      result.current.addRecent({ path: '/abs/b.md', name: 'b.md' });
    });
    act(() => {
      // Re-open "a.md" — same path, possibly a different display name.
      result.current.addRecent({ path: '/abs/a.md', name: 'a-renamed.md' });
    });

    expect(result.current.recent).toHaveLength(2);
    expect(result.current.recent.map((r) => r.name)).toEqual(['a-renamed.md', 'b.md']);
  });

  it('dedups by name when path is empty (browser mode has no filesystem path)', () => {
    const { result } = renderHook(() => useRecentFiles());

    act(() => {
      result.current.addRecent({ path: '', name: 'dropped.md' });
    });
    act(() => {
      result.current.addRecent({ path: '', name: 'dropped.md' });
    });

    expect(result.current.recent).toHaveLength(1);
  });

  it(`truncates to the most recent ${MAX_RECENT} entries`, () => {
    const { result } = renderHook(() => useRecentFiles());

    for (let i = 0; i < MAX_RECENT + 3; i++) {
      act(() => {
        result.current.addRecent({ path: `/abs/${i}.md`, name: `${i}.md` });
      });
    }

    expect(result.current.recent).toHaveLength(MAX_RECENT);
    // Newest-first: the last MAX_RECENT adds survive, oldest ones are dropped.
    expect(result.current.recent[0]).toMatchObject({ name: `${MAX_RECENT + 2}.md` });
    expect(result.current.recent.map((r) => r.name)).not.toContain('0.md');
  });

  it('persists every mutation to localStorage', () => {
    const { result } = renderHook(() => useRecentFiles());

    act(() => {
      result.current.addRecent({ path: '/abs/a.md', name: 'a.md' });
    });

    const stored = readStoredRecent();
    expect(Array.isArray(stored)).toBe(true);
    expect(stored).toMatchObject([{ path: '/abs/a.md', name: 'a.md' }]);
  });

  it('rehydrates from a previously persisted list on mount', () => {
    const persisted: RecentFile[] = [
      { path: '/abs/old.md', name: 'old.md', openedAt: 1000 },
    ];
    localStorage.setItem(STORAGE_KEY, JSON.stringify(persisted));

    const { result } = renderHook(() => useRecentFiles());

    expect(result.current.recent).toEqual(persisted);
  });

  it('removeRecent drops the matching entry by path and persists the change', () => {
    const { result } = renderHook(() => useRecentFiles());

    act(() => {
      result.current.addRecent({ path: '/abs/a.md', name: 'a.md' });
    });
    act(() => {
      result.current.addRecent({ path: '/abs/b.md', name: 'b.md' });
    });
    act(() => {
      result.current.removeRecent('/abs/a.md');
    });

    expect(result.current.recent.map((r) => r.path)).toEqual(['/abs/b.md']);
    expect(readStoredRecent()).toMatchObject([{ path: '/abs/b.md' }]);
  });

  it('removeRecent also matches by name (browser mode, empty path)', () => {
    const { result } = renderHook(() => useRecentFiles());

    act(() => {
      result.current.addRecent({ path: '', name: 'dropped.md' });
    });
    act(() => {
      result.current.removeRecent('dropped.md');
    });

    expect(result.current.recent).toEqual([]);
  });

  it('clearRecent empties the list and the persisted storage', () => {
    const { result } = renderHook(() => useRecentFiles());

    act(() => {
      result.current.addRecent({ path: '/abs/a.md', name: 'a.md' });
      result.current.addRecent({ path: '/abs/b.md', name: 'b.md' });
    });
    act(() => {
      result.current.clearRecent();
    });

    expect(result.current.recent).toEqual([]);
    expect(readStoredRecent()).toEqual([]);
  });

  it('falls back to an empty list when localStorage holds invalid JSON', () => {
    localStorage.setItem(STORAGE_KEY, '{not valid json');

    const { result } = renderHook(() => useRecentFiles());

    expect(result.current.recent).toEqual([]);
  });

  it('falls back to an empty list when the stored value is not an array', () => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ not: 'an array' }));

    const { result } = renderHook(() => useRecentFiles());

    expect(result.current.recent).toEqual([]);
  });

  it('filters out malformed entries while keeping well-formed ones', () => {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify([
        { path: '/abs/good.md', name: 'good.md', openedAt: 123 },
        { path: '/abs/missing-name.md', openedAt: 123 },
        { name: 'missing-openedAt.md' },
        null,
        'not an object',
      ]),
    );

    const { result } = renderHook(() => useRecentFiles());

    expect(result.current.recent).toEqual([{ path: '/abs/good.md', name: 'good.md', openedAt: 123 }]);
  });
});
