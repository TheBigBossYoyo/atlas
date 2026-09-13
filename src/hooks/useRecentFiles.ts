import { useState, useCallback } from 'react';
import type { RecentFile } from '../types';

const KEY = 'atlas-recent';
const MAX = 8;

function readRecent(): RecentFile[] {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw == null) return [];

    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];

    return parsed.filter((item): item is RecentFile => {
      if (typeof item !== 'object' || item === null) return false;
      const candidate = item as Partial<RecentFile>;
      return (
        typeof candidate.path === 'string' &&
        typeof candidate.name === 'string' &&
        typeof candidate.openedAt === 'number'
      );
    });
  } catch {
    return [];
  }
}

function persistRecent(recent: RecentFile[]): void {
  localStorage.setItem(KEY, JSON.stringify(recent));
}

export function useRecentFiles() {
  const [recent, setRecent] = useState<RecentFile[]>(readRecent);

  const addRecent = useCallback((file: Omit<RecentFile, 'openedAt'>) => {
    setRecent(current => {
      const key = file.path || file.name;
      const next: RecentFile[] = [
        { ...file, openedAt: Date.now() },
        ...current.filter(item => (item.path || item.name) !== key),
      ].slice(0, MAX);
      persistRecent(next);
      return next;
    });
  }, []);

  const removeRecent = useCallback((key: string) => {
    setRecent(current => {
      const next = current.filter(item => item.path !== key && item.name !== key);
      persistRecent(next);
      return next;
    });
  }, []);

  const clearRecent = useCallback(() => {
    setRecent(() => {
      persistRecent([]);
      return [];
    });
  }, []);

  return { recent, addRecent, removeRecent, clearRecent } as const;
}
