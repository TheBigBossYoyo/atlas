import { useEffect } from 'react';

const KEY = 'atlas-draft';

export interface Draft {
  markdown: string;
  fileName: string | null;
  savedAt: number;
}

function isDraft(value: unknown): value is Draft {
  if (typeof value !== 'object' || value === null) return false;
  const candidate = value as Partial<Draft>;
  return (
    typeof candidate.markdown === 'string' &&
    (typeof candidate.fileName === 'string' || candidate.fileName === null) &&
    typeof candidate.savedAt === 'number'
  );
}

export function useAutosave(markdown: string, fileName: string | null, debounceMs = 800): void {
  useEffect(() => {
    const trimmed = markdown.trim();
    if (trimmed.length === 0) {
      localStorage.removeItem(KEY);
      return;
    }

    const timer = window.setTimeout(() => {
      const draft: Draft = { markdown, fileName, savedAt: Date.now() };
      localStorage.setItem(KEY, JSON.stringify(draft));
    }, debounceMs);

    return () => {
      window.clearTimeout(timer);
    };
  }, [debounceMs, fileName, markdown]);
}

export function loadDraft(): Draft | null {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw == null) return null;
    const parsed: unknown = JSON.parse(raw);
    return isDraft(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

export function clearDraft(): void {
  localStorage.removeItem(KEY);
}
